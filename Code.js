// ==========================================
//  1. LIVE WEBHOOK TRAFFIC COP
// ==========================================
function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  var data = json.data;
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  if (json.event === "ticket.created") {
    return handleTicketCreated(ss, data);
  }
  if (json.event === "transaction.succeeded") {
    return handleTransactionSucceeded(ss, data);
  }
  return ContentService.createTextOutput(JSON.stringify({"status": "ignored", "event": json.event})).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
//  HANDLES: ticket.created (volunteer signups)
// ==========================================
function handleTicketCreated(ss, data) {
  var ticketTitle = (data.title || "").toLowerCase().trim();
  var donorName = (data.first_name || "") + " " + (data.last_name || "");
  var donorEmail = data.email || "";
  var currentDate = new Date();
  var ticketId = data.id || "";

  var VOLUNTEER_TICKET_MATCH = "general admission";
  var BANQUET_TICKET_KEYWORDS = ["banquet", "honorary"];
  
  var isBanquetTicket = BANQUET_TICKET_KEYWORDS.some(function(k) { return ticketTitle.includes(k); });
  if (isBanquetTicket) {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped banquet ticket"})).setMimeType(ContentService.MimeType.JSON);
  }

  var isVolunteerTicket = ticketTitle.includes(VOLUNTEER_TICKET_MATCH);
  if (!isVolunteerTicket) {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped unrecognized ticket", "title": data.title})).setMimeType(ContentService.MimeType.JSON);
  }

  var volunteerSheet = ss.getSheetByName("Volunteers");
  var volunteerLastRow = volunteerSheet.getLastRow();
  if (volunteerLastRow > 1 && ticketId) {
    var seenTicketIds = volunteerSheet.getRange("E2:E" + volunteerLastRow).getValues().flat().map(function(v) {
      return v.toString().trim();
    });
    if (seenTicketIds.indexOf(ticketId.toString().trim()) !== -1) {
      return ContentService.createTextOutput(JSON.stringify({"status": "skipped duplicate volunteer"})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  volunteerSheet.appendRow([currentDate, donorName, donorEmail, data.title || "Volunteer", ticketId]);

  var reportSheet = ss.getSheetByName("Weekly Report");
  var liveVolunteerCount = reportSheet.getRange("G2").getValue();

  var volunteerMessage = "🙋🏽‍♂️ *NEW VOLUNTEER REGISTERED!*\n" +
                         "*" + donorName + "* signed up for our *2026 Day Of Service*! 💛\n" +
                         "🚀 Total mobilized: *" + liveVolunteerCount + "* volunteers!";
  sendGoogleChatNotification(volunteerMessage, CONFIG.URL_VOLUNTEER_ALERTS);

  // Check volunteer milestones (e.g., hitting 60 and beyond)
  checkAndAnnounceVolunteerMilestone(ss);

  return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// HANDLES: transaction.succeeded (actual donations)
// ==========================================
function handleTransactionSucceeded(ss, data) {
  var BANQUET_CAMPAIGN_CODE = "TTMK9G"; 
  if (data.campaign_code && BANQUET_CAMPAIGN_CODE !== "TTMK9G" &&
      data.campaign_code.toLowerCase() === BANQUET_CAMPAIGN_CODE.toLowerCase()) {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped banquet"})).setMimeType(ContentService.MimeType.JSON);
  }

  var amount = (data.amount && !isNaN(data.amount)) ? parseFloat(data.amount) : 0.00;
  var transactionStatus = (data.status || "").toLowerCase().trim();

  if (transactionStatus === "revoked") {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped revoked"})).setMimeType(ContentService.MimeType.JSON);
  }
  if (amount <= 0) {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped"})).setMimeType(ContentService.MimeType.JSON);
  }

  var donorName = (data.first_name || "") + " " + (data.last_name || "");
  var donorEmail = data.email || "";
  var currentDate = new Date();
  var transactionId = data.id || "";

  var scholarName = "General Campaign";
  var teamCohort = "General Campaign";

  if (data.member_id) {
    var rosterEntry = getRosterEntry(ss, data.member_id);
    if (rosterEntry) {
      scholarName = rosterEntry.name;
      teamCohort = rosterEntry.cohort;
    } else {
      scholarName = "⚠️ Unmapped Member ID: " + data.member_id;
    }
  }

  var donorSheet = ss.getSheetByName("Donors");
  var donorLastRow = donorSheet.getLastRow();
  if (donorLastRow > 1 && transactionId) {
    var seenTransactionIds = donorSheet.getRange("G2:G" + donorLastRow).getValues().flat().map(function(v) {
      return v.toString().trim();
    });
    if (seenTransactionIds.indexOf(transactionId.toString().trim()) !== -1) {
      return ContentService.createTextOutput(JSON.stringify({"status": "skipped duplicate"})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  donorSheet.appendRow([currentDate, donorName, donorEmail, teamCohort, amount, scholarName, transactionId]);

  // Check fundraising milestone & approaching milestones
  checkAndAnnounceMilestone(ss);

  return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. CONDENSED DAILY COUNTDOWN RECAP
// ==========================================
var EVENT_COUNTDOWN_START = new Date("2026-07-20T00:00:00");
var EVENT_DATE            = new Date("2026-07-25T23:59:59");

function buildRecapMessage(headerTitle, bannerOverride) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var reportSheet = ss.getSheetByName("Weekly Report");
  
  var totalRaised = reportSheet.getRange("B2").getValue() || 0;
  var totalVolunteers = reportSheet.getRange("G2").getValue() || 0;
  var progressPercentage = (totalRaised / CONFIG.GOAL_AMOUNT) * 100;
  
  var leaderboard = getTopTeamData(reportSheet);
  var winningTeam = leaderboard.teamName; 
  var winningAmount = leaderboard.amountRaised;

  var analytics = getAllTimeCampaignAnalytics(ss);
  
  var top5String = "";
  var emojis = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣"];
  for (var i = 0; i < analytics.top5Rows.length; i++) {
    top5String += emojis[i] + " " + analytics.top5Rows[i].name + " — $" + analytics.top5Rows[i].amount.toFixed(2) + "\n";
  }
  if (top5String === "") { top5String = "No individual donations registered yet! 🌱\n"; }

  // 🎯 Dynamic Milestone / Approaching Milestone Banner Logic
  var milestoneInterval = 500;
  var nextTier = Math.ceil(totalRaised / milestoneInterval) * milestoneInterval;
  var distanceToMilestone = nextTier - totalRaised;
  var cushion = 100; // Triggers "Almost There" if within $100 of the next milestone
  
  var milestoneBanner = "";
  if (distanceToMilestone > 0 && distanceToMilestone <= cushion) {
    milestoneBanner = "👀 *ALMOST THERE:* Just *$" + distanceToMilestone.toFixed(2) + "* away from crossing *$" + nextTier.toLocaleString() + "*! 🏁💙\n\n";
  } else {
    var currentMilestoneTier = Math.floor(totalRaised / 500) * 500;
    if (currentMilestoneTier >= 500) {
      var percentageMatched = ((currentMilestoneTier / CONFIG.GOAL_AMOUNT) * 100).toFixed(0);
      milestoneBanner = "🎉 *MILESTONE UNLOCKED:* Crossed *$" + currentMilestoneTier.toLocaleString() + "* (" + percentageMatched + "% of our goal)! 🙌🏽💙\n\n";
    }
  }

  var highestDonorSection = "🥇 *WEEKLY HIGHEST DONOR:* Let's kick off the week with some donations!";
  if (analytics.weeklyTopDonorsList.length > 0) {
    var donorListStr = analytics.weeklyTopDonorsList.join(", ");
    highestDonorSection = "🥇 *WEEKLY HIGHEST DONOR(S):* Shoutout to *" + donorListStr + "* for leading this week with top donations of *$" + analytics.weeklyTopDonationAmount.toFixed(2) + "*! 🙌🏽💎";
  }

  var banner = bannerOverride || "";

  var message = headerTitle + "\n" +
                "━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
                banner +
                milestoneBanner +
                "💰 *TOTAL RAISED:* *$" + totalRaised.toFixed(2) + "* / $" + CONFIG.GOAL_AMOUNT.toFixed(2) + " (" + progressPercentage.toFixed(0) + "% progress) 🚀\n" +
                "🙋🏽‍♂️ *VOLUNTEERS:* *" + totalVolunteers + "* leaders registered!\n" +
                "☀️ *RAISED TODAY:* +$" + analytics.todayTotalRaised.toFixed(2) + " added today! 🚀\n" +
                // "📈 *THIS WEEK:* Added +$" + weeklyDifference.toFixed(2) + " to the campaign!\n\n" + (Commented out for Friday)
                "\n" +
                "🏆 *COHORT LEADERBOARD:*\n" +
                "👉 *" + winningTeam + "* leads the race with *$" + winningAmount.toFixed(2) + "* raised! 🔥\n\n" +
                "💎 *ALL-TIME TOP SCHOLARS:*\n" +
                top5String + 
                "\n" +
                highestDonorSection + "\n\n" +
                "☀️ *ACTIVE FUNDRASIERS TODAY:* " + (analytics.todayActiveScholarsList || "None yet today—let's get the momentum rolling!") + "\n\n" +
                "👏 *ACTIVE WEEKLY FUNDRAISERS:* " + (analytics.activeScholarsList || "None this week yet—let's secure that first donation!") + "\n\n" +
                "Every single share and donation brings us closer to making an impact. Let's keep supporting one another and push hard as a cohort! 💛💙";

  return message;
}

function sendDailyCountdownRecap() {
  var today = new Date();
  if (today < EVENT_COUNTDOWN_START || today > EVENT_DATE) {
    return; 
  }
  var msPerDay = 1000 * 60 * 60 * 24;
  var daysLeft = Math.ceil((EVENT_DATE.getTime() - today.getTime()) / msPerDay);
  var banner = (daysLeft > 0)
    ? ("⏳ *" + daysLeft + " DAY" + (daysLeft === 1 ? "" : "S") + " UNTIL DAY OF SERVICE!*\n\n")
    : ("🎉 *TODAY IS DAY OF SERVICE!!* 🎉\n\n");
  var message = buildRecapMessage("🚨 *SLS DAY OF SERVICE: DAILY COUNTDOWN* 🚨", banner);
  sendGoogleChatNotification(message, CONFIG.URL_WEEKLY_RECAP);
}

function checkAndAnnounceMilestone(ss) {
  SpreadsheetApp.flush(); 
  var reportSheet = ss.getSheetByName("Weekly Report");
  var totalRaised = reportSheet.getRange("B2").getValue() || 0;
  
  var currentTier = Math.floor(totalRaised / 500) * 500;
  var props = PropertiesService.getScriptProperties();
  var lastAnnounced = parseInt(props.getProperty("LAST_MILESTONE_ANNOUNCED") || "0", 10);
  
  if (currentTier > lastAnnounced && currentTier > 0) {
    var percentReached = ((currentTier / CONFIG.GOAL_AMOUNT) * 100).toFixed(0);
    var milestoneMsg = "🎉 *MILESTONE UNLOCKED!* 🎉\n" +
                       "We just crossed *$" + currentTier.toLocaleString() + "* — that's *" + percentReached + "%* of our $" + CONFIG.GOAL_AMOUNT.toLocaleString() + " goal! 🙌🏽💙\n" +
                       "Keep it going, SLS fam! 💛";
    sendGoogleChatNotification(milestoneMsg, CONFIG.URL_WEEKLY_RECAP);
    props.setProperty("LAST_MILESTONE_ANNOUNCED", currentTier.toString());
  }
  
  // Check if we are approaching a milestone ("Almost There" alert)
  checkApproachingMilestone(ss, totalRaised);
}

function checkApproachingMilestone(ss, totalRaised) {
  var milestoneInterval = 500; 
  var nextTier = Math.ceil(totalRaised / milestoneInterval) * milestoneInterval;
  var cushion = 100; // Triggers if within $100 of the next milestone
  var distanceToMilestone = nextTier - totalRaised;
  
  if (distanceToMilestone > 0 && distanceToMilestone <= cushion) {
    var props = PropertiesService.getScriptProperties();
    var lastApproachingAlert = props.getProperty("LAST_APPROACHING_MILESTONE") || "";
    
    if (lastApproachingAlert !== nextTier.toString()) {
      var approachingMsg = "👀 *ALMOST THERE, SLS FAM!* 👀\n" +
                           "We are just *$" + distanceToMilestone.toFixed(2) + "* away from crossing *$" + nextTier.toLocaleString() + "*! 🏁💙\n" +
                           "Let's secure a few final donations and smash this milestone right now! 💰🔥";
      sendGoogleChatNotification(approachingMsg, CONFIG.URL_WEEKLY_RECAP);
      props.setProperty("LAST_APPROACHING_MILESTONE", nextTier.toString());
    }
  }
}

function checkAndAnnounceVolunteerMilestone(ss) {
  SpreadsheetApp.flush();
  var reportSheet = ss.getSheetByName("Weekly Report");
  var totalVolunteers = reportSheet.getRange("G2").getValue() || 0;
  
  var currentVolTier = Math.floor(totalVolunteers / 50) * 50;
  var props = PropertiesService.getScriptProperties();
  var lastAnnouncedVol = parseInt(props.getProperty("LAST_VOLUNTEER_MILESTONE") || "0", 10);
  
  if ((currentVolTier > lastAnnouncedVol && currentVolTier > 0) || (totalVolunteers >= 60 && lastAnnouncedVol < 60)) {
    var volMilestoneMsg = "🎉 *VOLUNTEER MILESTONE UNLOCKED!* 🎉\n" +
                          "We have officially mobilized *" + totalVolunteers + "* leaders for our Day of Service! 🙌🏽💛\n" +
                          "The momentum is unreal. Let's keep pushing for even more change maker signups! 🚀";
    sendGoogleChatNotification(volMilestoneMsg, CONFIG.URL_WEEKLY_RECAP);
    props.setProperty("LAST_VOLUNTEER_MILESTONE", totalVolunteers.toString());
  }
}

// ==========================================
// 3. AUTOMATIC SNAPSHOT ARCHIVER
// ==========================================
function runWeeklySnapshot() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var reportSheet = ss.getSheetByName("Weekly Report");
  var snapshotSheet = ss.getSheetByName("Snapshot History");
  
  var totalRaised = reportSheet.getRange("B2").getValue() || 0;
  var totalVolunteers = reportSheet.getRange("G2").getValue() || 0;
  
  var atlantaTotal = 0;
  var athensTotal = 0;
  var pivotRows = reportSheet.getRange("I3:J10").getValues();
  for (var i = 0; i < pivotRows.length; i++) {
    var team = pivotRows[i][0] ? pivotRows[i][0].toString().toLowerCase().trim() : "";
    var amount = pivotRows[i][1] && !isNaN(pivotRows[i][1]) ? parseFloat(pivotRows[i][1]) : 0;
    if (team.includes("atlanta")) atlantaTotal = amount;
    if (team.includes("athens")) athensTotal = amount;
  }
  
  var analytics = getAllTimeCampaignAnalytics(ss);
  var weeklyHighestAmount = analytics.weeklyTopDonationAmount || 0.00;
  var weeklyHighestName = analytics.weeklyTopDonorsList.length > 0 ? analytics.weeklyTopDonorsList.join(", ") : "N/A";
  
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd");
  var weekLabel = "Week ending " + formattedDate;
  
  snapshotSheet.appendRow([
    weekLabel, 
    totalRaised, 
    atlantaTotal, 
    athensTotal, 
    weeklyHighestAmount, 
    weeklyHighestName, 
    totalVolunteers, 
    today
  ]);
}

// ==========================================
// 🔍 HELPERS & ANALYTICS (CALENDAR-WEEK ANCHORED)
// ==========================================
function getRosterEntry(ss, memberId) {
  var rosterSheet = ss.getSheetByName("Roster");
  if (!rosterSheet) return null;
  var lastRow = rosterSheet.getLastRow();
  if (lastRow < 2) return null;
  
  var rows = rosterSheet.getRange("A2:C" + lastRow).getValues();
  var cleanId = memberId.toString().trim();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0].toString().trim() === cleanId) {
      return { name: rows[i][1].toString().trim(), cohort: rows[i][2].toString().trim() };
    }
  }
  return null;
}

function reprocessUnmappedDonors() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var donorSheet = ss.getSheetByName("Donors");
  var lastRow = donorSheet.getLastRow();
  if (lastRow < 2) return;
  
  var data = donorSheet.getRange("A2:F" + lastRow).getValues(); 
  var fixedCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var scholarCell = data[i][5] ? data[i][5].toString() : "";
    var match = scholarCell.match(/Unmapped Member ID:\s*(\S+)/);
    if (match) {
      var memberId = match[1];
      var rosterEntry = getRosterEntry(ss, memberId);
      if (rosterEntry) {
        var sheetRow = i + 2; 
        donorSheet.getRange(sheetRow, 4).setValue(rosterEntry.cohort);   
        donorSheet.getRange(sheetRow, 6).setValue(rosterEntry.name);    
        fixedCount++;
      }
    }
  }
}

function getTopTeamData(sheet) {
  var lastRow = sheet.getLastRow();
  var defaultResult = { teamName: "No Cohort Yet", amountRaised: 0.00, contributors: 0 };
  if (lastRow < 3) return defaultResult;
  
  var dataRange = sheet.getRange("I3:K" + lastRow).getValues();
  var maxAmount = -1; var topTeam = ""; var topContributors = 0;
  
  for (var i = 0; i < dataRange.length; i++) {
    var team = dataRange[i][0] ? dataRange[i][0].toString().trim() : "";
    var amount = dataRange[i][1] && !isNaN(dataRange[i][1]) ? parseFloat(dataRange[i][1]) : 0;
    var count = dataRange[i][2] && !isNaN(dataRange[i][2]) ? parseInt(dataRange[i][2]) : 0;
    
    if (team === "" || team === "0" || team.toLowerCase().includes("grand total") || team.toLowerCase().includes("team / member") || team.toLowerCase() === "individual" || team.toLowerCase().includes("general")) continue;
    
    if (amount > maxAmount) {
      maxAmount = amount; topTeam = team; topContributors = count;
    }
  }
  return topTeam !== "" ? { teamName: topTeam, amountRaised: maxAmount, contributors: topContributors } : defaultResult;
}

function getAllTimeCampaignAnalytics(ss) {
  var donorSheet = ss.getSheetByName("Donors");
  var lastRow = donorSheet.getLastRow();
  
  var result = {
    top5Rows: [],
    weeklyTopDonorsList: [],
    weeklyTopDonationAmount: 0.00,
    activeScholarsList: "",
    todayActiveScholarsList: "",
    todayTotalRaised: 0.00
  };
  
  if (lastRow < 2) return result;
  
  var rawData = donorSheet.getRange("A2:F" + lastRow).getValues();
  
  var today = new Date();
  var dayOfWeek = today.getDay(); 
  var distanceToMonday = (dayOfWeek === 0) ? -6 : (1 - dayOfWeek);
  
  var startOfThisWeek = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  startOfThisWeek.setDate(today.getDate() + distanceToMonday);
  startOfThisWeek.setHours(0, 0, 0, 0);
  
  var allTimeTotals = {};
  var weeklyDonorMaxMap = {}; 
  var scholarsActiveThisWeek = [];
  var scholarsActiveToday = [];
  var todayTotalRaised = 0.00;
  
  var maxWeeklyDonationFound = 0;
  
  for (var i = 0; i < rawData.length; i++) {
    var timestamp = rawData[i][0];
    var donorName = rawData[i][1] ? rawData[i][1].toString().trim() : "";
    var amount = rawData[i][4] && !isNaN(rawData[i][4]) ? parseFloat(rawData[i][4]) : 0.00;
    var scholarName = rawData[i][5] ? rawData[i][5].toString().trim() : "";
    
    if (amount <= 0) continue;
    
    var cleanScholar = scholarName.toLowerCase();
    if (cleanScholar === "general campaign" || cleanScholar === "campaign" || cleanScholar === "athens" || cleanScholar === "atlanta" || cleanScholar === "") continue;
    
    if (!allTimeTotals[scholarName]) {
      allTimeTotals[scholarName] = 0.00;
    }
    allTimeTotals[scholarName] += amount;
    
    var transactionDate = new Date(timestamp);
    
    if (!isNaN(transactionDate.getTime()) && transactionDate >= startOfThisWeek) {
      if (scholarsActiveThisWeek.indexOf(scholarName) === -1) {
        scholarsActiveThisWeek.push(scholarName);
      }
      if (donorName) {
        if (!weeklyDonorMaxMap[donorName] || amount > weeklyDonorMaxMap[donorName]) {
          weeklyDonorMaxMap[donorName] = amount;
        }
      }
    }
    
    if (!isNaN(transactionDate.getTime()) && 
        transactionDate.getFullYear() === today.getFullYear() &&
        transactionDate.getMonth() === today.getMonth() &&
        transactionDate.getDate() === today.getDate()) {
      if (scholarsActiveToday.indexOf(scholarName) === -1) {
        scholarsActiveToday.push(scholarName);
      }
      todayTotalRaised += amount;
    }
  }
  
  for (var dName in weeklyDonorMaxMap) {
    var amt = weeklyDonorMaxMap[dName];
    if (amt > maxWeeklyDonationFound) {
      maxWeeklyDonationFound = amt;
      result.weeklyTopDonorsList = [dName];
    } else if (amt === maxWeeklyDonationFound && maxWeeklyDonationFound > 0) {
      if (result.weeklyTopDonorsList.indexOf(dName) === -1) {
        result.weeklyTopDonorsList.push(dName);
      }
    }
  }
  result.weeklyTopDonationAmount = maxWeeklyDonationFound;
  
  var sortedAllTime = [];
  for (var key in allTimeTotals) {
    sortedAllTime.push({ name: key, amount: allTimeTotals[key] });
  }
  sortedAllTime.sort(function(a, b) { return b.amount - a.amount; });
  
  result.top5Rows = sortedAllTime.slice(0, 5);
  result.activeScholarsList = scholarsActiveThisWeek.join(", ");
  result.todayActiveScholarsList = scholarsActiveToday.join(", ");
  result.todayTotalRaised = todayTotalRaised;
  
  return result;
}

// ==========================================
//  GOOGLE CHAT WEBHOOK SENDER
// ==========================================
function sendGoogleChatNotification(text, webhookUrl) {
  var targetUrl = webhookUrl || CONFIG.URL_VOLUNTEER_ALERTS;
  UrlFetchApp.fetch(targetUrl, {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({"text": text})
  });
}
