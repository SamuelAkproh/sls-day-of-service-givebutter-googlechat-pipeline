// ==========================================
//  1. LIVE WEBHOOK TRAFFIC COP (UPDATED DYNAMIC MAPPING)
// ==========================================
function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  var data = json.data;
  
  //  CAMPAIGN FILTER: Skip Honorary Banquet registrations
  var campaignTitle = (data.campaign && data.campaign.title) ? data.campaign.title.toLowerCase() : "";
  if (campaignTitle.includes("banquet")) {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped banquet"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Skip uncompleted transactions
  if (json.event === "transaction.success" && (!data.amount || parseFloat(data.amount) === 0) && !data.tickets) {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  var donorName = (data.first_name || "") + " " + (data.last_name || "");
  var donorEmail = data.email || "";
  var amount = (data.amount && !isNaN(data.amount)) ? parseFloat(data.amount) : 0.00;
  var transactionStatus = (data.status || "").toLowerCase().trim();
  
  //  STRICT DYNAMIC COHORT & SCHOLAR MAPPING (Aligned with Givebutter Team Assignments)
  var scholarName = "General Campaign";
  var teamCohort = "Individual"; // Matches Givebutter's non-team structures
  
  if (data.member) {
    scholarName = (data.member.first_name || "") + " " + (data.member.last_name || "");
    
    if (data.member.teams && data.member.teams.length > 0) {
      var teamName = data.member.teams[0].name.toLowerCase();
      if (teamName.includes("atlanta")) teamCohort = "Atlanta";
      else if (teamName.includes("athens")) teamCohort = "Athens";
    }
  }
  
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var currentDate = new Date(); 
  var cleanIncomingEmail = donorEmail.toString().toLowerCase().trim();
  
  // Detect if this is a volunteer registration
  var isVolunteerRegistration = false;
  if (data.tickets && data.tickets.length > 0) {
    for (var j = 0; j < data.tickets.length; j++) {
      var ticketName = (data.tickets[j].name || "").toLowerCase();
      if (ticketName.includes("volunteer") || ticketName.includes("t-shirt") || ticketName.includes("admission")) {
        isVolunteerRegistration = true;
      }
    }
  }

  // Skip revoked registrations
  if (transactionStatus === "revoked") {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped revoked"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (amount > 0 && !isVolunteerRegistration) {
    // ---- PROCESS ACTUAL DONATION ----
    var donorSheet = ss.getSheetByName("Donors");
    var donorLastRow = donorSheet.getLastRow();
    if (donorLastRow > 1) {
      var donorEmails = donorSheet.getRange("C1:C" + donorLastRow).getValues().flat().map(function(eStr) {
        return eStr.toString().toLowerCase().trim();
      });
      if (donorEmails.indexOf(cleanIncomingEmail) !== -1 && amount === 50.00) { 
        return ContentService.createTextOutput(JSON.stringify({"status": "skipped duplicate"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // Save dynamically to Google Sheets (silent log)
    donorSheet.appendRow([currentDate, donorName, donorEmail, teamCohort, amount, scholarName]);
    
  } else if (isVolunteerRegistration) {
    // ---- PROCESS VOLUNTEER REGISTRATION ----
    var volunteerSheet = ss.getSheetByName("Volunteers");
    var volunteerLastRow = volunteerSheet.getLastRow();
    if (volunteerLastRow > 1) {
      var volunteerEmails = volunteerSheet.getRange("C1:C" + volunteerLastRow).getValues().flat().map(function(eStr) {
        return eStr.toString().toLowerCase().trim();
      });
      if (volunteerEmails.indexOf(cleanIncomingEmail) !== -1) {
        return ContentService.createTextOutput(JSON.stringify({"status": "skipped duplicate volunteer"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    volunteerSheet.appendRow([currentDate, donorName, donorEmail, teamCohort]);
    
    var reportSheet = ss.getSheetByName("Weekly Report");
    var liveVolunteerCount = reportSheet.getRange("G2").getValue();
    
    var volunteerMessage = "🙋🏽‍♂️ *NEW VOLUNTEER REGISTERED!*\n" +
                           "*" + donorName + "* signed up for our *2026 Day Of Service*! 💛\n" +
                           "🚀 Total mobilized: *" + liveVolunteerCount + "* volunteers!";
                               
    sendGoogleChatNotification(volunteerMessage, CONFIG.URL_VOLUNTEER_ALERTS);
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
//  2. CONDENSED WEEKLY RECAP SCRIPT (ROLLING 7-DAY WEEKLY FILTERS)
// ==========================================
function sendWeeklyRecap() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var reportSheet = ss.getSheetByName("Weekly Report");
  
  var totalRaised = reportSheet.getRange("B2").getValue();
  var weeklyDifference = reportSheet.getRange("E2").getValue();
  var totalVolunteers = reportSheet.getRange("G2").getValue();
  var progressPercentage = (totalRaised / CONFIG.GOAL_AMOUNT) * 100;
  
  var leaderboard = getTopTeamData(reportSheet);
  var winningTeam = leaderboard.teamName; 
  var winningAmount = leaderboard.amountRaised;

  // Fetch true rolling 7-day calculations
  var weeklyAnalytics = getRollingWeeklyAnalytics(ss);
  
  var top5String = "";
  var emojis = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣"];
  for (var i = 0; i < weeklyAnalytics.top5Rows.length; i++) {
    top5String += emojis[i] + " " + weeklyAnalytics.top5Rows[i].name + " — $" + weeklyAnalytics.top5Rows[i].amount.toFixed(2) + "\n";
  }
  if (top5String === "") { top5String = "No individual donations registered this week yet! 🌱\n"; }

  //  DYNAMIC MILESTONE ENGINE (Every $500)
  var currentMilestoneTier = Math.floor(totalRaised / 500) * 500;
  var milestoneCelebration = "";
  if (currentMilestoneTier >= 500) {
    var percentageMatched = ((currentMilestoneTier / CONFIG.GOAL_AMOUNT) * 100).toFixed(0);
    milestoneCelebration = "🎉 *MILESTONE UNLOCKED:* Crossed *$" + currentMilestoneTier.toLocaleString() + "* (" + percentageMatched + "% of our goal)! 🙌🏽💙\n\n";
  }

  var highestDonorSection = "🥇 *WEEKLY HIGHEST DONOR:* No donations logged this week yet! Let's get the ball rolling!";
  if (weeklyAnalytics.topWeeklyDonor.amount > 0) {
    highestDonorSection = "🥇 *WEEKLY HIGHEST DONOR:* Shoutout to *" + weeklyAnalytics.topWeeklyDonor.name + "* for anchoring our efforts this week with a key donation of *$" + weeklyAnalytics.topWeeklyDonor.amount.toFixed(2) + "*! 🙌🏽💎";
  }

  var message = "🚨 *SLS DAY OF SERVICE: WEEKLY RECAP* 🚨\n" +
                "━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
                milestoneCelebration +
                "💰 *TOTAL RAISED:* *$" + totalRaised.toFixed(2) + "* / $" + CONFIG.GOAL_AMOUNT.toFixed(2) + " (" + progressPercentage.toFixed(0) + "% progress) 🚀\n" +
                "🙋🏽‍♂️ *VOLUNTEERS:* *" + totalVolunteers + "* leaders registered!\n" +
                "📈 *THIS WEEK:* Added *+$" + weeklyDifference.toFixed(2) + "* to the campaign!\n\n" +
                "🏆 *COHORT LEADERBOARD:*\n" +
                "👉 *" + winningTeam + "* leads the race with *$" + winningAmount.toFixed(2) + "* raised! 🔥\n\n" +
                "💎 *WEEKLY TOP INDIVIDUALS (Last 7 Days):*\n" +
                top5String + 
                "\n" +
                highestDonorSection + "\n\n" +
                "👏 *ACTIVE WEEKLY FUNDRAISERS:* " + (weeklyAnalytics.activeScholarsList || "None this week yet—let's secure that first donation!") + "\n\n" +
                "Every single share and donation brings us closer to making an impact. Let's keep supporting one another and push hard as a cohort over the weekend! 💛💙";

  // Note: To test in your sandbox, change CONFIG.URL_WEEKLY_RECAP to CONFIG.URL_VOLUNTEER_ALERTS here!
  sendGoogleChatNotification(message, CONFIG.URL_WEEKLY_RECAP);
}

// ==========================================
//  3. AUTOMATIC SNAPSHOT ARCHIVER (RUNS EVERY MONDAY MIDNIGHT)
// ==========================================
function runWeeklySnapshot() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var reportSheet = ss.getSheetByName("Weekly Report");
  var snapshotSheet = ss.getSheetByName("Snapshot History");
  
  var totalRaised = reportSheet.getRange("B2").getValue();
  var totalVolunteers = reportSheet.getRange("G2").getValue();
  
  // Calculate dynamic team totals from the Weekly Report's Pivot Section
  var atlantaTotal = 0;
  var athensTotal = 0;
  var pivotRows = reportSheet.getRange("I3:J10").getValues();
  for (var i = 0; i < pivotRows.length; i++) {
    var team = pivotRows[i][0] ? pivotRows[i][0].toString().toLowerCase().trim() : "";
    var amount = pivotRows[i][1] && !isNaN(pivotRows[i][1]) ? parseFloat(pivotRows[i][1]) : 0;
    if (team.includes("atlanta")) atlantaTotal = amount;
    if (team.includes("athens")) athensTotal = amount;
  }
  
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd");
  var weekLabel = "Week ending " + formattedDate;
  
  // Save dynamic records matching line chart metrics
  snapshotSheet.appendRow([weekLabel, totalRaised, atlantaTotal, athensTotal, totalVolunteers, today]);
}

// ==========================================
//  HELPER FUNCTION: DYNAMIC PIVOT TABLE SCANNER
// ==========================================
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
    
    // Ignore meta-totals and individuals outside teams
    if (team === "" || team === "0" || team.toLowerCase().includes("grand total") || team.toLowerCase().includes("team / member") || team.toLowerCase() === "individual" || team.toLowerCase().includes("general")) continue;
    
    if (amount > maxAmount) {
      maxAmount = amount; topTeam = team; topContributors = count;
    }
  }
  return topTeam !== "" ? { teamName: topTeam, amountRaised: maxAmount, contributors: topContributors } : defaultResult;
}

// ==========================================
//  HELPER FUNCTION: ROLLING 7-DAY WEEKLY FILTERS
// ==========================================
function getRollingWeeklyAnalytics(ss) {
  var donorSheet = ss.getSheetByName("Donors");
  var lastRow = donorSheet.getLastRow();
  
  var result = {
    top5Rows: [],
    topWeeklyDonor: { name: "", amount: 0.00 },
    activeScholarsList: ""
  };
  
  if (lastRow < 2) return result;
  
  var rawData = donorSheet.getRange("A2:F" + lastRow).getValues();
  
  var today = new Date();
  var sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(today.getDate() - 7);
  
  var weeklyScholarTotals = {};
  var scholarsActiveThisWeek = [];
  var maxWeeklyDonation = 0;
  var topWeeklyDonorName = "";
  
  for (var i = 0; i < rawData.length; i++) {
    var timestamp = rawData[i][0];
    var donorName = rawData[i][1] ? rawData[i][1].toString().trim() : "";
    var amount = rawData[i][4] && !isNaN(rawData[i][4]) ? parseFloat(rawData[i][4]) : 0.00;
    var scholarName = rawData[i][5] ? rawData[i][5].toString().trim() : "";
    
    // STRICT TIME WINDOW: Check if transaction is within the rolling 7-day window
    var transactionDate = new Date(timestamp);
    if (isNaN(transactionDate.getTime()) || transactionDate < sevenDaysAgo) continue;
    if (amount <= 0) continue;
    
    // Track weekly highest donor
    if (amount > maxWeeklyDonation) {
      maxWeeklyDonation = amount;
      topWeeklyDonorName = donorName;
    }
    
    // Ignore non-scholar system labels
    var cleanScholar = scholarName.toLowerCase();
    if (cleanScholar === "general campaign" || cleanScholar === "campaign" || cleanScholar === "") continue;
    
    // Aggregate weekly individual balances
    if (!weeklyScholarTotals[scholarName]) {
      weeklyScholarTotals[scholarName] = 0.00;
    }
    weeklyScholarTotals[scholarName] += amount;
  }
  
  result.topWeeklyDonor = { name: topWeeklyDonorName, amount: maxWeeklyDonation };
  
  var sortedList = [];
  for (var key in weeklyScholarTotals) {
    if (weeklyScholarTotals[key] > 0) {
      sortedList.push({ name: key, amount: weeklyScholarTotals[key] });
      scholarsActiveThisWeek.push(key);
    }
  }
  sortedList.sort(function(a, b) { return b.amount - a.amount; });
  
  result.top5Rows = sortedList.slice(0, 5);
  result.activeScholarsList = scholarsActiveThisWeek.join(", ");
  
  return result;
}

// ==========================================
//  GOOGLE CHAT WEBHOOK SENDER (ROUTED)
// ==========================================
function sendGoogleChatNotification(text, webhookUrl) {
  var targetUrl = webhookUrl || CONFIG.URL_VOLUNTEER_ALERTS;
  UrlFetchApp.fetch(targetUrl, {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({"text": text})
  });
}
