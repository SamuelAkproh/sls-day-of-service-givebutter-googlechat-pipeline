// ==========================================
// 1. THE LIVE WEBHOOK TRAFFIC COP (UPDATED FOR TICKETS/VOLUNTEERS)
// ==========================================
function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  var data = json.data;
  
  // Skip uncompleted transactions
  if (json.event === "transaction.success" && (!data.amount || parseFloat(data.amount) === 0) && !data.tickets) {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  var donorName = (data.first_name || "") + " " + (data.last_name || "");
  var donorEmail = data.email || "";
  var amount = (data.amount && !isNaN(data.amount)) ? parseFloat(data.amount) : 0.00;
  var transactionStatus = (data.status || "").toLowerCase().trim();
  
  var scholarName = "General Campaign";
  var scholarEmail = "";
  if (data.member) {
    scholarName = (data.member.first_name || "") + " " + (data.member.last_name || "");
    scholarEmail = (data.member.email || "").toString().toLowerCase().trim();
  }
  
  var atlantaEmails = ["samakproh@gmail.com", "amirion28@gmail.com", "birukgeremew05@gmail.com", "yanjohnny4@gmail.com", "bellamadison9108@gmail.com", "dawitkidaneyemane@gmail.com", "eaherndon07@gmail.com", "hassanb121286@gmail.com", "johnhana2006@gmail.com", "johnyilu60@gmail.com", "lajohnson0817@gmail.com", "mat.araujo.business@gmail.com", "mlhill1143@gmail.com", "napierceallen1@gmail.com", "naomitewodros6@gmail.com", "sutherlandnkosi@gmail.com", "ridhisaride01@gmail.com", "tmwilliams0924@gmail.com", "alicia@servantleaderscholars.org"];
  var athensEmails = ["kevincnwogu@gmail.com", "kevin@servantleaderscholars.org", "ameyah@servantleaderscholars.org", "angel@servantleaderscholars.org", "awesome@servantleaderscholars.org", "cayla@servantleaderscholars.org", "hadassah@servantleaderscholars.org", "india@servantleaderscholars.org", "jonina@servantleaderscholars.org", "kayanna@servantleaderscholars.org", "kayla@servantleaderscholars.org", "khiari@servantleaderscholars.org", "latrina@servantleaderscholars.org", "marc@servantleaderscholars.org", "marquesmckinney@servantleaderscholars.org", "marques@servantleaderscholars.org", "najma@servantleaderscholars.org", "nethen@servantleaderscholars.org", "nic@servantleaderscholars.org", "nyaboke@servantleaderscholars.org", "ugonna@servantleaderscholars.org"];
  
  var teamCohort = "General Campaign";
  if (atlantaEmails.indexOf(scholarEmail) !== -1) {
    teamCohort = "Atlanta";
  } else if (athensEmails.indexOf(scholarEmail) !== -1) {
    teamCohort = "Athens";
  } else {
    var cleanDonorEmail = donorEmail.toString().toLowerCase().trim();
    if (atlantaEmails.indexOf(cleanDonorEmail) !== -1) { teamCohort = "Atlanta"; scholarName = donorName; }
    else if (athensEmails.indexOf(cleanDonorEmail) !== -1) { teamCohort = "Athens"; scholarName = donorName; }
  }

  scholarName = getCleanScholarName(donorEmail, scholarName);

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var currentDate = new Date(); 
  var cleanIncomingEmail = donorEmail.toString().toLowerCase().trim();
  
  // Decide if this transaction is a donation or a volunteer ticket registration
  var isVolunteerRegistration = false;
  if (data.tickets && data.tickets.length > 0) {
    for (var j = 0; j < data.tickets.length; j++) {
      var ticketName = (data.tickets[j].name || "").toLowerCase();
      if (ticketName.includes("volunteer") || ticketName.includes("t-shirt") || ticketName.includes("admission")) {
        isVolunteerRegistration = true;
      }
    }
  }

  // EXCLUSION: If registration status is "revoked" or if the ticket is for a Banquet, skip volunteer log
  if (transactionStatus === "revoked") {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped revoked internal registration"})).setMimeType(ContentService.MimeType.JSON);
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
        return ContentService.createTextOutput(JSON.stringify({"status": "duplicate donation skipped"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    donorSheet.appendRow([currentDate, donorName, donorEmail, teamCohort, amount, scholarName]);
    sendGoogleChatNotification("💰 *NEW DONATION!* " + donorName + " just supported *" + scholarName + "* (" + teamCohort + ") with $" + amount.toFixed(2) + "! 🔥");
    
  } else if (isVolunteerRegistration) {
    // ---- PROCESS VOLUNTEER REGISTRATION (Includes paid $15 T-Shirt Volunteers) ----
    var volunteerSheet = ss.getSheetByName("Volunteers");
    var volunteerLastRow = volunteerSheet.getLastRow();
    if (volunteerLastRow > 1) {
      var volunteerEmails = volunteerSheet.getRange("C1:C" + volunteerLastRow).getValues().flat().map(function(eStr) {
        return eStr.toString().toLowerCase().trim();
      });
      if (volunteerEmails.indexOf(cleanIncomingEmail) !== -1) {
        return ContentService.createTextOutput(JSON.stringify({"status": "duplicate volunteer skipped"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    volunteerSheet.appendRow([currentDate, donorName, donorEmail, teamCohort]);
    
    var reportSheet = ss.getSheetByName("Weekly Report");
    var liveVolunteerCount = reportSheet.getRange("G2").getValue();
    
    var volunteerMessage = "🙋🏽‍♂️ *NEW VOLUNTEER REGISTERED!* \n\n" +
                           "Huge thank you to *" + donorName + "* for signing up to serve during our *2026 Day Of Service*! 💛🔥\n\n" +
                           "🚀 This takes our community total to *" + liveVolunteerCount + "* superstar volunteers mobilized!";
                               
    sendGoogleChatNotification(volunteerMessage);
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 📊 2. AUTOMATED FRIDAY RECAP SCRIPT (PRODUCTION ACTIVE)
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
  var contributorCount = leaderboard.contributors; 

  var advancedData = getAdvancedFundraisingData(ss);
  
  var top5String = "";
  var emojis = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣"];
  for (var i = 0; i < advancedData.top5Rows.length; i++) {
    top5String += emojis[i] + " " + advancedData.top5Rows[i].name + " — $" + advancedData.top5Rows[i].amount.toFixed(2) + "\n";
  }
  if (top5String === "") { top5String = "No individual donations registered yet! 🌱\n"; }

  var lastWeekTotal = totalRaised - weeklyDifference;
  var growthPercentageStr = "";
  if (lastWeekTotal > 0) {
    var pctIncrease = (weeklyDifference / lastWeekTotal) * 100;
    growthPercentageStr = " (a massive *" + pctIncrease.toFixed(1) + "%* explosion!)";
  }

  var wowSentence = "";
  if (weeklyDifference > 0) {
    wowSentence = "📈 *GROWTH MOMENTUM:* We pushed past last week's numbers by adding an incredible *$" + weeklyDifference.toFixed(2) + "* to our total this week alone" + growthPercentageStr + "!! 🎉\n";
  } else if (weeklyDifference === 0 && totalRaised > 0) {
    wowSentence = "✊ *HOLDING STRONG:* We are maintaining our steady momentum from last week! Let's push for a big breakout weekend! 🚀\n";
  }

  var milestoneCelebration = "";
  if (totalRaised >= 1500.00) {
    milestoneCelebration = "🎉🔥 *MAJOR MILESTONE ALERT:* We have officially shattered the *$1,500* mark together! Huge shoutout to the entire SLS family for locking in and pushing this movement forward. The momentum is unstoppable! 🙌🏽💙\n\n";
  }

  var mvpAtlantaName = advancedData.mvpAtlanta.name;
  var mvpAtlantaAmt = advancedData.mvpAtlanta.amount;
  var mvpAthensName = advancedData.mvpAthens.name;
  var mvpAthensAmt = advancedData.mvpAthens.amount;

  var leadershipShoutout = "👑 *LEADERSHIP SPOTLIGHT:*\n" +
                           "⚡ Massive respect to *Mckenzie Hill* ($" + mvpAthensAmt.toFixed(2) + ") and *Samuel Akproh* ($" + mvpAtlantaAmt.toFixed(2) + ") for absolutely pacing the field and anchoring both of our cohorts this week! You two are setting the standard! 🏆🔥\n\n";

  var message = "Good evening SLS Family 👋🏽💙! This is your favorite Bot, *Optimus Fine*, the finest bot on the internet 🤖😎! (ChatGPT got a secret crush on me y'all😉) \n\n" +
                "I hope y'all had an incredible week! 😄 \n\n" +
                "💻 *Microsoft Workshop Snapshot:*\n" +
                "Our session at Microsoft today was absolutely amazing! Shoutout to all the scholars who locked in and learned with us. Let's redirect that inspiration straight toward crushing our service campaign! ⚡🚀\n\n" +
                "Let's take a look at this week's recap for our 2026 DOS Campaign:\n\n" +
                "🚨 *SLS DAY OF SERVICE: WEEKLY SNAPSHOT* 🚨\n" +
                "━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
                milestoneCelebration +
                "💰 *TOTAL FUNDS RAISED:* *$" + totalRaised.toFixed(2) + "* / $" + CONFIG.GOAL_AMOUNT.toFixed(2) + " 🚀\n\n" +
                "📈 *PROGRESS TO GOAL:* We are officially *" + progressPercentage.toFixed(1) + "%* of the way to our $5,000 finish line 🙌🏽!\n\n" +
                "🙋🏽‍♂️ *VOLUNTEERS MOBILIZED:* *" + totalVolunteers + "* leaders signed up to serve!\n\n" +
                wowSentence + "\n" + 
                "🏆 *COHORT LEADERBOARD TOPPER:*\n" +
                "👉 *" + winningTeam + "* is officially leading the race with *$" + winningAmount.toFixed(2) + "* raised from *" + contributorCount + "* contributors! 🔥⚔️\n\n" +
                "🎖️ *COHORT MVP SHOUTOUTS:*\n" +
                "🍑 *Atlanta Top Supporter:* *" + mvpAtlantaName + "* ($" + mvpAtlantaAmt.toFixed(2) + " who else but HIM? 🤷‍♂️ )\n" +
                "🏛️ *Athens Top Supporter:* *" + mvpAthensName + "* ($" + mvpAthensAmt.toFixed(2) + ")\n\n" +
                leadershipShoutout +
                "💎 *INDIVIDUAL TOP 5 LEADERBOARD:*\n" +
                "━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
                top5String + 
                "━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
                "👑 *CAMPAIGN HIGHEST DONOR SHOUTOUT:*\n" +
                "🥇 Huge appreciation to *" + advancedData.topDonor.name + "* for anchoring our campaign with a massive single donation of *$" + advancedData.topDonor.amount.toFixed(2) + "*! You are a game-changer! 🙌🏽💎\n\n" +
                "👏 *ACTIVE FUNDRAISERS:*\n" +
                "Special shoutouts to our scholars actively generating donations: *" + advancedData.activeScholarsList + "*! Let's keep spreading the word and pushing this momentum forward over the weekend. All your hard work will reward us in the end! 🌟\n\n" +
                "Have a beautiful, restful weekend ahead! 💛💙";

  sendGoogleChatNotification(message);
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
  
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd");
  var weekLabel = "Week ending " + formattedDate;
  
  snapshotSheet.appendRow([weekLabel, totalRaised, totalVolunteers, today]);
  console.log("Snapshot archived for " + weekLabel);
}

// ==========================================
// 🔍 HELPER FUNCTION: DYNAMIC PIVOT TABLE SCANNER
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
    if (team === "" || team === "0" || team.toLowerCase().includes("grand total") || team.toLowerCase().includes("team / member")) continue;
    
    if (amount > maxAmount) {
      maxAmount = amount; topTeam = team; topContributors = count;
    }
  }
  return topTeam !== "" ? { teamName: topTeam, amountRaised: maxAmount, contributors: topContributors } : defaultResult;
}

// ==========================================
//  HELPER FUNCTION: PRECISION HISTORICAL MAP & MVP ANALYSIS
// ==========================================
function getAdvancedFundraisingData(ss) {
  var donorSheet = ss.getSheetByName("Donors");
  var lastRow = donorSheet.getLastRow();
  
  var result = {
    top5Rows: [],
    mvpAtlanta: { name: "No donations yet", amount: 0.00 },
    mvpAthens: { name: "No donations yet", amount: 0.00 },
    topDonor: { name: "No donors yet", amount: 0.00 },
    activeScholarsList: ""
  };
  
  if (lastRow < 2) return result;
  
  var rawData = donorSheet.getRange("A2:F" + lastRow).getValues();
  var individualMap = {};
  
  var maxAtlanta = 0;
  var maxAthens = 0;
  var absoluteMaxDonation = 0;
  var topDonorName = "";
  
  var authorizedScholars = [
    "India Knight", "Naomi Tewodros", "Samuel Akproh", "Kevin Nwogu", 
    "Marc Lewis", "Lisa Alexander", "Awesome David", "Mckenzie Hill", "Chelsea Jester"
  ];
  
  for (var i = 0; i < rawData.length; i++) {
    var donorName = rawData[i][1] ? rawData[i][1].toString().trim() : "";
    var donorEmail = rawData[i][2] ? rawData[i][2].toString().trim().toLowerCase() : "";
    var team = rawData[i][3] ? rawData[i][3].toString().trim().toLowerCase() : "";
    var amount = rawData[i][4] && !isNaN(rawData[i][4]) ? parseFloat(rawData[i][4]) : 0.00;
    var columnFScholar = rawData[i][5] ? rawData[i][5].toString().trim() : "";
    
    if (amount <= 0) continue;
    
    if (amount > absoluteMaxDonation) {
      absoluteMaxDonation = amount;
      topDonorName = donorName;
    }
    
    var scholarName = "";
    if (columnFScholar && columnFScholar !== "" && columnFScholar !== "General Campaign") {
      scholarName = getCleanScholarName(donorEmail, columnFScholar);
    } else {
      scholarName = getCleanScholarName(donorEmail, donorName);
    }
    
    if (scholarName === "General Campaign" || scholarName === "") continue;
    
    if (scholarName === "Mckenzie Hill") {
      team = "athens";
    } else if (scholarName === "Chelsea Jester") {
      team = "general campaign";
    }
    
    if (!individualMap[scholarName]) {
      individualMap[scholarName] = 0.00;
    }
    individualMap[scholarName] += amount;
    
    if (team === "atlanta" && individualMap[scholarName] > maxAtlanta) {
      maxAtlanta = individualMap[scholarName];
      result.mvpAtlanta = { name: scholarName, amount: individualMap[scholarName] };
    }
    if (team === "athens" && individualMap[scholarName] > maxAthens) {
      maxAthens = individualMap[scholarName];
      result.mvpAthens = { name: scholarName, amount: individualMap[scholarName] };
    }
  }
  
  result.topDonor = { name: topDonorName, amount: absoluteMaxDonation };
  
  var sortedList = [];
  var scholarsActive = [];
  
  for (var key in individualMap) {
    if (authorizedScholars.indexOf(key) !== -1 && individualMap[key] > 0) {
      sortedList.push({ name: key, amount: individualMap[key] });
      scholarsActive.push(key);
    }
  }
  sortedList.sort(function(a, b) { return b.amount - a.amount; });
  
  result.top5Rows = sortedList.slice(0, 5);
  result.activeScholarsList = scholarsActive.join(", ");
  
  return result;
}

// ==========================================
//  EMAIL-TO-SCHOLAR RESOLVER
// ==========================================
function getCleanScholarName(email, fallbackName) {
  var cleanEmail = email.toString().toLowerCase().trim();
  
  if (cleanEmail === "kevincnwogu@gmail.com" || cleanEmail === "mattgreene95@gmail.com") {
    return "Kevin Nwogu";
  } else if (cleanEmail === "marc@servantleaderscholars.org" || cleanEmail === "dxfdjr@gmail.com" || cleanEmail === "hbj2392@gmail.com" || cleanEmail === "elkan499@gmail.com" || cleanEmail === "acdaniel4103@gmail.com") {
    return "Marc Lewis";
  } else if (cleanEmail === "nardosaraya@hotmail.com" || cleanEmail === "naomitewodros6@gmail.com") {
    return "Naomi Tewodros";
  } else if (cleanEmail === "dakproh@gmail.com" || cleanEmail === "rwoodrum1@gmail.com") {
    return "Samuel Akproh";
  } else if (cleanEmail === "alicia@servantleaderscholars.org") {
    return "Lisa Alexander";
  } else if (cleanEmail === "cydneyjjohnson@gmail.com") {
    return "India Knight";
  } else if (cleanEmail === "awesomedav99@gmail.com") {
    return "Awesome David";
  } else if (cleanEmail === "ymwalker1143@gmail.com" || cleanEmail === "mlhill1143@gmail.com" || cleanEmail === "ambestseller1108@gmail.com" || cleanEmail === "accounts@ismllc-engr.com" || cleanEmail === "aubern.marshall@yahoo.com") {
    return "Mckenzie Hill";
  } else if (cleanEmail === "nikidinell@gmail.com" || cleanEmail === "vtowns108@gmail.com" || cleanEmail === "townsl@bellsouth.net") {
    return "Chelsea Jester";
  }
  
  return fallbackName;
}

// ==========================================
//  GOOGLE CHAT WEBHOOK SENDER (PRODUCTION LIVE!)
// ==========================================
function sendGoogleChatNotification(text) {
  UrlFetchApp.fetch(CONFIG.URL_MAIN_CHAT, {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({"text": text})
  });
}
