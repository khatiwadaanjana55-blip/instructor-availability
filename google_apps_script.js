// ================================================================
// GOOGLE APPS SCRIPT — Instructor Availability System Backend
// ================================================================

const SHEET_ID = "1MrnQP3AdckIITQEFT5XKPDV86XKITh-XxfeRX-CTImo";
const TERM = "Fall 2025";
const DEADLINE = "2025-08-15";
const ADMIN_EMAIL = "khatiwadaanjana55@gmail.com";
const FORM_BASE_URL =
  "https://shubhamrajpandey.github.io/instructor-availability/instructor_response_form.html";

const TABS = {
  INSTRUCTORS: "Instructors",
  RESPONSES: "Responses",
  TOKENS: "Tokens",
};

function doGet(e) {
  try {
    const action = e.parameter.action || "";
    if (action === "getData") return jsonOK(getAllData());
    if (action === "getInstructor")
      return jsonOK({ data: getInstructorByToken(e.parameter.token) });
    return jsonOK({ message: "API running OK" });
  } catch (err) {
    return jsonError(err.message);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || "submitResponse";
    if (action === "submitResponse") return jsonOK(submitResponse(body));
    if (action === "importInstructors")
      return jsonOK(importInstructors(body.data));
    if (action === "clearSheet") return jsonOK(clearSheetData(body.tab));
    if (action === "deleteInstructor")
      return jsonOK(deleteInstructorRow(body.email));
    return jsonError("Unknown action");
  } catch (err) {
    return jsonError(err.message);
  }
}

function jsonOK(data) {
  return ContentService.createTextOutput(
    JSON.stringify({ success: true, ...data }),
  ).setMimeType(ContentService.MimeType.JSON);
}
function jsonError(msg) {
  return ContentService.createTextOutput(
    JSON.stringify({ success: false, error: msg }),
  ).setMimeType(ContentService.MimeType.JSON);
}

function getAllData() {
  const ss = openSheet();
  setupSheets(ss);

  const instrRows = ss
    .getSheetByName(TABS.INSTRUCTORS)
    .getDataRange()
    .getValues()
    .slice(1);
  const respRows = ss
    .getSheetByName(TABS.RESPONSES)
    .getDataRange()
    .getValues()
    .slice(1);

  const COLORS = [
    { bg: "#E6F1FB", fg: "#0C447C" },
    { bg: "#FAEEDA", fg: "#633806" },
    { bg: "#EAF3DE", fg: "#27500A" },
    { bg: "#FBEAF0", fg: "#72243E" },
    { bg: "#E1F5EE", fg: "#085041" },
    { bg: "#EEEDFE", fg: "#3C3489" },
    { bg: "#FAECE7", fg: "#712B13" },
    { bg: "#F1EFE8", fg: "#444441" },
  ];

  const instructors = instrRows
    .filter((r) => r[0])
    .map((r, i) => {
      const clr = COLORS[i % COLORS.length];
      const name = String(r[0]);
      return {
        name: name,
        email: String(r[1] || ""),
        dept: String(r[2] || ""),
        courses: String(r[3] || ""),
        status: String(r[4] || "pending").toLowerCase(),
        token: String(r[5] || ""),
        contact: r[6] ? new Date(r[6]).toLocaleDateString("en-US") : "—",
        initials: name
          .split(" ")
          .map((w) => w[0])
          .filter((_, j, a) => j === 0 || j === a.length - 1)
          .join(""),
        bg: clr.bg,
        fg: clr.fg,
      };
    });

  // ✅ FIX 1: dates in en-US, ✅ FIX 2: availability field added
  const responses = respRows
    .filter((r) => r[0])
    .map((r) => ({
      time: r[0] ? new Date(r[0]).toLocaleString("en-US") : "—",
      token: String(r[1] || ""),
      name: String(r[2] || ""),
      email: String(r[3] || ""),
      dept: String(r[4] || ""),
      term: String(r[5] || ""),
      courseCode: String(r[6] || ""),
      courseName: String(r[7] || ""),
      availability: String(r[8] || "pending").toLowerCase(),
      status: String(r[8] || "pending").toLowerCase(),
      comment: String(r[9] || ""),
    }));

  return { instructors, responses };
}

function getInstructorByToken(token) {
  if (!token) return null;
  const ss = openSheet();
  setupSheets(ss);
  const sheet = ss.getSheetByName(TABS.TOKENS);
  const rows = sheet.getDataRange().getValues().slice(1);

  for (const row of rows) {
    if (
      String(row[0]).trim() === String(token).trim() &&
      String(row[5]).toLowerCase() !== "used"
    ) {
      let courses = [];
      try {
        courses = JSON.parse(row[6] || "[]");
      } catch (e) {}
      return {
        name: String(row[1]),
        email: String(row[2]),
        dept: String(row[3]),
        term: String(row[4]),
        courses: courses,
      };
    }
  }
  return null;
}

function submitResponse(data) {
  const ss = openSheet();
  setupSheets(ss);
  const sheet = ss.getSheetByName(TABS.RESPONSES);
  const ts = new Date(data.timestamp || new Date());

  (data.courses || []).forEach((course) => {
    sheet.appendRow([
      ts,
      data.token,
      data.instructorName,
      data.email,
      data.dept,
      data.term,
      course.code,
      course.name,
      course.availability,
      data.comments || "",
    ]);
  });

  // ✅ FIX 3: Correctly set unavailable when all courses are no
  const allUnavailable = (data.courses || []).every(
    (c) => c.availability === "no",
  );
  const finalStatus = allUnavailable ? "unavailable" : "available";
  updateInstructorStatus(ss, data.instructorName, finalStatus, ts);
  markTokenUsed(ss, data.token);
  return { message: "Response recorded successfully" };
}

function importInstructors(rows) {
  if (!rows || !rows.length) throw new Error("No data provided");
  const ss = openSheet();
  setupSheets(ss);
  const sheet = ss.getSheetByName(TABS.INSTRUCTORS);
  const existing = sheet.getDataRange().getValues().slice(1);

  rows.forEach((r) => {
    const idx = existing.findIndex((row) => row[1] === r.email);
    if (idx >= 0) {
      sheet
        .getRange(idx + 2, 1, 1, 4)
        .setValues([[r.name, r.email, r.dept, r.courses]]);
    } else {
      const token = "TOK-" + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([
        r.name,
        r.email,
        r.dept,
        r.courses,
        "pending",
        token,
        new Date(),
      ]);
    }
  });

  fixTokensNow();
  return { imported: rows.length };
}

function openSheet() {
  try {
    return SpreadsheetApp.openById(SHEET_ID);
  } catch (e) {
    throw new Error(
      "Cannot open spreadsheet. Check SHEET_ID. Error: " + e.message,
    );
  }
}

function setupSheets(ss) {
  if (!ss.getSheetByName(TABS.INSTRUCTORS)) {
    const s = ss.insertSheet(TABS.INSTRUCTORS);
    s.appendRow([
      "Name",
      "Email",
      "Department",
      "Courses",
      "Status",
      "Token",
      "Last Contact",
    ]);
    s.getRange(1, 1, 1, 7)
      .setFontWeight("bold")
      .setBackground("#185FA5")
      .setFontColor("white");
    s.setFrozenRows(1);
    s.setColumnWidth(1, 180);
    s.setColumnWidth(2, 220);
    s.setColumnWidth(3, 160);
    s.setColumnWidth(4, 160);
    s.setColumnWidth(5, 100);
    s.setColumnWidth(6, 130);
  }
  if (!ss.getSheetByName(TABS.RESPONSES)) {
    const s = ss.insertSheet(TABS.RESPONSES);
    s.appendRow([
      "Timestamp",
      "Token",
      "Instructor Name",
      "Email",
      "Department",
      "Term",
      "Course Code",
      "Course Name",
      "Availability",
      "Comments",
    ]);
    s.getRange(1, 1, 1, 10)
      .setFontWeight("bold")
      .setBackground("#185FA5")
      .setFontColor("white");
    s.setFrozenRows(1);
  }
  if (!ss.getSheetByName(TABS.TOKENS)) {
    const s = ss.insertSheet(TABS.TOKENS);
    s.appendRow([
      "Token",
      "Name",
      "Email",
      "Dept",
      "Term",
      "Status",
      "Courses JSON",
    ]);
    s.getRange(1, 1, 1, 7)
      .setFontWeight("bold")
      .setBackground("#185FA5")
      .setFontColor("white");
    s.setFrozenRows(1);
    s.setColumnWidth(1, 130);
    s.setColumnWidth(7, 400);
  }
}

function updateInstructorStatus(ss, name, status, timestamp) {
  const sheet = ss.getSheetByName(TABS.INSTRUCTORS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === name) {
      sheet.getRange(i + 1, 5).setValue(status);
      sheet.getRange(i + 1, 7).setValue(timestamp);
      break;
    }
  }
}

function markTokenUsed(ss, token) {
  const sheet = ss.getSheetByName(TABS.TOKENS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(token).trim()) {
      sheet.getRange(i + 1, 6).setValue("used");
      break;
    }
  }
}

function deleteInstructorRow(email) {
  const ss = openSheet();
  const instrSheet = ss.getSheetByName("Instructors");
  const tokenSheet = ss.getSheetByName("Tokens");
  const rows = instrSheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(email).trim()) {
      var token = String(rows[i][5]).trim();
      instrSheet.deleteRow(i + 1);
      Logger.log("Deleted instructor: " + email);

      var tokenRows = tokenSheet.getDataRange().getValues();
      for (var j = 1; j < tokenRows.length; j++) {
        if (
          String(tokenRows[j][0]).trim() === token ||
          String(tokenRows[j][2]).trim() === email
        ) {
          tokenSheet.deleteRow(j + 1);
          Logger.log("Deleted token for: " + email);
          break;
        }
      }
      return { deleted: email };
    }
  }
  return { error: "Instructor not found" };
}

function clearSheetData(tab) {
  const ss = openSheet();
  const tabs = tab === "All" ? ["Instructors", "Responses", "Tokens"] : [tab];

  tabs.forEach(function (tabName) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
    Logger.log("Cleared: " + tabName);
  });

  return { cleared: tabs };
}

function fixTokensNow() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const instrSheet = ss.getSheetByName("Instructors");
  const tokenSheet = ss.getSheetByName("Tokens");

  const lastRow = tokenSheet.getLastRow();
  if (lastRow > 1) {
    tokenSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
  Logger.log("Cleared old tokens");

  const instrRows = instrSheet
    .getDataRange()
    .getValues()
    .slice(1)
    .filter((r) => r[0]);
  Logger.log("Found " + instrRows.length + " instructors");

  instrRows.forEach(function (row) {
    var name = String(row[0]).trim();
    var email = String(row[1]).trim();
    var dept = String(row[2]).trim();
    var courses = String(row[3]).trim();
    var token = String(row[5]).trim();

    if (!token) {
      Logger.log("No token for: " + name + " — skipping");
      return;
    }

    var courseArr = courses.split(",").map(function (c) {
      return {
        code: c.trim(),
        section: "SEC-A",
        name: c.trim(),
        time: "See schedule",
        room: "TBD",
      };
    });

    tokenSheet.appendRow([
      token,
      name,
      email,
      dept,
      "Fall 2025",
      "active",
      JSON.stringify(courseArr),
    ]);

    Logger.log("Added: " + name + " → " + token);
  });

  Logger.log(
    "DONE! Tokens tab is ready. Total: " + instrRows.length + " tokens.",
  );
}

function initialSetup() {
  const ss = openSheet();
  setupSheets(ss);
  fixTokensNow();
  Logger.log("Setup complete! Sheets and tokens are ready.");
}

function sendReminderEmails() {
  const ss = openSheet();
  const sheet = ss.getSheetByName(TABS.INSTRUCTORS);
  if (!sheet) {
    Logger.log("No Instructors sheet");
    return;
  }

  const rows = sheet.getDataRange().getValues().slice(1);
  let sent = 0;

  rows.forEach(function (row) {
    var name = row[0];
    var email = row[1];
    var status = row[4];
    var token = row[5];
    if (String(status).toLowerCase() === "pending" && token && email) {
      var url = FORM_BASE_URL + "?token=" + token;
      var html = buildReminderEmail(name, url);
      try {
        GmailApp.sendEmail(
          email,
          "REMINDER: Course availability - " + TERM,
          "",
          {
            htmlBody: html,
            name: "University Scheduling System",
          },
        );
        sent++;
        Logger.log("Reminder sent to " + email);
      } catch (e) {
        Logger.log("Failed: " + email + " — " + e.message);
      }
    }
  });
  Logger.log("Sent " + sent + " reminder emails.");
  return sent;
}

function sendEscalationNotification() {
  var today = new Date();
  var deadline = new Date(DEADLINE);
  if (today < deadline) {
    Logger.log("Deadline not passed yet.");
    return;
  }

  var ss = openSheet();
  var sheet = ss.getSheetByName(TABS.INSTRUCTORS);
  var pending = sheet
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(function (r) {
      return String(r[4]).toLowerCase() === "pending" && r[0];
    })
    .map(function (r) {
      return { name: r[0], email: r[1], dept: r[2], courses: r[3] };
    });

  if (!pending.length) {
    Logger.log("All responded!");
    return;
  }

  var rows = pending
    .map(function (p) {
      return (
        "<tr>" +
        '<td style="padding:8px 12px;border-bottom:1px solid #f4f3ef">' +
        p.name +
        "</td>" +
        '<td style="padding:8px 12px;border-bottom:1px solid #f4f3ef">' +
        p.email +
        "</td>" +
        '<td style="padding:8px 12px;border-bottom:1px solid #f4f3ef">' +
        p.courses +
        "</td>" +
        "</tr>"
      );
    })
    .join("");

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px">' +
    '<div style="background:#A32D2D;padding:20px 28px;border-radius:8px 8px 0 0">' +
    '<div style="color:white;font-size:18px;font-weight:600">Escalation Alert</div></div>' +
    '<div style="background:white;padding:28px;border:1px solid #e4e2da;border-radius:0 0 8px 8px">' +
    '<p style="font-weight:600">Deadline Passed - ' +
    pending.length +
    " Instructor(s) Have Not Responded</p>" +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:12px">' +
    '<thead><tr style="background:#f4f3ef">' +
    '<th style="padding:8px 12px;text-align:left">Name</th>' +
    '<th style="padding:8px 12px;text-align:left">Email</th>' +
    '<th style="padding:8px 12px;text-align:left">Courses</th>' +
    "</tr></thead><tbody>" +
    rows +
    "</tbody></table>" +
    '<p style="margin-top:16px;font-size:11px;color:#a09e98">University Scheduling System</p>' +
    "</div></div>";

  GmailApp.sendEmail(
    ADMIN_EMAIL,
    "ESCALATION: " +
      pending.length +
      " instructors have not responded - " +
      TERM,
    "",
    { htmlBody: html },
  );
  Logger.log("Escalation sent for " + pending.length + " instructors.");
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendReminderEmails")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  ScriptApp.newTrigger("sendEscalationNotification")
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
  Logger.log("Triggers set up successfully!");
}

function buildReminderEmail(name, formUrl) {
  return (
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">' +
    '<div style="background:#BA7517;padding:20px 28px;border-radius:8px 8px 0 0">' +
    '<div style="color:white;font-size:18px;font-weight:600">Reminder: Course Availability</div>' +
    '<div style="color:#faeeda;font-size:12px;margin-top:4px">University Scheduling System - ' +
    TERM +
    "</div>" +
    "</div>" +
    '<div style="background:white;padding:28px;border-radius:0 0 8px 8px;border:1px solid #e4e2da;border-top:none">' +
    '<p style="font-size:16px;font-weight:600">Dear ' +
    name +
    ",</p>" +
    '<p style="color:#6b6860;font-size:14px;line-height:1.6;margin:12px 0">' +
    "We have not yet received your availability for <strong>" +
    TERM +
    "</strong>. Deadline: <strong>" +
    DEADLINE +
    "</strong>." +
    "</p>" +
    '<a href="' +
    formUrl +
    '" style="display:block;text-align:center;background:#185FA5;color:white;text-decoration:none;padding:13px;border-radius:6px;font-size:14px;font-weight:600;margin-top:20px">' +
    "Submit My Availability</a>" +
    '<p style="margin-top:12px;font-size:12px;text-align:center;color:#6b6860">' +
    'Or copy this link: <a href="' +
    formUrl +
    '" style="color:#185FA5;word-break:break-all">' +
    formUrl +
    "</a></p>" +
    '<p style="margin-top:16px;font-size:11px;color:#a09e98">Do not reply to this email.</p>' +
    "</div></div>"
  );
}
