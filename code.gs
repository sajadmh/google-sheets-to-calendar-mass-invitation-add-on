var InvitesPerBatch = 10;
var MaxTimePerBatch = 3 * 60 * 1000;

var dataResults = {
    "result": "continue",
    "executionTime": 0.0,
    "invitationsSent": 0,
    "success": 0,
    "skip": 0,
    "failure": 0,
    "nextEmail": 1
};

function onOpen(e) {
    var ui = SpreadsheetApp.getUi();

    ui.createMenu('► Invite Automation ◄')
        .addItem('Open Auto-Invite Program', 'openSideBar')
        .addToUi();
}

function openSideBar() {
    var html = HtmlService.createTemplateFromFile('Campaign') // Campaign.html
        .evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('Auto-Invitation Program') // Header of the sidebar
        .setWidth(300);

    var ui = SpreadsheetApp.getUi();
    ui.showSidebar(html);
}

function getEventIdFromLink(link) {
    const m = link.match(/eventedit\/(\w+)/);
    const sp = Utilities.newBlob(Utilities.base64Decode(m[1])).getDataAsString().split(" ");
    return sp[0];
}

function getEventsIds(cellValues) {
    var ids = [];
    for (var i = 0; i < cellValues.length; i++) {
        var link = cellValues[i][0];
        if (link == "")
            continue;
        else
            ids.push(getEventIdFromLink(link));
    }
    return ids;
}

function getEvents(calendarId, eventIds) {
    var events = []

    for (var i = 0; i < eventIds.length; i++) {
        var event = Calendar.Events.get(calendarId, eventIds[i]);
        events.push(event);
    }

    return events;
}

function getEvent(calendarId, eventId) {
    var event = Calendar.Events.get(calendarId, eventId);
    return event;
}

function updateEvent(calendarId, event, emailA, emailB) {
    console.log("Processing: " + event.summary);
    var attendees = event.attendees;
    if (attendees == null)
        attendees = [];
    var originalLength = attendees.length;
    console.log("Original Attendee Length")
    console.log(originalLength)

    var emailListA = emailA.split(", ")
    var emailListB = emailB.split(", ")

    for (var i = 0; i < emailListA.length; i++) {
        attendees.push({
            email: emailListA[i]
        });
    }

    for (var i = 0; i < emailListB.length; i++) {
        attendees.push({
            email: emailListB[i]
        });
    }
    console.log("Attendees Length after appending")
    console.log(attendees.length)
    var resource = {
        attendees: attendees
    };
    var args = {
        sendUpdates: "all"
    };
    if (attendees.length > originalLength) {
        Calendar.Events.patch(resource, calendarId, event.id, args);
        console.log("Finished with: " + event.summary + " with " + attendees.length + " emails invited");
    } else {
        console.log("No one to invite");
    }
}

const getUnique_ = array2d => [...new Set(array2d.flat())];

function getFirstUnInvitedCell(sheet) {
    var rng, sh, values, x;
    rng = sheet.getRange('D1:D')

    values = rng.getValues();
    x = 0;
    values.some(function(ca, i) {
        x = i;
        return ca[0] == "";
    });
    return x + 1;
}

function inviteGuests() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var settings = sheet.getSheetByName("Settings");
    var invitationList = sheet.getSheetByName("Email_List");
    var calendarId = settings.getRange(2, 10).getValue();

    var tempSheet = sheet.getSheetByName("Temp");
    var tempLastRow = tempSheet.getLastRow();

    if (tempLastRow == 0) {
        var cell = tempSheet.getRange("A1");
        cell.setFormula("=UNIQUE(Email_List!C6:C)");

        tempLastRow = tempSheet.getLastRow();

        var emailAFormula = [];
        var emailBFormula = [];
        for (var i = 0; i < tempLastRow; i++) {
            emailAFormula[i] = ['=JOIN(", " ,FILTER(Email_List!O6:O, Email_List!C6:C=A' + (i + 1).toString() + ', NOT(Email_List!O6:O = "")))'];
            emailBFormula[i] = ['=JOIN(", " ,FILTER(Email_List!P6:P, Email_List!C6:C=A' + (i + 1).toString() + ', NOT(Email_List!P6:P = "")))'];
        }

        // set the column values
        tempSheet.getRange(1, 2, tempLastRow, 1).setFormulas(emailAFormula);
        tempSheet.getRange(1, 3, tempLastRow, 1).setFormulas(emailBFormula);
    }

    var uniqueValues = tempSheet.getRange(1, 1, tempLastRow, 1).getValues();
    var emailListA = tempSheet.getRange(1, 2, tempLastRow, 1).getValues();
    var emailListB = tempSheet.getRange(1, 3, tempLastRow, 1).getValues();

    // remove null values
    for (var i = 0; i < uniqueValues.length; i++) {
        if (uniqueValues[i] == "") {
            uniqueValues.splice(i, 1);
            emailListA.splice(i, 1);
            emailListB.splice(i, 1);
        }
    }

    // remove labels that are not present in Settings
    var lastColumn = settings.getLastColumn();
    var lastRow = settings.getLastRow();
    var lookupLabelValues = settings.getRange(5, 1, 1, lastColumn).getValues();
    lookupLabelValues = lookupLabelValues[0];

    for (var i = 0; i < uniqueValues.length; i++) {
        var index = lookupLabelValues.indexOf(uniqueValues[i][0]);
        if (index < 0) {
            uniqueValues.splice(i, 1);
            emailListA.splice(i, 1);
            emailListB.splice(i, 1);
        }
    }
    console.log(uniqueValues);
    console.log(uniqueValues.length);

    // get LabelList
    var labelList = new Array();
    for (var i = 0; i < uniqueValues.length; i++) {
        var index = lookupLabelValues.indexOf(uniqueValues[i][0]);
        if (index > -1) {
            var labeltemp = settings.getRange(6, index + 1, lastRow, 1).getValues();
            labelList.push(labeltemp);
        }
    }

    // get event ids
    var labelIdsList = new Array();
    for (const labelValues of labelList) {
        var labelidtemp = getEventsIds(labelValues);
        labelIdsList.push(labelidtemp);
    }

    var eventTempSheet = sheet.getSheetByName("EventTemp");
    var rowCount = 1;
    if (eventTempSheet.getLastRow() == 0) {
        for (var i = 0; i < labelIdsList.length; i++) {
            for (var j = 0; j < labelIdsList[i].length; j++) {
                eventTempSheet.getRange(rowCount, 1, 1, 1).setValue(labelIdsList[i][j]);
                eventTempSheet.getRange(rowCount, 2, 1, 1).setValue(emailListA[i]);
                eventTempSheet.getRange(rowCount, 3, 1, 1).setValue(emailListB[i]);
                rowCount++;
            }
            rowCount++;
        }
    }

    var cellToInvite = getFirstUnInvitedCell(eventTempSheet);
    console.log(cellToInvite);
    var labelId = eventTempSheet.getRange(cellToInvite, 1, 1, 1).getValue();
    console.log(labelId);
    var emailAInvite = eventTempSheet.getRange(cellToInvite, 2, 1, 1).getValue();
    console.log(emailAInvite);
    var emailBInvite = eventTempSheet.getRange(cellToInvite, 3, 1, 1).getValue();
    console.log(emailBInvite);
    var eventtemp = getEvent(calendarId, labelId);
    console.log(eventtemp);
    updateEvent(calendarId, eventtemp, emailAInvite, emailBInvite);
    eventTempSheet.getRange(cellToInvite, 4, 1, 1).setValue("✅ Invited");
}
