import { SchedulerEntry, SchedulerAddress, Recurrence } from "./Interfaces";
import { createTrigger } from "./Utils";

function onOpen(e: any) {
    try {
        var spreadsheet = SpreadsheetApp.getActive();
        var menuItems = [
            { name: 'Install script', functionName: 'installScript' },
        ];
        spreadsheet.addMenu('Mail Scheduler', menuItems);
    }
    catch (err) {
        Browser.msgBox(err);
    }
}

function installScript(): void {
    try {
        // Copied document often lost triggers
        installScheduleSubmitTrigger();
        installCleanupTrigger();
        let authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
        if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED)
            Browser.msgBox('Could not install the script. Please try again!');
        else
            Browser.msgBox('Script installed successfully');
    }
    catch (err) {
        Browser.msgBox(err);
    }
}

function installScheduleSubmitTrigger(): void {
    const initializeHandler = 'onScheduleSubmit';
    var triggers = ScriptApp.getProjectTriggers();
    let found = triggers.filter(trigger => trigger.getHandlerFunction() == initializeHandler).length > 0;
    if (!found) {
        ScriptApp.newTrigger(initializeHandler)
            .forSpreadsheet(SpreadsheetApp.getActive())
            .onFormSubmit()
            .create();
        Logger.log(`New trigger created with handler ${initializeHandler}`);
    }
    else {
        Logger.log(`Trigger already existed with handler ${initializeHandler}`);
    }
}

function installCleanupTrigger(): void {
    const initializeHandler = 'onCleanup';
    var triggers = ScriptApp.getProjectTriggers();
    let found = triggers.filter(trigger => trigger.getHandlerFunction() == initializeHandler).length > 0;
    if (!found) {
        // Every 3 days, find & delete orphaned triggers
        ScriptApp.newTrigger(initializeHandler)
            .timeBased()
            .everyDays(3)
            .atHour(1)
            .create();
        Logger.log(`New trigger created with handler ${initializeHandler}`);
    } else {
        Logger.log(`Trigger already existed with handler ${initializeHandler}`);
    }
}

function onScheduleSubmit(): void {
    let sheet = SpreadsheetApp.getActiveSheet();
    let lastRow = sheet.getLastRow();
    let entry = readEntryByRow(lastRow);
    if (entry) {
        try {
            let trigger = createTrigger('onScheduleExecuted', entry);
            let uid = trigger.getUniqueId();
            sheet.getRange(`${SchedulerAddress.Uid}${lastRow}`).setValue(uid);
            Logger.log(`Trigger ${uid} created: ${entry.Subject}`);
        } catch (err) {
            Logger.log(err);
        }
    }
}

function onCleanup(): void {
    let excludedFunc = ['onScheduleSubmit', 'onCleanup']
    let triggers = ScriptApp.getProjectTriggers();
    let deleteItems: GoogleAppsScript.Script.Trigger[] = [];
    triggers.forEach(trigger => {
        let funcName = trigger.getHandlerFunction();
        if (excludedFunc.indexOf(funcName) < 0) {
            let uid = trigger.getUniqueId();
            let entry = readEntry(uid);
            if (!entry)
                deleteItems.push(trigger);
        }
    });
    Logger.log(`Deleting ${deleteItems.length} orphaned triggers`);
    deleteItems.forEach(item => {
        Logger.log(`Delete trigger uid: ${item.getUniqueId()}`);
        ScriptApp.deleteTrigger(item);
    })
}

function onScheduleExecuted(e: any): void {
    let uid = e.triggerUid;
    Logger.log(`Trigger ${uid} executed. Uid ${uid}`);
    let entry = readEntry(uid);
    if (entry) {
        try {
            MailApp.sendEmail(entry.To, entry.Subject, entry.Message, { htmlBody: entry.Message });
            Logger.log(`Mail '${entry.Subject}' sent to ${entry.To} successfully.`);
            if (entry.Mode == Recurrence.None) {
                var triggers = ScriptApp.getProjectTriggers();
                triggers.forEach(trigger => {
                    if (uid == trigger.getUniqueId())
                        ScriptApp.deleteTrigger(trigger);
                });
            }
        } catch (err) {
            Logger.log(err);
        }
    }
    else {
        Logger.log(`Error: Data with ${uid} is invalid`);
    }
}

function readEntry(uid: string): SchedulerEntry | null {
    try {
        let sheet = SpreadsheetApp.getActive().getSheetByName(SchedulerAddress.Sheet);
        let colIdx = sheet.getRange(`${SchedulerAddress.Uid}1`).getColumn();
        let columnValues = sheet.getRange(2, colIdx, sheet.getLastRow()).getValues();
        for (var i = 0; i < columnValues.length; i++) {
            if (columnValues[i][0] == uid) {
                // i + 2 is row index.
                return readEntryByRow(i + 2);
            }
        }
        return null;
    }
    catch (err) {
        Logger.log(err);
        return null;
    }
}

function readEntryByRow(row: number): SchedulerEntry | null {
    try {
        let sheet = SpreadsheetApp.getActive().getSheetByName(SchedulerAddress.Sheet);
        let entry = new SchedulerEntry();
        entry.To = sheet.getRange(`${SchedulerAddress.To}${row}`).getValue().toString();
        entry.Subject = sheet.getRange(`${SchedulerAddress.Subject}${row}`).getValue().toString();
        entry.Message = sheet.getRange(`${SchedulerAddress.Message}${row}`).getValue().toString();
        let mode = sheet.getRange(`${SchedulerAddress.Mode}${row}`).getValue().toString();
        switch (mode) {
            case 'None':
                entry.Mode = Recurrence.None;
                break;
            case 'Daily':
                entry.Mode = Recurrence.Daily;
                break;
            case 'Weekly':
                entry.Mode = Recurrence.Weekly;
                break;
            case 'Monthly':
                entry.Mode = Recurrence.Monthly;
                break;
            case 'Custom':
                entry.Mode = Recurrence.Custom;
                break;
        }

        let day = sheet.getRange(`${SchedulerAddress.Weekdays}${row}`).getValue().toString()
        switch (day.trim().toUpperCase()) {
            case 'MONDAY':
                entry.Weekday = ScriptApp.WeekDay.MONDAY;
                break;
            case 'TUESDAY':
                entry.Weekday = ScriptApp.WeekDay.TUESDAY;
                break;
            case 'WEDNESDAY':
                entry.Weekday = ScriptApp.WeekDay.WEDNESDAY;
                break;
            case 'THURSDAY':
                entry.Weekday = ScriptApp.WeekDay.THURSDAY;
                break;
            case 'FRIDAY':
                entry.Weekday = ScriptApp.WeekDay.FRIDAY;
                break;
            case 'SATURDAY':
                entry.Weekday = ScriptApp.WeekDay.SATURDAY;
                break;
            case 'SUNDAY':
                entry.Weekday = ScriptApp.WeekDay.SUNDAY;
                break;
        }
        let time = new Date(sheet.getRange(`${SchedulerAddress.SentOnTime}${row}`).getValue().toString());
        entry.Hour = time.getHours();
        entry.Minute = time.getMinutes();
        entry.SentOnDate = new Date(sheet.getRange(`${SchedulerAddress.SentOnDate}${row}`).getValue().toString());
        entry.Day = parseInt(sheet.getRange(`${SchedulerAddress.Day}${row}`).getValue().toString());
        entry.Timzone = sheet.getRange(`${SchedulerAddress.Timzone}${row}`).getValue().toString();
        entry.Uid = sheet.getRange(`${SchedulerAddress.Uid}${row}`).getValue().toString();
        return entry;
    }
    catch (err) {
        Logger.log(err);
        return null;
    }
}