import { SchedulerAddress, Recurrence } from "./interfaces";
import { createTrigger, readEntry, readEntryByRow, shouldStart, shouldEnd, getNextRun, deleteTrigger } from "./utils";
import { l } from "./localization";

function onOpen(e: any) {
    try {
        var spreadsheet = SpreadsheetApp.getActive();
        var menuItems = [
            { name: l('installScript'), functionName: 'installScript' },
        ];
        spreadsheet.addMenu(l('scriptName'), menuItems);
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
            Browser.msgBox(l('installFailure'));
        else
            Browser.msgBox(l('installSuccess'));
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
        // Every day, find & delete orphaned triggers
        ScriptApp.newTrigger(initializeHandler)
            .timeBased()
            .everyDays(1)
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
        // Check if we should start sending email
        if (!shouldStart(entry)) {
            Logger.log(`Well, not this time. The schedule is ${entry.StartDate}`);
            return;
        }

        // Let sending email
        let status: string = '';
        let nextRun = getNextRun(entry);
        try {
            let recipients = entry.UniqueMessage ? entry.To.split(',') : [entry.To];
            recipients.forEach(to => {
                MailApp.sendEmail(to, entry!.Subject, entry!.Message, { htmlBody: entry!.Message });
                Logger.log(`Mail '${entry!.Subject}' sent to ${to} successfully.`);
            });

            status = `OK. Next run: ${nextRun}`;
            // Check if we should stop sending email
            if (shouldEnd(entry)) {
                // Delete this row
                deleteTrigger(uid);
            }
        } catch (err) {
            Logger.log(err);
            status = `Error: ${err}. Next run: ${nextRun}`;
        }
        finally {
            // Write data back to status
            let sheet = SpreadsheetApp.getActive().getSheetByName(SchedulerAddress.Sheet);
            sheet.getRange(`${SchedulerAddress.Execution}${entry.RowNo}`).setValue(status);
        }
    }
    else {
        Logger.log(`Error: Data with ${uid} is invalid`);
    }
}
