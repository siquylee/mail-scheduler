import { SchedulerEntry, Recurrence } from "./Interfaces";

export function setProp(key: string, value: string): void {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty(key, value);
}

export function getProp(key: string): string {
    var userProperties = PropertiesService.getUserProperties();
    return userProperties.getProperty(key);
}

export function createTrigger(funcName: string, entry: SchedulerEntry): GoogleAppsScript.Script.Trigger {
    let trigger!: GoogleAppsScript.Script.Trigger;
    if (!entry.Timzone)
        entry.Timzone = Session.getScriptTimeZone();
    switch (entry.Mode) {
        case Recurrence.None:
            trigger = createOnceTrigger(funcName, entry);
            break;
        case Recurrence.Daily:
            trigger = createDailyTrigger(funcName, entry);
            break;
        case Recurrence.Weekly:
            trigger = createWeeklyTrigger(funcName, entry);
            break;
        case Recurrence.Monthly:
            trigger = createMonthlyTrigger(funcName, entry);
            break;
        default:
            break;
    }
    return trigger;
}

function createOnceTrigger(funcName: string, entry: SchedulerEntry): GoogleAppsScript.Script.Trigger {
    entry.SentOnDate!.setHours(entry.Hour, entry.Minute, 0, 0);
    return ScriptApp.newTrigger(funcName)
        .timeBased()
        .inTimezone(entry.Timzone!)
        .at(entry.SentOnDate!)
        .create();
}

function createDailyTrigger(funcName: string, entry: SchedulerEntry): GoogleAppsScript.Script.Trigger {
    return ScriptApp.newTrigger(funcName)
        .timeBased()
        .inTimezone(entry.Timzone!)
        .everyDays(1)
        .atHour(entry.Hour)
        .nearMinute(entry.Minute)
        .create();
}

function createWeeklyTrigger(funcName: string, entry: SchedulerEntry): GoogleAppsScript.Script.Trigger {
    return ScriptApp.newTrigger(funcName)
        .timeBased()
        .inTimezone(entry.Timzone!)
        .everyWeeks(1)
        .onWeekDay(entry.Weekday!)
        .atHour(entry.Hour)
        .nearMinute(entry.Minute)
        .create();
}

function createMonthlyTrigger(funcName: string, entry: SchedulerEntry): GoogleAppsScript.Script.Trigger {
    return ScriptApp.newTrigger(funcName)
        .timeBased()
        .inTimezone(entry.Timzone!)
        .onMonthDay(entry.Day!)
        .atHour(entry.Hour)
        .nearMinute(entry.Minute)
        .create();
}