import { SchedulerAddress, SchedulerEntry, Recurrence } from "./interfaces";
import { l } from "./localization";

export function readEntry(uid: string): SchedulerEntry | null {
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

export function readEntryByRow(row: number): SchedulerEntry | null {
    try {
        let sheet = SpreadsheetApp.getActive().getSheetByName(SchedulerAddress.Sheet);
        let entry = new SchedulerEntry();
        entry.To = sheet.getRange(`${SchedulerAddress.To}${row}`).getValue().toString();
        entry.Subject = sheet.getRange(`${SchedulerAddress.Subject}${row}`).getValue().toString();
        entry.Message = sheet.getRange(`${SchedulerAddress.Message}${row}`).getValue().toString();

        let mode = sheet.getRange(`${SchedulerAddress.Mode}${row}`).getValue().toString().trim();
        if (l('recurrence.None').trim() == mode)
            entry.Mode = Recurrence.None;
        else if (l('recurrence.Daily').trim() == mode)
            entry.Mode = Recurrence.Daily;
        else if (l('recurrence.Weekly').trim() == mode)
            entry.Mode = Recurrence.Weekly;
        else if (l('recurrence.Monthly').trim() == mode)
            entry.Mode = Recurrence.Monthly;
        else
            entry.Mode = Recurrence.Custom;

        let day = sheet.getRange(`${SchedulerAddress.Weekdays}${row}`).getValue().toString().trim().toUpperCase();
        if (l('weekDay.Monday').trim().toUpperCase() == day)
            entry.Weekday = ScriptApp.WeekDay.MONDAY;
        else if (l('weekDay.Tuesday').trim().toUpperCase() == day)
            entry.Weekday = ScriptApp.WeekDay.TUESDAY;
        else if (l('weekDay.Wednesday').trim().toUpperCase() == day)
            entry.Weekday = ScriptApp.WeekDay.WEDNESDAY;
        else if (l('weekDay.Thursday').trim().toUpperCase() == day)
            entry.Weekday = ScriptApp.WeekDay.THURSDAY;
        else if (l('weekDay.Friday').trim().toUpperCase() == day)
            entry.Weekday = ScriptApp.WeekDay.FRIDAY;
        else if (l('weekDay.Saturday').trim().toUpperCase() == day)
            entry.Weekday = ScriptApp.WeekDay.SATURDAY;
        else if (l('weekDay.Sunday').trim().toUpperCase() == day)
            entry.Weekday = ScriptApp.WeekDay.SUNDAY;

        let time = new Date(sheet.getRange(`${SchedulerAddress.SentOnTime}${row}`).getValue().toString());
        entry.Hour = time.getHours();
        entry.Minute = time.getMinutes();
        entry.SentOnDate = new Date(sheet.getRange(`${SchedulerAddress.SentOnDate}${row}`).getValue().toString());
        entry.Day = parseInt(sheet.getRange(`${SchedulerAddress.Day}${row}`).getValue().toString());
        entry.Timzone = getTimezoneByValue(sheet.getRange(`${SchedulerAddress.Timzone}${row}`).getValue().toString());
        entry.Uid = sheet.getRange(`${SchedulerAddress.Uid}${row}`).getValue().toString();
        return entry;
    }
    catch (err) {
        Logger.log(err);
        return null;
    }
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

function getTimezoneByValue(value: string): string {
    let timezones = [
        { key: 'Pacific/Midway', value: '(GMT-11:00) Midway' },
        { key: 'Pacific/Niue', value: '(GMT-11:00) Niue' },
        { key: 'Pacific/Pago_Pago', value: '(GMT-11:00) Pago Pago' },
        { key: 'Pacific/Honolulu', value: '(GMT-10:00) Hawaii Time' },
        { key: 'Pacific/Johnston', value: '(GMT-10:00) Johnston' },
        { key: 'Pacific/Rarotonga', value: '(GMT-10:00) Rarotonga' },
        { key: 'Pacific/Tahiti', value: '(GMT-10:00) Tahiti' },
        { key: 'Pacific/Marquesas', value: '(GMT-09:30) Marquesas' },
        { key: 'America/Anchorage', value: '(GMT-09:00) Alaska Time' },
        { key: 'Pacific/Gambier', value: '(GMT-09:00) Gambier' },
        { key: 'America/Dawson', value: '(GMT-08:00) Dawson' },
        { key: 'America/Los_Angeles', value: '(GMT-08:00) Pacific Time' },
        { key: 'America/Tijuana', value: '(GMT-08:00) Pacific Time - Tijuana' },
        { key: 'America/Vancouver', value: '(GMT-08:00) Pacific Time - Vancouver' },
        { key: 'America/Whitehorse', value: '(GMT-08:00) Pacific Time - Whitehorse' },
        { key: 'Pacific/Pitcairn', value: '(GMT-08:00) Pitcairn' },
        { key: 'America/Boise', value: '(GMT-07:00) Boise' },
        { key: 'America/Denver', value: '(GMT-07:00) Mountain Time' },
        { key: 'America/Phoenix', value: '(GMT-07:00) Mountain Time - Arizona' },
        { key: 'America/Mazatlan', value: '(GMT-07:00) Mountain Time - Chihuahua, Mazatlan' },
        { key: 'America/Dawson_Creek', value: '(GMT-07:00) Mountain Time - Dawson Creek' },
        { key: 'America/Edmonton', value: '(GMT-07:00) Mountain Time - Edmonton' },
        { key: 'America/Hermosillo', value: '(GMT-07:00) Mountain Time - Hermosillo' },
        { key: 'America/Yellowknife', value: '(GMT-07:00) Mountain Time - Yellowknife' },
        { key: 'America/Belize', value: '(GMT-06:00) Belize' },
        { key: 'America/Chicago', value: '(GMT-06:00) Central Time' },
        { key: 'America/Mexico_City', value: '(GMT-06:00) Central Time - Mexico City' },
        { key: 'America/Regina', value: '(GMT-06:00) Central Time - Regina' },
        { key: 'America/Tegucigalpa', value: '(GMT-06:00) Central Time - Tegucigalpa' },
        { key: 'America/Winnipeg', value: '(GMT-06:00) Central Time - Winnipeg' },
        { key: 'America/Costa_Rica', value: '(GMT-06:00) Costa Rica' },
        { key: 'Pacific/Easter', value: '(GMT-06:00) Easter Island' },
        { key: 'America/El_Salvador', value: '(GMT-06:00) El Salvador' },
        { key: 'Pacific/Galapagos', value: '(GMT-06:00) Galapagos' },
        { key: 'America/Guatemala', value: '(GMT-06:00) Guatemala' },
        { key: 'America/Managua', value: '(GMT-06:00) Managua' },
        { key: 'America/Cancun', value: '(GMT-05:00) America Cancun' },
        { key: 'America/Bogota', value: '(GMT-05:00) Bogota' },
        { key: 'America/Cayman', value: '(GMT-05:00) Cayman' },
        { key: 'America/Detroit', value: '(GMT-05:00) Detroit' },
        { key: 'America/New_York', value: '(GMT-05:00) Eastern Time' },
        { key: 'America/Iqaluit', value: '(GMT-05:00) Eastern Time - Iqaluit' },
        { key: 'America/Montreal', value: '(GMT-05:00) Eastern Time - Montreal' },
        { key: 'America/Toronto', value: '(GMT-05:00) Eastern Time - Toronto' },
        { key: 'America/Grand_Turk', value: '(GMT-05:00) Grand Turk' },
        { key: 'America/Guayaquil', value: '(GMT-05:00) Guayaquil' },
        { key: 'America/Havana', value: '(GMT-05:00) Havana' },
        { key: 'America/Jamaica', value: '(GMT-05:00) Jamaica' },
        { key: 'America/Lima', value: '(GMT-05:00) Lima' },
        { key: 'America/Nassau', value: '(GMT-05:00) Nassau' },
        { key: 'America/Panama', value: '(GMT-05:00) Panama' },
        { key: 'America/Port-au-Prince', value: '(GMT-05:00) Port-au-Prince' },
        { key: 'America/Rio_Branco', value: '(GMT-05:00) Rio Branco' },
        { key: 'America/Anguilla', value: '(GMT-04:00) Anguilla' },
        { key: 'America/Antigua', value: '(GMT-04:00) Antigua' },
        { key: 'America/Aruba', value: '(GMT-04:00) Aruba' },
        { key: 'America/Asuncion', value: '(GMT-04:00) Asuncion' },
        { key: 'America/Halifax', value: '(GMT-04:00) Atlantic Time - Halifax' },
        { key: 'America/Barbados', value: '(GMT-04:00) Barbados' },
        { key: 'Atlantic/Bermuda', value: '(GMT-04:00) Bermuda' },
        { key: 'America/Boa_Vista', value: '(GMT-04:00) Boa Vista' },
        { key: 'America/Campo_Grande', value: '(GMT-04:00) Campo Grande' },
        { key: 'America/Caracas', value: '(GMT-04:00) Caracas' },
        { key: 'America/Cuiaba', value: '(GMT-04:00) Cuiaba' },
        { key: 'America/Curacao', value: '(GMT-04:00) Curacao' },
        { key: 'America/Dominica', value: '(GMT-04:00) Dominica' },
        { key: 'America/Grenada', value: '(GMT-04:00) Grenada' },
        { key: 'America/Guadeloupe', value: '(GMT-04:00) Guadeloupe' },
        { key: 'America/Guyana', value: '(GMT-04:00) Guyana' },
        { key: 'America/La_Paz', value: '(GMT-04:00) La Paz' },
        { key: 'America/Manaus', value: '(GMT-04:00) Manaus' },
        { key: 'America/Martinique', value: '(GMT-04:00) Martinique' },
        { key: 'America/Montserrat', value: '(GMT-04:00) Montserrat' },
        { key: 'Antarctica/Palmer', value: '(GMT-04:00) Palmer' },
        { key: 'America/Port_of_Spain', value: '(GMT-04:00) Port of Spain' },
        { key: 'America/Porto_Velho', value: '(GMT-04:00) Porto Velho' },
        { key: 'America/Puerto_Rico', value: '(GMT-04:00) Puerto Rico' },
        { key: 'America/Punta_Arenas', value: '(GMT-04:00) Punta Arenas' },
        { key: 'America/Santiago', value: '(GMT-04:00) Santiago' },
        { key: 'America/Santo_Domingo', value: '(GMT-04:00) Santo Domingo' },
        { key: 'America/St_Kitts', value: '(GMT-04:00) St. Kitts' },
        { key: 'America/St_Lucia', value: '(GMT-04:00) St. Lucia' },
        { key: 'America/St_Thomas', value: '(GMT-04:00) St. Thomas' },
        { key: 'America/St_Vincent', value: '(GMT-04:00) St. Vincent' },
        { key: 'America/Thule', value: '(GMT-04:00) Thule' },
        { key: 'America/Tortola', value: '(GMT-04:00) Tortola' },
        { key: 'America/St_Johns', value: '(GMT-03:30) Newfoundland Time - St. Johns' },
        { key: 'America/Araguaina', value: '(GMT-03:00) Araguaina' },
        { key: 'America/Belem', value: '(GMT-03:00) Belem' },
        { key: 'America/Buenos_Aires', value: '(GMT-03:00) Buenos Aires' },
        { key: 'America/Cayenne', value: '(GMT-03:00) Cayenne' },
        { key: 'America/Cordoba', value: '(GMT-03:00) Cordoba' },
        { key: 'America/Fortaleza', value: '(GMT-03:00) Fortaleza' },
        { key: 'America/Godthab', value: '(GMT-03:00) Godthab' },
        { key: 'America/Maceio', value: '(GMT-03:00) Maceio' },
        { key: 'America/Miquelon', value: '(GMT-03:00) Miquelon' },
        { key: 'America/Montevideo', value: '(GMT-03:00) Montevideo' },
        { key: 'America/Paramaribo', value: '(GMT-03:00) Paramaribo' },
        { key: 'America/Recife', value: '(GMT-03:00) Recife' },
        { key: 'Antarctica/Rothera', value: '(GMT-03:00) Rothera' },
        { key: 'America/Bahia', value: '(GMT-03:00) Salvador' },
        { key: 'America/Sao_Paulo', value: '(GMT-03:00) Sao Paulo' },
        { key: 'Atlantic/Stanley', value: '(GMT-03:00) Stanley' },
        { key: 'America/Noronha', value: '(GMT-02:00) Noronha' },
        { key: 'Atlantic/South_Georgia', value: '(GMT-02:00) South Georgia' },
        { key: 'Atlantic/Azores', value: '(GMT-01:00) Azores' },
        { key: 'Atlantic/Cape_Verde', value: '(GMT-01:00) Cape Verde' },
        { key: 'America/Scoresbysund', value: '(GMT-01:00) Scoresbysund' },
        { key: 'Africa/Abidjan', value: '(GMT+00:00) Abidjan' },
        { key: 'Africa/Accra', value: '(GMT+00:00) Accra' },
        { key: 'Africa/Bamako', value: '(GMT+00:00) Bamako' },
        { key: 'Africa/Banjul', value: '(GMT+00:00) Banjul' },
        { key: 'Africa/Bissau', value: '(GMT+00:00) Bissau' },
        { key: 'Atlantic/Canary', value: '(GMT+00:00) Canary Islands' },
        { key: 'Africa/Casablanca', value: '(GMT+00:00) Casablanca' },
        { key: 'Africa/Conakry', value: '(GMT+00:00) Conakry' },
        { key: 'Africa/Dakar', value: '(GMT+00:00) Dakar' },
        { key: 'America/Danmarkshavn', value: '(GMT+00:00) Danmarkshavn' },
        { key: 'Europe/Dublin', value: '(GMT+00:00) Dublin' },
        { key: 'Africa/El_Aaiun', value: '(GMT+00:00) El Aaiun' },
        { key: 'Atlantic/Faeroe', value: '(GMT+00:00) Faeroe' },
        { key: 'Africa/Freetown', value: '(GMT+00:00) Freetown' },
        { key: 'Etc/GMT', value: '(GMT+00:00) GMT (no daylight saving)' },
        { key: 'Europe/Lisbon', value: '(GMT+00:00) Lisbon' },
        { key: 'Africa/Lome', value: '(GMT+00:00) Lome' },
        { key: 'Europe/London', value: '(GMT+00:00) London' },
        { key: 'Africa/Monrovia', value: '(GMT+00:00) Monrovia' },
        { key: 'Africa/Nouakchott', value: '(GMT+00:00) Nouakchott' },
        { key: 'Africa/Ouagadougou', value: '(GMT+00:00) Ouagadougou' },
        { key: 'Atlantic/Reykjavik', value: '(GMT+00:00) Reykjavik' },
        { key: 'Atlantic/St_Helena', value: '(GMT+00:00) St Helena' },
        { key: 'Africa/Algiers', value: '(GMT+01:00) Algiers' },
        { key: 'Europe/Amsterdam', value: '(GMT+01:00) Amsterdam' },
        { key: 'Europe/Andorra', value: '(GMT+01:00) Andorra' },
        { key: 'Africa/Bangui', value: '(GMT+01:00) Bangui' },
        { key: 'Europe/Berlin', value: '(GMT+01:00) Berlin' },
        { key: 'Africa/Brazzaville', value: '(GMT+01:00) Brazzaville' },
        { key: 'Europe/Brussels', value: '(GMT+01:00) Brussels' },
        { key: 'Europe/Budapest', value: '(GMT+01:00) Budapest' },
        { key: 'Europe/Belgrade', value: '(GMT+01:00) Central European Time - Belgrade' },
        { key: 'Europe/Prague', value: '(GMT+01:00) Central European Time - Prague' },
        { key: 'Africa/Ceuta', value: '(GMT+01:00) Ceuta' },
        { key: 'Europe/Copenhagen', value: '(GMT+01:00) Copenhagen' },
        { key: 'Africa/Douala', value: '(GMT+01:00) Douala' },
        { key: 'Europe/Gibraltar', value: '(GMT+01:00) Gibraltar' },
        { key: 'Africa/Kinshasa', value: '(GMT+01:00) Kinshasa' },
        { key: 'Africa/Lagos', value: '(GMT+01:00) Lagos' },
        { key: 'Africa/Libreville', value: '(GMT+01:00) Libreville' },
        { key: 'Africa/Luanda', value: '(GMT+01:00) Luanda' },
        { key: 'Europe/Luxembourg', value: '(GMT+01:00) Luxembourg' },
        { key: 'Europe/Madrid', value: '(GMT+01:00) Madrid' },
        { key: 'Africa/Malabo', value: '(GMT+01:00) Malabo' },
        { key: 'Europe/Malta', value: '(GMT+01:00) Malta' },
        { key: 'Europe/Monaco', value: '(GMT+01:00) Monaco' },
        { key: 'Africa/Ndjamena', value: '(GMT+01:00) Ndjamena' },
        { key: 'Africa/Niamey', value: '(GMT+01:00) Niamey' },
        { key: 'Europe/Oslo', value: '(GMT+01:00) Oslo' },
        { key: 'Europe/Paris', value: '(GMT+01:00) Paris' },
        { key: 'Africa/Porto-Novo', value: '(GMT+01:00) Porto-Novo' },
        { key: 'Europe/Rome', value: '(GMT+01:00) Rome' },
        { key: 'Africa/Sao_Tome', value: '(GMT+01:00) Sao Tome' },
        { key: 'Europe/Stockholm', value: '(GMT+01:00) Stockholm' },
        { key: 'Europe/Tirane', value: '(GMT+01:00) Tirane' },
        { key: 'Africa/Tunis', value: '(GMT+01:00) Tunis' },
        { key: 'Europe/Vaduz', value: '(GMT+01:00) Vaduz' },
        { key: 'Europe/Vienna', value: '(GMT+01:00) Vienna' },
        { key: 'Europe/Warsaw', value: '(GMT+01:00) Warsaw' },
        { key: 'Africa/Windhoek', value: '(GMT+01:00) Windhoek' },
        { key: 'Europe/Zurich', value: '(GMT+01:00) Zurich' },
        { key: 'Asia/Amman', value: '(GMT+02:00) Amman' },
        { key: 'Europe/Athens', value: '(GMT+02:00) Athens' },
        { key: 'Asia/Beirut', value: '(GMT+02:00) Beirut' },
        { key: 'Africa/Blantyre', value: '(GMT+02:00) Blantyre' },
        { key: 'Europe/Bucharest', value: '(GMT+02:00) Bucharest' },
        { key: 'Africa/Bujumbura', value: '(GMT+02:00) Bujumbura' },
        { key: 'Africa/Cairo', value: '(GMT+02:00) Cairo' },
        { key: 'Europe/Chisinau', value: '(GMT+02:00) Chisinau' },
        { key: 'Asia/Damascus', value: '(GMT+02:00) Damascus' },
        { key: 'Africa/Gaborone', value: '(GMT+02:00) Gaborone' },
        { key: 'Asia/Gaza', value: '(GMT+02:00) Gaza' },
        { key: 'Africa/Harare', value: '(GMT+02:00) Harare' },
        { key: 'Europe/Helsinki', value: '(GMT+02:00) Helsinki' },
        { key: 'Asia/Jerusalem', value: '(GMT+02:00) Jerusalem' },
        { key: 'Africa/Johannesburg', value: '(GMT+02:00) Johannesburg' },
        { key: 'Africa/Khartoum', value: '(GMT+02:00) Khartoum' },
        { key: 'Europe/Kiev', value: '(GMT+02:00) Kiev' },
        { key: 'Africa/Kigali', value: '(GMT+02:00) Kigali' },
        { key: 'Africa/Lubumbashi', value: '(GMT+02:00) Lubumbashi' },
        { key: 'Africa/Lusaka', value: '(GMT+02:00) Lusaka' },
        { key: 'Africa/Maputo', value: '(GMT+02:00) Maputo' },
        { key: 'Africa/Maseru', value: '(GMT+02:00) Maseru' },
        { key: 'Africa/Mbabane', value: '(GMT+02:00) Mbabane' },
        { key: 'Europe/Kaliningrad', value: '(GMT+02:00) Moscow-01 - Kaliningrad' },
        { key: 'Asia/Nicosia', value: '(GMT+02:00) Nicosia' },
        { key: 'Europe/Riga', value: '(GMT+02:00) Riga' },
        { key: 'Europe/Sofia', value: '(GMT+02:00) Sofia' },
        { key: 'Europe/Tallinn', value: '(GMT+02:00) Tallinn' },
        { key: 'Africa/Tripoli', value: '(GMT+02:00) Tripoli' },
        { key: 'Europe/Vilnius', value: '(GMT+02:00) Vilnius' },
        { key: 'Africa/Addis_Ababa', value: '(GMT+03:00) Addis Ababa' },
        { key: 'Asia/Aden', value: '(GMT+03:00) Aden' },
        { key: 'Indian/Antananarivo', value: '(GMT+03:00) Antananarivo' },
        { key: 'Africa/Asmera', value: '(GMT+03:00) Asmera' },
        { key: 'Asia/Baghdad', value: '(GMT+03:00) Baghdad' },
        { key: 'Asia/Bahrain', value: '(GMT+03:00) Bahrain' },
        { key: 'Indian/Comoro', value: '(GMT+03:00) Comoro' },
        { key: 'Africa/Dar_es_Salaam', value: '(GMT+03:00) Dar es Salaam' },
        { key: 'Africa/Djibouti', value: '(GMT+03:00) Djibouti' },
        { key: 'Europe/Istanbul', value: '(GMT+03:00) Istanbul' },
        { key: 'Africa/Kampala', value: '(GMT+03:00) Kampala' },
        { key: 'Asia/Kuwait', value: '(GMT+03:00) Kuwait' },
        { key: 'Indian/Mayotte', value: '(GMT+03:00) Mayotte' },
        { key: 'Europe/Minsk', value: '(GMT+03:00) Minsk' },
        { key: 'Africa/Mogadishu', value: '(GMT+03:00) Mogadishu' },
        { key: 'Europe/Moscow', value: '(GMT+03:00) Moscow+00 - Moscow' },
        { key: 'Africa/Nairobi', value: '(GMT+03:00) Nairobi' },
        { key: 'Asia/Qatar', value: '(GMT+03:00) Qatar' },
        { key: 'Asia/Riyadh', value: '(GMT+03:00) Riyadh' },
        { key: 'Antarctica/Syowa', value: '(GMT+03:00) Syowa' },
        { key: 'Asia/Tehran', value: '(GMT+03:30) Tehran' },
        { key: 'Asia/Aqtau', value: '(GMT+04:00) Aqtau' },
        { key: 'Asia/Baku', value: '(GMT+04:00) Baku' },
        { key: 'Asia/Dubai', value: '(GMT+04:00) Dubai' },
        { key: 'Indian/Mahe', value: '(GMT+04:00) Mahe' },
        { key: 'Indian/Mauritius', value: '(GMT+04:00) Mauritius' },
        { key: 'Europe/Samara', value: '(GMT+04:00) Moscow+01 - Samara' },
        { key: 'Asia/Muscat', value: '(GMT+04:00) Muscat' },
        { key: 'Indian/Reunion', value: '(GMT+04:00) Reunion' },
        { key: 'Asia/Tbilisi', value: '(GMT+04:00) Tbilisi' },
        { key: 'Asia/Yerevan', value: '(GMT+04:00) Yerevan' },
        { key: 'Asia/Kabul', value: '(GMT+04:30) Kabul' },
        { key: 'Asia/Aqtobe', value: '(GMT+05:00) Aqtobe' },
        { key: 'Asia/Ashgabat', value: '(GMT+05:00) Ashgabat' },
        { key: 'Asia/Bishkek', value: '(GMT+05:00) Bishkek' },
        { key: 'Asia/Dushanbe', value: '(GMT+05:00) Dushanbe' },
        { key: 'Asia/Karachi', value: '(GMT+05:00) Karachi' },
        { key: 'Indian/Kerguelen', value: '(GMT+05:00) Kerguelen' },
        { key: 'Indian/Maldives', value: '(GMT+05:00) Maldives' },
        { key: 'Antarctica/Mawson', value: '(GMT+05:00) Mawson' },
        { key: 'Asia/Yekaterinburg', value: '(GMT+05:00) Moscow+02 - Yekaterinburg' },
        { key: 'Asia/Tashkent', value: '(GMT+05:00) Tashkent' },
        { key: 'Asia/Colombo', value: '(GMT+05:30) Colombo' },
        { key: 'Asia/Calcutta', value: '(GMT+05:30) India Standard Time' },
        { key: 'Asia/Katmandu', value: '(GMT+05:45) Katmandu' },
        { key: 'Asia/Almaty', value: '(GMT+06:00) Almaty' },
        { key: 'Indian/Chagos', value: '(GMT+06:00) Chagos' },
        { key: 'Asia/Dhaka', value: '(GMT+06:00) Dhaka' },
        { key: 'Asia/Omsk', value: '(GMT+06:00) Moscow+03 - Omsk' },
        { key: 'Asia/Thimphu', value: '(GMT+06:00) Thimphu' },
        { key: 'Antarctica/Vostok', value: '(GMT+06:00) Vostok' },
        { key: 'Indian/Cocos', value: '(GMT+06:30) Cocos' },
        { key: 'Asia/Rangoon', value: '(GMT+06:30) Rangoon' },
        { key: 'Asia/Bangkok', value: '(GMT+07:00) Bangkok' },
        { key: 'Indian/Christmas', value: '(GMT+07:00) Christmas' },
        { key: 'Antarctica/Davis', value: '(GMT+07:00) Davis' },
        { key: 'Asia/Saigon" selected="', value: '(GMT+07:00) Hanoi' },
        { key: 'Asia/Hovd', value: '(GMT+07:00) Hovd' },
        { key: 'Asia/Jakarta', value: '(GMT+07:00) Jakarta' },
        { key: 'Asia/Krasnoyarsk', value: '(GMT+07:00) Moscow+04 - Krasnoyarsk' },
        { key: 'Asia/Phnom_Penh', value: '(GMT+07:00) Phnom Penh' },
        { key: 'Asia/Vientiane', value: '(GMT+07:00) Vientiane' },
        { key: 'Asia/Brunei', value: '(GMT+08:00) Brunei' },
        { key: 'Antarctica/Casey', value: '(GMT+08:00) Casey' },
        { key: 'Asia/Shanghai', value: '(GMT+08:00) China Time - Beijing' },
        { key: 'Asia/Choibalsan', value: '(GMT+08:00) Choibalsan' },
        { key: 'Asia/Hong_Kong', value: '(GMT+08:00) Hong Kong' },
        { key: 'Asia/Kuala_Lumpur', value: '(GMT+08:00) Kuala Lumpur' },
        { key: 'Asia/Macau', value: '(GMT+08:00) Macau' },
        { key: 'Asia/Makassar', value: '(GMT+08:00) Makassar' },
        { key: 'Asia/Manila', value: '(GMT+08:00) Manila' },
        { key: 'Asia/Irkutsk', value: '(GMT+08:00) Moscow+05 - Irkutsk' },
        { key: 'Asia/Singapore', value: '(GMT+08:00) Singapore' },
        { key: 'Asia/Taipei', value: '(GMT+08:00) Taipei' },
        { key: 'Asia/Ulaanbaatar', value: '(GMT+08:00) Ulaanbaatar' },
        { key: 'Australia/Perth', value: '(GMT+08:00) Western Time - Perth' },
        { key: 'Asia/Dili', value: '(GMT+09:00) Dili' },
        { key: 'Asia/Jayapura', value: '(GMT+09:00) Jayapura' },
        { key: 'Asia/Yakutsk', value: '(GMT+09:00) Moscow+06 - Yakutsk' },
        { key: 'Pacific/Palau', value: '(GMT+09:00) Palau' },
        { key: 'Asia/Pyongyang', value: '(GMT+09:00) Pyongyang' },
        { key: 'Asia/Seoul', value: '(GMT+09:00) Seoul' },
        { key: 'Asia/Tokyo', value: '(GMT+09:00) Tokyo' },
        { key: 'Australia/Adelaide', value: '(GMT+09:30) Central Time - Adelaide' },
        { key: 'Australia/Darwin', value: '(GMT+09:30) Central Time - Darwin' },
        { key: 'Antarctica/DumontDUrville', value: "(GMT+10:00) Dumont D'Urville" },
        { key: 'Australia/Brisbane', value: '(GMT+10:00) Eastern Time - Brisbane' },
        { key: 'Australia/Hobart', value: '(GMT+10:00) Eastern Time - Hobart' },
        { key: 'Australia/Melbourne', value: '(GMT+10:00) Eastern Time - Melbourne' },
        { key: 'Australia/Sydney', value: '(GMT+10:00) Eastern Time - Melbourne, Sydney' },
        { key: 'Pacific/Guam', value: '(GMT+10:00) Guam' },
        { key: 'Asia/Vladivostok', value: '(GMT+10:00) Moscow+07 - Vladivostok' },
        { key: 'Pacific/Port_Moresby', value: '(GMT+10:00) Port Moresby' },
        { key: 'Pacific/Saipan', value: '(GMT+10:00) Saipan' },
        { key: 'Pacific/Truk', value: '(GMT+10:00) Truk' },
        { key: 'Pacific/Efate', value: '(GMT+11:00) Efate' },
        { key: 'Pacific/Guadalcanal', value: '(GMT+11:00) Guadalcanal' },
        { key: 'Pacific/Kosrae', value: '(GMT+11:00) Kosrae' },
        { key: 'Asia/Magadan', value: '(GMT+11:00) Moscow+08 - Magadan' },
        { key: 'Pacific/Norfolk', value: '(GMT+11:00) Norfolk' },
        { key: 'Pacific/Noumea', value: '(GMT+11:00) Noumea' },
        { key: 'Pacific/Ponape', value: '(GMT+11:00) Ponape' },
        { key: 'Pacific/Auckland', value: '(GMT+12:00) Auckland' },
        { key: 'Pacific/Fiji', value: '(GMT+12:00) Fiji' },
        { key: 'Pacific/Funafuti', value: '(GMT+12:00) Funafuti' },
        { key: 'Pacific/Kwajalein', value: '(GMT+12:00) Kwajalein' },
        { key: 'Pacific/Majuro', value: '(GMT+12:00) Majuro' },
        { key: 'Asia/Kamchatka', value: '(GMT+12:00) Moscow+09 - Petropavlovsk-Kamchatskiy' },
        { key: 'Pacific/Nauru', value: '(GMT+12:00) Nauru' },
        { key: 'Pacific/Tarawa', value: '(GMT+12:00) Tarawa' },
        { key: 'Pacific/Wake', value: '(GMT+12:00) Wake' },
        { key: 'Pacific/Wallis', value: '(GMT+12:00) Wallis' },
        { key: 'Pacific/Apia', value: '(GMT+13:00) Apia' },
        { key: 'Pacific/Enderbury', value: '(GMT+13:00) Enderbury' },
        { key: 'Pacific/Fakaofo', value: '(GMT+13:00) Fakaofo' },
        { key: 'Pacific/Tongatapu', value: '(GMT+13:00) Tongatapu' },
        { key: 'Pacific/Kiritimati', value: '(GMT+14:00) Kiritimati' },
    ];
    let tz = timezones.filter(t => {
        return t.value == value;
    });
    return tz.length > 0 ? tz[0].key : '';
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