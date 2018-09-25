export function l(msgKey: string): string {
    return messages[msgKey] ? messages[msgKey] : msgKey;
}

const messages: any = {
    'scriptName': 'Mail Scheduler',
    'installScript': 'Install script',
    'installSuccess': 'Script installed successfully',
    'installFailure': 'Could not install the script. Please try again!',
    'recurrence.None': 'None',
    'recurrence.Daily': 'Daily',
    'recurrence.Weekly': 'Weekly',
    'recurrence.Monthly': 'Monthly',
    'weekDay.Monday': 'Monday',
    'weekDay.Tuesday': 'Tuesday',
    'weekDay.Wednesday': 'Wednesday',
    'weekDay.Thursday': 'Thursday',
    'weekDay.Friday': 'Friday',
    'weekDay.Saturday': 'Saturday',
    'weekDay.Sunday': 'Sunday',
};