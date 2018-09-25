export enum Recurrence { None, Daily, Weekly, Monthly, Custom };

export class SchedulerEntry {
    constructor() {
        this.Uid = '';
        this.To = '';
        this.Subject = '';
        this.Message = '';
        this.Mode = Recurrence.None;
        this.Hour = 0;
        this.Minute = 0;
    }

    public Uid: string;
    public To: string;
    public Subject: string;
    public Message: string;
    public Timzone?: string;
    public Mode: Recurrence;
    public Weekday?: GoogleAppsScript.Base.Weekday;
    public Hour: number;
    public Minute: number;
    public Day?: number;
    public SentOnDate?: Date;
}

export class SchedulerAddress {
    public static readonly Sheet = 'Schedule';
    public static readonly StartRow = 2;
    public static readonly To = 'B';
    public static readonly Subject = 'C';
    public static readonly Message = 'D';
    public static readonly Mode = 'F';
    public static readonly Weekdays = 'H';
    public static readonly SentOnTime = 'E';
    public static readonly SentOnDate = 'G';
    public static readonly Day = 'I';
    public static readonly Timzone = 'J';
    public static readonly Uid = 'K';
}