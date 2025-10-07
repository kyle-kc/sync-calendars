const PRIMARY_CALENDAR_ID_KEY = 'PRIMARY_CALENDAR_ID';
const SECONDARY_CALENDAR_IDS_KEY = 'SECONDARY_CALENDAR_IDS';

const DAYS_LOOKAHEAD = 365;
const STAGING_TITLE_PREFIX = "[STAGING]";
const SCRIPT_ID_TAG_KEY = "autoCreatedByScriptId";
const ORIGINAL_CALENDAR_ID_TAG_KEY = "originalCalendarId";
const ORIGINAL_EVENT_ID_TAG_KEY = "originalEventId";
const PRE_BUFFER_FOR_EVENT_ID_TAG = "preBufferForEventId";
const POST_BUFFER_FOR_EVENT_ID_TAG = "postBufferForEventId";
const SCRIPT_ID = ScriptApp.getScriptId();

const EVENT_COLORS_TO_HEX_CODES = new Map<GoogleAppsScript.Calendar.EventColor, string>([
    [CalendarApp.EventColor.PALE_BLUE, "#a4bdfc"],
    [CalendarApp.EventColor.PALE_GREEN, "#7ae7bf"],
    [CalendarApp.EventColor.MAUVE, "#dbadff"],
    [CalendarApp.EventColor.PALE_RED, "#ff887c"],
    [CalendarApp.EventColor.YELLOW, "#fbd75b"],
    [CalendarApp.EventColor.ORANGE, "#ffb878"],
    [CalendarApp.EventColor.CYAN, "#46d6db"],
    [CalendarApp.EventColor.GRAY, "#e1e1e1"],
    [CalendarApp.EventColor.BLUE, "#5484ed"],
    [CalendarApp.EventColor.GREEN, "#51b749"],
    [CalendarApp.EventColor.RED, "#dc2127"]
]);

const INITIAL_BACKOFF_MILLISECONDS = 200;
const MAX_RETRIES = 10;

const BUFFER_DURATION_MILLISECONDS = 30 * 60 * 1000; // 30 minutes

const hexCodeToClosestEventColorCache = new Map<string, GoogleAppsScript.Calendar.EventColor>();

// Custom type for extended Date
interface ExtendedDate extends Date {
    setTimeToMidnight(): ExtendedDate;
}

function extendDate(date: Date): ExtendedDate {
    const extended = date as ExtendedDate;
    extended.setTimeToMidnight = function(): ExtendedDate {
        this.setHours(0);
        this.setMinutes(0);
        this.setSeconds(0);
        this.setMilliseconds(0);
        return this;
    };
    return extended;
}

function hexToRgb(hex: string): [number, number, number] {
    const bigint = parseInt(hex.substring(1), 16);
    return [(bigint >> 16) & 255, (bigint >> 8) & 255, bigint & 255];
}

function rgbToXyz(rgb: [number, number, number]): [number, number, number] {
    let r = rgb[0] / 255, g = rgb[1] / 255, b = rgb[2] / 255;

    r = r > 0.04045 ? Math.pow((r + 0.055) / 1.055, 2.4) : r / 12.92;
    g = g > 0.04045 ? Math.pow((g + 0.055) / 1.055, 2.4) : g / 12.92;
    b = b > 0.04045 ? Math.pow((b + 0.055) / 1.055, 2.4) : b / 12.92;

    r *= 100;
    g *= 100;
    b *= 100;

    return [
        r * 0.4124564 + g * 0.3575761 + b * 0.1804375,
        r * 0.2126729 + g * 0.7151522 + b * 0.0721750,
        r * 0.0193339 + g * 0.1191920 + b * 0.9503041
    ];
}

function xyzToLab(xyz: [number, number, number]): [number, number, number] {
    let x = xyz[0] / 95.047, y = xyz[1] / 100.000, z = xyz[2] / 108.883;

    x = x > 0.008856 ? Math.pow(x, 1 / 3) : (7.787 * x) + (16 / 116);
    y = y > 0.008856 ? Math.pow(y, 1 / 3) : (7.787 * y) + (16 / 116);
    z = z > 0.008856 ? Math.pow(z, 1 / 3) : (7.787 * z) + (16 / 116);

    return [(116 * y) - 16, 500 * (x - y), 200 * (y - z)];
}

function deltaE76(lab1: [number, number, number], lab2: [number, number, number]): number {
    return Math.hypot(
        lab1[0] - lab2[0],
        lab1[1] - lab2[1],
        lab1[2] - lab2[2]
    );
}

function getClosestEventColor(hexCode: string): GoogleAppsScript.Calendar.EventColor {
    if (!hexCodeToClosestEventColorCache.has(hexCode)) {
        let closestColor: GoogleAppsScript.Calendar.EventColor | undefined;
        let minDistance = Number.MAX_VALUE;

        EVENT_COLORS_TO_HEX_CODES.forEach((hexColor, eventColor) => {
            const distance = deltaE76(xyzToLab(rgbToXyz(hexToRgb(hexCode))), xyzToLab(rgbToXyz(hexToRgb(hexColor))));
            if (distance < minDistance) {
                minDistance = distance;
                closestColor = eventColor;
            }
        });

        if (closestColor !== undefined) {
            hexCodeToClosestEventColorCache.set(hexCode, closestColor);
        }
    }
    return hexCodeToClosestEventColorCache.get(hexCode)!;
}

function callWithRetryAndExponentialBackoff<T>(apiFunction: () => T): T {
    let numberOfTries = 0;
    while (true) {
        try {
            return apiFunction();
        } catch (exception: any) {
            if (
                exception.message.includes("You have been creating or deleting too many calendars") &&
                numberOfTries <= MAX_RETRIES
            ) {
                Utilities.sleep(INITIAL_BACKOFF_MILLISECONDS * 2 ** numberOfTries);
                numberOfTries++;
            } else {
                throw exception;
            }
        }
    }
}

function setIfNeeded<T>(
    setMethod: (value: T) => void,
    getMethod: () => T | null | undefined,
    newValue: T | null | undefined
): void {
    const currentValue = getMethod();
    if ((currentValue || newValue) && currentValue !== newValue) {
        callWithRetryAndExponentialBackoff(() => setMethod(newValue as T));
    }
}

function setTagIfNeeded(
    event: GoogleAppsScript.Calendar.CalendarEvent,
    tag_key: string,
    tag_value: string
): void {
    if (event.getTag(tag_key) !== tag_value) {
        callWithRetryAndExponentialBackoff(() => event.setTag(tag_key, tag_value));
    }
}

function setStartAndEndTimesIfNeeded(
    setMethod: (startTime: Date, endTime: Date) => void,
    getStartTimeMethod: () => Date,
    getEndTimeMethod: () => Date,
    startTime: Date,
    endTime: Date
): void {
    if (getStartTimeMethod().getTime() !== startTime.getTime() || getEndTimeMethod().getTime() !== endTime.getTime()) {
        callWithRetryAndExponentialBackoff(() => setMethod(startTime, endTime));
    }
}

function setEventAttributesIfNeeded(
    targetEvent: GoogleAppsScript.Calendar.CalendarEvent,
    sourceEvent: GoogleAppsScript.Calendar.CalendarEvent,
    sourceCalendar: GoogleAppsScript.Calendar.Calendar,
    description: string | null = sourceEvent.getDescription(),
    location: string | null = sourceEvent.getLocation()
): void {
    const sourceEventColor = sourceEvent.getColor();
    const calendarColorHex = sourceCalendar.getColor();

    // Determine the color to set - if source event has a color use it, otherwise calculate from calendar color
    let colorToSet: GoogleAppsScript.Calendar.EventColor | null = null;
    if (sourceEventColor) {
        // sourceEventColor can be either EventColor or string depending on the API
        if (typeof sourceEventColor === 'string') {
            colorToSet = getClosestEventColor(sourceEventColor);
        } else {
            colorToSet = sourceEventColor as GoogleAppsScript.Calendar.EventColor;
        }
    } else if (calendarColorHex) {
        colorToSet = getClosestEventColor(calendarColorHex);
    }

    if (colorToSet) {
        const currentColor = targetEvent.getColor();
        // Convert current color to EventColor if it's a string
        let currentEventColor: GoogleAppsScript.Calendar.EventColor | null = null;
        if (currentColor) {
            if (typeof currentColor === 'string') {
                currentEventColor = getClosestEventColor(currentColor);
            } else {
                currentEventColor = currentColor as GoogleAppsScript.Calendar.EventColor;
            }
        }

        if (currentEventColor !== colorToSet) {
            callWithRetryAndExponentialBackoff(() => (targetEvent as any).setColor(colorToSet));
        }
    }
    setIfNeeded(
        (value: boolean) => targetEvent.setAnyoneCanAddSelf(value),
        () => targetEvent.anyoneCanAddSelf(),
        false
    );
    setIfNeeded(
        (desc: string | null) => targetEvent.setDescription(desc || ''),
        () => targetEvent.getDescription(),
        description
    );
    setIfNeeded(
        (value: boolean) => targetEvent.setGuestsCanInviteOthers(value),
        () => targetEvent.guestsCanInviteOthers(),
        false
    );
    setIfNeeded(
        (value: boolean) => targetEvent.setGuestsCanModify(value),
        () => targetEvent.guestsCanModify(),
        false
    );
    setIfNeeded(
        (value: boolean) => targetEvent.setGuestsCanSeeGuests(value),
        () => targetEvent.guestsCanSeeGuests(),
        false
    );
    setIfNeeded(
        (loc: string | null) => targetEvent.setLocation(loc || ''),
        () => targetEvent.getLocation(),
        location
    );
    // Transparency methods might not be in type definitions, using any cast
    setIfNeeded(
        (transparency: any) => (targetEvent as any).setTransparency(transparency),
        () => (targetEvent as any).getTransparency ? (targetEvent as any).getTransparency() : null,
        (sourceEvent as any).getTransparency ? (sourceEvent as any).getTransparency() : null
    );
    setIfNeeded(
        (visibility: GoogleAppsScript.Calendar.Visibility) => targetEvent.setVisibility(visibility),
        () => targetEvent.getVisibility(),
        CalendarApp.Visibility.DEFAULT
    );
    callWithRetryAndExponentialBackoff(() => targetEvent.removeAllReminders());  // the get methods don't actually return the correct data for this, so we just remove all reminders to be safe
}

function createOrUpdateBufferEvent(
    primaryCalendar: GoogleAppsScript.Calendar.Calendar,
    previouslyCreatedEvents: GoogleAppsScript.Calendar.CalendarEvent[],
    event: GoogleAppsScript.Calendar.CalendarEvent,
    bufferType: "Pre" | "Post",
    secondaryCalendar: GoogleAppsScript.Calendar.Calendar
): void {
    const bufferEventTitle = `${bufferType}-Buffer for ${event.getTitle()}`;
    const bufferForEventIdTag = bufferType === "Pre" ? PRE_BUFFER_FOR_EVENT_ID_TAG : POST_BUFFER_FOR_EVENT_ID_TAG;
    const bufferEventStartTime = new Date(
        bufferType === "Pre" ? event.getStartTime().getTime() - BUFFER_DURATION_MILLISECONDS : event.getEndTime().getTime()
    );
    const bufferEventEndTime = new Date(bufferEventStartTime.getTime() + BUFFER_DURATION_MILLISECONDS);

    const bufferEventIndex = previouslyCreatedEvents.findIndex(
        (previouslyCreatedEvent) =>
            previouslyCreatedEvent.getTag(ORIGINAL_CALENDAR_ID_TAG_KEY) === secondaryCalendar.getId() &&
            previouslyCreatedEvent.getTag(bufferForEventIdTag) === event.getId()
    );

    let bufferEvent: GoogleAppsScript.Calendar.CalendarEvent;
    if (bufferEventIndex === -1) {
        bufferEvent = callWithRetryAndExponentialBackoff(() =>
            primaryCalendar.createEvent(
                `${STAGING_TITLE_PREFIX} ${bufferEventTitle}`,
                bufferEventStartTime,
                bufferEventEndTime,
                {
                    description: null,
                    location: null,
                }
            )
        );
    } else {
        bufferEvent = previouslyCreatedEvents[bufferEventIndex];
        previouslyCreatedEvents.splice(bufferEventIndex, 1);
    }

    setTagIfNeeded(bufferEvent, SCRIPT_ID_TAG_KEY, SCRIPT_ID);
    setTagIfNeeded(bufferEvent, ORIGINAL_CALENDAR_ID_TAG_KEY, secondaryCalendar.getId());
    setTagIfNeeded(bufferEvent, bufferType === "Pre" ? PRE_BUFFER_FOR_EVENT_ID_TAG : POST_BUFFER_FOR_EVENT_ID_TAG, event.getId());

    setStartAndEndTimesIfNeeded(
        (start: Date, end: Date) => bufferEvent.setTime(start, end),
        () => new Date(bufferEvent.getStartTime() as any),
        () => new Date(bufferEvent.getEndTime() as any),
        bufferEventStartTime,
        bufferEventEndTime
    );
    setEventAttributesIfNeeded(bufferEvent, event, secondaryCalendar, null, null);
    setIfNeeded(
        (title: string) => bufferEvent.setTitle(title),
        () => bufferEvent.getTitle(),
        bufferEventTitle
    );
}

function cleanUpStagingEvents(
    primaryCalendar: GoogleAppsScript.Calendar.Calendar,
    today: Date,
    endDate: Date
): void {
    const stagedEvents = primaryCalendar
        .getEvents(today, endDate)
        .filter(event => event.getTitle().startsWith(STAGING_TITLE_PREFIX));

    for (const event of stagedEvents) {
        console.warn(`Deleting orphaned staged event: ${event.getTitle()} (${event.getId()})`);
        callWithRetryAndExponentialBackoff(() => event.deleteEvent());
    }
}

function main(): void {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(1)) {
        return;
    }

    try {
        const scriptProperties = PropertiesService.getScriptProperties();

        const PRIMARY_CALENDAR_ID = scriptProperties.getProperty(PRIMARY_CALENDAR_ID_KEY);
        if (!PRIMARY_CALENDAR_ID) {
            throw new Error(`${PRIMARY_CALENDAR_ID_KEY} not set. Add it in Project Settings > Script Properties`);
        }

        const SECONDARY_CALENDAR_IDS_STRING = scriptProperties.getProperty(SECONDARY_CALENDAR_IDS_KEY);
        if (!SECONDARY_CALENDAR_IDS_STRING) {
            throw new Error(`${SECONDARY_CALENDAR_IDS_KEY} not set. Add it in Project Settings > Script Properties (comma-separated)`);
        }
        const SECONDARY_CALENDAR_IDS = SECONDARY_CALENDAR_IDS_STRING.split(',').map(id => id.trim());

        const today = extendDate(new Date());
        today.setTimeToMidnight();
        const endDate = extendDate(new Date());
        endDate.setDate(today.getDate() + DAYS_LOOKAHEAD);
        endDate.setTimeToMidnight();

        const primaryCalendar = callWithRetryAndExponentialBackoff(() =>
            CalendarApp.getCalendarById(PRIMARY_CALENDAR_ID)
        );

        cleanUpStagingEvents(primaryCalendar, today, endDate);

        const previouslyCreatedEvents = callWithRetryAndExponentialBackoff(() =>
            primaryCalendar.getEvents(today, endDate)
        ).filter(event => event.getTag(SCRIPT_ID_TAG_KEY) === SCRIPT_ID);

        const orphanedEvents = callWithRetryAndExponentialBackoff(() =>
            primaryCalendar.getEvents(today, endDate)
        ).filter(event =>
            !event.getAllTagKeys().includes(SCRIPT_ID_TAG_KEY) &&
            (event.getTitle().startsWith("Pre-Buffer") || event.getTitle().startsWith("Post-Buffer"))
        );

        for (const event of orphanedEvents) {
            console.log("Deleting event: " + event.getTitle() + " " + new Date(event.getStartTime() as any));
            callWithRetryAndExponentialBackoff(() => event.deleteEvent());
        }

        for (const secondaryCalendarId of SECONDARY_CALENDAR_IDS) {
            const secondaryCalendar = callWithRetryAndExponentialBackoff(() =>
                CalendarApp.getCalendarById(secondaryCalendarId)
            );

            for (const secondaryEvent of callWithRetryAndExponentialBackoff(() => secondaryCalendar.getEvents(today, endDate))) {
                const primaryEventIndex = previouslyCreatedEvents.findIndex(
                    event =>
                        event.getTag(ORIGINAL_CALENDAR_ID_TAG_KEY) === secondaryCalendarId &&
                        event.getTag(ORIGINAL_EVENT_ID_TAG_KEY) === secondaryEvent.getId()
                );

                let primaryEvent: GoogleAppsScript.Calendar.CalendarEvent;
                if (primaryEventIndex === -1) {
                    primaryEvent = secondaryEvent.isAllDayEvent()
                        ? primaryCalendar.createAllDayEvent(
                            secondaryEvent.getTitle(),
                            new Date(secondaryEvent.getAllDayStartDate() as any),
                            new Date(secondaryEvent.getAllDayEndDate() as any),
                            {description: secondaryEvent.getDescription(), location: secondaryEvent.getLocation()}
                        )
                        : primaryCalendar.createEvent(
                            `${STAGING_TITLE_PREFIX} ${secondaryEvent.getTitle()}`,
                            secondaryEvent.getStartTime(),
                            secondaryEvent.getEndTime(),
                            {description: secondaryEvent.getDescription(), location: secondaryEvent.getLocation()}
                        );
                } else {
                    primaryEvent = previouslyCreatedEvents[primaryEventIndex];
                    previouslyCreatedEvents.splice(primaryEventIndex, 1);
                }

                setTagIfNeeded(primaryEvent, SCRIPT_ID_TAG_KEY, SCRIPT_ID);
                setTagIfNeeded(primaryEvent, ORIGINAL_CALENDAR_ID_TAG_KEY, secondaryCalendarId);
                setTagIfNeeded(primaryEvent, ORIGINAL_EVENT_ID_TAG_KEY, secondaryEvent.getId());
                setEventAttributesIfNeeded(primaryEvent, secondaryEvent, secondaryCalendar);

                if (secondaryEvent.isAllDayEvent()) {
                    const startDate = new Date(secondaryEvent.getAllDayStartDate() as any);
                    const endDate = new Date(secondaryEvent.getAllDayEndDate() as any);
                    setStartAndEndTimesIfNeeded(
                        (start: Date, end: Date) => primaryEvent.setAllDayDates(start, end),
                        () => (primaryEvent.isAllDayEvent() ? new Date(primaryEvent.getAllDayStartDate() as any) : new Date()),
                        () => (primaryEvent.isAllDayEvent() ? new Date(primaryEvent.getAllDayEndDate() as any) : new Date()),
                        startDate,
                        endDate
                    );
                    setIfNeeded(
                        (title: string) => primaryEvent.setTitle(title),
                        () => primaryEvent.getTitle(),
                        secondaryEvent.getTitle()
                    );
                } else {
                    setStartAndEndTimesIfNeeded(
                        (start: Date, end: Date) => primaryEvent.setTime(start, end),
                        () => new Date(primaryEvent.getStartTime() as any),
                        () => new Date(primaryEvent.getEndTime() as any),
                        new Date(secondaryEvent.getStartTime() as any),
                        new Date(secondaryEvent.getEndTime() as any)
                    );
                    setIfNeeded(
                        (title: string) => primaryEvent.setTitle(title),
                        () => primaryEvent.getTitle(),
                        secondaryEvent.getTitle()
                    );
                    createOrUpdateBufferEvent(primaryCalendar, previouslyCreatedEvents, primaryEvent, "Pre", secondaryCalendar);
                    createOrUpdateBufferEvent(primaryCalendar, previouslyCreatedEvents, primaryEvent, "Post", secondaryCalendar);
                }
            }
        }

        // anything still in the list must have been deleted on the secondary calendar or accidentally duplicated, so let's delete it
        for (const primaryEvent of previouslyCreatedEvents) {
            callWithRetryAndExponentialBackoff(() => primaryEvent.deleteEvent());
        }

    } catch (exception) {
        throw exception;
    } finally {
        lock.releaseLock();
    }
}