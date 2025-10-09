type LAB = readonly [l: number, a: number, b: number];

const PRIMARY_CALENDAR_ID_KEY = "PRIMARY_CALENDAR_ID";
const SECONDARY_CALENDAR_IDS_KEY = "SECONDARY_CALENDAR_IDS";
const DAYS_LOOKAHEAD = 365;
const STAGING_TITLE_PREFIX = "[STAGING]";
const SCRIPT_ID_TAG_KEY = "autoCreatedByScriptId";
const ORIGINAL_CALENDAR_ID_TAG_KEY = "originalCalendarId";
const ORIGINAL_EVENT_ID_TAG_KEY = "originalEventId";
const PRE_BUFFER_FOR_EVENT_ID_TAG = "preBufferForEventId";
const POST_BUFFER_FOR_EVENT_ID_TAG = "postBufferForEventId";
const SCRIPT_ID = ScriptApp.getScriptId();

const EVENT_COLOR_KEYS: GoogleAppsScript.Calendar.EventColor[] = [
  CalendarApp.EventColor.PALE_BLUE,
  CalendarApp.EventColor.PALE_GREEN,
  CalendarApp.EventColor.MAUVE,
  CalendarApp.EventColor.PALE_RED,
  CalendarApp.EventColor.YELLOW,
  CalendarApp.EventColor.ORANGE,
  CalendarApp.EventColor.CYAN,
  CalendarApp.EventColor.GRAY,
  CalendarApp.EventColor.BLUE,
  CalendarApp.EventColor.GREEN,
  CalendarApp.EventColor.RED,
];

const EVENT_COLORS_TO_HEX_CODES: Record<
  GoogleAppsScript.Calendar.EventColor,
  string
> = {
  [CalendarApp.EventColor.PALE_BLUE]: "#a4bdfc",
  [CalendarApp.EventColor.PALE_GREEN]: "#7ae7bf",
  [CalendarApp.EventColor.MAUVE]: "#dbadff",
  [CalendarApp.EventColor.PALE_RED]: "#ff887c",
  [CalendarApp.EventColor.YELLOW]: "#fbd75b",
  [CalendarApp.EventColor.ORANGE]: "#ffb878",
  [CalendarApp.EventColor.CYAN]: "#46d6db",
  [CalendarApp.EventColor.GRAY]: "#e1e1e1",
  [CalendarApp.EventColor.BLUE]: "#5484ed",
  [CalendarApp.EventColor.GREEN]: "#51b749",
  [CalendarApp.EventColor.RED]: "#dc2127",
};

const INITIAL_BACKOFF_MILLISECONDS = 200;
const MAX_RETRIES = 10;
const BUFFER_DURATION_MILLISECONDS = 30 * 60 * 1000; // 30 minutes

const hexCodeToClosestEventColorCache = new Map<
  string,
  GoogleAppsScript.Calendar.EventColor
>();

function main(): void {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1)) {
    return;
  }

  try {
    const scriptProperties = PropertiesService.getScriptProperties();

    const primaryCalendarId = scriptProperties.getProperty(
      PRIMARY_CALENDAR_ID_KEY,
    );
    if (!primaryCalendarId) {
      throw new Error(
        `${PRIMARY_CALENDAR_ID_KEY} not set. Add it in Project Settings > Script Properties.`,
      );
    }

    const secondaryCalendarIdsString = scriptProperties.getProperty(
      SECONDARY_CALENDAR_IDS_KEY,
    );
    if (!secondaryCalendarIdsString) {
      throw new Error(
        `${SECONDARY_CALENDAR_IDS_KEY} not set. Add it in Project Settings > Script Properties (comma-separated).`,
      );
    }
    const secondaryCalendarIds = secondaryCalendarIdsString
      .split(",")
      .map((secondaryCalendarId) => secondaryCalendarId.trim());

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const endDate = new Date();
    endDate.setDate(today.getDate() + DAYS_LOOKAHEAD);
    endDate.setHours(0, 0, 0, 0);

    const primaryCalendar = callWithRetryAndExponentialBackoff(() =>
      CalendarApp.getCalendarById(primaryCalendarId),
    );

    cleanUpStagingEvents(primaryCalendar, today, endDate);

    const previouslyCreatedEvents = callWithRetryAndExponentialBackoff(() =>
      primaryCalendar.getEvents(today, endDate),
    ).filter(
      (event: GoogleAppsScript.Calendar.CalendarEvent) =>
        event.getTag(SCRIPT_ID_TAG_KEY) === SCRIPT_ID,
    );

    const orphanedEvents = callWithRetryAndExponentialBackoff(() =>
      primaryCalendar.getEvents(today, endDate),
    ).filter(
      (event: GoogleAppsScript.Calendar.CalendarEvent) =>
        !event.getAllTagKeys().includes(SCRIPT_ID_TAG_KEY) &&
        (event.getTitle().startsWith("Pre-Buffer") ||
          event.getTitle().startsWith("Post-Buffer")),
    );

    for (const event of orphanedEvents) {
      Logger.log(
        "Deleting event: " + event.getTitle() + " " + event.getStartTime(),
      );
      callWithRetryAndExponentialBackoff(event.deleteEvent);
    }

    for (const secondaryCalendarId of secondaryCalendarIds) {
      const secondaryCalendar = callWithRetryAndExponentialBackoff(() =>
        CalendarApp.getCalendarById(secondaryCalendarId),
      );

      for (const secondaryEvent of callWithRetryAndExponentialBackoff(() =>
        secondaryCalendar.getEvents(today, endDate),
      )) {
        const primaryEventIndex = previouslyCreatedEvents.findIndex(
          (event: GoogleAppsScript.Calendar.CalendarEvent) =>
            event.getTag(ORIGINAL_CALENDAR_ID_TAG_KEY) ===
              secondaryCalendarId &&
            event.getTag(ORIGINAL_EVENT_ID_TAG_KEY) === secondaryEvent.getId(),
        );

        const primaryEvent: GoogleAppsScript.Calendar.CalendarEvent = (() => {
          if (primaryEventIndex === -1) {
            return secondaryEvent.isAllDayEvent()
              ? primaryCalendar.createAllDayEvent(
                  secondaryEvent.getTitle(),
                  secondaryEvent.getAllDayStartDate(),
                  secondaryEvent.getAllDayEndDate(),
                  {
                    description: secondaryEvent.getDescription(),
                    location: secondaryEvent.getLocation(),
                  },
                )
              : primaryCalendar.createEvent(
                  `${STAGING_TITLE_PREFIX} ${secondaryEvent.getTitle()}`,
                  secondaryEvent.getStartTime(),
                  secondaryEvent.getEndTime(),
                  {
                    description: secondaryEvent.getDescription(),
                    location: secondaryEvent.getLocation(),
                  },
                );
          } else {
            const event = previouslyCreatedEvents[primaryEventIndex];
            previouslyCreatedEvents.splice(primaryEventIndex, 1);
            return event;
          }
        })();

        setTagIfNeeded(primaryEvent, SCRIPT_ID_TAG_KEY, SCRIPT_ID);
        setTagIfNeeded(
          primaryEvent,
          ORIGINAL_CALENDAR_ID_TAG_KEY,
          secondaryCalendarId,
        );
        setTagIfNeeded(
          primaryEvent,
          ORIGINAL_EVENT_ID_TAG_KEY,
          secondaryEvent.getId(),
        );
        setEventAttributesIfNeeded(
          primaryEvent,
          secondaryEvent,
          secondaryCalendar,
        );

        if (secondaryEvent.isAllDayEvent()) {
          const startDate = new Date(
            secondaryEvent.getAllDayStartDate().getTime(),
          );
          const endDate = new Date(secondaryEvent.getAllDayEndDate().getTime());
          setStartAndEndTimesIfNeeded(
            primaryEvent.setAllDayDates,
            () =>
              primaryEvent.isAllDayEvent()
                ? new Date(primaryEvent.getAllDayStartDate().getTime())
                : new Date(),
            () =>
              primaryEvent.isAllDayEvent()
                ? new Date(primaryEvent.getAllDayEndDate().getTime())
                : new Date(),
            startDate,
            endDate,
          );
          setIfNeeded(
            primaryEvent.setTitle,
            primaryEvent.getTitle,
            secondaryEvent.getTitle(),
          );
        } else {
          setStartAndEndTimesIfNeeded(
            primaryEvent.setTime,
            () => new Date(primaryEvent.getStartTime().getTime()),
            () => new Date(primaryEvent.getEndTime().getTime()),
            new Date(secondaryEvent.getStartTime().getTime()),
            new Date(secondaryEvent.getEndTime().getTime()),
          );
          setIfNeeded(
            primaryEvent.setTitle,
            primaryEvent.getTitle,
            secondaryEvent.getTitle(),
          );
          createOrUpdateBufferEvent(
            primaryCalendar,
            previouslyCreatedEvents,
            primaryEvent,
            "Pre",
            secondaryCalendar,
          );
          createOrUpdateBufferEvent(
            primaryCalendar,
            previouslyCreatedEvents,
            primaryEvent,
            "Post",
            secondaryCalendar,
          );
        }
      }
    }

    for (const primaryEvent of previouslyCreatedEvents) {
      callWithRetryAndExponentialBackoff(primaryEvent.deleteEvent);
    }
  } finally {
    lock.releaseLock();
  }
}

function hexToLab(hex: string): LAB {
  // Parse hex to RGB
  const number = parseInt(hex.slice(1), 16);
  const [r, g, b] = [
    (number >> 16) & 255,
    (number >> 8) & 255,
    number & 255,
  ].map((value) => {
    const normalizedChannel = value / 255;
    return normalizedChannel > 0.04045
      ? ((normalizedChannel + 0.055) / 1.055) ** 2.4
      : normalizedChannel / 12.92;
  });

  // Convert RGB → XYZ (D65)
  const [x, y, z] = [
    r * 0.4124564 + g * 0.3575761 + b * 0.1804375,
    r * 0.2126729 + g * 0.7151522 + b * 0.072175,
    r * 0.0193339 + g * 0.119192 + b * 0.9503041,
  ].map((value) => value * 100);

  // Normalize for LAB
  const [xr, yr, zr] = [x / 95.047, y / 100.0, z / 108.883].map((value) =>
    value > 0.008856 ? Math.cbrt(value) : 7.787 * value + 16 / 116,
  );

  // Return LAB
  return [116 * yr - 16, 500 * (xr - yr), 200 * (yr - zr)] as const;
}

function calculateColorDistance(colorHex1: string, colorHex2: string): number {
  return Math.hypot(
    hexToLab(colorHex1)[0] - hexToLab(colorHex2)[0],
    hexToLab(colorHex1)[1] - hexToLab(colorHex2)[1],
    hexToLab(colorHex1)[2] - hexToLab(colorHex2)[2],
  );
}

function getClosestEventColor(
  targetColorHex: string,
): GoogleAppsScript.Calendar.EventColor {
  if (!hexCodeToClosestEventColorCache.has(targetColorHex)) {
    const closestColor = EVENT_COLOR_KEYS.reduce<{
      color: GoogleAppsScript.Calendar.EventColor | undefined;
      distance: number;
    }>(
      (eventColorCandidate, colorKey) => {
        const distance = calculateColorDistance(
          targetColorHex,
          EVENT_COLORS_TO_HEX_CODES[colorKey],
        );
        return distance < eventColorCandidate.distance
          ? {
              color: colorKey,
              distance,
            }
          : eventColorCandidate;
      },
      { color: undefined, distance: Infinity },
    ).color;

    hexCodeToClosestEventColorCache.set(targetColorHex, closestColor!);
  }
  return hexCodeToClosestEventColorCache.get(targetColorHex)!;
}

function callWithRetryAndExponentialBackoff<T>(
  apiFunction: () => T,
  attempt = 0,
): T {
  try {
    return apiFunction();
  } catch (error: unknown) {
    if (
      !(error instanceof Error) ||
      error.message.includes(
        "You have been creating or deleting too many calendars",
      ) ||
      attempt >= MAX_RETRIES
    ) {
      throw error;
    }
    Utilities.sleep(
      INITIAL_BACKOFF_MILLISECONDS * 2 ** attempt + Math.random() * 500,
    );
    return callWithRetryAndExponentialBackoff(apiFunction, attempt + 1);
  }
}

function setIfNeeded<T>(
  setMethod: (value: T) => void,
  getMethod: () => T,
  newValue: T,
): void {
  const currentValue = getMethod();
  if ((currentValue || newValue) && currentValue !== newValue) {
    callWithRetryAndExponentialBackoff(() => setMethod(newValue as T));
  }
}

function setTagIfNeeded(
  event: GoogleAppsScript.Calendar.CalendarEvent,
  tag_key: string,
  tag_value: string,
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
  endTime: Date,
): void {
  if (
    getStartTimeMethod().getTime() !== startTime.getTime() ||
    getEndTimeMethod().getTime() !== endTime.getTime()
  ) {
    callWithRetryAndExponentialBackoff(() => setMethod(startTime, endTime));
  }
}

function setEventAttributesIfNeeded(
  targetEvent: GoogleAppsScript.Calendar.CalendarEvent,
  sourceEvent: GoogleAppsScript.Calendar.CalendarEvent,
  sourceCalendar: GoogleAppsScript.Calendar.Calendar,
  description: string | null = sourceEvent.getDescription(),
  location: string | null = sourceEvent.getLocation(),
): void {
  const sourceEventColor = sourceEvent.getColor();
  const calendarColorHex = sourceCalendar.getColor();

  let colorToSet: GoogleAppsScript.Calendar.EventColor | null = null;
  if (sourceEventColor) {
    if (typeof sourceEventColor === "string") {
      colorToSet = getClosestEventColor(sourceEventColor);
    } else {
      colorToSet = sourceEventColor as GoogleAppsScript.Calendar.EventColor;
    }
  } else if (calendarColorHex) {
    colorToSet = getClosestEventColor(calendarColorHex);
  }

  if (colorToSet) {
    const currentColor = targetEvent.getColor();
    let currentEventColor: GoogleAppsScript.Calendar.EventColor | null = null;
    if (currentColor) {
      if (typeof currentColor === "string") {
        currentEventColor = getClosestEventColor(currentColor);
      } else {
        currentEventColor =
          currentColor as GoogleAppsScript.Calendar.EventColor;
      }
    }

    if (currentEventColor !== colorToSet) {
      callWithRetryAndExponentialBackoff(() =>
        targetEvent.setColor(colorToSet),
      );
    }
  }
  setIfNeeded(
    targetEvent.setAnyoneCanAddSelf,
    targetEvent.anyoneCanAddSelf,
    false,
  );
  if (description) {
    setIfNeeded(
      targetEvent.setDescription,
      targetEvent.getDescription,
      description,
    );
  }
  setIfNeeded(
    targetEvent.setGuestsCanInviteOthers,
    targetEvent.guestsCanInviteOthers,
    false,
  );
  setIfNeeded(
    targetEvent.setGuestsCanModify,
    targetEvent.guestsCanModify,
    false,
  );
  setIfNeeded(
    targetEvent.setGuestsCanSeeGuests,
    targetEvent.guestsCanSeeGuests,
    false,
  );
  if (location) {
    setIfNeeded(targetEvent.setLocation, targetEvent.getLocation, location);
  }
  setIfNeeded(
    targetEvent.setTransparency,
    targetEvent.getTransparency,
    sourceEvent.getTransparency(),
  );
  setIfNeeded(
    targetEvent.setVisibility,
    targetEvent.getVisibility,
    CalendarApp.Visibility.DEFAULT,
  );
  callWithRetryAndExponentialBackoff(targetEvent.removeAllReminders);
}

function createOrUpdateBufferEvent(
  primaryCalendar: GoogleAppsScript.Calendar.Calendar,
  previouslyCreatedEvents: GoogleAppsScript.Calendar.CalendarEvent[],
  event: GoogleAppsScript.Calendar.CalendarEvent,
  bufferType: "Pre" | "Post",
  secondaryCalendar: GoogleAppsScript.Calendar.Calendar,
): void {
  const bufferEventTitle = `${bufferType}-Buffer for ${event.getTitle()}`;
  const bufferForEventIdTag =
    bufferType === "Pre"
      ? PRE_BUFFER_FOR_EVENT_ID_TAG
      : POST_BUFFER_FOR_EVENT_ID_TAG;
  const bufferEventStartTime = new Date(
    bufferType === "Pre"
      ? event.getStartTime().getTime() - BUFFER_DURATION_MILLISECONDS
      : event.getEndTime().getTime(),
  );
  const bufferEventEndTime = new Date(
    bufferEventStartTime.getTime() + BUFFER_DURATION_MILLISECONDS,
  );

  const bufferEventIndex = previouslyCreatedEvents.findIndex(
    (previouslyCreatedEvent) =>
      previouslyCreatedEvent.getTag(ORIGINAL_CALENDAR_ID_TAG_KEY) ===
        secondaryCalendar.getId() &&
      previouslyCreatedEvent.getTag(bufferForEventIdTag) === event.getId(),
  );

  const bufferEvent: GoogleAppsScript.Calendar.CalendarEvent = (() => {
    if (bufferEventIndex === -1) {
      return callWithRetryAndExponentialBackoff(() =>
        primaryCalendar.createEvent(
          `${STAGING_TITLE_PREFIX} ${bufferEventTitle}`,
          bufferEventStartTime,
          bufferEventEndTime,
          {
            description: null,
            location: null,
          },
        ),
      );
    } else {
      const event = previouslyCreatedEvents[bufferEventIndex];
      previouslyCreatedEvents.splice(bufferEventIndex, 1);
      return event;
    }
  })();

  setTagIfNeeded(bufferEvent, SCRIPT_ID_TAG_KEY, SCRIPT_ID);
  setTagIfNeeded(
    bufferEvent,
    ORIGINAL_CALENDAR_ID_TAG_KEY,
    secondaryCalendar.getId(),
  );
  setTagIfNeeded(
    bufferEvent,
    bufferType === "Pre"
      ? PRE_BUFFER_FOR_EVENT_ID_TAG
      : POST_BUFFER_FOR_EVENT_ID_TAG,
    event.getId(),
  );

  setStartAndEndTimesIfNeeded(
    bufferEvent.setTime,
    () => new Date(bufferEvent.getStartTime().getTime()),
    () => new Date(bufferEvent.getEndTime().getTime()),
    bufferEventStartTime,
    bufferEventEndTime,
  );
  setEventAttributesIfNeeded(bufferEvent, event, secondaryCalendar, null, null);
  setIfNeeded(bufferEvent.setTitle, bufferEvent.getTitle, bufferEventTitle);
}

function cleanUpStagingEvents(
  primaryCalendar: GoogleAppsScript.Calendar.Calendar,
  today: Date,
  endDate: Date,
): void {
  for (const event of primaryCalendar
    .getEvents(today, endDate)
    .filter((event) => event.getTitle().startsWith(STAGING_TITLE_PREFIX))) {
    Logger.log(
      `[WARNING] Deleting orphaned staged event: ${event.getTitle()} (${event.getId()})`,
    );
    callWithRetryAndExponentialBackoff(event.deleteEvent);
  }
}

// Make main function available to Google Apps Script runtime
globalThis.main = main;
