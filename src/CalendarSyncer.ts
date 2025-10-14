import { ColorCalculator } from "./ColorCalculator";

type Event = GoogleAppsScript.Calendar.Schema.Event & {
  id: string; // Required - guaranteed by the API
};

type SecondaryCalendar = GoogleAppsScript.Calendar.Schema.Calendar & {
  id: string; // Required - guaranteed by the API
  backgroundColor: string; // Required - guaranteed by the API
};

class BufferType {
  static readonly PRE_BUFFER_FOR_EVENT_ID_TAG = "preBufferForEventId";
  static readonly POST_BUFFER_FOR_EVENT_ID_TAG = "postBufferForEventId";
  static readonly BUFFER_DURATION_MILLISECONDS = 30 * 60 * 1000; // 30 minutes
  static readonly SUMMARY_WHEN_NOT_INCLUDING_SOURCE_EVENT_DETAILS =
    "Travel Time";

  static readonly PRE = new BufferType(
    "Pre",
    BufferType.PRE_BUFFER_FOR_EVENT_ID_TAG,
    "Pre-Buffer",
    (start: Date) =>
      new Date(start.getTime() - BufferType.BUFFER_DURATION_MILLISECONDS),
  );
  static readonly POST = new BufferType(
    "Post",
    BufferType.POST_BUFFER_FOR_EVENT_ID_TAG,
    "Post-Buffer",
    (_start: Date, end: Date) => new Date(end.getTime()),
  );
  static readonly ALL = [BufferType.PRE, BufferType.POST] as const;

  private constructor(
    readonly name: string,
    readonly tagKey: string,
    readonly titlePrefix: string,
    readonly calculateStartTime: (start: Date, end: Date) => Date,
  ) {}

  getTitle(mainEventSummary: string, includeDetails: boolean): string {
    return includeDetails
      ? `${this.titlePrefix} for ${mainEventSummary}`
      : BufferType.SUMMARY_WHEN_NOT_INCLUDING_SOURCE_EVENT_DETAILS;
  }
}

export class CalendarSyncer {
  private static readonly PRIMARY_CALENDAR_ID_KEY = "PRIMARY_CALENDAR_ID";
  private static readonly SECONDARY_CALENDAR_IDS_KEY = "SECONDARY_CALENDAR_IDS";
  private static readonly INCLUDE_SOURCE_EVENT_DETAILS_KEY =
    "INCLUDE_SOURCE_EVENT_DETAILS";
  private static readonly DAYS_LOOKAHEAD = 365;
  private static readonly SCRIPT_ID_TAG_KEY = "autoCreatedByScriptId";
  private static readonly ORIGINAL_CALENDAR_ID_TAG_KEY = "originalCalendarId";
  private static readonly ORIGINAL_EVENT_ID_TAG_KEY = "originalEventId";
  private static readonly INITIAL_BACKOFF_MILLISECONDS = 200;
  private static readonly MAX_RETRIES = 10;
  private static readonly SUMMARY_WHEN_NOT_INCLUDING_SOURCE_EVENT_DETAILS =
    "Appointment";

  private readonly Calendars: GoogleAppsScript.Calendar.Collection.CalendarsCollection;
  private readonly CalendarList: GoogleAppsScript.Calendar.Collection.CalendarListCollection;
  private readonly Events: GoogleAppsScript.Calendar.Collection.EventsCollection;

  private readonly colorCalculator = new ColorCalculator();

  private readonly scriptId;
  private readonly primaryCalendarId: string;
  private readonly secondaryCalendarIds: string[];
  private readonly includeSourceEventDetails: boolean;

  constructor() {
    this.scriptId = ScriptApp.getScriptId();

    if (!Calendar?.Calendars || !Calendar?.Events || !Calendar?.CalendarList) {
      throw new Error(
        "Calendar API is not available. Ensure the Calendar service is enabled.",
      );
    }
    this.Calendars = Calendar.Calendars;
    this.CalendarList = Calendar.CalendarList;
    this.Events = Calendar.Events;

    const scriptProperties = PropertiesService.getScriptProperties();

    const primaryCalendarId = scriptProperties.getProperty(
      CalendarSyncer.PRIMARY_CALENDAR_ID_KEY,
    );
    if (!primaryCalendarId) {
      throw new Error(
        `${CalendarSyncer.PRIMARY_CALENDAR_ID_KEY} not set. Add it in Project Settings > Script Properties.`,
      );
    }
    this.primaryCalendarId = primaryCalendarId;

    const secondaryCalendarIdsString = scriptProperties.getProperty(
      CalendarSyncer.SECONDARY_CALENDAR_IDS_KEY,
    );
    if (!secondaryCalendarIdsString) {
      throw new Error(
        `${CalendarSyncer.SECONDARY_CALENDAR_IDS_KEY} not set. Add it in Project Settings > Script Properties (comma-separated).`,
      );
    }
    this.secondaryCalendarIds = secondaryCalendarIdsString
      .split(",")
      .map((id) => id.trim());

    this.includeSourceEventDetails =
      scriptProperties.getProperty(
        CalendarSyncer.INCLUDE_SOURCE_EVENT_DETAILS_KEY,
      ) !== "false";
  }

  private static isAllDayEvent(event: Event): boolean {
    return !!event.start?.date;
  }

  private static callWithRetryAndExponentialBackoff<T>(
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
        attempt >= CalendarSyncer.MAX_RETRIES
      ) {
        throw error;
      }
      Utilities.sleep(
        CalendarSyncer.INITIAL_BACKOFF_MILLISECONDS * 2 ** attempt +
        Math.random() * 500,
      );
      return CalendarSyncer.callWithRetryAndExponentialBackoff(
        apiFunction,
        attempt + 1,
      );
    }
  }

  private static atMidnight(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  syncCalendars(): void {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(1)) {
      return;
    }

    try {
      const primaryCalendar = CalendarSyncer.callWithRetryAndExponentialBackoff(
        () => this.Calendars.get(this.primaryCalendarId),
      );
      if (!primaryCalendar) {
        throw new Error(
          `Could not access primary calendar: ${this.primaryCalendarId}`,
        );
      }

      const today = CalendarSyncer.atMidnight(new Date());
      const endDate = CalendarSyncer.atMidnight(
        new Date(
          today.getFullYear(),
          today.getMonth(),
          today.getDate() + CalendarSyncer.DAYS_LOOKAHEAD,
        ),
      );

      this.deleteOrphanedBufferEvents(today, endDate);

      const previouslyCreatedEvents = [
        ...this.getEventsInRange(
          this.primaryCalendarId,
          today,
          endDate,
          `${CalendarSyncer.SCRIPT_ID_TAG_KEY}=${this.scriptId}`,
        ),
      ];
      const processedPrimaryEventIds = new Set();

      this.getSecondaryCalendars().forEach((secondaryCalendar) => {
        [
          ...this.getEventsInRange(secondaryCalendar.id, today, endDate),
        ].forEach((secondaryMainEvent) => {
          const primaryMainEvent = this.createOrUpdateMainEvent(
            previouslyCreatedEvents,
            secondaryCalendar,
            secondaryMainEvent,
          );
          processedPrimaryEventIds.add(primaryMainEvent.id);

          if (!CalendarSyncer.isAllDayEvent(secondaryMainEvent)) {
            [BufferType.PRE, BufferType.POST].forEach((bufferType) =>
              processedPrimaryEventIds.add(
                this.createOrUpdateBufferEvent(
                  previouslyCreatedEvents,
                  primaryMainEvent,
                  bufferType,
                  secondaryCalendar,
                ),
              ),
            );
          }
        });
      });

      previouslyCreatedEvents
        .filter((event) => !(event.id in processedPrimaryEventIds))
        .forEach((event) => {
          CalendarSyncer.callWithRetryAndExponentialBackoff(() =>
            this.Events.remove(this.primaryCalendarId, event.id),
          );
        });
    } finally {
      lock.releaseLock();
    }
  }

  private getSecondaryCalendars(): SecondaryCalendar[] {
    // Calendar background colors are only returned in the CalendarList API, so we have to call it separately and merge it with the regular calendar data
    const calendarListEntriesById =
      CalendarSyncer.callWithRetryAndExponentialBackoff(() =>
        this.CalendarList.list(),
      )?.items?.reduce(
        (calendarListEntriesById, calendarListEntry) => {
          if (calendarListEntry.id)
            calendarListEntriesById[calendarListEntry.id!] = calendarListEntry;
          return calendarListEntriesById;
        },
        {} as Record<
          string,
          GoogleAppsScript.Calendar.Schema.CalendarListEntry
        >,
      ) ?? {};

    return this.secondaryCalendarIds
      .map((calendarId): SecondaryCalendar | null => {
        const calendar = CalendarSyncer.callWithRetryAndExponentialBackoff(() =>
          this.Calendars.get(calendarId),
        );

        if (!calendar || !(calendarId in calendarListEntriesById)) {
          Logger.log(`Could not access secondary calendar: ${calendarId}`);
          return null;
        }

        return {
          ...calendar,
          id: calendar.id!,
          backgroundColor: calendarListEntriesById[calendarId].backgroundColor!,
        };
      })
      .filter((calendar) => calendar !== null);
  }

  private *getEventsInRange(
    calendarId: string,
    startDate: Date,
    endDate: Date,
    privateExtendedProperty?: string,
  ): Generator<Event, void, unknown> {
    let pageToken: string | undefined;
    const startDateIsoString = startDate.toISOString();
    const endDateIsoString = endDate.toISOString();

    do {
      const response = CalendarSyncer.callWithRetryAndExponentialBackoff(() =>
        this.Events.list(calendarId, {
          timeMin: startDateIsoString,
          timeMax: endDateIsoString,
          singleEvents: true,
          orderBy: "startTime",
          maxResults: 2500,
          pageToken,
          privateExtendedProperty,
        }),
      );

      for (const item of response.items ?? []) {
        yield item as Event;
      }

      pageToken = response.nextPageToken;
    } while (pageToken);
  }

  private buildEventProperties(
    summary: string | undefined,
    start: GoogleAppsScript.Calendar.Schema.EventDateTime,
    end: GoogleAppsScript.Calendar.Schema.EventDateTime,
    transparency: string | undefined,
    colorId: string,
    privateExtendedProperties: Record<string, string>,
    description?: string,
    location?: string,
  ) {
    return {
      summary,
      description,
      location,
      start,
      end,
      transparency,
      visibility: "default",
      guestsCanInviteOthers: false,
      guestsCanModify: false,
      guestsCanSeeOtherGuests: false,
      anyoneCanAddSelf: false,
      reminders: { useDefault: false, overrides: [] },
      extendedProperties: {
        private: privateExtendedProperties,
      },
      colorId,
    };
  }

  private createOrUpdateMainEvent(
    previouslyCreatedEvents: Event[],
    secondaryCalendar: SecondaryCalendar,
    secondaryMainEvent: Event,
  ): Event {
    const eventProperties = this.buildEventProperties(
      this.includeSourceEventDetails
        ? (secondaryMainEvent.summary ?? undefined)
        : CalendarSyncer.SUMMARY_WHEN_NOT_INCLUDING_SOURCE_EVENT_DETAILS,
      secondaryMainEvent.start!,
      secondaryMainEvent.end!,
      secondaryMainEvent.transparency,
      secondaryMainEvent.colorId ??
      this.colorCalculator.getClosestColorId(
        secondaryCalendar.backgroundColor,
      ),
      {
        [CalendarSyncer.SCRIPT_ID_TAG_KEY]: this.scriptId,
        [CalendarSyncer.ORIGINAL_CALENDAR_ID_TAG_KEY]: secondaryCalendar.id,
        [CalendarSyncer.ORIGINAL_EVENT_ID_TAG_KEY]: secondaryMainEvent.id,
      },
      this.includeSourceEventDetails
        ? secondaryMainEvent.description
        : undefined,
      this.includeSourceEventDetails ? secondaryMainEvent.location : undefined,
    );

    const primaryMainEvent = previouslyCreatedEvents.find(
      (event) =>
        event.extendedProperties?.private?.[
          CalendarSyncer.ORIGINAL_CALENDAR_ID_TAG_KEY
          ] === secondaryCalendar.id &&
        event.extendedProperties?.private?.[
          CalendarSyncer.ORIGINAL_EVENT_ID_TAG_KEY
          ] === secondaryMainEvent.id,
    );

    return CalendarSyncer.callWithRetryAndExponentialBackoff(
      primaryMainEvent
        ? () =>
          this.Events.patch(
            eventProperties,
            this.primaryCalendarId,
            primaryMainEvent.id,
          )
        : () => this.Events.insert(eventProperties, this.primaryCalendarId),
    ) as Event;
  }

  private createOrUpdateBufferEvent(
    previouslyCreatedEvents: Event[],
    primaryMainEvent: Event,
    bufferType: BufferType,
    secondaryCalendar: SecondaryCalendar,
  ) {
    const bufferEventStartTime = bufferType.calculateStartTime(
      // we know the start and end dateTimes are set because we validated that these are not all-day events
      new Date(primaryMainEvent.start!.dateTime!),
      new Date(primaryMainEvent.end!.dateTime!),
    );

    const eventProperties = this.buildEventProperties(
      bufferType.getTitle(
        primaryMainEvent.summary ?? "",
        this.includeSourceEventDetails,
      ),
      { dateTime: bufferEventStartTime.toISOString() },
      {
        dateTime: new Date(
          bufferEventStartTime.getTime() +
          BufferType.BUFFER_DURATION_MILLISECONDS,
        ).toISOString(),
      },
      primaryMainEvent.transparency,
      primaryMainEvent.colorId ??
      this.colorCalculator.getClosestColorId(
        secondaryCalendar.backgroundColor,
      ),
      {
        [CalendarSyncer.SCRIPT_ID_TAG_KEY]: this.scriptId,
        [CalendarSyncer.ORIGINAL_CALENDAR_ID_TAG_KEY]: secondaryCalendar.id,
        [bufferType.tagKey]: primaryMainEvent.id,
      },
    );

    const existingBufferEvent = previouslyCreatedEvents.find(
      (event) =>
        event.extendedProperties?.private?.[
          CalendarSyncer.ORIGINAL_CALENDAR_ID_TAG_KEY
          ] === secondaryCalendar.id &&
        event.extendedProperties?.private?.[bufferType.tagKey] ===
        primaryMainEvent.id,
    );

    return CalendarSyncer.callWithRetryAndExponentialBackoff(
      existingBufferEvent
        ? () =>
          this.Events.patch(
            eventProperties,
            this.primaryCalendarId,
            existingBufferEvent.id,
          )
        : () => this.Events.insert(eventProperties, this.primaryCalendarId),
    ) as Event;
  }

  private deleteOrphanedBufferEvents(startDate: Date, endDate: Date): void {
    [...this.getEventsInRange(this.primaryCalendarId, startDate, endDate)]
      .filter(
        (event) =>
          !event.extendedProperties?.private?.[
            CalendarSyncer.SCRIPT_ID_TAG_KEY
            ] &&
          (BufferType.ALL.some((bufferType) =>
              event.summary?.startsWith(bufferType.titlePrefix),
            ) ||
            event.summary ===
            BufferType.SUMMARY_WHEN_NOT_INCLUDING_SOURCE_EVENT_DETAILS),
      )
      .forEach((event) => {
        if (event.id) {
          CalendarSyncer.callWithRetryAndExponentialBackoff(() =>
            this.Events.remove(this.primaryCalendarId, event.id!),
          );
        }
      });
  }
}
