const PRIMARY_CALENDAR_ID = "YOUR_CALENDAR_ID";  // probably your email

const SECONDARY_CALENDAR_IDS = ["OTHER_CALENDAR_1", "OTHER_CALENDAR_2"]; // probably other emails

const DAYS_LOOKAHEAD = 365;
const SCRIPT_ID_TAG_KEY = "autoCreatedByScriptId";
const ORIGINAL_CALENDAR_ID_TAG_KEY = "originalCalendarId";
const ORIGINAL_EVENT_ID_TAG_KEY = "originalEventId";
const PRE_BUFFER_FOR_EVENT_ID_TAG = "preBufferForEventId";
const POST_BUFFER_FOR_EVENT_ID_TAG = "postBufferForEventId";
const SCRIPT_ID = ScriptApp.getScriptId();

const EVENT_COLORS_TO_HEX_CODES = new Map([
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

const hexCodeToClosestEventColorCache = new Map();

Date.prototype.setTimeToMidnight = function () {
  this.setHours(0);
  this.setMinutes(0);
  this.setSeconds(0);
  this.setMilliseconds(0);
  return this;
};

function hexToRgb(hex) {
  let bigint = parseInt(hex.substring(1), 16);
  return [(bigint >> 16) & 255, (bigint >> 8) & 255, bigint & 255];
}

function rgbToXyz(rgb) {
  let r = rgb[0] / 255, g = rgb[1] / 255, b = rgb[2] / 255;

  r = r > 0.04045 ? Math.pow((r + 0.055) / 1.055, 2.4) : r / 12.92;
  g = g > 0.04045 ? Math.pow((g + 0.055) / 1.055, 2.4) : g / 12.92;
  b = b > 0.04045 ? Math.pow((b + 0.055) / 1.055, 2.4) : b / 12.92;

  r *= 100; g *= 100; b *= 100;

  return [
    r * 0.4124564 + g * 0.3575761 + b * 0.1804375,
    r * 0.2126729 + g * 0.7151522 + b * 0.0721750,
    r * 0.0193339 + g * 0.1191920 + b * 0.9503041
  ];
}

function xyzToLab(xyz) {
  let x = xyz[0] / 95.047, y = xyz[1] / 100.000, z = xyz[2] / 108.883;

  x = x > 0.008856 ? Math.pow(x, 1 / 3) : (7.787 * x) + (16 / 116);
  y = y > 0.008856 ? Math.pow(y, 1 / 3) : (7.787 * y) + (16 / 116);
  z = z > 0.008856 ? Math.pow(z, 1 / 3) : (7.787 * z) + (16 / 116);

  return [(116 * y) - 16, 500 * (x - y), 200 * (y - z)];
}

function colorDistance(rgb1, rgb2) {
  return Math.sqrt(
    Math.pow(rgb1[0] - rgb2[0], 2) +
    Math.pow(rgb1[1] - rgb2[1], 2) +
    Math.pow(rgb1[2] - rgb2[2], 2)
  );
}

function deltaE76(lab1, lab2) {
  return Math.sqrt(
    Math.pow(lab1[0] - lab2[0], 2) +
    Math.pow(lab1[1] - lab2[1], 2) +
    Math.pow(lab1[2] - lab2[2], 2)
  );
}

function getClosestEventColor(hesCode) {
  if (!hexCodeToClosestEventColorCache.has(hesCode)) {
    let targetLab = xyzToLab(rgbToXyz(hexToRgb(hesCode)));
    let closestColor = null;
    let minDistance = Number.MAX_VALUE;

    EVENT_COLORS_TO_HEX_CODES.forEach((hexColor, eventColor) => {
      let colorLab = xyzToLab(rgbToXyz(hexToRgb(hexColor)));
      let distance = deltaE76(targetLab, colorLab);
      if (distance < minDistance) {
        minDistance = distance;
        closestColor = eventColor;
      }
    });

    hexCodeToClosestEventColorCache.set(hesCode, closestColor);
  }
  return hexCodeToClosestEventColorCache.get(hesCode);
}

function callWithRetryAndExponentialBackoff(apiFunction) {
  let numberOfTries = 0;

  while (true) {
    try {
      return apiFunction();
    } catch (exception) {
      if (exception.message.includes("You have been creating or deleting too many calendars or calendar events in a short time. Please try again later.") && numberOfTries <= MAX_RETRIES) {
        Utilities.sleep(INITIAL_BACKOFF_MILLISECONDS * Math.pow(2, numberOfTries));
        numberOfTries++;
      } else {
        throw exception;
      }
    }
  }
}

function setIfNeeded(setMethod, getMethod, newValue) {
  currentValue = getMethod()
  if ((currentValue || newValue) && currentValue !== newValue) {
    callWithRetryAndExponentialBackoff(() => setMethod(newValue));
  }
}

function setTagIfNeeded(event, tag_key, tag_value) {
  if (event.getTag(tag_key) !== tag_value) {
    callWithRetryAndExponentialBackoff(() => event.setTag(tag_key, tag_value));
  }
}

function setStartAndEndTimesIfNeeded(setMethod, getStartTimeMethod, getEndTimeMethod, startTime, endTime) {
  if (getStartTimeMethod().getTime() != startTime.getTime() || getEndTimeMethod().getTime() != endTime.getTime()) {
    callWithRetryAndExponentialBackoff(() => setMethod(startTime, endTime));
  }
}

function setEventAttributesIfNeeded(targetEvent, sourceEvent, sourceCalendar, description = sourceEvent.getDescription(), location = sourceEvent.getLocation(), title = sourceEvent.getTitle()) {
  setIfNeeded(targetEvent.setColor, () => targetEvent.getColor(), sourceEvent.getColor() || getClosestEventColor(sourceCalendar.getColor()));
  setIfNeeded(targetEvent.setAnyoneCanAddSelf, () => targetEvent.anyoneCanAddSelf(), false);
  setIfNeeded(targetEvent.setDescription, () => targetEvent.getDescription(), description);
  setIfNeeded(targetEvent.setGuestsCanInviteOthers, () => targetEvent.guestsCanInviteOthers(), false);
  setIfNeeded(targetEvent.setGuestsCanModify, () => targetEvent.guestsCanModify(), false);
  setIfNeeded(targetEvent.setGuestsCanSeeGuests, () => targetEvent.guestsCanSeeGuests(), false);
  setIfNeeded(targetEvent.setLocation, () => targetEvent.getLocation(), location);
  setIfNeeded(targetEvent.setTitle, () => targetEvent.getTitle(), title);
  setIfNeeded(targetEvent.setTransparency, () => targetEvent.getTransparency(), sourceEvent.getTransparency());
  setIfNeeded(targetEvent.setVisibility, () => targetEvent.getVisibility(), CalendarApp.Visibility.DEFAULT);
  if ([...targetEvent.getEmailReminders(), ...targetEvent.getPopupReminders(), ...targetEvent.getSmsReminders()].length > 0) {
    callWithRetryAndExponentialBackoff(targetEvent.removeAllReminders);
  }
}

function createOrUpdateBufferEvent(primaryCalendar, previouslyCreatedEvents, event, bufferType, secondaryCalendar) {
  const bufferEventTitle = `${bufferType}-Buffer for ${event.getTitle()}`;
  const bufferForEventIdTag = bufferType === "Pre" ? PRE_BUFFER_FOR_EVENT_ID_TAG : POST_BUFFER_FOR_EVENT_ID_TAG
  const bufferEventStartTime = new Date(
    bufferType === "Pre" ? event.getStartTime().getTime() - BUFFER_DURATION_MILLISECONDS : event.getEndTime().getTime()
  );
  const bufferEventEndTime = new Date(bufferEventStartTime.getTime() + BUFFER_DURATION_MILLISECONDS);

  const bufferEventIndex = previouslyCreatedEvents.findIndex(
    (previouslyCreatedEvent) => previouslyCreatedEvent.getTag(ORIGINAL_CALENDAR_ID_TAG_KEY) === secondaryCalendar.getId() && previouslyCreatedEvent.getTag(bufferForEventIdTag) === event.getId()
  );

  let bufferEvent;
  if (bufferEventIndex === -1) {
    bufferEvent = callWithRetryAndExponentialBackoff(() => primaryCalendar.createEvent(bufferEventTitle, bufferEventStartTime, bufferEventEndTime, {
      description: null,
      location: null,
    }));
  } else {
    bufferEvent = previouslyCreatedEvents[bufferEventIndex];
    previouslyCreatedEvents.splice(bufferEventIndex, 1);
  }

  setTagIfNeeded(bufferEvent, SCRIPT_ID_TAG_KEY, SCRIPT_ID);
  setTagIfNeeded(bufferEvent, ORIGINAL_CALENDAR_ID_TAG_KEY, secondaryCalendar.getId());
  setTagIfNeeded(bufferEvent, bufferType === "Pre" ? PRE_BUFFER_FOR_EVENT_ID_TAG : POST_BUFFER_FOR_EVENT_ID_TAG, event.getId());

  setStartAndEndTimesIfNeeded(bufferEvent.setTime, bufferEvent.getStartTime, bufferEvent.getEndTime, bufferEventStartTime, bufferEventEndTime);
  setEventAttributesIfNeeded(bufferEvent, event, secondaryCalendar, null, null, bufferEventTitle);
}

function main() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1)) {
    return;
  }

  try {
    const today = new Date();
    today.setTimeToMidnight();
    const endDate = new Date();
    endDate.setDate(today.getDate() + DAYS_LOOKAHEAD);
    endDate.setTimeToMidnight();

    const primaryCalendar = callWithRetryAndExponentialBackoff(() => CalendarApp.getCalendarById(PRIMARY_CALENDAR_ID));

    const previouslyCreatedEvents = callWithRetryAndExponentialBackoff(() => primaryCalendar.getEvents(today, endDate)).filter(event => event.getTag(SCRIPT_ID_TAG_KEY) === SCRIPT_ID)

    for (secondaryCalendarId of SECONDARY_CALENDAR_IDS) {
      const secondaryCalendar = callWithRetryAndExponentialBackoff(() => CalendarApp.getCalendarById(secondaryCalendarId));

      for (const secondaryEvent of callWithRetryAndExponentialBackoff(() => secondaryCalendar.getEvents(today, endDate))) {
        primaryEventIndex = previouslyCreatedEvents.findIndex(event => event.getTag(ORIGINAL_CALENDAR_ID_TAG_KEY) === secondaryCalendarId && event.getTag(ORIGINAL_EVENT_ID_TAG_KEY) == secondaryEvent.getId());
        if (primaryEventIndex === -1) {
          if (secondaryEvent.isAllDayEvent()) {
            primaryEvent = primaryCalendar.createAllDayEvent(
              secondaryEvent.getTitle(),
              secondaryEvent.getAllDayStartDate(),
              secondaryEvent.getAllDayEndDate(),
              {
                description: secondaryEvent.getDescription(),
                location: secondaryEvent.getLocation(),
              }
            );
          }
          else {
            primaryEvent = primaryCalendar.createEvent(
              secondaryEvent.getTitle(),
              secondaryEvent.getStartTime(),
              secondaryEvent.getEndTime(),
              {
                description: secondaryEvent.getDescription(),
                location: secondaryEvent.getLocation(),
              }
            );
          }
        } else {
          primaryEvent = previouslyCreatedEvents[primaryEventIndex];
          previouslyCreatedEvents.splice(primaryEventIndex, 1);
        }

        setTagIfNeeded(primaryEvent, SCRIPT_ID_TAG_KEY, SCRIPT_ID);
        setTagIfNeeded(primaryEvent, ORIGINAL_CALENDAR_ID_TAG_KEY, secondaryCalendarId);
        setTagIfNeeded(primaryEvent, ORIGINAL_EVENT_ID_TAG_KEY, secondaryEvent.getId());

        setEventAttributesIfNeeded(primaryEvent, secondaryEvent, secondaryCalendar);

        if (secondaryEvent.isAllDayEvent()) {
          setStartAndEndTimesIfNeeded(primaryEvent.setAllDayDates, () => primaryEvent.isAllDayEvent() ? primaryEvent.getAllDayStartDate() : null, () => primaryEvent.isAllDayEvent() ? primaryEvent.getAllDayEndDate() : null, secondaryEvent.getAllDayStartDate(), secondaryEvent.getAllDayEndDate())
        }
        else {
          setStartAndEndTimesIfNeeded(primaryEvent.setTime, primaryEvent.getStartTime, primaryEvent.getEndTime, secondaryEvent.getStartTime(), secondaryEvent.getEndTime());
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
