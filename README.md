# sync-calendars

Google Apps script which syncs events from multiple external calendars into another calendar

# Setup

1. Install Node.js and npm.
2. Clone this repository.
3. Identify your primary and secondary calendar IDs.
    1. Your primary calendar ID is usually the email address of the account _to_ which you want to sync events..
    2. Your secondary calendar IDs are usually the email addresses of the accounts _from_ which you want to sync events.
4. Share the secondary calendars with your primary account.
5. Subscribe to the secondary calendars in your primary account.
6. Copy `appsscript.json.example` to `appsscript.json` and update the `"timeZone"` key with your timezone (e.g.,
   `"America/Los_Angeles"`, `"Europe/London"`, `"Asia/Tokyo"`).
7. Install dependencies: `npm install`
8. Log into Clasp and create a new Apps Script project: `npm run set-up`
9. Copy `.clasp.json.example` to `.clasp.json` and update the `"scriptId"` key with your script ID.
10. Enable the Google Apps Script API [here](https://script.google.com/home/usersettings).
11. Push the code to Google Apps Script: `npm run push`.
12. Configure Script Properties.
    1. Open your Apps Script Project: `npm run open`.
    2. Choose Project Settings (gear icon on the left).
    3. Under Script Properties, choose Edit script properties.
    4. Add the following properties:
        1. Property: `PRIMARY_CALENDAR_ID`, Value: Your primary calendar ID
        2. Property: `SECONDARY_CALENDAR_IDS`, Value: Comma-separated list of secondary calendar IDs
    5. Choose Save script properties.
13. Set up a trigger.
    1. Open your Apps Script Project: `npm run open`.
    2. Choose Triggers (clock icon on the left).
    3. Choose Add Trigger.
    4. Configure the trigger with the following properties:
        1. Choose which function to run: main
        2. Choose which deployment should run: Head
        3. Select event source: Time-driven
        4. Select type of time based trigger: Minutes timer
        5. Select minute interval: every minute
        6. Failure notification settings: as desired. It can get noisy with transient failures from Google, so I have
           mine set to notify me daily.