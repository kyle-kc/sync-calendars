# sync-calendars
Google Apps script which syncs events from multiple external calendars into another calendar

First, share the secondary calendars with your primary account. Then, in your primary account, subscribe to the calendars.

Then, copy the code in a Google Apps script and edit the calendar ID constants. Then you can add a trigger to run on a schedule. I have mine running every minute.

The script is actually fairly smart and self-recovers from most errors. It works for me, but use at your own risk.
