# calsync
GSuite Calendar Syncing, a hacky way to go about getting your Gsuite Work Calendar on your WearOS device if your organization doesn't allow sharing calendar with your personal profile.

## Steps
* Create a new calendar in your personal account, and share it with your work account with write permissions.
* Run this script in your work account's GScripts (scripts.google.com) to sync your work calendar with the shared personal calendar.
  * Install the script, i.e, copy/paste/save as a new .gs script.
  * Run the function `createSpreadsheet()` first.
    * This will log a spreadsheet doc ID in your script logs.
    * Copy the DocID, and update the sheetID variable on top of the script. (Like I said, this is pretty hacky)
    * Also, make sure to update the calendar ID on top of the script with your newly minted personal calendar ID
  * Now, schedule the function `syncCalendars()` to run at the frequency with which you want the calendars to stay in sync.

## TODO / Eventual Improvements
* Automagically generate spreadsheet if one doesn't exist and automagically search for your shared personal calendar.
