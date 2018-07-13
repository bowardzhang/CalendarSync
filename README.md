# CalendarSync
Synchronize MS Outlook calendar to Google calendar

Author: Bo Zhang <bowardzhang@gmail.com>

### Features
1. Can be configured to hide subject as "%COMPANY% busy" for security reasons.
2. Deal with recurring events and exceptions in a reccurring series.

### Preparation
1.  Visit the [Google API Console](https://console.developers.google.com/) and create a new project.
2.  Under Library, enable the  [Google Calendar API](https://console.developers.google.com/apis/api/calendar-json.googleapis.com/overview).
3.  Under Credentials, create a new OAuth client ID for application type Other and download it as a JSON file.
4.  Rename the downloaded JSON file to  `credentials.json`  and place it in the Python script folder.
