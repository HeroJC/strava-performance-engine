# Strava Performance Engine v4.0

A Google Apps Script for analyzing inline speed skating performance data from Strava. It pulls your recent activities and segments into a Google Sheet to track your velocity maintenance, heart rate performance, and race pacing.

## How to use this project

Instead of hard-coding your secrets into the script, this project includes a secure setup menu right inside Google Sheets. Anyone who makes a copy of the sheet can use it with their own Strava account.

### Step 1: Create a Google Sheet
1. Create a new Google Sheet at [sheets.new](https://sheets.new).
2. Go to **Extensions > Apps Script**.
3. Copy the entire contents of `Code.gs` from this repository and paste it into the script editor, replacing any existing code.
4. Click the **Save** icon (or press Ctrl+S / Cmd+S).
5. Refresh your Google Sheet. You should now see a new **🏃‍♂️ Strava** menu at the top.

### Step 2: Get your Strava API Developer Keys
To pull data from Strava, you need to create a free "API Application" on your Strava account.

1. Log into Strava on your computer.
2. Go to the API settings page: [https://www.strava.com/settings/api](https://www.strava.com/settings/api)
3. If you've never created an app before, fill out the form:
   - **Application Name:** (e.g., "My Skate Tracker")
   - **Category:** "Data Importer"
   - **Website:** (Any valid URL, e.g., `https://google.com`)
   - **Authorization Callback Domain:** `localhost` *(This is important!)*
4. Once created, look for your **Client ID** and **Client Secret**. Keep this page open.

### Step 3: Connect the Script
1. Go back to your Google Sheet.
2. Click **🏃‍♂️ Strava > 1. Setup API Credentials**.
3. Paste in your **Client ID** and **Client Secret** when prompted.
4. Next, click **🏃‍♂️ Strava > 2. Authorize Strava**. A dialog will appear with an authorization link.
5. Click the link. Strava will ask if you want to authorize the app. Click Authorize.
6. **IMPORTANT:** You will be redirected to an error page that says the site can't be reached (localhost). **This is totally normal!**
7. Look at the URL bar in your browser. It will look something like this:
   `http://localhost/exchange_token?state=&code=abc123def456ghi789jkl012mno345pqr678stu90&scope=read,activity:read_all`
8. Copy the long string of letters and numbers immediately after **`code=`** and before the next `&` symbol.
9. Go back to your Google Sheet and click **🏃‍♂️ Strava > 3. Complete Authorization**, then paste that code.

### Step 4: Run It!
You're done! You can now click **🏃‍♂️ Strava > Sync Strava Data** at any time to pull your latest 15 activities and their segments into the spreadsheet.

## Privacy & Security Note
* Your Strava API credentials are saved securely to your Google Account's hidden `PropertiesService`.
* No one else who looks at the `Code.gs` file can see your Client ID, Secret, or Refresh Token.
* If you want to remove the script's access, use the **🏃‍♂️ Strava > Clear Credentials** menu option, or revoke access from your Strava settings page.
