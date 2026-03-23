/**
 * Strava Performance Engine v4.0
 * Features: Secure user credentials via PropertiesService and custom UI menu.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏃‍♂️ Strava')
      .addItem('1. Setup API Credentials', 'setupCredentials')
      .addItem('2. Authorize Strava', 'showAuthUrl')
      .addItem('3. Complete Authorization', 'completeAuth')
      .addSeparator()
      .addItem('Sync Strava Data', 'syncStravaData')
      .addItem('Clear Credentials', 'clearCredentials')
      .addToUi();
}

function setupCredentials() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
  
  const clientIdResponse = ui.prompt('Strava Setup (Step 1 of 3)', 'Enter your Strava Client ID:', ui.ButtonSet.OK_CANCEL);
  if (clientIdResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const clientSecretResponse = ui.prompt('Strava Setup (Step 1 of 3)', 'Enter your Strava Client Secret:', ui.ButtonSet.OK_CANCEL);
  if (clientSecretResponse.getSelectedButton() !== ui.Button.OK) return;

  props.setProperty('CLIENT_ID', clientIdResponse.getResponseText().trim());
  props.setProperty('CLIENT_SECRET', clientSecretResponse.getResponseText().trim());
  
  ui.alert('Success', 'Credentials saved! Now click "2. Authorize Strava" from the menu.', ui.ButtonSet.OK);
}

function showAuthUrl() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
  const clientId = props.getProperty('CLIENT_ID');
  
  if (!clientId) {
    ui.alert('Error', 'Please run "1. Setup API Credentials" first.', ui.ButtonSet.OK);
    return;
  }
  
  const authUrl = `https://www.strava.com/oauth/authorize?client_id=${clientId}&response_type=code&redirect_uri=http://localhost/exchange_token&approval_prompt=force&scope=activity:read_all`;
  
  const htmlOutput = HtmlService
    .createHtmlOutput(`<div style="font-family: Arial, sans-serif;">
                       <p>Click the link below to authorize this script to access your Strava data.</p>
                       <p><a href="${authUrl}" target="_blank" style="background-color: #fc4c02; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; display: inline-block;">Authorize with Strava</a></p>
                       <p><b>Important:</b></p>
                       <ol>
                         <li>After clicking, you will be redirected to an error page (localhost). This is expected!</li>
                         <li>Look at the URL in your browser. Copy the text immediately after <b>code=</b> (and before any & symbol).</li>
                         <li>Close that tab and run "3. Complete Authorization" in the menu.</li>
                       </ol>
                       </div>`)
    .setWidth(450)
    .setHeight(300);
    
  ui.showModalDialog(htmlOutput, 'Authorize Strava (Step 2 of 3)');
}

function completeAuth() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
  
  const clientId = props.getProperty('CLIENT_ID');
  const clientSecret = props.getProperty('CLIENT_SECRET');
  
  if (!clientId || !clientSecret) {
    ui.alert('Error', 'Missing Client ID or Secret. Please run Step 1.', ui.ButtonSet.OK);
    return;
  }

  const codeResponse = ui.prompt('Complete Auth (Step 3 of 3)', 'Paste the "code" from the redirect URL:', ui.ButtonSet.OK_CANCEL);
  if (codeResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const code = codeResponse.getResponseText().trim();
  
  try {
    const tokenUrl = `https://www.strava.com/oauth/token`;
    const payload = {
      client_id: clientId,
      client_secret: clientSecret,
      code: code,
      grant_type: 'authorization_code'
    };
    
    const response = UrlFetchApp.fetch(tokenUrl, {
      method: 'post',
      payload: payload
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.refresh_token) {
      props.setProperty('REFRESH_TOKEN', data.refresh_token);
      ui.alert('Success 🎉', 'Authorization complete! You can now use "Sync Strava Data".', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Failed to get refresh token. ' + JSON.stringify(data), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Authentication failed. Please make sure you copied the correct code. Details: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearCredentials() {
  PropertiesService.getUserProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert('Success', 'All Strava credentials have been removed from this script.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncStravaData() {
  const props = PropertiesService.getUserProperties();
  const clientId = props.getProperty('CLIENT_ID');
  const clientSecret = props.getProperty('CLIENT_SECRET');
  let refreshToken = props.getProperty('REFRESH_TOKEN');

  if (!clientId || !clientSecret || !refreshToken) {
    SpreadsheetApp.getUi().alert('Configuration Missing', 'Please complete the Strava setup menu first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const summaryHeaders = ['Date', 'Neighborhood', 'Name', 'Dist (mi)', 'Avg MPH', 'Max MPH', 'Suffer Score', 'Predicted 26.2 Time', 'Map Polyline'];
  const segmentHeaders = [
    'Date', 'Segment Name', 'Avg MPH', 'Avg HR', 
    'Velocity Maintenance %', 'Target Gap Ratio', 'Aerobic Power (S/HR)', 'Segment ID'
  ];

  const summarySheet = checkAndCreateSheet(ss, 'Summary_Data', summaryHeaders);
  const segmentSheet = checkAndCreateSheet(ss, 'Segment_Data', segmentHeaders);

  // Clear data for fresh sync
  if (summarySheet.getLastRow() > 1) summarySheet.getRange(2, 1, summarySheet.getLastRow(), summaryHeaders.length).clearContent();
  if (segmentSheet.getLastRow() > 1) segmentSheet.getRange(2, 1, segmentSheet.getLastRow(), segmentHeaders.length).clearContent();

  try {
    const tokenUrl = `https://www.strava.com/oauth/token?client_id=${clientId}&client_secret=${clientSecret}&refresh_token=${refreshToken}&grant_type=refresh_token`;
    const tokenResponse = UrlFetchApp.fetch(tokenUrl, { method: 'post' });
    const tokenData = JSON.parse(tokenResponse.getContentText());
    
    // Update refresh token if Strava rotated it
    if (tokenData.refresh_token && tokenData.refresh_token !== refreshToken) {
      props.setProperty('REFRESH_TOKEN', tokenData.refresh_token);
    }
    
    const accessToken = tokenData.access_token;

    // Fetching last 15 activities
    const activities = JSON.parse(UrlFetchApp.fetch('https://www.strava.com/api/v3/athlete/activities?per_page=15', {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    }).getContentText());

    activities.forEach(activity => {
      const detail = JSON.parse(UrlFetchApp.fetch(`https://www.strava.com/api/v3/activities/${activity.id}`, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      }).getContentText());

      let neighborhood = "NYC Area";
      if (detail.start_latlng) {
        neighborhood = fetchFreeNeighborhood(detail.start_latlng[0], detail.start_latlng[1]);
        Utilities.sleep(1000); // Respect OSM rate limits
      }

      const avgMph = (detail.average_speed * 2.23694);
      
      // Predicted 26.2 Time
      const hours = 26.2 / avgMph;
      const h = Math.floor(hours);
      const m = Math.floor((hours - h) * 60);
      const predictedTime = `${h}h ${m}m`;

      summarySheet.appendRow([
        new Date(detail.start_date_local).toLocaleDateString(),
        neighborhood,
        detail.name,
        (detail.distance * 0.000621371).toFixed(2),
        avgMph.toFixed(1),
        (detail.max_speed * 2.23694).toFixed(1),
        detail.suffer_score || 0,
        predictedTime,
        detail.map ? detail.map.summary_polyline : "N/A"
      ]);

      if (detail.segment_efforts) {
        detail.segment_efforts.forEach(effort => {
          const mph = (effort.average_speed * 2.23694);
          
          const hr = (effort.average_heartrate && !isNaN(effort.average_heartrate)) ? effort.average_heartrate : 0;
          
          const velocityMaintenance = ((mph / 21.0) * 100).toFixed(1);
          const targetGap = (mph / 21.0).toFixed(2);
          
          const aerobicPower = hr > 0 ? (mph / hr).toFixed(3) : "N/A";

          segmentSheet.appendRow([
            new Date(detail.start_date_local).toLocaleDateString(),
            effort.name,
            mph.toFixed(2),
            hr > 0 ? hr.toFixed(0) : "N/A",
            velocityMaintenance + "%",
            targetGap,
            aerobicPower,
            effort.segment.id
          ]);
        });
      }
    });

  } catch (e) {
    SpreadsheetApp.getUi().alert('Sync Error', 'An error occurred during sync: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function fetchFreeNeighborhood(lat, lng) {
  try {
    const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${lat}&lon=${lng}`;
    const params = { 'headers': { 'User-Agent': 'StravaSkateTracker/1.0 (jeff@example.com)' } };
    const response = JSON.parse(UrlFetchApp.fetch(url, params).getContentText());
    return response.address.neighbourhood || response.address.suburb || "NYC Area";
  } catch (e) { return "NYC Area"; }
}

function checkAndCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
  }
  return sheet;
}
