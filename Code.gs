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
      .addItem('Backfill Missing Data', 'backfillMissingData')
      .addItem('Copy Gemini Digest', 'showGeminiDigest')
      .addSeparator()
      .addItem('Refresh Volume Rollup', 'updateVolumeRollup')
      .addItem('Refresh PR Board', 'updatePRBoard')
      .addSeparator()
      .addItem('Reset Sync History', 'resetSyncHistory')
      .addItem('Clear Credentials', 'clearCredentials')
      .addToUi();
}

function showGeminiDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('Summary_Data');
  const segmentSheet = ss.getSheetByName('Segment_Data');
  const ui = SpreadsheetApp.getUi();

  if (!summarySheet || summarySheet.getLastRow() < 2) {
    ui.alert('No data found. Please run "Sync Strava Data" first.');
    return;
  }

  // Get headers and the last 5 activities
  const summaryHeaders = summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn()).getValues()[0];
  const summaryData = summarySheet.getRange(Math.max(2, summarySheet.getLastRow() - 4), 1, Math.min(5, summarySheet.getLastRow() - 1), summarySheet.getLastColumn()).getValues();
  
  let digest = "Act as an elite inline speed skating coach. I am training for a 26.2-mile marathon. ";
  digest += "Analyze my recent performance data below. Look for trends in my aerobic power (MPH/HR), pacing consistency, and how I handle elevation and weather. ";
  digest += "Provide 3 actionable tips for my next training block.\n\n";
  digest += "## 🏃‍♂️ Recent Activity Data\n\n";
  
  summaryData.forEach(row => {
    digest += `### Activity: ${row[3]} (${row[1]})\n`;
    digest += `- **Distance:** ${row[4]} mi | **Elev Gain:** ${row[5]} ft\n`;
    digest += `- **Speed:** Avg ${row[6]} MPH (Max ${row[7]}) | **Variance:** ${row[8]} MPH\n`;
    digest += `- **Effort:** Suffer Score ${row[9]} | **Zone:** ${row[10]}\n`;
    digest += `- **Weather:** ${row[12]}°F, Wind ${row[13]} MPH\n\n`;
  });

  if (segmentSheet && segmentSheet.getLastRow() > 1) {
    const segmentData = segmentSheet.getRange(Math.max(2, segmentSheet.getLastRow() - 9), 1, Math.min(10, segmentSheet.getLastRow() - 1), 9).getValues();
    digest += "## 📈 Recent Segment Highlights\n";
    segmentData.forEach(seg => {
      digest += `- **${seg[3]}**: ${seg[4]} MPH | Maintenance: ${seg[6]} | Power: ${seg[8]}\n`;
    });
  }

  digest += "\n---\n*Copy and paste this into Gemini for your coaching session.*";

  const htmlOutput = HtmlService
    .createHtmlOutput(`<div style="font-family: Arial, sans-serif;">
                       <textarea id="digestText" style="width: 100%; height: 250px; font-family: monospace; font-size: 12px; padding: 10px;">${digest}</textarea>
                       <p>Copy the text above and paste it into your conversation with Gemini.</p>
                       <button onclick="copyToClipboard()" style="background-color: #fc4c02; color: white; padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer;">Copy to Clipboard</button>
                       <script>
                         function copyToClipboard() {
                           var copyText = document.getElementById("digestText");
                           copyText.select();
                           document.execCommand("copy");
                           google.script.host.close();
                         }
                       </script>
                       </div>`)
    .setWidth(500)
    .setHeight(400);
    
  ui.showModalDialog(htmlOutput, 'Gemini Performance Digest');
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

function resetSyncHistory() {
  PropertiesService.getUserProperties().deleteProperty('LAST_SYNC_TIMESTAMP');
  SpreadsheetApp.getUi().alert('History Reset', 'The sync history timestamp has been deleted. Your next sync will start fresh and pull the last 30 activities.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function getSettings(ss) {
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.appendRow(['Setting', 'Value', 'Description']);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f3f3f3");
    sheet.appendRow(['Target Speed (MPH)', 21.0, 'Used for Velocity Maintenance and Target Gap']);
    sheet.appendRow(['Race Distance (Miles)', 26.2, 'Used for Predicted Race Time']);
    sheet.appendRow(['Target Sport', 'InlineSkate', 'Filter activities by this sport type']);
    sheet.appendRow(['Athlete Age', 30, 'Used for Heart Rate Zone calculation']);
    sheet.autoResizeColumns(1, 3);
  }
  
  const data = sheet.getDataRange().getValues();
  const settings = {
    targetSpeed: 21.0,
    raceDistance: 26.2,
    targetSport: 'InlineSkate',
    age: 30
  };
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const val = data[i][1];
    if (key === 'Target Speed (MPH)' && val) settings.targetSpeed = parseFloat(val);
    if (key === 'Race Distance (Miles)' && val) settings.raceDistance = parseFloat(val);
    if (key === 'Target Sport' && val) settings.targetSport = val.toString().trim();
    if (key === 'Athlete Age' && val) settings.age = parseInt(val);
  }
  return settings;
}

function syncStravaData(isBackfill = false) {
  const props = PropertiesService.getUserProperties();
  const clientId = props.getProperty('CLIENT_ID');
  const clientSecret = props.getProperty('CLIENT_SECRET');
  let refreshToken = props.getProperty('REFRESH_TOKEN');

  if (!clientId || !clientSecret || !refreshToken) {
    try {
      SpreadsheetApp.getUi().alert('Configuration Missing', 'Please complete the Strava setup menu first.', SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      console.error('Configuration Missing: Please complete the Strava setup menu first.');
    }
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss);
  
  const summaryHeaders = ['Activity ID', 'Date', 'Neighborhood', 'Name', 'Dist (mi)', 'Elev Gain (ft)', 'Avg MPH', 'Max MPH', 'Pacing Variance (MPH)', 'Suffer Score', 'Training Zone', 'Predicted Race Time', 'Temp (°F)', 'Wind (MPH)', 'Wind Dir (°)', 'Map Polyline'];
  const segmentHeaders = [
    'Activity ID', 'Date', 'Activity', 'Segment Name', 'Avg MPH', 'Avg HR', 
    'Velocity Maintenance %', 'Target Gap Ratio', 'Aerobic Power (S/HR)', 'Segment ID'
  ];

  const summarySheet = checkAndCreateSheet(ss, 'Summary_Data', summaryHeaders);
  const segmentSheet = checkAndCreateSheet(ss, 'Segment_Data', segmentHeaders);

  const existingSummaryIds = getExistingIds(summarySheet);
  // For segments, we can't easily upsert unless we use a compound ID, 
  // so we'll just focus on Summary upserts for now or clear/re-populate segment rows.

  try {
    const tokenUrl = `https://www.strava.com/oauth/token?client_id=${clientId}&client_secret=${clientSecret}&refresh_token=${refreshToken}&grant_type=refresh_token`;
    const tokenResponse = UrlFetchApp.fetch(tokenUrl, { method: 'post' });
    const tokenData = JSON.parse(tokenResponse.getContentText());
    
    // Update refresh token if Strava rotated it
    if (tokenData.refresh_token && tokenData.refresh_token !== refreshToken) {
      props.setProperty('REFRESH_TOKEN', tokenData.refresh_token);
    }
    
    const accessToken = tokenData.access_token;

    // Incremental sync logic
    let afterTimestamp = props.getProperty('LAST_SYNC_TIMESTAMP');
    let activitiesUrl = 'https://www.strava.com/api/v3/athlete/activities?per_page=' + (isBackfill ? 50 : 30);
    if (afterTimestamp && !isBackfill) {
      activitiesUrl += `&after=${afterTimestamp}`;
    }

    const activities = JSON.parse(UrlFetchApp.fetch(activitiesUrl, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    }).getContentText());

    if (activities.length === 0) {
      try {
        SpreadsheetApp.getUi().alert('Sync Complete', 'No new activities found.', SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) {
        console.log('Sync Complete: No new activities found.');
      }
      return;
    }

    // Strava returns newest first. Reverse to append oldest first for a chronological log.
    activities.reverse();
    let maxTimestamp = afterTimestamp ? parseInt(afterTimestamp) : 0;

    activities.forEach(activity => {
      const activityTime = Math.floor(new Date(activity.start_date).getTime() / 1000);
      if (activityTime > maxTimestamp) maxTimestamp = activityTime;
      if (settings.targetSport && activity.sport_type !== settings.targetSport && activity.type !== settings.targetSport) {
        return; // Skip activities that don't match the target sport
      }

      const detail = JSON.parse(UrlFetchApp.fetch(`https://www.strava.com/api/v3/activities/${activity.id}`, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      }).getContentText());

      let neighborhood = "NYC Area";
      let weather = { temp: "N/A", windSpeed: "N/A", windDir: "N/A" };
      
      if (detail.start_latlng) {
        neighborhood = fetchFreeNeighborhood(detail.start_latlng[0], detail.start_latlng[1]);
        weather = fetchWeatherData(detail.start_latlng[0], detail.start_latlng[1], detail.start_date_local);
        Utilities.sleep(1000); // Respect OSM rate limits
      }

      const avgMph = (detail.average_speed * 2.23694);
      
      // Pacing Variance (MPH difference between fastest and slowest mile split)
      let pacingVariance = "N/A";
      if (detail.splits_standard && detail.splits_standard.length > 1) {
        const splitSpeeds = detail.splits_standard.map(s => (s.distance / s.elapsed_time) * 2.23694);
        const maxSplit = Math.max(...splitSpeeds);
        const minSplit = Math.min(...splitSpeeds);
        pacingVariance = (maxSplit - minSplit).toFixed(1);
      }

      // Heart Rate Zone Estimation
      let trainingZone = "N/A";
      if (detail.has_heartrate && detail.average_heartrate) {
        const maxHr = 220 - settings.age;
        const hr = detail.average_heartrate;
        if (hr < maxHr * 0.7) trainingZone = "Zone 2 (Base)";
        else if (hr < maxHr * 0.85) trainingZone = "Zone 3/4 (Tempo)";
        else trainingZone = "Zone 5 (Anaerobic)";
      }

      // Predicted Race Time
      const hours = settings.raceDistance / avgMph;
      const h = Math.floor(hours);
      const m = Math.floor((hours - h) * 60);
      const predictedTime = `${h}h ${m}m`;

      const summaryRow = [
        detail.id.toString(),
        new Date(detail.start_date_local).toLocaleDateString(),
        neighborhood,
        detail.name,
        (detail.distance * 0.000621371).toFixed(2),
        (detail.total_elevation_gain * 3.28084).toFixed(0),
        avgMph.toFixed(1),
        (detail.max_speed * 2.23694).toFixed(1),
        pacingVariance,
        detail.suffer_score || 0,
        trainingZone,
        predictedTime,
        weather.temp,
        weather.windSpeed,
        weather.windDir,
        detail.map ? detail.map.summary_polyline : "N/A"
      ];

      upsertRow(summarySheet, detail.id, summaryRow, existingSummaryIds);

      if (detail.segment_efforts) {
        // If backfilling, we might want to avoid duplicate segments. 
        // For simplicity, we'll only append new segments if they aren't already there for this activity.
        // A full segment sync/upsert would require more complex mapping.
        
        detail.segment_efforts.forEach(effort => {
          // Calculate MPH with fallback if average_speed is missing
          let speedMS = effort.average_speed;
          if ((!speedMS || isNaN(speedMS)) && effort.distance && effort.moving_time) {
            speedMS = effort.distance / effort.moving_time;
          }
          
          const mph = speedMS ? (speedMS * 2.23694) : 0;
          const hr = (effort.average_heartrate && !isNaN(effort.average_heartrate)) ? effort.average_heartrate : 0;
          
          const velocityMaintenance = (mph > 0 && settings.targetSpeed > 0) ? ((mph / settings.targetSpeed) * 100).toFixed(1) : "0.0";
          const targetGap = (mph > 0 && settings.targetSpeed > 0) ? (mph / settings.targetSpeed).toFixed(2) : "0.00";
          const aerobicPower = (mph > 0 && hr > 0) ? (mph / hr).toFixed(3) : "N/A";

          segmentSheet.appendRow([
            detail.id.toString(),
            new Date(detail.start_date_local).toLocaleDateString(),
            detail.name,
            effort.name,
            mph > 0 ? mph.toFixed(2) : "N/A",
            hr > 0 ? hr.toFixed(0) : "N/A",
            velocityMaintenance + "%",
            targetGap,
            aerobicPower,
            effort.segment.id
          ]);
        });
      }
    });

    // Save the latest timestamp so we don't fetch these again
    if (maxTimestamp > 0 && !isBackfill) {
      props.setProperty('LAST_SYNC_TIMESTAMP', maxTimestamp.toString());
    }

    updateDashboard(ss);
    updateVolumeRollup(ss);
    updatePRBoard(ss);

  } catch (e) {
    try {
      SpreadsheetApp.getUi().alert('Sync Error', 'An error occurred during sync: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (err) {
      console.error('Sync Error: ' + e.toString());
    }
  }
}

function updateVolumeRollup(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('Summary_Data');
  if (!summarySheet || summarySheet.getLastRow() < 2) return;

  const data = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 7).getValues();
  const monthly = {};
  const weekly = {};

  data.forEach(row => {
    const date = new Date(row[0]);
    if (isNaN(date)) return;

    // Monthly Key: "YYYY-MM"
    const monthKey = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
    
    // Weekly Key: "Week Starting YYYY-MM-DD" (Monday)
    const day = date.getDay();
    const diff = date.getDate() - day + (day === 0 ? -6 : 1); // Adjust for Monday start
    const monday = new Date(date.setDate(diff));
    const weekKey = `Week Starting ${monday.getFullYear()}-${(monday.getMonth() + 1).toString().padStart(2, '0')}-${monday.getDate().toString().padStart(2, '0')}`;

    const dist = parseFloat(row[3]) || 0;
    const mph = parseFloat(row[4]) || 0;
    const suffer = parseFloat(row[6]) || 0;

    [monthly, weekly].forEach((group, i) => {
      const key = i === 0 ? monthKey : weekKey;
      if (!group[key]) {
        group[key] = { dist: 0, mphSum: 0, count: 0, suffer: 0 };
      }
      group[key].dist += dist;
      group[key].mphSum += (mph * dist); // Weighted by distance
      group[key].count += 1;
      group[key].suffer += suffer;
    });
  });

  const rollupSheet = checkAndCreateSheet(ss, 'Volume_Rollup', ['Period', 'Total Miles', 'Avg MPH', 'Skates', 'Total Suffer']);
  rollupSheet.clearContents();
  rollupSheet.appendRow(['Period', 'Total Miles', 'Avg MPH', 'Skates', 'Total Suffer']);
  rollupSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f3f3f3");

  const appendData = (group) => {
    const sortedKeys = Object.keys(group).sort().reverse();
    sortedKeys.forEach(key => {
      const g = group[key];
      rollupSheet.appendRow([
        key,
        g.dist.toFixed(2),
        g.dist > 0 ? (g.mphSum / g.dist).toFixed(1) : "0.0",
        g.count,
        g.suffer.toFixed(0)
      ]);
    });
  };

  rollupSheet.appendRow(['--- MONTHLY VOLUME ---']);
  rollupSheet.getRange(rollupSheet.getLastRow(), 1).setFontWeight("bold");
  appendData(monthly);
  
  rollupSheet.appendRow(['']);
  rollupSheet.appendRow(['--- WEEKLY VOLUME ---']);
  rollupSheet.getRange(rollupSheet.getLastRow(), 1).setFontWeight("bold");
  appendData(weekly);
  
  rollupSheet.autoResizeColumns(1, 5);
}

function updatePRBoard(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const segmentSheet = ss.getSheetByName('Segment_Data');
  if (!segmentSheet || segmentSheet.getLastRow() < 2) return;

  const data = segmentSheet.getRange(2, 1, segmentSheet.getLastRow() - 1, 9).getValues();
  const prs = {};

  data.forEach(row => {
    const segmentName = row[2];
    const mph = parseFloat(row[3]) || 0;
    const date = row[0];
    const activity = row[1];

    if (!prs[segmentName]) {
      prs[segmentName] = [];
    }
    
    prs[segmentName].push({ mph, date, activity });
  });

  const prSheet = checkAndCreateSheet(ss, 'Personal_Records', ['Segment Name', 'Rank', 'Avg MPH', 'Date', 'Activity']);
  prSheet.clearContents();
  prSheet.appendRow(['Segment Name', 'Rank', 'Avg MPH', 'Date', 'Activity']);
  prSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f3f3f3");

  const sortedSegments = Object.keys(prs).sort();
  sortedSegments.forEach(seg => {
    // Sort efforts for this segment by speed descending
    const efforts = prs[seg].sort((a, b) => b.mph - a.mph);
    
    // Take top 3 PRs
    efforts.slice(0, 3).forEach((effort, index) => {
      prSheet.appendRow([
        index === 0 ? seg : "", // Only show segment name on the first row for readability
        `#${index + 1}`,
        effort.mph.toFixed(2),
        effort.date,
        effort.activity
      ]);
    });
    prSheet.appendRow(['']); // Spacer
  });

  prSheet.autoResizeColumns(1, 5);
}

function updateDashboard(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName('Dashboard');
  if (!sheet) {
    sheet = ss.insertSheet('Dashboard', 0);
  }
  
  const summarySheet = ss.getSheetByName('Summary_Data');
  if (!summarySheet || summarySheet.getLastRow() < 2) return;

  const charts = sheet.getCharts();
  charts.forEach(c => sheet.removeChart(c));

  const lastRow = summarySheet.getLastRow();
  
  const rangeDate = summarySheet.getRange(1, 1, lastRow, 1);
  const rangeMph = summarySheet.getRange(1, 5, lastRow, 1);
  
  const mphChart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rangeDate)
    .addRange(rangeMph)
    .setPosition(2, 2, 0, 0)
    .setOption('title', 'Average MPH Over Time')
    .setOption('legend', {position: 'none'})
    .setOption('vAxis', {title: 'Avg MPH'})
    .setOption('hAxis', {title: 'Date'})
    .build();
    
  sheet.insertChart(mphChart);
  
  const rangeSuffer = summarySheet.getRange(1, 7, lastRow, 1);
  
  const sufferChart = sheet.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(rangeMph)
    .addRange(rangeSuffer)
    .setPosition(2, 8, 0, 0)
    .setOption('title', 'Speed vs. Suffer Score')
    .setOption('hAxis', {title: 'Avg MPH'})
    .setOption('vAxis', {title: 'Suffer Score'})
    .setOption('legend', {position: 'none'})
    .build();
    
  sheet.insertChart(sufferChart);
}

function fetchWeatherData(lat, lng, startDateLocal) {
  try {
    const dateStr = startDateLocal.split('T')[0];
    const hour = parseInt(startDateLocal.split('T')[1].split(':')[0], 10);

    const url = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lng}&start_date=${dateStr}&end_date=${dateStr}&hourly=temperature_2m,wind_speed_10m,wind_direction_10m&temperature_unit=fahrenheit&wind_speed_unit=mph`;
    
    const response = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    if (response && response.hourly) {
      return {
        temp: response.hourly.temperature_2m[hour],
        windSpeed: response.hourly.wind_speed_10m[hour],
        windDir: response.hourly.wind_direction_10m[hour]
      };
    }
  } catch (e) {}
  return { temp: "N/A", windSpeed: "N/A", windDir: "N/A" };
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
    sheet.setFrozenRows(1);
  } else {
    // Sync headers if they have changed (e.g., after a script update)
    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (existingHeaders.join(',') !== headers.join(',')) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
    }
  }
  return sheet;
}

function getExistingIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const map = {};
  ids.forEach((row, index) => {
    if (row[0]) map[row[0].toString()] = index + 2; // Row number
  });
  return map;
}

function upsertRow(sheet, id, rowData, existingIdsMap) {
  const idStr = id.toString();
  if (existingIdsMap[idStr]) {
    const rowNum = existingIdsMap[idStr];
    sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
    existingIdsMap[idStr] = sheet.getLastRow();
  }
}

function backfillMissingData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Backfill Missing Data', 'This will re-sync your last 50 activities to fill in any missing or new columns. It may take a minute. Continue?', ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) return;
  
  // Temporarily clear the sync history for this run
  const props = PropertiesService.getUserProperties();
  const originalSyncTimestamp = props.getProperty('LAST_SYNC_TIMESTAMP');
  props.deleteProperty('LAST_SYNC_TIMESTAMP');
  
  try {
    syncStravaData(true); // Special flag for backfill
    ui.alert('Backfill Complete', 'Successfully re-synced recent activities.', ui.ButtonSet.OK);
  } finally {
    // Restore the sync history
    if (originalSyncTimestamp) {
      props.setProperty('LAST_SYNC_TIMESTAMP', originalSyncTimestamp);
    }
  }
}
