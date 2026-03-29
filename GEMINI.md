# Strava Performance Engine v4.0

## Project Overview
The **Strava Performance Engine** is a Google Apps Script (GAS) project designed to synchronize, analyze, and visualize inline speed skating performance data from Strava within a Google Spreadsheet. It provides advanced metrics such as velocity maintenance, heart rate performance, and race pacing (specifically targeted at marathon distances).

### Main Technologies
- **Language:** Google Apps Script (JavaScript-based)
- **APIs:** 
  - **Strava API:** Primary source for athlete activities and segment efforts.
  - **Open-Meteo API:** Fetches historical weather data (temperature, wind speed/direction) for activities.
  - **Nominatim (OpenStreetMap):** Reverse geocoding to identify activity "neighborhoods."
- **Storage:** Google Sheets for data; `PropertiesService` (UserProperties) for secure API credential storage.

### Key Features
- **Secure Authentication:** OAuth2 flow with Strava, storing Client IDs, Secrets, and Refresh Tokens securely.
- **Incremental Sync:** Pulls only new activities since the last sync to optimize performance and API limits.
- **Performance Analytics:** Calculates metrics like "Velocity Maintenance %" (comparison against a target speed) and "Aerobic Power" (speed/HR ratio).
- **Automated Dashboards:** Generates charts (Avg MPH over time, Speed vs. Suffer Score) and rollup sheets (Monthly/Weekly volume, Personal Records).
- **Gemini Integration:** A "Gemini Digest" feature generates a Markdown summary of recent performance for easy copy-pasting into LLMs for deep-dive analysis.

---

## Getting Started

### Installation
1. Create a new Google Sheet.
2. Open **Extensions > Apps Script**.
3. Replace the contents of `Code.gs` in the editor with the `Code.gs` file from this repository.
4. Save and refresh the spreadsheet to see the **🏃‍♂️ Strava** menu.

### Configuration
The project uses a custom menu for setup:
1. **🏃‍♂️ Strava > 1. Setup API Credentials:** Enter Strava Client ID and Secret.
2. **🏃‍♂️ Strava > 2. Authorize Strava:** Follow the OAuth link and copy the `code` from the redirect URL (localhost).
3. **🏃‍♂️ Strava > 3. Complete Authorization:** Paste the code to obtain a Refresh Token.

---

## Development & Architecture

### File Structure
- `Code.gs`: Contains the entire logic, including UI creation, API interaction, data processing, and sheet management.
- `README.md`: User-facing documentation for setup and usage.

### Key Functions
- `onOpen()`: Initializes the custom spreadsheet menu.
- `syncStravaData()`: The core execution engine that fetches data, processes it, and updates the sheets.
- `showGeminiDigest()`: Generates the performance summary for LLM analysis.
- `updateVolumeRollup()`, `updatePRBoard()`, `updateDashboard()`: Handle post-sync data aggregation and visualization.

### Development Conventions
- **Secrets Management:** NEVER hardcode API keys. Use the `setupCredentials()` flow which utilizes `PropertiesService.getUserProperties()`.
- **Rate Limiting:** Includes `Utilities.sleep(1000)` when calling Nominatim to respect OpenStreetMap's usage policy.
- **Data Integrity:** Uses `checkAndCreateSheet()` to ensure necessary tabs exist and headers are consistent.

---

## Usage for Gemini CLI
When interacting with this project via Gemini:
1. **Code Modification:** Edits should be made directly to `Code.gs`.
2. **Feature Requests:** Can include adding new metrics, integrating additional weather parameters, or refining the "Gemini Digest" output.
3. **Data Analysis:** If you have access to the "Digest" output from a user, you can provide detailed training advice based on the metrics defined in `Code.gs`.
