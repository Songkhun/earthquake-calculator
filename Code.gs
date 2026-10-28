/**
 * Earthquake Calculator — Google Apps Script web app.
 *
 * Server side (this file): the Google Sheet (earthquake list, travel-time tables,
 * settings, saved results) and the internet requests to USGS and EarthScope (IRIS).
 * Browser side (Index.html): decodes the seismograms, picks P and S, locates the
 * epicentre and computes the magnitudes.
 */

const SHEETS = { events: 'Events', tt: 'TravelTimes', settings: 'Settings', results: 'Results' };
const FDSN_STATION = 'https://service.earthscope.org/fdsnws/station/1/query';
const FDSN_DATA = 'https://service.earthscope.org/fdsnws/dataselect/1/query';
const USGS_EVENT = 'https://earthquake.usgs.gov/fdsnws/event/1/query';

// Earthquakes from the worksheets (USGS ComCat values).
const DEFAULT_EVENTS = [
  ['usp000hyfx', '2011 Tarlay, Myanmar', '2011-03-24T13:55:12.010Z', 20.687, 99.822, 8, 6.9, 'Mww'],
  ['usb000qack', '2014 Mae Lao, Chiang Rai, Thailand', '2014-05-05T11:08:43.420Z', 19.6557, 99.6696, 6, 6.1, 'Mww'],
  ['us10006gbf', '2016 Chauk, Myanmar', '2016-08-24T10:34:54.580Z', 20.9228, 94.569, 82, 6.8, 'Mww'],
  ['us7000pn9s', '2025 Mandalay, Myanmar', '2025-03-28T06:20:52.715Z', 22.011, 95.9363, 10, 7.7, 'Mww'],
  ['official20041226005853450_30', '2004 Sumatra–Andaman', '2004-12-26T00:58:53.450Z', 3.295, 95.982, 30, 9.1, 'Mw'],
];

const DEFAULT_SETTINGS = [
  ['radius_deg', 20, 'Look for stations up to this distance from the earthquake (degrees, max 20)'],
  ['max_stations', 12, 'Use at most this many stations'],
  ['min_distance_deg', 0.3, 'Skip stations closer than this (degrees)'],
  ['min_separation_deg', 0.5, 'Skip a station this close to one already chosen (avoids duplicates at one site)'],
  ['channels', 'BH?', 'Channel codes to download (BH = broadband, 20–40 samples/s)'],
];

/** Web app entry point. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Earthquake Calculator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Earthquake app')
    .addItem('Set up sheets', 'setup')
    .addItem('Show web app link', 'showLink')
    .addToUi();
}

/** Creates the sheets. Run from the menu or the editor; the web app also runs it if the tabs are missing. */
function setup() {
  setupSheets_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getActive().toast('Sheets are ready. Next: Deploy → New deployment → Web app.');
}

function setupSheets_(ss) {
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  const ev = sheet_(ss, SHEETS.events);
  if (ev.getLastRow() < 2) {
    ev.clear();
    ev.getRange(1, 1, 1, 8).setValues([['USGS ID', 'Name', 'Origin time (UTC)', 'Latitude', 'Longitude', 'Depth (km)', 'Magnitude', 'Mag. type']]);
    ev.getRange('C:C').setNumberFormat('@');  // keep times as text
    ev.getRange(2, 1, DEFAULT_EVENTS.length, 8).setValues(DEFAULT_EVENTS);
  }

  const tt = sheet_(ss, SHEETS.tt);
  tt.clear();
  const head = ['Δ (deg)'];
  TT_DEPTHS.forEach(h => head.push(`tP h=${h} km (s)`, `S−P h=${h} km (s)`));
  tt.getRange(1, 1, 1, head.length).setValues([head]);
  tt.getRange(2, 1, TT_ROWS.length, head.length).setValues(TT_ROWS);
  tt.getRange(1, head.length + 2).setValue('iasp91 first-arrival P and S times (TauP). The app interpolates between depths.');

  const st = sheet_(ss, SHEETS.settings);
  if (st.getLastRow() < 2) {
    st.clear();
    st.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Meaning']]);
    st.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  }

  const res = sheet_(ss, SHEETS.results);
  if (res.getLastRow() < 1) {
    res.getRange(1, 1, 1, 14).setValues([['Saved at', 'Student', 'Email', 'Event ID', 'Event', 'Epicentre lat', 'Epicentre lon',
      'Distance from USGS (km)', 'Stations used', 'mb', 'Ms', 'Catalogue M', 'Picks (JSON)', 'Station magnitudes (JSON)']]);
  }

  [ev, tt, st, res].forEach(s => s.setFrozenRows(1));
}

function showLink() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? 'Web app: ' + url : 'Not deployed yet. In the Apps Script editor: Deploy → New deployment → Web app.');
}

// ---------------------------------------------------------------- called from Index.html

/** Everything the page needs at start: earthquakes, travel-time table, settings. */
function getInitData() {
  const ss = ss_();
  if (!ss) throw new Error('No Google Sheet is linked. Create the script from the Sheet (Extensions → Apps Script), not at script.google.com.');
  if (!ss.getSheetByName(SHEETS.events) || !ss.getSheetByName(SHEETS.tt)) setupSheets_(ss);  // first run
  const ev = ss.getSheetByName(SHEETS.events);
  const events = ev.getLastRow() < 2 ? [] : ev.getRange(2, 1, ev.getLastRow() - 1, 8).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({ id: String(r[0]), name: r[1], time: r[2] instanceof Date ? r[2].toISOString() : String(r[2]),
                 lat: +r[3], lon: +r[4], depth: +r[5], mag: +r[6], magType: r[7] }));
  const tt = ss.getSheetByName(SHEETS.tt);
  const rows = tt.getRange(2, 1, tt.getLastRow() - 1, 1 + 2 * TT_DEPTHS.length).getValues();
  return { events, tt: { depths: TT_DEPTHS, rows }, settings: settings_() };
}

/** Searches the USGS catalogue. opts: {start, end, minMag, lat, lon, radiusDeg}. */
function searchUsgs(opts) {
  const url = USGS_EVENT + '?' + query_({
    format: 'geojson', starttime: opts.start, endtime: opts.end, minmagnitude: opts.minMag,
    latitude: opts.lat, longitude: opts.lon, maxradius: opts.radiusDeg, orderby: 'magnitude', limit: 50,
  });
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('USGS search failed (HTTP ' + res.getResponseCode() + ')');
  return JSON.parse(res.getContentText()).features.map(f => ({
    id: f.id, name: f.properties.place, time: new Date(f.properties.time).toISOString(),
    lat: f.geometry.coordinates[1], lon: f.geometry.coordinates[0], depth: f.geometry.coordinates[2],
    mag: f.properties.mag, magType: f.properties.magType,
  }));
}

/** Adds an earthquake from a USGS search to the Events sheet (once). */
function addEvent(e) {
  const sh = ss_().getSheetByName(SHEETS.events);
  const ids = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  if (!ids.includes(e.id)) sh.appendRow([e.id, e.name, e.time, e.lat, e.lon, e.depth, e.mag, e.magType]);
  return getInitData().events;
}

/**
 * Finds broadband stations around the earthquake and downloads their records
 * (miniSEED, base64-encoded). Returns the stations that actually returned data.
 */
function getSeismograms(ev) {
  const s = settings_();
  const t0 = new Date(ev.time);
  const stations = findStations_(ev, s);
  if (!stations.length) return { stations: [], skipped: [] };

  // Window: 60 s before the origin until the Rayleigh wave (≈3 km/s) has passed, plus 60 s.
  const requests = stations.map(st => {
    const end = new Date(t0.getTime() + Math.min(1500, 60 + st.distDeg * 111.2 / 3.0 + 60) * 1000);
    return {
      url: FDSN_DATA + '?' + query_({ net: st.net, sta: st.sta, loc: st.loc || '--', cha: st.channels.map(c => c.cha).join(','),
                                      starttime: iso_(new Date(t0.getTime() - 60000)), endtime: iso_(end) }),
      muteHttpExceptions: true,
    };
  });
  const replies = UrlFetchApp.fetchAll(requests);
  const ok = [], skipped = [];
  replies.forEach((r, i) => {
    const st = stations[i];
    if (r.getResponseCode() === 200 && r.getContent().length > 0 && ok.length < s.max_stations) {
      st.mseed = Utilities.base64Encode(r.getContent());
      ok.push(st);
    } else if (r.getResponseCode() !== 200) {
      skipped.push(`${st.net}.${st.sta}: ${r.getResponseCode() === 204 ? 'no data' : 'HTTP ' + r.getResponseCode()}`);
    }
  });
  return { stations: ok, skipped };
}

/** Saves one student's result as a new row of the Results sheet. */
function saveResult(r) {
  const sh = ss_().getSheetByName(SHEETS.results);
  let email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) { /* not available for public web apps */ }
  sh.appendRow([new Date(), r.student, email, r.eventId, r.eventName, r.lat, r.lon, r.errKm, r.nStations,
                r.mb, r.Ms, r.catalogMag, JSON.stringify(r.picks), JSON.stringify(r.stationMags)]);
  return sh.getLastRow();
}

// ---------------------------------------------------------------- helpers

function findStations_(ev, s) {
  const t0 = new Date(ev.time);
  const url = FDSN_STATION + '?' + query_({
    latitude: ev.lat, longitude: ev.lon, maxradius: Math.min(20, s.radius_deg), channel: s.channels,
    level: 'channel', format: 'text', starttime: iso_(t0), endtime: iso_(new Date(t0.getTime() + 600000)),
  });
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() === 204) return [];
  if (res.getResponseCode() !== 200) throw new Error('Station search failed (HTTP ' + res.getResponseCode() + ')');

  // Text columns: Net|Sta|Loc|Cha|Lat|Lon|Elev|Depth|Azimuth|Dip|Sensor|Scale|ScaleFreq|ScaleUnits|SampleRate|Start|End
  const groups = {};
  res.getContentText().split('\n').filter(l => l && l[0] !== '#').forEach(l => {
    const c = l.split('|');
    const key = `${c[0]}.${c[1]}.${c[2]}`;
    const g = groups[key] = groups[key] || { net: c[0], sta: c[1], loc: c[2], lat: +c[4], lon: +c[5], channels: [] };
    if (!g.channels.some(x => x.cha === c[3]) && c[13].toUpperCase() === 'M/S')
      g.channels.push({ cha: c[3], az: +c[8], dip: +c[9], scale: +c[11], rate: +c[14] });
  });

  // Keep one location code per station: complete (Z + 2 horizontals), lowest sample rate, prefer "00".
  const best = {};
  Object.values(groups).forEach(g => {
    const z = g.channels.filter(c => Math.abs(c.dip) > 60), h = g.channels.filter(c => Math.abs(c.dip) <= 30);
    if (z.length < 1 || h.length < 2) return;
    g.channels = [z[0], h[0], h[1]];
    g.rate = z[0].rate;
    g.distDeg = distDeg_(ev.lat, ev.lon, g.lat, g.lon);
    const id = `${g.net}.${g.sta}`, cur = best[id];
    const score = x => x.rate * 10 + (x.loc === '00' ? 0 : x.loc === '' ? 1 : 2);
    if (!cur || score(g) < score(cur)) best[id] = g;
  });

  // Nearest first, skipping duplicates at the same site. Ask for four times as many as needed,
  // because some networks (e.g. recent Thai TMD data) are not shared through EarthScope.
  const chosen = [];
  Object.values(best).filter(g => g.distDeg >= s.min_distance_deg).sort((a, b) => a.distDeg - b.distDeg)
    .forEach(g => {
      if (chosen.length >= 4 * s.max_stations) return;
      if (chosen.some(c => distDeg_(c.lat, c.lon, g.lat, g.lon) < s.min_separation_deg)) return;
      chosen.push(g);
    });
  return chosen;
}

function settings_() {
  const out = {};
  DEFAULT_SETTINGS.forEach(r => out[r[0]] = r[1]);
  const sh = ss_().getSheetByName(SHEETS.settings);
  if (sh && sh.getLastRow() > 1)
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] !== '') out[r[0]] = r[1]; });
  return out;
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function distDeg_(la1, lo1, la2, lo2) {
  const r = Math.PI / 180;
  const a = Math.sin((la2 - la1) * r / 2) ** 2 + Math.cos(la1 * r) * Math.cos(la2 * r) * Math.sin((lo2 - lo1) * r / 2) ** 2;
  return 2 * Math.asin(Math.min(1, Math.sqrt(a))) / r;
}

function iso_(d) {
  return d.toISOString().replace('Z', '');
}

function query_(params) {
  return Object.keys(params).filter(k => params[k] !== undefined && params[k] !== '')
    .map(k => k + '=' + encodeURIComponent(params[k])).join('&');
}

// ---------------------------------------------------------------- travel-time table

/**
 * iasp91 travel times (first-arriving P and S), computed with TauP (ObsPy) for the
 * Earthquake Calculator. Each row: [Δ (deg), tP h=10, S−P h=10, tP h=35, S−P h=35,
 * tP h=80, S−P h=80, tP h=150, S−P h=150] in seconds. setup() copies it to the TravelTimes sheet.
 */
const TT_DEPTHS = [10, 35, 80, 150];
const TT_ROWS = [
  [0.1, 2.6, 1.9, 6, 4.4, 11.5, 8.7, 20.1, 15.6],
  [0.2, 4.2, 3.1, 6.8, 5, 11.8, 9, 20.2, 15.7],
  [0.3, 6, 4.4, 7.9, 5.8, 12.3, 9.4, 20.5, 15.9],
  [0.4, 7.9, 5.7, 9.3, 6.8, 12.9, 9.9, 20.9, 16.2],
  [0.5, 9.7, 7.1, 10.6, 7.9, 13.7, 10.5, 21.3, 16.5],
  [0.6, 11.6, 8.4, 12, 9, 14.6, 11.2, 21.8, 17],
  [0.7, 13.5, 9.8, 13.4, 10.1, 15.6, 12, 22.5, 17.4],
  [0.8, 15.4, 11.2, 14.8, 11.2, 16.7, 12.8, 23.2, 18],
  [0.9, 17.3, 12.6, 16.1, 12.3, 17.9, 13.7, 23.9, 18.6],
  [1, 19.2, 14, 17.5, 13.4, 19, 14.6, 24.7, 19.2],
  [1.1, 21.1, 15.4, 18.9, 14.5, 20.2, 15.6, 25.6, 19.9],
  [1.2, 22.8, 16.6, 20.3, 15.6, 21.5, 16.6, 26.5, 20.6],
  [1.3, 24.2, 18.2, 21.6, 16.6, 22.7, 17.5, 27.5, 21.4],
  [1.4, 25.6, 19.3, 23, 17.7, 24, 18.6, 28.5, 22.2],
  [1.5, 27, 20.4, 24.4, 18.8, 25.3, 19.6, 29.5, 23],
  [1.6, 28.3, 21.5, 25.8, 19.9, 26.6, 20.6, 30.6, 23.8],
  [1.7, 29.7, 22.6, 27.1, 21, 27.9, 21.6, 31.7, 24.7],
  [1.8, 31.1, 23.7, 28.5, 22.1, 29.2, 22.7, 32.8, 25.6],
  [1.9, 32.5, 24.8, 29.9, 23.2, 30.5, 23.7, 33.9, 26.5],
  [2, 33.8, 25.9, 31.3, 24.3, 31.9, 24.8, 35.1, 27.4],
  [2.1, 35.2, 27, 32.6, 25.4, 33.2, 25.8, 36.3, 28.3],
  [2.2, 36.6, 28.1, 34, 26.5, 34.5, 26.9, 37.4, 29.2],
  [2.3, 38, 29.2, 35.4, 27.6, 35.9, 28, 38.6, 30.2],
  [2.4, 39.3, 30.3, 36.8, 28.7, 37.2, 29, 39.9, 31.2],
  [2.5, 40.7, 31.4, 38.1, 29.8, 38.6, 30.1, 41.1, 32.1],
  [2.6, 42.1, 32.5, 39.5, 30.9, 39.9, 31.2, 42.3, 33.1],
  [2.7, 43.5, 33.6, 40.9, 32, 41.3, 32.2, 43.5, 34.1],
  [2.8, 44.8, 34.7, 42.3, 33.1, 42.6, 33.3, 44.8, 35.1],
  [2.9, 46.2, 35.8, 43.6, 34.2, 44, 34.4, 46.1, 36.1],
  [3, 47.6, 36.9, 45, 35.3, 45.3, 35.5, 47.3, 37.1],
  [3.1, 49, 38, 46.4, 36.4, 46.7, 36.5, 48.6, 38.1],
  [3.2, 50.3, 39.1, 47.8, 37.5, 48, 37.6, 49.9, 39.1],
  [3.3, 51.7, 40.2, 49.1, 38.6, 49.4, 38.7, 51.1, 40.1],
  [3.4, 53.1, 41.3, 50.5, 39.7, 50.7, 39.8, 52.4, 41.1],
  [3.5, 54.5, 42.4, 51.9, 40.8, 52.1, 40.9, 53.7, 42.1],
  [3.6, 55.8, 43.5, 53.3, 41.9, 53.4, 41.9, 55, 43.2],
  [3.7, 57.2, 44.6, 54.6, 43, 54.8, 43, 56.3, 44.2],
  [3.8, 58.6, 45.7, 56, 44.1, 56.2, 44.1, 57.6, 45.2],
  [3.9, 60, 46.8, 57.4, 45.2, 57.5, 45.2, 58.9, 46.3],
  [4, 61.3, 47.9, 58.8, 46.3, 58.9, 46.3, 60.2, 47.3],
  [4.1, 62.7, 49, 60.1, 47.4, 60.2, 47.3, 61.5, 48.3],
  [4.2, 64.1, 50.1, 61.5, 48.5, 61.6, 48.4, 62.8, 49.4],
  [4.3, 65.5, 51.2, 62.9, 49.6, 63, 49.5, 64.1, 50.4],
  [4.4, 66.8, 52.3, 64.3, 50.7, 64.3, 50.6, 65.4, 51.5],
  [4.5, 68.2, 53.4, 65.6, 51.8, 65.7, 51.7, 66.8, 52.5],
  [4.6, 69.6, 54.5, 67, 52.9, 67.1, 52.8, 68.1, 53.5],
  [4.7, 70.9, 55.6, 68.4, 54, 68.4, 53.8, 69.4, 54.6],
  [4.8, 72.3, 56.7, 69.8, 55.1, 69.8, 54.9, 70.7, 55.6],
  [4.9, 73.7, 57.8, 71.1, 56.2, 71.1, 56, 72, 56.7],
  [5, 75.1, 58.9, 72.5, 57.2, 72.5, 57.1, 73.4, 57.7],
  [5.1, 76.4, 59.9, 73.9, 58.3, 73.9, 58.2, 74.7, 58.8],
  [5.2, 77.8, 61, 75.3, 59.4, 75.2, 59.3, 76, 59.8],
  [5.3, 79.2, 62.1, 76.6, 60.5, 76.6, 60.3, 77.3, 60.9],
  [5.4, 80.6, 63.2, 78, 61.6, 78, 61.4, 78.7, 62],
  [5.5, 81.9, 64.3, 79.4, 62.7, 79.3, 62.5, 80, 63],
  [5.6, 83.3, 65.4, 80.8, 63.8, 80.7, 63.6, 81.3, 64.1],
  [5.7, 84.7, 66.5, 82.1, 64.9, 82.1, 64.7, 82.7, 65.1],
  [5.8, 86.1, 67.6, 83.5, 66, 83.4, 65.8, 84, 66.2],
  [5.9, 87.4, 68.7, 84.9, 67.1, 84.8, 66.8, 85.3, 67.3],
  [6, 88.8, 69.8, 86.2, 68.2, 86.1, 67.9, 86.6, 68.3],
  [6.1, 90.2, 70.9, 87.6, 69.3, 87.5, 69, 88, 69.4],
  [6.2, 91.6, 72, 89, 70.4, 88.9, 70.1, 89.3, 70.4],
  [6.3, 92.9, 73.1, 90.4, 71.5, 90.2, 71.2, 90.6, 71.5],
  [6.4, 94.3, 74.2, 91.7, 72.6, 91.6, 72.3, 92, 72.6],
  [6.5, 95.7, 75.3, 93.1, 73.7, 93, 73.3, 93.3, 73.6],
  [6.6, 97.1, 76.4, 94.5, 74.8, 94.3, 74.4, 94.6, 74.7],
  [6.7, 98.4, 77.5, 95.9, 75.8, 95.7, 75.5, 96, 75.8],
  [6.8, 99.8, 78.6, 97.2, 76.9, 97.1, 76.6, 97.3, 76.8],
  [6.9, 101.2, 79.6, 98.6, 78, 98.4, 77.7, 98.7, 77.9],
  [7, 102.5, 80.7, 100, 79.1, 99.8, 78.7, 100, 79],
  [7.1, 103.9, 81.8, 101.4, 80.2, 101.2, 79.8, 101.3, 80],
  [7.2, 105.3, 82.9, 102.7, 81.3, 102.5, 80.9, 102.7, 81.1],
  [7.3, 106.7, 84, 104.1, 82.4, 103.9, 82, 104, 82.2],
  [7.4, 108, 85.1, 105.5, 83.5, 105.3, 83.1, 105.3, 83.2],
  [7.5, 109.4, 86.2, 106.8, 84.6, 106.6, 84.2, 106.7, 84.3],
  [7.6, 110.8, 87.3, 108.2, 85.7, 108, 85.2, 108, 85.4],
  [7.7, 112.2, 88.4, 109.6, 86.8, 109.3, 86.3, 109.3, 86.4],
  [7.8, 113.5, 89.5, 111, 87.9, 110.7, 87.4, 110.7, 87.5],
  [7.9, 114.9, 90.6, 112.3, 88.9, 112.1, 88.5, 112, 88.6],
  [8, 116.3, 91.7, 113.7, 90, 113.4, 89.6, 113.3, 89.7],
  [8.1, 117.6, 92.7, 115.1, 91.1, 114.8, 90.6, 114.7, 90.7],
  [8.2, 119, 93.8, 116.4, 92.2, 116.2, 91.7, 116, 91.8],
  [8.3, 120.4, 94.9, 117.8, 93.3, 117.5, 92.8, 117.3, 92.9],
  [8.4, 121.8, 96, 119.2, 94.4, 118.9, 93.9, 118.7, 94],
  [8.5, 123.1, 97.1, 120.6, 95.5, 120.3, 95, 120, 95],
  [8.6, 124.5, 98.2, 121.9, 96.6, 121.6, 96, 121.3, 96.1],
  [8.7, 125.9, 99.3, 123.3, 97.7, 123, 97.1, 122.7, 97.2],
  [8.8, 127.2, 100.4, 124.7, 98.7, 124.4, 98.2, 124, 98.3],
  [8.9, 128.6, 101.5, 126, 99.8, 125.7, 99.3, 125.3, 99.3],
  [9, 130, 102.5, 127.4, 100.9, 127.1, 100.4, 126.6, 100.4],
  [9.1, 131.4, 103.6, 128.8, 102, 128.4, 101.4, 128, 101.5],
  [9.2, 132.7, 104.7, 130.2, 103.1, 129.8, 102.5, 129.3, 102.6],
  [9.3, 134.1, 105.8, 131.5, 104.2, 131.2, 103.6, 130.6, 103.7],
  [9.4, 135.5, 106.9, 132.9, 105.3, 132.5, 104.7, 132, 104.7],
  [9.5, 136.8, 108, 134.3, 106.4, 133.9, 105.7, 133.3, 105.8],
  [9.6, 138.2, 109.1, 135.6, 107.4, 135.3, 106.8, 134.6, 106.9],
  [9.7, 139.6, 110.2, 137, 108.5, 136.6, 107.9, 135.9, 108],
  [9.8, 141, 111.2, 138.4, 109.6, 138, 109, 137.3, 109.1],
  [9.9, 142.3, 112.3, 139.7, 110.7, 139.4, 110.1, 138.6, 110.2],
  [10, 143.7, 113.4, 141.1, 111.8, 140.7, 111.1, 139.9, 111.3],
  [10.1, 145.1, 114.5, 142.5, 112.9, 142.1, 112.2, 141.2, 112.3],
  [10.2, 146.4, 115.6, 143.9, 114, 143.4, 113.3, 142.5, 113.4],
  [10.3, 147.8, 116.7, 145.2, 115, 144.8, 114.4, 143.9, 114.5],
  [10.4, 149.2, 117.8, 146.6, 116.1, 146.2, 115.4, 145.2, 115.6],
  [10.5, 150.5, 118.8, 148, 117.2, 147.5, 116.5, 146.5, 116.7],
  [10.6, 151.9, 119.9, 149.3, 118.3, 148.9, 117.6, 147.8, 117.8],
  [10.7, 153.3, 121, 150.7, 119.4, 150.3, 118.7, 149.1, 118.9],
  [10.8, 154.6, 122.1, 152.1, 120.5, 151.6, 119.7, 150.5, 120],
  [10.9, 156, 123.2, 153.4, 121.5, 153, 120.8, 151.8, 121.1],
  [11, 157.4, 124.3, 154.8, 122.6, 154.3, 121.9, 153.1, 122.2],
  [11.1, 158.8, 125.3, 156.2, 123.7, 155.7, 123, 154.4, 123.3],
  [11.2, 160.1, 126.4, 157.5, 124.8, 157.1, 124, 155.7, 124.4],
  [11.3, 161.5, 127.5, 158.9, 125.9, 158.4, 125.1, 157, 125.5],
  [11.4, 162.9, 128.6, 160.3, 126.9, 159.8, 126.2, 158.3, 126.6],
  [11.5, 164.2, 129.7, 161.6, 128, 161.2, 127.2, 159.6, 127.7],
  [11.6, 165.6, 130.8, 163, 129.1, 162.5, 128.3, 160.9, 128.8],
  [11.7, 167, 131.8, 164.4, 130.2, 163.9, 129.4, 162.3, 129.9],
  [11.8, 168.3, 132.9, 165.7, 131.3, 165.2, 130.5, 163.6, 131],
  [11.9, 169.7, 134, 167.1, 132.3, 166.6, 131.5, 164.9, 132.1],
  [12, 171.1, 135.1, 168.5, 133.4, 168, 132.6, 166.2, 133.2],
  [12.1, 172.4, 136.2, 169.8, 134.5, 169.3, 133.7, 167.5, 134.3],
  [12.2, 173.8, 137.2, 171.2, 135.6, 170.7, 134.7, 168.8, 135.4],
  [12.3, 175.2, 138.3, 172.6, 136.7, 172, 135.8, 170.1, 136.5],
  [12.4, 176.5, 139.4, 173.9, 137.7, 173.4, 136.9, 171.4, 137.6],
  [12.5, 177.9, 140.5, 175.3, 138.8, 174.8, 137.9, 172.7, 138.7],
  [12.6, 179.3, 141.5, 176.7, 139.9, 176.1, 139, 174, 139.8],
  [12.7, 180.6, 142.6, 178, 141, 177.4, 140.1, 175.3, 140.9],
  [12.8, 182, 143.7, 179.4, 142, 178.8, 141.2, 176.5, 142.1],
  [12.9, 183.4, 144.8, 180.8, 143.1, 180.1, 142.3, 177.8, 143.2],
  [13, 184.7, 145.9, 182.1, 144.2, 181.4, 143.4, 179.1, 144.3],
  [13.1, 186.1, 146.9, 183.5, 145.3, 182.7, 144.5, 180.4, 145.4],
  [13.2, 187.5, 148, 184.9, 146.3, 184.1, 145.6, 181.7, 146.6],
  [13.3, 188.8, 149.1, 186.2, 147.4, 185.4, 146.8, 182.9, 147.7],
  [13.4, 190.2, 150.2, 187.6, 148.5, 186.7, 147.9, 184.2, 148.8],
  [13.5, 191.6, 151.2, 189, 149.6, 188, 149, 185.5, 149.9],
  [13.6, 192.9, 152.3, 190.3, 150.6, 189.3, 150.1, 186.7, 151.1],
  [13.7, 194.3, 153.4, 191.7, 151.7, 190.6, 151.2, 188, 152.2],
  [13.8, 195.7, 154.5, 193.1, 152.8, 191.9, 152.3, 189.3, 153.4],
  [13.9, 197, 155.5, 194.4, 153.8, 193.3, 153.4, 190.5, 154.5],
  [14, 198.4, 156.6, 195.8, 154.9, 194.6, 154.5, 191.8, 155.6],
  [14.1, 199.7, 157.7, 197.2, 156, 195.9, 155.6, 193.1, 156.8],
  [14.2, 201.1, 158.7, 198.5, 157.1, 197.2, 156.8, 194.3, 157.9],
  [14.3, 202.5, 159.8, 199.9, 158.1, 198.5, 157.9, 195.6, 159.1],
  [14.4, 203.8, 160.9, 201.2, 159.2, 199.8, 159, 196.8, 160.2],
  [14.5, 205.2, 162, 202.6, 160.3, 201.1, 160.1, 198.1, 161.4],
  [14.6, 206.6, 163, 204, 161.3, 202.4, 161.2, 199.3, 162.5],
  [14.7, 207.9, 164.1, 205.3, 162.4, 203.7, 162.4, 200.6, 163.7],
  [14.8, 209.3, 165.2, 206.7, 163.5, 205, 163.5, 201.8, 164.8],
  [14.9, 210.7, 166.2, 208, 164.6, 206.2, 164.6, 203.1, 166],
  [15, 212, 167.3, 209.3, 165.7, 207.5, 165.8, 204.3, 167.1],
  [15.1, 213.3, 168.4, 210.6, 166.9, 208.8, 166.9, 205.6, 168.3],
  [15.2, 214.6, 169.5, 211.9, 168, 210.1, 168.1, 206.8, 169.5],
  [15.3, 216, 170.6, 213.2, 169.1, 211.3, 169.2, 208, 170.6],
  [15.4, 217.3, 171.8, 214.5, 170.2, 212.6, 170.4, 209.3, 171.8],
  [15.5, 218.6, 172.9, 215.8, 171.3, 213.9, 171.5, 210.5, 172.9],
  [15.6, 219.9, 174, 217.1, 172.5, 215.1, 172.7, 211.7, 174.1],
  [15.7, 221.2, 175.1, 218.4, 173.6, 216.4, 173.8, 213, 175.3],
  [15.8, 222.5, 176.2, 219.7, 174.7, 217.7, 175, 214.2, 176.4],
  [15.9, 223.8, 177.4, 221, 175.8, 218.9, 176.2, 215.4, 177.6],
  [16, 225.1, 178.5, 222.3, 177, 220.2, 177.3, 216.6, 178.7],
  [16.1, 226.4, 179.6, 223.6, 178.1, 221.4, 178.5, 217.9, 179.7],
  [16.2, 227.7, 180.8, 224.9, 179.3, 222.7, 179.7, 219.1, 180.8],
  [16.3, 229, 181.9, 226.1, 180.4, 223.9, 180.8, 220.3, 181.9],
  [16.4, 230.2, 183.1, 227.4, 181.6, 225.2, 182, 221.5, 183],
  [16.5, 231.5, 184.2, 228.7, 182.7, 226.4, 183.2, 222.7, 184.1],
  [16.6, 232.8, 185.4, 229.9, 183.9, 227.6, 184.3, 223.8, 185.3],
  [16.7, 234, 186.5, 231.2, 185.1, 228.9, 185.5, 224.9, 186.4],
  [16.8, 235.3, 187.7, 232.4, 186.2, 230.1, 186.7, 226, 187.6],
  [16.9, 236.6, 188.9, 233.7, 187.4, 231.3, 187.9, 227.1, 188.8],
  [17, 237.8, 190, 235, 188.6, 232.6, 189.1, 228.2, 189.9],
  [17.1, 239.1, 191.2, 236.2, 189.7, 233.8, 190.2, 229.3, 191],
  [17.2, 240.3, 192.4, 237.4, 190.9, 235, 191.4, 230.4, 192],
  [17.3, 241.6, 193.5, 238.7, 192.1, 236.3, 192.6, 231.5, 192.9],
  [17.4, 242.8, 194.7, 239.9, 193.3, 237.5, 193.8, 232.6, 193.8],
  [17.5, 244.1, 195.9, 241.2, 194.4, 238.7, 195, 233.7, 194.7],
  [17.6, 245.3, 197, 242.4, 195.6, 239.8, 196.2, 234.8, 195.7],
  [17.7, 246.5, 198.2, 243.6, 196.8, 240.9, 197.6, 235.9, 196.6],
  [17.8, 247.8, 199.4, 244.9, 198, 242, 198.9, 237, 197.5],
  [17.9, 249, 200.6, 246.1, 199.2, 243.1, 200.2, 238.1, 198.4],
  [18, 250.3, 201.8, 247.3, 200.4, 244.2, 201.5, 239.1, 199.3],
  [18.1, 251.5, 203, 248.6, 201.6, 245.3, 202.8, 240.2, 200.2],
  [18.2, 252.7, 204.2, 249.8, 202.8, 246.4, 204.1, 241.3, 201.2],
  [18.3, 253.9, 205.3, 250.9, 204.1, 247.5, 205.4, 242.4, 202.1],
  [18.4, 255.1, 206.6, 252, 205.4, 248.6, 206.6, 243.5, 203],
  [18.5, 256.2, 207.9, 253.1, 206.7, 249.7, 207.5, 244.6, 203.9],
  [18.6, 257.3, 209.2, 254.2, 208, 250.8, 208.4, 245.7, 204.8],
  [18.7, 258.4, 210.5, 255.3, 209.3, 251.9, 209.4, 246.8, 205.7],
  [18.8, 259.5, 211.8, 256.4, 210.7, 253, 210.3, 247.9, 206.6],
  [18.9, 260.6, 213.1, 257.5, 212, 254.1, 211.2, 248.9, 207.5],
  [19, 261.7, 214.4, 258.6, 213.3, 255.2, 212.1, 250, 208.4],
  [19.1, 262.8, 215.8, 259.7, 214.6, 256.3, 213, 251.1, 209.3],
  [19.2, 263.9, 217.1, 260.8, 215.9, 257.4, 213.9, 252.2, 210.2],
  [19.3, 265, 218.4, 261.8, 217.2, 258.5, 214.9, 253.3, 211.1],
  [19.4, 266.1, 219.7, 262.9, 218.2, 259.6, 215.8, 254.4, 212],
  [19.5, 267.2, 221, 264, 219.1, 260.6, 216.7, 255.4, 212.9],
  [19.6, 268.3, 222.1, 265.1, 220, 261.7, 217.6, 256.5, 213.8],
  [19.7, 269.4, 223.1, 266.2, 221, 262.8, 218.5, 257.6, 214.7],
  [19.8, 270.5, 224, 267.3, 221.9, 263.9, 219.4, 258.7, 215.6],
  [19.9, 271.6, 224.9, 268.4, 222.8, 265, 220.3, 259.8, 216.5],
  [20, 272.7, 225.8, 269.5, 223.7, 266.1, 221.2, 260.8, 217.4]
];
