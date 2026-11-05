/**
 * ============================================================================
 * BOOKSTALL OPERATIONS & SALES TRACKER - BACKEND SCRIPT (code.gs)
 * Deployment Mode: Execute as ME (Database Private) with PIN Authentication
 * ============================================================================
 */

const SHEET_NAMES = {
  books: 'Books',
  locations: 'Locations',
  volunteers: 'Volunteers',
  events: 'Events',
  stock: 'Stock',
  sales: 'Sales'
};
const SESSION_TOKEN_PROPERTY_PREFIX = 'bookstall_session_';
const SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000;

function getSheetByName(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  for (let s of sheets) {
    if (s.getName().trim().toLowerCase() === sheetName.trim().toLowerCase()) {
      return s;
    }
  }

  const aliases = {
    books: ['catalogue', 'catalog', 'book', 'titles', 'booklist'],
    locations: ['location', 'places', 'locs'],
    volunteers: ['volunteer', 'users', 'members'],
    sales: ['sale', 'transactions', 'orders'],
    stock: ['inventory', 'stocks'],
    events: ['event', 'stall_events']
  };

  const checkAliases = aliases[sheetName.toLowerCase()] || [];
  for (let s of sheets) {
    if (checkAliases.includes(s.getName().trim().toLowerCase())) {
      return s;
    }
  }

  return spreadsheet.insertSheet(sheetName);
}

function getAllRows(sheetName) {
  const sheet = getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0].map((h) => String(h || '').trim());
  const rows = values.slice(1);

  return rows.map((row) => {
    const item = {};
    headers.forEach((header, index) => {
      if (header) item[header] = row[index];
    });
    return item;
  });
}

function ensureSheetsExist() {
  const defaultHeaders = {
    Books: ['bookId', 'bookName', 'category', 'language', 'active'],
    Locations: ['locationId', 'locationName', 'monthlyTarget', 'active'],
    Volunteers: ['volunteerId', 'name', 'email', 'role', 'pin', 'active', 'username'],
    Events: ['eventId', 'location', 'subLocation', 'date', 'volunteers', 'interactions', 'durationHours', 'booksDistributed', 'appInstalls', 'status', 'startTime'],
    Stock: ['stockId', 'bookName', 'volunteerName', 'quantityReceived', 'quantityRemaining'],
    Sales: ['saleId', 'timestamp', 'bookName', 'eventId', 'location', 'subLocation', 'sellerName', 'stockOwnerName', 'quantity', 'unitPrice', 'recordedBy', 'notes']
  };

  Object.entries(defaultHeaders).forEach(([sheetName, headers]) => {
    let sheet = getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      sheet.appendRow(headers);
    } else {
      const existingHeaders = data[0].map((h) => String(h || '').trim().toLowerCase());
      headers.forEach((expectedCol) => {
        if (!existingHeaders.includes(expectedCol.toLowerCase())) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(expectedCol);
          existingHeaders.push(expectedCol.toLowerCase());
        }
      });
    }
  });

  backfillUsernames_();

  const locSheet = getSheetByName(SHEET_NAMES.locations);
  if (locSheet.getLastRow() <= 1) {
    const defaultLocs = [
      ['LOC001', 'AECS Layout', 100, 'True'],
      ['LOC002', 'HSR Layout', 120, 'True'],
      ['LOC003', 'Seetharampalya Lake', 80, 'True'],
      ['LOC004', 'JP Nagar', 75, 'True'],
      ['LOC005', 'Malleshwaram', 100, 'True'],
      ['LOC006', 'Mahadevpura', 100, 'True'],
      ['LOC007', 'Bellandur', 80, 'True']
    ];
    locSheet.getRange(locSheet.getLastRow() + 1, 1, defaultLocs.length, 4).setValues(defaultLocs);
  }
}

// =============================================================================
// PIN AUTHENTICATION & SESSION MANAGEMENT
// =============================================================================

function normalizeUsername_(value) {
  return String(value || '').trim().toLowerCase();
}

// Existing volunteers get a username derived from email (or name) so they can log in.
function backfillUsernames_() {
  const sheet = getSheetByName(SHEET_NAMES.volunteers);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const headers = data[0].map((h) => String(h || '').trim().toLowerCase());
  const userIdx = headers.indexOf('username');
  const emailIdx = headers.indexOf('email');
  const nameIdx = headers.indexOf('name');
  if (userIdx === -1) return;

  const taken = new Set();
  for (let i = 1; i < data.length; i++) {
    const existing = normalizeUsername_(data[i][userIdx]);
    if (existing) taken.add(existing);
  }

  for (let i = 1; i < data.length; i++) {
    if (normalizeUsername_(data[i][userIdx])) continue;
    const email = emailIdx === -1 ? '' : String(data[i][emailIdx] || '');
    const base = (email.split('@')[0] || (nameIdx === -1 ? '' : String(data[i][nameIdx] || '')))
      .toLowerCase().replace(/[^a-z0-9._-]/g, '') || 'user' + i;
    let candidate = base;
    let suffix = 2;
    while (taken.has(candidate)) candidate = base + suffix++;
    taken.add(candidate);
    sheet.getRange(i + 1, userIdx + 1).setValue(candidate);
  }
}

function createSessionToken(volunteerId) {
  const properties = PropertiesService.getScriptProperties();
  const now = Date.now();
  Object.entries(properties.getProperties()).forEach(([key, value]) => {
    if (!key.startsWith(SESSION_TOKEN_PROPERTY_PREFIX)) return;
    try {
      const session = JSON.parse(value);
      if (!Number.isFinite(session.expiresAt) || session.expiresAt <= now) properties.deleteProperty(key);
    } catch (e) {
      properties.deleteProperty(key);
    }
  });

  const token = Utilities.getUuid();
  const session = {
    volunteerId: String(volunteerId),
    expiresAt: now + SESSION_TTL_MS
  };
  properties.setProperty(
    SESSION_TOKEN_PROPERTY_PREFIX + token,
    JSON.stringify(session)
  );
  return token;
}

function loginWithCredentials(username, inputPin) {
  ensureSheetsExist();
  const rows = getAllRows(SHEET_NAMES.volunteers);
  const targetUsername = normalizeUsername_(username);
  const pinEntered = String(inputPin || '').trim();

  const user = targetUsername && rows.find((v) => normalizeUsername_(v.username) === targetUsername);
  const targetId = user ? String(user.volunteerId || '').trim() : '';

  if (!user) {
    return { success: false, message: 'Invalid username or PIN.' };
  }

  const actVal = String(user.active ?? 'true').trim().toLowerCase();
  if (['0', 'false', 'no', 'inactive'].includes(actVal)) {
    return { success: false, message: 'Your volunteer account is currently disabled. Contact admin.' };
  }

  const storedPin = String(user.pin !== undefined && user.pin !== null ? user.pin : '').trim();

  if (!storedPin) {
    return { success: false, message: 'No PIN is set for this account. Contact admin.' };
  }
  if (storedPin !== pinEntered) {
    return { success: false, message: 'Invalid username or PIN.' };
  }

  const role = String(user.role || 'user').trim().toLowerCase();
  const token = createSessionToken(targetId);

  return {
    success: true,
    token: token,
    currentUser: {
      volunteerId: String(user.volunteerId || ''),
      name: String(user.name || '').trim(),
      email: String(user.email || '').trim(),
      role: role,
      isAdmin: role === 'admin'
    }
  };
}

function validateSessionToken(token) {
  if (!token) return null;
  const properties = PropertiesService.getScriptProperties();
  let serializedSession = properties.getProperty(SESSION_TOKEN_PROPERTY_PREFIX + token);

  if (!serializedSession) {
    try {
      const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
      const parts = decoded.split('::');
      if (parts.length < 3) return null;

      const issuedAt = Number(parts[2]);
      if (!Number.isFinite(issuedAt) || Date.now() - issuedAt > SESSION_TTL_MS || issuedAt > Date.now()) return null;

      const legacyRows = getAllRows(SHEET_NAMES.volunteers);
      const legacyUser = legacyRows.find((row) => String(row.volunteerId || '').trim() === parts[0]);
      if (!legacyUser || String(legacyUser.pin || '').trim() !== parts[1]) return null;

      serializedSession = JSON.stringify({ volunteerId: parts[0], expiresAt: issuedAt + SESSION_TTL_MS });
      properties.setProperty(SESSION_TOKEN_PROPERTY_PREFIX + token, serializedSession);
    } catch (e) {
      return null;
    }
  }

  try {
    const session = JSON.parse(serializedSession);
    if (!session.volunteerId || !Number.isFinite(session.expiresAt) || Date.now() >= session.expiresAt) {
      properties.deleteProperty(SESSION_TOKEN_PROPERTY_PREFIX + token);
      return null;
    }

    const rows = getAllRows(SHEET_NAMES.volunteers);
    const user = rows.find((v) => String(v.volunteerId || '').trim() === String(session.volunteerId));

    if (!user) return null;

    const actVal = String(user.active ?? 'true').trim().toLowerCase();
    if (['0', 'false', 'no', 'inactive'].includes(actVal)) return null;

    const role = String(user.role || 'user').trim().toLowerCase();
    return {
      volunteerId: String(user.volunteerId || ''),
      name: String(user.name || '').trim(),
      email: String(user.email || '').trim(),
      role: role,
      isAdmin: role === 'admin'
    };
  } catch (e) {
    return null;
  }
}

function revokeSessionToken(sessionToken) {
  if (sessionToken) {
    PropertiesService.getScriptProperties().deleteProperty(SESSION_TOKEN_PROPERTY_PREFIX + String(sessionToken));
  }
  return { success: true };
}

function updateVolunteerPinInSheet_(volunteerId, newPin) {
  const sheet = getSheetByName(SHEET_NAMES.volunteers);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map((h) => String(h || '').trim().toLowerCase());
  const idIdx = headers.indexOf('volunteerid');
  let pinIdx = headers.indexOf('pin');

  if (pinIdx === -1) {
    sheet.getRange(1, headers.length + 1).setValue('pin');
    pinIdx = headers.length;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx] || '').trim() === String(volunteerId).trim()) {
      sheet.getRange(i + 1, pinIdx + 1).setValue(String(newPin).trim());
      return true;
    }
  }
  return false;
}

// =============================================================================
// DATA LOOKUP
// =============================================================================

function getLookupData(sessionToken) {
  ensureSheetsExist();

  const currentUser = validateSessionToken(sessionToken);

  const rawBooks = getAllRows(SHEET_NAMES.books);
  const allVolunteers = getAllRows(SHEET_NAMES.volunteers);
  const allLocations = getAllRows(SHEET_NAMES.locations);
  const allEvents = getAllRows(SHEET_NAMES.events);

  const activeBooks = [];
  const categoriesSet = new Set();
  const languagesSet = new Set();

  rawBooks.forEach((row, idx) => {
    let name = '';
    let category = '';
    let language = '';
    let bookId = idx + 1;
    let isActive = true;

    Object.entries(row).forEach(([key, val]) => {
      const cleanKey = String(key || '').trim().toLowerCase().replace(/[\s\-_]/g, '');
      const strVal = String(val !== undefined && val !== null ? val : '').trim();

      if (['bookname', 'booktitle', 'title', 'book', 'name'].includes(cleanKey)) {
        if (strVal) name = strVal;
      } else if (['category', 'cat', 'genre'].includes(cleanKey)) {
        category = strVal;
      } else if (['language', 'lang'].includes(cleanKey)) {
        language = strVal;
      } else if (['bookid', 'id', 'sno', 'sn'].includes(cleanKey)) {
        if (strVal) bookId = strVal;
      } else if (['active', 'status', 'isactive'].includes(cleanKey)) {
        if (['0', 'false', 'no', 'inactive', 'disabled'].includes(strVal.toLowerCase())) {
          isActive = false;
        }
      }
    });

    if (name && isActive) {
      activeBooks.push({ bookId, bookName: name, category, language });
      if (category) categoriesSet.add(category);
      if (language) languagesSet.add(language);
    }
  });

  const activeVolunteers = allVolunteers
    .filter((v) => {
      const act = String(v.active ?? 'true').trim().toLowerCase();
      return !['0', 'false', 'no', 'inactive'].includes(act);
    })
    .map((v) => ({
      volunteerId: String(v.volunteerId || ''),
      name: String(v.name || ''),
      email: String(v.email || ''),
      role: String(v.role || 'user').toLowerCase(),
      active: true
    }));

  const safeAllVolunteers = allVolunteers.map((v) => ({
    volunteerId: String(v.volunteerId || ''),
    name: String(v.name || ''),
    username: String(v.username || ''),
    email: String(v.email || ''),
    role: String(v.role || 'user').toLowerCase(),
    active: !['0', 'false', 'no', 'inactive'].includes(String(v.active ?? 'true').trim().toLowerCase())
  }));

  const activeLocations = allLocations.filter((l) => {
    const act = String(l.active ?? 'true').trim().toLowerCase();
    return !['0', 'false', 'no', 'inactive'].includes(act);
  });

  const locationSubLocationsMap = {};
  allEvents.forEach((ev) => {
    const loc = String(ev.location || '').trim();
    const sub = String(ev.subLocation || '').trim();
    if (loc && sub) {
      if (!locationSubLocationsMap[loc]) locationSubLocationsMap[loc] = new Set();
      locationSubLocationsMap[loc].add(sub);
    }
  });

  const subLocationsExport = {};
  Object.entries(locationSubLocationsMap).forEach(([k, set]) => {
    subLocationsExport[k] = Array.from(set);
  });

  return {
    currentUser: currentUser,
    books: activeBooks,
    events: allEvents
      .map((ev) => ({
        eventId: String(ev.eventId || '').trim(),
        location: String(ev.location || '').trim(),
        subLocation: String(ev.subLocation || '').trim(),
        date: normalizeEventDate(ev.date),
        startTime: normalizeEventTime_(ev.startTime),
        volunteers: String(ev.volunteers || '').trim(),
        status: String(ev.status || '').trim().toLowerCase()
      }))
      .filter((ev) => ev.eventId && ev.date)
      .sort((a, b) => b.date.localeCompare(a.date)),
    volunteers: activeVolunteers,
    allVolunteers: safeAllVolunteers,
    locations: activeLocations,
    subLocationsMap: subLocationsExport,
    categories: Array.from(categoriesSet).sort(),
    languages: Array.from(languagesSet).sort()
  };
}

// =============================================================================
// SALES & EVENTS LOGIC
// =============================================================================

function normalizeEventDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const dateText = String(value || '').trim();
  const isoDate = dateText.match(/^\d{4}-\d{2}-\d{2}/);
  if (isoDate) return isoDate[0];
  if (!dateText) return '';

  const parsedDate = new Date(dateText);
  return isNaN(parsedDate.getTime())
    ? dateText
    : Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeEventTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return match ? match[1].padStart(2, '0') + ':' + match[2] : '';
}

function createOrFindEvent(location, subLocation, date) {
  const eventsSheet = getSheetByName(SHEET_NAMES.events);
  const values = eventsSheet.getDataRange().getValues();
  const headers = values[0].map((h) => String(h || '').trim().toLowerCase());
  const eventIdIdx = headers.indexOf('eventid');
  const locIdx = headers.indexOf('location');
  const subLocIdx = headers.indexOf('sublocation');
  const dateIdx = headers.indexOf('date');
  const statusIdx = headers.indexOf('status');

  const loc = String(location || '').trim();
  const subLoc = String(subLocation || '').trim();
  const dt = normalizeEventDate(date);

  const matchedIndex = values.findIndex((row, index) => index > 0 &&
    String(row[locIdx] || '').trim().toLowerCase() === loc.toLowerCase() &&
    normalizeEventDate(row[dateIdx]) === dt &&
    String(row[subLocIdx] || '').trim().toLowerCase() === subLoc.toLowerCase());

  if (matchedIndex !== -1) {
    const existingStatus = statusIdx === -1 ? '' : String(values[matchedIndex][statusIdx] || '').trim().toLowerCase();
    if (statusIdx !== -1 && existingStatus !== 'completed') {
      eventsSheet.getRange(matchedIndex + 1, statusIdx + 1).setValue('live');
    }
    return String(values[matchedIndex][eventIdIdx] || '');
  }

  const eventId = 'E-' + Date.now();
  const newRow = new Array(headers.length).fill('');
  const setCell = (col, val) => {
    const idx = headers.indexOf(col.toLowerCase());
    if (idx !== -1) newRow[idx] = val;
  };

  setCell('eventid', eventId);
  setCell('location', loc);
  setCell('sublocation', subLoc);
  setCell('date', dt);
  setCell('status', 'live');

  eventsSheet.appendRow(newRow);
  return eventId;
}

function createEvent(eventData, sessionToken) {
  ensureSheetsExist();
  if (!validateSessionToken(sessionToken)) {
    throw new Error('Your session has expired. Please log in again.');
  }

  const location = String((eventData && eventData.location) || '').trim();
  const subLocation = String((eventData && eventData.subLocation) || '').trim();
  const date = normalizeEventDate(eventData && eventData.date);
  if (!location) throw new Error('Major Location is required.');
  if (!subLocation) throw new Error('Sub-Location / Landmark is required.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('A valid event date is required.');

  const startTime = normalizeEventTime_(eventData && eventData.startTime);
  const rawVolunteers = eventData && eventData.volunteers;
  const volunteers = Array.isArray(rawVolunteers)
    ? rawVolunteers.map((name) => String(name || '').trim()).filter(Boolean).join(', ')
    : String(rawVolunteers || '').trim();

  const eventId = createOrFindEvent(location, subLocation, date);
  const found = findEventRowById_(eventId);
  const rowNumber = found.rowIndex + 1;
  const writeCell = (col, val, asText) => {
    const idx = found.headers.indexOf(col);
    if (idx === -1 || !val) return;
    const range = found.sheet.getRange(rowNumber, idx + 1);
    // Plain-text format stops Sheets turning "09:30" into a time value.
    if (asText) range.setNumberFormat('@');
    range.setValue(val);
  };
  writeCell('starttime', startTime, true);
  writeCell('volunteers', volunteers, false);
  const statusIdx = found.headers.indexOf('status');
  const status = statusIdx === -1 ? 'live' : String(found.values[found.rowIndex][statusIdx] || 'live').trim().toLowerCase();

  return {
    success: true,
    event: {
      eventId: eventId, location: location, subLocation: subLocation, date: date,
      startTime: startTime || normalizeEventTime_(found.values[found.rowIndex][found.headers.indexOf('starttime')]),
      volunteers: volunteers || String(found.values[found.rowIndex][found.headers.indexOf('volunteers')] || '').trim(),
      status: status
    },
    message: `Bookstall ready: ${subLocation}, ${location} on ${date}.`
  };
}

function findEventRowById_(eventId) {
  const eventsSheet = getSheetByName(SHEET_NAMES.events);
  const values = eventsSheet.getDataRange().getValues();
  const headers = values[0].map((h) => String(h || '').trim().toLowerCase());
  const idIdx = headers.indexOf('eventid');
  const targetId = String(eventId || '').trim();
  const rowIndex = targetId ? values.findIndex((row, i) => i > 0 && String(row[idIdx] || '').trim() === targetId) : -1;
  return { sheet: eventsSheet, values: values, headers: headers, rowIndex: rowIndex };
}

function saveEventSummary(summaryData, sessionToken) {
  ensureSheetsExist();
  if (!validateSessionToken(sessionToken)) {
    throw new Error('Your session has expired. Please log in again.');
  }

  const volunteers = String(summaryData.volunteers || '').trim();
  const interactions = String(summaryData.interactions || '').trim();
  const durationHours = summaryData.durationHours === '' || summaryData.durationHours == null ? '' : Number(summaryData.durationHours);
  const booksDistributed = Number(summaryData.booksDistributed || 0);
  const appInstalls = Number(summaryData.appInstalls || 0);

  if (durationHours !== '' && (!Number.isFinite(durationHours) || durationHours < 0)) {
    throw new Error('Duration must be a non-negative number of hours.');
  }
  if (!Number.isInteger(booksDistributed) || booksDistributed < 0) {
    throw new Error('Books distributed must be a non-negative whole number.');
  }

  const found = findEventRowById_(summaryData.eventId);
  if (found.rowIndex === -1) throw new Error('Event not found. Create the event first.');

  const { sheet, values, headers, rowIndex } = found;
  const rowNumber = rowIndex + 1;
  const setIfPresent = (col, val) => {
    const idx = headers.indexOf(col);
    if (idx !== -1) sheet.getRange(rowNumber, idx + 1).setValue(val);
  };
  setIfPresent('volunteers', volunteers);
  setIfPresent('interactions', interactions);
  setIfPresent('durationhours', durationHours);
  setIfPresent('booksdistributed', booksDistributed);
  setIfPresent('appinstalls', appInstalls);
  setIfPresent('status', 'completed');

  return {
    success: true,
    eventId: String(values[rowIndex][headers.indexOf('eventid')] || ''),
    message: 'Event summary saved.'
  };
}

function recordMultiItemSale(payload, sessionToken) {
  ensureSheetsExist();

  const currentUser = validateSessionToken(sessionToken);
  if (!currentUser) {
    throw new Error('Your session has expired. Please log in with your PIN again.');
  }

  if (!payload || !payload.items || !Array.isArray(payload.items) || payload.items.length === 0) {
    throw new Error('Please add at least one book to record a transaction.');
  }

  const found = findEventRowById_(payload.eventId);
  if (found.rowIndex === -1) throw new Error('Select an event. Create the event first if it does not exist.');
  const eventRow = found.values[found.rowIndex];
  const eventCell = (col) => String(eventRow[found.headers.indexOf(col)] || '').trim();
  const eventId = eventCell('eventid');
  const location = eventCell('location');
  const subLocation = eventCell('sublocation');
  const seller = String(payload.sellerName || '').trim();
  const recordedBy = currentUser.name;
  const notes = String(payload.notes || '').trim();

  if (!seller) throw new Error('Seller (Volunteer) is required.');

  payload.items.forEach((item, index) => {
    const book = String(item.bookName || '').trim();
    const qty = Number(item.quantity);
    const price = Number(item.unitPrice);

    if (!book) throw new Error(`Item #${index + 1}: Book selection is required.`);
    if (!Number.isInteger(qty) || qty <= 0) throw new Error(`Item #${index + 1} (${book}): Quantity must be at least 1.`);
    if (!Number.isFinite(price) || price < 0) throw new Error(`Item #${index + 1} (${book}): Unit price must be valid.`);
  });

  const salesSheet = getSheetByName(SHEET_NAMES.sales);
  const values = salesSheet.getDataRange().getValues();
  const headers = values[0].map((h) => String(h || '').trim().toLowerCase());

  const baseTxnId = 'TXN-' + Date.now();
  const timestamp = new Date().toISOString();

  payload.items.forEach((item, index) => {
    const qty = Number(item.quantity);
    const price = Number(item.unitPrice);
    const book = String(item.bookName).trim();
    const saleId = `${baseTxnId}-${index + 1}`;

    const row = new Array(headers.length).fill('');
    const setCell = (colName, val) => {
      const idx = headers.indexOf(colName.toLowerCase());
      if (idx !== -1) row[idx] = val;
    };

    setCell('saleid', saleId);
    setCell('timestamp', timestamp);
    setCell('bookname', book);
    setCell('eventid', eventId);
    setCell('location', location);
    setCell('sublocation', subLocation);
    setCell('sellername', seller);
    setCell('quantity', qty);
    setCell('unitprice', price);
    setCell('recordedby', recordedBy);
    setCell('notes', notes);

    salesSheet.appendRow(row);
  });

  return {
    success: true,
    transactionId: baseTxnId,
    itemCount: payload.items.length,
    message: `Recorded ${payload.items.length} book sale(s) successfully at ${location} (${subLocation}).`
  };
}

function getNormalizedSalesWithLocations() {
  const sales = getAllRows(SHEET_NAMES.sales);
  const events = getAllRows(SHEET_NAMES.events);

  const eventMap = {};
  events.forEach((e) => {
    const id = String(e.eventId || '').trim();
    if (id) {
      eventMap[id] = {
        location: String(e.location || e.name || '').trim(),
        subLocation: String(e.subLocation || '').trim(),
        date: String(e.date || '').trim()
      };
    }
  });

  return sales.map((s) => {
    const eventId = String(s.eventId || '').trim();
    const ev = eventMap[eventId] || { location: '', subLocation: '', date: '' };

    const location = String(s.location || ev.location || 'Unspecified').trim();
    const subLocation = String(s.subLocation || ev.subLocation || '').trim();

    let isoString = '';
    const rawTs = s.timestamp;
    if (rawTs instanceof Date) {
      try {
        isoString = rawTs.toISOString();
      } catch (e) {
        isoString = String(rawTs);
      }
    } else if (rawTs) {
      isoString = String(rawTs).trim();
    }

    const qty = Number(s.quantity || 0);
    const unitPrice = Number(s.unitPrice || 0);

    return {
      saleId: String(s.saleId || ''),
      timestamp: isoString,
      date: normalizeEventDate(ev.date) || (isoString ? isoString.substring(0, 10) : ''),
      bookName: String(s.bookName || '').trim(),
      eventId: eventId,
      location: location,
      subLocation: subLocation,
      sellerName: String(s.sellerName || '').trim(),
      stockOwnerName: String(s.stockOwnerName || '').trim(),
      quantity: qty,
      unitPrice: unitPrice,
      totalAmount: qty * unitPrice,
      recordedBy: String(s.recordedBy || s.sellerName || '').trim(),
      notes: String(s.notes || '').trim()
    };
  });
}

function getLocationAnalyticsAndTargets(monthFilter) {
  ensureSheetsExist();
  const sales = getNormalizedSalesWithLocations();
  const locations = getAllRows(SHEET_NAMES.locations);
  const filterMonth = String(monthFilter || '').trim();

  const monthSales = filterMonth ? sales.filter((s) => s.timestamp.startsWith(filterMonth) || s.date.startsWith(filterMonth)) : sales;

  const locTargetMap = {};
  locations.forEach((loc) => {
    const name = String(loc.locationName || '').trim();
    if (name) {
      locTargetMap[name.toLowerCase()] = {
        locationId: loc.locationId,
        locationName: name,
        target: Number(loc.monthlyTarget || 100),
        soldCopies: 0,
        revenue: 0,
        progressPercent: 0
      };
    }
  });

  const locationBookPrices = {};

  sales.forEach((s) => {
    const loc = s.location || 'Unspecified';
    const price = Number(s.unitPrice || 0);
    const book = s.bookName;

    if (loc && book && price > 0) {
      if (!locationBookPrices[loc]) locationBookPrices[loc] = {};
      if (!locationBookPrices[loc][book]) locationBookPrices[loc][book] = [];
      locationBookPrices[loc][book].push(price);
    }
  });

  monthSales.forEach((s) => {
    const loc = s.location || 'Unspecified';
    const locKey = loc.toLowerCase();
    const qty = Number(s.quantity || 0);
    const price = Number(s.unitPrice || 0);

    if (!locTargetMap[locKey]) {
      locTargetMap[locKey] = {
        locationId: 'LOC-CUSTOM',
        locationName: loc,
        target: 100,
        soldCopies: 0,
        revenue: 0,
        progressPercent: 0
      };
    }

    locTargetMap[locKey].soldCopies += qty;
    locTargetMap[locKey].revenue += (qty * price);
  });

  Object.values(locTargetMap).forEach((item) => {
    item.progressPercent = item.target > 0 ? Math.min(100, Math.round((item.soldCopies / item.target) * 100)) : 0;
  });

  const locationPriceInsights = [];
  Object.entries(locationBookPrices).forEach(([loc, booksMap]) => {
    Object.entries(booksMap).forEach(([book, prices]) => {
      const avgPrice = Math.round(prices.reduce((a, b) => a + b, 0) / prices.length);
      const minPrice = Math.min(...prices);
      const maxPrice = Math.max(...prices);
      locationPriceInsights.push({
        location: loc,
        bookName: book,
        averagePrice: avgPrice,
        priceRange: minPrice === maxPrice ? `₹${minPrice}` : `₹${minPrice} - ₹${maxPrice}`
      });
    });
  });

  locationPriceInsights.sort((a, b) => {
    const locCmp = a.location.localeCompare(b.location);
    return locCmp !== 0 ? locCmp : a.bookName.localeCompare(b.bookName);
  });

  return {
    filterMonth: filterMonth,
    totalRevenue: monthSales.reduce((sum, s) => sum + s.totalAmount, 0),
    totalCopiesSold: monthSales.reduce((sum, s) => sum + s.quantity, 0),
    locationTargets: Object.values(locTargetMap),
    locationPriceInsights: locationPriceInsights
  };
}

function getTransactionsHistory(filters) {
  ensureSheetsExist();
  const sales = getNormalizedSalesWithLocations();

  const filterRecordedBy = String((filters && filters.recordedBy) || '').trim().toLowerCase();
  const filterLocation = String((filters && filters.location) || '').trim().toLowerCase();
  const filterMonth = String((filters && filters.month) || '').trim();

  let matched = sales;

  if (filterRecordedBy) {
    matched = matched.filter((s) => s.recordedBy.toLowerCase() === filterRecordedBy || s.sellerName.toLowerCase() === filterRecordedBy);
  }
  if (filterLocation) {
    matched = matched.filter((s) => s.location.toLowerCase() === filterLocation);
  }
  if (filterMonth) {
    matched = matched.filter((s) => s.timestamp.startsWith(filterMonth) || s.date.startsWith(filterMonth));
  }

  return matched.reverse();
}

function getSalesPeriodSummary(startDate, endDate, sessionToken) {
  const currentUser = validateSessionToken(sessionToken);
  if (!currentUser) throw new Error('Your session has expired. Please log in with your PIN again.');
  if (!currentUser.isAdmin) throw new Error('Only admins can view sales summaries.');

  const start = String(startDate || '').trim();
  const end = String(endDate || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end) || start > end) {
    throw new Error('A valid summary date range is required.');
  }

  ensureSheetsExist();
  const sales = getNormalizedSalesWithLocations();
  const events = getAllRows(SHEET_NAMES.events);
  const inRange = (date) => Boolean(date) && date >= start && date <= end;
  const eventKeys = new Set();
  const eventDateById = new Map();
  const volunteers = new Map();
  let totalBooks = 0;
  let amountEarned = 0;

  const addVolunteer = (name) => {
    const cleanName = String(name || '').trim();
    if (cleanName && !volunteers.has(cleanName.toLowerCase())) {
      volunteers.set(cleanName.toLowerCase(), cleanName);
    }
  };

  events.forEach((event) => {
    let eventDate = '';
    if (event.date instanceof Date && !isNaN(event.date.getTime())) {
      eventDate = Utilities.formatDate(event.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      const eventDateText = String(event.date || '').trim();
      const isoDate = eventDateText.match(/^\d{4}-\d{2}-\d{2}/);
      if (isoDate) {
        eventDate = isoDate[0];
      } else if (eventDateText) {
        const parsedDate = new Date(eventDateText);
        if (!isNaN(parsedDate.getTime())) {
          eventDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      }
    }

    const eventId = String(event.eventId || '').trim();
    if (eventId && eventDate) eventDateById.set(eventId, eventDate);
    if (!inRange(eventDate)) return;
    const eventKey = eventId || [eventDate, event.location, event.subLocation].join('|');
    eventKeys.add(eventKey);
    String(event.volunteers || '').split(/[,;]|\s+and\s+/i).forEach(addVolunteer);
  });

  sales.forEach((sale) => {
    const eventDate = eventDateById.get(String(sale.eventId || '').trim()) || sale.date;
    if (!inRange(eventDate)) return;
    const quantity = Number(sale.quantity || 0);
    const unitPrice = Number(sale.unitPrice || 0);
    totalBooks += Number.isFinite(quantity) ? quantity : 0;
    amountEarned += Number.isFinite(quantity * unitPrice) ? quantity * unitPrice : 0;
  });

  return {
    startDate: start,
    endDate: end,
    totalBooks: totalBooks,
    amountEarned: amountEarned,
    eventCount: eventKeys.size,
    volunteerCount: volunteers.size
  };
}

function getInventory(stockOwnerFilter, statusFilter, sessionToken) {
  if (!validateSessionToken(sessionToken)) {
    throw new Error('Your session has expired. Please log in with your PIN again.');
  }

  ensureSheetsExist();
  const stockRows = getAllRows(SHEET_NAMES.stock);
  const normalizedOwner = String(stockOwnerFilter || '').trim().toLowerCase();
  const normalizedStatus = String(statusFilter || 'all').trim().toLowerCase();

  const inventoryMap = {};

  stockRows.forEach((row) => {
    const book = String(row.bookName || '').trim();
    const volunteer = String(row.volunteerName || '').trim();
    const qtyRemaining = Number(row.quantityRemaining ?? row.quantityReceived ?? 0);
    const qtyReceived = Number(row.quantityReceived ?? 0);

    if (!book || !volunteer) return;
    if (normalizedOwner && volunteer.toLowerCase() !== normalizedOwner) return;

    const key = volunteer + '|' + book;
    if (!inventoryMap[key]) {
      inventoryMap[key] = {
        volunteerName: volunteer,
        bookName: book,
        quantityRemaining: 0,
        quantityReceived: 0
      };
    }

    inventoryMap[key].quantityRemaining += Number.isFinite(qtyRemaining) ? qtyRemaining : 0;
    inventoryMap[key].quantityReceived += Number.isFinite(qtyReceived) ? qtyReceived : 0;
  });

  let results = Object.values(inventoryMap);

  if (normalizedStatus === 'available' || normalizedStatus === 'in_stock') {
    results = results.filter((r) => r.quantityRemaining > 0);
  } else if (normalizedStatus === 'out_of_stock' || normalizedStatus === 'sold_out') {
    results = results.filter((r) => r.quantityRemaining <= 0);
  }

  return results.sort((a, b) => a.bookName.localeCompare(b.bookName));
}

function adjustInventory(itemsArray, requestedVolunteerName, adjustmentType, sessionToken) {
  const currentUser = validateSessionToken(sessionToken);
  if (!currentUser) throw new Error('Your session has expired. Please log in with your PIN again.');

  ensureSheetsExist();
  const targetVolunteer = currentUser.isAdmin
    ? String(requestedVolunteerName || '').trim()
    : currentUser.name.trim();
  const operation = String(adjustmentType || '').trim().toLowerCase();
  if (!targetVolunteer) throw new Error('Please select a volunteer to manage stock for.');
  if (!['add', 'remove'].includes(operation)) throw new Error('Choose whether to add or remove stock.');
  if (!itemsArray || !Array.isArray(itemsArray) || itemsArray.length === 0) throw new Error('No stock changes to apply.');

  if (currentUser.isAdmin) {
    const activeVolunteers = getAllRows(SHEET_NAMES.volunteers).filter((volunteer) => {
      const active = String(volunteer.active ?? 'true').trim().toLowerCase();
      return !['0', 'false', 'no', 'inactive'].includes(active);
    });
    if (!activeVolunteers.some((volunteer) => String(volunteer.name || '').trim().toLowerCase() === targetVolunteer.toLowerCase())) {
      throw new Error('Select an active volunteer.');
    }
  }

  const quantitiesByBook = new Map();
  itemsArray.forEach((item) => {
    const book = String(item.bookName || '').trim();
    const quantity = Number(item.quantity);
    if (!book || !Number.isInteger(quantity) || quantity <= 0) {
      throw new Error('Each stock change needs a book and a positive whole-number quantity.');
    }
    const key = book.toLowerCase();
    const current = quantitiesByBook.get(key);
    quantitiesByBook.set(key, {
      bookName: current ? current.bookName : book,
      quantity: (current ? current.quantity : 0) + quantity
    });
  });

  const stockSheet = getSheetByName(SHEET_NAMES.stock);
  const values = stockSheet.getDataRange().getValues();
  const headers = values[0].map((header) => String(header || '').trim().toLowerCase());
  const bookIdx = headers.indexOf('bookname');
  const ownerIdx = headers.indexOf('volunteername');
  const recIdx = headers.indexOf('quantityreceived');
  const remIdx = headers.indexOf('quantityremaining');
  const stockIdIdx = headers.indexOf('stockid');
  if ([bookIdx, ownerIdx, recIdx, remIdx].some((index) => index === -1)) {
    throw new Error('Inventory columns were not found.');
  }

  const plannedUpdates = [];
  const rowsToAdd = [];
  let totalUnits = 0;
  quantitiesByBook.forEach(({ bookName, quantity }) => {
    const matchingRows = [];
    for (let index = 1; index < values.length; index++) {
      if (
        String(values[index][bookIdx] || '').trim().toLowerCase() === bookName.toLowerCase() &&
        String(values[index][ownerIdx] || '').trim().toLowerCase() === targetVolunteer.toLowerCase()
      ) {
        matchingRows.push(index);
      }
    }

    if (operation === 'add') {
      if (matchingRows.length) {
        const index = matchingRows[0];
        plannedUpdates.push({
          rowNumber: index + 1,
          quantityReceived: Number(values[index][recIdx] || 0) + quantity,
          quantityRemaining: Number(values[index][remIdx] || 0) + quantity
        });
      } else {
        const newRow = new Array(headers.length).fill('');
        if (stockIdIdx !== -1) newRow[stockIdIdx] = 'STK-' + Date.now() + '-' + rowsToAdd.length;
        newRow[bookIdx] = bookName;
        newRow[ownerIdx] = targetVolunteer;
        newRow[recIdx] = quantity;
        newRow[remIdx] = quantity;
        rowsToAdd.push(newRow);
      }
      totalUnits += quantity;
      return;
    }

    const available = matchingRows.reduce((sum, index) => sum + Number(values[index][remIdx] || 0), 0);
    if (available < quantity) {
      throw new Error(`${bookName} has only ${available} copies available for ${targetVolunteer}.`);
    }

    let remainingToRemove = quantity;
    matchingRows.forEach((index) => {
      if (remainingToRemove <= 0) return;
      const currentRemaining = Number(values[index][remIdx] || 0);
      const removed = Math.min(currentRemaining, remainingToRemove);
      plannedUpdates.push({ rowNumber: index + 1, quantityRemaining: currentRemaining - removed });
      remainingToRemove -= removed;
    });
    totalUnits += quantity;
  });

  plannedUpdates.forEach((update) => {
    if (update.quantityReceived !== undefined) {
      stockSheet.getRange(update.rowNumber, recIdx + 1).setValue(update.quantityReceived);
    }
    stockSheet.getRange(update.rowNumber, remIdx + 1).setValue(update.quantityRemaining);
  });
  rowsToAdd.forEach((row) => stockSheet.appendRow(row));

  return {
    success: true,
    adjustment: operation,
    adjustedTitles: quantitiesByBook.size,
    totalUnits: totalUnits,
    message: `${operation === 'add' ? 'Added' : 'Removed'} ${totalUnits} copies across ${quantitiesByBook.size} title(s) ${operation === 'add' ? 'to' : 'from'} ${targetVolunteer}'s inventory.`
  };
}

function normalizeText(str) {
  return String(str || '')
    .toLowerCase()
    .replace(/[\s\-_:',\.\(\)]/g, '');
}

const COMMON_ALIASES = {
  'bgv1': 'Bhagavad Gita - Vol 1',
  'bg1': 'Bhagavad Gita - Vol 1',
  'bhagavadgitavol1': 'Bhagavad Gita - Vol 1',
  'bgv2': 'Bhagavad Gita - Vol 2',
  'bg2': 'Bhagavad Gita - Vol 2',
  'ashtavakra': 'Ashtavakra Gita Eng',
  'womensrevolution': "Women's Revolution",
  'womanrevolution': "Women's Revolution",
  'upanishad': 'Upanishads',
  'upanishads': 'Upanishads',
  'stree': 'स्त्री',
  'darr': 'डर',
  'fear': 'Fear',
  'maa': 'माँ',
  'samay': 'समय',
  'sangharsh': 'संघर्ष',
  'joyfulrelationships': 'The Secret of Joyful Relationships HB',
  'secretofjoyfulrelationships': 'The Secret of Joyful Relationships HB'
};

function previewAndMatchBulkStock(rawText) {
  const lookup = getLookupData('');
  const catalogBooks = lookup.books.map((b) => b.bookName);
  const normalizedCatalogMap = {};

  catalogBooks.forEach((name) => {
    normalizedCatalogMap[normalizeText(name)] = name;
  });

  const lines = String(rawText || '').split('\n');
  const results = [];

  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) return;

    const parts = trimmed.split(/[,\t]/);
    const rawTitle = parts[0] ? parts[0].trim() : '';
    const qty = parts[1] ? parseInt(parts[1].trim(), 10) : 10;

    if (!rawTitle) return;

    const normInput = normalizeText(rawTitle);
    let matchedTitle = null;

    if (normalizedCatalogMap[normInput]) {
      matchedTitle = normalizedCatalogMap[normInput];
    }
    if (!matchedTitle && COMMON_ALIASES[normInput]) {
      matchedTitle = COMMON_ALIASES[normInput];
    }
    if (!matchedTitle) {
      for (let [normCat, canonical] of Object.entries(normalizedCatalogMap)) {
        if (normCat.includes(normInput) || normInput.includes(normCat)) {
          matchedTitle = canonical;
          break;
        }
      }
    }

    results.push({
      originalInput: rawTitle,
      matchedTitle: matchedTitle || '',
      isMatched: !!matchedTitle,
      quantity: isNaN(qty) || qty <= 0 ? 1 : qty
    });
  });

  return {
    items: results,
    catalogBooks: catalogBooks
  };
}

function updateLocationTarget(locationId, monthlyTarget, sessionToken) {
  const currentUser = validateSessionToken(sessionToken);
  if (!currentUser) throw new Error('Your session has expired. Please log in with your PIN again.');
  if (!currentUser.isAdmin) throw new Error('Only admins can update location targets.');

  const target = Number(monthlyTarget);
  if (!Number.isInteger(target) || target < 1) {
    throw new Error('Monthly target must be a positive whole number.');
  }

  ensureSheetsExist();
  const sheet = getSheetByName(SHEET_NAMES.locations);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map((header) => String(header || '').trim().toLowerCase());
  const idIdx = headers.indexOf('locationid');
  const targetIdx = headers.indexOf('monthlytarget');
  if (idIdx === -1 || targetIdx === -1) throw new Error('Location target columns were not found.');

  const normalizedId = String(locationId || '').trim();
  const rowIdx = values.findIndex((row, index) => index > 0 && String(row[idIdx] || '').trim() === normalizedId);
  if (rowIdx === -1) throw new Error('Location not found.');

  sheet.getRange(rowIdx + 1, targetIdx + 1).setValue(target);
  return { success: true, message: `Monthly target updated to ${target} books.` };
}

function addNewLocation(locationName, monthlyTarget) {
  ensureSheetsExist();
  const locName = String(locationName || '').trim();
  const target = Number(monthlyTarget) > 0 ? Number(monthlyTarget) : 100;

  if (!locName) throw new Error('Location name is required.');

  const sheet = getSheetByName(SHEET_NAMES.locations);
  const rows = getAllRows(SHEET_NAMES.locations);

  const exists = rows.some((r) => String(r.locationName || '').trim().toLowerCase() === locName.toLowerCase());
  if (exists) throw new Error('Location already exists.');

  const locId = 'LOC' + String(sheet.getLastRow()).padStart(3, '0');
  sheet.appendRow([locId, locName, target, 'True']);

  return { success: true, message: `Added location: "${locName}" with monthly target of ${target} books.` };
}

function addVolunteer(volunteerData, sessionToken) {
  const currentUser = validateSessionToken(sessionToken);
  if (!currentUser) throw new Error('Your session has expired. Please log in with your PIN again.');
  if (!currentUser.isAdmin) throw new Error('Only admins can add volunteers.');

  ensureSheetsExist();
  const name = String(volunteerData.name || '').trim();
  const username = normalizeUsername_(volunteerData.username);
  const email = String(volunteerData.email || '').trim().toLowerCase();
  const role = String(volunteerData.role || 'user').trim().toLowerCase();
  const initialPin = '1234';

  if (!name) throw new Error('Volunteer Name is required.');
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('Username must be 3-30 characters: letters, numbers, dot, dash or underscore.');
  if (!email || !email.includes('@')) throw new Error('A valid Gmail/Email address is required.');

  const volSheet = getSheetByName(SHEET_NAMES.volunteers);
  const rows = getAllRows(SHEET_NAMES.volunteers);

  const exists = rows.some((r) => String(r.email || '').trim().toLowerCase() === email);
  if (exists) throw new Error(`A volunteer with email "${email}" is already registered.`);
  if (rows.some((r) => normalizeUsername_(r.username) === username)) throw new Error(`Username "${username}" is already taken.`);

  const nextNum = volSheet.getLastRow();
  const volId = 'VOL' + String(nextNum).padStart(3, '0');

  const sheetHeaders = volSheet.getRange(1, 1, 1, volSheet.getLastColumn()).getValues()[0]
    .map((h) => String(h || '').trim().toLowerCase());
  const newRow = new Array(sheetHeaders.length).fill('');
  const values = { volunteerid: volId, name: name, email: email, role: role, pin: initialPin, active: 'True', username: username };
  Object.entries(values).forEach(([col, val]) => {
    const idx = sheetHeaders.indexOf(col);
    if (idx !== -1) newRow[idx] = val;
  });
  volSheet.appendRow(newRow);
  return { success: true, message: `Volunteer "${name}" added. Username: ${username}, initial PIN: 1234 (they can change it after signing in).` };
}

function updateOwnVolunteerPin(currentPin, newPin, sessionToken) {
  const currentUser = validateSessionToken(sessionToken);
  if (!currentUser) throw new Error('Your session has expired. Please log in with your PIN again.');

  const enteredPin = String(currentPin || '').trim();
  const updatedPin = String(newPin || '').trim();
  if (!/^\d{4,6}$/.test(updatedPin)) throw new Error('PIN must be 4 to 6 digits.');

  const rows = getAllRows(SHEET_NAMES.volunteers);
  const user = rows.find((row) => String(row.volunteerId || '').trim() === currentUser.volunteerId);
  if (!user || String(user.pin || '').trim() !== enteredPin) throw new Error('Current PIN is incorrect.');
  if (enteredPin === updatedPin) throw new Error('Choose a different PIN.');
  if (!updateVolunteerPinInSheet_(currentUser.volunteerId, updatedPin)) throw new Error('Could not update your PIN.');

  revokeSessionToken(sessionToken);
  const token = createSessionToken(currentUser.volunteerId);
  return { success: true, token: token, message: 'Your PIN was updated.' };
}

function updateVolunteerRole(volunteerId, newRole) {
  const volSheet = getSheetByName(SHEET_NAMES.volunteers);
  const data = volSheet.getDataRange().getValues();
  const headers = data[0].map((h) => String(h || '').trim().toLowerCase());
  const idIdx = headers.indexOf('volunteerid');
  const roleIdx = headers.indexOf('role');

  const normalizedRole = String(newRole).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx] || '').trim() === String(volunteerId).trim()) {
      volSheet.getRange(i + 1, roleIdx + 1).setValue(normalizedRole);
      return { success: true, message: `Updated volunteer role to ${normalizedRole}.` };
    }
  }
  throw new Error('Volunteer not found.');
}

function toggleVolunteerStatus(volunteerId, makeActive) {
  const volSheet = getSheetByName(SHEET_NAMES.volunteers);
  const data = volSheet.getDataRange().getValues();
  const headers = data[0].map((h) => String(h || '').trim().toLowerCase());
  const idIdx = headers.indexOf('volunteerid');
  const actIdx = headers.indexOf('active');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx] || '').trim() === String(volunteerId).trim()) {
      volSheet.getRange(i + 1, actIdx + 1).setValue(makeActive ? 'True' : 'False');
      return { success: true, message: `Volunteer status set to ${makeActive ? 'Active' : 'Inactive'}.` };
    }
  }
  throw new Error('Volunteer not found.');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Book Stall Operations & Sales')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}