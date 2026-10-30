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
    Volunteers: ['volunteerId', 'name', 'email', 'role', 'pin', 'active'],
    Events: ['eventId', 'location', 'subLocation', 'date', 'volunteers', 'interactions', 'durationHours', 'booksDistributed', 'appInstalls', 'status'],
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

function getVolunteersForLogin() {
  ensureSheetsExist();
  const rows = getAllRows(SHEET_NAMES.volunteers);

  return rows
    .filter((v) => {
      const act = String(v.active ?? 'true').trim().toLowerCase();
      return !['0', 'false', 'no', 'inactive'].includes(act);
    })
    .map((v) => ({
      volunteerId: String(v.volunteerId || ''),
      name: String(v.name || '').trim()
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

function verifyVolunteerPin(volunteerId, inputPin) {
  ensureSheetsExist();
  const rows = getAllRows(SHEET_NAMES.volunteers);
  const targetId = String(volunteerId || '').trim();
  const pinEntered = String(inputPin || '').trim();

  const user = rows.find((v) => String(v.volunteerId || '').trim() === targetId);

  if (!user) {
    return { success: false, message: 'Volunteer profile not found.' };
  }

  const actVal = String(user.active ?? 'true').trim().toLowerCase();
  if (['0', 'false', 'no', 'inactive'].includes(actVal)) {
    return { success: false, message: 'Your volunteer account is currently disabled. Contact admin.' };
  }

  const storedPin = String(user.pin !== undefined && user.pin !== null ? user.pin : '').trim();

  if (!storedPin) {
    updateVolunteerPinInSheet(targetId, pinEntered);
  } else if (storedPin !== pinEntered) {
    return { success: false, message: 'Incorrect PIN. Please try again.' };
  }

  const role = String(user.role || 'user').trim().toLowerCase();
  const token = Utilities.base64Encode(targetId + '::' + pinEntered + '::' + Date.now());

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
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts = decoded.split('::');
    if (parts.length < 2) return null;

    const volunteerId = parts[0];
    const pin = parts[1];

    const rows = getAllRows(SHEET_NAMES.volunteers);
    const user = rows.find((v) => String(v.volunteerId || '').trim() === volunteerId);

    if (!user) return null;

    const actVal = String(user.active ?? 'true').trim().toLowerCase();
    if (['0', 'false', 'no', 'inactive'].includes(actVal)) return null;

    const storedPin = String(user.pin !== undefined && user.pin !== null ? user.pin : '').trim();
    if (storedPin !== pin) return null;

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

function updateVolunteerPinInSheet(volunteerId, newPin) {
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
    email: String(v.email || ''),
    role: String(v.role || 'user').toLowerCase(),
    pin: String(v.pin || ''),
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
    if (statusIdx !== -1) eventsSheet.getRange(matchedIndex + 1, statusIdx + 1).setValue('live');
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

function saveEventSummary(summaryData) {
  ensureSheetsExist();
  const location = String(summaryData.location || '').trim();
  const subLocation = String(summaryData.subLocation || '').trim();
  const date = String(summaryData.date || '').trim();
  const volunteers = String(summaryData.volunteers || '').trim();
  const interactions = String(summaryData.interactions || '').trim();
  const durationHours = summaryData.durationHours === '' || summaryData.durationHours == null ? '' : Number(summaryData.durationHours);
  const booksDistributed = Number(summaryData.booksDistributed || 0);
  const appInstalls = Number(summaryData.appInstalls || 0);

  if (!location) throw new Error('Location is required.');
  if (!subLocation) throw new Error('Sub-location is required.');
  if (!date) throw new Error('Date is required.');
  if (durationHours !== '' && (!Number.isFinite(durationHours) || durationHours < 0)) {
    throw new Error('Duration must be a non-negative number of hours.');
  }
  if (!Number.isInteger(booksDistributed) || booksDistributed < 0) {
    throw new Error('Books distributed must be a non-negative whole number.');
  }

  const eventsSheet = getSheetByName(SHEET_NAMES.events);
  const values = eventsSheet.getDataRange().getValues();

  const headers = values[0].map((h) => String(h || '').trim().toLowerCase());
  const locIdx = headers.indexOf('location');
  const subLocIdx = headers.indexOf('sublocation');
  const dateIdx = headers.indexOf('date');
  const volIdx = headers.indexOf('volunteers');
  const intIdx = headers.indexOf('interactions');
  const durationHoursIdx = headers.indexOf('durationhours');
  const booksDistributedIdx = headers.indexOf('booksdistributed');
  const appIdx = headers.indexOf('appinstalls');

  const matchedIndex = values.findIndex((row, index) => index > 0 &&
    String(row[locIdx] || '').trim().toLowerCase() === location.toLowerCase() &&
    String(row[subLocIdx] || '').trim().toLowerCase() === subLocation.toLowerCase() &&
    normalizeEventDate(row[dateIdx]) === normalizeEventDate(date));

  if (matchedIndex !== -1) {
    const matchedRowIndex = matchedIndex + 1;
    if (volIdx !== -1) eventsSheet.getRange(matchedRowIndex, volIdx + 1).setValue(volunteers);
    if (intIdx !== -1) eventsSheet.getRange(matchedRowIndex, intIdx + 1).setValue(interactions);
    if (durationHoursIdx !== -1) eventsSheet.getRange(matchedRowIndex, durationHoursIdx + 1).setValue(durationHours);
    if (booksDistributedIdx !== -1) eventsSheet.getRange(matchedRowIndex, booksDistributedIdx + 1).setValue(booksDistributed);
    if (appIdx !== -1) eventsSheet.getRange(matchedRowIndex, appIdx + 1).setValue(appInstalls);

    return {
      success: true,
      eventId: String(values[matchedIndex][headers.indexOf('eventid')] || ''),
      message: `Updated stall event (${location} - ${subLocation} on ${date}).`
    };
  }

  const eventId = 'E-' + Date.now();
  const newRow = new Array(headers.length).fill('');
  const setCell = (col, val) => {
    const idx = headers.indexOf(col.toLowerCase());
    if (idx !== -1) newRow[idx] = val;
  };

  setCell('eventid', eventId);
  setCell('location', location);
  setCell('sublocation', subLocation);
  setCell('date', date);
  setCell('volunteers', volunteers);
  setCell('interactions', interactions);
  setCell('durationhours', durationHours);
  setCell('booksdistributed', booksDistributed);
  setCell('appinstalls', appInstalls);

  eventsSheet.appendRow(newRow);
  return { success: true, eventId: eventId, message: `Created new event for ${location} - ${subLocation}.` };
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

  const location = String(payload.location || '').trim();
  const subLocation = String(payload.subLocation || '').trim();
  const date = String(payload.eventDate || '').trim();
  const seller = String(payload.sellerName || '').trim();
  const stockOwner = String(payload.stockOwnerName || '').trim();
  const recordedBy = currentUser.name;
  const notes = String(payload.notes || '').trim();

  if (!location) throw new Error('Major Location is required.');
  if (!subLocation) throw new Error('Sub-Location / Landmark is mandatory.');
  if (!date) throw new Error('Event date is required.');
  if (!seller) throw new Error('Seller (Volunteer) is required.');

  payload.items.forEach((item, index) => {
    const book = String(item.bookName || '').trim();
    const qty = Number(item.quantity);
    const price = Number(item.unitPrice);

    if (!book) throw new Error(`Item #${index + 1}: Book selection is required.`);
    if (!Number.isInteger(qty) || qty <= 0) throw new Error(`Item #${index + 1} (${book}): Quantity must be at least 1.`);
    if (!Number.isFinite(price) || price < 0) throw new Error(`Item #${index + 1} (${book}): Unit price must be valid.`);
  });

  const eventId = createOrFindEvent(location, subLocation, date);

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
    setCell('stockownername', stockOwner);
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

function getInventory(stockOwnerFilter, statusFilter) {
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

function commitStockAllocation(itemsArray, volunteerName) {
  ensureSheetsExist();
  const targetVolunteer = String(volunteerName || '').trim();
  if (!targetVolunteer) throw new Error('Please select a volunteer to assign the stock to.');
  if (!itemsArray || !Array.isArray(itemsArray) || itemsArray.length === 0) throw new Error('No items to allocate.');

  const stockSheet = getSheetByName(SHEET_NAMES.stock);
  const values = stockSheet.getDataRange().getValues();
  const headers = values[0].map((h) => String(h || '').trim().toLowerCase());

  const bookIdx = headers.indexOf('bookname');
  const ownerIdx = headers.indexOf('volunteername');
  const recIdx = headers.indexOf('quantityreceived');
  const remIdx = headers.indexOf('quantityremaining');

  let addedCount = 0;
  let totalUnits = 0;

  itemsArray.forEach((item) => {
    const book = String(item.bookName || '').trim();
    const qty = Number(item.quantity);
    if (!book || !Number.isInteger(qty) || qty <= 0) return;

    let matched = false;
    for (let i = 1; i < values.length; i++) {
      if (
        String(values[i][bookIdx] || '').trim().toLowerCase() === book.toLowerCase() &&
        String(values[i][ownerIdx] || '').trim().toLowerCase() === targetVolunteer.toLowerCase()
      ) {
        const curRec = Number(values[i][recIdx] ?? 0);
        const curRem = Number((remIdx !== -1 ? values[i][remIdx] : values[i][recIdx]) ?? 0);
        stockSheet.getRange(i + 1, recIdx + 1).setValue(curRec + qty);
        if (remIdx !== -1) stockSheet.getRange(i + 1, remIdx + 1).setValue(curRem + qty);
        matched = true;
        break;
      }
    }

    if (!matched) {
      const newRow = Array(headers.length).fill('');
      newRow[bookIdx] = book;
      newRow[ownerIdx] = targetVolunteer;
      newRow[recIdx] = qty;
      if (remIdx !== -1) newRow[remIdx] = qty;
      stockSheet.appendRow(newRow);
    }
    addedCount++;
    totalUnits += qty;
  });

  return {
    success: true,
    addedTitles: addedCount,
    totalUnits: totalUnits,
    message: `Allocated ${totalUnits} copies across ${addedCount} title(s) to ${targetVolunteer}.`
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

function addVolunteer(volunteerData) {
  ensureSheetsExist();
  const name = String(volunteerData.name || '').trim();
  const email = String(volunteerData.email || '').trim().toLowerCase();
  const role = String(volunteerData.role || 'user').trim().toLowerCase();
  const pin = String(volunteerData.pin || '1234').trim();

  if (!name) throw new Error('Volunteer Name is required.');
  if (!email || !email.includes('@')) throw new Error('A valid Gmail/Email address is required.');

  const volSheet = getSheetByName(SHEET_NAMES.volunteers);
  const rows = getAllRows(SHEET_NAMES.volunteers);

  const exists = rows.some((r) => String(r.email || '').trim().toLowerCase() === email);
  if (exists) throw new Error(`A volunteer with email "${email}" is already registered.`);

  const nextNum = volSheet.getLastRow();
  const volId = 'VOL' + String(nextNum).padStart(3, '0');

  volSheet.appendRow([volId, name, email, role, pin, 'True']);
  return { success: true, message: `Volunteer "${name}" added successfully with PIN ${pin}.` };
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