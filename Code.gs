const SHEET_NAMES = {
  books: 'Books',
  volunteers: 'Volunteers',
  events: 'Events',
  stock: 'Stock',
  sales: 'Sales'
};

function getSheetByName(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Missing sheet: ' + sheetName);
  }

  return sheet;
}

function getAllRows(sheetName) {
  const sheet = getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  const headers = values[0];
  const rows = values.slice(1);

  return rows.map((row) => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  });
}

function ensureSheetsExist() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEET_NAMES).forEach((sheetName) => {
    if (!spreadsheet.getSheetByName(sheetName)) {
      spreadsheet.insertSheet(sheetName);
    }
  });
}

function getBooks() {
  ensureSheetsExist();
  return getAllRows(SHEET_NAMES.books);
}

function getVolunteers() {
  ensureSheetsExist();
  return getAllRows(SHEET_NAMES.volunteers);
}

function getEvents() {
  ensureSheetsExist();
  return getAllRows(SHEET_NAMES.events);
}

function getStock() {
  ensureSheetsExist();
  return getAllRows(SHEET_NAMES.stock);
}

function getSales() {
  ensureSheetsExist();
  return getAllRows(SHEET_NAMES.sales);
}

function getInventoryByVolunteer(stockOwnerName) {
  const stockRows = getStock();
  const salesRows = getSales();
  const summaryMap = {};
  const normalizedFilter = String(stockOwnerName || '').trim();

  stockRows.forEach((row) => {
    const bookName = String(row.bookName || '').trim();
    const volunteerName = String(row.volunteerName || '').trim();
    const quantity = Number(row.quantityReceived || 0);

    if (!bookName || !volunteerName || !Number.isFinite(quantity)) {
      return;
    }

    if (normalizedFilter && volunteerName !== normalizedFilter) {
      return;
    }

    const key = volunteerName + '|' + bookName;
    summaryMap[key] = (summaryMap[key] || 0) + quantity;
  });

  salesRows.forEach((row) => {
    const bookName = String(row.bookName || '').trim();
    const volunteerName = String(row.stockOwnerName || '').trim();
    const quantity = Number(row.quantity || 0);

    if (!bookName || !volunteerName || !Number.isFinite(quantity)) {
      return;
    }

    if (normalizedFilter && volunteerName !== normalizedFilter) {
      return;
    }

    const key = volunteerName + '|' + bookName;
    summaryMap[key] = (summaryMap[key] || 0) - quantity;
  });

  return Object.keys(summaryMap)
    .sort()
    .map((key) => {
      const [volunteerName, bookName] = key.split('|');
      const available = Number(summaryMap[key]) || 0;

      return {
        volunteerName: volunteerName || '',
        bookName: bookName || '',
        quantity: available > 0 ? available : 0
      };
    })
    .filter((row) => row.volunteerName && row.bookName && row.quantity > 0);
}

function getLookupData() {
  const books = getBooks();
  const volunteers = getVolunteers();

  return {
    books: books.filter((book) => {
      const active = String(book.active || '').trim().toLowerCase();
      return active === 'true' || active === 'yes' || active === '1';
    }),
    volunteers: volunteers.filter((volunteer) => {
      const active = String(volunteer.active || '').trim().toLowerCase();
      return active === 'true' || active === 'yes' || active === '1';
    })
  };
}

function getBookByName(bookName) {
  const books = getBooks();
  const matchedBook = books.find((book) => {
    return String(book.bookName || book.name || '').trim() === String(bookName || '').trim();
  });

  return matchedBook || null;
}

function validateSaleRequest(data) {
  if (!data) {
    throw new Error('No entry data was provided.');
  }

  const normalizedBookName = String(data.bookName || '').trim();
  if (!normalizedBookName) {
    throw new Error('Missing required field: bookName');
  }

  const matchedBook = getBookByName(normalizedBookName);
  if (!matchedBook) {
    throw new Error('Book not found: ' + normalizedBookName);
  }

  const requiredFields = [
    'eventLocation',
    'eventDate',
    'sellerName',
    'stockOwnerName',
    'quantity',
    'unitPrice',
    'recordedBy'
  ];

  requiredFields.forEach((field) => {
    const value = data[field];
    if (value === undefined || value === null || String(value).trim() === '') {
      throw new Error('Missing required field: ' + field);
    }
  });

  const quantity = Number(data.quantity);
  if (!Number.isInteger(quantity) || quantity <= 0) {
    throw new Error('Quantity must be a positive whole number.');
  }

  const unitPrice = Number(data.unitPrice);
  if (!Number.isFinite(unitPrice) || unitPrice < 0) {
    throw new Error('Unit price must be zero or greater.');
  }

  return {
    bookId: String(matchedBook.bookId || '').trim(),
    bookName: normalizedBookName,
    eventLocation: String(data.eventLocation).trim(),
    eventDate: String(data.eventDate).trim(),
    sellerName: String(data.sellerName).trim(),
    stockOwnerName: String(data.stockOwnerName).trim(),
    quantity: quantity,
    unitPrice: unitPrice,
    recordedBy: String(data.recordedBy).trim()
  };
}

function createOrFindEvent(eventLocation, eventDate) {
  const eventsSheet = getSheetByName(SHEET_NAMES.events);
  const rows = getAllRows(SHEET_NAMES.events);

  const normalizedLocation = String(eventLocation || '').trim();
  const normalizedDate = String(eventDate || '').trim();

  if (!normalizedLocation) {
    throw new Error('Event location is required.');
  }

  if (!normalizedDate) {
    throw new Error('Event date is required.');
  }

  const existingEvent = rows.find((row) => {
    return String(row.location || '').trim() === normalizedLocation &&
      String(row.date || '').trim() === normalizedDate;
  });

  if (existingEvent && existingEvent.eventId) {
    return String(existingEvent.eventId);
  }

  const eventId = 'E-' + Date.now();
  const eventName = normalizedLocation;

  eventsSheet.appendRow([eventId, normalizedLocation, eventName, normalizedDate]);

  return eventId;
}

function getAvailableStockForBookAndOwner(bookName, stockOwnerName) {
  const stockRows = getStock();
  const salesRows = getSales();
  const normalizedBookName = String(bookName || '').trim();

  const matchingStock = stockRows.filter((row) => {
    const rowBookName = String(row.bookName || row.name || '').trim();
    const rowOwner = String(row.volunteerName || '').trim();

    return rowBookName && normalizedBookName && rowBookName === normalizedBookName && rowOwner === String(stockOwnerName).trim();
  });

  const totalReceived = matchingStock.reduce((sum, row) => {
    const received = Number(row.quantityReceived || 0);
    return sum + (Number.isFinite(received) ? received : 0);
  }, 0);

  const soldQuantity = salesRows.reduce((sum, row) => {
    const rowBookName = String(row.bookName || row.bookId || '').trim();
    const matchesBook = rowBookName === normalizedBookName ||
      String(row.bookId || '').trim() === String(getBookByName(normalizedBookName)?.bookId || '').trim();
    const matchesOwner = String(row.stockOwnerName).trim() === String(stockOwnerName).trim();

    if (matchesBook && matchesOwner) {
      const qty = Number(row.quantity || 0);
      return sum + (Number.isFinite(qty) ? qty : 0);
    }

    return sum;
  }, 0);

  return totalReceived - soldQuantity;
}

function updateStockAfterSale(bookName, stockOwnerName, quantity) {
  const stockSheet = getSheetByName(SHEET_NAMES.stock);
  const rows = stockSheet.getDataRange().getValues();

  if (rows.length <= 1) {
    throw new Error('No stock rows found for this book and owner.');
  }

  const headers = rows[0];
  const bookIndex = headers.indexOf('bookName');
  const ownerIndex = headers.indexOf('volunteerName');
  const qtyIndex = headers.indexOf('quantityReceived');

  if (bookIndex === -1 || ownerIndex === -1 || qtyIndex === -1) {
    throw new Error('Stock sheet is missing required columns: bookName, volunteerName, or quantityReceived.');
  }

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rowBookName = String(row[bookIndex] || '').trim();
    const rowOwnerName = String(row[ownerIndex] || '').trim();

    if (rowBookName === String(bookName).trim() && rowOwnerName === String(stockOwnerName).trim()) {
      const currentQty = Number(row[qtyIndex] || 0);
      const updatedQty = currentQty - Number(quantity);

      if (updatedQty < 0) {
        throw new Error('Stock would go negative.');
      }

      stockSheet.getRange(i + 1, qtyIndex + 1).setValue(updatedQty);
      return;
    }
  }

  throw new Error('No stock row found for this book and owner.');
}

function validateStockEntry(data) {
  if (!data) {
    throw new Error('No stock data was provided.');
  }

  const bookName = String(data.bookName || '').trim();
  if (!bookName) {
    throw new Error('Missing required field: bookName');
  }

  if (!getBookByName(bookName)) {
    throw new Error('Book not found: ' + bookName);
  }

  const volunteerName = String(data.volunteerName || '').trim();
  if (!volunteerName) {
    throw new Error('Missing required field: volunteerName');
  }

  const quantityReceived = Number(data.quantityReceived);
  if (!Number.isInteger(quantityReceived) || quantityReceived <= 0) {
    throw new Error('Quantity received must be a positive whole number.');
  }

  return {
    bookName: bookName,
    volunteerName: volunteerName,
    quantityReceived: quantityReceived
  };
}

function addStock(data) {
  const cleanedData = validateStockEntry(data);
  const stockSheet = getSheetByName(SHEET_NAMES.stock);
  const rows = stockSheet.getDataRange().getValues();

  const headers = rows.length
    ? rows[0]
    : ['stockId', 'bookName', 'volunteerName', 'quantityReceived'];

  const bookIndex = headers.indexOf('bookName');
  const ownerIndex = headers.indexOf('volunteerName');
  const qtyIndex = headers.indexOf('quantityReceived');
  const stockIdIndex = headers.indexOf('stockId');

  if (bookIndex === -1 || ownerIndex === -1 || qtyIndex === -1) {
    throw new Error('Stock sheet is missing required columns: bookName, volunteerName, or quantityReceived.');
  }

  if (rows.length === 0) {
    stockSheet.appendRow(headers);
  }

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const rowBookName = String(row[bookIndex] || '').trim();
    const rowVolunteerName = String(row[ownerIndex] || '').trim();

    if (rowBookName === cleanedData.bookName && rowVolunteerName === cleanedData.volunteerName) {
      const currentQty = Number(row[qtyIndex] || 0);
      const updatedQty = currentQty + cleanedData.quantityReceived;
      stockSheet.getRange(i + 1, qtyIndex + 1).setValue(updatedQty);
      return {
        success: true,
        message: 'Stock added successfully.'
      };
    }
  }

  const newRow = Array(headers.length).fill('');
  if (stockIdIndex !== -1) {
    newRow[stockIdIndex] = '';
  }
  newRow[bookIndex] = cleanedData.bookName;
  newRow[ownerIndex] = cleanedData.volunteerName;
  newRow[qtyIndex] = cleanedData.quantityReceived;

  stockSheet.appendRow(newRow);

  return {
    success: true,
    message: 'Stock added successfully.'
  };
}

function recordSale(data) {
  const cleanedData = validateSaleRequest(data);
  const eventId = createOrFindEvent(cleanedData.eventLocation, cleanedData.eventDate);

  const availableStock = getAvailableStockForBookAndOwner(
    cleanedData.bookName,
    cleanedData.stockOwnerName
  );

  if (cleanedData.quantity > availableStock) {
    throw new Error('Sale cannot exceed the available stock for this book and stock owner.');
  }

  updateStockAfterSale(cleanedData.bookName, cleanedData.stockOwnerName, cleanedData.quantity);

  const salesSheet = getSheetByName(SHEET_NAMES.sales);
  const saleId = 'SALE-' + Date.now();

  salesSheet.appendRow([
    saleId,
    new Date().toISOString(),
    cleanedData.bookName,
    eventId,
    cleanedData.sellerName,
    cleanedData.stockOwnerName,
    cleanedData.quantity,
    cleanedData.unitPrice,
    cleanedData.recordedBy
  ]);

  return {
    success: true,
    message: 'Entry saved successfully.',
    saleId: saleId
  };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}
