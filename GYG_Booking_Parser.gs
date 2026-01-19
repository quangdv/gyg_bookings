/******************** CONFIG ********************/
const ROOT_FOLDER_ID = '1bv4A4wUX757bTALr9FbIns90O7iRc3p1';
const SOURCE_LABEL = 'GYG_BOOKING';
const PROCESSED_LABEL = 'GYG_BOOKING_READ';
const TIMEZONE = 'Asia/Ho_Chi_Minh';

/******************** MAIN ********************/
function processGYGBookings() {
  const sourceLabel = getOrCreateLabel(SOURCE_LABEL);
  const processedLabel = getOrCreateLabel(PROCESSED_LABEL);

  const threads = sourceLabel.getThreads();
  Logger.log('Found threads: ' + threads.length);

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1];

    const booking = parseGYGBooking(msg);
    if (!booking) {
      Logger.log('Skip email: cannot parse booking');
      return;
    }

    const sheet = getOrCreateDailySheet(booking.checkinDate);
    appendBookingRow(sheet, booking);

    thread.removeLabel(sourceLabel);
    thread.addLabel(processedLabel);
  });
}

/******************** PARSER ********************/
function parseGYGBooking(message) {
  if (!message) return null;

  const body = message.getBody();
  const text = body.replace(/<[^>]+>/g, '\n').replace(/\s+/g, ' ').trim();

  // Fixed regex to prevent duplicate tour names - stop at first occurrence
  const tourMatch = text.match(/Your offer has been booked:\s*([^]*?)\s*(?:Reference number|Date)/i);
  if (!tourMatch) return null;

  // Extract tour name and clean it up
  let tour = tourMatch[1].trim();

  // Remove duplicate tour names - detect if tour name is repeated
  // Split the tour text and check if it contains the same pattern twice
  const tourLength = tour.length;
  if (tourLength > 20) {
    // Check if the second half starts with the first half (indicating duplication)
    const midPoint = Math.floor(tourLength / 2);
    const firstPart = tour.substring(0, midPoint).trim();
    const secondPart = tour.substring(midPoint).trim();

    // If second part starts with first part (allowing for some variation), it's likely duplicated
    if (secondPart.length > 0 && firstPart.length > 0) {
      const firstPartStart = firstPart.substring(0, Math.min(30, firstPart.length));
      if (secondPart.substring(0, Math.min(30, secondPart.length)) === firstPartStart) {
        tour = firstPart;
      }
    }
  }

  // Replace &amp; with &
  tour = tour.replace(/&amp;/g, '&');

  const dateMatch = text.match(/Date\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})/);
  const checkinDate = dateMatch ? new Date(dateMatch[1]) : null;

  const customerMatch = text.match(/Main customer\s*([A-Za-z\s]+)/i);
  const customer = customerMatch ? customerMatch[1].trim() : '';

  const guestMatch = text.match(/(\d+)\s*x\s*Adults?/i);
  const adults = guestMatch ? Number(guestMatch[1]) : 0;

  const pickupMatch = text.match(/Pickup\s*(.*?)\s*(Open in Google Maps|Price)/i);
  const pickup = pickupMatch ? pickupMatch[1].trim() : '';

  return {
    tour,
    customer,
    checkinDate,
    checkoutDate: addDays(checkinDate, 1),
    adults,
    children: 0,
    infant: 0,
    pickup,
    pickupTime: '8:00 to 8:30 AM'
  };
}

/******************** SHEET ********************/
function getOrCreateDailySheet(dateObj) {
  const safeDate = dateObj instanceof Date && !isNaN(dateObj) ? dateObj : new Date();
  const year = safeDate.getFullYear();
  const month = String(safeDate.getMonth() + 1).padStart(2, '0');
  const dateStr = Utilities.formatDate(safeDate, TIMEZONE, 'yyyy-MM-dd');

  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const yearFolder = getOrCreateFolder(root, year);
  const monthFolder = getOrCreateFolder(yearFolder, month);

  const fileName = `GetYourGuide_Bookings_${dateStr}`;
  const files = monthFolder.getFilesByName(fileName);

  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(ss.getId()).moveTo(monthFolder);
    setupSheet(ss.getActiveSheet());
  }

  return ss.getActiveSheet();
}

function setupSheet(sheet) {
  const headers = [
    'Tour','Customer Name','Checkin','Checkout',
    'Adult','Children','Infant',
    'Double/Twin','Triple','Single',
    'Peak season','Bus','VAT','Holiday','Other','Cruise',
    'Pickup','Pickup time','Status'
  ];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Chờ xử lý','Chờ gửi email confirmation','Đã gửi email confirm'], true)
    .build();

  sheet.getRange('S2:S').setDataValidation(rule);
}

/******************** APPEND ********************/
function appendBookingRow(sheet, b) {
  const rooms = calculateRooms(b.adults + b.children);

  sheet.appendRow([
    b.tour,
    b.customer,
    b.checkinDate,
    b.checkoutDate,
    b.adults,
    b.children,
    b.infant,
    rooms.double,
    rooms.triple,
    rooms.single,
    '', '', '', '', '', '',
    b.pickup,
    b.pickupTime,
    'Chờ xử lý'
  ]);
}

/******************** HELPERS ********************/
function calculateRooms(x) {
  if (x <= 0) return { double:0, triple:0, single:0 };
  if (x === 1) return { double:0, triple:0, single:1 };
  if (x === 2) return { double:1, triple:0, single:0 };
  if (x === 3) return { double:0, triple:1, single:0 };
  if (x % 2 === 0) return { double:x/2, triple:0, single:0 };
  return { double:(x-3)/2, triple:1, single:0 };
}

function addDays(d, days) {
  if (!(d instanceof Date)) return '';
  const r = new Date(d);
  r.setDate(r.getDate() + days);
  return r;
}

function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function getOrCreateFolder(parent, name) {
  const folders = parent.getFoldersByName(String(name));
  return folders.hasNext() ? folders.next() : parent.createFolder(String(name));
}
