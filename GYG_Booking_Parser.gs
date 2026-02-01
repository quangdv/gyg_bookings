/*************************************************
 * CONFIG
 *************************************************/
const ROOT_FOLDER_ID = '1bv4A4wUX757bTALr9FbIns90O7iRc3p1';
const SOURCE_LABEL = 'GYG_BOOKING';
const PROCESSED_LABEL = 'GYG_BOOKING_READ';
const TIMEZONE = 'Asia/Ho_Chi_Minh';

// Status (EN)
const STATUS_NEW = 'NEW';
const STATUS_READY = 'READY_TO_CONFIRM';
const STATUS_DRAFTED = 'CONFIRMATION_DRAFTED';

/*************************************************
 * MAIN – PARSE BOOKING EMAIL
 *************************************************/
function processGYGBookings() {
  const sourceLabel = getOrCreateLabel(SOURCE_LABEL);
  const processedLabel = getOrCreateLabel(PROCESSED_LABEL);

  const threads = sourceLabel.getThreads();
  Logger.log('Found threads: ' + threads.length);

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const msg = messages.pop();
    const booking = parseGYGBooking(msg);

    if (!booking) {
      Logger.log('Skip email: cannot parse booking');
      return;
    }

    // Create 1 daily sheet per original email received date (first message in thread, not latest)
    const firstMsg = messages.length > 0 ? messages[0] : msg;
    const receivedAt = firstMsg.getDate();
    const receivedYmd = Utilities.formatDate(receivedAt, TIMEZONE, 'yyyy-MM-dd');
    const [y, m, d] = receivedYmd.split('-').map(Number);
    const receivedDate = new Date(y, m - 1, d);

    const sheet = getOrCreateDailySheet(receivedDate);
    upsertBookingRowByReference(sheet, booking);

    thread.removeLabel(sourceLabel);
    thread.addLabel(processedLabel);
  });
}

/*************************************************
 * PARSER – GETYOURGUIDE EMAIL
 *************************************************/
function parseGYGBooking(message) {
  if (!message) return null;

  const html = message.getBody();
  const text = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  // TOUR (only first occurrence)
  const tourMatch = text.match(/Your offer has been booked:\s*(.*?)\s*Reference number/i);
  if (!tourMatch) return null;
  let tour = tourMatch[1].trim().replace(/&amp;/g, '&');
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

  // DATE
  const dateMatch = text.match(/Date\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})/);
  const checkinDate = dateMatch ? new Date(dateMatch[1]) : null;
  if (!(checkinDate instanceof Date) || isNaN(checkinDate)) return null;

  // CUSTOMER
  // Use \p{L} to match any Unicode letter (including ü, ö, é, ñ, etc.)
  const customerMatch = text.match(/Main customer\s*([\p{L}\s]+)/iu);
  let customer = customerMatch ? customerMatch[1].trim() : '';
  // Remove "customer" word if present (at beginning or end)
  customer = customer.replace(/^customer\s*|\s*customer$/gi, '').trim();

  // EMAIL
  const emailMatch = text.match(/([a-z0-9._%+-]+@reply\.getyourguide\.com)/i);
  const email = emailMatch ? emailMatch[1] : '';

  // PHONE
  const phoneMatch = text.match(/Phone:\s*([+\d\s]+)/i);
  const phone = phoneMatch ? phoneMatch[1].trim() : '';

  // ADULT
  const adultMatch = text.match(/(\d+)\s*x\s*Adults?/i);
  const adults = adultMatch ? Number(adultMatch[1]) : 0;

  // CHILDREN
  const childrenMatch = text.match(/(\d+)\s*x\s*Child(?:ren)?/i);
  const children = childrenMatch ? Number(childrenMatch[1]) : 0;

  // INFANT
  const infantMatch = text.match(/(\d+)\s*x\s*Infants?/i);
  const infant = infantMatch ? Number(infantMatch[1]) : 0;

  // PICKUP
  const pickupMatch = text.match(/Pickup\s*(.*?)\s*(Open in Google Maps|Price)/i);
  const pickup = pickupMatch ? pickupMatch[1].trim().replace(/&amp;/g, '&') : '';

  // REFERENCE
  const referenceMatch = text.match(/Reference number\s*:?\s*([A-Z0-9\-]+?)(?=Date|$)/i);
  const reference = referenceMatch ? referenceMatch[1].trim() : '';

  return {
    tour,
    customer,
    email,
    phone,
    checkinDate,
    checkoutDate: addDays(checkinDate, 1),
    adults,
    children,
    infant,
    pickup,
    pickupTime: '8:00 to 8:30 AM',
    reference
  };
}

/*************************************************
 * SHEET
 *************************************************/
function getOrCreateDailySheet(dateObj) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const year = dateObj.getFullYear();
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const dateStr = Utilities.formatDate(dateObj, TIMEZONE, 'yyyy-MM-dd');

  const yearFolder = getOrCreateFolder(root, year);
  const monthFolder = getOrCreateFolder(yearFolder, month);

  const fileName = `GYG_Bookings_${dateStr}`;
  const files = monthFolder.getFilesByName(fileName);

  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
    // Ensure formatting rules exist for already-created sheets
    applyStatusRowColorRules_(ss.getActiveSheet());
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
    'Peak season','Bus','Single Cabin','VAT','Holiday','Other','Cruise',
    'Pickup','Pickup time',
    'Status','Email','Phone','Reference'
  ];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS_NEW, STATUS_READY, STATUS_DRAFTED], true)
    .build();

  sheet.getRange('T2:T').setDataValidation(rule);

  applyStatusRowColorRules_(sheet);
}

function applyStatusRowColorRules_(sheet) {
  const dataRange = sheet.getRange('A2:W'); // 23 columns (A..W), exclude header row

  // Remove old rules for these exact formulas (avoid duplicates)
  const existing = sheet.getConditionalFormatRules();
  const kept = existing.filter(r => {
    const cond = r.getBooleanCondition && r.getBooleanCondition();
    if (!cond) return true;
    if (cond.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return true;
    const v = cond.getCriteriaValues();
    const f = v && v[0];
    return !(
      f === '=$T2="READY_TO_CONFIRM"' ||
      f === '=$T2="CONFIRMATION_DRAFTED"'
    );
  });

  const readyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$T2="READY_TO_CONFIRM"')
    .setBackground('#D9EAD3')
    .setRanges([dataRange])
    .build();

  const draftedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$T2="CONFIRMATION_DRAFTED"')
    .setBackground('#B6D7A8')
    .setRanges([dataRange])
    .build();

  sheet.setConditionalFormatRules([...kept, readyRule, draftedRule]);
}

/*************************************************
 * UPSERT ROW (BY REFERENCE)
 *************************************************/
function upsertBookingRowByReference(sheet, b) {
  // If reference is missing, fall back to append (can't dedupe reliably)
  const ref = String(b.reference || '').trim();
  if (!ref) {
    appendBookingRow(sheet, b);
    return;
  }

  const existingRowIndex = findRowIndexByReference_(sheet, ref);
  if (!existingRowIndex) {
    appendBookingRow(sheet, b);
    return;
  }

  // Update only fields sourced from email; keep manual fields + existing status
  const lastCol = 23; // headers length
  const current = sheet.getRange(existingRowIndex, 1, 1, lastCol).getValues()[0];

  const rooms = calculateRooms(b.adults + b.children);

  // 0-based mapping in current[]
  current[0] = b.tour;
  current[1] = b.customer;
  current[2] = b.checkinDate;
  current[3] = b.checkoutDate;
  current[4] = b.adults;
  current[5] = b.children;
  current[6] = b.infant;

  // Room counts (derived)
  current[7] = rooms.double;
  current[8] = rooms.triple;
  current[9] = rooms.single;

  // Pickup
  current[17] = b.pickup;
  current[18] = b.pickupTime;

  // Status: keep existing if present, else set NEW
  if (!current[19]) current[19] = STATUS_NEW;

  // Contact
  current[20] = b.email;
  current[21] = b.phone;

  // Reference
  current[22] = ref;

  sheet.getRange(existingRowIndex, 1, 1, lastCol).setValues([current]);
}

function appendBookingRow(sheet, b) {
  const rooms = calculateRooms(b.adults + b.children);

  // Build row data with explicit column mapping
  const rowData = [
    b.tour,           // Col A (0): Tour
    b.customer,       // Col B (1): Customer Name
    b.checkinDate,    // Col C (2): Checkin
    b.checkoutDate,   // Col D (3): Checkout
    b.adults,         // Col E (4): Adult
    b.children,       // Col F (5): Children
    b.infant,         // Col G (6): Infant
    rooms.double,     // Col H (7): Double/Twin
    rooms.triple,     // Col I (8): Triple
    rooms.single,     // Col J (9): Single
    '',               // Col K (10): Peak season
    '',               // Col L (11): Bus
    '',               // Col M (12): Single Cabin
    '',               // Col N (13): VAT
    '',               // Col O (14): Holiday
    '',               // Col P (15): Other
    '',               // Col Q (16): Cruise
    b.pickup,         // Col R (17): Pickup
    b.pickupTime,     // Col S (18): Pickup time
    STATUS_NEW,       // Col T (19): Status
    b.email,          // Col U (20): Email
    b.phone,          // Col V (21): Phone
    b.reference       // Col W (22): Reference
  ];
  sheet.appendRow(rowData);
}

function findRowIndexByReference_(sheet, reference) {
  const refCol = 23; // 1-based column index of 'Reference'
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, refCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === reference) {
      return i + 2; // actual sheet row index
    }
  }
  return null;
}

/*************************************************
 * BUILD EMAIL DRAFTS (MANUAL RUN)
 *************************************************/
function buildConfirmationDrafts() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);

  const years = root.getFolders();
  while (years.hasNext()) {
    const y = years.next();
    const months = y.getFolders();

    while (months.hasNext()) {
      const m = months.next();
      const files = m.getFilesByType(MimeType.GOOGLE_SHEETS);

      while (files.hasNext()) {
        const sheet = SpreadsheetApp.open(files.next()).getActiveSheet();
        const rows = sheet.getDataRange().getValues();

        for (let i = 1; i < rows.length; i++) {
          if (rows[i][19] === STATUS_READY && rows[i][20]) {
            const reference = String(rows[i][22] || '').trim(); // Reference column (index 22)
            const email = rows[i][20];
            const subject = `Booking Confirmation – ${rows[i][0]}`;
            const htmlBody = buildEmailHTML(rows[i]);

            // Try to find original email by reference and reply to it
            const originalMessage = findOriginalEmailByReference(reference);
            if (originalMessage) {
              try {
                // Create draft reply - this automatically sets up reply headers
                const draft = originalMessage.createDraftReply('', {
                  htmlBody: htmlBody
                });

                // Verify draft was created and log details
                const draftId = draft.getId();
                const threadId = originalMessage.getThread().getId();
                Logger.log(`✅ Created reply draft for reference: ${reference}`);
                Logger.log(`   Draft ID: ${draftId}, Thread ID: ${threadId}`);
                Logger.log(`   Original subject: ${originalMessage.getSubject()}`);

                // Force save by accessing draft properties (ensures it's saved)
                const draftSubject = draft.getMessage().getSubject();
                Logger.log(`   Draft subject: ${draftSubject}`);

              } catch (e) {
                Logger.log(`❌ Error creating reply draft for reference: ${reference}`);
                Logger.log(`   Error: ${e.toString()}`);
                // Fallback: create new draft if reply fails
                GmailApp.createDraft(
                  email,
                  subject,
                  '',
                  { htmlBody: htmlBody }
                );
                Logger.log(`⚠️ Fallback: created new draft for reference: ${reference}`);
              }
            } else {
              // Fallback: create new draft if original email not found
              GmailApp.createDraft(
                email,
                subject,
                '',
                { htmlBody: htmlBody }
              );
              Logger.log(`⚠️ Original email not found for reference: ${reference}, created new draft`);
            }

            sheet.getRange(i+1, 20).setValue(STATUS_DRAFTED);
          }
        }
      }
    }
  }
}

/*************************************************
 * FIND ORIGINAL EMAIL BY REFERENCE
 *************************************************/
function findOriginalEmailByReference(reference) {
  if (!reference) return null;

  // Limit search to processed booking emails for better performance
  const labelQuery = `label:${SOURCE_LABEL} OR label:${PROCESSED_LABEL}`;

  // First: Search in subject (faster and more reliable)
  // Reference can be at end: "Booking - S68147 - GYGVN3HW223V" or just "GYGVN3HW223V"
  const subjectQuery = `${labelQuery} subject:"${reference}"`;
  let threads = GmailApp.search(subjectQuery, 0, 20);

  for (let thread of threads) {
    const messages = thread.getMessages();
    // Get the first message in thread (original email)
    if (messages.length > 0) {
      const msg = messages[0];
      const subject = msg.getSubject();
      // Check if reference appears in subject
      if (subject.includes(reference)) {
        return msg;
      }
    }
  }

  // Second: If not found in subject, search in body content
  // Look for "Reference number GYGVN3HW223V" or "Reference number: GYGVN3HW223V"
  const bodyQuery = `${labelQuery} "Reference number ${reference}"`;
  threads = GmailApp.search(bodyQuery, 0, 20);

  for (let thread of threads) {
    const messages = thread.getMessages();
    // Get the first message in thread (original email)
    if (messages.length > 0) {
      const msg = messages[0];
      const body = msg.getBody();
      const text = body
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/p>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/\s+/g, ' ')
        .trim();

      // Verify reference appears in content
      const refPattern = new RegExp(`Reference number\\s*:?\\s*${reference.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`, 'i');
      if (refPattern.test(text)) {
        return msg;
      }
    }
  }

  return null;
}

/*************************************************
 * EMAIL HTML (TABLE – GMAIL SAFE)
 *************************************************/
function buildEmailHTML(r) {
  /*
    Column mapping (theo sheet của bạn):
    0  Tour
    1  Customer name
    2  Checkin
    3  Checkout
    4  Adult
    5  Children
    6  Infant
    7  Double/Twin
    8  Triple
    9  Single
    10 Peak season
    11 Bus
    12 Single Cabin
    13 VAT
    14 Holiday
    15 Other
    16 Cruise
    17 Pickup
    18 Pickup time
    19 Status
    20 Email
    21 Phone
  */

  const fmtDate = d =>
    d instanceof Date
      ? Utilities.formatDate(d, TIMEZONE, 'dd-MMM-yy')
      : d;

  const tourName = String(r[0] || '');
  const roomSuffix = tourName.toLowerCase().indexOf('bai tu long') !== -1
    ? 'Delux Window Cabin'
    : 'Private Balcony';
  const rooms = [];
  if (r[7]) rooms.push(`${r[7]} Double/ Twin - ${roomSuffix}`);
  if (r[8]) rooms.push(`${r[8]} Triple - ${roomSuffix}`);
  if (r[9]) rooms.push(`${r[9]} Single - ${roomSuffix}`);

  const surcharges = [
    ['Peak season from 1 Oct to 30 Apr $13/person', r[10]],
    ['Single Cabin 80 USD', r[12]],
    ['Limousine Bus 2 way HN <--> HL $25/person', r[11]],
    ['The Government VAT Tax 12 USD/person', r[13]],
    ['Holiday', r[14]],
    ['Other', r[15]],
  ].filter(x => Number(x[1]) > 0);

  const total = surcharges.reduce((s, x) => s + Number(x[1]), 0);
  const sig = getDefaultGmailSignature();

  return `
<table border="0" cellpadding="0" cellspacing="0" width="725"
 style="border-collapse:collapse;font-family:'Times New Roman';font-size:11pt;color:#000">

<tr>
  <td colspan="3" style="padding-bottom:12px">
    Dear ${r[1]},<br><br>
    Thank you for booking with us.<br>
    I would like to confirm your booking as follows:
  </td>
</tr>

${row('Tour code', r[0])}
${row('Guest name', r[1])}
${row('Number of guest', `${r[4]} x Adults`)}
${row('Check-in date', fmtDate(r[2]))}
${row('Check out date', fmtDate(r[3]))}
${row('Room', rooms.join(' / '))}
${row('Pick up/Drop off address', r[17] || 'Please provide')}
${row('Pick up time', r[18])}

<tr>
  <td rowspan="${surcharges.length + 1}"
      style="border:1px solid #000;text-align:center;vertical-align:middle">
    Surcharge (USD)
  </td>
  <td style="border:1px solid #000">${surcharges[0][0]}</td>
  <td style="border:1px solid #000;text-align:right">$ ${surcharges[0][1]}</td>
</tr>

${surcharges.slice(1).map(x => `
<tr>
  <td style="border:1px solid #000">${x[0]}</td>
  <td style="border:1px solid #000;text-align:right">$ ${x[1]}</td>
</tr>`).join('')}

<tr>
  <td style="border:1px solid #000;font-weight:bold">Total</td>
  <td style="border:1px solid #000;text-align:right;font-weight:bold">$ ${total}</td>
</tr>

<tr><td colspan="3" style="height:15px"></td></tr>

<tr>
<td colspan="3">
${getNoteContent()}
</td>
</tr>

</table>
${sig ? '<div style="margin-top:15px">' + sig + '</div>' : ''}
`;
}

function row(label, value) {
  return `
<tr>
  <td style="border:1px solid #000;width:180px">${label}</td>
  <td colspan="2" style="border:1px solid #000;text-align:center">${value || ''}</td>
</tr>`;
}

/*************************************************
 * GET DEFAULT GMAIL SIGNATURE
 *************************************************/
function getDefaultGmailSignature() {
  try {
    const list = Gmail.Users.Settings.SendAs.list('me');
    if (!list.sendAs || !list.sendAs.length) return '';
    const defaultSendAs = list.sendAs.find(function (account) {
      return account.isDefault;
    });
    return (defaultSendAs && defaultSendAs.signature) ? defaultSendAs.signature : '';
  } catch (e) {
    Logger.log('Could not get default Gmail signature: ' + e.toString());
    return '';
  }
}

/*************************************************
 * GET NOTE CONTENT FROM FILE
 *************************************************/
function getNoteContent() {
  try {
    // Try to read from HTML template file (note.html)
    // In Google Apps Script, create a file named "note.html" in the same project
    const noteContent = HtmlService.createTemplateFromFile('note').getRawContent();
    if (noteContent && noteContent.trim()) {
      return noteContent;
    }
  } catch (e) {
    Logger.log('Note file not found, using default note content. Error: ' + e.toString());
  }

  // Fallback to default note content if file not found
  // OLD NOTE CONTENT (commented out - now reading from note.html file):
  /*
  <b>Note:</b><br>
  - About the surcharge, cash is recommended. Card payments will incur a 3–10% bank commission fee.<br>
  - Do you have any food allergies or are you vegetarian?<br>
  - Please provide details of passport information of all guests before check-in cruise.<br>
  - Pick up & drop off point can be Hanoi Old Quarter or Ninh Binh.<br>
  - Estimated pickup time: 8:00–8:30 AM (Hanoi), 7:00–7:15 AM (Ninh Binh).<br>
  - Please reply to this email to confirm you received the information.<br>
  - Please give us your Whatsapp number or personal email so we can contact you easily.
  */

  // Return default content as fallback
  return `<b>Note:</b><br>
- About the surcharge, cash is recommended. Card payments will incur a 3–10% bank commission fee.<br>
- Do you have any food allergies or are you vegetarian?<br>
- Please provide details of passport information of all guests before check-in cruise.<br>
- Pick up & drop off point can be Hanoi Old Quarter or Ninh Binh.<br>
- Estimated pickup time: 8:00–8:30 AM (Hanoi), 7:00–7:15 AM (Ninh Binh).<br>
- Please reply to this email to confirm you received the information.<br>
- Please give us your Whatsapp number or personal email so we can contact you easily.`;
}


/*************************************************
 * HELPERS
 *************************************************/
function calculateRooms(x) {
  if (x <= 0) return { double:0, triple:0, single:0 };
  if (x === 1) return { double:0, triple:0, single:1 };
  if (x === 2) return { double:1, triple:0, single:0 };
  if (x === 3) return { double:0, triple:1, single:0 };
  if (x % 2 === 0) return { double:x/2, triple:0, single:0 };
  return { double:(x-3)/2, triple:1, single:0 };
}

function addDays(d, n) {
  const r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}

function formatDate(d) {
  return Utilities.formatDate(d, TIMEZONE, 'dd-MMM-yyyy');
}

function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function getOrCreateFolder(parent, name) {
  const f = parent.getFoldersByName(String(name));
  return f.hasNext() ? f.next() : parent.createFolder(String(name));
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Booking')
    .addItem('Create confirmation drafts', 'buildConfirmationDrafts')
    .addToUi();
}


function createConfirmationDraftFromRow(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const colIndex = name => headers.indexOf(name);

  const email = data[colIndex('Email')];
  const statusIndex = colIndex('Status');

  if (!email) {
    Logger.log('❌ No email, skip row ' + row);
    return;
  }

  const subject = `Booking Confirmation – ${data[colIndex('Tour')]}`;
  const htmlBody = buildEmailHTML(data);

  GmailApp.createDraft(email, subject, '', {
    htmlBody: htmlBody
  });

  // Update status sau khi tạo draft
  sheet.getRange(row, statusIndex + 1).setValue('DRAFT_CREATED');

  Logger.log(`✅ Draft created for row ${row}`);
}

