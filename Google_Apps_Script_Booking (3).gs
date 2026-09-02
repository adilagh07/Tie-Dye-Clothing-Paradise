/**
 * Tie & Dye Clothing Paradise - BOOKING BACKEND
 *
 * 1. Create a Google Sheet.
 * 2. Open Extensions > Apps Script.
 * 3. Replace the placeholder code with this file.
 * 4. Set OWNER_EMAIL below.
 * 5. Run setupBookingSheet() once and authorize it.
 * 6. Deploy > New deployment > Web app:
 *      Execute as: Me
 *      Who has access: Anyone
 * 7. Copy the /exec URL into BOOKING_API_URL in index_booking_ready.html.
 */

const OWNER_EMAIL = "kurivitu07@gmail.com";
const SHEET_NAME = "Bookings";
const TIME_ZONE = Session.getScriptTimeZone() || "Pacific/Fiji";

function setupBookingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headers = [
    "Booking ID",
    "Submitted At",
    "Reservation Date",
    "Time",
    "Customer Name",
    "Email",
    "Phone",
    "Guests",
    "Message",
    "Status"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
  }

  return "Booking sheet is ready.";
}

function doGet() {
  return jsonResponse({
    success: true,
    message: "Booking API is running."
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No booking data was received.");
    }

    const data = JSON.parse(e.postData.contents);

    validateBooking(data);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      setupBookingSheet();
      sheet = ss.getSheetByName(SHEET_NAME);
    }

    const bookingId =
      "BK-" +
      Utilities.formatDate(new Date(), TIME_ZONE, "yyyyMMdd-HHmmss");

    const submittedAt = new Date();

    sheet.appendRow([
      bookingId,
      submittedAt,
      data.date,
      data.time,
      data.name,
      data.email,
      data.phone,
      Number(data.guests),
      data.message || "",
      "New"
    ]);

    sendOwnerEmail(data, bookingId);
    sendCustomerEmail(data, bookingId);

    return jsonResponse({
      success: true,
      bookingId: bookingId,
      message: "Booking received successfully."
    });

  } catch (error) {
    console.error(error);

    return jsonResponse({
      success: false,
      message: error.message || "Booking could not be processed."
    });
  }
}

function validateBooking(data) {
  const required = ["name", "email", "phone", "guests", "date", "time"];

  required.forEach(function (field) {
    if (!data[field] || String(data[field]).trim() === "") {
      throw new Error("Missing required field: " + field);
    }
  });

  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailPattern.test(String(data.email).trim())) {
    throw new Error("Please provide a valid email address.");
  }

  const guests = Number(data.guests);
  if (!Number.isInteger(guests) || guests < 1 || guests > 12) {
    throw new Error("Number of guests must be between 1 and 12.");
  }

  // Accept YYYY-MM-DD from the HTML date input.
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(data.date))) {
    throw new Error("Invalid reservation date.");
  }
}

function sendOwnerEmail(data, bookingId) {
  if (!OWNER_EMAIL || OWNER_EMAIL.indexOf("YOUR_CLIENT_EMAIL") === 0) {
    throw new Error("OWNER_EMAIL has not been configured in Apps Script.");
  }

  const subject = "New Table Reservation - " + bookingId;

  const body =
    "A new table reservation has been received.\n\n" +
    "Booking ID: " + bookingId + "\n" +
    "Reservation Date: " + data.date + "\n" +
    "Time: " + data.time + "\n\n" +
    "Customer Information\n" +
    "--------------------\n" +
    "Name: " + data.name + "\n" +
    "Email: " + data.email + "\n" +
    "Phone: " + data.phone + "\n" +
    "Guests: " + data.guests + "\n" +
    "Message: " + (data.message || "None") + "\n\n" +
    "The booking has also been added to the Google Sheet.";

  MailApp.sendEmail({
    to: OWNER_EMAIL,
    subject: subject,
    body: body,
    replyTo: data.email,
    name: "Tie & Dye Clothing Paradise Booking"
  });
}

function sendCustomerEmail(data, bookingId) {
  const subject = "Reservation Request Received - " + bookingId;

  const body =
    "Hello " + data.name + ",\n\n" +
    "Thank you for your reservation request. We have received your booking details.\n\n" +
    "Booking ID: " + bookingId + "\n" +
    "Date: " + data.date + "\n" +
    "Time: " + data.time + "\n" +
    "Guests: " + data.guests + "\n" +
    "Message: " + (data.message || "None") + "\n\n" +
    "We will contact you if any further confirmation is required.\n\n" +
    "Thank you,\n" +
    "Tie & Dye Clothing Paradise";

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    body: body,
    name: "Tie & Dye Clothing Paradise"
  });
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
