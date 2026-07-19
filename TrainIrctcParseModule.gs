/**
 * Parses IRCTC Ticket Confirmation emails (2025+ tab-separated format).
 */
const TrainIrctcParseModule = (function () {
  'use strict';

  var matchPlaceLabelsCache_ = null;

  function clearMatchPlaceCache_() {
    matchPlaceLabelsCache_ = null;
  }

  function normalizeText_(text) {
    return String(text || '')
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/[ \t]+\n/g, '\n')
      .trim();
  }

  function cleanValue_(value) {
    return String(value || '')
      .replace(/\*/g, '')
      .replace(/\s+/g, ' ')
      .replace(/\|/g, '')
      .trim();
  }

  function cleanStationLabel_(stationRaw) {
    return cleanValue_(stationRaw);
  }

  function cleanMobile_(value) {
    const onlyDigits = String(value || '').replace(/\D/g, '');
    if (onlyDigits.length >= 10) {
      return onlyDigits.slice(-10);
    }
    return cleanValue_(value);
  }

  function looksLikeEmailAddress_(value) {
    const v = String(value || '').trim();
    return /@/.test(v) || /^<[^>]+>$/.test(v);
  }

  function stripHtmlToText_(html) {
    return String(html)
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/tr>/gi, '\n')
      .replace(/<\/td>/gi, '\t')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&');
  }

  function scoreBody_(text) {
    if (!text || text.length < 80) {
      return -1;
    }
    var s = 0;
    if (/\t/.test(text)) {
      s += 10;
    }
    if (/PNR\s+No\.?\s*:/i.test(text)) {
      s += 5;
    }
    if (/From\s+:\s*.+?\s+Date\s+of\s+Journey\s*:/i.test(text)) {
      s += 8;
    }
    if (/Passenger\s+Details/i.test(text)) {
      s += 3;
    }
    return s;
  }

  function pickBetterBody_(plain, html) {
    const plainScore = scoreBody_(plain);
    const htmlScore = scoreBody_(html);
    if (htmlScore > plainScore) {
      return html;
    }
    if (plainScore > htmlScore) {
      return plain;
    }
    if (html.length > plain.length) {
      return html;
    }
    return plain || html;
  }

  function getMessageBody_(message) {
    if (!message) {
      return '';
    }
    const plain = normalizeText_(message.getPlainBody() || '');
    const htmlRaw = message.getBody() || '';
    const html = htmlRaw ? normalizeText_(stripHtmlToText_(htmlRaw)) : '';
    return pickBetterBody_(plain, html);
  }

  function sliceBookingBlock_(text) {
    const startMatch = text.match(/(?:Ticket\s+Confirmation|PNR\s+No\.?\s*:)/i);
    const start = startMatch ? startMatch.index : 0;
    const tail = text.substring(start);
    const endMatch = tail.match(/\n\s*Fare\s+Details\b/i);
    return endMatch ? tail.substring(0, endMatch.index) : tail;
  }

  function extractDateOnly_(value) {
    const m = String(value || '').match(/\d{1,2}-[A-Za-z]{3}-\d{4}/i);
    return m ? m[0] : cleanValue_(value);
  }

  function extractDateTimeValue_(value) {
    const m = String(value || '').match(/\d{1,2}-[A-Za-z]{3}-\d{4}(?:\s+\d{1,2}:\d{2})?/i);
    return m ? m[0] : cleanValue_(value);
  }

  function extractLabelValue_(text, label) {
    const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(escaped + '\\*?\\s+:', 'gi');
    var match;
    while ((match = re.exec(text)) !== null) {
      var start = match.index + match[0].length;
      while (start < text.length && /[\s\t]/.test(text[start]) && text[start] !== '\n') {
        start++;
      }

      const rest = text.substring(start);
      const nextTabField = rest.search(/\t[A-Za-z][^\n\t]{0,60}?\s+:/);
      const nextInlineField = rest.search(
        /\s+(?:Date\s+of\s+Journey|To|Boarding\s+At|Reservation\s+Up|Scheduled\s+Departure|Passenger\s+Mobile|Distance|Quota|Class|Transaction\s+ID|Train\s+No)\s*:/i
      );
      const nl = rest.indexOf('\n');
      var end = text.length;
      if (nextTabField >= 0) {
        end = Math.min(end, start + nextTabField);
      }
      if (nextInlineField >= 0) {
        end = Math.min(end, start + nextInlineField);
      }
      if (nl >= 0) {
        end = Math.min(end, start + nl);
      }

      const value = cleanValue_(text.substring(start, end));
      if (value && !looksLikeEmailAddress_(value)) {
        return value;
      }
    }
    return '';
  }

  function extractRouteFields_(text) {
    const fromM = text.match(/From\s+:\s*(.+?)(?=\s+Date\s+of\s+Journey\s*:)/i);
    const dateM = text.match(/Date\s+of\s+Journey\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{4})/i);
    const toM = text.match(/To\s+:\s*(.+?)(?=\s+Boarding\s+At\s*:|(?:\t|\n)|\s+Reservation\s+Up)/i);
    return {
      from: fromM ? cleanStationLabel_(fromM[1]) : '',
      to: toM ? cleanStationLabel_(toM[1]) : '',
      dateOfJourney: dateM ? dateM[1] : '',
    };
  }

  function extractScheduledDeparture_(booking) {
    const m = booking.match(
      /Scheduled\s+Departure\*?\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{4}\s+\d{1,2}:\d{2})/i
    );
    if (m) {
      return m[1];
    }
    return extractDateTimeValue_(extractLabelValue_(booking, 'Scheduled Departure'));
  }

  function extractPassengerMobile_(booking) {
    const m = booking.match(/Passenger\s+Mobile\s+No\s*:\s*(\d{10,12})/i);
    return m ? m[1] : extractLabelValue_(booking, 'Passenger Mobile No');
  }

  function extractUserId_(text) {
    const m = text.match(/User\s*Id\s*:\s*([A-Za-z0-9._-]+)/i);
    return m ? m[1].trim() : '';
  }

  function extractDearName_(text) {
    const m = text.match(/Dear\s+([^(,\n]+?)\s*\(\s*User\s*Id\s*:/i);
    if (!m) {
      return '';
    }
    return cleanValue_(m[1]);
  }

  function extractTrainNumber_(trainRaw) {
    if (!trainRaw) {
      return '';
    }
    const num = String(trainRaw).match(/(\d{1,5})/);
    return num ? num[1] : cleanValue_(trainRaw);
  }

  function isPassengerStatus_(value) {
    const s = String(value || '').trim().toUpperCase();
    if (!s) {
      return false;
    }
    if (s === 'CNF' || s === 'CONFIRMED' || s === 'RAC') {
      return true;
    }
    if (s === 'WL' || /^WL\d+$/.test(s)) {
      return true;
    }
    return false;
  }

  function mapPassengerStatus_(statusRaw) {
    const s = String(statusRaw || '').trim().toUpperCase();
    if (s === 'CNF' || s === 'CONFIRMED') {
      return 'CONFIRMED';
    }
    if (s === 'RAC') {
      return 'RAC';
    }
    if (s.indexOf('WL') === 0 || s === 'WL') {
      return 'WL';
    }
    return s || 'BOOKED';
  }

  function inferTicketStatus_(text, seatBerth, coach) {
    const blob = (text + ' ' + seatBerth + ' ' + coach).toUpperCase();
    if (/\bWL\s*\d+\b/.test(blob) || /\bWAITLIST\b/.test(blob)) {
      return 'WL';
    }
    if (/\bRAC\b/.test(blob)) {
      return 'RAC';
    }
    if (/\bCNF\b/.test(blob) || /\bCONFIRM/.test(blob)) {
      return 'CONFIRMED';
    }
    if (seatBerth && /\d/.test(seatBerth)) {
      return 'CONFIRMED';
    }
    return 'BOOKED';
  }

  function buildFirstPassengerRow_(name, status, coach, seatBerth) {
    return {
      name: cleanValue_(name),
      status: status,
      coach: cleanValue_(coach),
      seatBerth: cleanValue_(seatBerth),
    };
  }

  function parsePassengerDetailsTable_(text) {
    const sectionMatch = text.match(/Passenger\s+Details[\s\S]*?(?=Fare\s+Details|$)/i);
    if (!sectionMatch) {
      return null;
    }

    const section = sectionMatch[0];
    const rowPatterns = [
      /^\s*1(?:\t+|\s{2,})([^\t\n]+?)(?:\t+|\s{2,})(\d+)(?:\t+|\s{2,})(Male|Female)(?:\t+|\s{2,})([^\t\n]+?)(?:\t+|\s{2,})(CNF|RAC|WL\d*)(?:\t+|\s{2,})(\S+?)(?:\t+|\s+)(\S+)/im,
      /^\s*1[\t ]+([A-Za-z][A-Za-z\s]+?)[\t ]+(\d+)[\t ]+(Male|Female)[\t ]+[^\t\n]+[\t ]+(CNF|RAC|WL\d*)[\t ]+(\S+)[\t ]+(\S+)/im,
    ];

    for (var p = 0; p < rowPatterns.length; p++) {
      const rowMatch = section.match(rowPatterns[p]);
      if (!rowMatch) {
        continue;
      }
      const statusIdx = p === 0 ? 5 : 4;
      const coachIdx = p === 0 ? 6 : 5;
      const seatIdx = p === 0 ? 7 : 6;
      if (!isPassengerStatus_(rowMatch[statusIdx])) {
        continue;
      }
      return buildFirstPassengerRow_(
        rowMatch[1],
        rowMatch[statusIdx],
        rowMatch[coachIdx],
        rowMatch[seatIdx]
      );
    }

    const lines = section.split('\n');
    for (var i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (!/\t/.test(line)) {
        continue;
      }
      const cols = line.split('\t').map(function (c) {
        return c.trim();
      });
      if (cols.length < 8) {
        continue;
      }
      if (cols[0] !== '1') {
        continue;
      }
      if (!isPassengerStatus_(cols[5])) {
        continue;
      }
      return buildFirstPassengerRow_(cols[1], cols[5], cols[6], cols[7]);
    }
    return null;
  }

  function stationCodeFromLabel_(label) {
    const m = String(label || '').match(/\(\s*([A-Za-z0-9]+)\s*\)\s*$/);
    return m ? m[1].toUpperCase() : '';
  }

  function getMatchPlaceLabels_() {
    if (matchPlaceLabelsCache_) {
      return matchPlaceLabelsCache_;
    }

    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      matchPlaceLabelsCache_ = [];
      return matchPlaceLabelsCache_;
    }

    const sheet = SpreadsheetUtils.getSheetByName(
      spreadsheet,
      AppConfig.getTrainMatchPlaceSheetName()
    );
    if (!sheet) {
      matchPlaceLabelsCache_ = [];
      return matchPlaceLabelsCache_;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      matchPlaceLabelsCache_ = [];
      return matchPlaceLabelsCache_;
    }

    try {
      const values = sheet.getRange(1, 1, lastRow, 1).getValues();
      matchPlaceLabelsCache_ = values
        .map(function (row) {
          return String(row[0] || '').trim();
        })
        .filter(function (label) {
          return label.length > 0;
        });
    } catch (err) {
      console.error('TrainIrctcParseModule.getMatchPlaceLabels_ failed', {
        error: err,
        stack: err && err.stack,
      });
      matchPlaceLabelsCache_ = [];
    }
    return matchPlaceLabelsCache_;
  }

  function resolveMatchPlaceLabel_(rawStation) {
    const raw = cleanStationLabel_(rawStation);
    if (!raw) {
      return '';
    }
    const labels = getMatchPlaceLabels_();
    if (labels.length === 0) {
      return raw;
    }

    const code = stationCodeFromLabel_(raw);
    if (code) {
      for (var i = 0; i < labels.length; i++) {
        if (stationCodeFromLabel_(labels[i]) === code) {
          return labels[i];
        }
      }
    }

    const rawName = raw.replace(/\s*\([^)]*\)\s*$/i, '').trim().toUpperCase();
    for (var j = 0; j < labels.length; j++) {
      const labelName = labels[j].replace(/\s*\([^)]*\)\s*$/i, '').trim().toUpperCase();
      if (
        labelName &&
        (rawName === labelName ||
          rawName.indexOf(labelName) >= 0 ||
          labelName.indexOf(rawName) >= 0)
      ) {
        return labels[j];
      }
    }
    return raw;
  }

  function getTimezone_() {
    try {
      return AppConfig.getTrainDateTimezone() || Session.getScriptTimeZone() || 'Asia/Kolkata';
    } catch (err) {
      return 'Asia/Kolkata';
    }
  }

  /**
   * @param {string} body
   * @param {{getDate: function(): Date}|GoogleAppsScript.Gmail.GmailMessage} message
   * @returns {Object}
   */
  function parseEmailBody_(body, message) {
    const text = normalizeText_(body);
    const booking = sliceBookingBlock_(text);
    const passenger = parsePassengerDetailsTable_(booking);
    const route = extractRouteFields_(booking);

    const pnr = extractLabelValue_(booking, 'PNR No.');
    const trainRaw = extractLabelValue_(booking, 'Train No. / Name');
    const transactionId = extractLabelValue_(booking, 'Transaction ID');

    var bookingDateTime = extractLabelValue_(booking, 'Date & Time of Booking');
    bookingDateTime = bookingDateTime.replace(/\s+HRS\s*$/i, '').trim();

    const fromRaw = route.from || extractLabelValue_(booking, 'From');
    const toRaw = route.to || extractLabelValue_(booking, 'To');
    const dateOfJourney =
      route.dateOfJourney || extractDateOnly_(extractLabelValue_(booking, 'Date of Journey'));
    const travelClass = extractLabelValue_(booking, 'Class');
    const dateOfBoarding = extractDateOnly_(extractLabelValue_(booking, 'Date Of Boarding'));
    const scheduledDeparture = extractScheduledDeparture_(booking);
    const passengerMobile = extractPassengerMobile_(booking);
    const userId = extractUserId_(text);

    const name = passenger ? passenger.name : extractDearName_(text);
    const status = passenger
      ? mapPassengerStatus_(passenger.status)
      : inferTicketStatus_(booking, passenger ? passenger.seatBerth : '', passenger ? passenger.coach : '');
    const coach = passenger ? passenger.coach : '';
    const seatBerth = passenger ? passenger.seatBerth : '';

    const messageDate = message && message.getDate ? message.getDate() : new Date();
    const bookingFallback = Utilities.formatDate(messageDate, getTimezone_(), 'dd-MMM-yyyy HH:mm');

    var txnDigits = String(transactionId || '').replace(/\D/g, '');
    if (txnDigits.length > 18) {
      txnDigits = txnDigits.slice(0, 18);
    }

    return {
      from: resolveMatchPlaceLabel_(cleanStationLabel_(fromRaw)),
      to: resolveMatchPlaceLabel_(cleanStationLabel_(toRaw)),
      dateOfJourney: dateOfJourney,
      name: name,
      status: status,
      coach: coach,
      seatBerth: seatBerth,
      travelClass: travelClass,
      pnr: String(pnr || '').replace(/\D/g, '').slice(0, 10),
      trainNo: extractTrainNumber_(trainRaw),
      scheduledDeparture: scheduledDeparture,
      dateOfBoarding: dateOfBoarding,
      transactionId: txnDigits,
      bookingDateTime: extractDateTimeValue_(bookingDateTime) || bookingFallback,
      userId: userId,
      passengerMobile: cleanMobile_(passengerMobile),
      threadId: '',
    };
  }

  return {
    clearMatchPlaceCache: clearMatchPlaceCache_,
    getMessageBody: getMessageBody_,
    parseEmailBody: parseEmailBody_,
    resolveMatchPlaceLabel: resolveMatchPlaceLabel_,
  };
})();
