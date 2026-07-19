/**
 * Searches Gmail for IRCTC confirmations and appends to Train Booking History.
 */
const TrainIrctcImportModule = (function () {
  'use strict';

  const LOCK_TIMEOUT_MS = 30000;

  function getCustomGmailQuery_() {
    try {
      const key = AppConfig.getIrctcGmailQueryPropertyKey();
      const custom = PropertiesService.getScriptProperties().getProperty(key);
      if (custom && String(custom).trim()) {
        return String(custom).trim();
      }
    } catch (err) {
      console.error('TrainIrctcImportModule.getCustomGmailQuery_ failed', {
        error: err,
        stack: err && err.stack,
      });
    }
    return '';
  }

  function getForwardFromEmails_() {
    const configured = AppConfig.getIrctcForwardFromEmails();
    if (configured && configured.length) {
      return configured;
    }
    return [];
  }

  function buildGmailQuery_() {
    const clauses = [
      'from:' + AppConfig.getIrctcEmailFrom(),
      'subject:"Booking Confirmation on IRCTC"',
      'subject:"Ticket Confirmation"',
      'subject:"Fwd: Booking Confirmation"',
    ];

    const irctcSubjectHint =
      '(subject:"Booking Confirmation" OR subject:"IRCTC" OR subject:"Ticket Confirmation" OR subject:"PNR No.")';

    getForwardFromEmails_().forEach(function (email) {
      if (email) {
        clauses.push('(from:' + email + ' ' + irctcSubjectHint + ')');
      }
    });

    return 'is:unread (' + clauses.join(' OR ') + ')';
  }

  function getGmailQuery_() {
    const custom = getCustomGmailQuery_();
    if (custom) {
      return custom;
    }
    return buildGmailQuery_();
  }

  /**
   * Import unread IRCTC emails into Train Booking History.
   * @returns {{appended: number, skipped: number, query: string}}
   */
  function extractEmails_() {
    TrainIrctcParseModule.clearMatchPlaceCache();

    const sheet = TrainBookingHistoryModule.getBookingSheet();
    if (!sheet) {
      ErrorLogModule.error(
        'TrainIrctcImportModule',
        'extractEmails_',
        'Train Booking History sheet unavailable',
        { triggerSource: ErrorLogModule.getTriggerSource() || 'manual' }
      );
      return { appended: 0, skipped: 0, query: '', error: 'Sheet unavailable' };
    }

    const query = getGmailQuery_();
    console.log('TrainIrctcImportModule: Gmail query = ' + query);

    var threads = [];
    try {
      threads = GmailApp.search(query, 0, AppConfig.getIrctcMaxThreadsPerRun());
    } catch (err) {
      console.error('TrainIrctcImportModule.extractEmails_: Gmail search failed', {
        query: query,
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error('TrainIrctcImportModule', 'extractEmails_', 'Gmail search failed: ' + err, {
        details: query,
      });
      return { appended: 0, skipped: 0, query: query, error: String(err) };
    }

    if (!threads || threads.length === 0) {
      console.log('TrainIrctcImportModule: no unread emails matched query');
      return { appended: 0, skipped: 0, query: query };
    }

    var appended = 0;
    var skipped = 0;
    const markRead = AppConfig.getIrctcMarkReadAfterImport();

    threads.forEach(function (thread) {
      try {
        const threadId = thread.getId();
        const messages = thread.getMessages();

        messages.forEach(function (message) {
          try {
            if (!message.isUnread()) {
              return;
            }

            const body = TrainIrctcParseModule.getMessageBody(message);
            if (!body) {
              console.log('TrainIrctcImportModule: empty body, thread=' + threadId);
              skipped++;
              return;
            }

            const parsed = TrainIrctcParseModule.parseEmailBody(body, message);
            parsed.threadId = threadId;

            if (!parsed.pnr && !parsed.transactionId) {
              console.log('TrainIrctcImportModule: could not parse PNR/Transaction ID — skipped');
              skipped++;
              return;
            }

            if (TrainBookingHistoryModule.rowExists(sheet, parsed)) {
              console.log(
                'TrainIrctcImportModule: duplicate skipped (PNR=' +
                  parsed.pnr +
                  ', txn=' +
                  parsed.transactionId +
                  ')'
              );
              skipped++;
              if (markRead) {
                message.markRead();
              }
              return;
            }

            const ok = TrainBookingHistoryModule.appendParsedRow(sheet, parsed);
            if (ok) {
              appended++;
              console.log(
                'TrainIrctcImportModule: added PNR ' +
                  parsed.pnr +
                  ' ' +
                  parsed.from +
                  '→' +
                  parsed.to +
                  ' ' +
                  parsed.dateOfJourney
              );
            } else {
              skipped++;
            }

            if (markRead) {
              message.markRead();
            }
          } catch (msgErr) {
            skipped++;
            console.error('TrainIrctcImportModule: message failed', {
              error: msgErr,
              stack: msgErr && msgErr.stack,
            });
            ErrorLogModule.error(
              'TrainIrctcImportModule',
              'extractEmails_',
              'Message import failed: ' + msgErr,
              {}
            );
          }
        });
      } catch (threadErr) {
        skipped++;
        console.error('TrainIrctcImportModule: thread failed', {
          error: threadErr,
          stack: threadErr && threadErr.stack,
        });
      }
    });

    console.log(
      'TrainIrctcImportModule finished: appended=' + appended + ' skipped=' + skipped
    );

    if (sheet.getLastRow() >= AppConfig.getDataStartRow()) {
      TrainBookingHistoryModule.sortByJourneyDate(sheet);
    }

    return { appended: appended, skipped: skipped, query: query };
  }

  function extractEmailsWithLock_() {
    const lock = LockService.getScriptLock();
    const hasLock = lock.tryLock(LOCK_TIMEOUT_MS);
    if (!hasLock) {
      console.log('TrainIrctcImportModule: could not acquire lock');
      return { appended: 0, skipped: 0, query: '', error: 'Lock timeout' };
    }

    try {
      ErrorLogModule.setTriggerSource('scheduled');
      return extractEmails_();
    } finally {
      ErrorLogModule.clearTriggerSource();
      lock.releaseLock();
    }
  }

  return {
    extractEmails: extractEmails_,
    extractEmailsWithLock: extractEmailsWithLock_,
    getGmailQuery: getGmailQuery_,
    buildGmailQuery: buildGmailQuery_,
    LOCK_TIMEOUT_MS: LOCK_TIMEOUT_MS,
  };
})();
