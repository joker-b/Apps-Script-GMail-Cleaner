//
// Email retention rules for base-account (free) gmail.
//
// CAUTION: THIS SCRIPT DELETES EMAIL MESSAGES FROM YOUR ACCOUNT. *THAT'S WHAT IT'S FOR.*
//   The author bears no responsibility for email that you may erroneously or accidentally
//   or otherwise delete if you run this script.
//
// GMail will only let you delete around 300 emails per run before it times out,
//   so the best way to run this is to put it on a timer and let it slowly whittle-away
//   at your existing emails until it catches up with the day-to-day. Do this via the
//  "Resources->Current project's triggers..." menu item, and select "retentionRulesMain"
//  as the function. Every few hours or once per day should be plenty for most people...
//  you may want to run it more often just at first during the initial "catch-up phase."
//
// This script uses two Google Drive documents: a spreadsheet containing rules for
//   email retention, and a simple doc for storing logs of actions the script has taken.
//   You will need to create these two drive docs and pass their keys into the script. This script
//   can be attached to the control spreadsheet, in which case the spreedsheet FileID (aka ControlID)
//   is ignored.
//
// To define retention rules, create the control spreadsheet with these column headers:
//  Label  NumDays ProtectUnread ProtectRead ProtectStarred ProtectImportant daysArchive
// and then fill in the rows appropriately. The boolean values for each row can
// be blank (don't protect) or 1 (do protect).
//
// if both ProtectUnread and ProtectRead are true, this label goes into a special
//    "always protect" category
//
// Enter the Google key (part of the url) in the SpreadsheetApp.openById(id) function found below.
//
// In addition, create a simple Google Document, called "Retention Log" or similar,
//  for collecting logs of the script behavior. Enter the key for that document in
//  the DocumentApp.openById(id) function found below
//

/* globals SpreadsheetApp */
/* globals DocumentApp */
/* globals GmailApp */
/* globals Logger */

// Set these as you please -- these apply to labelled cels that are otherwise empty.
//  Labels that are NOT listed get NO special handling
var DEFAULT = {
  DAYS: 365,
  MINIMUM_DAYS: 7,
  PROTECT_UNREAD: false,
  PROTECT_READ: false,
  PROTECT_STARRED: true,
  PROTECT_IMPORTANT: true,
  MAX_THREADS: 500,
};

var _gDebug = false;

// never delete messages with these labels....
var Immortals = {};

// doc keys for the control spreadsheet and the text logfile
var CONTROL_ID = 'the_url_key_for_the_control_spreadsheet_here'; // optional -- we may be attached to the current doc
var LOG_DOC_ID = 'the_url_key_for_the_retention_log_file__here'; // also optional?

//
//
//
function _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived,nForever,pretrash) {
  'use strict';
  if ((nRemoved+nArchived+pretrash) > 0) {
    var msg = ('  Removed ' + nRemoved + '/' + pretrash + ' items ' + daysBack + '+ days old');
    msg = (msg + ', archived ' + nArchived + ' (' + daysArchive + '+ days), '+nForever+' immortals');
    LogPar.appendText(msg);
  }
}

//
// find out if this thread is among the immortals
//
function _hasImmortality_(MsgThread) {
  'use strict';
  var labs = MsgThread.getLabels();
  if (labs.length < 1) {
    return false;
  }
  for (var ml in labs) {
    var n = labs[ml].getName();
    if (n in Immortals) {
      if (_gDebug) {
        Logger.log('Msg "'+MsgThread.getFirstMessageSubject()+'" is immortal: '+n);
      }
      return true;
    }
  }
  return false;
}

//
// Apply retention rules to a single label
//
function _applyRetention_(labelName,daysBack,ProtectUnread,ProtectRead,ProtectStarred,ProtectImportant,daysArchive,LogPar) {
  'use strict';
  //var logMsg = ('applyRetention("'+labelName+'",'+daysBack+','+ProtectUnread+','+ProtectRead+','+ProtectStarred+','+ProtectImportant+','+daysArchive+',...)');
  //Logger.log(logMsg);
  var nRemoved = 0;
  var nArchived = 0;
  var nForever = 0;
  var operations = 0;
  var pretrash = 0;
  var maxThreads = DEFAULT.MAX_THREADS;
  var label;
  try {
    label = GmailApp.getUserLabelByName(labelName);
  } catch(err) {
    Logger.log('Too many mail calls for "'+labelName+'"? "'+err.message+'"');
    LogPar.appendText(' read error');
    return -1; // error
  }
  if (label === undefined) {
    Logger.log('Could not find label "'+labelName+'"');
    LogPar.appendText('No such label available');
    LogPar.setForegroundColor('#800000');
    LogPar.setBold(true);
    return nRemoved;
  }
  var now = new Date();
  var millisecsPerDay = 1000 * 60 * 60 * 24;
  var nowDay = Math.round(now.getTime() / millisecsPerDay);
  var cutoff = nowDay - daysBack;
  var cutoffArch = nowDay - daysArchive;
  if (daysArchive <= 0) { // hack to avoid archiving 
    cutoffArch = cutoff - 1; // one day before the cutoff
  }
  var startThread = 0;
  var pending = true;
  var i, msgThread, lastMsgDay, unrd;
  while (pending) {
    var threads = label.getThreads(startThread,maxThreads); // seems to max at 500 - TO-DO: use getThreads(start,max) to fetch more past 500
    var nThreads = threads.length;
    if (nThreads === maxThreads) {
      startThread = startThread + maxThreads; // for the next cycle
      LogPar.appendText('>');
    } else {
      LogPar.appendText((nThreads+startThread)+' initial items');
      pending = false;
    }
    for (i = 0; i < threads.length; i += 1) {
      msgThread = threads[i];
      if (msgThread.isInTrash()) {
        pretrash += 1;
        continue;
      }
      if (ProtectStarred && msgThread.hasStarredMessages()) {
        continue;
      }
      if (ProtectImportant && msgThread.isImportant()) {
        continue;
      }
      unrd = msgThread.isUnread();
      if ((ProtectUnread && unrd) || (ProtectRead && (!unrd))) {
        continue;
      }
      if (_hasImmortality_(msgThread)) {
        nForever += 1;
        continue;
      }
      lastMsgDay = Math.round(msgThread.getLastMessageDate().getTime()/millisecsPerDay);
      if (lastMsgDay > cutoff) {  // old enough?
        if (lastMsgDay <= cutoffArch) { // within archive range?
           if (msgThread.isInInbox()) {
             nArchived += 1;
             operations += 1;
             try {
                if (_gDebug) {
                  Logger.log('Archive "'+msgThread.getFirstMessageSubject()+'"');
                } else {
                  msgThread.moveToArchive();
                }
             } catch(err) {
                Logger.log('Too much archiving for "'+labelName+'"? "'+err.message+'"');
                LogPar.appendText('('+(startThread+nThreads)+') ');
                _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived);
                LogPar.appendText(' Archive error');
                return -1; // error
             }
           }
        } else { // not within archive range, so bye bye
          if (! msgThread.isInTrash()) {
            try {
              if (_gDebug) {
                  Logger.log('Trash "'+msgThread.getFirstMessageSubject()+'"');
              } else {
                msgThread.moveToTrash();
              }
              operations += 1;
              nRemoved += 1;
            } catch(err) {
              Logger.log('Too much trash for "'+labelName+'"? "'+err.message+'"');
              LogPar.appendText('('+(startThread+nThreads)+') ');
              _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived);
              LogPar.appendText(' Trash error');
              return -1; // error
            }
          }
        }
      }
      if (operations >= 100) {
        // _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived);
        operations = 0;
      }
    }
  }
  if ((nRemoved+nArchived) > 0) {
    _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived,nForever,pretrash);
  } else {
    LogPar.removeFromParent(); // don't log no-ops
  }
  return (nRemoved+nArchived);
}

//
// This is the entry point for the whole script. Be sure to set the two doc keys
//   correctly before beginning!
//
function retentionRulesMain() {
  'use strict';
  var ssDoc, logDoc;
  ssDoc = SpreadsheetApp.getActiveSpreadsheet();
  if (ssDoc === null) {
    try {
      ssDoc = SpreadsheetApp.openById(CONTROL_ID);
    } catch(err) {
      Logger.log('Error: Unable to open the retention-rules spreadsheet: "'+err.message+'"');
      return;
    }
  }
  try {
    logDoc = DocumentApp.openById(LOG_DOC_ID);
  } catch(err) {
    Logger.log('Error: Unable to open the retention log file: "'+err.message+'"');
    return;
  }
  if (logDoc.getNumChildren() < 1) {
    logDoc.appendParagraph('(end of log file)');
  }
  var now = new Date();
  var par = logDoc.insertParagraph(0, 'Retention '+(_gDebug?'TEST ':'Log ')+now.toLocaleString());
  par.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  var ss = ssDoc.getActiveSheet();
  var lc = ss.getLastColumn();
  var lr = ss.getLastRow();
  var v = ss.getRange(2,1,lr-1,lc).getValues(); // read whole spreadsheet
  var i, days, daysArch, pUnread, pRead, pStar, pImp, nItems;
  var pars = [];
  for (i=0; i<(lr-1); i+=1) {
    pars[i] = logDoc.insertListItem(i+1, '"'+v[i][0]+'": ');
    pars[i].setSpacingBefore(0);
    pars[i].setSpacingAfter(0);
    pars[i].setFontSize(8);
    pars[i].setBold(false);
  }
  for (i=1; i<(lr-1); i+=1) {
    pars[i].setListId(pars[0]);
  }
  logDoc.insertHorizontalRule(i+1);
  var earlyHalt = false;
  for (i=0; i<(lr-1); i+=1) {
    if (earlyHalt) {
      pars[i].removeFromParent();
      continue;
    }
    if (v[i][0] === '') {
      continue;
    }
    pUnread = (v[i][2] !== '') ? (parseInt(v[i][2]) !== 0) : DEFAULT.PROTECT_UNREAD;
    pRead   = (v[i][3] !== '') ? (parseInt(v[i][3]) !== 0) : DEFAULT.PROTECT_READ;
    if (pUnread && pRead) {
      Immortals[v[i][0]] = true;
    }
 }
  for (i=0; i<(lr-1); i+=1) {
    if (earlyHalt) {
      pars[i].removeFromParent();
      continue;
    }
    var label = v[i][0];
    if (label in Immortals) {
      if (_gDebug) {
        Logger.log('Never delete "'+label+'" items');
      }
    }
    if (label === '') {
      Logger.log('Skipping item '+i+': No Label');
      continue;
    }
    days = (v[i][1] !== '') ? parseInt(v[i][1]) : DEFAULT.DAYS;
    if (days < DEFAULT.MINIMUM_DAYS) {
      Logger.log('Hmm, item '+i+' ('+label+') had length of only '+days+' days. Skipping');
      continue;
    }
    //for (var j=2; j<= 6; j+=1) {
    //  Logger.log('Item '+j+': "'+v[i][j]+'" is ('+parseInt(v[i][j])+')');
    //}
    pUnread  = (v[i][2] !== '') ? (parseInt(v[i][2]) !== 0) : DEFAULT.PROTECT_UNREAD;
    pRead    = (v[i][3] !== '') ? (parseInt(v[i][3]) !== 0) : DEFAULT.PROTECT_READ;
    pStar    = (v[i][4] !== '') ? (parseInt(v[i][4]) !== 0) : DEFAULT.PROTECT_STARRED;
    pImp     = (v[i][5] !== '') ? (parseInt(v[i][5]) !== 0) : DEFAULT.PROTECT_IMPORTANT;
    daysArch = (v[i][6] !== '') ?  parseInt(v[i][6]) : 0;  
    //Logger.log(i+': "'+label+'":'+days+' days, unread:'+pUnread+', read:'+pRead+', star:'+pStar+', impo:'+pImp+', daysArch:'+daysArch);
    try {
      nItems = _applyRetention_(label,days,pUnread,pRead,pStar,pImp,daysArch,pars[i]);
    } catch(err) {
      Logger.log('Error for "'+label+'"? "'+err.message+'"');
      pars[i].appendText(' Error "'+err.message+'"');
      nItems = -1;
    }
    if (nItems < 0) { // error
      earlyHalt = true;
    }
  }
}

//
//
//
function testRules() {
  'use strict';
  _gDebug = true;
  DEFAULT.MAX_THREADS = 40;
  retentionRulesMain();
}

/// eof 
