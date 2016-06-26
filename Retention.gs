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
//  Label  NumDays Protect_Unread Protect_Read Protect_Starred Protect_Important days_Archive
// and then fill in the rows appropriately. The boolean values for each row can
// be blank (don't protect) or 1 (do protect).
//
// if both Protect_Unread and Protect_Read are true, this label goes into a special
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


// utility: shuffle an array in place
function _shuffle_array_(a) {
  'use strict';
    var j, x, i;
    for (i = a.length; i; i -= 1) {
        j = Math.floor(Math.random() * i);
        x = a[i - 1];
        a[i - 1] = a[j];
        a[j] = x;
    }
}

//
//
//
function _logUpdate_(Rule,counts) {
  'use strict';
  if ((counts.nRemoved+counts.nArchived+counts.pretrash) > 0) {
    var msg = (',  Removed ' + counts.nRemoved)
    if (counts.pretrash > 0) {
       msg = (msg + '(/' + counts.pretrash + ')');
    }
    msg = (msg + ' items >' + Rule.days + '+ days old');
    if (counts.nArchived > 0) {
      msg = (msg + ', archived ' + counts.nArchived + ' ( >' + Rule.daysArch + ' days)');
    }
    if (counts.nForever > 0) {
      msg = (msg + ', '+counts.nForever+' immortal');
    }
    if (counts.nUnread > 0) {
      msg = (msg + ', '+counts.nUnread+' unread');
    }
    if (counts.nUnchanged > 0) {
      msg = (msg + ', '+counts.nUnchanged+' protected');
    }
    if (counts.nStarSafe > 0) {
      msg = (msg + ', '+counts.nStarSafe+'*');
    }
    if (counts.nImpSafe > 0) {
      msg = (msg + ', '+counts.nImpSafe+'Imp');
    }
    if (_gDebug) {
      Logger.log(msg);
    }
    Rule.listItem.appendText(msg);
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
      // if (_gDebug) {
      //   Logger.log('Msg "'+MsgThread.getFirstMessageSubject()+'" is immortal: '+n);
      // }
      return true;
    }
  }
  return false;
}

//
// Apply retention rules to a single label
//
function _applyRetention_(LabelName,Rule) {
  'use strict';
  //var logMsg = ('applyRetention("'+LabelName+'",'+Rule.days+','+Rule.pUnread+','+Rule.pRead+','+Rule.pStar+','+Rule.pImp+','+Rule.daysArch+',...)');
  //Logger.log(logMsg);
  var now = new Date();
  var millisecsPerDay = 1000 * 60 * 60 * 24;
  var nowDay = Math.round(now.getTime() / millisecsPerDay);
  var cutoff = nowDay - Rule.days;
  var cutoffArch = nowDay - Rule.daysArch;
  if (Rule.daysArch <= 0) { // hack to avoid archiving 
    cutoffArch = cutoff - 1; // one day before the cutoff
  }
  if (_gDebug) {
    Rule.listItem.appendText('['+now.toLocaleString()+'] ');
  }
  if (Rule.immortal) {
    if (_gDebug) {
      Logger.log('Never delete "'+LabelName+'" items');
    }
    Rule.listItem.appendText('-protected-');
    return 0;
  }
  Logger.log('Starting on unprotected label "'+LabelName+'"');
  var counts = {
    nRemoved: 0,
    nArchived: 0,
    nForever: 0,
    nStarSafe: 0,
    nImpSafe: 0,
    nUnchanged: 0,
    nUnread: 0,
    operations: 0,
    pretrash: 0,
  }
  var maxThreads = DEFAULT.MAX_THREADS;
  var label;
  try {
    label = GmailApp.getUserLabelByName(LabelName);
  } catch(err) {
    Logger.log('Too many mail calls for "'+LabelName+'"? "'+err.message+'"');
    Rule.listItem.appendText(' read error');
    return -1; // error
  }
  if (! label) {
    Logger.log('Could not find label "'+LabelName+'" - check your spreadsheet');
    Rule.listItem.appendText('No such label available: "'+LabelName+'" - check your definitions!');
    Rule.listItem.setForegroundColor('#800000');
    Rule.listItem.setBold(true);
    return counts.nRemoved;
  }
  var startThread = 0;
  var pending = true;
  var i, msgThread, lastMsgDay, unrd, okayToDelete;
  var passes = 0;
  while (pending) {
    var passStart = new Date();
    if (_gDebug && (passes>0)) {
      Rule.listItem.appendText(' ..['+passStart.toLocaleString()+'] ');
    }
    var threads = label.getThreads(startThread,maxThreads); // seems to max at 500 - TO-DO: use getThreads(start,max) to fetch more past 500
    var nThreads = threads.length;
    if (nThreads === maxThreads) {
      if (_gDebug && (passes > 2)) {
        Rule.listItem.appendText(nThreads+' (+'+startThread+') initial test items');
        pending = false;
      } else {
        startThread = startThread + maxThreads; // for the next cycle
        Rule.listItem.appendText('>');
      }
    } else {
      Rule.listItem.appendText(nThreads+' (+'+startThread+') initial items');
      pending = false;
    }
    passes += 1;
    for (i = 0; i < threads.length; i += 1) {
      msgThread = threads[i];
      okayToDelete = true;
      if (msgThread.isInTrash()) {
        counts.pretrash += 1;
        continue;
      }
      if (Rule.pStar && msgThread.hasStarredMessages()) {
        counts.nStarSafe += 1;
        continue;
      }
      if (Rule.pImp && msgThread.isImportant()) {
        counts.nImpSafe += 1;
        continue;
      }
      unrd = msgThread.isUnread();
      if (Rule.pUnread && unrd) {
        counts.nUnread += 1;
        continue;
      }
      if (Rule.pRead && !unrd) {
        counts.nUnchanged += 1;
        continue;
      }
      if (_hasImmortality_(msgThread)) {
        okayToDelete = false;
        continue;
      }
      lastMsgDay = Math.round(msgThread.getLastMessageDate().getTime()/millisecsPerDay);
      if (lastMsgDay > cutoff) {  // old enough?
        if (lastMsgDay <= cutoffArch) { // within archive range?
           if (msgThread.isInInbox()) {
             counts.nArchived += 1;
             counts.operations += 1;
             try {
                if (_gDebug) {
                  Logger.log('Archive "'+msgThread.getFirstMessageSubject()+'"');
                } else {
                  msgThread.moveToArchive();
                }
             } catch(err) {
                Logger.log('Too much archiving for "'+LabelName+'"? "'+err.message+'"');
                Rule.listItem.appendText('('+(startThread+nThreads)+') ');
                _logUpdate_(Rule,counts);
                Rule.listItem.appendText(' Archive error');
                return -1; // error
             }
           }
        } else if (okayToDelete) { // not within archive range, so bye bye
          if (! msgThread.isInTrash()) {
            try {
              if (_gDebug) {
                  Logger.log('Trash "'+msgThread.getFirstMessageSubject()+'"');
              } else {
                msgThread.moveToTrash();
              }
              counts.nRemoved += 1;
              counts.operations += 1;
            } catch(err) {
              Logger.log('Too much trash for "'+LabelName+'"? "'+err.message+'"');
              Rule.listItem.appendText('('+(startThread+'->'+nThreads)+') ');
              _logUpdate_(Rule,counts);
              Rule.listItem.appendText(' Trash error');
              return -1; // error
            }
          }
        } else {
          counts.nForever += 1;
        }
      } else {
         counts.nUnchanged += 1;
      }
      if (counts.operations >= 100) {
        counts.operations = 0;
      }
    }
  }
  var totalAltered = counts.nRemoved+counts.nArchived;
  if (totalAltered > 0) {
    _logUpdate_(Rule,counts);
  } else {
    if (_gDebug) {
      Logger.log('No changes for label "'+LabelName+'"');
    }
    Rule.listItem.removeFromParent(); // don't log no-ops
  }
  return totalAltered;
}

//
// read entire sheet, extract all rules into an object. This lets us safely manage duplicates
//
function _buildRules_(SsDoc) {
  'use strict';
  var rulesObject = {};
  var ss = SsDoc.getActiveSheet();
  var lc = ss.getLastColumn();
  var lr = ss.getLastRow();
  var v = ss.getRange(2,1,lr-1,lc).getValues(); // read whole spreadsheet as 2D array
  for (i=1; i<(lr-1); i+=1) {
    var name = v[i][0];
    if (name == '') {
      continue;
    }
    var days = (v[i][1] !== '') ? parseInt(v[i][1]) : DEFAULT.DAYS;
    if (days < DEFAULT.MINIMUM_DAYS) {
      Logger.log('Odd rule '+i+' ('+labelName+') had length of only '+days+' days. Skipping');
      continue;
    }
    if (rulesObject[name] === undefined) {
      rulesObject[name] = {
        days: days,
        pUnread:  (v[i][2] !== '') ? (parseInt(v[i][2]) !== 0) : DEFAULT.PROTECT_UNREAD,
        pRead:    (v[i][3] !== '') ? (parseInt(v[i][3]) !== 0) : DEFAULT.PROTECT_READ,
        pStar:    (v[i][4] !== '') ? (parseInt(v[i][4]) !== 0) : DEFAULT.PROTECT_STARRED,
        pImp:     (v[i][5] !== '') ? (parseInt(v[i][5]) !== 0) : DEFAULT.PROTECT_IMPORTANT,
        daysArch: (v[i][6] !== '') ?  parseInt(v[i][6]) : 0,
      }
    } else {
        Logger.log('Label "'+name+'" is duplicated, opting for the least damage');
        rulesObject[name].days = Math.max(rulesObject[name].days, days),
        rulesObject[name].pUnread |=  ((v[i][2] !== '') ? (parseInt(v[i][2]) !== 0) : DEFAULT.PROTECT_UNREAD);
        rulesObject[name].pRead |=    ((v[i][3] !== '') ? (parseInt(v[i][3]) !== 0) : DEFAULT.PROTECT_READ);
        rulesObject[name].pStar |=    ((v[i][4] !== '') ? (parseInt(v[i][4]) !== 0) : DEFAULT.PROTECT_STARRED);
        rulesObject[name].pImp |=     ((v[i][5] !== '') ? (parseInt(v[i][5]) !== 0) : DEFAULT.PROTECT_IMPORTANT);
        rulesObject[name].daysArch = Math.max(rulesObject[name].daysArch, ((v[i][6] !== '') ?  parseInt(v[i][6]) : 0));
   }
   rulesObject[name].immortal = ((rulesObject[name].pUnread && rulesObject[name].pRead));
   Immortals[name] = true;
  }
  return rulesObject;
}

function _getSSRules_(DocID) // DocID may be ignored
{
  var ssDoc = SpreadsheetApp.getActiveSpreadsheet();
  if (ssDoc === null) {
    try {
      ssDoc = SpreadsheetApp.openById(DocID);
    } catch(err) {
      Logger.log('Error: Unable to open the retention-rules spreadsheet: "'+err.message+'"');
      return null;
    }
  }
  return _buildRules_(ssDoc);
}

function _getLogDoc_(DocID) {
  var logDoc;
  try {
    logDoc = DocumentApp.openById(DocID);
  } catch(err) {
    Logger.log('Error: Unable to open the retention log file: "'+err.message+'"');
    return null;
  }
  if (logDoc.getNumChildren() < 1) {
    logDoc.appendParagraph('(end of log file)');
  }
  return logDoc;
}

//
// This is the entry point for the whole script. Be sure to set the two doc keys
//   correctly before beginning!
//
function retentionRulesMain() {
  'use strict';
  var now = new Date();
  var rules = _getSSRules_(CONTROL_ID);
  if (!rules) {
    return;
  }
  var labels = Object.keys(rules);
  if (_gDebug) {
    Logger.log('Found '+labels.length+' unique labels');
  }
  labels.sort();
  //
  var logDoc = _getLogDoc_(LOG_DOC_ID);
  if (!logDoc) {
    return;
  }
  var title = logDoc.insertParagraph(0, 'Retention '+(_gDebug?'TEST ':'Log ')+now.toLocaleString());
  title.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  logDoc.insertHorizontalRule(1);
  // from here...
  var i, j,  nItems, rule;
  var listItems = [];
  var firstItem = null;
  for (i in labels) {
    j = labels.length - 1 - i;
    rule = rules[labels[j]];
    //listItems[i] = logDoc.insertListItem(i+1, (labels[i]+': ') );
    rule.listItem = logDoc.insertListItem(1, (labels[j]+': ') );
    rule.listItem.setSpacingBefore(0);
    rule.listItem.setSpacingAfter(0);
    rule.listItem.setFontSize(8);
    rule.listItem.setBold(false);
    if (j === 0) {
      firstItem = rule.listItem;
    }
    if (_gDebug) {
      Logger.log(labels[i]+' '+i);
    }
  }
  for (i in labels) {
    if (i > 0) {
      rules[labels[i]].listItem.setListId(firstItem);
    }
  }
  _shuffle_array_(labels);  // randomize
  var earlyHalt = false;
  for (i in labels) {
    var labelName = labels[i];
    rule = rules[labelName];
    if (earlyHalt) {
      if (_gDebug) {
        Logger.log('wrap '+labels[i]);
      }
      rule.listItem.removeFromParent();
      continue;
    }
    try {
      nItems = _applyRetention_(labelName, rule);
    } catch(err) {
      Logger.log('Error for "'+labelName+'"? "'+err.message+'"');
      rule.listItem.appendText(' Error "'+err.message+'"');
      nItems = -1;
    }
    if (nItems < 0) { // error
      earlyHalt = true;
    }
  }
  var finished = new Date();
  Logger.log('Completed after '+Math.floor(0.5+((finished.getTime()-now.getTime())/1000))+' seconds');
}

//
//
//
function testRules() {
  'use strict';
  _gDebug = true;
  // DEFAULT.MAX_THREADS = 40;
  alert('wow');
  return;
  retentionRulesMain();
}

/// eof 
