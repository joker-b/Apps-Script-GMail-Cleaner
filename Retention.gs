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
//   You will need to create these two drive docs and pass their keys into the script.
//
// To define retention rules, create a spreadsheet with these column headers:
//  Label  NumDays ProtectUnread ProtectRead ProtectStarred ProtectImportant daysArchive
// and then fill in the rows appropriately. The boolean values for each row can
// be blank (don't protect) or 1 (do protect).
//
// Enter the Google key (part of the url) in the SpreadsheetApp.openById(id) function found below.
//
// In addition, create a simple Google Document, called "Retention Log" or similar,
//  for collecting logs of the script behavior. Enter the key for that document in
//  the DocumentApp.openById(id) function found below
//

// Set these as you please -- these apply to labelled cels that are otherwise empty.
//  Labels that are NOT listed get NO handling (that is, they are never culled)
var DEFAULT_DAYS = 365;
var DEFAULT_PROTECT_UNREAD = false;
var DEFAULT_PROTECT_READ = false;
var DEFAULT_PROTECT_STARRED = true;
var DEFAULT_PROTECT_IMPORTANT = true;

// doc keys for the control spreadsheet and the text logfile
var CONTROL_ID = "the_url_key_for_the_control_spreadsheet_here";
var LOG_DOC_ID = "the_url_key_for_the_retention_log_file__here";

//
//
//
function _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived) {
  if ((nRemoved+nArchived) > 0) {
    var msg = ('  Removed ' + nRemoved + ' items ' + daysBack + '+ days old');
    //if (daysArchive > 0) {
      msg = (msg + ', archived ' + nArchived + ' (' + daysArchive + '+ days).');
    //} else {
    //  msg = (msg + '.');
    //}
    //Logger.log(msg);
    LogPar.appendText(msg);
  }
}

//
// Apply retention rules to a single label
//
function _applyRetention_(labelName,daysBack,ProtUnread,ProtRead,ProtStar,ProtImp,daysArchive,LogPar) {
  var logMsg = ('applyRetention("'+labelName+'",'+daysBack+','+ProtUnread+','+ProtRead+','+ProtStar+','+ProtImp+','+daysArchive+',...)');
  //Logger.log(logMsg);
  var nRemoved = 0;
  var nArchived = 0;
  var operations = 0;
  var maxThreads = 500;
  try {
    var label = GmailApp.getUserLabelByName(labelName);
  } catch(err) {
    Logger.log('Too many mail calls for "'+labelName+'"? "'+err.message+'"');
    LogPar.appendText(" read error");
    return -1; // error
  }
  if (label == undefined) {
    Logger.log('Could not find label "'+labelName+'"');
    LogPar.appendText("No such label available");
    LogPar.setForegroundColor('#800000');
    LogPar.setBold(true);
    return nRemoved;
  }
  var now = new Date();
  var nowTime = now.getTime();
  var dayMilli = 1000 * 60 * 60 * 24;
  nowTime = Math.round(nowTime / dayMilli);
  var cutoff = nowTime - daysBack; // *dayMilli;
  var cutoffArch = nowTime - daysArchive; // *dayMilli;
  //Logger.log("nowTime:"+nowTime);
  //Logger.log("cutoff:"+cutoff);
  //Logger.log("cutoffArch:"+cutoffArch);
  //var halt = cutoff/0;
  //Logger.log("halt:"+halt);
  if (daysArchive <= 0) { // hack to avoid archiving at all
    cutoffArch = cutoff - dayMilli; // one day before the cutoff
  }
  //Logger.log('"'+labelName+'": '+threads.length+' initial items');
  var startThread = 0;
  var pending = true;
  var i, thrd, dThrd, dayThrd, unrd, star, impo;
  while (pending) {
    var threads = label.getThreads(startThread,maxThreads); // seems to max at 500 - TO-DO: use getThreads(start,max) to fetch more past 500
    var nThreads = threads.length;
    if (nThreads == maxThreads) {
      startThread = startThread + maxThreads; // for the next cycle
      LogPar.appendText('>');
    } else {
      LogPar.appendText((nThreads+startThread)+' initial items');
      pending = false;
    }
    for (i = 0; i < threads.length; i += 1) {
      thrd = threads[i];
      dThrd = thrd.getLastMessageDate();
      dayThrd = dThrd.getTime();
      dayThrd = Math.round(dayThrd/dayMilli);
      //halt = int(halt);
      star = thrd.hasStarredMessages();
      impo = thrd.isImportant();
      //Logger.log("message days:"+dayThrd+' star '+star+' impo '+impo);
      if ((ProtStar && star) || (ProtImp && impo)) {
        continue;
      }
      if (dayThrd > cutoff) {
        if (dayThrd <= cutoffArch) {
           //Logger.log("message days:"+dayThrd);
           if (thrd.isInInbox()) {
             nArchived += 1;
             operations += 1;
             try {
               thrd.moveToArchive();
             } catch(err) {
                Logger.log('Too much archiving for "'+labelName+'"? "'+err.message+'"');
                LogPar.appendText('('+(startThread+nThreads)+') ');
                _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived);
                  LogPar.appendText(" Archive error");
                return -1; // error
             }
           //} else {
           //  Logger.log('archive of thread that was already archived? '+nArchived+' "'+thrd.getFirstMessageSubject()+'"');
           }
        }
      } else {
        unrd = thrd.isUnread();
        if ((ProtUnread && unrd) || (ProtRead && (!unrd))) {
          continue;
        }
        if (! thrd.isInTrash()) {
          try {
            thrd.moveToTrash();
            operations += 1;
            nRemoved += 1;
          } catch(err) {
            Logger.log('Too much trash for "'+labelName+'"? "'+err.message+'"');
            LogPar.appendText('('+(startThread+nThreads)+') ');
            _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived);
            LogPar.appendText(" Trash error");
            return -1; // error
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
    _logUpdate_(daysBack,daysArchive,LogPar,nRemoved,nArchived);
  } else {
    LogPar.removeFromParent();
  }
  return (nRemoved+nArchived);
};

//
// This is the entry point for the whole script. Be sure to set the two doc keys
//   correctly before beginning!
//
function retentionRulesMain() {
  var s, logDoc;
  try {
    s = SpreadsheetApp.openById(CONTROL_ID);
  } catch(err) {
    Logger.log('Error: Unable to open the retention-rules spreadsheet: "'+err.message+'"');
    return;
  }
  try {
    logDoc = DocumentApp.openById(LOG_DOC_ID);
  } catch(err) {
    Logger.log('Error: Unable to open the retention log file: "'+err.message+'"');
    return;
  }
  if (logDoc.getNumChildren() < 1) {
    logDoc.appendParagraph("(end of log file)");
  }
  var now = new Date();
  var par = logDoc.insertParagraph(0, 'Retention Log '+now);
  par.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  var ss = s.getActiveSheet();
  var lc = ss.getLastColumn();
  var lr = ss.getLastRow();
  var v = ss.getRange(2,1,lr-1,lc).getValues(); // read whole spreadsheet
  var i, days, daysArch, pUnread, pRead, pStar, pImp, nItems;
  var pars = new Array();
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
    days = DEFAULT_DAYS;
    pUnread = DEFAULT_PROTECT_UNREAD;
    pRead = DEFAULT_PROTECT_READ;
    pStar = DEFAULT_PROTECT_STARRED;
    pImp = DEFAULT_PROTECT_IMPORTANT;
    daysArch = 0;
    lab = v[i][0];
    if (lab =="") {
      Logger.log("skipping item "+i+': no label');
      continue;
    }
    if (v[i][1] != "") days = parseInt(v[i][1]);
    if (days < 7) {
      Logger.log("Hmm, item "+i+' ('+lab+') had length of only '+days+' days. Skipping');
      continue;
    }
    //for (var j=2; j<= 6; j+=1) {
    //  Logger.log('Item '+j+': "'+v[i][j]+'" is ('+parseInt(v[i][j])+')');
    //}
    if (v[i][2] !== "") {pUnread = (parseInt(v[i][2]) != 0);/*Logger.log("pUnread:"+pUnread);*/}
    if (v[i][3] !== "") {pRead = (parseInt(v[i][3]) != 0);/*Logger.log("pRead:"+pRead);*/}
    if (v[i][4] !== "") {pStar = (parseInt(v[i][4]) != 0);/*Logger.log("pStar:"+pStar);*/}
    if (v[i][5] !== "") {pImp = (parseInt(v[i][5]) != 0);/*Logger.log("pImp:"+pImp);*/}
    if (v[i][6] !== "") {daysArch = parseInt(v[i][6]);/*Logger.log("dayArch:"+daysArch);*/}
    //Logger.log(i+': "'+v[i][0]+'":'+days+' days, unread:'+pUnread+', read:'+pRead+', star:'+pStar+', impo:'+pImp+', daysArch:'+daysArch);
    try {
      nItems = _applyRetention_(v[i][0],days,pUnread,pRead,pStar,pImp,daysArch,pars[i]);
    } catch(err) {
      Logger.log('Error for "'+lab+'"? "'+err.message+'"');
      pars[i].appendText(' Error "'+err.message+'"');
      nItems = -1;
    }
    if (nItems < 0) { // error
      earlyHalt = true;
    }
    //Logger.log('   '+nItems+" items removed");
  }
  //logDoc.saveAndClose(); // no need in this specific case
}

