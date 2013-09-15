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
//   YOU WILL NEED TO CREATE THESE TWO DOCS AND PASS THEIR DOC KEYS INTO THE SCRIPT.
//   See the notes below on how to do this.
//
// To define retention rules, create a spreadsheet with these column headers:
//  Label  NumDays ProtectUnread ProtectRead ProtectStarred ProtectImportant
// and then fill in the rows appropriately. The boolean values for each row can
// be blank (don't protect) or 1 (do protect).
//
// A simple sample spreadsheet (the on I'm using today, actually) is included in this
//   github archive in CSV format. If you want to start from that, import it into GDocs.
//
// Enter the Google key (part of the url) in the SpreadsheetApp.openById(id) function found below.
//
// In addition, create a simple Google Document, called "Retention Log" or similar,
//  for collecting logs of the script behavior. Enter the key for that document in
//  the DocumentApp.openById(id) function found below
//

//
// Apply retention rules to a single label
//
function _applyRetention_(labelName,daysBack,ProtUnread,ProtRead,ProtStar,ProtImp,LogPar) {
  var removed = 0;
  var label = GmailApp.getUserLabelByName(labelName);
  if (label == undefined) {
    Logger.log('Could not find label "'+labelName+'"');
    LogPar.appendText("No such label available");
    LogPar.setForegroundColor('#800000');
    LogPar.setBold(true);
    return removed;
  }
  var threads = label.getThreads(); // seems to max at 500
  var now = new Date();
  var dayMilli = 1000 * 60 * 60 * 24;
  var cutoff = now.getTime() - daysBack*dayMilli;
  //Logger.log('"'+labelName+'": '+threads.length+' initial items');
  LogPar.appendText((threads.length+' initial items'));
  for (var i = 0; i < threads.length; i += 1) {
    var t = threads[i];
    var d = t.getLastMessageDate();
    var dt = d.getTime();
    if (dt > cutoff) continue;
    var u = t.isUnread();
    if (ProtUnread && u) continue;
    if (ProtRead && (!u)) continue;
    if (ProtStar && t.hasStarredMessages()) continue;
    if (ProtImp && t.isImportant()) continue;
    if (! t.isInTrash()) {
      // var n = t.getFirstMessageSubject();
      // Logger.log("trashing "+d+": "+n);
      t.moveToTrash();
      removed += 1;
    }
  }
  if (removed > 0) {
    LogPar.appendText('  removed '+removed+' old items.');
    //Logger.log(labelName+': removed '+removed+' old items.');
  } else {
    LogPar.removeFromParent();
  }
  return removed;
};

//
// This is the entry point for the whole script. Be sure to set the two doc keys
//   correctly before beginning!
//
function retentionRulesMain() {
  var s = SpreadsheetApp.openById("PUT_YOUR_SPREADSHEET_DOC_KEY_HERE");
  var logDoc = DocumentApp.openById("PUT_YOUR_LOG_DOCUMENT_DOC_KEY_HERE");
  if (logDoc.getNumChildren() < 1) {
    logDoc.appendParagraph("(end of log file)");
  }
  var now = new Date();
  var par = logDoc.insertParagraph(0, 'Retention Log '+now);
  par.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  var ss = s.getActiveSheet();
  var lc = ss.getLastColumn();
  var lr = ss.getLastRow();
  var v = ss.getRange(2,1,lr-1,lc).getValues();
  var i, days, pUnread, pRead, pStar, pImp, nItems;
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
  for (i=0; i<(lr-1); i+=1) {
    days = 365;
    pUnread = false;
    pRead = false;
    pStar = false;
    pImp = false;
    lab = v[i][0];
    if (lab =="") {
      Logger.log("skipping item "+i+': no label');
      continue;
    }
    if (v[i][1] != "") days = v[i][1];
    if (days < 7) {
      Logger.log("Hmm, item "+i+'had length of only '+days+' days. Skipping');
      continue;
    }
    if (v[i][2] != "") pUnread = (v[i][2] != 0);
    if (v[i][3] != "") pRead = (v[i][3] != 0);
    if (v[i][4] != "") pStar = (v[i][4] != 0);
    if (v[i][5] != "") pImp = (v[i][5] != 0);
    //Logger.log(i+': "'+v[i][0]+'":'+days+' days, '+pUnread+', '+pRead+', '+pStar+', '+pImp+'...');
    nItems = _applyRetention_(v[i][0],days,pUnread,pRead,pStar,pImp,pars[i]);
    //Logger.log('   '+nItems+" items removed");
  }
  //logDoc.saveAndClose(); // no need in this specific case
}

