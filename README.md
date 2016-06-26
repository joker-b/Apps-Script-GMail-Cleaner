Apps-Script-GMail-Cleaner
=========================

This is an Apps Script tool that I use for managing my free-level GMail account.
It implements a simple, GDocs-driven retention scheme -- that is, it removes old, uninteresting mail. What
defines "uninteresting" is specified in a Google docs spreadsheet.

<pre>
// CAUTION: THIS SCRIPT DELETES EMAIL MESSAGES FROM YOUR ACCOUNT. *THAT'S WHAT IT'S FOR.*
//   The author bears no responsibility for email that you may erroneously or accidentally
//   or otherwise delete if you run this script.
</pre>

To use this, create a spreadsheet containing retention rules (a sample is provided), and an empty Google Docs text document which will be used to store the log data from Retention.gs. With the spreadsheet app open, select "Tools->Script Editor..." and copy-paste to replace the contents the edit window with the code here in Retention.gs.

Once all the pieces are ready, you will need to get the GDocs doc ID for the log document.
The quickest way to get this ID is to open the doc and check the URL -- extract the long "base64" string
from it, e.g. https://docs.google.com/document/d/3ckYOu8kuIfBzbu-Dtu9XwGHUnUJG32PK7wHe5xMv3VG/ has document
id 3ckYOu8kuIfBzbu-Dtu9XwGHUnUJG32PK7wHe5xMv3VG

Paste this ID into your copy of the script as `LOG_DOC_ID` -- just near the top. Save.

To test your rules, run `testRules` -- to actually start archiving, run `retentionRulesMain`

RunRetention.gs by hand in Apps Script, or (my method) set it on a timer, known in App Script as "triggers." Triggers
can be assigned under the App Script Resource menu.

Bon app!
