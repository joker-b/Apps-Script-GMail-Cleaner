Apps-Script-GMail-Cleaner
=========================

This is an Apps Script tool that I use for managing my free-level GMail account.
It implements a simple, GDocs-driven retention scheme.

To run this, create a copy of Retention.gs in your Google Docs folder.

Also create a spreadsheet containing retention rules (a sample is provided), and an empty Google Docs text document which will be used to store the log data from Retention.gs

Once all the pieces are ready, you will need to get the GDocs doc ID's for the spreadsheet and the log document.
The quickest way to get the doc ID's is to open each docs and chack the URL -- extract the long "base64" string from it, e.g. https://docs.google.com/document/d/3ckYOu8kuIfBzbu-Dtu9XwGHUnUJG32PK7wHe5xMv3VG/ has document id 3ckYOu8kuIfBzbu-Dtu9XwGHUnUJG32PK7wHe5xMv3VG

Paste these two ID's into your copy of Retention.gs in the indicated locations (it's pretty obvious -- this is a short program!). Save.

RunRetention.gs by hand in Apps Script, or (my method) set it on a timer, known in App Script as "triggers."

Bon app!
