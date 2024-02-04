# Google Apps Script sampler
A collection of functions I've written using Google Apps Script.

Google Apps Script is a javascript-based language with tools connected to different Google apps, such as reading from/writing to a spreadsheet or sending an email through Gmail.

Currently, 4 functions are included:
- copyText: Reads a delimited text file (csv or txt) and copies to a spreadsheet
- updateNames: Given an attendance form and a spreadsheet of its responses, checks for names already marked and flags them in the form for future submissions
- sendAttendanceEmail: Takes a list of present/absent names from a spreadsheet and formats them into an email report
- tableHtml: Turns an array into a string of that array's contents as an HTML table
