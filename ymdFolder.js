/*
 * ymdFolder
 * v0.2.0
 * https://github.com/anseki/ymd-folder
 *
 * Copyright (c) 2015 anseki
 * Licensed under the MIT license.
 */

var
  FORMAT = '%Y%M%D',
  /*
    %y :    year      (4 digits)              e.g. 2015
    %Y :    year      (2 digits)              e.g. 15
    %m :    month     (1 or 2 digits)         e.g. 5
    %M :    month     (2 digits)              e.g. 05
    %d :    date      (1 or 2 digits)         e.g. 5
    %D :    date      (2 digits)              e.g. 05
    %H :    hours     (2 digits)              e.g. 23
    %I :    minutes   (2 digits)              e.g. 59
    %S :    seconds   (2 digits)              e.g. 59
  */

  APP_NAME = 'ymdFolder',
  shell = WScript.CreateObject('WScript.Shell'),
  fso = new ActiveXObject('Scripting.FileSystemObject'),
  parentPath = WScript.Arguments(0),
  folderName = dateFormat(FORMAT),
  inc = 0, path;

// Run by WScript
if (/\\cscript\.exe$/i.test(WScript.FullName)) {
  shell.Run('wscript.exe "' + WScript.ScriptFullName + '" "' + parentPath + '"');
  WScript.Quit();
}

while (true) {
  path = parentPath + '\\' + folderName + (inc ? '-' + inc : '');
  try {
    fso.CreateFolder(path);
    break;
  } catch (e) {
    if (e.number !== -2146828230) {
      shell.Popup(e.description + '(' + e.number + ')', 0, APP_NAME, 48);
      WScript.Quit(1);
    }
    inc++;
  }
}

WScript.Quit();

function dateFormat(template, date) {
  var dateValues;
  function rightDigits(value, digits) {
    // return ('0000' + value).substr(-digits);
    // negative arg is not supported by JScript.
    var text = '0000' + value;
    return text.substr(text.length - digits);
  }

  date = date || new Date();
  dateValues = {
    year:           date.getFullYear(),
    month:          date.getMonth() + 1,
    date:           date.getDate(),
    hours:          date.getHours(),
    minutes:        date.getMinutes(),
    seconds:        date.getSeconds()
  };
  return template
    .replace(/\%y/g, dateValues.year)
    .replace(/\%Y/g, rightDigits(dateValues.year, 2))
    .replace(/\%m/g, dateValues.month)
    .replace(/\%M/g, rightDigits(dateValues.month, 2))
    .replace(/\%d/g, dateValues.date)
    .replace(/\%D/g, rightDigits(dateValues.date, 2))
    .replace(/\%H/g, rightDigits(dateValues.hours, 2))
    .replace(/\%I/g, rightDigits(dateValues.minutes, 2))
    .replace(/\%S/g, rightDigits(dateValues.seconds, 2));
}
