
var
  APP_NAME = 'ymdFolder',
  FILE_NAME = 'ymdFolder.js',
  COMMAND_LABEL = 'Create Date/Time Folder(&Y)',

  shell = WScript.CreateObject('WScript.Shell'),
  fso = new ActiveXObject('Scripting.FileSystemObject'),
  path = fso.GetFile(WScript.ScriptFullName).ParentFolder.Path + '\\' + FILE_NAME,
  keyDir = 'HKCR\\Directory\\shell\\' + APP_NAME,
  keyDrv = 'HKCR\\Drive\\shell\\' + APP_NAME,
  keyExists;

// Run by WScript
if (/\\cscript\.exe$/i.test(WScript.FullName)) {
  shell.Run('wscript "' + WScript.ScriptFullName + '" "' + parentPath + '"');
  WScript.Quit();
}

if (!fso.FileExists(path)) {
  shell.Popup('Not Found File: ' + path, 0, APP_NAME, 48);
  WScript.Quit(1);
}

try {
  shell.RegRead(keyDir + '\\');
  keyExists = true;
} catch (e) {}

if (shell.Popup((keyExists ? 'Unregister ' : 'Register ') + APP_NAME +
    '.\nContinue?', 0, APP_NAME, 36) !== 6) {
  WScript.Quit(1);
}

if (keyExists) { // Remove
  shell.RegDelete(keyDir + '\\command\\');
  shell.RegDelete(keyDir + '\\');
  shell.RegDelete(keyDrv + '\\command\\');
  shell.RegDelete(keyDrv + '\\');
} else {
  shell.RegWrite(keyDir + '\\', COMMAND_LABEL);
  shell.RegWrite(keyDir + '\\command\\', 'wscript \"' + path + '\" \"%1\"');
  shell.RegWrite(keyDrv + '\\', COMMAND_LABEL);
  shell.RegWrite(keyDrv + '\\command\\', 'wscript \"' + path + '\" \"%1\"');
}

shell.Popup('Success!', 0, APP_NAME, 64);
WScript.Quit();
