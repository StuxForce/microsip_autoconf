/* ============================================================================
 * MicroSIP softphone auto configuration script for using in enterprise
 * Powered by: Denis Pantsyrev <denis.pantsyrev@gmail.com>
 * ============================================================================
*/


// Settings below may be changed by system administrator.
var allowedPC = /^.*/i;
var desktopLink = true;
var scriptSrvDir = '\\\\path\\to\\microsip_autoconf.js\\shared\\folder';


// Script internal variables. DO NOT CHANGE!
var WSN = WScript.CreateObject('WScript.Network');
var FSO = WScript.CreateObject('Scripting.FileSystemObject');
var WSS = WScript.CreateObject('WScript.Shell');

var distFolderName = 'Dist';
var usersFolderName = 'Users';
var confFolderName = 'MicroSIP';
var confFileName = 'MicroSIP.ini';
var execFileName = 'microsip.exe';
var lnkFileName = 'MicroSIP.lnk';

var srvDistPath = scriptSrvDir + '\\' + distFolderName;
var usrConfPath = WSS.ExpandEnvironmentStrings('%APPDATA%') + '\\' + confFolderName;
var mainConfFileOnSrv = scriptSrvDir + '\\' + confFileName;
var usrConfFileOnSrv = scriptSrvDir + '\\' + usersFolderName + '\\' + WSN.UserName + '.ini';
var usrConfFileOnLoc = usrConfPath + '\\' + confFileName;
var usrExecFilePath = usrConfPath + '\\' + execFileName;
var mainSrvINI = {};
var usrSrvINI = {};
var usrLocINI = {};

var needRestart = 0;
var cmdLine = 'taskkill.exe /F /FI "USERNAME eq ' + WSN.UserName + '" /IM microsip.exe';


if (!(allowedPC.test(WSN.ComputerName))
	|| !(FSO.FileExists(usrConfFileOnSrv) && FSO.GetFile(usrConfFileOnSrv).Size > 0)
) {
	WScript.Quit();
}

// Close microsip.exe for update if it running
if (/ [0-9]+/.test(WSS.Exec(cmdLine).StdOut.ReadAll())) {
	needRestart = 1;
	WScript.Sleep(1000);
}

// Update MicroSIP files
if (FSO.FolderExists(srvDistPath)) {
	FSO.CopyFolder(srvDistPath, usrConfPath);
	if (desktopLink) {
		strDesktop = WSS.SpecialFolders('Desktop');
		oMyShortcut = WSS.CreateShortcut(strDesktop + '\\' + lnkFileName);
		oMyShortcut.WindowStyle = 4;
		oMyShortcut.IconLocation = usrExecFilePath + ', 0';
		oMyShortcut.TargetPath = usrExecFilePath;
		oMyShortcut.WorkingDirectory = usrConfPath;
		oMyShortcut.Save();
	}
}

ReadINIFile(mainSrvINI, mainConfFileOnSrv);
ReadINIFile(usrSrvINI, usrConfFileOnSrv);
ReadINIFile(usrLocINI, usrConfFileOnLoc);

MergeINIObj(usrLocINI, mainSrvINI);
MergeINIObj(usrLocINI, usrSrvINI);

SaveINIFile(usrLocINI, usrConfPath, usrConfFileOnLoc);

// Update Registry to associate MicroSIP application with sip:// uri
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\', 'Internet Call Protocol', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\URL Protocol', '', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\Owner Name', 'MicroSIP', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\DefaultIcon', usrExecFilePath + ',0', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\shell\\open\\command\\', '\"' + usrExecFilePath + '\" %1', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\RegisteredApplications\\MicroSIP', 'SOFTWARE\\MicroSIP\\Capabilities', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\', usrConfPath, 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\Start Menu Folder', 'MicroSIP', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\Capabilities\\ApplicationDescription', 'MicroSIP Softphone', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\Capabilities\\ApplicationName', 'MicroSIP', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\Capabilities\\UrlAssociations\\sip', 'MicroSIP', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\MicroSIP\\', 'Internet Call Protocol', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\MicroSIP\\DefaultIcon\\', usrExecFilePath + ',0', 'REG_SZ');
WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\MicroSIP\\shell\\open\\command\\', '\"' + usrExecFilePath + '\" %1', 'REG_SZ');

if (needRestart) {
	WSS.run(usrConfPath + '\\' + execFileName, 1, false);
}


// Read .ini file to object
function ReadINIFile(INIObj, filename) {
	if ((FSO.FileExists(filename)) && (FSO.GetFile(filename).Size > 0)) {
		var fileh = FSO.OpenTextFile(filename, 1);
		while (!fileh.AtEndOfStream) {
			var line = fileh.ReadLine();
			if (/^\[(\w+)\]/.test(line)) {
				var section = RegExp.$1.toLowerCase();
				INIObj[section] = {};
			}
			if (/^([^;#][^=]*?)\s*=\s*([^\r\n]*?)\s*$/.test(line)) {
				var param = RegExp.$1.toLowerCase();
				var value = RegExp.$2;
				INIObj[section][param] = value;
			}
		}
		fileh.Close();
	}
}

// Read object to .ini file
function SaveINIFile(INIObj, confPath, filename) {
	function WriteArray(arr) {
		for (var i in arr) {
			var value = arr[i];
			if (typeof (value) === 'object') {
				if (value) {
					file.Write('[' + i + ']\r\n');
					WriteArray(value);
				}
			} else {
				if (value) {
					file.Write(i + '=' + value + '\r\n');
				}
			}
		}
	}

	if (!FSO.FolderExists(confPath)) {
		FSO.CreateFolder(confPath);
	}
	var file = FSO.OpenTextFile(filename, 2, true);
	WriteArray(INIObj);
	file.Close();
}

// Merge objects
function MergeINIObj(firstINIObj, secondINIObj) {
	for (var section in secondINIObj) {
		if (!firstINIObj[section]) {
			firstINIObj[section] = {};
		}
		for (var param in secondINIObj[section]) {
			if (!firstINIObj[section][param] || firstINIObj[section][param] !== secondINIObj[section][param]) {
				firstINIObj[section][param] = secondINIObj[section][param];
			}
		}
	}
}
