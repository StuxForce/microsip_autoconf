/* ============================================================================
 * MicroSIP updater script for using in Enterprise
 * Powered by: Denis Pantsyrev <denis.pantsyrev@gmail.com>
 * ============================================================================
*/

// Script internal variables. DO NOT CHANGE!
var WSN = WScript.CreateObject('WScript.Network');
var FSO = WScript.CreateObject('Scripting.FileSystemObject');
var WSS = WScript.CreateObject('WScript.Shell');
var srvScriptDir = WScript.ScriptFullName.substring(0, WScript.ScriptFullName.lastIndexOf(WScript.ScriptName) - 1)

try {
	// Get updater config
	var updaterINI = {};
	ReadINIFile(updaterINI, srvScriptDir + '\\updater.ini');
	var allowedPC = new RegExp(updaterINI.settings.allowedpc, "i");
	var desktopLink = updaterINI.settings.desktoplink;

	// Prepare names
	var srvDistFolderName = 'Dist';
	var srvUsersFolderName = 'Users';
	var distFolderName = 'MicroSIP';
	var confFileName = 'MicroSIP.ini';
	var contactsFileName = 'Contacts.xml';
	var execFileName = 'microsip.exe';
	var lnkFileName = 'MicroSIP.lnk';
	var tmpDistFolderName = 'tmp_MicroSIP';

	var srvDistPath = srvScriptDir + '\\' + srvDistFolderName;
	var usrDistPath = WSS.ExpandEnvironmentStrings('%APPDATA%') + '\\' + distFolderName;
	var tmpDistPath = WSS.ExpandEnvironmentStrings('%APPDATA%') + '\\' + tmpDistFolderName;

	var tmpConfFile = tmpDistPath + '\\' + confFileName
	var tmpContactsFile = tmpDistPath + '\\' + contactsFileName;
	var srvMainConfFile = srvScriptDir + '\\' + confFileName;
	var srvUserConfFile = srvScriptDir + '\\' + srvUsersFolderName + '\\' + WSN.UserName + '.ini';
	var usrConfFile = usrDistPath + '\\' + confFileName;
	var usrContactsFile = usrDistPath + '\\' + contactsFileName;
	var usrExecFile = usrDistPath + '\\' + execFileName;

	var mainSrvINI = {};
	var usrSrvINI = {};
	var usrLocINI = {};

	var needRestart = 0;
	var cmdLine = 'taskkill.exe /FI "USERNAME eq ' + WSN.UserName + '" /IM ' + execFileName;

	if (!(allowedPC.test(WSN.ComputerName))
		|| !(FSO.FileExists(srvUserConfFile) && FSO.GetFile(srvUserConfFile).Size > 0)
	) {
		WScript.Quit();
	}

	// Update MicroSIP files
	if (FSO.FolderExists(srvDistPath)) {
		FSO.CopyFolder(srvDistPath, tmpDistPath);
		if (FSO.FileExists(usrConfFile)) {
			FSO.CopyFile(usrConfFile, tmpConfFile);
		}
		if (FSO.FileExists(usrContactsFile)) {
			FSO.CopyFile(usrContactsFile, tmpContactsFile);
		}
		if (desktopLink == "true") {
			strDesktop = WSS.SpecialFolders('Desktop');
			oMyShortcut = WSS.CreateShortcut(strDesktop + '\\' + lnkFileName);
			oMyShortcut.WindowStyle = 4;
			oMyShortcut.IconLocation = usrExecFile + ', 0';
			oMyShortcut.TargetPath = usrExecFile;
			oMyShortcut.WorkingDirectory = usrDistPath;
			oMyShortcut.Save();
		}
	}

	ReadINIFile(mainSrvINI, srvMainConfFile);
	ReadINIFile(usrSrvINI, srvUserConfFile);
	ReadINIFile(usrLocINI, tmpConfFile, -1);

	MergeINIObj(usrLocINI, mainSrvINI);
	MergeINIObj(usrLocINI, usrSrvINI);

	SaveINIFile(usrLocINI, tmpDistPath, tmpConfFile);

	// Update Registry to associate MicroSIP application with sip:// uri
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\', 'Internet Call Protocol', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\URL Protocol', '', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\Owner Name', 'MicroSIP', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\DefaultIcon', usrExecFile + ',0', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\sip\\shell\\open\\command\\', '\"' + usrExecFile + '\" %1', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\RegisteredApplications\\MicroSIP', 'SOFTWARE\\MicroSIP\\Capabilities', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\', usrDistPath, 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\Start Menu Folder', 'MicroSIP', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\Capabilities\\ApplicationDescription', 'MicroSIP Softphone', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\Capabilities\\ApplicationName', 'MicroSIP', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Wow6432Node\\MicroSIP\\Capabilities\\UrlAssociations\\sip', 'MicroSIP', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\MicroSIP\\', 'Internet Call Protocol', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\MicroSIP\\DefaultIcon\\', usrExecFile + ',0', 'REG_SZ');
	WSS.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Classes\\MicroSIP\\shell\\open\\command\\', '\"' + usrExecFile + '\" %1', 'REG_SZ');

	// Close microsip.exe for update if it running
	if (/ [0-9]+/.test(WSS.Exec(cmdLine).StdOut.ReadAll())) {
		needRestart = 1;
		WScript.Sleep(3000);
	}

	if (FSO.FolderExists(usrDistPath)) {
		FSO.DeleteFolder(usrDistPath, 1);
	}

	FSO.MoveFolder(tmpDistPath, usrDistPath);

	if (needRestart) {
		WSS.run(usrDistPath + '\\' + execFileName, 1, false);
	}

} catch (e) {
	WSS.LogEvent(1, WScript.ScriptName + ' ' + e.name + ': ' + e.message)
}


// Read .ini file to object
function ReadINIFile(INIObj, filename, format) {
	if ((FSO.FileExists(filename)) && (FSO.GetFile(filename).Size > 0)) {
		var fileh = FSO.OpenTextFile(filename, 1, false, format);
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
