# microsip_autoconf
MicroSIP softphone auto configuration script for using in Enterprise

Allow to simplify distribution, configuring and management of softphone software in Enterprise 
by using simple JScript script and Group Policy

## Files and folders definitions:

**Dist** - Folder with a portable version of MicroSIP. Distributed to workstations

**Users** - User configuration files. '__tpl.ini' is a template file with sample of some options. 
Settings in this file are override all other settings (user settings on local computer and from the main MicroSIP.ini file on the server side)

**MicroSIP.ini** - The main configuration file (general settings). Overlapped by settings from the user settings file (in Users dir)

**microsip_autoconf.js** - Group Policy distribution file that must be run at user logon time on workstation.

**updater.ini** - Some autoconf settings

**updater.log** - Some log data of script execution (errors, etc)

## How it work

- clone this repo in shared folder.
- allow all domain users *read* permittions on all files and folders. For Users dir turn off inheritance (each user must have read access only to his config file)
- allow all domain users *write* permittions on **updater.log**
- rename **updater.ini.sample** to **updater.ini** and change settings (if needed) 
- change Group Policy to run **microsip_autoconf.js** file at user logon
- change settings in **MicroSIP.ini** file (if needed)
	- allowedPC: Reulal expression of allowed computers. If set updater will work only on this computers (Default: .*)
	- desktopLink: if set true, it create link to MicroSIP.exe on desktop (Default: true)
	- saveUserContacts: if set false, it replaces user's **Contacts.xml** file by file **Contacts.xml** from Dist folder (Default: false)

To add new user's config do these:
- make a copy of **__tpl.ini** file in Users dir
- rename it to **<SAMAccountName_of_user>.ini**
- give the user rights to read 
- change options in this new config file.

Now, if user has his own config file in Users directory, the script will install MicroSIP distro at user's local PC and apply config settings.
