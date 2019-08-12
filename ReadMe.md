# Session Arbiter

A free tool for managing time limits for local login sessions on workstations running Windows Vista/7/8/10. Essentially, it aims to reproduce on regular workstations the functionality that exists for Remote Desktop Services that allows administrators to impose time limits on different kinds of login session.

Session Arbiter can also be used to configure a laptop to log off the current user when the lid is closed (an option not natively available in Windows), as well as optionally placing the laptop into Sleep or Hibernate once the logoff has finished.

Developed by [James Schlackman](https://www.schlackman.org).

## Requirements

Session Arbiter was originally designed for use with Windows Vista and Windows 7. It should also work on Windows 8 and Windows 10. It may work on earlier versions of Windows down to Windows 2000, but it will require .NET Framework 2.0 to be installed.

## Development

Session Arbiter is developed using [Visual Basic 2010 Express](http://www.microsoft.com/visualstudio/en-us/products/2010-editions/express). The installer is authored using the freeware edition of [Advanced Installer](http://www.advancedinstaller.com/), with some additions made using [InstEd](http://www.instedit.com/).

This software uses functionality provided by the [Cassia library](https://code.google.com/archive/p/cassia/) by Dan Ports.
