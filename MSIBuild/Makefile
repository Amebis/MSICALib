#
#    Copyright 1991-2016 Amebis
#
#    This file is part of MSICA.
#
#    MSICA is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    MSICA is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with MSICA. If not, see <http://www.gnu.org/licenses/>.
#

!INCLUDE "..\..\..\include\MSIBuildCfg.mak"

MSIBUILD_IS_LOCALIZEABLE=1


######################################################################
# Error

All :: \
	"$(LANG).$(PLAT).$(CFG).Error-1.idt" \
	"$(LANG).$(PLAT).$(CFG).Error-2.idt"

"$(LANG).$(PLAT).$(CFG).Error-1.idt" : "Makefile" "..\..\..\include\MSIBuildCfg.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Error	Message
i2	L0
Error	Error
<<NOKEEP

"en_US.$(PLAT).$(CFG).Error-2.idtx" : "Makefile" "..\..\..\include\MSIBuildCfg.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Error	Message
i2	L0
1252	Error	Error
2550	Error opening installation package. Please, contact your support personnel.
2551	Error [2] creating task list. Please, contact your support personnel.
2552	Error [3] writing to "[2]" task list file. Please, contact your support personnel.
2560	Error [3] reading from "[2]" task list file. Please, contact your support personnel.
2553	Error [3] setting "[2]" parameter. Please, contact your support personnel.
2554	Error [3] deleting "[2]" file. Please, contact your support personnel.
2555	Error [4] moving "[2]" file to "[3]". Please, contact your support personnel.
2556	Error [3] creating "[2]" scheduled task. Please, contact your support personnel.
2557	Error [3] deleting "[2]" scheduled task. Please, contact your support personnel.
2558	Error [3] enabling/disabling "[2]" scheduled task. Please, contact your support personnel.
2559	Error [4] copying "[2]" scheduled task to "[3]". Please, contact your support personnel.
2569	Error [3] installing certificate to certificate store "[2]". Please, contact your support personnel.
2570	Error [3] removing certificate from certificate store "[2]". Please, contact your support personnel.
2571	Error [3] changing service "[2]" start type. Please, contact your support personnel.
2572	Error [3] starting service "[2]". Please, contact your support personnel.
2573	Error [3] stopping service "[2]". Please, contact your support personnel.
2580	Error installing WLAN profiles, because WLAN is not installed. Please, install Wireless LAN Service Windows feature, or contact your support personnel.
2579	Error opening WLAN handle, because WLAN AutoConfig service is not started. Please, enable and start WLAN AutoConfig service, or contact your support personnel.
2577	Error [2] opening WLAN handle. Please, contact your support personnel.
2578	WLAN profile "[2]" XML data is not UTF-16 encoded. Please, contact your support personnel.
2574	Error [4] deleting profile "[3]" on WLAN interface [2]. Please, contact your support personnel.
2575	Error [5] renaming profile "[3]" to "[4]" on WLAN interface [2]. Please, contact your support personnel.
2576	Error [5] setting profile "[3]" on WLAN interface [2] (WLAN reason is [4]). Please, contact your support personnel.
<<NOKEEP

!IF "$(LANG)" == "en_US"
"en_US.$(PLAT).$(CFG).Error-2.idt" : "en_US.$(PLAT).$(CFG).Error-2.idtx"
	copy /y $** $@ > NUL
!ELSE
"$(LANG).$(PLAT).$(CFG).Error-2.idt" : "en_US.$(PLAT).$(CFG).Error-2.idtx" "..\locale\$(LANG).po"
	rcxgettext.exe idtp $@ $**
!ENDIF


######################################################################
# Build MSM module!
######################################################################

!INCLUDE "..\..\..\MSI\MSIBuild\MSM.mak"
