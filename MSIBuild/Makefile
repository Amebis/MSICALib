#    SPDX-License-Identifier: GPL-3.0-or-later
#    Copyright © 1991-2022 Amebis

!INCLUDE "..\..\..\..\..\include\MSIBuildCfg.mak"

MSIBUILD_IS_LOCALIZEABLE=1
!IF "$(LANG)" != "en_US"
MSIBUILD_PO="..\locale\$(LANG).po"
!ENDIF


######################################################################
# Error

All :: \
	"$(LANG).$(PLAT).$(CFG).Error-1.idt" \
	"$(LANG).$(PLAT).$(CFG).Error-2.idt"

"$(LANG).$(PLAT).$(CFG).Error-1.idt" : "Makefile" "..\..\..\..\..\include\MSIBuildCfg.mak"
	move /y << $@ > NUL
Error	Message
i2	L0
Error	Error
<<NOKEEP

"en_US.$(PLAT).$(CFG).Error-2.idtx" : "Makefile" "..\..\..\..\..\include\MSIBuildCfg.mak"
	move /y << $@ > NUL
Error	Message
i2	L0
Error	Error
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

"$(LANG).$(PLAT).$(CFG).Error-2.idt" : "en_US.$(PLAT).$(CFG).Error-2.idtx" $(MSIBUILD_PO)
	cscript.exe "..\..\..\..\..\MSI\MSIBuild\MSI.wsf" //Job:IDTTranslate //Nologo $@ $** /CP:$(MSIBUILD_CODEPAGE)


######################################################################
# Build MSM module!
######################################################################

!INCLUDE "..\..\..\..\..\MSI\MSIBuild\MSM.mak"
