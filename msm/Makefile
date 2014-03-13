!INCLUDE "..\..\include\MSIBuildCfg.mak"

MSIBUILD_IS_LOCALIZEABLE=1


######################################################################
# Error

All :: \
	"$(LANG).$(PLAT).$(CFG).Error-1.idt" \
	"$(LANG).$(PLAT).$(CFG).Error-2.idt"

"$(LANG).$(PLAT).$(CFG).Error-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Error	Message
i2	L0
Error	Error
<<NOKEEP

"En.$(PLAT).$(CFG).Error-2.idtx" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Error	Message
i2	L0
1252	Error	Error
2550	Error opening installation package. Please, contact your support personnel.
2551	Error [2] creating scheduled task list. Please, contact your support personnel.
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
<<NOKEEP

"De.$(CFG).$(PLAT).Error-2.idt" : "En.$(PLAT).$(CFG).Error-2.idtx" "..\L10N\de_DE.po"
	rcxgettext.exe idtp $@ $**

"En.$(PLAT).$(CFG).Error-2.idt" : "En.$(PLAT).$(CFG).Error-2.idtx"
	copy /y $** $@ > NUL

"It.$(CFG).$(PLAT).Error-2.idt" : "En.$(PLAT).$(CFG).Error-2.idtx" "..\L10N\it_IT.po"
	rcxgettext.exe idtp $@ $**

"Sl.$(CFG).$(PLAT).Error-2.idt" : "En.$(PLAT).$(CFG).Error-2.idtx" "..\L10N\sl_SI.po"
	rcxgettext.exe idtp $@ $**


######################################################################
# Build MSM module!
######################################################################

!INCLUDE "..\..\MSIBuild\MSM.mak"
