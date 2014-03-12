!INCLUDE "..\..\include\MSINast.mak"

MSM_IMA_LOKALIZACIJO=1


######################################################################
# Error

Vse :: \
	"$(JEZIK).$(CFG).$(PLAT).Error-1.idt" \
	"$(JEZIK).$(CFG).$(PLAT).Error-2.idt"

"$(JEZIK).$(CFG).$(PLAT).Error-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Error	Message
i2	L0
Error	Error
<<NOKEEP

"Sl.$(CFG).$(PLAT).Error-2.idtx" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Error	Message
i2	L0
1250	Error	Error
2550	Pri odpiranju namestitvenega paketa je prišlo do napake. Obrnite se na svojo tehnièno službo.
2552	Pri pisanju v datoteko seznama opravil »[2]« je prišlo do napake [3]. Obrnite se na svojo tehnièno službo.
2560	Pri branju iz datoteke seznama opravil »[2]« je prišlo do napake [3]. Obrnite se na svojo tehnièno službo.
2553	Pri nastavljanju parametra »[2]« je prišlo do napake [3]. Obrnite se na svojo tehnièno službo.
2554	Pri brisanju datoteke »[2]« je prišlo do napake [3]. Obrnite se na svojo tehnièno službo.
2555	Pri premikanju datoteke »[2]« v »[3]« je prišlo do napake [4]. Obrnite se na svojo tehnièno službo.
2556	Pri ustvarjanju razporejenega opravila »[2]« je prišlo do napake [3]. Obrnite se na svojo tehnièno službo.
2557	Pri brisanju razporejenega opravila »[2]« je prišlo do napake [3]. Obrnite se na svojo tehnièno službo.
2558	Pri o(ne)mogoèanju razporejenega opravila »[2]« je prišlo do napake [3]. Obrnite se na svojo tehnièno službo.
2559	Pri kopiranju razporejenega opravila »[2]« v »[3]« je prišlo do napake [4]. Obrnite se na svojo tehnièno službo.
<<NOKEEP

"De.$(CFG).$(PLAT).Error-2.idt" : "Sl.$(CFG).$(PLAT).Error-2.idtx" "..\res\de_DE.po"
	rcxgettext.exe idtp $@ $**

"En.$(CFG).$(PLAT).Error-2.idt" : "Sl.$(CFG).$(PLAT).Error-2.idtx" "..\res\en_GB.po"
	rcxgettext.exe idtp $@ $**

"It.$(CFG).$(PLAT).Error-2.idt" : "Sl.$(CFG).$(PLAT).Error-2.idtx" "..\res\it_IT.po"
	rcxgettext.exe idtp $@ $**

"Sl.$(CFG).$(PLAT).Error-2.idt" : "Sl.$(CFG).$(PLAT).Error-2.idtx"
	copy /y $** $@ > NUL


######################################################################
# Izdelava modula MSM
######################################################################

!INCLUDE "..\..\AOsn\msm\MSM.mak"
