!INCLUDE "..\..\include\MSINast.mak"

MSM_IMA_LOKALIZACIJO=1
!IF "$(CFG)" == "ReleaseU"
CFG_VC=Unicode Release
!ELSEIF "$(CFG)" == "DebugU"
CFG_VC=Unicode Debug
!ELSE
CFG_VC=$(CFG)
!ENDIF


######################################################################
# _Validation

Vse :: "$(JEZIK).$(CFG).$(PLAT)._Validation-1.idt"

"$(JEZIK).$(CFG).$(PLAT)._Validation-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Table	Column	Nullable	MinValue	MaxValue	KeyTable	KeyColumn	Category	Set	Description
s32	s32	s4	I4	I4	S255	I2	S32	S255	S255
_Validation	Table	Column
ScheduledTask	Task	N					Identifier		Primary key, non-localized token.
ScheduledTask	DisplayName	N					Formatted		Task name.
ScheduledTask	Application	N					Formatted		Task application name.
ScheduledTask	Parameters	Y					Formatted		Task command line parameters.
ScheduledTask	WorkingDir	N					Formatted		Task working directory. If the working directory is set to "", when the application is run, the current directory will be the directory in which the task scheduler service executable, Mstask.exe, resides.
ScheduledTask	Flags	N					DoubleInteger		Task's flags to pass in IScheduledWorkItem::SetFlags() call.
ScheduledTask	Priority	N					DoubleInteger	32;64;128;256;16384,32768	Task's priority to pass in ITask::SetPriority() call.
ScheduledTask	User	Y					Formatted		Account name to run task as. Use empty string for system account.
ScheduledTask	Password	Y					Formatted		Account password to run task as. Use NULL for system account.
ScheduledTask	Description	Y					Formatted		Task description.
ScheduledTask	IdleMin	Y					Integer		A value that specifies how long, in minutes, the system must remain idle before the work item can run.
ScheduledTask	IdleDeadline	Y					Integer		A value that specifies the maximum number of minutes that the Task Scheduler will wait for the idle-time period.
ScheduledTask	MaxRuntime	N					DoubleInteger		Specifies the maximum run time (in milliseconds), for the task. This parameter may be set to -1 to specify an unlimited time.
ScheduledTask	Condition	Y					Condition		Optional expression which skips the task if evaluates to expFalse. If the expression syntax is invalid, the engine will terminate, returning iesBadActionData.
ScheduledTask	Component_	N			Component	1	Identifier		Component the task is part of.
TaskTrigger	Trigger	N					Identifier		Primary key, non-localized token.
TaskTrigger	BeginDate	N					DoubleInteger		Date that the task trigger activates (counted in days from January 1st, 1980). The beginning date must be specified when setting a task.
TaskTrigger	EndDate	Y					DoubleInteger		Date that the task trigger deactivates (counted in days from January 1st, 1980).
TaskTrigger	StartTime	N					DoubleInteger		Time of the day the task runs (counted in minutes since midnight).
TaskTrigger	StartTimeRand	Y					Integer		Maximum random number of minutes to add to start time.
TaskTrigger	MinutesDuration	Y					DoubleInteger		Number of minutes after the task starts that the trigger will remain active. The number of minutes specified here must be greater than or equal to the MinutesInterval column.
TaskTrigger	MinutesInterval	Y					DoubleInteger		Number of minutes between consecutive task executions. This number is counted from the start of the previous scheduled task. The number of minutes specified here must be less than the MinutesDuration column.
TaskTrigger	Flags	N					DoubleInteger		Value that describes the behavior of the trigger. This value is a combination of the flags in TASK_TRIGGER::rgFlags.
TaskTrigger	Type	N	0	7			Integer		A TASK_TRIGGER_TYPE enumerated value that specifies the type of trigger. This member is used with Type. The type of trigger specified here determines which fields of the TRIGGER_TYPE_UNION specified in Type member will be used. Trigger type is based on when the trigger will run the task.
TaskTrigger	DaysInterval	Y					Integer		Specifies the number of days between task runs.
TaskTrigger	WeeksInterval	Y					Integer		Number of weeks between invocations of a task.
TaskTrigger	DaysOfTheWeek	Y	0	127			Integer		Value that describes the days of the week the task runs. This value is a bitfield and is a combination of the following values: 1=Sun, 2=Mon, 4=Tue, 8=Wed, 16=Thu, 32=Fri, 64=Sat.
TaskTrigger	DaysOfMonth	Y					DoubleInteger		Specifies the day of the month a task runs. This value is a bitfield that specifies the day(s) the task will run. Bit 0 corresponds to the first of the month, bit 1 to the second, and so forth.
TaskTrigger	WeekOfMonth	Y					Integer	1;2;3;4;5	Specifies the week of the month when the task runs. This value is exclusive and is one of the following: 1=The task will run between the first and seventh day of the month, 2=The task will run between the eighth and 14th day of the month, 3=The task will run between the 15th and 21st day of the month, 4=The task will run between the 22nd and 28th of the month, 5=The task will run between the last seven days of the month.
TaskTrigger	MonthsOfYear	Y	0	4095			Integer		Specifies the month(s) when the task runs. This value is a combination of the following values: 1=Jan, 2=Feb, 4=Mar, 8=Apr, 16=May, 32=Jun, 64=Jul, 128=Aug, 256=Sep, 512=Oct, 1024=Nov, 2048=Dec.
TaskTrigger	Task_	N			ScheduledTask	1	Identifier		Key of the Trigger's task.
<<NOKEEP


######################################################################
# ActionText

Vse :: "$(JEZIK).$(CFG).$(PLAT).ActionText-2.idt"

"Sl.$(CFG).$(PLAT).ActionText-2.idtx" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Action	Description	Template
s$(MSI_TIP_ID)	L0	L0
1250	ActionText	Action
InstallScheduledTasks	Registracija razporejenih opravil	Razporejeno opravilo: [1]
<<NOKEEP

"De.$(CFG).$(PLAT).ActionText-2.idt" : "Sl.$(CFG).$(PLAT).ActionText-2.idtx" "..\res\de_DE.po"
	rcxgettext.exe idtp $@ $**

"En.$(CFG).$(PLAT).ActionText-2.idt" : "Sl.$(CFG).$(PLAT).ActionText-2.idtx" "..\res\en_GB.po"
	rcxgettext.exe idtp $@ $**

"It.$(CFG).$(PLAT).ActionText-2.idt" : "Sl.$(CFG).$(PLAT).ActionText-2.idtx" "..\res\it_IT.po"
	rcxgettext.exe idtp $@ $**

"Sl.$(CFG).$(PLAT).ActionText-2.idt" : "Sl.$(CFG).$(PLAT).ActionText-2.idtx"
	copy /y $** $@ > NUL


######################################################################
# Binary

Vse :: \
	"$(JEZIK).$(CFG).$(PLAT).Binary-1.idt" \
	"$(JEZIK).$(CFG).$(PLAT).Binary-1\MSITSCA.dll"

"$(JEZIK).$(CFG).$(PLAT).Binary-1.idt" : "Makefile" "..\..\include\MSINast.mak" "$(JEZIK).$(CFG).$(PLAT).Binary-1"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Name	Data
s$(MSI_TIP_ID)	v0
Binary	Name
binMSITSCA.dll	MSITSCA.dll
<<NOKEEP

"$(JEZIK).$(CFG).$(PLAT).Binary-1\MSITSCA.dll" : "..\..\Test\$(CFG_VC).$(PLAT)\MSITSCA.dll"
	copy /y $** $@ > NUL


######################################################################
# CustomAction

Vse :: "$(JEZIK).$(CFG).$(PLAT).CustomAction-1.idt"

"$(JEZIK).$(CFG).$(PLAT).CustomAction-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Action	Type	Source	Target
s$(MSI_TIP_ID)	i2	S$(MSI_TIP_ID)	S255
CustomAction	Action
EvaluateScheduledTasks	1	binMSITSCA.dll	EvaluateScheduledTasks
InstallScheduledTasks	1025	binMSITSCA.dll	InstallScheduledTasks
RollbackScheduledTasks	1281	binMSITSCA.dll	InstallScheduledTasks
CommitScheduledTasks	1537	binMSITSCA.dll	InstallScheduledTasks
<<NOKEEP


######################################################################
# Error

Vse :: "$(JEZIK).$(CFG).$(PLAT).Error-2.idt"

"Sl.$(CFG).$(PLAT).Error-2.idtx" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Error	Message
i2	L0
1250	Error	Error
2550	Pri odpiranju namestitvenega paketa je prišlo do napake.
2551	Pri pripravi seznama razvršèenih opravil je prišlo do napake [2].
2552	Pri pisanju v datoteko seznama razporejenih opravil »[2]« je prišlo do napake [3].
2553	Pri nastavljanju parametra »[2]« je prišlo do napake [3].
2554	Pri brisanju datoteke »[2]« je prišlo do napake [3].
2555	Pri premikanju datoteke »[2]« v »[3]« je prišlo do napake [4].
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
# InstallExecuteSequence

Vse :: "$(JEZIK).$(CFG).$(PLAT).InstallExecuteSequence-1.idt"

"$(JEZIK).$(CFG).$(PLAT).InstallExecuteSequence-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Action	Condition	Sequence
s$(MSI_TIP_ID)	S255	I2
InstallExecuteSequence	Action
EvaluateScheduledTasks		6580
InstallScheduledTasks		6581
RollbackScheduledTasks		6582
CommitScheduledTasks		6583
<<NOKEEP


######################################################################
# ScheduledTask

Vse :: "$(JEZIK).$(CFG).$(PLAT).ScheduledTask-1.idt"

"$(JEZIK).$(CFG).$(PLAT).ScheduledTask-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Task	DisplayName	Application	Parameters	WorkingDir	Flags	Priority	User	Password	Description	IdleMin	IdleDeadline	MaxRuntime	Condition	Component_
s$(MSI_TIP_ID)	s255	s255	S255	s255	i4	i4	S255	S255	S255	I2	I2	i4	S255	s$(MSI_TIP_ID)
ScheduledTask	Task
<<NOKEEP


######################################################################
# TaskTrigger

Vse :: "$(JEZIK).$(CFG).$(PLAT).TaskTrigger-1.idt"

"$(JEZIK).$(CFG).$(PLAT).TaskTrigger-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Trigger	BeginDate	EndDate	StartTime	StartTimeRand	MinutesDuration	MinutesInterval	Flags	Type	DaysInterval	WeeksInterval	DaysOfTheWeek	DaysOfMonth	WeekOfMonth	MonthsOfYear	Task_
s$(MSI_TIP_ID)	i2	I2	i2	I2	I4	I4	i4	i2	I2	I2	I2	I4	I2	I2	s$(MSI_TIP_ID)
TaskTrigger	Trigger
<<NOKEEP


######################################################################
# Izdelava modula MSM
######################################################################

!INCLUDE "..\..\AOsn\AOsnMSI\MSM.mak"
