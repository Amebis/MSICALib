!INCLUDE "..\..\include\MSINast.mak"

#MSM_IMA_LOKALIZACIJO=1


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
ScheduledTask	Application	N					Formatted		Task application name.
ScheduledTask	Parameters	N					Formatted		Task command line parameters.
ScheduledTask	WorkingDir	N					Formatted		Task working directory. If the working directory is set to "", when the application is run, the current directory will be the directory in which the task scheduler service executable, Mstask.exe, resides.
ScheduledTask	Flags	N					DoubleInteger		Task's flags to pass in IScheduledWorkItem::SetFlags() call.
ScheduledTask	Priority	N					DoubleInteger	32;64;128;256	Task's priority to pass in IScheduledWorkItem::SetPriority() call.
ScheduledTask	User	N					Text		Account name to run task as. Use empty string for system account.
ScheduledTask	Password	Y					Text		Account password to run task as. Use NULL for system account.
ScheduledTask	Comment	Y					Text		Task description.
ScheduledTask	IdleMin	Y					Integer		A value that specifies how long, in minutes, the system must remain idle before the work item can run.
ScheduledTask	IdleDeadline	Y					Integer		A value that specifies the maximum number of minutes that the Task Scheduler will wait for the idle-time period.
ScheduledTask	MaxRuntime	N					DoubleInteger		Specifies the maximum run time (in milliseconds), for the task. This parameter may be set to -1 to specify an unlimited time.
ScheduledTask	Condition	Y					Condition		Optional expression which skips the task if evaluates to expFalse. If the expression syntax is invalid, the engine will terminate, returning iesBadActionData.
ScheduledTask	Component_	N			Component	1	Identifier		Component the task is part of.
TaskTrigger	Trigger	N					Identifier		Primary key, non-localized token.
TaskTrigger	BeginDate	N					Date/Time		Date that the task trigger activates. The beginning date must be specified when setting a task.
TaskTrigger	EndDate	Y					Date/Time		Date that the task trigger deactivates.
TaskTrigger	StartTime	N					Date/Time		Time of the day the task runs.
TaskTrigger	MinutesDuration	Y					DoubleInteger		Number of minutes after the task starts that the trigger will remain active. The number of minutes specified here must be greater than or equal to the MinutesInterval column.
TaskTrigger	MinutesInterval	Y					DoubleInteger		Number of minutes between consecutive task executions. This number is counted from the start of the previous scheduled task. The number of minutes specified here must be less than the MinutesDuration column.
TaskTrigger	Flags	N					DoubleInteger		Value that describes the behavior of the trigger. This value is a combination of the flags in TASK_TRIGGER::rgFlags.
TaskTrigger	Type	N	0	7			Integer		A TASK_TRIGGER_TYPE enumerated value that specifies the type of trigger. This member is used with Type. The type of trigger specified here determines which fields of the TRIGGER_TYPE_UNION specified in Type member will be used. Trigger type is based on when the trigger will run the task.
TaskTrigger	DaysInterval	Y					Integer		Specifies the number of days between task runs.
TaskTrigger	WeeksInterval	Y					Integer		Number of weeks between invocations of a task.
TaskTrigger	DaysOfTheWeek	Y	0	127			Integer		Value that describes the days of the week the task runs. This value is a bitfield and is a combination of the following flags: 1=Sun, 2=Mon, 4=Tue, 8=Wed, 16=Thu, 32=Fri, 64=Sat.
TaskTrigger	DaysOfMonth	Y					DoubleInteger		Specifies the day of the month a task runs. This value is a bitfield that specifies the day(s) the task will run. Bit 0 corresponds to the first of the month, bit 1 to the second, and so forth.
TaskTrigger	WeekOfMonth	Y					Integer	1;2;3;4;5	Specifies the week of the month when the task runs. This value is exclusive and is one of the following flags: 1=The task will run between the first and seventh day of the month, 2=The task will run between the eighth and 14th day of the month, 3=The task will run between the 15th and 21st day of the month, 4=The task will run between the 22nd and 28th of the month, 5=The task will run between the last seven days of the month.
TaskTrigger	MonthsOfYear	Y	0	4095			Integer		Specifies the month(s) when the task runs. This value is a combination of the following flags: 1=Jan, 2=Feb, 4=Mar, 8=Apr, 16=May, 32=Jun, 64=Jul, 128=Aug, 256=Sep, 512=Oct, 1024=Nov, 2048=Dec.
TaskTrigger	Task_	N			ScheduledTask	1	Identifier		Key of the Trigger's task.
<<NOKEEP


######################################################################
# ScheduledTask

Vse :: "$(JEZIK).$(CFG).$(PLAT).ScheduledTask-1.idt"

"$(JEZIK).$(CFG).$(PLAT).ScheduledTask-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Task	Application	Parameters	WorkingDir	Flags	Priority	User	Password	Comment	IdleMin	IdleDeadline	MaxRuntime	Condition	Component_
s$(MSI_TIP_ID)	s255	s255	s255	i4	i4	s255	S255	S255	I2	I2	i4	S255	s$(MSI_TIP_ID)
ScheduledTask	Task
<<NOKEEP


######################################################################
# TaskTrigger

Vse :: "$(JEZIK).$(CFG).$(PLAT).TaskTrigger-1.idt"

"$(JEZIK).$(CFG).$(PLAT).TaskTrigger-1.idt" : "Makefile" "..\..\include\MSINast.mak"
	-if exist $@ del /f /q $@
	move /y << $@ > NUL
Trigger	BeginDate	EndDate	StartTime	MinutesDuration	MinutesInterval	Flags	Type	DaysInterval	WeeksInterval	DaysOfTheWeek	DaysOfMonth	WeekOfMonth	MonthsOfYear	Task_
s$(MSI_TIP_ID)	i2	I2	i2	I4	I4	i4	i2	I2	I2	I2	I4	I2	I2	s$(MSI_TIP_ID)
TaskTrigger	Trigger
<<NOKEEP


######################################################################
# Izdelava modula MSM
######################################################################

!INCLUDE "..\..\AOsn\AOsnMSI\MSM.mak"
