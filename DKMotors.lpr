program DKMotors;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  {$IFDEF HASAMIGA}
  athreads,
  {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, UMainForm, UImages, lazcontrols, datetimectrls, UReclamationForm,
  UBuildAddForm, USheets, UTestAddForm, UAboutForm, UCargoEditForm,
  UStoreForm, UMotorListForm, UBuildLogForm, UBuildEditForm, UReportForm,
  UTestLogForm, UCargoForm, UReclamationEditForm, UDataBase, UStatisticForm,
  URepairForm, UControlListForm, UCalendar, UCalendarForm, UCalendarEditForm,
  UCardForm, URepairEditForm, UControlListEditForm, UVars,
  UPackingSheetForm;

{$R *.res}

begin
  RequireDerivedFormResource:=True;
  Application.Scaled:=True;
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.

