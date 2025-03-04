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
  Forms, UMainForm, lazcontrols, datetimectrls, UReclamationForm,
  UBuildAddForm, USheetUtils, UTestAddForm, UAboutForm, UCargoEditForm,
  UStoreForm, UMotorListForm, UBuildLogForm, UBuildEditForm, UReportForm,
  UTestLogForm, UShipmentForm, UReclamationEditForm, USQLite, UStatisticForm,
  URepairForm, UControlListForm, UCalendar, UCalendarForm, UCalendarEditForm,
  UCardForm, URepairEditForm, UControlListEditForm, UUtils, UPackingSheetForm;

{$R *.res}

begin
  RequireDerivedFormResource:=True;
  Application.Scaled:=True;
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.

