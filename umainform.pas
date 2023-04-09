unit UMainForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, DividerBevel, DK_LCLStrRus, DK_Vector, rxctrls, USQLite3ListForm,
  UBuildLogForm, UShipmentForm, UReclamationForm, UStoreForm, UAboutForm,
  UTestLogForm, UMotorListForm, UReportForm, UStatisticForm, URepairForm,
  UControlListForm, UCalendarForm, USQLite, USheetUtils;

type

  { TMainForm }

  TMainForm = class(TForm)
    ControlListButton: TRxSpeedButton;
    DividerBevel5: TDividerBevel;
    DividerBevel6: TDividerBevel;
    DividerBevel7: TDividerBevel;
    DividerBevel8: TDividerBevel;
    MotorListButton: TRxSpeedButton;
    ChooseMotorNamesButton: TSpeedButton;
    ChooseRecieverNamesButton: TSpeedButton;
    ExitButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    ImageListEdit24: TImageList;
    ImageListCategory24: TImageList;
    ImageList16: TImageList;
    Label2: TLabel;
    Label3: TLabel;
    MainPanel: TPanel;
    MotorNamesLabel: TLabel;
    MotorNamesPanel: TPanel;
    RepairButton: TRxSpeedButton;
    CalendarButton: TRxSpeedButton;
    ToolPanel: TPanel;
    ReceiverNamesLabel: TLabel;
    ReceiverNamesPanel: TPanel;
    BuildLogButton: TRxSpeedButton;
    StatisticButton: TRxSpeedButton;
    ShipmentButton: TRxSpeedButton;
    ReclamationButton: TRxSpeedButton;
    StoreButton: TRxSpeedButton;
    TestLogButton: TRxSpeedButton;
    ReportButton: TRxSpeedButton;
    AboutButton: TSpeedButton;
    procedure AboutButtonClick(Sender: TObject);
    procedure CalendarButtonClick(Sender: TObject);
    procedure ChooseMotorNamesButtonClick(Sender: TObject);
    procedure ChooseRecieverNamesButtonClick(Sender: TObject);
    procedure ControlListButtonClick(Sender: TObject);
    procedure ExitButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Label2Click(Sender: TObject);
    procedure Label3Click(Sender: TObject);
    procedure MotorListButtonClick(Sender: TObject);
    procedure MotorNamesLabelClick(Sender: TObject);
    procedure MotorNamesPanelClick(Sender: TObject);
    procedure ReceiverNamesLabelClick(Sender: TObject);
    procedure ReceiverNamesPanelClick(Sender: TObject);
    procedure RepairButtonClick(Sender: TObject);
    procedure StatisticButtonClick(Sender: TObject);
    procedure BuildLogButtonClick(Sender: TObject);
    procedure ShipmentButtonClick(Sender: TObject);
    procedure ReclamationButtonClick(Sender: TObject);
    procedure StoreButtonClick(Sender: TObject);
    procedure TestLogButtonClick(Sender: TObject);
    procedure ReportButtonClick(Sender: TObject);
  private
    BuildLogForm: TBuildLogForm;
    TestLogForm: TTestLogForm;
    ShipmentForm: TShipmentForm;
    ReclamationForm: TReclamationForm;
    StoreForm: TStoreForm;
    MotorListForm: TMotorListForm;
    ReportForm: TReportForm;
    StatisticForm: TStatisticForm;
    RepairForm: TRepairForm;
    ControlListForm: TControlListForm;
    CalendarForm: TCalendarForm;

    procedure DBConnect;
    procedure Choose;
    procedure FreeForms;
    procedure SetFormPosition(AForm: TForm);
    procedure BuildLogFormOpen;
    procedure TestFormOpen;
    procedure ShipmentFormOpen;
    procedure ReclamationFormOpen;
    procedure StoreFormOpen;
    procedure MotorListFormOpen;
    procedure ReportFormOpen;
    procedure StatisticFormOpen;
    procedure RepairFormOpen;
    procedure ControlListFormOpen;
    procedure CalendarFormOpen;

    procedure ChangeUsedMotorList;
    procedure ChangeUsedReceiverList;

    procedure ShowData;
  public


    UsedNameIDs: TIntVector;
    UsedNames: TStrVector;

    UsedReceiverIDs: TIntVector;
    UsedReceiverNames: TStrVector;

    procedure SetNamesPanelsVisible(const AMotorNamesPanelVisible,
                                          AReceiverNamesPanelVisible: Boolean);
  end;

var
  MainForm: TMainForm;


implementation

{$R *.lfm}

procedure TMainForm.FormCreate(Sender: TObject);
begin
  USQLite3ListForm.ImageList:= ImageListEdit24;
  SQLite:= TSQLite.Create;
  SQLite.SetEditListSettings(COLOR_BACKGROUND_SELECTED, clWindowText);
  DBConnect;
  SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, False, UsedNameIDs, UsedNames);
  SQLite.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, False, UsedReceiverIDs, UsedReceiverNames);
  Choose;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if Assigned(SQLite) then FreeAndNil(SQLite);
  FreeForms;
end;

procedure TMainForm.ChooseMotorNamesButtonClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TMainForm.Label2Click(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TMainForm.MotorNamesLabelClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TMainForm.MotorNamesPanelClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TMainForm.ChooseRecieverNamesButtonClick(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.ControlListButtonClick(Sender: TObject);
begin
  if not Assigned(ControlListForm) then Choose;
end;

procedure TMainForm.Label3Click(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.MotorListButtonClick(Sender: TObject);
begin
  if not Assigned(MotorListForm) then Choose;
end;

procedure TMainForm.ReceiverNamesLabelClick(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.ReceiverNamesPanelClick(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.RepairButtonClick(Sender: TObject);
begin
  if not Assigned(RepairForm) then Choose;
end;

procedure TMainForm.ExitButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.AboutButtonClick(Sender: TObject);
var
  AboutForm: TAboutForm;
begin
  AboutForm:= TAboutForm.Create(MainForm);
  AboutForm.ShowModal;
  FreeAndNil(AboutForm);
end;

procedure TMainForm.CalendarButtonClick(Sender: TObject);
begin
  if not Assigned(CalendarForm) then Choose;
end;

procedure TMainForm.BuildLogButtonClick(Sender: TObject);
begin
  if not Assigned(BuildLogForm) then Choose;
end;

procedure TMainForm.ShipmentButtonClick(Sender: TObject);
begin
  if not Assigned(ShipmentForm) then Choose;
end;

procedure TMainForm.ReclamationButtonClick(Sender: TObject);
begin
  if not Assigned(ReclamationForm) then Choose;
end;

procedure TMainForm.StoreButtonClick(Sender: TObject);
begin
  if not Assigned(StoreForm) then Choose;
end;

procedure TMainForm.TestLogButtonClick(Sender: TObject);
begin
  if not Assigned(TestLogForm) then Choose;
end;

procedure TMainForm.ReportButtonClick(Sender: TObject);
begin
  if not Assigned(ReportForm) then Choose;
end;

procedure TMainForm.StatisticButtonClick(Sender: TObject);
begin
  if not Assigned(StatisticForm) then Choose;
end;

procedure TMainForm.DBConnect;
var
  ProgPath, DBName, DDLName: String;
begin
  ProgPath:= ExtractFilePath(Application.ExeName);
  DBName:= ProgPath + 'DKMotors.db';
  DDLName:= ProgPath + 'DDL.sql';
  SQLite.Connect(DBName);
  SQLite.ExecuteScript(DDLName);
end;

procedure TMainForm.Choose;
begin
  Screen.Cursor:= crHourGlass;
  try
    FreeForms;
    if BuildLogButton.Down       then BuildLogFormOpen
    else if ShipmentButton.Down  then ShipmentFormOpen
    else if ReclamationButton.Down  then ReclamationFormOpen
    else if StoreButton.Down  then StoreFormOpen
    else if TestLogButton.Down  then TestFormOpen
    else if ReportButton.Down  then ReportFormOpen
    else if StatisticButton.Down then StatisticFormOpen
    else if MotorListButton.Down then MotorListFormOpen
    else if RepairButton.Down then RepairFormOpen
    else if ControlListButton.Down then ControlListFormOpen
    else if CalendarButton.Down then CalendarFormOpen;
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TMainForm.FreeForms;
begin
  if Assigned(BuildLogForm) then FreeAndNil(BuildLogForm);
  if Assigned(TestLogForm) then FreeAndNil(TestLogForm);
  if Assigned(ShipmentForm) then FreeAndNil(ShipmentForm);
  if Assigned(ReclamationForm) then FreeAndNil(ReclamationForm);
  if Assigned(StoreForm) then FreeAndNil(StoreForm);
  if Assigned(ReportForm) then FreeAndNil(ReportForm);
  if Assigned(StatisticForm) then FreeAndNil(StatisticForm);
  if Assigned(MotorListForm) then FreeAndNil(MotorListForm);
  if Assigned(RepairForm) then FreeAndNil(RepairForm);
  if Assigned(ControlListForm) then FreeAndNil(ControlListForm);
  if Assigned(CalendarForm) then FreeAndNil(CalendarForm);
end;

procedure TMainForm.SetFormPosition(AForm: TForm);
begin
  AForm.Parent:= MainPanel;
  AForm.Left:= 0;
  AForm.Top:= 0;
  AForm.MakeFullyVisible();
end;

procedure TMainForm.BuildLogFormOpen;
begin
  BuildLogForm:= TBuildLogForm.Create(MainForm);
  SetFormPosition(TForm(BuildLogForm));
  BuildLogForm.Show;
end;

procedure TMainForm.TestFormOpen;
begin
  TestLogForm:= TTestLogForm.Create(MainForm);
  SetFormPosition(TForm(TestLogForm));
  TestLogForm.Show;
end;

procedure TMainForm.ShipmentFormOpen;
begin
  ShipmentForm:= TShipmentForm.Create(MainForm);
  SetFormPosition(TForm(ShipmentForm));
  ShipmentForm.Show;
end;

procedure TMainForm.ReclamationFormOpen;
begin
  ReclamationForm:= TReclamationForm.Create(MainForm);
  SetFormPosition(TForm(ReclamationForm));
  ReclamationForm.Show;
end;

procedure TMainForm.StoreFormOpen;
begin
  StoreForm:= TStoreForm.Create(MainForm);
  SetFormPosition(TForm(StoreForm));
  StoreForm.Show;
end;

procedure TMainForm.MotorListFormOpen;
begin
  MotorListForm:= TMotorListForm.Create(MainForm);
  SetFormPosition(TForm(MotorListForm));
  MotorListForm.Show;
end;

procedure TMainForm.ReportFormOpen;
begin
  ReportForm:= TReportForm.Create(MainForm);
  SetFormPosition(TForm(ReportForm));
  ReportForm.Show;
end;

procedure TMainForm.StatisticFormOpen;
begin
  StatisticForm:= TStatisticForm.Create(MainForm);
  SetFormPosition(TForm(StatisticForm));
  StatisticForm.Show;
end;

procedure TMainForm.RepairFormOpen;
begin
  RepairForm:= TRepairForm.Create(MainForm);
  SetFormPosition(TForm(RepairForm));
  RepairForm.Show;
end;

procedure TMainForm.ControlListFormOpen;
begin
  ControlListForm:= TControlListForm.Create(MainForm);
  SetFormPosition(TForm(ControlListForm));
  ControlListForm.Show;
end;

procedure TMainForm.CalendarFormOpen;
begin
  CalendarForm:= TCalendarForm.Create(MainForm);
  SetFormPosition(TForm(CalendarForm));
  CalendarForm.Show;
end;

procedure TMainForm.ChangeUsedMotorList;
begin
  if SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, True,
    UsedNameIDs, UsedNames) then ShowData;
end;

procedure TMainForm.ChangeUsedReceiverList;
begin
  if SQLite.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, True,
    UsedReceiverIDs, UsedReceiverNames) then ShowData;
end;

procedure TMainForm.ShowData;
begin
  if BuildLogButton.Down         then BuildLogForm.ShowBuildLog
  else if ShipmentButton.Down    then ShipmentForm.ShowShipment
  else if ReclamationButton.Down then ReclamationForm.ShowReclamation
  else if StoreButton.Down       then StoreForm.ShowStore
  else if TestLogButton.Down     then TestLogForm.ShowTestLog
  else if ReportButton.Down      then ReportForm.ShowReport
  else if StatisticButton.Down   then StatisticForm.ShowStatistic
  else if MotorListButton.Down   then MotorListForm.ShowMotorList
  else if RepairButton.Down      then RepairForm.ShowRepair
  else if ControlListButton.Down then ControlListForm.ShowControlList;
end;

procedure TMainForm.SetNamesPanelsVisible(const AMotorNamesPanelVisible,
  AReceiverNamesPanelVisible: Boolean);
begin
  MotorNamesPanel.Visible:= AMotorNamesPanelVisible;
  ReceiverNamesPanel.Visible:= AReceiverNamesPanelVisible;
end;


end.

