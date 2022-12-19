unit UMainForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  DividerBevel, DK_LCLStrRus, rxctrls, USQLite3ListForm,
  UBuildLogForm, UShipmentForm, UReclamationForm, UStoreForm, UAboutForm,
  UTestLogForm, UMotorListForm, UReportForm, UStatisticForm, USQLite, SheetUtils;

type

  { TMainForm }

  TMainForm = class(TForm)
    ExitButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    ImageListEdit24: TImageList;
    ImageListCategory24: TImageList;
    ImageList16: TImageList;
    MainPanel: TPanel;
    Panel2: TPanel;
    RxSpeedButton1: TRxSpeedButton;
    RxSpeedButton10: TRxSpeedButton;
    RxSpeedButton2: TRxSpeedButton;
    RxSpeedButton3: TRxSpeedButton;
    RxSpeedButton4: TRxSpeedButton;
    RxSpeedButton5: TRxSpeedButton;
    RxSpeedButton9: TRxSpeedButton;
    AboutButton: TSpeedButton;
    procedure AboutButtonClick(Sender: TObject);
    procedure ExitButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RxSpeedButton10Click(Sender: TObject);
    procedure RxSpeedButton1Click(Sender: TObject);
    procedure RxSpeedButton2Click(Sender: TObject);
    procedure RxSpeedButton3Click(Sender: TObject);
    procedure RxSpeedButton5Click(Sender: TObject);
    procedure RxSpeedButton9Click(Sender: TObject);
  private
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
  public
    BuildLogForm: TBuildLogForm;
    TestLogForm: TTestLogForm;
    ShipmentForm: TShipmentForm;
    ReclamationForm: TReclamationForm;
    StoreForm: TStoreForm;
    MotorListForm: TMotorListForm;
    ReportForm: TReportForm;
    StatisticForm: TStatisticForm;
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
  MotorListFormOpen;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if Assigned(SQLite) then FreeAndNil(SQLite);
  if Assigned(MotorListForm) then FreeAndNil(MotorListForm);
  FreeForms;
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

procedure TMainForm.RxSpeedButton1Click(Sender: TObject);
begin
  Choose;
end;

procedure TMainForm.RxSpeedButton2Click(Sender: TObject);
begin
  Choose;
end;

procedure TMainForm.RxSpeedButton3Click(Sender: TObject);
begin
  Choose;
end;

procedure TMainForm.RxSpeedButton5Click(Sender: TObject);
begin
  Choose;
end;

procedure TMainForm.RxSpeedButton9Click(Sender: TObject);
begin
  Choose;
end;

procedure TMainForm.RxSpeedButton10Click(Sender: TObject);
begin
  Choose;
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
    if RxSpeedButton1.Down       then BuildLogFormOpen
    else if RxSpeedButton2.Down  then ShipmentFormOpen
    else if RxSpeedButton3.Down  then ReclamationFormOpen
    else if RxSpeedButton4.Down  then StoreFormOpen
    else if RxSpeedButton5.Down  then TestFormOpen
    else if RxSpeedButton9.Down  then ReportFormOpen
    else if RxSpeedButton10.Down then StatisticFormOpen;
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


end.

