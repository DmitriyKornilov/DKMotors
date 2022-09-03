unit UReportForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, rxctrls, fpspreadsheetgrid, DividerBevel, DateTimePicker,
  DK_DateUtils, SheetUtils, USQLite, DK_SheetExporter, fpstypes, DK_Vector,
  DK_Matrix;

type

  { TReportForm }

  TReportForm = class(TForm)
    ChooseMotorNamesButton: TSpeedButton;
    ChooseRecieverNamesButton: TSpeedButton;
    CloseButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    Label2: TLabel;
    Label3: TLabel;
    MotorNamesLabel: TLabel;
    ReceiverNamesLabel: TLabel;
    MotorNamesPanel: TPanel;
    ReceiverNamesPanel: TPanel;
    OrderNumCheckBox: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Label1: TLabel;
    LogGrid: TsWorksheetGrid;
    Panel1: TPanel;
    Panel2: TPanel;
    ReceiverPanel: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    ReportPeriodPanel: TPanel;
    ToolPanel: TPanel;
    procedure ChooseMotorNamesButtonClick(Sender: TObject);
    procedure ChooseRecieverNamesButtonClick(Sender: TObject);
    procedure OrderNumCheckBoxChange(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);

    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure RadioButton3Click(Sender: TObject);
    procedure RadioButton4Click(Sender: TObject);

  private
    NameIDs, ReceiverIDs: TIntVector;
    MotorBuildSheet: TMotorBuildSheet;
    MotorTestSheet: TMotorTestSheet;
    ReportShipmentSheet: TReportShipmentSheet;
    ReportReclamationSheet: TReportReclamationSheet;

    UsedNameIDs: TIntVector;
    UsedNames: TStrVector;

    UsedReceiverIDs: TIntVector;
    UsedReceiverNames: TStrVector;

    procedure ShowReport;
    procedure ShowBuildReport;
    procedure ShowTestReport;
    procedure ShowShipmentReport;
    procedure ShowReclamationReport;

    procedure FreeSheets;
  public

  end;

var
  ReportForm: TReportForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TReportForm }

procedure TReportForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton9.Down:= False;
  MainForm.ReportForm:= nil;
  CloseAction:= caFree;
end;

procedure TReportForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(LogGrid);
  try
    //Exporter.SheetName:= 'Отчет';
    Exporter.PageSettings(spoPortrait, pfWidth);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TReportForm.DateTimePicker2Change(Sender: TObject);
begin
  ShowReport;
end;

procedure TReportForm.DateTimePicker1Change(Sender: TObject);
begin
  ShowReport;
end;

procedure TReportForm.OrderNumCheckBoxChange(Sender: TObject);
begin
  ShowReport;
end;

procedure TReportForm.ChooseMotorNamesButtonClick(Sender: TObject);
begin
   if SQLite.EditKeyPickList(UsedNameIDs, UsedNames, 'Наименования электродвигателей',
    'MOTORNAMES', 'NameID', 'MotorName', False, True) then  ShowReport;
  MotorNamesLabel.Caption:= VVectorToStr(UsedNames, ', ');
  MotorNamesLabel.Hint:= MotorNamesLabel.Caption;
end;

procedure TReportForm.ChooseRecieverNamesButtonClick(Sender: TObject);
begin
  if SQLite.EditKeyPickList(UsedReceiverIDs, UsedReceiverNames,
    'Наименования грузополучателей',
    'CARGORECEIVERS', 'ReceiverID', 'ReceiverName', True, True) then  ShowReport;
  ReceiverNamesLabel.Caption:= VVectorToStr(UsedReceiverNames, ', ');
  ReceiverNamesLabel.Hint:= ReceiverNamesLabel.Caption;
end;

procedure TReportForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TReportForm.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= FirstDayInMonth(Date);

  SQLite.KeyPickList('MOTORNAMES', 'NameID', 'MotorName',
                     UsedNameIDs, UsedNames, True, 'NameID');
  MotorNamesLabel.Caption:= VVectorToStr(UsedNames, ', ');
  MotorNamesLabel.Hint:= MotorNamesLabel.Caption;

  SQLite.KeyPickList('CARGORECEIVERS', 'ReceiverID', 'ReceiverName',
                     UsedReceiverIDs, UsedReceiverNames, True, 'ReceiverName');
  ReceiverNamesLabel.Caption:= VVectorToStr(UsedReceiverNames, ', ');
  ReceiverNamesLabel.Hint:= ReceiverNamesLabel.Caption;



end;


procedure TReportForm.FormDestroy(Sender: TObject);
begin
  FreeSheets;
end;

procedure TReportForm.FormShow(Sender: TObject);
begin
  ShowReport;
end;



procedure TReportForm.RadioButton1Click(Sender: TObject);
begin
  ShowReport;
end;

procedure TReportForm.RadioButton2Click(Sender: TObject);
begin
  ShowReport;
end;

procedure TReportForm.RadioButton3Click(Sender: TObject);
begin
  ShowReport;
end;

procedure TReportForm.RadioButton4Click(Sender: TObject);
begin
  ShowReport;
end;

procedure TReportForm.FreeSheets;
begin
  if Assigned(MotorBuildSheet) then FreeAndNil(MotorBuildSheet);
  if Assigned(MotorTestSheet) then FreeAndNil(MotorTestSheet);
  if Assigned(ReportShipmentSheet) then FreeAndNil(ReportShipmentSheet);
  if Assigned(ReportReclamationSheet) then FreeAndNil(ReportReclamationSheet);
end;

procedure TReportForm.ShowReport;
begin
  ExportButton.Align:= alRight;
  DividerBevel2.Align:= alRight;
  ReceiverPanel.Visible:= RadioButton3.Checked;//
  ReceiverNamesPanel.Visible:= RadioButton3.Checked;
  OrderNumCheckBox.Visible:= not RadioButton3.Checked;
  DividerBevel2.Align:= alLeft;
  ExportButton.Align:= alLeft;

  Screen.Cursor:= crHourGlass;
  try
    LogGrid.Clear;
    FreeSheets;

    if RadioButton1.Checked then
      ShowBuildReport
    else if RadioButton2.Checked then
      ShowTestReport
    else if RadioButton3.Checked then
      ShowShipmentReport
    else if RadioButton4.Checked then
      ShowReclamationReport;

  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TReportForm.ShowBuildReport;
var
  BD, ED: TDate;
  TotalMotorNames: TStrVector;
  TotalMotorCounts: TIntVector;

  BuildDates: TDateVector;
  MotorNames, MotorNums, RotorNums: TStrVector;
  X,Y,Z: TIntVector;
begin
  ED:= DateTimePicker1.Date;
  BD:= DateTimePicker2.Date;
  if BD>ED then Exit;

  SQLite.BuildListLoad(BD, ED, UsedNameIDs,
                    OrderNumCheckBox.Checked, X, Y, Z,
                    BuildDates, MotorNames, MotorNums, RotorNums);
  SQLite.BuildTotalLoad(BD, ED, UsedNameIDs,
                     TotalMotorNames, TotalMotorCounts);
  MotorBuildSheet:= TMotorBuildSheet.Create(LogGrid);
  MotorBuildSheet.DrawReport(BD, ED,
                             BuildDates, MotorNames, MotorNums, RotorNums,
                             TotalMotorNames, TotalMotorCounts);
end;

procedure TReportForm.ShowTestReport;
var
  BD, ED: TDate;
  TotalMotorNames: TStrVector;
  TotalMotorCounts, TotalFailCounts: TIntVector;

  TestResults, X: TIntVector;
  TestDates: TDateVector;
  MotorNames, MotorNums, TestNotes: TStrVector;
begin
  ED:= DateTimePicker1.Date;
  BD:= DateTimePicker2.Date;
  if BD>ED then Exit;

  SQLite.TestListLoad(BD, ED, UsedNameIDs,
              OrderNumCheckBox.Checked, X, TestResults, TestDates,
              MotorNames, MotorNums, TestNotes);

  SQLite.TestTotalLoad(BD, ED, UsedNameIDs,
                     TotalMotorNames, TotalMotorCounts, TotalFailCounts);
  MotorTestSheet:= TMotorTestSheet.Create(LogGrid);
  MotorTestSheet.DrawReport(BD, ED,
                        TestDates, MotorNames, MotorNums,
                        TestNotes, TestResults,
                        TotalMotorNames, TotalMotorCounts, TotalFailCounts);
end;

procedure TReportForm.ShowShipmentReport;
var
  BD, ED: TDate;
  TotalMotorNames: TStrVector;
  TotalMotorCounts: TIntVector;

  RecieverNames: TStrVector;
  RecieverMotorNames: TStrMatrix;
  RecieverMotorCounts: TIntMatrix;

  ListSendDates: TDateVector;
  ListMotorNames, ListMotorNums, ListReceiverNames: TStrVector;

begin
  ED:= DateTimePicker1.Date;
  BD:= DateTimePicker2.Date;
  if BD>ED then Exit;

  SQLite.ShipmentTotalLoad(BD, ED, UsedNameIDs, TotalMotorNames, TotalMotorCounts);
  SQLite.ShipmentRecieversTotalLoad(BD, ED, UsedNameIDs, UsedReceiverIDs,
                RecieverNames, RecieverMotorNames, RecieverMotorCounts);
  SQLite.ShipmentMotorListLoad(BD, ED, UsedNameIDs, UsedReceiverIDs,
                ListSendDates, ListMotorNames, ListMotorNums, ListReceiverNames);

  ReportShipmentSheet:= TReportShipmentSheet.Create(LogGrid);
  ReportShipmentSheet.DrawReport(Length(UsedReceiverIDs)=1,
                BD, ED, TotalMotorNames, TotalMotorCounts,
                RecieverNames, RecieverMotorNames, RecieverMotorCounts,
                ListSendDates, ListMotorNames, ListMotorNums, ListReceiverNames);

end;

procedure TReportForm.ShowReclamationReport;
var
  BD, ED: TDate;
  TotalMotorNames: TStrVector;
  TotalMotorCounts: TIntVector;

  PlaceNames: TStrVector;
  PlaceMotorCounts: TIntVector;

  DefectNames: TStrVector;
  DefectMotorCounts: TIntVector;

  ReasonNames: TStrVector;
  ReasonMotorCounts: TIntVector;
begin
  ED:= DateTimePicker1.Date;
  BD:= DateTimePicker2.Date;
  if BD>ED then Exit;

  SQLite.ReclamationTotalLoad(BD, ED, UsedNameIDs, TotalMotorNames, TotalMotorCounts);
  SQLite.ReclamationPlacesLoad(BD, ED, UsedNameIDs, PlaceNames, PlaceMotorCounts);
  SQLite.ReclamationDefectsLoad(BD, ED, UsedNameIDs, DefectNames, DefectMotorCounts);
  SQLite.ReclamationReasonsLoad(BD, ED, UsedNameIDs, ReasonNames, ReasonMotorCounts);

  ReportReclamationSheet:= TReportReclamationSheet.Create(LogGrid);
  ReportReclamationSheet.DrawReport(Length(UsedNameIDs)=1,
                BD, ED, TotalMotorNames, TotalMotorCounts,
                PlaceNames, PlaceMotorCounts,
                DefectNames, DefectMotorCounts,
                ReasonNames, ReasonMotorCounts);
end;



end.

