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
    OrderNumCheckBox: TCheckBox;
    CloseButton: TSpeedButton;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    DividerBevel4: TDividerBevel;
    DividerBevel5: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Label1: TLabel;
    LogGrid: TsWorksheetGrid;
    MotorNameComboBox: TComboBox;
    ReceiverNameComboBox: TComboBox;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel6: TPanel;
    ReceiverPanel: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    ReportPeriodPanel: TPanel;
    ToolPanel: TPanel;
    procedure OrderNumCheckBoxChange(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNameComboBoxChange(Sender: TObject);
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

procedure TReportForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TReportForm.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= FirstDayInMonth(Date);
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs, False);
  SQLite.ReceiverIDsAndNamesLoad(ReceiverNameComboBox, ReceiverIDs, False);
end;

procedure TReportForm.FormDestroy(Sender: TObject);
begin
  FreeSheets;
end;

procedure TReportForm.FormShow(Sender: TObject);
begin
  ShowReport;
end;

procedure TReportForm.MotorNameComboBoxChange(Sender: TObject);
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
  ReceiverPanel.Visible:= RadioButton3.Checked;
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

  SQLite.BuildListLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
                    OrderNumCheckBox.Checked, X, Y, Z,
                    BuildDates, MotorNames, MotorNums, RotorNums);
  SQLite.BuildTotalLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
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

  SQLite.TestListLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
              OrderNumCheckBox.Checked, X, TestResults, TestDates,
              MotorNames, MotorNums, TestNotes);

  SQLite.TestTotalLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
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

  SQLite.ShipmentTotalLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
                         TotalMotorNames, TotalMotorCounts);
  SQLite.ShipmentRecieversTotalLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
                ReceiverIDs[ReceiverNameComboBox.ItemIndex],
                RecieverNames, RecieverMotorNames, RecieverMotorCounts);
  SQLite.ShipmentMotorListLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
                ReceiverIDs[ReceiverNameComboBox.ItemIndex],
                ListSendDates, ListMotorNames, ListMotorNums, ListReceiverNames);

  ReportShipmentSheet:= TReportShipmentSheet.Create(LogGrid);
  ReportShipmentSheet.DrawReport(ReceiverNameComboBox.ItemIndex>0,
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

  SQLite.ReclamationTotalLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
                         TotalMotorNames, TotalMotorCounts);
  SQLite.ReclamationPlacesLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
                         PlaceNames, PlaceMotorCounts);
  SQLite.ReclamationDefectsLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
                         DefectNames, DefectMotorCounts);
  SQLite.ReclamationReasonsLoad(BD, ED, NameIDs[MotorNameComboBox.ItemIndex],
                         ReasonNames, ReasonMotorCounts);

  ReportReclamationSheet:= TReportReclamationSheet.Create(LogGrid);
  ReportReclamationSheet.DrawReport(MotorNameComboBox.ItemIndex>0,
                BD, ED, TotalMotorNames, TotalMotorCounts,
                PlaceNames, PlaceMotorCounts,
                DefectNames, DefectMotorCounts,
                ReasonNames, ReasonMotorCounts);
end;



end.

