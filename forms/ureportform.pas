unit UReportForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, fpspreadsheetgrid, DateTimePicker, BCButton,
  //DK packages utils
  DK_DateUtils, DK_SheetExporter, DK_Vector, DK_Matrix, DK_CtrlUtils,
  //Project utils
  UVars, USheets;

type

  { TReportForm }

  TReportForm = class(TForm)
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    ExportButton: TBCButton;
    NumberListCheckBox: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label1: TLabel;
    LogGrid: TsWorksheetGrid;
    OrderNumCheckBox: TCheckBox;
    Panel1: TPanel;
    Panel2: TPanel;
    NumberListPanel: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    ReportPeriodPanel: TPanel;
    ToolPanel: TPanel;
    procedure ExportButtonClick(Sender: TObject);
    procedure NumberListCheckBoxChange(Sender: TObject);
    procedure OrderNumCheckBoxChange(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure RadioButton3Click(Sender: TObject);
  private
    MotorBuildSheet: TMotorBuildSheet;
    MotorTestSheet: TMotorTestSheet;
    ReportShipmentSheet: TReportShipmentSheet;

    procedure ShowBuildReport;
    procedure ShowTestReport;
    procedure ShowShipmentReport;

    procedure FreeSheets;

    procedure SetControlsVisible;
  public
    procedure ViewUpdate;
  end;

var
  ReportForm: TReportForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TReportForm }

procedure TReportForm.DateTimePicker2Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TReportForm.DateTimePicker1Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TReportForm.OrderNumCheckBoxChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TReportForm.FormCreate(Sender: TObject);
begin
  SetToolPanels([
    ToolPanel
  ]);

  SetControlsVisible;
  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= FirstDayInMonth(Date);
end;

procedure TReportForm.NumberListCheckBoxChange(Sender: TObject);
begin
  SetControlsVisible;
  ViewUpdate;
end;

procedure TReportForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(LogGrid);
  try
    Exporter.PageSettings(spoPortrait, pfWidth);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TReportForm.FormDestroy(Sender: TObject);
begin
  FreeSheets;
end;

procedure TReportForm.FormShow(Sender: TObject);
begin
  SetControlsVisible;
  ViewUpdate;
end;

procedure TReportForm.RadioButton1Click(Sender: TObject);
begin
  SetControlsVisible;
  ViewUpdate;
end;

procedure TReportForm.RadioButton2Click(Sender: TObject);
begin
  SetControlsVisible;
  ViewUpdate;
end;

procedure TReportForm.RadioButton3Click(Sender: TObject);
begin
  SetControlsVisible;
  ViewUpdate;
end;

procedure TReportForm.FreeSheets;
begin
  if Assigned(MotorBuildSheet) then FreeAndNil(MotorBuildSheet);
  if Assigned(MotorTestSheet) then FreeAndNil(MotorTestSheet);
  if Assigned(ReportShipmentSheet) then FreeAndNil(ReportShipmentSheet);
end;

procedure TReportForm.SetControlsVisible;
begin
  MainForm.SetNamesPanelsVisible(True, RadioButton3.Checked);
  OrderNumCheckBox.Visible:= NumberListCheckBox.Checked;
end;

procedure TReportForm.ViewUpdate;
begin
  Screen.Cursor:= crHourGlass;
  try
    LogGrid.Clear;
    FreeSheets;
    if RadioButton1.Checked then
      ShowBuildReport
    else if RadioButton2.Checked then
      ShowTestReport
    else if RadioButton3.Checked then
      ShowShipmentReport;
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

  DataBase.BuildListLoad(BD, ED, MainForm.UsedNameIDs,
                    OrderNumCheckBox.Checked, X, Y, Z,
                    BuildDates, MotorNames, MotorNums, RotorNums);
  DataBase.BuildTotalLoad(BD, ED, MainForm.UsedNameIDs,
                     TotalMotorNames, TotalMotorCounts);
  MotorBuildSheet:= TMotorBuildSheet.Create(LogGrid);
  MotorBuildSheet.DrawReport(BD, ED, NumberListCheckBox.Checked,
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

  DataBase.TestListLoad(BD, ED, MainForm.UsedNameIDs,
              OrderNumCheckBox.Checked, X, TestResults, TestDates,
              MotorNames, MotorNums, TestNotes);

  DataBase.TestTotalLoad(BD, ED, MainForm.UsedNameIDs,
                     TotalMotorNames, TotalMotorCounts, TotalFailCounts);
  MotorTestSheet:= TMotorTestSheet.Create(LogGrid);
  MotorTestSheet.DrawReport(BD, ED, NumberListCheckBox.Checked,
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

  DataBase.ShipmentTotalLoad(BD, ED, MainForm.UsedNameIDs, TotalMotorNames, TotalMotorCounts);
  DataBase.ShipmentRecieversTotalLoad(BD, ED, MainForm.UsedNameIDs, MainForm.UsedReceiverIDs,
                RecieverNames, RecieverMotorNames, RecieverMotorCounts);
  DataBase.ShipmentMotorListLoad(BD, ED, MainForm.UsedNameIDs, MainForm.UsedReceiverIDs,
                OrderNumCheckBox.Checked,
                ListSendDates, ListMotorNames, ListMotorNums, ListReceiverNames);

  ReportShipmentSheet:= TReportShipmentSheet.Create(LogGrid);
  ReportShipmentSheet.DrawReport(Length(MainForm.UsedReceiverIDs)=1,
                NumberListCheckBox.Checked,
                BD, ED, TotalMotorNames, TotalMotorCounts,
                RecieverNames, RecieverMotorNames, RecieverMotorCounts,
                ListSendDates, ListMotorNames, ListMotorNums, ListReceiverNames);
end;

end.

