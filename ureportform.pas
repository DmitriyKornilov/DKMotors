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
    NumberListCheckBox: TCheckBox;
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
    NumberListPanel: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    ReportPeriodPanel: TPanel;
    ToolPanel: TPanel;
    procedure ChooseMotorNamesButtonClick(Sender: TObject);
    procedure ChooseRecieverNamesButtonClick(Sender: TObject);
    procedure NumberListCheckBoxChange(Sender: TObject);
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
  PageOrientation: TsPageOrientation;
begin
  Exporter:= TGridExporter.Create(LogGrid);
  try
    //Exporter.SheetName:= 'Отчет';
    PageOrientation:= spoPortrait;
    if RadioButton4.Checked then
      PageOrientation:= spoLandscape;
    Exporter.PageSettings(PageOrientation, pfWidth);
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
  SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, False, UsedNameIDs, UsedNames);
  SQLite.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, False, UsedReceiverIDs, UsedReceiverNames);
end;

procedure TReportForm.ChooseMotorNamesButtonClick(Sender: TObject);
begin
  if SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, True, UsedNameIDs, UsedNames) then
    ShowReport;
end;

procedure TReportForm.ChooseRecieverNamesButtonClick(Sender: TObject);
begin
  if SQLite.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, True, UsedReceiverIDs, UsedReceiverNames) then
    ShowReport;
end;

procedure TReportForm.NumberListCheckBoxChange(Sender: TObject);
begin
  ShowReport;
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
  DividerBevel2.Visible:= not RadioButton4.Checked;
  ExportButton.Align:= alRight;
  DividerBevel2.Align:= alRight;
  NumberListPanel.Visible:= not RadioButton4.Checked;
  ReceiverNamesPanel.Visible:= RadioButton3.Checked;
  OrderNumCheckBox.Visible:= (RadioButton1.Checked or RadioButton2.Checked) and
                             NumberListCheckBox.Checked;
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

  SQLite.TestListLoad(BD, ED, UsedNameIDs,
              OrderNumCheckBox.Checked, X, TestResults, TestDates,
              MotorNames, MotorNums, TestNotes);

  SQLite.TestTotalLoad(BD, ED, UsedNameIDs,
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

  SQLite.ShipmentTotalLoad(BD, ED, UsedNameIDs, TotalMotorNames, TotalMotorCounts);
  SQLite.ShipmentRecieversTotalLoad(BD, ED, UsedNameIDs, UsedReceiverIDs,
                RecieverNames, RecieverMotorNames, RecieverMotorCounts);
  SQLite.ShipmentMotorListLoad(BD, ED, UsedNameIDs, UsedReceiverIDs,
                ListSendDates, ListMotorNames, ListMotorNums, ListReceiverNames);

  ReportShipmentSheet:= TReportShipmentSheet.Create(LogGrid);
  ReportShipmentSheet.DrawReport(Length(UsedReceiverIDs)=1,
                NumberListCheckBox.Checked,
                BD, ED, TotalMotorNames, TotalMotorCounts,
                RecieverNames, RecieverMotorNames, RecieverMotorCounts,
                ListSendDates, ListMotorNames, ListMotorNums, ListReceiverNames);

end;

procedure TReportForm.ShowReclamationReport;
var
  BD, ED: TDate;
  TotalMotorNames: TStrVector;
  TotalMotorCounts: TIntVector;
  TotalMotorReasonCounts: TIntMatrix;

  TitleReasonIDs: TIntVector;
  TitleReasonNames: TStrVector;

  PlaceNames: TStrVector;
  PlaceMotorCounts: TIntVector;
  PlaceReasonCounts: TIntMatrix;

  DefectNames: TStrVector;
  DefectMotorCounts: TIntVector;
  DefectReasonCounts: TIntMatrix;
begin
  ED:= DateTimePicker1.Date;
  BD:= DateTimePicker2.Date;
  if BD>ED then Exit;

  SQLite.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
                     TitleReasonIDs, TitleReasonNames, True {TODO: не указано тоже нужно учесть!});

  SQLite.ReclamationTotalWithReasonsLoad(BD, ED, UsedNameIDs, TitleReasonIDs,
                   TotalMotorNames, TotalMotorCounts, TotalMotorReasonCounts);
  SQLite.ReclamationDefectsWithReasonsLoad(BD, ED, UsedNameIDs, TitleReasonIDs,
                   DefectNames, DefectMotorCounts, DefectReasonCounts);
  SQLite.ReclamationPlacesWithReasonsLoad(BD, ED, UsedNameIDs, TitleReasonIDs,
                   PlaceNames, PlaceMotorCounts, PlaceReasonCounts);


  ReportReclamationSheet:= TReportReclamationSheet.Create(LogGrid, Length(TitleReasonNames));
  ReportReclamationSheet.DrawReport(Length(UsedNameIDs)=1,
                BD, ED, TotalMotorNames, TotalMotorCounts,
                TitleReasonNames, TotalMotorReasonCounts,
                PlaceNames, PlaceMotorCounts, PlaceReasonCounts,
                DefectNames, DefectMotorCounts, DefectReasonCounts);
end;



end.

