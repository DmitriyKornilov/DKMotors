unit UStatisticForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons, BCButton,
  fpspreadsheetgrid, VirtualTrees, LCLType, StdCtrls, Spin, ComCtrls, DateTimePicker,
  //DK packages utils
  DK_Vector, DK_VSTTables, DK_DateUtils, DK_Matrix, DK_SheetExporter, DK_SheetConst,
  DK_StrUtils, DK_Zoom, DK_VSTTableTools, DK_CtrlUtils,
  //Project utils
  UVars, USheets;

type

  { TStatisticForm }

  TStatisticForm = class(TForm)
    ANEMAsSameNameCheckBox: TCheckBox;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    ExportButton: TBCButton;
    Label10: TLabel;
    ANEMPanel: TPanel;
    SeveralPeriodsCheckBox: TCheckBox;
    ShowPercentCheckBox: TCheckBox;
    ShowGraphicsCheckBox: TCheckBox;
    ShowLinePercentCheckBox: TCheckBox;
    TotalCountForUsedParamsOnlyCheckBox: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Grid1: TsWorksheetGrid;
    Label3: TLabel;
    Label4: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    MainPanel: TPanel;
    Panel1: TPanel;
    Panel10: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel7: TPanel;
    Panel9: TPanel;
    ReportPeriodPanel: TPanel;
    AdditionYearCountSpinEdit: TSpinEdit;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    ToolPanel: TPanel;
    ClientPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    VT3: TVirtualStringTree;
    ZoomPanel: TPanel;
    procedure ANEMAsSameNameCheckBoxChange(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure SeveralPeriodsCheckBoxChange(Sender: TObject);
    procedure ShowGraphicsCheckBoxChange(Sender: TObject);
    procedure ShowLinePercentCheckBoxChange(Sender: TObject);
    procedure ShowPercentCheckBoxChange(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure AdditionYearCountSpinEditChange(Sender: TObject);
    procedure TotalCountForUsedParamsOnlyCheckBoxChange(Sender: TObject);
  private
    SelectedIndex: Integer;
    CanShow: Boolean;

    ReasonIDs: TIntVector;
    ReasonNames: TStrVector;

    ReasonList: TVSTCheckList;
    StatisticList: TVSTStringList;
    MonthReportTypeList: TVSTStringList;

    ParamNames: TStrVector;
    Counts: TIntMatrix3D;
    MotorNames: String;
    AdditionYearsCount: Integer;

    ZoomPercent: Integer;

    procedure VerifyDates;

    procedure CreateStatisticList;
    procedure SelectStatistic;

    procedure CreateReasonList;

    procedure CreateMonthReportTypeList;
    procedure SelectMonthReportType;

    //AParamType
    //0 - Распределение по наименованиям двигателей
    //1 - Распределение по неисправным элементам
    //2 - Распределение по предприятиям
    //3 - Распределение по месяцам
    procedure LoadStatistic(const AParamType: Integer);

    procedure Draw(const AZoomPercent: Integer);
    procedure DrawStatistic(const AParamType: Integer);
    procedure DrawStatisticForSinglePeriod(const AParamType: Integer);
    procedure DrawStatisticForSeveralPeriods(const AParamType: Integer);

    procedure ExportStatistic(const AParamType: Integer);
    procedure ExportStatisticForSinglePeriod(const AParamType: Integer);
    procedure ExportStatisticForSeveralPeriods(const AParamType: Integer);

  public
    procedure ViewUpdate;
  end;

var
  StatisticForm: TStatisticForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TStatisticForm }

procedure TStatisticForm.FormCreate(Sender: TObject);
begin
  SetToolPanels([
    ToolPanel
  ]);

  CanShow:= False;

  SelectedIndex:= -1;
  MainForm.SetNamesPanelsVisible(True, False);

  ZoomPercent:= 100;
  CreateZoomControls(50, 150, ZoomPercent, ZoomPanel, @Draw, True);

  CreateReasonList;
  CreateStatisticList;
  CreateMonthReportTypeList;

  DateTimePicker2.Date:= FirstDayInYear(Date);
  DateTimePicker1.Date:= LastDayInMonth(Date);
  CanShow:= True;

  ViewUpdate;
end;

procedure TStatisticForm.DateTimePicker1Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.DateTimePicker2Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.VerifyDates;
var
  BD, ED: TDate;
begin
  CanShow:= False;
  BD:= DateTimePicker2.Date;
  ED:= DateTimePicker1.Date;

  if ED<BD then
  begin
    ED:= DateTimePicker2.Date;
    BD:= DateTimePicker1.Date;
  end;

  if SeveralPeriodsCheckBox.Checked or (SelectedIndex=3) then
  begin
    if YearOfDate(BD)<>YearOfDate(ED) then
      ED:= LastDayInYear(BD);
  end;

  DateTimePicker2.Date:= BD;
  DateTimePicker1.Date:= ED;

  CanShow:= True;
end;

procedure TStatisticForm.SeveralPeriodsCheckBoxChange(Sender: TObject);
begin
  AdditionYearCountSpinEdit.Enabled:= SeveralPeriodsCheckBox.Checked;
  ShowLinePercentCheckBox.Visible:= not SeveralPeriodsCheckBox.Checked;
  ViewUpdate;
end;

procedure TStatisticForm.ShowGraphicsCheckBoxChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.ShowLinePercentCheckBoxChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.ANEMAsSameNameCheckBoxChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.ExportButtonClick(Sender: TObject);
begin
  ExportStatistic(SelectedIndex);
end;

procedure TStatisticForm.ShowPercentCheckBoxChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.FormDestroy(Sender: TObject);
begin
  if Assigned(StatisticList) then FreeAndNil(StatisticList);
  if Assigned(ReasonList) then FreeAndNil(ReasonList);
  if Assigned(MonthReportTypeList) then FreeAndNil(MonthReportTypeList);
end;

procedure TStatisticForm.AdditionYearCountSpinEditChange(Sender: TObject);
begin
  if AdditionYearCountSpinEdit.Value=1 then
    Label4.Caption:= 'предыдущий год'
  else if (AdditionYearCountSpinEdit.Value>1) and (AdditionYearCountSpinEdit.Value<5) then
    Label4.Caption:= 'предыдущиx года'
  else
    Label4.Caption:= 'предыдущиx лет';
  ViewUpdate;
end;

procedure TStatisticForm.TotalCountForUsedParamsOnlyCheckBoxChange(
  Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.ViewUpdate;
begin
  if not CanShow then Exit;
  VerifyDates;

  Screen.Cursor:= crHourGlass;
  try
    Grid1.Clear;
    LoadStatistic(SelectedIndex);
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TStatisticForm.CreateStatisticList;
var
  S: String;
  V: TStrVector;
begin
  S:= 'Статистика:';
  V:= VCreateStr([
    'Распределение по наименованиям двигателей',
    'Распределение по неисправным элементам',
    'Распределение по предприятиям',
    'Распределение по месяцам',
    'Распределение по пробегу локомотива'
  ]);
  SelectedIndex:= 0;
  StatisticList:= TVSTStringList.Create(VT1, S, @SelectStatistic);
  StatisticList.Update(V, SelectedIndex);
  Panel2.Height:= StatisticList.AutoHeightValue + 10;
end;

procedure TStatisticForm.SelectStatistic;
begin
  if SelectedIndex=StatisticList.SelectedIndex then Exit;
  SelectedIndex:= StatisticList.SelectedIndex;
  ANEMPanel.Visible:= SelectedIndex=0;
  VT3.Visible:= SelectedIndex=3;
  ViewUpdate;
end;

procedure TStatisticForm.CreateReasonList;
var
  i: Integer;
  S: String;
  V: TStrVector;
  B: TBoolVector;
begin
  ReasonIDs:= VCreateInt([
    -1,
    0,
    1,
    2,
    3,
    4
  ]);
  ReasonNames:= VCreateStr([
    'Общее количество за период',
    'Не расследовано',
    'Некачественные комплектующие',
    'Дефект сборки / изготовления',
    'Нарушение условий эксплуатации',
    'Электродвигатель исправен'
  ]);
  VDim(V{%H-}, Length(ReasonNames));
  for i:= 0 to High(ReasonNames) do
    V[i]:= SFirstLower(ReasonNames[i]);
  S:= 'Включать в отчёт:';
  B:= VCreateBool(Length(V), True);
  ReasonList:= TVSTCheckList.Create(VT2, S, @ViewUpdate);
  ReasonList.Update(V, B);
end;

procedure TStatisticForm.CreateMonthReportTypeList;
var
  S: String;
  V: TStrVector;
begin
  S:= 'Вид отчёта:';
  V:= VCreateStr([
    'распределение количества рекламаций',
    'накопление количества рекламаций'
  ]);
  MonthReportTypeList:= TVSTStringList.Create(VT3, S, @SelectMonthReportType);
  MonthReportTypeList.Update(V);
end;

procedure TStatisticForm.SelectMonthReportType;
begin
  ViewUpdate;
end;

procedure TStatisticForm.LoadStatistic(const AParamType: Integer);
var
  BD, ED: TDate;
begin
  if not ReasonList.IsSelected then Exit;

  if AParamType<0 then Exit;
  BD:= DateTimePicker2.Date;
  ED:= DateTimePicker1.Date;
  MotorNames:= VVectorToStr(MainForm.UsedNames, ', ');

  AdditionYearsCount:= 0;
  if SeveralPeriodsCheckBox.Checked then
    AdditionYearsCount:= AdditionYearCountSpinEdit.Value;

  case AParamType of
    0: DataBase.ReclamationMotorsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs,
                                  ANEMAsSameNameCheckBox.Checked, ParamNames, Counts);
    1: DataBase.ReclamationDefectsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
    2: DataBase.ReclamationPlacesWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
    3: DataBase.ReclamationMonthsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
    4: DataBase.ReclamationMileagesWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
  end;

  DrawStatistic(AParamType);
end;

procedure TStatisticForm.Draw(const AZoomPercent: Integer);
begin
  ZoomPercent:= AZoomPercent;
  DrawStatistic(SelectedIndex);
end;

procedure TStatisticForm.DrawStatisticForSinglePeriod(const AParamType: Integer);
var
  Drawer: TStatisticSinglePeriodSheet;
begin
  case AParamType of
    0: Drawer:= TStatisticSinglePeriodAtMotorNamesSheet.Create(
                  Grid1.Worksheet, Grid1, ReasonList.Checkeds,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    1: Drawer:= TStatisticSinglePeriodAtDefectNamesSheet.Create(
                  Grid1.Worksheet, Grid1, ReasonList.Checkeds,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    2: Drawer:= TStatisticSinglePeriodAtPlaceNamesSheet.Create(
                  Grid1.Worksheet, Grid1, ReasonList.Checkeds,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    3: if MonthReportTypeList.SelectedIndex=0 then
         Drawer:= TStatisticSinglePeriodAtMonthNamesSheet.Create(
                  Grid1.Worksheet, Grid1, ReasonList.Checkeds,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked)
       else
         Drawer:= TStatisticSinglePeriodAtMonthNamesSumSheet.Create(
                  Grid1.Worksheet, Grid1, ReasonList.Checkeds,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    4: Drawer:= TStatisticSinglePeriodAtMileagesSheet.Create(
                  Grid1.Worksheet, Grid1, ReasonList.Checkeds,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
  end;

  try
    Drawer.Zoom(ZoomPercent);
    Drawer.Draw(DateTimePicker2.Date, DateTimePicker1.Date,
               MotorNames, ParamNames, ReasonNames, Counts[0],
               TotalCountForUsedParamsOnlyCheckBox.Checked,
               ShowGraphicsCheckBox.Checked);
  finally
    FreeAndNil(Drawer);
  end;
end;

procedure TStatisticForm.DrawStatisticForSeveralPeriods(const AParamType: Integer);
var
  Drawer: TStatisticSeveralPeriodsSheet;
begin
  case AParamType of
    0: Drawer:= TStatisticSeveralPeriodsAtMotorNamesSheet.Create(
                  Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                  ReasonList.Checkeds, ShowPercentCheckBox.Checked);
    1: Drawer:= TStatisticSeveralPeriodsAtDefectNamesSheet.Create(
                  Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                  ReasonList.Checkeds, ShowPercentCheckBox.Checked);
    2: Drawer:= TStatisticSeveralPeriodsAtPlaceNamesSheet.Create(
                  Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                  ReasonList.Checkeds, ShowPercentCheckBox.Checked);
    3: if MonthReportTypeList.SelectedIndex=0 then
         Drawer:= TStatisticSeveralPeriodsAtMonthNamesSheet.Create(
                  Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                  ReasonList.Checkeds, ShowPercentCheckBox.Checked)
       else
         Drawer:= TStatisticSeveralPeriodsAtMonthNamesSumSheet.Create(
                  Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                  ReasonList.Checkeds, ShowPercentCheckBox.Checked);
    4: Drawer:= TStatisticSeveralPeriodsAtMileagesSheet.Create(
                  Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                  ReasonList.Checkeds, ShowPercentCheckBox.Checked);
 end;

  try
    Drawer.Zoom(ZoomPercent);
    Drawer.Draw(DateTimePicker2.Date, DateTimePicker1.Date,
               MotorNames, ParamNames, ReasonNames, Counts,
               TotalCountForUsedParamsOnlyCheckBox.Checked,
               ShowGraphicsCheckBox.Checked);
  finally
    FreeAndNil(Drawer);
  end;
end;

procedure TStatisticForm.DrawStatistic(const AParamType: Integer);
begin
  if MIsNil(Counts) then Exit;
  if SeveralPeriodsCheckBox.Checked then
    DrawStatisticForSeveralPeriods(AParamType)
  else
    DrawStatisticForSinglePeriod(AParamType);
end;

procedure TStatisticForm.ExportStatistic(const AParamType: Integer);
begin
  if MIsNil(Counts) then Exit;
  if SeveralPeriodsCheckBox.Checked then
    ExportStatisticForSeveralPeriods(AParamType)
  else
    ExportStatisticForSinglePeriod(AParamType);
end;

procedure TStatisticForm.ExportStatisticForSinglePeriod(const AParamType: Integer);
var
  Drawer: TStatisticSinglePeriodSheet;
  Sheet: TsWorksheet;
  Exporter: TSheetsExporter;
begin
  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');

    case AParamType of
      0: Drawer:= TStatisticSinglePeriodAtMotorNamesSheet.Create(
                    Sheet, nil, ReasonList.Checkeds,
                    ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
      1: Drawer:= TStatisticSinglePeriodAtDefectNamesSheet.Create(
                    Sheet, nil, ReasonList.Checkeds,
                    ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
      2: Drawer:= TStatisticSinglePeriodAtPlaceNamesSheet.Create(
                    Sheet, nil, ReasonList.Checkeds,
                    ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
      3: Drawer:= TStatisticSinglePeriodAtMonthNamesSheet.Create(
                    Sheet, nil, ReasonList.Checkeds,
                    ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
      4: Drawer:= TStatisticSinglePeriodAtMileagesSheet.Create(
                    Sheet, nil, ReasonList.Checkeds,
                    ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    end;

    try
      Drawer.Draw(DateTimePicker2.Date, DateTimePicker1.Date,
                  MotorNames, ParamNames, ReasonNames, Counts[0],
                  TotalCountForUsedParamsOnlyCheckBox.Checked,
                  ShowGraphicsCheckBox.Checked);

    finally
      FreeAndNil(Drawer);
    end;
    Exporter.PageSettings(spoPortrait);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TStatisticForm.ExportStatisticForSeveralPeriods(const AParamType: Integer);
var
  Drawer: TStatisticSeveralPeriodsSheet;
  Sheet: TsWorksheet;
  Exporter: TSheetsExporter;
begin
  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');

    case AParamType of
      0: Drawer:= TStatisticSeveralPeriodsAtMotorNamesSheet.Create(
                      Sheet, nil, AdditionYearCountSpinEdit.Value,
                      ReasonList.Checkeds, ShowPercentCheckBox.Checked);
      1: Drawer:= TStatisticSeveralPeriodsAtDefectNamesSheet.Create(
                      Sheet, nil, AdditionYearCountSpinEdit.Value,
                      ReasonList.Checkeds, ShowPercentCheckBox.Checked);
      2: Drawer:= TStatisticSeveralPeriodsAtPlaceNamesSheet.Create(
                      Sheet, nil, AdditionYearCountSpinEdit.Value,
                      ReasonList.Checkeds, ShowPercentCheckBox.Checked);
      3: Drawer:= TStatisticSeveralPeriodsAtMonthNamesSheet.Create(
                      Sheet, nil, AdditionYearCountSpinEdit.Value,
                      ReasonList.Checkeds, ShowPercentCheckBox.Checked);
      4: Drawer:= TStatisticSeveralPeriodsAtMileagesSheet.Create(
                      Sheet, nil, AdditionYearCountSpinEdit.Value,
                      ReasonList.Checkeds, ShowPercentCheckBox.Checked);
    end;

    try
      Drawer.Draw(DateTimePicker2.Date, DateTimePicker1.Date,
                  MotorNames, ParamNames, ReasonNames, Counts,
                  TotalCountForUsedParamsOnlyCheckBox.Checked,
                  ShowGraphicsCheckBox.Checked);

    finally
      FreeAndNil(Drawer);
    end;
    Exporter.PageSettings(spoLandscape);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

end.

