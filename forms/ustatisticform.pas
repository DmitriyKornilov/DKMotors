unit UStatisticForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons, ComCtrls,
  fpspreadsheetgrid, VirtualTrees, LCLType, StdCtrls, Spin, DateTimePicker, DividerBevel,
  //DK packages utils
  DK_Vector, DK_VSTTables, DK_DateUtils, DK_Matrix, DK_SheetExporter, DK_SheetConst,
  DK_StrUtils, DK_Zoom, DK_VSTTableTools, DK_CtrlUtils, DK_VSTParamList,
  //Project utils
  UVars, UStatSheets, UStatSheetNew;

type

  { TStatisticForm }

  TStatisticForm = class(TForm)
    Bevel1: TBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    Label1: TLabel;
    ReportTypeComboBox: TComboBox;
    DividerBevel1: TDividerBevel;
    ExportButton: TSpeedButton;
    ReportTypePanel: TPanel;
    SettingClientPanel: TPanel;
    EndDateTimePicker: TDateTimePicker;
    BeginDatePicker: TDateTimePicker;
    ViewGrid: TsWorksheetGrid;
    Label3: TLabel;
    Label4: TLabel;
    MainPanel: TPanel;
    LeftPanel: TPanel;
    StatisticPanel: TPanel;
    RightPanel: TPanel;
    AdditionYearsPanel: TPanel;
    GridPanel: TPanel;
    ReportPeriodPanel: TPanel;
    AdditionYearCountSpinEdit: TSpinEdit;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    ToolPanel: TPanel;
    ClientPanel: TPanel;
    StatisticVT: TVirtualStringTree;
    ZoomPanel: TPanel;
    procedure ExportButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ReportTypeComboBoxChange(Sender: TObject);
    procedure EndDateTimePickerChange(Sender: TObject);
    procedure BeginDatePickerChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure AdditionYearCountSpinEditChange(Sender: TObject);
  private
    CanShow: Boolean;
    ZoomPercent: Integer;
    SelectedIndex: Integer;
    StatisticList: TVSTStringList;
    ParamList: TVSTParamList;

    //MotorNames - список наименований двигателей, включаемых в отчет
    MotorNames: String;
    //AdditionYearsCount - количество предыдущих лет для сопроставления данных
    AdditionYearsCount: Integer;

    //ParamNames - список параметров, по которым считается статистика
    // - наименования двигателей
    // - неисправные элементы двигателя
    // - депо обнаружения неисправности
    // - месяцы
    // - пробеги локомотива
    ParamNames: TStrVector;

    //ReasonIDs - ID критерия статистики
    //ReasonNames - наименование критерия статистики
    // 0 : Не расследовано
    // 1 : Некачественные комплектующие
    // 2 : Дефект сборки / изготовления
    // 3 : Нарушение условий эксплуатации
    // 4 : Электродвигатель исправен
    ReasonIDs: TIntVector;
    ReasonNames: TStrVector;

    //ClaimCounts - количество рекламаций
    // - Index1 - индекс периода
    //      0 - данные за текущий произвольный период (в этом случае это единственный индекс)
    //          или указанный период года
    //      1 - данные за указанный период 1-ого предшествующего года
    //      2 - данные за указанный период 2-ого предшествующего года
    //      3 - данные за указанный период 3-его предшествующего года
    //      4 - данные за указанный период 4-ого предшествующего года (max значение индекса)
    // - Index2 - [0..4] - индекс критерия статистики, соответствующего ReasonNames
    // - Index3 - [0..High(ParamNames)] - индекс параметра статистики, соответствующего ParamNames
    ClaimCounts: TIntMatrix3D;

    procedure VerifyDates;

    procedure CreateParamList;

    procedure CreateStatisticList;
    procedure SelectStatistic;
    //AParamType
    //0 - Распределение по наименованиям двигателей
    //1 - Распределение по неисправным элементам
    //2 - Распределение по предприятиям
    //3 - Распределение по месяцам
    //4 - Распределение по пробегу локомотива
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
  CanShow:= False;

  MainForm.SetNamesPanelsVisible(True, False);

  ZoomPercent:= 100;
  CreateZoomControls(50, 150, ZoomPercent, ZoomPanel, @Draw, True);

  DataBase.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName', ReasonIDs, ReasonNames);
  ReasonNames[0]:= 'Не расследовано';

  CreateParamList;

  SelectedIndex:= -1;
  CreateStatisticList;

  BeginDatePicker.Date:= FirstDayInYear(Date);
  EndDateTimePicker.Date:= LastDayInMonth(Date);
  CanShow:= True;
end;

procedure TStatisticForm.FormDestroy(Sender: TObject);
begin
  if Assigned(ParamList) then FreeAndNil(ParamList);
  if Assigned(StatisticList) then FreeAndNil(StatisticList);
end;

procedure TStatisticForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);
  Images.ToButtons([ExportButton]);

  StatisticPanel.Height:= StatisticList.AutoHeightValue + 10;

  ViewUpdate;
end;

procedure TStatisticForm.EndDateTimePickerChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.BeginDatePickerChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.VerifyDates;
var
  BD, ED: TDate;
begin
  CanShow:= False;
  BD:= BeginDatePicker.Date;
  ED:= EndDateTimePicker.Date;

  if ED<BD then
  begin
    ED:= BeginDatePicker.Date;
    BD:= EndDateTimePicker.Date;
  end;

  if ( ReportTypeComboBox.ItemIndex=1) or (SelectedIndex=3) then
  begin
    if YearOfDate(BD)<>YearOfDate(ED) then
      ED:= LastDayInYear(BD);
  end;

  BeginDatePicker.Date:= BD;
  EndDateTimePicker.Date:= ED;

  CanShow:= True;
end;

procedure TStatisticForm.ReportTypeComboBoxChange(Sender: TObject);
begin
  AdditionYearsPanel.Visible:= ReportTypeComboBox.ItemIndex=1;
  ParamList.ItemVisibles['AdditionShow']:= VCreateBool([True, True, ReportTypeComboBox.ItemIndex=0]);
  ViewUpdate;
end;

procedure TStatisticForm.ExportButtonClick(Sender: TObject);
begin
  ExportStatistic(SelectedIndex);
end;

procedure TStatisticForm.AdditionYearCountSpinEditChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.ViewUpdate;
begin
  if not CanShow then Exit;
  VerifyDates;

  Screen.Cursor:= crHourGlass;
  try
    ViewGrid.Clear;
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

  StatisticList:= TVSTStringList.Create(StatisticVT, S, @SelectStatistic);
  StatisticList.Update(V);
  StatisticPanel.Height:= StatisticList.AutoHeightValue + 10;
end;

procedure TStatisticForm.SelectStatistic;
begin
  if SelectedIndex=StatisticList.SelectedIndex then Exit;
  SelectedIndex:= StatisticList.SelectedIndex;

  ParamList.Visibles['ANEMIsSingleName']:= SelectedIndex=0;
  ParamList.Visibles['MonthTypeList']:= SelectedIndex=3;

  ViewUpdate;
end;

procedure TStatisticForm.CreateParamList;
var
  S: String;
  V: TStrVector;
  B: TBoolVector;
  i: Integer;
begin
  ParamList:= TVSTParamList.Create(SettingClientPanel);

  S:= 'Включать в отчёт столбцы:';
  VDim(V{%H-}, Length(ReasonNames));
  for i:= 0 to High(ReasonNames) do
    V[i]:= SFirstLower(ReasonNames[i]);
  B:= VCreateBool(Length(V), True);
  ParamList.AddCheckList('ReasonList', S, V, @ViewUpdate, B);

  S:= 'Считать общее количество (сумму) рекламаций:';
  V:= VCreateStr([
    'по всем критериям',
    'только по включенным в отчёт критериям',
    'не считать и не отображать'
  ]);
  ParamList.AddStringList('SumTypeList', S, V, @ViewUpdate);

  S:= 'Дополнительно отображать:';
  V:= VCreateStr([
    'гистограммы',
    '% от общего количества рекламаций за период',
    '% от общего количества рекламаций в строке'
  ]);
  ParamList.AddCheckList('AdditionShow', S, V, @ViewUpdate);

  //для отчета по месяцам
  S:= 'Вид отчёта:';
  V:= VCreateStr([
    'распределение количества рекламаций',
    'накопление количества рекламаций'
  ]);
  ParamList.AddStringList('MonthTypeList', S, V, @ViewUpdate);

  //для отчета по наименованиям двигателей
  S:= 'Оба исполнения АНЭМ225L4УХЛ2 IM1001 и IM1002:';
  V:= VCreateStr([
    'считать одним наименованием электродвигателя'
  ]);
  ParamList.AddCheckList('ANEMIsSingleName', S, V, @ViewUpdate);
end;

procedure TStatisticForm.LoadStatistic(const AParamType: Integer);
var
  BD, ED: TDate;
begin
  if not ParamList.IsSelected['ReasonList'] then Exit;

  //if not ParamList.IsSelected['ReasonList'] then Exit;
  //
  //if AParamType<0 then Exit;
  //BD:= BeginDatePicker.Date;
  //ED:= EndDateTimePicker.Date;
  //MotorNames:= VVectorToStr(MainForm.UsedNames, ', ');
  //
  //AdditionYearsCount:= 0;
  //if SeveralPeriodsCheckBox.Checked then
  //  AdditionYearsCount:= AdditionYearCountSpinEdit.Value;
  //
  //case AParamType of
  //  0: DataBase.ReclamationMotorsWithReasonsLoad(BD, ED, AdditionYearsCount,
  //                                MainForm.UsedNameIDs, ReasonIDs,
  //                                ParamList.Checked['ANEMIsSingleName', 0],
  //                                ParamNames, ClaimCounts);
  //  1: DataBase.ReclamationDefectsWithReasonsLoad(BD, ED, AdditionYearsCount,
  //                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, ClaimCounts);
  //  2: DataBase.ReclamationPlacesWithReasonsLoad(BD, ED, AdditionYearsCount,
  //                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, ClaimCounts);
  //  3: DataBase.ReclamationMonthsWithReasonsLoad(BD, ED, AdditionYearsCount,
  //                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, ClaimCounts);
  //  4: DataBase.ReclamationMileagesWithReasonsLoad(BD, ED, AdditionYearsCount,
  //                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, ClaimCounts);
  //end;
  //
  //DrawStatistic(AParamType);
end;

procedure TStatisticForm.Draw(const AZoomPercent: Integer);
begin
  ZoomPercent:= AZoomPercent;
  DrawStatistic(SelectedIndex);
end;

procedure TStatisticForm.DrawStatisticForSinglePeriod(const AParamType: Integer);
var
  Drawer: TStatisticSinglePeriodSheet;
  //Drawer: TStatSheet;
  //TmpReasonValues, TmpParamValues: TIntVector;
  //TmpReasonNames: TStrVector;
begin
  //if AParamType<>0 then Exit;
  //
  //TmpReasonValues:= VCreateInt([0, 4, 25, 10, 3]);
  //TmpParamValues:= VCreateInt([39, 3]);
  //TmpReasonNames:= VCut(ReasonNames, 1);
  //
  //Drawer:= TStatSheet.Create(ViewGrid.Worksheet, ViewGrid, GridFont);
  //try
  //  Drawer.Zoom(ZoomPercent);
  //  Drawer.Draw('наименованиям двигателей',
  //              TmpReasonNames, ParamNames, TmpReasonValues, TmpParamValues);
  //finally
  //  FreeAndNil(Drawer);
  //end;





  //case AParamType of
  //  0: Drawer:= TStatisticSinglePeriodAtMotorNamesSheet.Create(
  //                ViewGrid.Worksheet, ViewGrid, GridFont,
  //                ParamList.Checkeds['ReasonList'],
  //                ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //  1: Drawer:= TStatisticSinglePeriodAtDefectNamesSheet.Create(
  //                ViewGrid.Worksheet, ViewGrid, GridFont,
  //                ParamList.Checkeds['ReasonList'],
  //                ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //  2: Drawer:= TStatisticSinglePeriodAtPlaceNamesSheet.Create(
  //                ViewGrid.Worksheet, ViewGrid, GridFont,
  //                ParamList.Checkeds['ReasonList'],
  //                ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //  3: if ParamList.Selected['MonthTypeList']=0 then //распределение по месяцам
  //       Drawer:= TStatisticSinglePeriodAtMonthNamesSheet.Create(
  //                ViewGrid.Worksheet, ViewGrid, GridFont,
  //                ParamList.Checkeds['ReasonList'],
  //                ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                )
  //     else // накопление по месяцам
  //       Drawer:= TStatisticSinglePeriodAtMonthNamesSumSheet.Create(
  //                ViewGrid.Worksheet, ViewGrid, GridFont,
  //                ParamList.Checkeds['ReasonList'],
  //                ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //  4: Drawer:= TStatisticSinglePeriodAtMileagesSheet.Create(
  //                ViewGrid.Worksheet, ViewGrid, GridFont,
  //                ParamList.Checkeds['ReasonList'],
  //                ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //end;
  //
  //try
  //  Drawer.Zoom(ZoomPercent);
  //  Drawer.Draw(BeginDatePicker.Date, EndDateTimePicker.Date,
  //             MotorNames, ParamNames, ReasonNames, ClaimCounts[0],
  //             ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
  //             ParamList.Checked['AdditionShow', 0] //отображать гистограммы
  //             );
  //finally
  //  FreeAndNil(Drawer);
  //end;
end;

procedure TStatisticForm.DrawStatisticForSeveralPeriods(const AParamType: Integer);
var
  Drawer: TStatisticSeveralPeriodsSheet;
begin
 // case AParamType of
 //   0: Drawer:= TStatisticSeveralPeriodsAtMotorNamesSheet.Create(
 //                 ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
 //                 ParamList.Checkeds['ReasonList'],
 //                 ParamList.Checked['AdditionShow', 1] //% от суммы за период
 //                 );
 //   1: Drawer:= TStatisticSeveralPeriodsAtDefectNamesSheet.Create(
 //                 ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
 //                 ParamList.Checkeds['ReasonList'],
 //                 ParamList.Checked['AdditionShow', 1] //% от суммы за период
 //                 );
 //   2: Drawer:= TStatisticSeveralPeriodsAtPlaceNamesSheet.Create(
 //                 ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
 //                 ParamList.Checkeds['ReasonList'],
 //                 ParamList.Checked['AdditionShow', 1] //% от суммы за период
 //                 );
 //   3: if ParamList.Selected['MonthTypeList']=0 then //распределение по месяцам
 //        Drawer:= TStatisticSeveralPeriodsAtMonthNamesSheet.Create(
 //                 ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
 //                 ParamList.Checkeds['ReasonList'],
 //                 ParamList.Checked['AdditionShow', 1] //% от суммы за период
 //                 )
 //      else  // накопление по месяцам
 //        Drawer:= TStatisticSeveralPeriodsAtMonthNamesSumSheet.Create(
 //                 ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
 //                 ParamList.Checkeds['ReasonList'],
 //                 ParamList.Checked['AdditionShow', 1] //% от суммы за период
 //                 );
 //   4: Drawer:= TStatisticSeveralPeriodsAtMileagesSheet.Create(
 //                 ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
 //                 ParamList.Checkeds['ReasonList'],
 //                 ParamList.Checked['AdditionShow', 1] //% от суммы за период
 //                 );
 //end;
 //
 // try
 //   Drawer.Zoom(ZoomPercent);
 //   Drawer.Draw(BeginDatePicker.Date, EndDateTimePicker.Date,
 //              MotorNames, ParamNames, ReasonNames, ClaimCounts,
 //              ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
 //              ParamList.Checked['AdditionShow', 0] //отображать гистограммы
 //              );
 // finally
 //   FreeAndNil(Drawer);
 // end;
end;

procedure TStatisticForm.DrawStatistic(const AParamType: Integer);
begin
  if MIsNil(ClaimCounts) then Exit;
  if ReportTypeComboBox.ItemIndex=0 then
    DrawStatisticForSinglePeriod(AParamType)
  else if ReportTypeComboBox.ItemIndex=1 then
    DrawStatisticForSeveralPeriods(AParamType);
end;

procedure TStatisticForm.ExportStatistic(const AParamType: Integer);
begin
  if MIsNil(ClaimCounts) then Exit;
  if ReportTypeComboBox.ItemIndex=0 then
    ExportStatisticForSinglePeriod(AParamType)
  else if ReportTypeComboBox.ItemIndex=1 then
    ExportStatisticForSeveralPeriods(AParamType);
end;

procedure TStatisticForm.ExportStatisticForSinglePeriod(const AParamType: Integer);
var
  Drawer: TStatisticSinglePeriodSheet;

  Sheet: TsWorksheet;
  Exporter: TSheetsExporter;

  //Drawer: TStatSheet;
  //TmpReasonValues, TmpParamValues: TIntVector;
  //TmpReasonNames: TStrVector;
begin
  //if AParamType<>0 then Exit;
  //Exporter:= TSheetsExporter.Create;
  //try
  //  Sheet:= Exporter.AddWorksheet('Лист1');
  //
  //  TmpReasonValues:= VCreateInt([0, 4, 25, 10, 3]);
  //  TmpParamValues:= VCreateInt([39, 3]);
  //  TmpReasonNames:= VCut(ReasonNames, 1);
  //
  //  Drawer:= TStatSheet.Create(Sheet, nil, GridFont);
  //  try
  //    Drawer.Draw('наименованиям двигателей',
  //                TmpReasonNames, ParamNames, TmpReasonValues, TmpParamValues);
  //  finally
  //    FreeAndNil(Drawer);
  //  end;
  //
  //Exporter.PageSettings(spoPortrait, pfWidth, True, False);
  //  Exporter.Save('Выполнено!');
  //finally
  //  FreeAndNil(Exporter);
  //end;


  //Exporter:= TSheetsExporter.Create;
  //try
  //  Sheet:= Exporter.AddWorksheet('Лист1');
  //  case AParamType of
  //    0: Drawer:= TStatisticSinglePeriodAtMotorNamesSheet.Create(Sheet, nil, GridFont,
  //                  ParamList.Checkeds['ReasonList'],
  //                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //    1: Drawer:= TStatisticSinglePeriodAtDefectNamesSheet.Create(Sheet, nil, GridFont,
  //                  ParamList.Checkeds['ReasonList'],
  //                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //    2: Drawer:= TStatisticSinglePeriodAtPlaceNamesSheet.Create(Sheet, nil, GridFont,
  //                  ParamList.Checkeds['ReasonList'],
  //                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //    3: if ParamList.Selected['MonthTypeList']=0 then //распределение по месяцам
  //       Drawer:= TStatisticSinglePeriodAtMonthNamesSheet.Create(Sheet, nil, GridFont,
  //                ParamList.Checkeds['ReasonList'],
  //                ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                )
  //     else // накопление по месяцам
  //       Drawer:= TStatisticSinglePeriodAtMonthNamesSumSheet.Create(Sheet, nil, GridFont,
  //                ParamList.Checkeds['ReasonList'],
  //                ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //    4: Drawer:= TStatisticSinglePeriodAtMileagesSheet.Create(Sheet, nil, GridFont,
  //                  ParamList.Checkeds['ReasonList'],
  //                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
  //                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
  //                );
  //  end;
  //
  //  try
  //    Drawer.Draw(BeginDatePicker.Date, EndDateTimePicker.Date,
  //                MotorNames, ParamNames, ReasonNames, ClaimCounts[0],
  //                ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
  //                ParamList.Checked['AdditionShow', 0] //отображать гистограммы
  //                );
  //
  //  finally
  //    FreeAndNil(Drawer);
  //  end;
  //  Exporter.PageSettings(spoPortrait, pfWidth, True, False);
  //  Exporter.Save('Выполнено!');
  //finally
  //  FreeAndNil(Exporter);
  //end;
end;

procedure TStatisticForm.ExportStatisticForSeveralPeriods(const AParamType: Integer);
var
  Drawer: TStatisticSeveralPeriodsSheet;
  Sheet: TsWorksheet;
  Exporter: TSheetsExporter;
begin
  //Exporter:= TSheetsExporter.Create;
  //try
  //  Sheet:= Exporter.AddWorksheet('Лист1');
  //
  //  case AParamType of
  //    0: Drawer:= TStatisticSeveralPeriodsAtMotorNamesSheet.Create(
  //                    Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
  //                    ParamList.Checkeds['ReasonList'],
  //                    ParamList.Checked['AdditionShow', 1] //% от суммы за период
  //                );
  //    1: Drawer:= TStatisticSeveralPeriodsAtDefectNamesSheet.Create(
  //                    Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
  //                    ParamList.Checkeds['ReasonList'],
  //                    ParamList.Checked['AdditionShow', 1] //% от суммы за период
  //                );
  //    2: Drawer:= TStatisticSeveralPeriodsAtPlaceNamesSheet.Create(
  //                    Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
  //                    ParamList.Checkeds['ReasonList'],
  //                    ParamList.Checked['AdditionShow', 1] //% от суммы за период
  //                );
  //    3: if ParamList.Selected['MonthTypeList']=0 then //распределение по месяцам
  //         Drawer:= TStatisticSeveralPeriodsAtMonthNamesSheet.Create(
  //                  Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
  //                  ParamList.Checkeds['ReasonList'],
  //                  ParamList.Checked['AdditionShow', 1] //% от суммы за период
  //                  )
  //       else  // накопление по месяцам
  //         Drawer:= TStatisticSeveralPeriodsAtMonthNamesSumSheet.Create(
  //                  Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
  //                  ParamList.Checkeds['ReasonList'],
  //                  ParamList.Checked['AdditionShow', 1] //% от суммы за период
  //                  );
  //    4: Drawer:= TStatisticSeveralPeriodsAtMileagesSheet.Create(
  //                    Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
  //                    ParamList.Checkeds['ReasonList'],
  //                    ParamList.Checked['AdditionShow', 1] //% от суммы за период
  //                );
  //  end;
  //
  //  try
  //    Drawer.Draw(BeginDatePicker.Date, EndDateTimePicker.Date,
  //                MotorNames, ParamNames, ReasonNames, ClaimCounts,
  //                ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
  //                ParamList.Checked['AdditionShow', 0] //отображать гистограммы
  //                );
  //
  //  finally
  //    FreeAndNil(Drawer);
  //  end;
  //  Exporter.PageSettings(spoLandscape, pfWidth, True, False);
  //  Exporter.Save('Выполнено!');
  //finally
  //  FreeAndNil(Exporter);
  //end;
end;

end.

