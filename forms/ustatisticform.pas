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
  UVars, UStatSheets, UStatistic;

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
    StatSelectedIndex: Integer;
    StatisticList: TVSTStringList;
    ParamList: TVSTParamList;

    //MotorNamesStr - список наименований двигателей, включаемых в отчет
    MotorNamesStr: String;
    //PeriodStr - отчетный период
    PeriodStr: String;
    //AdditionYearsCount - количество предыдущих лет для сопроставления данных
    AdditionYearsCount: Integer;

    //ParamNames - список параметров, по которым считается статистика
    // - наименования двигателей
    // - неисправные элементы двигателя
    // - депо обнаружения неисправности
    // - месяцы
    // - пробеги локомотива
    ParamNames: TStrVector;
    // ParamNeeds - флаги использования параметров
    ParamNeeds: TBoolVector;

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
    //      0 - данные за произвольный период (в этом случае это единственный индекс)
    //          или указанный период года, который нужно сопоставить с предыдущими
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
    procedure LoadStatistic;

    procedure Draw(const AZoomPercent: Integer);
    procedure DrawStatistic;
    procedure DrawStatisticForSinglePeriod;
    procedure DrawStatisticForSeveralPeriods;

    procedure ExportStatistic;
    procedure ExportStatisticForSinglePeriod;
    procedure ExportStatisticForSeveralPeriods;

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

  DataBase.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName', ReasonIDs, ReasonNames, False, 'ReasonID');
  ReasonNames[0]:= 'Не расследовано';

  CreateParamList;

  StatSelectedIndex:= -1;
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
  ParamList.AutoHeight;

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

  if ( ReportTypeComboBox.ItemIndex=1) or (StatSelectedIndex=3) then
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
  ExportStatistic;
end;

procedure TStatisticForm.AdditionYearCountSpinEditChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.ViewUpdate;
begin
  if not CanShow then Exit;

  ParamList.Visibles['MotorTypeAsSingleName']:= StatSelectedIndex=0;
  ParamList.Visibles['MonthTypeList']:= StatSelectedIndex=3;

  VerifyDates;

  Screen.Cursor:= crHourGlass;
  try
    ViewGrid.Clear;
    LoadStatistic;
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TStatisticForm.CreateStatisticList;
var
  S: String;
  V: TStrVector;
begin
  S:= 'Вид статистики:';
  V:= VCreateStr([
    'Распределение по наименованиям электродвигателей',
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
  if StatSelectedIndex=StatisticList.SelectedIndex then Exit;
  StatSelectedIndex:= StatisticList.SelectedIndex;
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

  S:= 'Включать в отчёт критерии неисправности:';
  VDim(V{%H-}, Length(ReasonNames));
  for i:= 0 to High(ReasonNames) do
    V[i]:= SFirstLower(ReasonNames[i]);
  B:= VCreateBool(Length(V), True);
  ParamList.AddCheckList('ReasonList', S, V, @ViewUpdate, B);

  S:= 'Включать в отчёт данные рекламаций:';
  V:= VCreateStr([
    'Общее количество по виду статистики',
    'Общее количество по критериям неисправности',
    'Количество по виду статистики и критериям неисправности'
  ]);
  B:= VCreateBool(Length(V), True);
  ParamList.AddCheckList('DataList', S, V, @ViewUpdate, B);

  S:= 'Считать общее количество (сумму) рекламаций:';
  V:= VCreateStr([
    'по всем критериям',
    'только по включенным в отчёт критериям'//,
    //'не считать и не отображать'
  ]);
  ParamList.AddStringList('SumTypeList', S, V, @ViewUpdate);

  S:= 'Дополнительно отображать:';
  V:= VCreateStr([
    'гистограммы',
    '% от общего количества рекламаций',
    'данные с общим количеством рекламаций = 0'
  ]);
  B:= VCreateBool(Length(V), True);
  ParamList.AddCheckList('AdditionShow', S, V, @ViewUpdate, B);

  //для отчета по месяцам
  S:= 'Вид отчёта:';
  V:= VCreateStr([
    'распределение количества рекламаций',
    'накопление количества рекламаций'
  ]);
  ParamList.AddStringList('MonthTypeList', S, V, @ViewUpdate);

  //для отчета по наименованиям двигателей
  S:= 'Все наименования электродвигателей одного типа:';
  V:= VCreateStr([
    'считать одним наименованием электродвигателя'
  ]);
  ParamList.AddCheckList('MotorTypeAsSingleName', S, V, @ViewUpdate);
end;

procedure TStatisticForm.LoadStatistic;
var
  BD, ED: TDate;
begin
  if StatSelectedIndex<0 then Exit;               //не выбрана статистика
  if (not ParamList.IsSelected['ReasonList']) and //не выбран ни один критерий
     (ParamList.Selected['SumTypeList']>0) then   //не указано считать кол-во по всем критериям
       Exit;

  if ReportTypeComboBox.ItemIndex=0 then //текущий произвольный период
  begin
    AdditionYearsCount:= 0;
    BD:= BeginDatePicker.Date;
    ED:= EndDateTimePicker.Date;
    PeriodStr:= 'с ' + FormatDateTime('dd.mm.yyyy', BD) +
                ' по ' + FormatDateTime('dd.mm.yyyy', ED);
  end
  else if ReportTypeComboBox.ItemIndex=1 then
  begin
    AdditionYearsCount:= AdditionYearCountSpinEdit.Value;
    BD:= BeginDatePicker.Date; //???
    ED:= EndDateTimePicker.Date;//???
  end;

  MotorNamesStr:= VVectorToStr(MainForm.UsedNames, ', ');
  case StatSelectedIndex of
    0: begin
         ParamNames:= VCut(MainForm.UsedNames);
         DataBase.ReclamationByMotorNamesLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs,
                                  ParamList.Checked['MotorTypeAsSingleName', 0],
                                  not ParamList.Checked['AdditionShow', 2{отображать count=0}],
                                  ParamNames, ParamNeeds, ClaimCounts);
       end;
    1: DataBase.ReclamationByDefectsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs,
                                  not ParamList.Checked['AdditionShow', 2{отображать count=0}],
                                  ParamNames, ParamNeeds, ClaimCounts);
    2: DataBase.ReclamationByPlacesLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs,
                                  not ParamList.Checked['AdditionShow', 2{отображать count=0}],
                                  ParamNames, ParamNeeds, ClaimCounts);
    3: ;
    4: ;
  end;

  DrawStatistic;


  //if not ParamList.IsSelected['ReasonList'] then Exit;
  //
  //if AParamType<0 then Exit;
  //BD:= BeginDatePicker.Date;
  //ED:= EndDateTimePicker.Date;
  //MotorNamesStr:= VVectorToStr(MainForm.UsedNames, ', ');
  //
  //AdditionYearsCount:= 0;
  //if SeveralPeriodsCheckBox.Checked then
  //  AdditionYearsCount:= AdditionYearCountSpinEdit.Value;
  //
  //case AParamType of
  //  0: DataBase.ReclamationMotorsWithReasonsLoad(BD, ED, AdditionYearsCount,
  //                                MainForm.UsedNameIDs, ReasonIDs,
  //                                ParamList.Checked['MotorTypeAsSingleName', 0],
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
  DrawStatistic;
end;

procedure TStatisticForm.DrawStatisticForSinglePeriod;
var
  //Drawer: TStatisticSinglePeriodSheet;
  Drawer: TStatSheet;
  ParamColName, PartTitle, PartTitle2: String;
begin
  case StatSelectedIndex of
    0: begin
         ParamColName:= 'Наименование электродвигателя';
         PartTitle:= 'наименованиям электродвигателей';
         PartTitle2:= 'наименованию электродвигателя';
       end;
    1: begin
         ParamColName:= 'Неисправный элемент';
         PartTitle:= 'неисправным элементам';
         PartTitle2:= 'неисправному элементу';
       end;
    2: begin
         ParamColName:= 'Предприятие';
         PartTitle:= 'предприятиям';
         PartTitle2:= 'предприятию';
       end;
    3: begin

       end;
    4: begin

       end;
  end;

  Drawer:= TStatSheet.Create(ViewGrid.Worksheet, ViewGrid, GridFont);
  try
    Drawer.Zoom(ZoomPercent);
    Drawer.Draw(ParamColName, PartTitle, PartTitle2, MotorNamesStr, PeriodStr,
                ParamList.Checkeds['ReasonList'], ReasonNames,
                ParamNeeds, ParamNames, ClaimCounts,
                ParamList.Checkeds['DataList'],
                ParamList.Selected['SumTypeList'],
                ParamList.Checked['AdditionShow', 0{гистограммы}],
                ParamList.Checked['AdditionShow', 1{% от кол-ва}]//,
                //ParamList.Checked['AdditionShow', 2{подробные гистограммы}]
                );
  finally
    FreeAndNil(Drawer);
  end;





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
  //             MotorNamesStr, ParamNames, ReasonNames, ClaimCounts[0],
  //             ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
  //             ParamList.Checked['AdditionShow', 0] //отображать гистограммы
  //             );
  //finally
  //  FreeAndNil(Drawer);
  //end;
end;

procedure TStatisticForm.DrawStatisticForSeveralPeriods;
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
 //              MotorNamesStr, ParamNames, ReasonNames, ClaimCounts,
 //              ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
 //              ParamList.Checked['AdditionShow', 0] //отображать гистограммы
 //              );
 // finally
 //   FreeAndNil(Drawer);
 // end;
end;

procedure TStatisticForm.DrawStatistic;
begin
  if MIsNil(ClaimCounts) then Exit;
  if ReportTypeComboBox.ItemIndex=0 then
    DrawStatisticForSinglePeriod
  else if ReportTypeComboBox.ItemIndex=1 then
    DrawStatisticForSeveralPeriods;
end;

procedure TStatisticForm.ExportStatistic;
begin
  if MIsNil(ClaimCounts) then Exit;
  if ReportTypeComboBox.ItemIndex=0 then
    ExportStatisticForSinglePeriod
  else if ReportTypeComboBox.ItemIndex=1 then
    ExportStatisticForSeveralPeriods;
end;

procedure TStatisticForm.ExportStatisticForSinglePeriod;
var
  //Drawer: TStatisticSinglePeriodSheet;

  Sheet: TsWorksheet;
  Exporter: TSheetsExporter;
  Drawer: TStatSheet;
  ParamColName, PartTitle, PartTitle2: String;
begin
  case StatSelectedIndex of
    0: begin
         ParamColName:= 'Наименование электродвигателя';
         PartTitle:= 'наименованиям электродвигателей';
         PartTitle2:= 'наименованию электродвигателя';
       end;
    1: begin
         ParamColName:= 'Неисправный элемент';
         PartTitle:= 'неисправным элементам';
         PartTitle2:= 'неисправному элементу';
       end;
    2: begin
         ParamColName:= 'Предприятие';
         PartTitle:= 'предприятиям';
         PartTitle2:= 'предприятию';
       end;
    3: begin

       end;
    4: begin

       end;
  end;

  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    Drawer:= TStatSheet.Create(Sheet, nil, GridFont);
    try
      Drawer.Draw(ParamColName, PartTitle, PartTitle2, MotorNamesStr, PeriodStr,
                  ParamList.Checkeds['ReasonList'], ReasonNames,
                  ParamNeeds, ParamNames, ClaimCounts,
                  ParamList.Checkeds['DataList'],
                  ParamList.Selected['SumTypeList'],
                  ParamList.Checked['AdditionShow', 0{гистограммы}],
                  ParamList.Checked['AdditionShow', 1{% от кол-ва}]//,
                  //ParamList.Checked['AdditionShow', 2{подробные гистограммы}]
                  );
    finally
      FreeAndNil(Drawer);
    end;

    Exporter.PageSettings(spoPortrait, pfWidth, True, False);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;


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
  //                MotorNamesStr, ParamNames, ReasonNames, ClaimCounts[0],
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

procedure TStatisticForm.ExportStatisticForSeveralPeriods;
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
  //                MotorNamesStr, ParamNames, ReasonNames, ClaimCounts,
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

