unit UStatisticForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons, ComCtrls,
  fpspreadsheetgrid, VirtualTrees, LCLType, StdCtrls, Spin, DateTimePicker, DividerBevel,
  DateUtils,
  //DK packages utils
  DK_Vector, DK_VSTTables, DK_DateUtils, DK_Matrix, DK_SheetTypes, DK_Const,
  DK_StrUtils, DK_Zoom, DK_VSTTableTools, DK_CtrlUtils, DK_VSTParamList,
  //Project utils
  UVars, UStatSheets, UStatistic;

type

  { TStatisticForm }

  TStatisticForm = class(TForm)
    AdditionYearCountSpinEdit: TSpinEdit;
    AdditionYearsPanel: TPanel;
    Bevel1: TBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    ReportTypeComboBox: TComboBox;
    ExportButton: TSpeedButton;
    ReportTypePanel: TPanel;
    SettingClientPanel: TPanel;
    EndDatePicker: TDateTimePicker;
    BeginDatePicker: TDateTimePicker;
    YearSpinEdit: TSpinEdit;
    ViewGrid: TsWorksheetGrid;
    Label3: TLabel;
    MainPanel: TPanel;
    LeftPanel: TPanel;
    StatisticPanel: TPanel;
    RightPanel: TPanel;
    GridPanel: TPanel;
    PeriodPanel: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    ToolPanel: TPanel;
    ClientPanel: TPanel;
    StatisticVT: TVirtualStringTree;
    YearPanel: TPanel;
    ZoomPanel: TPanel;
    procedure ExportButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ReportTypeComboBoxChange(Sender: TObject);
    procedure EndDatePickerChange(Sender: TObject);
    procedure BeginDatePickerChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure AdditionYearCountSpinEditChange(Sender: TObject);
    procedure YearSpinEditChange(Sender: TObject);
  private
    CanShow: Boolean;
    ZoomPercent: Integer;
    StatSelectedIndex: Integer;
    StatisticList: TVSTStringList;
    ParamList: TVSTParamList;
    Drawer: TStatSheet;

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

    procedure CreateParamList;

    procedure CreateStatisticList;
    procedure SelectStatistic;
    procedure LoadStatistic;
    procedure ExportStatistic;

    procedure StatisticSettings(out AParamColName, APartTitle, APartTitle2: String;
                                out ANeedAccumCount, ATotalCountHistSort: Boolean);

    procedure Draw(const AZoomPercent: Integer);
    procedure DrawStatistic;

    procedure LimitDatePickers;
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
  Drawer:= TStatSheet.Create(ViewGrid.Worksheet, ViewGrid, GridFont);

  DataBase.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName', ReasonIDs, ReasonNames, False, 'ReasonID');
  ReasonNames[0]:= 'Не расследовано';

  CreateParamList;

  StatSelectedIndex:= -1;
  CreateStatisticList;

  YearSpinEdit.Value:= YearOfDate(Date);
  BeginDatePicker.Date:= FirstDayInYear(Date);
  EndDatePicker.Date:= LastDayInMonth(Date);
  CanShow:= True;
end;

procedure TStatisticForm.FormDestroy(Sender: TObject);
begin
  if Assigned(Drawer) then FreeAndNil(Drawer);
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

procedure TStatisticForm.EndDatePickerChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.BeginDatePickerChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStatisticForm.LimitDatePickers;
begin
  BeginDatePicker.MinDate:= NULDATE;
  BeginDatePicker.MaxDate:= INFDATE;
  EndDatePicker.MinDate:= BeginDatePicker.MinDate;
  EndDatePicker.MaxDate:= BeginDatePicker.MaxDate;

  if ReportTypeComboBox.ItemIndex=1 then //comparison
  begin
    BeginDatePicker.Date:= RecodeYear(BeginDatePicker.Date, YearSpinEdit.Value);
    EndDatePicker.Date:= RecodeYear(EndDatePicker.Date, YearSpinEdit.Value);
    BeginDatePicker.MinDate:= FirstDayInYear(YearSpinEdit.Value);
    BeginDatePicker.MaxDate:= LastDayInYear(YearSpinEdit.Value);
    EndDatePicker.MinDate:= BeginDatePicker.MinDate;
    EndDatePicker.MaxDate:= BeginDatePicker.MaxDate;
  end;
end;

procedure TStatisticForm.ReportTypeComboBoxChange(Sender: TObject);
begin
  YearPanel.Visible:= ReportTypeComboBox.ItemIndex=1;
  AdditionYearsPanel.Visible:= ReportTypeComboBox.ItemIndex=1;
  Application.ProcessMessages;
  LimitDatePickers;
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

procedure TStatisticForm.YearSpinEditChange(Sender: TObject);
begin
  LimitDatePickers;
  ViewUpdate;
end;

procedure TStatisticForm.ViewUpdate;
begin
  if not CanShow then Exit;

  ParamList.Visibles['MotorTypeAsSingleName']:= StatSelectedIndex=0;
  ParamList.Visibles['AccumTotalCounts']:= StatSelectedIndex=3;
  ParamList.Visibles['MileageStepList']:= StatSelectedIndex=4;

  Screen.Cursor:= crHourGlass;
  try
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
  S:= 'Общее количество рекламаций считать как:';
  V:= VCreateStr([
    'распределение количества рекламаций',
    'накопление количества рекламаций'
  ]);
  ParamList.AddStringList('AccumTotalCounts', S, V, @ViewUpdate);

  //для отчета по наименованиям двигателей
  S:= 'Все наименования электродвигателей одного типа:';
  V:= VCreateStr([
    'считать одним наименованием электродвигателя'
  ]);
  ParamList.AddCheckList('MotorTypeAsSingleName', S, V, @ViewUpdate);

  //для отчета по пробегу
  S:= 'Считать пробег локомотива с шагом:';
  V:= VCreateStr([
    '50 тыс.км',
    '25 тыс.км',
    '10 тыс.км'
  ]);
  ParamList.AddStringList('MileageStepList', S, V, @ViewUpdate);
end;

procedure TStatisticForm.LoadStatistic;
var
  BD, ED: TDate;
  i: Integer;
begin
  if StatSelectedIndex<0 then Exit;               //не выбрана статистика
  if (not ParamList.IsSelected['ReasonList']) and //не выбран ни один критерий
     (ParamList.Selected['SumTypeList']>0) then   //не указано считать кол-во по всем критериям
       Exit;

  BD:= BeginDatePicker.Date;
  ED:= EndDatePicker.Date;
  if ReportTypeComboBox.ItemIndex=0 then //текущий произвольный период
  begin
    AdditionYearsCount:= 0;
    PeriodStr:= FormatDateTime('с dd.mm.yyyy', BD) + FormatDateTime(' по dd.mm.yyyy', ED);
  end
  else if ReportTypeComboBox.ItemIndex=1 then
  begin
    AdditionYearsCount:= AdditionYearCountSpinEdit.Value;
    PeriodStr:= IntToStr(YearSpinEdit.Value-AdditionYearsCount);
    for i:= AdditionYearsCount-1 downto 0 do
      PeriodStr:= PeriodStr + '/' + IntToStr(YearSpinEdit.Value-i);
    PeriodStr:= FormatDateTime('с dd ', BD) + MONTHSGEN[MonthOf(BD)] +
                FormatDateTime(' по dd ', ED) + MONTHSGEN[MonthOf(ED)] +
                SYMBOL_SPACE + PeriodStr + ' гг.';
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
    3: DataBase.ReclamationByMonthsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs,
                                  not ParamList.Checked['AdditionShow', 2{отображать count=0}],
                                  ParamNames, ParamNeeds, ClaimCounts);
    4: begin
         //MileageStep
         case ParamList.Selected['MileageStepList'] of
           0: i:= 50;
           1: i:= 25;
           2: i:= 10;
         end;
         DataBase.ReclamationByMileagesLoad(BD, ED, i, AdditionYearsCount,
                                  MainForm.UsedNameIDs,
                                  not ParamList.Checked['AdditionShow', 2{отображать count=0}],
                                  ParamNames, ParamNeeds, ClaimCounts);
       end;
  end;

  DrawStatistic;
end;

procedure TStatisticForm.StatisticSettings(out AParamColName, APartTitle, APartTitle2: String;
                                           out ANeedAccumCount, ATotalCountHistSort: Boolean);
begin
  case StatSelectedIndex of
    0: begin
         AParamColName:= 'Наименование электродвигателя';
         APartTitle:= 'наименованиям электродвигателей';
         APartTitle2:= 'наименованию электродвигателя';
         ANeedAccumCount:= False;
         ATotalCountHistSort:= True;
       end;
    1: begin
         AParamColName:= 'Неисправный элемент';
         APartTitle:= 'неисправным элементам';
         APartTitle2:= 'неисправному элементу';
         ANeedAccumCount:= False;
         ATotalCountHistSort:= True;
       end;
    2: begin
         AParamColName:= 'Предприятие';
         APartTitle:= 'предприятиям';
         APartTitle2:= 'предприятию';
         ANeedAccumCount:= False;
         ATotalCountHistSort:= True;
       end;
    3: begin
         AParamColName:= 'Месяц';
         APartTitle:= 'месяцам';
         APartTitle2:= 'месяцу';
         ANeedAccumCount:= ParamList.Selected['AccumTotalCounts']=1;
         ATotalCountHistSort:= False;
       end;
    4: begin
         AParamColName:= 'Пробег локомотива, тыс. км';
         APartTitle:= 'пробегам локомотва';
         APartTitle2:= 'пробегу локомотива';
         ANeedAccumCount:= False;
         ATotalCountHistSort:= False;
       end;
  end;
end;

procedure TStatisticForm.Draw(const AZoomPercent: Integer);
begin
  ZoomPercent:= AZoomPercent;
  DrawStatistic;
end;

procedure TStatisticForm.DrawStatistic;
var
  ParamColName, PartTitle, PartTitle2: String;
  NeedAccumCount, TotalCountHistSort: Boolean;
begin
  if MIsNil(ClaimCounts) then Exit;

  StatisticSettings(ParamColName, PartTitle, PartTitle2, NeedAccumCount, TotalCountHistSort);
  Drawer.Zoom(ZoomPercent);
  if ReportTypeComboBox.ItemIndex=0 then
    Drawer.PeriodDraw(ParamColName, PartTitle, PartTitle2, MotorNamesStr, PeriodStr,
                      ParamList.Checkeds['ReasonList'], ReasonNames,
                      ParamNeeds, ParamNames, ClaimCounts,
                      ParamList.Checkeds['DataList'],
                      ParamList.Selected['SumTypeList'],
                      ParamList.Checked['AdditionShow', 0{гистограммы}],
                      ParamList.Checked['AdditionShow', 1{% от кол-ва}],
                      NeedAccumCount, TotalCountHistSort)
  else if ReportTypeComboBox.ItemIndex=1 then
    Drawer.ComparisonDraw(YearSpinEdit.Value, AdditionYearCountSpinEdit.Value,
                      ParamColName, PartTitle, PartTitle2, MotorNamesStr, PeriodStr,
                      ParamList.Checkeds['ReasonList'], ReasonNames,
                      ParamNeeds, ParamNames, ClaimCounts,
                      //ParamList.Checkeds['DataList'],
                      ParamList.Selected['SumTypeList'],
                      ParamList.Checked['AdditionShow', 0{гистограммы}],
                      ParamList.Checked['AdditionShow', 1{% от кол-ва}]);
end;

procedure TStatisticForm.ExportStatistic;
begin
  if MIsNil(ClaimCounts) then Exit;
  Drawer.Save('Лист1', 'Выполнено!', spoPortrait, pfWidth,
              False{no headers}, False{no grid lines});
end;

end.

