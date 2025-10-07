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
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    ExportButton: TSpeedButton;
    SettingClientPanel: TPanel;
    SeveralPeriodsCheckBox: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    ViewGrid: TsWorksheetGrid;
    Label3: TLabel;
    Label4: TLabel;
    MainPanel: TPanel;
    LeftPanel: TPanel;
    StatisticPanel: TPanel;
    RightPanel: TPanel;
    Panel5: TPanel;
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
    procedure SeveralPeriodsCheckBoxChange(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure AdditionYearCountSpinEditChange(Sender: TObject);
  private
    SelectedIndex: Integer;
    CanShow: Boolean;

    ReasonIDs: TIntVector;
    ReasonNames: TStrVector;

    ParamList: TVSTParamList;
    StatisticList: TVSTStringList;

    ParamNames: TStrVector;
    MotorCounts: TIntMatrix3D;
    MotorNames: String;
    AdditionYearsCount: Integer;

    ZoomPercent: Integer;

    procedure VerifyDates;

    procedure CreateStatisticList;
    procedure SelectStatistic;

    procedure LoadReasons;
    procedure CreateParamList;

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

  //to use!!!!
  //DataBase.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName', ReasonIDs, ReasonNames);
  LoadReasons; //to delete!!!

  CreateParamList;

  SelectedIndex:= -1;
  CreateStatisticList;

  DateTimePicker2.Date:= FirstDayInYear(Date);
  DateTimePicker1.Date:= LastDayInMonth(Date);
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
  ParamList.ItemVisibles['AdditionShow']:= VCreateBool([True, True, not SeveralPeriodsCheckBox.Checked]);
  ViewUpdate;
end;

procedure TStatisticForm.ExportButtonClick(Sender: TObject);
begin
  ExportStatistic(SelectedIndex);
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

procedure TStatisticForm.LoadReasons;
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
                                  ParamList.Checked['ANEMIsSingleName', 0],
                                  ParamNames, MotorCounts);
    1: DataBase.ReclamationDefectsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs, ParamNames, MotorCounts);
    2: DataBase.ReclamationPlacesWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs, ParamNames, MotorCounts);
    3: DataBase.ReclamationMonthsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs, ParamNames, MotorCounts);
    4: DataBase.ReclamationMileagesWithReasonsLoad(BD, ED, AdditionYearsCount,
                                  MainForm.UsedNameIDs, ReasonIDs, ParamNames, MotorCounts);
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





  case AParamType of
    0: Drawer:= TStatisticSinglePeriodAtMotorNamesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
    1: Drawer:= TStatisticSinglePeriodAtDefectNamesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
    2: Drawer:= TStatisticSinglePeriodAtPlaceNamesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
    3: if ParamList.Selected['MonthTypeList']=0 then //распределение по месяцам
         Drawer:= TStatisticSinglePeriodAtMonthNamesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  )
       else // накопление по месяцам
         Drawer:= TStatisticSinglePeriodAtMonthNamesSumSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
    4: Drawer:= TStatisticSinglePeriodAtMileagesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
  end;

  try
    Drawer.Zoom(ZoomPercent);
    Drawer.Draw(DateTimePicker2.Date, DateTimePicker1.Date,
               MotorNames, ParamNames, ReasonNames, MotorCounts[0],
               ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
               ParamList.Checked['AdditionShow', 0] //отображать гистограммы
               );
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
                  ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
    1: Drawer:= TStatisticSeveralPeriodsAtDefectNamesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
    2: Drawer:= TStatisticSeveralPeriodsAtPlaceNamesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
    3: if ParamList.Selected['MonthTypeList']=0 then //распределение по месяцам
         Drawer:= TStatisticSeveralPeriodsAtMonthNamesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  )
       else  // накопление по месяцам
         Drawer:= TStatisticSeveralPeriodsAtMonthNamesSumSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
    4: Drawer:= TStatisticSeveralPeriodsAtMileagesSheet.Create(
                  ViewGrid.Worksheet, ViewGrid, GridFont, AdditionYearCountSpinEdit.Value,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
 end;

  try
    Drawer.Zoom(ZoomPercent);
    Drawer.Draw(DateTimePicker2.Date, DateTimePicker1.Date,
               MotorNames, ParamNames, ReasonNames, MotorCounts,
               ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
               ParamList.Checked['AdditionShow', 0] //отображать гистограммы
               );
  finally
    FreeAndNil(Drawer);
  end;
end;

procedure TStatisticForm.DrawStatistic(const AParamType: Integer);
begin
  if MIsNil(MotorCounts) then Exit;
  if SeveralPeriodsCheckBox.Checked then
    DrawStatisticForSeveralPeriods(AParamType)
  else
    DrawStatisticForSinglePeriod(AParamType);
end;

procedure TStatisticForm.ExportStatistic(const AParamType: Integer);
begin
  if MIsNil(MotorCounts) then Exit;
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


  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    case AParamType of
      0: Drawer:= TStatisticSinglePeriodAtMotorNamesSheet.Create(Sheet, nil, GridFont,
                    ParamList.Checkeds['ReasonList'],
                    ParamList.Checked['AdditionShow', 1], //% от суммы за период
                    ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
      1: Drawer:= TStatisticSinglePeriodAtDefectNamesSheet.Create(Sheet, nil, GridFont,
                    ParamList.Checkeds['ReasonList'],
                    ParamList.Checked['AdditionShow', 1], //% от суммы за период
                    ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
      2: Drawer:= TStatisticSinglePeriodAtPlaceNamesSheet.Create(Sheet, nil, GridFont,
                    ParamList.Checkeds['ReasonList'],
                    ParamList.Checked['AdditionShow', 1], //% от суммы за период
                    ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
      3: if ParamList.Selected['MonthTypeList']=0 then //распределение по месяцам
         Drawer:= TStatisticSinglePeriodAtMonthNamesSheet.Create(Sheet, nil, GridFont,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  )
       else // накопление по месяцам
         Drawer:= TStatisticSinglePeriodAtMonthNamesSumSheet.Create(Sheet, nil, GridFont,
                  ParamList.Checkeds['ReasonList'],
                  ParamList.Checked['AdditionShow', 1], //% от суммы за период
                  ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
      4: Drawer:= TStatisticSinglePeriodAtMileagesSheet.Create(Sheet, nil, GridFont,
                    ParamList.Checkeds['ReasonList'],
                    ParamList.Checked['AdditionShow', 1], //% от суммы за период
                    ParamList.Checked['AdditionShow', 2]  //% от суммы по строке
                  );
    end;

    try
      Drawer.Draw(DateTimePicker2.Date, DateTimePicker1.Date,
                  MotorNames, ParamNames, ReasonNames, MotorCounts[0],
                  ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
                  ParamList.Checked['AdditionShow', 0] //отображать гистограммы
                  );

    finally
      FreeAndNil(Drawer);
    end;
    Exporter.PageSettings(spoPortrait, pfWidth, True, False);
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
                      Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
                      ParamList.Checkeds['ReasonList'],
                      ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
      1: Drawer:= TStatisticSeveralPeriodsAtDefectNamesSheet.Create(
                      Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
                      ParamList.Checkeds['ReasonList'],
                      ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
      2: Drawer:= TStatisticSeveralPeriodsAtPlaceNamesSheet.Create(
                      Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
                      ParamList.Checkeds['ReasonList'],
                      ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
      3: if ParamList.Selected['MonthTypeList']=0 then //распределение по месяцам
           Drawer:= TStatisticSeveralPeriodsAtMonthNamesSheet.Create(
                    Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
                    ParamList.Checkeds['ReasonList'],
                    ParamList.Checked['AdditionShow', 1] //% от суммы за период
                    )
         else  // накопление по месяцам
           Drawer:= TStatisticSeveralPeriodsAtMonthNamesSumSheet.Create(
                    Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
                    ParamList.Checkeds['ReasonList'],
                    ParamList.Checked['AdditionShow', 1] //% от суммы за период
                    );
      4: Drawer:= TStatisticSeveralPeriodsAtMileagesSheet.Create(
                      Sheet, nil, GridFont, AdditionYearCountSpinEdit.Value,
                      ParamList.Checkeds['ReasonList'],
                      ParamList.Checked['AdditionShow', 1] //% от суммы за период
                  );
    end;

    try
      Drawer.Draw(DateTimePicker2.Date, DateTimePicker1.Date,
                  MotorNames, ParamNames, ReasonNames, MotorCounts,
                  ParamList.Selected['SumTypeList']=1, //общее кол-во только по включенным в отчёт критериям
                  ParamList.Checked['AdditionShow', 0] //отображать гистограммы
                  );

    finally
      FreeAndNil(Drawer);
    end;
    Exporter.PageSettings(spoLandscape, pfWidth, True, False);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

end.

