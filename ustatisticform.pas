unit UStatisticForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  fpspreadsheetgrid, VirtualTrees, DK_Vector, DK_VSTTables, USheetUtils, LCLType,
  StdCtrls, Spin, ComCtrls, DateTimePicker, DK_DateUtils, rxctrls, DividerBevel,
  USQLite, DK_Matrix, DK_SheetExporter, DK_SheetConst;

type

  { TStatisticForm }

  TStatisticForm = class(TForm)
    ANEMAsSameNameCheckBox: TCheckBox;
    MonthPanel: TPanel;
    Label10: TLabel;
    ANEMPanel: TPanel;
    Label11: TLabel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    SeveralPeriodsCheckBox: TCheckBox;
    ShowPercentCheckBox: TCheckBox;
    ShowGraphicsCheckBox: TCheckBox;
    ShowLinePercentCheckBox: TCheckBox;
    TotalCountForUsedParamsOnlyCheckBox: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    DividerBevel3: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Grid1: TsWorksheetGrid;
    Label1: TLabel;
    Label2: TLabel;
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
    Panel8: TPanel;
    Panel9: TPanel;
    ReportPeriodPanel: TPanel;
    AdditionYearCountSpinEdit: TSpinEdit;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    TopPanel: TPanel;
    ClientPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    ZoomCaptionLabel: TLabel;
    ZoomInButton: TSpeedButton;
    ZoomOutButton: TSpeedButton;
    ZoomPanel: TPanel;
    ZoomTrackBar: TTrackBar;
    ZoomValueLabel: TLabel;
    ZoomValuePanel: TPanel;
    procedure ANEMAsSameNameCheckBoxChange(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure SeveralPeriodsCheckBoxChange(Sender: TObject);
    procedure ShowGraphicsCheckBoxChange(Sender: TObject);
    procedure ShowLinePercentCheckBoxChange(Sender: TObject);

    procedure ShowPercentCheckBoxChange(Sender: TObject);

    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

    procedure AdditionYearCountSpinEditChange(Sender: TObject);
    procedure TotalCountForUsedParamsOnlyCheckBoxChange(Sender: TObject);
    procedure ZoomInButtonClick(Sender: TObject);
    procedure ZoomOutButtonClick(Sender: TObject);
    procedure ZoomTrackBarChange(Sender: TObject);

  private
    SelectedIndex: Integer;
    CanShow: Boolean;

    ReasonIDs: TIntVector;
    ReasonNames: TStrVector;
    ReasonList: TVSTCheckTable;

    StatisticList: TVSTTable;

    ParamNames: TStrVector;
    Counts: TIntMatrix3D;
    MotorNames: String;
    AdditionYearsCount: Integer;


    procedure VerifyDates;

    procedure SetStatisticList;
    procedure StatisticListSelectItem;

    procedure SetReasonList;
    procedure ReasonListSelectItem(const {%H-}AIndex: Integer; const {%H-}AChecked: Boolean);

    //AParamType
    //0 - Распределение по наименованиям двигателей
    //1 - Распределение по неисправным элементам
    //2 - Распределение по предприятиям
    //3 - Распределение по месяцам
    procedure LoadStatistic(const AParamType: Integer);

    procedure DrawStatistic(const AParamType: Integer);
    procedure DrawStatisticForSinglePeriod(const AParamType: Integer);
    procedure DrawStatisticForSeveralPeriods(const AParamType: Integer);

    procedure ExportStatistic(const AParamType: Integer);
    procedure ExportStatisticForSinglePeriod(const AParamType: Integer);
    procedure ExportStatisticForSeveralPeriods(const AParamType: Integer);

  public
    procedure ShowStatistic;
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

  SelectedIndex:= -1;
  MainForm.SetNamesPanelsVisible(True, False);


  ReasonList:= TVSTCheckTable.Create(VT2);
  ReasonList.OnSelect:= @ReasonListSelectItem;
  SetReasonList;

  StatisticList:= TVSTTable.Create(VT1);
  StatisticList.OnSelect:= @StatisticListSelectItem;
  SetStatisticList;

  DateTimePicker2.Date:= FirstDayInYear(Date);
  CanShow:= True;
  DateTimePicker1.Date:= LastDayInMonth(Date);

  ShowStatistic;
end;

procedure TStatisticForm.ExportButtonClick(Sender: TObject);
begin
  ExportStatistic(SelectedIndex);
end;

procedure TStatisticForm.DateTimePicker1Change(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.DateTimePicker2Change(Sender: TObject);
begin
  ShowStatistic;
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
  ShowStatistic;
end;

procedure TStatisticForm.ShowGraphicsCheckBoxChange(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.ShowLinePercentCheckBoxChange(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.ANEMAsSameNameCheckBoxChange(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.RadioButton1Click(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.RadioButton2Click(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.ShowPercentCheckBoxChange(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.FormDestroy(Sender: TObject);
begin
  if Assigned(StatisticList) then FreeAndNil(StatisticList);
  if Assigned(ReasonList) then FreeAndNil(ReasonList);
end;

procedure TStatisticForm.AdditionYearCountSpinEditChange(Sender: TObject);
begin
  if AdditionYearCountSpinEdit.Value=1 then
    Label4.Caption:= 'предыдущий год'
  else if (AdditionYearCountSpinEdit.Value>1) and (AdditionYearCountSpinEdit.Value<5) then
    Label4.Caption:= 'предыдущиx года'
  else
    Label4.Caption:= 'предыдущиx лет';
  ShowStatistic;
end;

procedure TStatisticForm.TotalCountForUsedParamsOnlyCheckBoxChange(
  Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.ZoomInButtonClick(Sender: TObject);
begin
  ZoomTrackBar.Position:= ZoomTrackBar.Position + 5;
end;

procedure TStatisticForm.ZoomOutButtonClick(Sender: TObject);
begin
  ZoomTrackBar.Position:= ZoomTrackBar.Position - 5;
end;

procedure TStatisticForm.ZoomTrackBarChange(Sender: TObject);
begin
  ZoomValueLabel.Caption:= IntToStr(ZoomTrackBar.Position) + ' %';
  DrawStatistic(SelectedIndex);
end;

procedure TStatisticForm.ShowStatistic;
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

procedure TStatisticForm.SetStatisticList;
var
  V: TStrVector;
begin
  V:= VCreateStr([
    'Распределение по наименованиям двигателей',
    'Распределение по неисправным элементам',
    'Распределение по предприятиям',
    'Распределение по месяцам',
    'Распределение по пробегу локомотива'
  ]);

  StatisticList.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  StatisticList.HeaderVisible:= False;
  StatisticList.GridLinesVisible:= False;
  StatisticList.CanSelect:= True;
  StatisticList.CanUnselect:= False;
  StatisticList.AddColumn('Список');
  StatisticList.SetColumn('Список', V, taLeftJustify);
  StatisticList.Draw;
  StatisticList.Select(0);
end;

procedure TStatisticForm.StatisticListSelectItem;
begin
  if SelectedIndex=StatisticList.SelectedIndex then Exit;
  SelectedIndex:= StatisticList.SelectedIndex;
  ANEMPanel.Visible:= SelectedIndex=0;
  MonthPanel.Visible:= SelectedIndex=3;
  ShowStatistic;
end;

procedure TStatisticForm.SetReasonList;
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

  ReasonList.GridLinesVisible:= False;
  ReasonList.HeaderVisible:= False;
  ReasonList.SelectedBGColor:= VT2.Color;
  ReasonList.AddColumn('Список');
  ReasonList.SetColumn('Список', ReasonNames, taLeftJustify);
  ReasonList.Draw;
  ReasonList.CheckAll(True);
end;

procedure TStatisticForm.ReasonListSelectItem(const AIndex: Integer; const AChecked: Boolean);
begin
  ShowStatistic;
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
  0: SQLite.ReclamationMotorsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs,
                                ANEMAsSameNameCheckBox.Checked, ParamNames, Counts);
  1: SQLite.ReclamationDefectsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
  2: SQLite.ReclamationPlacesWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
  3: SQLite.ReclamationMonthsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
  4: SQLite.ReclamationMileagesWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
  end;

  DrawStatistic(AParamType);
end;

procedure TStatisticForm.DrawStatisticForSinglePeriod(const AParamType: Integer);
var
  Drawer: TStatisticSinglePeriodSheet;
begin
  case AParamType of
  0: Drawer:= TStatisticSinglePeriodAtMotorNamesSheet.Create(
                Grid1.Worksheet, Grid1, ReasonList.Selected,
                ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
  1: Drawer:= TStatisticSinglePeriodAtDefectNamesSheet.Create(
                Grid1.Worksheet, Grid1, ReasonList.Selected,
                ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
  2: Drawer:= TStatisticSinglePeriodAtPlaceNamesSheet.Create(
                Grid1.Worksheet, Grid1, ReasonList.Selected,
                ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
  3: if RadioButton1.Checked then
       Drawer:= TStatisticSinglePeriodAtMonthNamesSheet.Create(
                Grid1.Worksheet, Grid1, ReasonList.Selected,
                ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked)
     else
       Drawer:= TStatisticSinglePeriodAtMonthNamesSumSheet.Create(
                Grid1.Worksheet, Grid1, ReasonList.Selected,
                ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
  4: Drawer:= TStatisticSinglePeriodAtMileagesSheet.Create(
                Grid1.Worksheet, Grid1, ReasonList.Selected,
                ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
  end;

  try
    Drawer.Zoom(ZoomTrackBar.Position);
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
                ReasonList.Selected, ShowPercentCheckBox.Checked);
  1: Drawer:= TStatisticSeveralPeriodsAtDefectNamesSheet.Create(
                Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                ReasonList.Selected, ShowPercentCheckBox.Checked);
  2: Drawer:= TStatisticSeveralPeriodsAtPlaceNamesSheet.Create(
                Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                ReasonList.Selected, ShowPercentCheckBox.Checked);
  3: if RadioButton1.Checked then
       Drawer:= TStatisticSeveralPeriodsAtMonthNamesSheet.Create(
                Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                ReasonList.Selected, ShowPercentCheckBox.Checked)
     else
       Drawer:= TStatisticSeveralPeriodsAtMonthNamesSumSheet.Create(
                Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                ReasonList.Selected, ShowPercentCheckBox.Checked);
  4: Drawer:= TStatisticSeveralPeriodsAtMileagesSheet.Create(
                Grid1.Worksheet, Grid1, AdditionYearCountSpinEdit.Value,
                ReasonList.Selected, ShowPercentCheckBox.Checked);
 end;

  try
    Drawer.Zoom(ZoomTrackBar.Position);
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
  Exporter: TSheetExporter;
begin
  Exporter:= TSheetExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');

    case AParamType of
    0: Drawer:= TStatisticSinglePeriodAtMotorNamesSheet.Create(
                  Sheet, nil, ReasonList.Selected,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    1: Drawer:= TStatisticSinglePeriodAtDefectNamesSheet.Create(
                  Sheet, nil, ReasonList.Selected,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    2: Drawer:= TStatisticSinglePeriodAtPlaceNamesSheet.Create(
                  Sheet, nil, ReasonList.Selected,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    3: Drawer:= TStatisticSinglePeriodAtMonthNamesSheet.Create(
                  Sheet, nil, ReasonList.Selected,
                  ShowPercentCheckBox.Checked, ShowLinePercentCheckBox.Checked);
    4: Drawer:= TStatisticSinglePeriodAtMileagesSheet.Create(
                  Sheet, nil, ReasonList.Selected,
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
  Exporter: TSheetExporter;
begin
  Exporter:= TSheetExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');

    case AParamType of
    0: Drawer:= TStatisticSeveralPeriodsAtMotorNamesSheet.Create(
                    Sheet, nil, AdditionYearCountSpinEdit.Value,
                    ReasonList.Selected, ShowPercentCheckBox.Checked);
    1: Drawer:= TStatisticSeveralPeriodsAtDefectNamesSheet.Create(
                    Sheet, nil, AdditionYearCountSpinEdit.Value,
                    ReasonList.Selected, ShowPercentCheckBox.Checked);
    2: Drawer:= TStatisticSeveralPeriodsAtPlaceNamesSheet.Create(
                    Sheet, nil, AdditionYearCountSpinEdit.Value,
                    ReasonList.Selected, ShowPercentCheckBox.Checked);
    3: Drawer:= TStatisticSeveralPeriodsAtMonthNamesSheet.Create(
                    Sheet, nil, AdditionYearCountSpinEdit.Value,
                    ReasonList.Selected, ShowPercentCheckBox.Checked);
    4: Drawer:= TStatisticSeveralPeriodsAtMileagesSheet.Create(
                    Sheet, nil, AdditionYearCountSpinEdit.Value,
                    ReasonList.Selected, ShowPercentCheckBox.Checked);
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

