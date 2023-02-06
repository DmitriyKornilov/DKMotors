unit UStatisticForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  fpspreadsheetgrid, VirtualTrees, DK_Vector, DK_VSTTables, USheetUtils, LCLType,
  StdCtrls, Spin, ComCtrls, DateTimePicker, DK_DateUtils, rxctrls, DividerBevel,
  USQLite, DK_Matrix, DK_SheetExporter, DK_SheetConst, DK_StrUtils;

type

  { TStatisticForm }

  TStatisticForm = class(TForm)
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
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
    Label8: TLabel;
    MainPanel: TPanel;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    ReportPeriodPanel: TPanel;
    SpinEdit1: TSpinEdit;
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
    procedure CheckBox1Change(Sender: TObject);
    procedure CheckBox2Change(Sender: TObject);
    procedure CheckBox3Change(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
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
    ParamCaption, ReportName, MotorNames: String;
    ParamNameColumnWidth, GroupType, SecondColumnType, AdditionYearsCount: Integer;

    procedure VerifyDates;

    procedure SetStatisticList;
    procedure StatisticListSelectItem;

    procedure SetReasonList;
    procedure ReasonListSelectItem(const AIndex: Integer; const AChecked: Boolean);

    //AParamType
    //0 - Распределение по наименованиям двигателей
    //1 - Распределение по неисправным элементам
    //2 - Распределение по предприятиям
    //3 - Распределение по месяцам
    procedure Statistic(const AParamType: Integer);
    procedure DrawStatistic;
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
var
  Exporter: TGridExporter;
  PageOrientation: TsPageOrientation;
begin
  Exporter:= TGridExporter.Create(Grid1);
  try
    //Exporter.SheetName:= 'Отчет';
    PageOrientation:= spoPortrait;
    Exporter.PageSettings(PageOrientation, pfWidth);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
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

  if CheckBox1.Checked or (SelectedIndex=3) then
  begin
    if YearOfDate(BD)<>YearOfDate(ED) then
      ED:= LastDayInYear(BD);
  end;

  DateTimePicker2.Date:= BD;
  DateTimePicker1.Date:= ED;

  CanShow:= True;
end;

procedure TStatisticForm.CheckBox1Change(Sender: TObject);
begin
  SpinEdit1.Enabled:= CheckBox1.Checked;
  CheckBox3.Visible:= CheckBox1.Checked;
  Panel6.Visible:= CheckBox1.Checked;

  ShowStatistic;
end;

procedure TStatisticForm.CheckBox2Change(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.CheckBox3Change(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.FormDestroy(Sender: TObject);
begin
  if Assigned(StatisticList) then FreeAndNil(StatisticList);
  if Assigned(ReasonList) then FreeAndNil(ReasonList);
end;

procedure TStatisticForm.RadioButton1Click(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.RadioButton2Click(Sender: TObject);
begin
  ShowStatistic;
end;

procedure TStatisticForm.SpinEdit1Change(Sender: TObject);
begin
  if SpinEdit1.Value=1 then
    Label4.Caption:= 'предыдущий год'
  else if (SpinEdit1.Value>1) and (SpinEdit1.Value<5) then
    Label4.Caption:= 'предыдущиx года'
  else
    Label4.Caption:= 'предыдущиx лет';
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
  DrawStatistic;
end;

procedure TStatisticForm.ShowStatistic;
begin
  if not CanShow then Exit;
  VerifyDates;

  Screen.Cursor:= crHourGlass;
  try
    Grid1.Clear;
    Statistic(SelectedIndex);
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
  if SelectedIndex=3 then
    CheckBox2.Caption:= 'накопление количества по месяцам'
  else
    CheckBox2.Caption:= '% от общего количества рекламаций';
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

procedure TStatisticForm.Statistic(const AParamType: Integer);
var
  BD, ED: TDate;
  i,W: Integer;
begin
  if AParamType<0 then Exit;

  BD:= DateTimePicker2.Date;
  ED:= DateTimePicker1.Date;

  //if RadioButton1.Checked then
  GroupType:= 2;
  if RadioButton2.Checked then
    GroupType:= 1;

  SecondColumnType:= 1;
  if SelectedIndex=3 then
    SecondColumnType:= 2;

  AdditionYearsCount:= 0;
  if CheckBox1.Checked then
    AdditionYearsCount:= SpinEdit1.Value;

  if AParamType=0 then
  begin
    SQLite.ReclamationMotorsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
    ParamCaption:= 'Наименование';
    ReportName:= 'распределение по наименованиям электродвигателей';
  end
  else if AParamType=1 then
  begin
    SQLite.ReclamationDefectsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
    ParamCaption:= 'Неисправность';
    ReportName:= 'распределение по неисправным элементам';
  end
  else if AParamType=2 then
  begin
    SQLite.ReclamationPlacesWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
    ParamCaption:= 'Предприятие';
    ReportName:= 'распределение по предприятиям';
  end
  else if AParamType=3 then
  begin
    SQLite.ReclamationMonthsWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
    ParamCaption:= 'Месяц';
    ReportName:= 'распределение по месяцам';
  end
  else if AParamType=4 then
  begin
    SQLite.ReclamationMileagesWithReasonsLoad(BD, ED, AdditionYearsCount,
                                MainForm.UsedNameIDs, ReasonIDs, ParamNames, Counts);
    ParamCaption:= 'Пробег локомотива, тыс. км';
    ReportName:= 'распределение по пробегу локомотива';
  end;

  MotorNames:= VVectorToStr(MainForm.UsedNames, ', ');

  //ширина столбца с наименованиями параметра распределения
  ParamNameColumnWidth:= 120;
  for i:=0 to High(ParamNames) do
  begin
    W:= SWidth(ParamNames[i], SHEET_FONT_NAME, SHEET_FONT_SIZE)+50;
    W:= Round(W/DIMENTION_FACTOR);
    if W>ParamNameColumnWidth then
      ParamNameColumnWidth:= W;

  end;

  DrawStatistic;
end;

procedure TStatisticForm.DrawStatistic;
var
  BD, ED: TDate;
  Sheet: TStatisticReclamationSheet;
begin
  Sheet:= TStatisticReclamationSheet.Create(Grid1,
                    ParamNameColumnWidth, GroupType, SecondColumnType, AdditionYearsCount,
                    ReasonList.Selected, CheckBox2.Checked, CheckBox3.Checked);
  try
    BD:= DateTimePicker2.Date;
    ED:= DateTimePicker1.Date;
    Sheet.Zoom(ZoomTrackBar.Position);
    Sheet.Draw(BD, ED, ReportName, MotorNames, ParamCaption,
             ParamNames, ReasonNames, Counts, True);
  finally
    FreeAndNil(Sheet);
  end;
end;






end.

