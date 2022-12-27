unit UStatisticCountForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  VirtualTrees, rxctrls, fpspreadsheetgrid, DividerBevel, DateTimePicker,
  DK_DateUtils, SheetUtils, DK_Vector, DK_Matrix, USQLite, DK_SheetExporter,
  fpstypes, DK_VSTTables;

type

  { TStatisticCountForm }

  TStatisticCountForm = class(TForm)
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    DividerBevel3: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Grid1: TsWorksheetGrid;
    Label1: TLabel;
    Panel1: TPanel;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    Panel4: TPanel;
    Panel6: TPanel;
    ReportPeriodPanel: TPanel;
    VT1: TVirtualStringTree;

    procedure CheckBox1Change(Sender: TObject);
    procedure CheckBox2Change(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);


  private
    Sheet: TStatisticReclamationCountSheet;
    CategoryList: TVSTTable;

    TitleReasonIDs: TIntVector;
    TitleReasonNames: TStrVector;

    procedure SetCategoryList;
    procedure CategoryListSelectItem;
  public
    procedure ShowData;
  end;

var
  StatisticCountForm: TStatisticCountForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TStatisticCountForm }

procedure TStatisticCountForm.FormCreate(Sender: TObject);
begin
  SQLite.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
                     TitleReasonIDs, TitleReasonNames, CheckBox2.Checked);

  CategoryList:= TVSTTable.Create(VT1);
  CategoryList.OnSelect:= @CategoryListSelectItem;
  SetCategoryList;

  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= FirstDayInYear(Date);
end;

procedure TStatisticCountForm.CheckBox1Change(Sender: TObject);
begin
  ShowData;
end;

procedure TStatisticCountForm.CheckBox2Change(Sender: TObject);
begin
  SQLite.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
                     TitleReasonIDs, TitleReasonNames, CheckBox2.Checked);
  ShowData;
end;

procedure TStatisticCountForm.ExportButtonClick(Sender: TObject);
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

procedure TStatisticCountForm.FormDestroy(Sender: TObject);
begin
  if Assigned(Sheet) then FreeAndNil(Sheet);
  if Assigned(CategoryList) then FreeAndNil(CategoryList);
end;

procedure TStatisticCountForm.SetCategoryList;
var
  V: TStrVector;
begin
  V:= VCreateStr([
    'Электродвигатели',
    'Неисправности',
    'Предприятия'
  ]);

  CategoryList.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  CategoryList.HeaderVisible:= False;
  CategoryList.CanSelect:= True;
  CategoryList.CanUnselect:= False;
  CategoryList.AddColumn('Список');
  CategoryList.SetColumn('Список', V, taLeftJustify);
  CategoryList.Draw;
  CategoryList.Select(0);

end;

procedure TStatisticCountForm.CategoryListSelectItem;
begin
  ShowData;
end;

procedure TStatisticCountForm.ShowData;
var
  BD, ED: TDate;
  ReportName, FirstColName, MotorNames: String;
  NameValues: TStrVector;
  CountValues, JointlyReasonIDs: TIntVector;
  ReasonCountValues: TIntMatrix;
  N,i: Integer;
  ReasonNames: TStrVector;
  JointlyReasonName: String;
begin
  if not CategoryList.IsSelected then Exit;

  ED:= DateTimePicker1.Date;
  BD:= DateTimePicker2.Date;
  if BD>ED then Exit;

  //Считать вместе неисправности локомотива и неправильную эксплуатацию
  JointlyReasonIDs:= nil;
  JointlyReasonName:= EmptyStr;
  if CheckBox1.Checked then
  begin
    VDim(JointlyReasonIDs, 2);
    N:= VIndexOf(TitleReasonNames, 'Неправильная эксплуатация');
    if N>=0 then
    begin
      JointlyReasonIDs[0]:= TitleReasonIDs[N];
      if JointlyReasonName=EmptyStr then
        JointlyReasonName:= 'Неправильная эксплуатация'
      else
        JointlyReasonName:= JointlyReasonName + ', ' + 'Неправильная эксплуатация';
    end;
    N:= VIndexOf(TitleReasonNames, 'Неисправность локомотива');
    if N>=0 then
    begin
      JointlyReasonIDs[1]:= TitleReasonIDs[N];
      if JointlyReasonName=EmptyStr then
        JointlyReasonName:= 'Неисправность локомотива'
      else
        JointlyReasonName:= JointlyReasonName + ', ' + 'Неисправность локомотива';
    end;
  end;

  ReportName:= EmptyStr;

  if CategoryList.SelectedIndex=0 then
  begin
    SQLite.ReclamationTotalWithReasonsLoad(BD, ED,
                   MainForm.UsedNameIDs, TitleReasonIDs, JointlyReasonIDs,
                   NameValues, CountValues, ReasonCountValues);
    FirstColName:= 'Наименование';
    ReportName:= 'общее количество';
  end
  else if CategoryList.SelectedIndex=1 then
  begin
    SQLite.ReclamationDefectsWithReasonsLoad(BD, ED,
                   MainForm.UsedNameIDs, TitleReasonIDs, JointlyReasonIDs,
                   NameValues, CountValues, ReasonCountValues);
    FirstColName:= 'Неисправность';
    ReportName:= 'распределение по неисправным элементам';
  end
  else if CategoryList.SelectedIndex=2 then
  begin
    SQLite.ReclamationPlacesWithReasonsLoad(BD, ED,
                   MainForm.UsedNameIDs, TitleReasonIDs, JointlyReasonIDs,
                   NameValues, CountValues, ReasonCountValues);
    FirstColName:= 'Предприятие';
    ReportName:= 'распределение по предприятиям';
  end;

  MotorNames:= VVectorToStr(MainForm.UsedNames, ', ');

  ReasonNames:= nil;
  ReasonNames:= VCut(TitleReasonNames);
  if not VIsNil(JointlyReasonIDs) then
  begin
    for i:= 0 to High(JointlyReasonIDs) do
    begin
      N:= VIndexOf(TitleReasonIDs, JointlyReasonIDs[i]);
      if N>=0 then VDel(ReasonNames, N);
    end;
    VAppend(ReasonNames, JointlyReasonName);
  end;


  N:= Length(TitleReasonIds);
  if not VIsNil(JointlyReasonIDs) then
  N:= Length(TitleReasonIds) - Length(JointlyReasonIDs) + 1;

  Grid1.Clear;
  if Assigned(Sheet) then FreeAndNil(Sheet);
  Sheet:= TStatisticReclamationCountSheet.Create(Grid1, N);
  Sheet.Draw(BD, ED, ReportName, MotorNames, FirstColName,
             NameValues, ReasonNames, CountValues, ReasonCountValues, True);
end;

end.

