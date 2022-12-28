unit UStatisticMonthForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  VirtualTrees, fpspreadsheetgrid, rxctrls, DK_VSTTables, DK_Vector, DK_Matrix,
  USQLite, SheetUtils, DK_SheetExporter, fpstypes;

type

  { TStatisticMonthForm }

  TStatisticMonthForm = class(TForm)
    ExportButton: TRxSpeedButton;
    Grid1: TsWorksheetGrid;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    ToolPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

  private
    Years: TStrVector;
    YearList: TVSTCheckTable;

    CategoryNames: TStrVector;
    CategoryList: TVSTTable;

    Sheet: TStatisticReclamationMonthSheet;

    procedure YearListSelectItem(const AIndex: Integer; const AChecked: Boolean);
    procedure SetYearList;

    procedure CategoryListSelectItem;
    procedure SetCategoryList;
  public
    procedure ShowData;
  end;

var
  StatisticMonthForm: TStatisticMonthForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TStatisticMonthForm }

procedure TStatisticMonthForm.FormCreate(Sender: TObject);
begin
  YearList:= TVSTCheckTable.Create(VT1);
  YearList.OnSelect:= @YearListSelectItem;
  YearList.MaxCheckedCount:= 3; //!!!!
  SetYearList;

  CategoryList:= TVSTTable.Create(VT2);
  CategoryList.OnSelect:= @CategoryListSelectItem;
  SetCategoryList;
end;

procedure TStatisticMonthForm.ExportButtonClick(Sender: TObject);
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

procedure TStatisticMonthForm.FormDestroy(Sender: TObject);
begin
  if Assigned(YearList) then FreeAndNil(YearList);
  if Assigned(CategoryList) then FreeAndNil(CategoryList);
  if Assigned(Sheet) then FreeAndNil(Sheet);
end;

procedure TStatisticMonthForm.YearListSelectItem(const AIndex: Integer; const AChecked: Boolean);
begin
  ShowData;
end;

procedure TStatisticMonthForm.SetYearList;
begin
  SQLite.ReclamationExistYearsLoad(Years);
  YearList.HeaderVisible:= False;
  YearList.SelectedBGColor:= VT1.Color;
  YearList.AddColumn('Список');
  YearList.SetColumn('Список', Years, taLeftJustify);
  YearList.Draw;
end;

procedure TStatisticMonthForm.CategoryListSelectItem;
begin
  ShowData;
end;

procedure TStatisticMonthForm.SetCategoryList;
begin
  CategoryNames:= VCreateStr([
    'Общее количество',
    'Некачественные комплектующие',
    'Дефект сборки / изготовления',
    'Неправильная эксплуатация',
    'Неисправность локомотива'
  ]);

  CategoryList.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  CategoryList.HeaderVisible:= False;
  CategoryList.CanSelect:= True;
  CategoryList.CanUnselect:= False;
  CategoryList.AddColumn('Список');
  CategoryList.SetColumn('Список', CategoryNames, taLeftJustify);
  CategoryList.Draw;
  CategoryList.Select(0);

end;

procedure TStatisticMonthForm.ShowData;
var
  UsedYears: TStrVector;
  Counts, Accumulations: TIntMatrix;
  ReportName, MotorNames: String;
begin
  UsedYears:= nil;
  if YearList.IsSelected then
    UsedYears:= VCut(Years, YearList.Selected);

  Counts:= nil;
  Accumulations:= nil;

  if Assigned(CategoryList) and CategoryList.IsSelected then
    SQLite.ReclamationPerMonthLoad(CategoryList.SelectedIndex,
       UsedYears, MainForm.UsedNameIDs, Counts, Accumulations);

  ReportName:= EmptyStr;
  if (not VIsNil(CategoryNames)) and (CategoryList.IsSelected) then
    ReportName:= CategoryNames[CategoryList.SelectedIndex];

  MotorNames:= VVectorToStr(MainForm.UsedNames, ', ');

  Grid1.Clear;
  if Assigned(Sheet) then FreeAndNil(Sheet);
  Sheet:= TStatisticReclamationMonthSheet.Create(Grid1, Length(UsedYears));
  Sheet.Draw(UsedYears, ReportName, MotorNames, Counts, Accumulations, True);
end;

end.

