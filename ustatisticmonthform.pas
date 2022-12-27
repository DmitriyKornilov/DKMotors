unit UStatisticMonthForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  VirtualTrees, fpspreadsheetgrid, rxctrls, DK_VSTTables, DK_Vector, DK_Matrix,
  USQLite, SheetUtils;

type

  { TStatisticMonthForm }

  TStatisticMonthForm = class(TForm)
    ExportButton: TRxSpeedButton;
    Grid1: TsWorksheetGrid;
    ImageList1: TImageList;
    Panel1: TPanel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    VT1: TVirtualStringTree;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

  private
    Years: TStrVector;
    YearList: TVSTCheckTable;

    Sheet: TStatisticReclamationMonthSheet;

    procedure YearListSelectItem(const AIndex: Integer; const AChecked: Boolean);
    procedure SetYearList;
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
end;

procedure TStatisticMonthForm.FormDestroy(Sender: TObject);
begin
  if Assigned(YearList) then FreeAndNil(YearList);
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

procedure TStatisticMonthForm.ShowData;
var
  UsedYears: TStrVector;
  Counts, Accumulations: TIntMatrix;
  ReportName, MotorNames: String;
begin
  UsedYears:= nil;
  if YearList.IsSelected then
    UsedYears:= VCut(Years, YearList.Selected);

  SQLite.ReclamationPerMonthTotalLoad(UsedYears, MainForm.UsedNameIDs,
                                      Counts, Accumulations);


  ReportName:= EmptyStr;
  MotorNames:= VVectorToStr(MainForm.UsedNames, ', ');

  Grid1.Clear;
  if Assigned(Sheet) then FreeAndNil(Sheet);
  Sheet:= TStatisticReclamationMonthSheet.Create(Grid1, Length(UsedYears));
  Sheet.Draw(UsedYears, ReportName, MotorNames, Counts, Accumulations, True);
end;

end.

