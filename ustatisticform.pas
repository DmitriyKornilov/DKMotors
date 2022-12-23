unit UStatisticForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, fpspreadsheetgrid, VirtualTrees, rxctrls, DateTimePicker,
  DividerBevel, USQLite, DK_Vector, DK_DateUtils, DK_VSTTables, SheetUtils,
  LCLType,

  UStatisticCountForm;

type

  { TStatisticForm }

  TStatisticForm = class(TForm)
    Panel1: TPanel;
    MainPanel: TPanel;
    Splitter1: TSplitter;
    VT1: TVirtualStringTree;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

  private
    SelectedIndex: Integer;

    StatisticList: TVSTTable;

    StatisticCountForm: TStatisticCountForm;

    procedure Choose;
    procedure FreeForms;
    procedure SetFormPosition(AForm: TForm);

    procedure StatisticCountFormOpen;

    procedure SetStatisticList;
    procedure StatisticListSelectItem;


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
  SelectedIndex:= -1;
  MainForm.SetNamesPanelsVisible(True, False);
  StatisticList:= TVSTTable.Create(VT1);
  StatisticList.OnSelect:= @StatisticListSelectItem;
  SetStatisticList;
end;

procedure TStatisticForm.FormDestroy(Sender: TObject);
begin
  if Assigned(StatisticList) then FreeAndNil(StatisticList);
  FreeForms;
end;

procedure TStatisticForm.Choose;
begin
  Screen.Cursor:= crHourGlass;
  try
    FreeForms;
    case SelectedIndex of
    0: StatisticCountFormOpen;

    end;

  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TStatisticForm.FreeForms;
begin
  if Assigned(StatisticCountForm) then FreeAndNil(StatisticCountForm);
end;

procedure TStatisticForm.SetFormPosition(AForm: TForm);
begin
  AForm.Parent:= MainPanel;
  AForm.Left:= 0;
  AForm.Top:= 0;
  AForm.MakeFullyVisible();
end;

procedure TStatisticForm.StatisticCountFormOpen;
begin
  StatisticCountForm:= TStatisticCountForm.Create(StatisticForm);
  SetFormPosition(TForm(StatisticCountForm));
  StatisticCountForm.Show;
end;

procedure TStatisticForm.SetStatisticList;
var
  V: TStrVector;
begin
  V:= VCreateStr([
    'Количество за период',
    'Статистика 2'
  ]);

  StatisticList.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  StatisticList.HeaderVisible:= False;
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
  Choose;
end;

procedure TStatisticForm.ShowStatistic;
begin
  Choose;
end;


end.

