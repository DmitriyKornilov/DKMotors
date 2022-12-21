unit UStatisticForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, VirtualTrees, rxctrls, DateTimePicker, DividerBevel, USQLite,
  DK_Vector, DK_DateUtils, DK_VSTTables, SheetUtils;

type

  { TStatisticForm }

  TStatisticForm = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Splitter1: TSplitter;
    VT1: TVirtualStringTree;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    StatisticList: TVSTTable;
    procedure SetStatisticList;
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
  MainForm.SetNamesPanelsVisible(True, False);
  StatisticList:= TVSTTable.Create(VT1);
  SetStatisticList;
  //DateTimePicker1.Date:= Date;
  //DateTimePicker2.Date:= FirstDayInMonth(Date);
end;

procedure TStatisticForm.FormDestroy(Sender: TObject);
begin
  if Assigned(StatisticList) then FreeAndNil(StatisticList);
end;

procedure TStatisticForm.SetStatisticList;
var
  V: TStrVector;
begin
  V:= VCreateStr([
    'Статистика 1',
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

procedure TStatisticForm.ShowStatistic;
begin

end;


end.

