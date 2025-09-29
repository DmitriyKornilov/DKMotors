unit UPackingSheetForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  fpspreadsheetgrid,
  //DK packages utils
  DK_Vector, DK_CtrlUtils,
  //Project utils
  UVars, USheets;

type

  { TPackingSheetForm }

  TPackingSheetForm = class(TForm)
    EditButtonPanel: TPanel;
    ExportButton: TSpeedButton;
    LogGrid: TsWorksheetGrid;
    Panel2: TPanel;
    ToolPanel: TPanel;
    YearPanel: TPanel;
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    CargoPackingSheet: TCargoPackingSheet;
    MotorNames, MotorNums, TestDates: TStrVector;
  public
    CargoID: Integer;
  end;

var
  PackingSheetForm: TPackingSheetForm;

implementation

{$R *.lfm}

{ TPackingSheetForm }

procedure TPackingSheetForm.FormCreate(Sender: TObject);
begin
  CargoPackingSheet:= TCargoPackingSheet.Create(LogGrid, GridFont);
  CargoID:= 0;
end;

procedure TPackingSheetForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(CargoPackingSheet);
end;

procedure TPackingSheetForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);
  Images.ToButtons([ExportButton]);

  DataBase.CargoPackingSheet(CargoID, MotorNames, MotorNums, TestDates);
  CargoPackingSheet.Update(MotorNames, MotorNums, TestDates);
end;

procedure TPackingSheetForm.ExportButtonClick(Sender: TObject);
begin
  if not Assigned(CargoPackingSheet) then Exit;
  CargoPackingSheet.Save('Выполнено!');
end;

end.

