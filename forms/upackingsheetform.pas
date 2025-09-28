unit UPackingSheetForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, BCButton,
  fpspreadsheetgrid,
  //DK packages utils
  DK_Vector,
  //Project utils
  UDataBase, USheetUtils;

type

  { TPackingSheetForm }

  TPackingSheetForm = class(TForm)
    EditButtonPanel: TPanel;
    ExportButton: TBCButton;
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
    procedure CreatePackingSheet;
  public
    CargoID: Integer;
  end;

var
  PackingSheetForm: TPackingSheetForm;

implementation

{$R *.lfm}

{ TPackingSheetForm }

procedure TPackingSheetForm.FormShow(Sender: TObject);
begin
  CreatePackingSheet;
end;

procedure TPackingSheetForm.CreatePackingSheet;
var
  MotorNames, MotorNums, TestDates: TStrVector;
begin
  if CargoID=0 then Exit;
  DataBase.CargoPackingSheet(CargoID, MotorNames, MotorNums, TestDates);
  CargoPackingSheet.Update(MotorNames, MotorNums, TestDates);
end;

procedure TPackingSheetForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(CargoPackingSheet);
end;

procedure TPackingSheetForm.ExportButtonClick(Sender: TObject);
begin
  if not Assigned(CargoPackingSheet) then Exit;
  CargoPackingSheet.Save('Выполнено!');
end;

procedure TPackingSheetForm.FormCreate(Sender: TObject);
begin
  CargoPackingSheet:= TCargoPackingSheet.Create(LogGrid);
  CargoID:= 0;
end;

end.

