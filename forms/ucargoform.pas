unit UCargoForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  Spin, ComCtrls, DividerBevel, VirtualTrees, fpspreadsheetgrid,
  //DK packages utils
  DK_DateUtils, DK_VSTTableTools, DK_Dialogs, DK_Vector, DK_Matrix, DK_Const,
  DK_SheetExporter, DK_Zoom, DK_CtrlUtils,
  //Project utils
  UVars, USheets,
  //Forms
  UCargoEditForm, UPackingSheetForm;

type

  { TShipmentForm }

  TShipmentForm = class(TForm)
    AddButton: TSpeedButton;
    Bevel1: TBevel;
    DelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    EditButton: TSpeedButton;
    EditButtonPanel: TPanel;
    ExportButton: TSpeedButton;
    PackingSheetButton: TSpeedButton;
    LogGrid: TsWorksheetGrid;
    Panel1: TPanel;
    Panel2: TPanel;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    VT: TVirtualStringTree;
    YearPanel: TPanel;
    ZoomPanel: TPanel;
    procedure AddButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure PackingSheetButtonClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    Months: TStrVector;
    Shipments: TStrMatrix;
    CargoIDs: TIntMatrix;

    ZoomPercent: Integer;

    VSTCargoList: TVSTCategoryIDList;
    CargoSheet: TCargoSheet;

    SendDate: TDate;
    ReceiverName: String;
    MotorNames: TStrVector;
    MotorCounts: TIntVector;
    MotorNums, Series: TStrMatrix;

    procedure OpenCargoList(const ACargoID: Integer);
    procedure SelectCargo;
    procedure DrawCargo(const AZoomPercent: Integer);
    procedure ExportCargo;

    procedure OpenCargoEditForm(const AEditMode: Byte); //1 - add, 2 - edit
    procedure SetButtonsEnabled(const AEnabled: Boolean);
  public
    procedure ViewUpdate;
  end;

var
  ShipmentForm: TShipmentForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TShipmentForm }

procedure TShipmentForm.FormCreate(Sender: TObject);
begin
  MainForm.SetNamesPanelsVisible(True, True);

  ZoomPercent:= 100;
  CreateZoomControls(50, 150, ZoomPercent, ZoomPanel, @DrawCargo, True);

  VSTCargoList:= TVSTCategoryIDList.Create(VT, EmptyStr, @SelectCargo);
  CargoSheet:= TCargoSheet.Create(LogGrid.Worksheet, LogGrid, GridFont);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TShipmentForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CargoSheet) then FReeAndNil(CargoSheet);
  if Assigned(VSTCargoList) then FReeAndNil(VSTCargoList);
end;

procedure TShipmentForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);
  SetToolButtons([AddButton, DelButton, EditButton]);
  Images.ToButtons([ExportButton, PackingSheetButton, AddButton, DelButton, EditButton]);

  OpenCargoList(0);
end;

procedure TShipmentForm.AddButtonClick(Sender: TObject);
begin
  OpenCargoEditForm(1);
end;

procedure TShipmentForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить отгрузку ' +  SYMBOL_BREAK +
                  Shipments[VSTCargoList.SelectedIndex1, VSTCargoList.SelectedIndex2] + '?') then Exit;
  DataBase.Delete('CARGOLIST', 'CargoID', CargoIDs[VSTCargoList.SelectedIndex1, VSTCargoList.SelectedIndex2]);
  OpenCargoList(0);
end;

procedure TShipmentForm.EditButtonClick(Sender: TObject);
begin
  OpenCargoEditForm(2);
end;

procedure TShipmentForm.ExportButtonClick(Sender: TObject);
begin
  ExportCargo;
end;

procedure TShipmentForm.ViewUpdate;
begin
  OpenCargoList(0);
end;

procedure TShipmentForm.PackingSheetButtonClick(Sender: TObject);
var
  PackingSheetForm: TPackingSheetForm;
begin
  if not VSTCargoList.IsSelected then Exit;

  PackingSheetForm:= TPackingSheetForm.Create(nil);
  try
    PackingSheetForm.CargoID:= CargoIDs[VSTCargoList.SelectedIndex1, VSTCargoList.SelectedIndex2];
    PackingSheetForm.ShowModal;
  finally
    FreeAndNil(PackingSheetForm);
  end;
end;

procedure TShipmentForm.SpinEdit1Change(Sender: TObject);
begin
  OpenCargoList(0);
end;

procedure TShipmentForm.OpenCargoList(const ACargoID: Integer);
begin
  LogGrid.Clear;
  DataBase.ShipmentListLoad(MainForm.UsedNameIDs, MainForm.UsedReceiverIDs,
                          SpinEdit1.Value, Months, Shipments, CargoIDs);
  VSTCargoList.Update(Months, Shipments, CargoIDs, ACargoID);
  SetButtonsEnabled(VSTCargoList.IsSelected);
end;

procedure TShipmentForm.SelectCargo;
begin
  LogGrid.Clear;
  if not VSTCargoList.IsSelected then Exit;
  DataBase.CargoLoad(CargoIDs[VSTCargoList.SelectedIndex1, VSTCargoList.SelectedIndex2],
             SendDate, ReceiverName, MotorNames, MotorCounts, MotorNums, Series);
  DrawCargo(ZoomPercent);
end;

procedure TShipmentForm.DrawCargo(const AZoomPercent: Integer);
begin
  ZoomPercent:= AZoomPercent;
  CargoSheet.Zoom(ZoomPercent);
  CargoSheet.Draw(SendDate, ReceiverName, MotorNames, MotorCounts, MotorNums, Series);
end;

procedure TShipmentForm.ExportCargo;
var
  Drawer: TCargoSheet;
  Sheet: TsWorksheet;
  Exporter: TSheetsExporter;
begin
  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    Drawer:= TCargoSheet.Create(Sheet, nil, GridFont);
    try
      Drawer.Draw(SendDate, ReceiverName, MotorNames, MotorCounts, MotorNums, Series);
    finally
      FreeAndNil(Drawer);
    end;
    Exporter.PageSettings(spoPortrait);

    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TShipmentForm.OpenCargoEditForm(const AEditMode: Byte);
var
  CargoEditForm: TCargoEditForm;
  ID: Integer;
begin
  CargoEditForm:= TCargoEditForm.Create(ShipmentForm);
  try
    CargoEditForm.CargoID:= 0;
    if AEditMode=2 then
      CargoEditForm.CargoID:= CargoIDs[VSTCargoList.SelectedIndex1, VSTCargoList.SelectedIndex2];
    CargoEditForm.ShowModal;
    ID:= CargoEditForm.CargoID;
  finally
    FreeAndNil(CargoEditForm);
  end;
  OpenCargoList(ID);
end;

procedure TShipmentForm.SetButtonsEnabled(const AEnabled: Boolean);
begin
  DelButton.Enabled:= AEnabled;
  EditButton.Enabled:= DelButton.Enabled;
  ExportButton.Enabled:= DelButton.Enabled;
end;

end.

