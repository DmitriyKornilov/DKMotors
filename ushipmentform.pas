unit UShipmentForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  Spin, ComCtrls, DK_DateUtils, VirtualTrees, DK_VSTTableTools, USQLite,
  USheetUtils, UCargoEditForm, BCButton, UUtils, fpspreadsheetgrid,
  DK_Dialogs, DK_Vector, DK_Matrix, DK_Const, DK_SheetExporter, DK_Zoom,
  UPackingSheetForm;

type

  { TShipmentForm }

  TShipmentForm = class(TForm)
    AddButton: TSpeedButton;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    DelButton: TSpeedButton;
    EditButton: TSpeedButton;
    EditButtonPanel: TPanel;
    ExportButton: TBCButton;
    PackingSheetButton: TBCButton;
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

    procedure OpenShipmentList(const ACargoID: Integer);

    procedure SelectShipment;

    procedure DrawShipment(const AZoomPercent: Integer);
    procedure ExportShipment;

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

procedure TShipmentForm.AddButtonClick(Sender: TObject);
begin
  OpenCargoEditForm(1);
end;

procedure TShipmentForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить отгрузку ' +  SYMBOL_BREAK +
                  Shipments[VSTCargoList.SelectedIndex1, VSTCargoList.SelectedIndex2] + '?') then Exit;
  SQLite.Delete('CARGOLIST', 'CargoID', CargoIDs[VSTCargoList.SelectedIndex1, VSTCargoList.SelectedIndex2]);
  OpenShipmentList(0);
end;

procedure TShipmentForm.EditButtonClick(Sender: TObject);
begin
  OpenCargoEditForm(2);
end;

procedure TShipmentForm.ExportButtonClick(Sender: TObject);
begin
  ExportShipment;
end;

procedure TShipmentForm.FormCreate(Sender: TObject);
begin
  SetToolPanels([
    ToolPanel
  ]);
  SetToolButtons([
    AddButton, DelButton, EditButton
  ]);

  MainForm.SetNamesPanelsVisible(True, True);

  ZoomPercent:= 100;
  CreateZoomControls(50, 150, ZoomPercent, ZoomPanel, @DrawShipment, True);

  VSTCargoList:= TVSTCategoryIDList.Create(VT, EmptyStr, @SelectShipment);
  CargoSheet:= TCargoSheet.Create(LogGrid.Worksheet, LogGrid);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TShipmentForm.ViewUpdate;
begin
  OpenShipmentList(0);
end;

procedure TShipmentForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CargoSheet) then FReeAndNil(CargoSheet);
  if Assigned(VSTCargoList) then FReeAndNil(VSTCargoList);
end;

procedure TShipmentForm.FormShow(Sender: TObject);
begin
  OpenShipmentList(0);
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
  OpenShipmentList(0);
end;

procedure TShipmentForm.OpenShipmentList(const ACargoID: Integer);
begin
  LogGrid.Clear;
  SQLite.ShipmentListLoad(MainForm.UsedNameIDs, MainForm.UsedReceiverIDs,
                          SpinEdit1.Value, Months, Shipments, CargoIDs);
  VSTCargoList.Update(Months, Shipments, CargoIDs, ACargoID);
  SetButtonsEnabled(VSTCargoList.IsSelected);
end;

procedure TShipmentForm.SelectShipment;
begin
  LogGrid.Clear;
  if not VSTCargoList.IsSelected then Exit;
  SQLite.CargoLoad(CargoIDs[VSTCargoList.SelectedIndex1, VSTCargoList.SelectedIndex2],
             SendDate, ReceiverName, MotorNames, MotorCounts, MotorNums, Series);
  DrawShipment(ZoomPercent);
end;

procedure TShipmentForm.DrawShipment(const AZoomPercent: Integer);
begin
  ZoomPercent:= AZoomPercent;
  CargoSheet.Zoom(ZoomPercent);
  CargoSheet.Draw(SendDate, ReceiverName, MotorNames, MotorCounts, MotorNums, Series);
end;

procedure TShipmentForm.ExportShipment;
var
  Drawer: TCargoSheet;
  Sheet: TsWorksheet;
  Exporter: TSheetsExporter;
begin
  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    Drawer:= TCargoSheet.Create(Sheet);
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
  OpenShipmentList(ID);
end;

procedure TShipmentForm.SetButtonsEnabled(const AEnabled: Boolean);
begin
  DelButton.Enabled:= AEnabled;
  EditButton.Enabled:= DelButton.Enabled;
  ExportButton.Enabled:= DelButton.Enabled;
end;

end.

