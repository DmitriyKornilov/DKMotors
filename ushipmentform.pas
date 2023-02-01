unit UShipmentForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  Spin, rxctrls, DK_DateUtils, VirtualTrees, DK_VSTTables, USQLite,
  USheetUtils, UCargoEditForm, DividerBevel, fpspreadsheetgrid, DK_Dialogs,
  DK_Vector, DK_Matrix, DK_Const, DK_SheetExporter, fpstypes;

type

  { TShipmentForm }

  TShipmentForm = class(TForm)
    AddButton: TSpeedButton;
    LogGrid: TsWorksheetGrid;
    DelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    DividerBevel3: TDividerBevel;
    EditButton: TSpeedButton;
    EditButtonPanel: TPanel;
    ExportButton: TRxSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    RxSpeedButton5: TRxSpeedButton;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    VT: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RxSpeedButton5Click(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure VTClick(Sender: TObject);
  private
    Months: TStrVector;
    Shipments: TStrMatrix;
    CargoIDs: TIntMatrix;

    VST: TVSTCategoryRadioButtonTable;
    CargoSheet: TCargoSheet;

    procedure OpenShipmentList(const ACargoID: Integer);
    function OpenShipment: Boolean;

    procedure OpenCargoEditForm(const AEditMode: Byte); //1 - add, 2 - edit
    procedure SetButtonsEnabled(const AEnabled: Boolean);
  public
    procedure ShowShipment;
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
                  Shipments[VST.SelectedIndex1, VST.SelectedIndex2] + '?') then Exit;
  SQLite.Delete('CARGOLIST', 'CargoID', CargoIDs[VST.SelectedIndex1, VST.SelectedIndex2]);
  OpenShipmentList(0);
end;

procedure TShipmentForm.EditButtonClick(Sender: TObject);
begin
  OpenCargoEditForm(2);
end;

procedure TShipmentForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(LogGrid);
  try
    //Exporter.SheetName:= 'Отчет';
    Exporter.PageSettings(spoPortrait, pfWidth);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TShipmentForm.FormCreate(Sender: TObject);
begin
  MainForm.SetNamesPanelsVisible(False, True);

  VST:= TVSTCategoryRadioButtonTable.Create(VT);
  VST.SelectedFont.Style:= [fsBold];
  VST.CanUnselect:= False;
  VST.HeaderVisible:= False;
  VST.GridLinesVisible:= False;
  VST.AddColumn('Shipments');

  CargoSheet:= TCargoSheet.Create(LogGrid);

  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TShipmentForm.ShowShipment;
begin
  OpenShipmentList(0);
end;

procedure TShipmentForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CargoSheet) then FReeAndNil(CargoSheet);
  if Assigned(VST) then FReeAndNil(VST);
end;

procedure TShipmentForm.FormShow(Sender: TObject);
begin
  OpenShipmentList(0);
end;

procedure TShipmentForm.RxSpeedButton5Click(Sender: TObject);
begin
  if SQLite.EditList('Грузополучатели',
    'CARGORECEIVERS', 'ReceiverID', 'ReceiverName', True, True) then
  begin
    SQLite.ShipmentListLoad(MainForm.UsedReceiverIDs, SpinEdit1.Value, Months, Shipments, CargoIDs);
    OpenShipment;
  end;
end;

procedure TShipmentForm.SpinEdit1Change(Sender: TObject);
begin
  OpenShipmentList(0);
end;

procedure TShipmentForm.VTClick(Sender: TObject);
begin
  if VST.IsSelected then
    OpenShipment;
end;

procedure TShipmentForm.OpenShipmentList(const ACargoID: Integer);
var
  I1, I2: Integer;
begin
  LogGrid.Clear;
  VST.ValuesClear;
  SetButtonsEnabled(False);
  if not SQLite.ShipmentListLoad(MainForm.UsedReceiverIDs, SpinEdit1.Value,
                                 Months, Shipments, CargoIDs) then Exit;
  VST.SetCategories(Months);
  VST.SetColumn('Shipments', Shipments, taLeftJustify);
  VST.Draw;

  if MIndexOf(CargoIDs, ACargoID, I1, I2) then
  begin
    VST.Select(I1, I2);
    VST.Show(I1, I2);
  end
  else begin
    VST.Select(0, 0);
    VST.Show(0, 0);
  end;
  SetButtonsEnabled(True);
  OpenShipment;
end;

function TShipmentForm.OpenShipment: Boolean;
var
  SendDate: TDate;
  ReceiverName: String;
  MotorNames: TStrVector;
  MotorCounts: TIntVector;
  MotorNums, Series: TStrMatrix;
begin
  Result:= SQLite.CargoLoad(CargoIDs[VST.SelectedIndex1, VST.SelectedIndex2],
             SendDate, ReceiverName, MotorNames, MotorCounts, MotorNums, Series);
  CargoSheet.Draw(
             SendDate, ReceiverName, MotorNames, MotorCounts, MotorNums, Series);

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
      CargoEditForm.CargoID:= CargoIDs[VST.SelectedIndex1, VST.SelectedIndex2];
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

