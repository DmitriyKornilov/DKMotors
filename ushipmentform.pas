unit UShipmentForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  Spin, StdCtrls, rxctrls, DK_DateUtils, VirtualTrees, DK_VSTUtils, USQLite,
  SheetUtils, UCargoEditForm, DividerBevel, fpspreadsheetgrid, DK_Dialogs,
  DK_Vector, DK_Matrix, DK_Const, DK_SheetExporter, fpstypes;

type

  { TShipmentForm }

  TShipmentForm = class(TForm)
    AddButton: TSpeedButton;
    ChooseRecieverNamesButton: TSpeedButton;
    CloseButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    Label3: TLabel;
    LogGrid: TsWorksheetGrid;
    DelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    DividerBevel3: TDividerBevel;
    EditButton: TSpeedButton;
    EditButtonPanel: TPanel;
    ExportButton: TRxSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    ReceiverNamesLabel: TLabel;
    ReceiverNamesPanel: TPanel;
    RxSpeedButton5: TRxSpeedButton;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    VT: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure ChooseRecieverNamesButtonClick(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RxSpeedButton5Click(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure VTDrawText(Sender: TBaseVirtualTree; TargetCanvas: TCanvas;
      Node: PVirtualNode; Column: TColumnIndex; const CellText: String;
      const CellRect: TRect; var DefaultDraw: Boolean);
    procedure VTGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: String);
    procedure VTInitNode(Sender: TBaseVirtualTree; ParentNode,
      Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
    procedure VTMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    Months: TStrVector;
    Shipments: TStrMatrix;
    CargoIDs: TIntMatrix;

    UsedReceiverIDs: TIntVector;
    UsedReceiverNames: TStrVector;

    SelectedIndex1, SelectedIndex2: Integer;
    SelectedNode: PVirtualNode;

    CargoSheet: TCargoSheet;

    procedure OpenShipmentList(const ACargoID: Integer);
    function OpenShipment: Boolean;

    procedure OpenCargoEditForm(const AEditMode: Byte); //1 - add, 2 - edit

    procedure UnselectNode;
    procedure SelectNode(Node: PVirtualNode);
  public

  end;

var
  ShipmentForm: TShipmentForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TShipmentForm }

procedure TShipmentForm.FormClose(Sender: TObject;
  var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton2.Down:= False;
  MainForm.ShipmentForm:= nil;
  CloseAction:= caFree;
end;

procedure TShipmentForm.AddButtonClick(Sender: TObject);
begin
  OpenCargoEditForm(1);
end;

procedure TShipmentForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TShipmentForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить отгрузку ' +  SYMBOL_BREAK +
                  Shipments[SelectedIndex1, SelectedIndex2] + '?') then Exit;
  SQLite.Delete('CARGOLIST', 'CargoID', CargoIDs[SelectedIndex1, SelectedIndex2]);
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
  SelectedIndex1:= -1;
  SelectedIndex2:= -1;
  CargoSheet:= TCargoSheet.Create(LogGrid);

  SQLite.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, False, UsedReceiverIDs, UsedReceiverNames);

  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TShipmentForm.ChooseRecieverNamesButtonClick(Sender: TObject);
begin
  if SQLite.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, True, UsedReceiverIDs, UsedReceiverNames) then
    OpenShipmentList(0);
end;

procedure TShipmentForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CargoSheet) then FReeAndNil(CargoSheet);
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
    SQLite.ShipmentListLoad(UsedReceiverIDs, SpinEdit1.Value, Months, Shipments, CargoIDs);
    OpenShipment;
  end;
end;

procedure TShipmentForm.SpinEdit1Change(Sender: TObject);
begin
  OpenShipmentList(0);
end;

procedure TShipmentForm.VTDrawText(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  const CellText: String; const CellRect: TRect; var DefaultDraw: Boolean);
var
  i,j: Integer;
begin
  if VT.GetNodeLevel(Node)=1 then
  begin
    i:= (Node^.Parent)^.Index;
    j:= Node^.Index;
    if (SelectedIndex1=i) and (SelectedIndex2=j) then
      TargetCanvas.Font.Style:= [fsBold];
  end;
end;

procedure TShipmentForm.VTGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
  Column: TColumnIndex; TextType: TVSTTextType; var CellText: String);
begin
  VSTGetText(Node, CellText, VT, Months, Shipments);
end;

procedure TShipmentForm.VTInitNode(Sender: TBaseVirtualTree; ParentNode,
  Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
var
  Level: Integer;
begin
  Level := VT.GetNodeLevel(Node);
  if Level=1 then
    Node^.CheckType:= ctRadioButton;
end;

procedure TShipmentForm.VTMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  Node: PVirtualNode;
  Level: Integer;
begin
  if Button<>mbLeft then Exit;
  Node:= VT.GetNodeAt(X, Y);
  if not Assigned(Node) then Exit;
  Level:= VT.GetNodeLevel(Node);
  if Level=1 then
  begin
    UnselectNode;
    SelectNode(Node);
  end;
end;

procedure TShipmentForm.OpenShipmentList(const ACargoID: Integer);
var
  I1, I2: Integer;
begin
  VT.Clear;
  UnselectNode;
  if not SQLite.ShipmentListLoad(UsedReceiverIDs, SpinEdit1.Value, Months, Shipments, CargoIDs) then Exit;
  VSTLoad(VT, Shipments, 0);

  if MIndexOf(CargoIDs, ACargoID, I1, I2) then
    SelectNode(VSTShowNode(VT, I1, I2))
  else //выделение первой отгрузки по списку
    SelectNode(VT.GetNext(VT.GetNext(VT.RootNode)));
end;

function TShipmentForm.OpenShipment: Boolean;
var
  SendDate: TDate;
  ReceiverName: String;
  MotorNames: TStrVector;
  MotorCounts: TIntVector;
  MotorNums, Series: TStrMatrix;
begin
  Result:= SQLite.CargoLoad(CargoIDs[SelectedIndex1, SelectedIndex2],
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
      CargoEditForm.CargoID:= CargoIDs[SelectedIndex1, SelectedIndex2];
    CargoEditForm.ShowModal;
    ID:= CargoEditForm.CargoID;
  finally
    FreeAndNil(CargoEditForm);
  end;
  OpenShipmentList(ID);
end;

procedure TShipmentForm.UnselectNode;
begin
  if SelectedIndex1>=0 then
    SelectedNode^.CheckState:= csUnCheckedNormal;
  SelectedIndex1:= -1;
  SelectedIndex2:= -1;
  SelectedNode:= nil;
  VT.Refresh;
  LogGrid.Clear;
  DelButton.Enabled:= False;
  EditButton.Enabled:= DelButton.Enabled;
  ExportButton.Enabled:= DelButton.Enabled;
end;

procedure TShipmentForm.SelectNode(Node: PVirtualNode);
begin
  SelectedIndex1:= (Node^.Parent)^.Index;
  SelectedIndex2:= Node^.Index;
  SelectedNode:= Node;
  SelectedNode^.CheckState:= csCheckedNormal;
  VT.Refresh;
  DelButton.Enabled:= OpenShipment;
  EditButton.Enabled:= DelButton.Enabled;
  ExportButton.Enabled:= DelButton.Enabled;
end;

end.

