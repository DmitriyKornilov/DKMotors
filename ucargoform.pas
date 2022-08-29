unit UCargoForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, UListForm, fpspreadsheetgrid, DateTimePicker, DividerBevel, rxctrls,
  RxDBGrid, DBUtils, DK_Vector, DK_Matrix, SheetUtils, DateUtils,
  UCargoEditForm, VirtualTrees, DK_Dialogs, DK_Const, DK_SheetExporter,
  fpstypes;

type

  { TCargoForm }

  TCargoForm = class(TForm)
    AddButton: TSpeedButton;
    CloseButton: TSpeedButton;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DelButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    DividerBevel5: TDividerBevel;
    DividerBevel6: TDividerBevel;
    EditButton: TSpeedButton;
    ExportButton: TRxSpeedButton;
    Label1: TLabel;
    CargoGrid: TsWorksheetGrid;
    CargoItemGrid: TsWorksheetGrid;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    RxSpeedButton5: TRxSpeedButton;
    Splitter1: TSplitter;
    TopToolsPanel: TPanel;
    VT: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CargoGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure CloseButtonClick(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);

    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);

    procedure RxSpeedButton5Click(Sender: TObject);
  private
    Months: TStrVector;
    Cargos: TStrMatrix;
    MCargoIDs: TIntMatrix;

    SelectedIndex1, SelectedIndex2: Integer;
    SelectedNode: PVirtualNode;


    CargoSheet: TCargoSheet;
    CargoItemSheet: TCargoSheet;

    CargoIDs: TIntVector;
    SendDates: TDateVector;
    ReceiverNames: TStrVector;

    CargoSelectedIndex: Integer;

    procedure CargoOpen;
    procedure CargoSelectionClear;
    procedure CargoSelectLine(const ARow: Integer);

    procedure CargoItemOpen(const ACargoID: Integer);

    procedure CargoEditFormOpen(const AEditMode: Byte); //1 - add, 2 - edit
    procedure DelCargo;

    procedure ExportSheet;

    procedure UnselectNode;
    procedure SelectNode(Node: PVirtualNode);
  public

  end;

var
  CargoForm: TCargoForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TCargoForm }

procedure TCargoForm.DelCargo;
begin
  if CargoSelectedIndex<0 then Exit;
  if not Confirm('Удалить отгрузку ' +  SYMBOL_BREAK +
                  DateToStr(SendDates[CargoSelectedIndex]) + ' в ' +
                  ReceiverNames[CargoSelectedIndex] + '?') then Exit;
  DeleteWithID('CARGOLIST', 'CargoID', CargoIDs[CargoSelectedIndex]);
  CargoItemGrid.Clear;
  CargoOpen;
end;

procedure TCargoForm.ExportSheet;
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(CargoItemGrid);
  try
    //Exporter.SheetName:= 'Отчет';
    Exporter.PageSettings(spoPortrait, pfWidth);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TCargoForm.UnselectNode;
begin
  if SelectedIndex1>=0 then
    SelectedNode^.CheckState:= csUnCheckedNormal;
  SelectedIndex1:= -1;
  SelectedIndex2:= -1;
  SelectedNode:= nil;
  VT.Refresh;
  CargoItemGrid.Clear;
  DelButton.Enabled:= False;
  EditButton.Enabled:= DelButton.Enabled;
end;

procedure TCargoForm.SelectNode(Node: PVirtualNode);
begin
  SelectedIndex1:= (Node^.Parent)^.Index;
  SelectedIndex2:= Node^.Index;
  SelectedNode:= Node;
  SelectedNode^.CheckState:= csCheckedNormal;
  VT.Refresh;
  //OpenListQuery;
  //DelButton.Enabled:= not ListQuery.IsEmpty;
  EditButton.Enabled:= DelButton.Enabled;
end;

procedure TCargoForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton2.Down:= False;
 // MainForm.CargoForm:= nil;
  CloseAction:= caFree;
end;

procedure TCargoForm.FormCreate(Sender: TObject);
begin
  CargoSelectedIndex:= -1;
  CargoSheet:= TCargoSheet.Create(CargoGrid);
  CargoItemSheet:= TCargoSheet.Create(CargoItemGrid);

  DateTimePicker1.Date:= IncDay(Date, -9);
  DateTimePicker2.Date:= Date;
end;

procedure TCargoForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CargoSheet) then FReeAndNil(CargoSheet);
  if Assigned(CargoItemSheet) then FReeAndNil(CargoItemSheet);
end;

procedure TCargoForm.FormShow(Sender: TObject);
begin
  CargoOpen;
end;



procedure TCargoForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TCargoForm.CargoGridMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  R,C: Integer;
begin
  if Button=mbRight then
    CargoSelectionClear;
  if Button=mbLeft  then
  begin
    CargoGrid.MouseToCell(X,Y,C{%H-},R{%H-});
    CargoSelectLine(R);
  end;
end;

procedure TCargoForm.AddButtonClick(Sender: TObject);
begin
  CargoEditFormOpen(1);
end;

procedure TCargoForm.DateTimePicker1Change(Sender: TObject);
begin
  CargoOpen;
end;

procedure TCargoForm.DateTimePicker2Change(Sender: TObject);
begin
  CargoOpen;
end;

procedure TCargoForm.DelButtonClick(Sender: TObject);
begin
  DelCargo;
end;

procedure TCargoForm.EditButtonClick(Sender: TObject);
begin
  CargoEditFormOpen(2);
end;

procedure TCargoForm.ExportButtonClick(Sender: TObject);
begin
  ExportSheet;
end;



procedure TCargoForm.RxSpeedButton5Click(Sender: TObject);
begin
  if ListFormOpen('CARGORECEIVERS', 'ReceiverID', 'ReceiverName', True) then
    CargoOpen;
end;

procedure TCargoForm.CargoItemOpen(const ACargoID: Integer);
var
  SendDate: TDate;
  ReceiverName: String;
  MotorNames: TStrVector;
  MotorCounts: TIntVector;
  MotorNums, Series: TStrMatrix;
begin
   CargoLoad(ACargoID, SendDate, ReceiverName, MotorNames, MotorCounts,
             MotorNums, Series);
   CargoItemSheet.Draw(SendDate, ReceiverName, MotorNames, MotorCounts,
                       MotorNums, Series);
end;

procedure TCargoForm.CargoOpen;
begin
  Screen.Cursor:= crHourGlass;
  try
    CargoSelectionClear;

    CargoListLoad(DateTimePicker1.Date, DateTimePicker2.Date,
                CargoIDs, SendDates, ReceiverNames);
    CargoSheet.Draw(SendDates, ReceiverNames);

    DelButton.Enabled:= CargoSelectedIndex>-1;
    EditButton.Enabled:= CargoSelectedIndex>-1;

  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TCargoForm.CargoSelectionClear;
begin
  if CargoSelectedIndex>-1 then
  begin
    CargoSheet.DrawLine(CargoSelectedIndex, False);
    CargoSelectedIndex:= -1;
  end;
  CargoItemGrid.Clear;
  DelButton.Enabled:= False;
  EditButton.Enabled:= False;
  ExportButton.Enabled:= False;
end;

procedure TCargoForm.CargoSelectLine(const ARow: Integer);
const
  FirstRow = 2;
begin
  if VIsNil(SendDates) then Exit;
  if (ARow>=FirstRow) and (ARow<CargoGrid.RowCount-1) then  //клик ниже заголовка
  begin
    if CargoSelectedIndex>-1 then
      CargoSheet.DrawLine(CargoSelectedIndex, False);
    CargoSelectedIndex:= ARow - FirstRow;
    CargoSheet.DrawLine(CargoSelectedIndex, True);

    DelButton.Enabled:= CargoSelectedIndex>-1;
    EditButton.Enabled:= CargoSelectedIndex>-1;
    ExportButton.Enabled:= CargoSelectedIndex>-1;

    if CargoSelectedIndex>-1 then
      CargoItemOpen(CargoIDs[CargoSelectedIndex]);
  end;
end;



procedure TCargoForm.CargoEditFormOpen(const AEditMode: Byte);
var
  CargoEditForm: TCargoEditForm;
begin
  CargoEditForm:= TCargoEditForm.Create(CargoForm);
  try
    CargoEditForm.CargoID:= 0;
    if AEditMode=2 then
      CargoEditForm.CargoID:= CargoIDs[CargoSelectedIndex];
    CargoEditForm.ShowModal;
  finally
    FreeAndNil(CargoEditForm);
  end;
  CargoOpen;

end;

end.

