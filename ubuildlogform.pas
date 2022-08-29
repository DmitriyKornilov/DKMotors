unit UBuildLogForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, StdCtrls, VirtualTrees, DividerBevel, USQLite,
  rxctrls, DK_DateUtils, DK_Vector, DK_Matrix, DK_VSTUtils,
  DK_Dialogs, DK_Const, UBuildAddForm, UBuildEditForm,
  SheetUtils, fpspreadsheetgrid;

type

  { TBuildLogForm }

  TBuildLogForm = class(TForm)
    AddButton: TSpeedButton;
    CheckBox1: TCheckBox;
    CloseButton: TSpeedButton;
    DelButton: TSpeedButton;
    DividerBevel5: TDividerBevel;
    EditButton: TSpeedButton;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    EditButtonPanel: TPanel;
    DividerBevel1: TDividerBevel;
    DividerBevel4: TDividerBevel;
    LogGrid: TsWorksheetGrid;
    MotorNameComboBox: TComboBox;
    MotorNamesButton: TRxSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel6: TPanel;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    VT: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CheckBox1Change(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LogGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure MotorNamesButtonClick(Sender: TObject);
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
    Dates: TDateMatrix;
    StrDates: TStrMatrix;

    SelectedIndex1, SelectedIndex2: Integer;
    SelectedNode: PVirtualNode;

    MotorBuildSheet: TMotorBuildSheet;

    MotorIDs, NameIDs, OldMotors, ViewNameIDs: TIntVector;
    MotorNames, MotorNums, RotorNums: TStrVector;

    SelectedIndex: Integer;
    procedure SelectionClear;
    procedure SelectLine(const ARow: Integer);

    procedure OpenDatesList(const ASelectDate: TDate);
    procedure OpenMotorsList;

    procedure UnselectNode;
    procedure SelectNode(Node: PVirtualNode);
  public

  end;

var
  BuildLogForm: TBuildLogForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TBuildLogForm }

procedure TBuildLogForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton1.Down:= False;
  MainForm.BuildLogForm:= nil;
  CloseAction:= caFree;
end;

procedure TBuildLogForm.FormCreate(Sender: TObject);
begin
  SelectedIndex:= -1;
  SelectedIndex1:= -1;
  SelectedIndex2:= -1;
  MotorBuildSheet:= TMotorBuildSheet.Create(LogGrid);
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, ViewNameIDs, False);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TBuildLogForm.FormDestroy(Sender: TObject);
begin
  if Assigned(MotorBuildSheet) then FreeAndNil(MotorBuildSheet);
end;

procedure TBuildLogForm.FormShow(Sender: TObject);
begin
  OpenDatesList(Date);
end;

procedure TBuildLogForm.LogGridMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  R,C: Integer;
begin
  if Button=mbRight then
    SelectionClear;
  if Button=mbLeft  then
  begin
    (Sender As TsWorksheetGrid).MouseToCell(X,Y,C{%H-},R{%H-});
    SelectLine(R);
  end;
end;

procedure TBuildLogForm.MotorNameComboBoxChange(Sender: TObject);
begin
  OpenMotorsList;
end;

procedure TBuildLogForm.MotorNamesButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Наименования электродвигателей',
    'MOTORNAMES', 'NameID', 'MotorName', False, True) then
      OpenMotorsList;
end;

procedure TBuildLogForm.SpinEdit1Change(Sender: TObject);
begin
  OpenDatesList(LastDayInYear(SpinEdit1.Value));
end;

procedure TBuildLogForm.VTDrawText(Sender: TBaseVirtualTree;
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

procedure TBuildLogForm.VTGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
  Column: TColumnIndex; TextType: TVSTTextType; var CellText: String);
begin
  VSTGetText(Node, CellText, VT, Months, StrDates);
end;

procedure TBuildLogForm.VTInitNode(Sender: TBaseVirtualTree; ParentNode,
  Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
var
  Level: Integer;
begin
  Level := VT.GetNodeLevel(Node);
  if Level=1 then
    Node^.CheckType:= ctRadioButton;
end;

procedure TBuildLogForm.VTMouseDown(Sender: TObject; Button: TMouseButton;
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

procedure TBuildLogForm.SelectionClear;
begin
  if SelectedIndex>-1 then
  begin
    MotorBuildSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= -1;
  end;
  DelButton.Enabled:= False;
  EditButton.Enabled:= DelButton.Enabled;
end;

procedure TBuildLogForm.SelectLine(const ARow: Integer);
const
  FirstRow = 3;
begin
  if VIsNil(MotorNames) then Exit;
  if (ARow>=FirstRow) and (ARow<LogGrid.RowCount-1) then  //клик ниже заголовка
  begin
    if SelectedIndex>-1 then
      MotorBuildSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= ARow - FirstRow;
    MotorBuildSheet.DrawLine(SelectedIndex, True);
    DelButton.Enabled:= True;
    EditButton.Enabled:= DelButton.Enabled;
  end;
end;

procedure TBuildLogForm.UnselectNode;
begin
  if SelectedIndex1>=0 then
    SelectedNode^.CheckState:= csUnCheckedNormal;
  SelectedIndex1:= -1;
  SelectedIndex2:= -1;
  SelectedNode:= nil;
  VT.Refresh;
  SelectionClear;
end;

procedure TBuildLogForm.SelectNode(Node: PVirtualNode);
begin
  SelectedIndex1:= (Node^.Parent)^.Index;
  SelectedIndex2:= Node^.Index;
  SelectedNode:= Node;
  SelectedNode^.CheckState:= csCheckedNormal;
  VT.Refresh;
  OpenMotorsList;
end;

procedure TBuildLogForm.OpenDatesList(const ASelectDate: TDate);
var
  I1, I2: Integer;
begin
  VT.Clear;
  UnselectNode;

  if not SQLite.MonthAndDatesForBuildLogLoad(SpinEdit1.Value, Months, Dates) then Exit;

  StrDates:= MFormatDateTime('dd.mm.yyyy', Dates);
  VSTLoad(VT, StrDates, 0);
  if MIndexOf(Dates, ASelectDate, I1, I2) then
    SelectNode(VSTShowNode(VT, I1, I2))
  else
    SelectNode(VT.GetNext(VT.GetNext(VT.RootNode)));
end;

procedure TBuildLogForm.OpenMotorsList;
var
  Tmp: TDateVector;
begin
  SelectionClear;
  if (SelectedIndex1<0) or (SelectedIndex2<0) then Exit;

  SQLite.BuildListLoad(Dates[SelectedIndex1, SelectedIndex2],
                    Dates[SelectedIndex1, SelectedIndex2],
                    ViewNameIDs[MotorNameComboBox.ItemIndex],//0 {все наименования},
                    CheckBox1.Checked, MotorIDs, NameIDs, OldMotors,
                    Tmp, MotorNames, MotorNums, RotorNums);

  MotorBuildSheet.Draw(Dates[SelectedIndex1, SelectedIndex2],
                    MotorNames, MotorNums, RotorNums);
end;

procedure TBuildLogForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TBuildLogForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить двигатель ' +  SYMBOL_BREAK +
                  MotorNames[SelectedIndex] +
                 ' № ' +
                  MotorNums[SelectedIndex] +
                 '?') then Exit;
  SQLite.Delete('MOTORLIST', 'MotorID', MotorIDs[SelectedIndex]);
  OpenDatesList(Dates[SelectedIndex1, SelectedIndex2]);
end;

procedure TBuildLogForm.EditButtonClick(Sender: TObject);
var
  BuildEditForm: TBuildEditForm;
  D: TDate;
begin
  BuildEditForm:= TBuildEditForm.Create(BuildLogForm);
  try
    BuildEditForm.DateTimePicker1.Date:= Dates[SelectedIndex1, SelectedIndex2];
    BuildEditForm.MotorID:= MotorIDs[SelectedIndex];
    BuildEditForm.OldNameID:= NameIDs[SelectedIndex];
    BuildEditForm.MotorNumEdit.Text:= MotorNums[SelectedIndex];
    BuildEditForm.RotorNumEdit.Text:= RotorNums[SelectedIndex];
    BuildEditForm.OldMotorCheckBox.Checked:= OldMotors[SelectedIndex]=1;
    BuildEditForm.ShowModal;
    D:= BuildEditForm.DateTimePicker1.Date;
  finally
    FreeAndNil(BuildEditForm);
  end;
   OpenDatesList(D);
end;

procedure TBuildLogForm.CheckBox1Change(Sender: TObject);
begin
  OpenMotorsList;
end;

procedure TBuildLogForm.AddButtonClick(Sender: TObject);
var
  BuildAddForm: TBuildAddForm;
  D: TDate;
begin
  BuildAddForm:= TBuildAddForm.Create(BuildLogForm);
  try
    if (SelectedIndex1>=0) and (SelectedIndex2>=0) then
      BuildAddForm.DateTimePicker1.Date:= Dates[SelectedIndex1, SelectedIndex2]
    else
      BuildAddForm.DateTimePicker1.Date:= Date;
    BuildAddForm.ShowModal;
    D:= BuildAddForm.DateTimePicker1.Date;
  finally
    FreeAndNil(BuildAddForm);
  end;
  OpenDatesList(D);
end;

end.

