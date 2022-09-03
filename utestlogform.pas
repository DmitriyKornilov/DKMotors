unit UTestLogForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, StdCtrls, VirtualTrees, DividerBevel, USQLite,
  DK_DateUtils, DK_Vector, DK_Matrix, DK_VSTUtils,
  DK_Dialogs, DK_Const, UTestAddForm, fpspreadsheetgrid,
  SheetUtils;

type

  { TTestLogForm }

  TTestLogForm = class(TForm)
    AddButton: TSpeedButton;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    ChooseMotorNamesButton: TSpeedButton;
    CloseButton: TSpeedButton;
    DelButton: TSpeedButton;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    DividerBevel4: TDividerBevel;
    EditButtonPanel: TPanel;
    DividerBevel1: TDividerBevel;
    Label1: TLabel;
    LogGrid: TsWorksheetGrid;
    MotorNamesLabel: TLabel;
    MotorNamesPanel: TPanel;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    TestGrid: TsWorksheetGrid;
    ToolPanel: TPanel;
    VT: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CheckBox1Change(Sender: TObject);
    procedure CheckBox2Change(Sender: TObject);
    procedure ChooseMotorNamesButtonClick(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);

    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);


    procedure SpinEdit1Change(Sender: TObject);
    procedure TestGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
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

    TestIDs, TestResults: TIntVector;
    MotorNames, MotorNums, TestNotes: TStrVector;

    SelectedIndex: Integer;

    UsedNameIDs: TIntVector;
    UsedNames: TStrVector;

    BeforeTestSheet: TBeforeTestSheet;
    MotorTestSheet: TMotorTestSheet;

    procedure SelectionClear;
    procedure SelectLine(const ARow: Integer);

    procedure OpenBeforeTestList;
    procedure OpenDatesList(const ASelectDate: TDate);
    procedure OpenTestsList;

    procedure UnselectNode;
    procedure SelectNode(Node: PVirtualNode);

  public

  end;

var
  TestLogForm: TTestLogForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TTestLogForm }

procedure TTestLogForm.OpenBeforeTestList;
var
  BuildDates: TDateVector;
  MNames, MNums, TotalMNames: TStrVector;
  TotalMCounts: TIntVector;
  BTestDates: TDateMatrix;
  BTestFails: TIntMatrix;
  BTestNotes: TStrMatrix;
begin
  Screen.Cursor:= crHourGlass;
  try
    SQLite.TestBeforeListLoad(UsedNameIDs,
                         Checkbox1.Checked, BuildDates, MNames, MNums,
                         TotalMNames, TotalMCounts,
                         BTestDates, BTestFails, BTestNotes);
    BeforeTestSheet.Draw(BuildDates, MNames, MNums,
                         TotalMNames, TotalMCounts,
                         BTestDates, BTestFails, BTestNotes);
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TTestLogForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton5.Down:= False;
  MainForm.TestLogForm:= nil;
  CloseAction:= caFree;
end;

procedure TTestLogForm.FormCreate(Sender: TObject);
begin
  SelectedIndex:= -1;
  SelectedIndex1:= -1;
  SelectedIndex2:= -1;

  SQLite.KeyPickList('MOTORNAMES', 'NameID', 'MotorName',
                     UsedNameIDs, UsedNames, True, 'NameID');
  MotorNamesLabel.Caption:= VVectorToStr(UsedNames, ', ');
  MotorNamesLabel.Hint:= MotorNamesLabel.Caption;

  SpinEdit1.Value:= YearOfDate(Date);
  BeforeTestSheet:= TBeforeTestSheet.Create(LogGrid);
  MotorTestSheet:= TMotorTestSheet.Create(TestGrid);
end;

procedure TTestLogForm.FormDestroy(Sender: TObject);
begin
  if Assigned(BeforeTestSheet) then FreeAndNil(BeforeTestSheet);
  if Assigned(MotorTestSheet) then FreeAndNil(MotorTestSheet);
end;

procedure TTestLogForm.FormShow(Sender: TObject);
begin
  OpenDatesList(Date);
  OpenBeforeTestList;
end;

procedure TTestLogForm.SpinEdit1Change(Sender: TObject);
begin
  OpenDatesList(LastDayInYear(SpinEdit1.Value));
end;

procedure TTestLogForm.TestGridMouseDown(Sender: TObject; Button: TMouseButton;
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

procedure TTestLogForm.VTDrawText(Sender: TBaseVirtualTree;
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

procedure TTestLogForm.VTGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
  Column: TColumnIndex; TextType: TVSTTextType; var CellText: String);
begin
  VSTGetText(Node, CellText, VT, Months, StrDates);
end;

procedure TTestLogForm.VTInitNode(Sender: TBaseVirtualTree; ParentNode,
  Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
var
  Level: Integer;
begin
  Level := VT.GetNodeLevel(Node);
  if Level=1 then
    Node^.CheckType:= ctRadioButton;
end;

procedure TTestLogForm.VTMouseDown(Sender: TObject; Button: TMouseButton;
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

procedure TTestLogForm.SelectionClear;
begin
  if SelectedIndex>-1 then
  begin
    MotorTestSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= -1;
  end;
  DelButton.Enabled:= False;
end;

procedure TTestLogForm.SelectLine(const ARow: Integer);
const
  FirstRow = 3;
begin
  if VIsNil(MotorNames) then Exit;
  if (ARow>=FirstRow) and (ARow<TestGrid.RowCount-1-1 {минус строка итогов}) then  //клик ниже заголовка
  begin
    if SelectedIndex>-1 then
      MotorTestSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= ARow - FirstRow;
    MotorTestSheet.DrawLine(SelectedIndex, True);
    DelButton.Enabled:= True;
  end;
end;

procedure TTestLogForm.UnselectNode;
begin
  if SelectedIndex1>=0 then
    SelectedNode^.CheckState:= csUnCheckedNormal;
  SelectedIndex1:= -1;
  SelectedIndex2:= -1;
  SelectedNode:= nil;
  VT.Refresh;
  SelectionClear;
end;

procedure TTestLogForm.SelectNode(Node: PVirtualNode);
begin
  SelectedIndex1:= (Node^.Parent)^.Index;
  SelectedIndex2:= Node^.Index;
  SelectedNode:= Node;
  SelectedNode^.CheckState:= csCheckedNormal;
  VT.Refresh;
  OpenTestsList;
end;

procedure TTestLogForm.OpenDatesList(const ASelectDate: TDate);
var
  I1, I2: Integer;
begin
  TestGrid.Clear;
  VT.Clear;
  UnselectNode;
  if not SQLite.MonthAndDatesForTestLogLoad(SpinEdit1.Value, Months, Dates) then Exit;
  StrDates:= MFormatDateTime('dd.mm.yyyy', Dates);
  VSTLoad(VT, StrDates, 0);

  if MIndexOf(Dates, ASelectDate, I1, I2) then
    SelectNode(VSTShowNode(VT, I1, I2))
  else
    SelectNode(VT.GetNext(VT.GetNext(VT.RootNode)));
end;

procedure TTestLogForm.OpenTestsList;
var
  TotalMotorNames: TStrVector;
  TotalMotorCounts, TotalFailCounts: TIntVector;
  X: TDateVector;
begin
  SelectionClear;
  if (SelectedIndex1<0) or (SelectedIndex2<0) then Exit;

  SQLite.TestListLoad(Dates[SelectedIndex1, SelectedIndex2],
                    Dates[SelectedIndex1, SelectedIndex2],
                    UsedNameIDs,
                    CheckBox1.Checked, TestIDs, TestResults, X,
                    MotorNames, MotorNums, TestNotes);
  SQLite.TestTotalLoad(Dates[SelectedIndex1, SelectedIndex2],
                     Dates[SelectedIndex1, SelectedIndex2],
                     UsedNameIDs,
                     TotalMotorNames, TotalMotorCounts, TotalFailCounts);
  MotorTestSheet.Draw(Dates[SelectedIndex1, SelectedIndex2],
                      MotorNames, MotorNums, TestNotes, TestResults,
                      TotalMotorCounts, TotalFailCounts);
end;

procedure TTestLogForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TTestLogForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить результаты испытаний?') then Exit;
  SQLite.Delete('MOTORTEST', 'TestID', TestIDs[SelectedIndex]);
  OpenDatesList(Dates[SelectedIndex1, SelectedIndex2]);
  OpenBeforeTestList;
end;

procedure TTestLogForm.CheckBox1Change(Sender: TObject);
begin
  OpenTestsList;
  OpenBeforeTestList;
end;

procedure TTestLogForm.CheckBox2Change(Sender: TObject);
begin
  Panel4.Visible:= CheckBox2.Checked;
  Splitter2.Visible:= CheckBox2.Checked;
end;

procedure TTestLogForm.ChooseMotorNamesButtonClick(Sender: TObject);
begin
  if SQLite.EditKeyPickList(UsedNameIDs, UsedNames, 'Наименования электродвигателей',
    'MOTORNAMES', 'NameID', 'MotorName', False, True) then
  begin
    OpenTestsList;
    OpenBeforeTestList;
  end;
  MotorNamesLabel.Caption:= VVectorToStr(UsedNames, ', ');
  MotorNamesLabel.Hint:= MotorNamesLabel.Caption;
end;

procedure TTestLogForm.AddButtonClick(Sender: TObject);
var
  TestAddForm: TTestAddForm;
  D: TDate;
begin
  TestAddForm:= TTestAddForm.Create(TestLogForm);
  try
    if (SelectedIndex1>=0) and (SelectedIndex2>=0) then
      TestAddForm.DateTimePicker1.Date:= Dates[SelectedIndex1, SelectedIndex2]
    else
      TestAddForm.DateTimePicker1.Date:= Date;
    TestAddForm.ShowModal;
    D:= TestAddForm.DateTimePicker1.Date;
  finally
    FreeAndNil(TestAddForm);
  end;
  OpenDatesList(D);
  OpenBeforeTestList;
end;

end.

