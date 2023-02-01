unit UTestLogForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, StdCtrls, VirtualTrees, DividerBevel, USQLite,
  DK_DateUtils, DK_Vector, DK_Matrix, DK_VSTTables,
  DK_Dialogs, DK_Const, UTestAddForm, fpspreadsheetgrid,
  USheetUtils;

type

  { TTestLogForm }

  TTestLogForm = class(TForm)
    AddButton: TSpeedButton;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    DelButton: TSpeedButton;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    EditButtonPanel: TPanel;
    DividerBevel1: TDividerBevel;
    LogGrid: TsWorksheetGrid;
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
    procedure DelButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure TestGridMouseDown(Sender: TObject; Button: TMouseButton;
      {%H-}Shift: TShiftState; X, Y: Integer);
    procedure VTClick(Sender: TObject);
  private
    Months: TStrVector;
    Dates: TDateMatrix;
    StrDates: TStrMatrix;

    VST: TVSTCategoryRadioButtonTable;

    TestIDs, TestResults: TIntVector;
    MotorNames, MotorNums, TestNotes: TStrVector;

    SelectedIndex: Integer;

    BeforeTestSheet: TBeforeTestSheet;
    MotorTestSheet: TMotorTestSheet;

    procedure SelectionClear;
    procedure SelectLine(const ARow: Integer);

    procedure OpenBeforeTestList;
    procedure OpenDatesList(const ASelectDate: TDate);
    procedure OpenTestsList;

  public
    procedure ShowTestLog;
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
    SQLite.TestBeforeListLoad(MainForm.UsedNameIDs,
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

procedure TTestLogForm.FormCreate(Sender: TObject);
begin
  MainForm.SetNamesPanelsVisible(True, False);

  SelectedIndex:= -1;

  VST:= TVSTCategoryRadioButtonTable.Create(VT);
  VST.SelectedFont.Style:= [fsBold];
  VST.CanUnselect:= False;
  VST.HeaderVisible:= False;
  VST.GridLinesVisible:= False;
  VST.AddColumn('Dates');


  BeforeTestSheet:= TBeforeTestSheet.Create(LogGrid);
  MotorTestSheet:= TMotorTestSheet.Create(TestGrid);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TTestLogForm.ShowTestLog;
begin
  OpenTestsList;
  OpenBeforeTestList;
end;

procedure TTestLogForm.FormDestroy(Sender: TObject);
begin
  if Assigned(BeforeTestSheet) then FreeAndNil(BeforeTestSheet);
  if Assigned(MotorTestSheet) then FreeAndNil(MotorTestSheet);
  if Assigned(VST) then FreeAndNil(VST);
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

procedure TTestLogForm.VTClick(Sender: TObject);
begin
  if VST.IsSelected then
    OpenTestsList;
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

procedure TTestLogForm.OpenDatesList(const ASelectDate: TDate);
var
  I1, I2: Integer;
begin
  TestGrid.Clear;
  VST.ValuesClear;

  if not SQLite.MonthAndDatesForTestLogLoad(SpinEdit1.Value, Months, Dates) then Exit;
  StrDates:= MFormatDateTime('dd.mm.yyyy', Dates);
  VST.SetCategories(Months);
  VST.SetColumn('Dates', StrDates, taLeftJustify);
  VST.Draw;

  if MIndexOfDate(Dates, ASelectDate, I1, I2) then
  begin
    VST.Select(I1, I2);
    VST.Show(I1, I2);
  end
  else begin
    VST.Select(0, 0);
    VST.Show(0, 0);
  end;
  OpenTestsList;
end;

procedure TTestLogForm.OpenTestsList;
var
  TotalMotorNames: TStrVector;
  TotalMotorCounts, TotalFailCounts: TIntVector;
  X: TDateVector;
  D: TDate;
begin
  SelectionClear;
  if not VST.IsSelected then Exit;

  D:= Dates[VST.SelectedIndex1, VST.SelectedIndex2];

  SQLite.TestListLoad(D,D, MainForm.UsedNameIDs,
                    CheckBox1.Checked, TestIDs, TestResults, X,
                    MotorNames, MotorNums, TestNotes);
  SQLite.TestTotalLoad(D,D, MainForm.UsedNameIDs, TotalMotorNames,
                       TotalMotorCounts, TotalFailCounts);
  MotorTestSheet.Draw(D, MotorNames, MotorNums, TestNotes, TestResults,
                      TotalMotorCounts, TotalFailCounts);
end;

procedure TTestLogForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить результаты испытаний?') then Exit;
  SQLite.Delete('MOTORTEST', 'TestID', TestIDs[SelectedIndex]);
  OpenDatesList(Dates[VST.SelectedIndex1, VST.SelectedIndex2]);
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

procedure TTestLogForm.AddButtonClick(Sender: TObject);
var
  TestAddForm: TTestAddForm;
  D: TDate;
begin
  TestAddForm:= TTestAddForm.Create(TestLogForm);
  try
    if VST.IsSelected then
      TestAddForm.DateTimePicker1.Date:= Dates[VST.SelectedIndex1, VST.SelectedIndex2]
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

