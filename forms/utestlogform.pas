unit UTestLogForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, StdCtrls, DividerBevel, VirtualTrees, fpspreadsheetgrid,
  //DK packages utils
  DK_DateUtils, DK_Vector, DK_Matrix, DK_VSTTableTools, DK_Dialogs, DK_Const,
  DK_CtrlUtils,
  //Project utils
  UVars, USheets,
  //Forms
  UTestAddForm;

type

  { TTestLogForm }

  TTestLogForm = class(TForm)
    AddButton: TSpeedButton;
    OrderByNumCheckBox: TCheckBox;
    CheckBox2: TCheckBox;
    DelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    EditButtonPanel: TPanel;
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
    YearPanel: TPanel;
    procedure AddButtonClick(Sender: TObject);
    procedure OrderByNumCheckBoxChange(Sender: TObject);
    procedure CheckBox2Change(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    Months: TStrVector;
    Dates: TDateMatrix;
    SelectedDate: TDate;
    TestIDs, TestResults: TIntVector;
    MotorNames, MotorNums, TestNotes: TStrVector;

    VSTDateList: TVSTCategoryDateList;
    TestLog: TTestLogTable;
    BeforeTestSheet: TBeforeTestSheet;

    procedure SelectMotor;
    procedure SelectDate;

    procedure OpenBeforeTestList;
    procedure OpenDatesList(const ASelectedDate: TDate);

  public
    procedure ViewUpdate;
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
    DataBase.TestBeforeListLoad(MainForm.UsedNameIDs,
                         OrderByNumCheckBox.Checked, BuildDates, MNames, MNums,
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
  VSTDateList:= TVSTCategoryDateList.Create(VT, EmptyStr, @SelectDate);
  TestLog:= TTestLogTable.Create(TestGrid, @SelectMotor, GridFont);
  BeforeTestSheet:= TBeforeTestSheet.Create(LogGrid.Worksheet, LogGrid, GridFont);
  SelectedDate:= Date;
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TTestLogForm.FormDestroy(Sender: TObject);
begin
  if Assigned(TestLog) then FreeAndNil(TestLog);
  if Assigned(BeforeTestSheet) then FreeAndNil(BeforeTestSheet);
  if Assigned(VSTDateList) then FreeAndNil(VSTDateList);
end;

procedure TTestLogForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);
  SetToolButtons([AddButton, DelButton]);
  Images.ToButtons([AddButton, DelButton]);

  ViewUpdate;
end;

procedure TTestLogForm.ViewUpdate;
begin
  OpenDatesList(SelectedDate);
  OpenBeforeTestList;
end;

procedure TTestLogForm.SpinEdit1Change(Sender: TObject);
begin
  OpenDatesList(LastDayInYear(SpinEdit1.Value));
end;

procedure TTestLogForm.SelectMotor;
begin
  if not Assigned(TestLog) then Exit;
  DelButton.Enabled:= TestLog.IsSelected;
end;

procedure TTestLogForm.OpenDatesList(const ASelectedDate: TDate);
begin
  TestGrid.Clear;
  DataBase.MonthAndDatesForTestLogLoad(SpinEdit1.Value, MainForm.UsedNameIDs, Months, Dates);
  VSTDateList.Update(Months, Dates, ASelectedDate);
end;

procedure TTestLogForm.SelectDate;
var
  TotalMotorNames: TStrVector;
  TotalMotorCounts, TotalFailCounts: TIntVector;
  X: TDateVector;
begin
  TestGrid.CLear;
  if not VSTDateList.IsSelected then Exit;

  SelectedDate:= Dates[VSTDateList.SelectedIndex1, VSTDateList.SelectedIndex2];

  DataBase.TestListLoad(SelectedDate,SelectedDate, MainForm.UsedNameIDs,
                    OrderByNumCheckBox.Checked, TestIDs, TestResults, X,
                    MotorNames, MotorNums, TestNotes);
  DataBase.TestTotalLoad(SelectedDate,SelectedDate, MainForm.UsedNameIDs, TotalMotorNames,
                       TotalMotorCounts, TotalFailCounts);

  TestLog.Update(SelectedDate, TestResults, MotorNames, MotorNums, TestNotes,
                 VSum(TotalMotorCounts), VSum(TotalFailCounts));
end;

procedure TTestLogForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить результаты испытаний?') then Exit;
  DataBase.Delete('MOTORTEST', 'TestID', TestIDs[TestLog.SelectedIndex]);
  ViewUpdate;
end;

procedure TTestLogForm.OrderByNumCheckBoxChange(Sender: TObject);
begin
  SelectDate;
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
    if VSTDateList.IsSelected then
      TestAddForm.DateTimePicker1.Date:= Dates[VSTDateList.SelectedIndex1, VSTDateList.SelectedIndex2]
    else
      TestAddForm.DateTimePicker1.Date:= Date;
    TestAddForm.UsedNameID:= 0;
    if Length(MainForm.UsedNameIDs)=1 then
      TestAddForm.UsedNameID:= VFirst(MainForm.UsedNameIDs);
    TestAddForm.ShowModal;
    D:= TestAddForm.DateTimePicker1.Date;
  finally
    FreeAndNil(TestAddForm);
  end;
  OpenDatesList(D);
  OpenBeforeTestList;
end;

end.

