unit UBuildLogForm;

{$mode ObjFPC}{$H+}
interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, StdCtrls, VirtualTrees, DividerBevel, USQLite,
  rxctrls, DK_DateUtils, DK_Vector, DK_Matrix, DK_VSTTools,
  DK_Dialogs, DK_Const, UBuildAddForm, UBuildEditForm,
  USheetUtils, fpspreadsheetgrid, DK_SheetTables;

type

  { TBuildLogForm }

  TBuildLogForm = class(TForm)
    AddButton: TSpeedButton;
    CheckBox1: TCheckBox;
    DelButton: TSpeedButton;
    EditButton: TSpeedButton;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    EditButtonPanel: TPanel;
    DividerBevel1: TDividerBevel;
    LogGrid: TsWorksheetGrid;
    Grid1: TsWorksheetGrid;
    MotorNamesButton: TRxSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    VT: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CheckBox1Change(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LogGridMouseDown(Sender: TObject; Button: TMouseButton;
      {%H-}Shift: TShiftState; X, Y: Integer);
    procedure MotorNamesButtonClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    Months: TStrVector;
    Dates: TDateMatrix;

    VSTDateList: TVSTCategoryDateList;
    MotorBuildSheet: TMotorBuildSheet;

    Test: TSheetTable;

    MotorIDs, NameIDs, OldMotors: TIntVector;
    MotorNames, MotorNums, RotorNums: TStrVector;

    SelectedIndex: Integer;
    procedure ClearSelection;
    procedure SelectLine(const ARow: Integer);

    procedure OpenDatesList(const ASelectDate: TDate);
  public
    procedure ShowBuildLog;
  end;

var
  BuildLogForm: TBuildLogForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TBuildLogForm }

procedure TBuildLogForm.FormCreate(Sender: TObject);
begin
  MainForm.SetNamesPanelsVisible(True, False);
  SelectedIndex:= -1;
  MotorBuildSheet:= TMotorBuildSheet.Create(LogGrid);

  Test:= TSheetTable.Create(Grid1);
  Test.AddColumn('№ п/п', 50);
  Test.AddColumn('Наименование двигателя', 250);
  Test.AddColumn('Номер двигателя', 150);
  Test.AddColumn('Номер ротора', 100);
  Test.AddToHeader(1, 1, 1, 1, '№ п/п');
  Test.AddToHeader(1, 2, 1, 2, 'Наименование двигателя');
  Test.AddToHeader(1, 3, 1, 3, 'Номер двигателя');
  Test.AddToHeader(1, 4, 1, 4, 'Номер ротора');

  VSTDateList:= TVSTCategoryDateList.Create(VT, EmptyStr, @ShowBuildLog);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TBuildLogForm.FormDestroy(Sender: TObject);
begin
  if Assigned(MotorBuildSheet) then FreeAndNil(MotorBuildSheet);
  if Assigned(VSTDateList) then FreeAndNil(VSTDateList);
  if Assigned(Test) then FreeAndNil(Test);
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
    ClearSelection;
  if Button=mbLeft  then
  begin
    (Sender As TsWorksheetGrid).MouseToCell(X,Y,C{%H-},R{%H-});
    SelectLine(R);
  end;
end;

procedure TBuildLogForm.MotorNamesButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Наименования электродвигателей',
    'MOTORNAMES', 'NameID', 'MotorName', False, True) then
      ShowBuildLog;
end;

procedure TBuildLogForm.SpinEdit1Change(Sender: TObject);
begin
  OpenDatesList(LastDayInYear(SpinEdit1.Value));
end;

procedure TBuildLogForm.ClearSelection;
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

procedure TBuildLogForm.OpenDatesList(const ASelectDate: TDate);
begin
  LogGrid.Clear;
  SQLite.MonthAndDatesForBuildLogLoad(SpinEdit1.Value, Months, Dates);
  VSTDateList.Update(Months, Dates, ASelectDate);
end;

procedure TBuildLogForm.ShowBuildLog;
var
  Tmp: TDateVector;
  D: TDate;
begin
  ClearSelection;
  if not VSTDateList.IsSelected then Exit;

  D:= Dates[VSTDateList.SelectedIndex1, VSTDateList.SelectedIndex2];
  SQLite.BuildListLoad(D, D, MainForm.UsedNameIDs,
                    CheckBox1.Checked, MotorIDs, NameIDs, OldMotors,
                    Tmp, MotorNames, MotorNums, RotorNums);

  MotorBuildSheet.Draw(D, MotorNames, MotorNums, RotorNums);


  Test.SetColumnOrder('№ п/п');
  Test.SetColumnString('Наименование двигателя', MotorNames);
  Test.SetColumnString('Номер двигателя', MotorNums);
  Test.SetColumnString('Номер ротора', RotorNums);
  Test.Draw;
end;

procedure TBuildLogForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить двигатель ' +  SYMBOL_BREAK +
                  MotorNames[SelectedIndex] +
                 ' № ' +
                  MotorNums[SelectedIndex] +
                 '?') then Exit;
  SQLite.Delete('MOTORLIST', 'MotorID', MotorIDs[SelectedIndex]);

  OpenDatesList(Dates[VSTDateList.SelectedIndex1, VSTDateList.SelectedIndex2]);
end;

procedure TBuildLogForm.EditButtonClick(Sender: TObject);
var
  BuildEditForm: TBuildEditForm;
  D: TDate;
begin
  BuildEditForm:= TBuildEditForm.Create(BuildLogForm);
  try
    BuildEditForm.DateTimePicker1.Date:= Dates[VSTDateList.SelectedIndex1, VSTDateList.SelectedIndex2];
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
  ShowBuildLog;
end;

procedure TBuildLogForm.AddButtonClick(Sender: TObject);
var
  BuildAddForm: TBuildAddForm;
  D: TDate;
begin
  BuildAddForm:= TBuildAddForm.Create(BuildLogForm);
  try
    if VSTDateList.IsSelected then
      BuildAddForm.DateTimePicker1.Date:= Dates[VSTDateList.SelectedIndex1, VSTDateList.SelectedIndex2]
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

