unit UBuildLogForm;

{$mode ObjFPC}{$H+}
interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, StdCtrls, VirtualTrees, DividerBevel, USQLite,
  rxctrls, DK_DateUtils, DK_Vector, DK_Matrix, DK_VSTTools,
  DK_Dialogs, DK_Const, UBuildAddForm, UBuildEditForm,
  USheetUtils, fpspreadsheetgrid;

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
    procedure MotorNamesButtonClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    Months: TStrVector;
    Dates: TDateMatrix;
    SelectedDate: TDate;

    VSTDateList: TVSTCategoryDateList;
    BuildLog: TBuildLogTable;

    MotorIDs, NameIDs, OldMotors: TIntVector;
    MotorNames, MotorNums, RotorNums: TStrVector;

    procedure SelectMotor;
    procedure SelectDate;

    procedure OpenDatesList(const ASelectedDate: TDate);
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
  BuildLog:= TBuildLogTable.Create(LogGrid, @SelectMotor);
  VSTDateList:= TVSTCategoryDateList.Create(VT, EmptyStr, @SelectDate);
  SelectedDate:= Date;
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TBuildLogForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTDateList) then FreeAndNil(VSTDateList);
  if Assigned(BuildLog) then FreeAndNil(BuildLog);
end;

procedure TBuildLogForm.FormShow(Sender: TObject);
begin
  ShowBuildLog;
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

procedure TBuildLogForm.SelectMotor;
begin
  if not Assigned(BuildLog) then Exit;
  DelButton.Enabled:= BuildLog.IsSelected;
  EditButton.Enabled:= DelButton.Enabled;
end;

procedure TBuildLogForm.SelectDate;
var
  Tmp: TDateVector;
begin
  LogGrid.Clear;
  if not VSTDateList.IsSelected then Exit;

  SelectedDate:= Dates[VSTDateList.SelectedIndex1, VSTDateList.SelectedIndex2];
  SQLite.BuildListLoad(SelectedDate, SelectedDate, MainForm.UsedNameIDs,
                    CheckBox1.Checked, MotorIDs, NameIDs, OldMotors,
                    Tmp, MotorNames, MotorNums, RotorNums);
  BuildLog.Update(SelectedDate, MotorNames, MotorNums, RotorNums);
end;

procedure TBuildLogForm.OpenDatesList(const ASelectedDate: TDate);
begin
  LogGrid.Clear;
  SQLite.MonthAndDatesForBuildLogLoad(SpinEdit1.Value, MainForm.UsedNameIDs, Months, Dates);
  VSTDateList.Update(Months, Dates, ASelectedDate);
end;

procedure TBuildLogForm.ShowBuildLog;
begin
  OpenDatesList(SelectedDate);
end;

procedure TBuildLogForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить двигатель ' +  SYMBOL_BREAK +
                  MotorNames[BuildLog.SelectedIndex] +
                 ' № ' +
                  MotorNums[BuildLog.SelectedIndex] +
                 '?') then Exit;
  SQLite.Delete('MOTORLIST', 'MotorID', MotorIDs[BuildLog.SelectedIndex]);
  ShowBuildLog;
end;

procedure TBuildLogForm.EditButtonClick(Sender: TObject);
var
  BuildEditForm: TBuildEditForm;
  D: TDate;
begin
  BuildEditForm:= TBuildEditForm.Create(BuildLogForm);
  try
    BuildEditForm.DateTimePicker1.Date:= Dates[VSTDateList.SelectedIndex1, VSTDateList.SelectedIndex2];
    BuildEditForm.MotorID:= MotorIDs[BuildLog.SelectedIndex];
    BuildEditForm.OldNameID:= NameIDs[BuildLog.SelectedIndex];
    BuildEditForm.MotorNumEdit.Text:= MotorNums[BuildLog.SelectedIndex];
    BuildEditForm.RotorNumEdit.Text:= RotorNums[BuildLog.SelectedIndex];
    BuildEditForm.OldMotorCheckBox.Checked:= OldMotors[BuildLog.SelectedIndex]=1;
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
    BuildAddForm.UsedNameID:= 0;
    if Length(MainForm.UsedNameIDs)=1 then
      BuildAddForm.UsedNameID:= VFirst(MainForm.UsedNameIDs);
    BuildAddForm.ShowModal;
    D:= BuildAddForm.DateTimePicker1.Date;
  finally
    FreeAndNil(BuildAddForm);
  end;
  OpenDatesList(D);
end;

end.

