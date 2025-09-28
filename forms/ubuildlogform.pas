unit UBuildLogForm;

{$mode ObjFPC}{$H+}
interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, StdCtrls, VirtualTrees,  fpspreadsheetgrid,
  //DK packages utils
  DK_DateUtils, DK_Vector, DK_Matrix, DK_VSTTableTools, DK_Dialogs, DK_Const,
  //Project utils
  UDataBase, UUtils, USheetUtils,
  //Forms
  UBuildAddForm, UBuildEditForm;

type

  { TBuildLogForm }

  TBuildLogForm = class(TForm)
    AddButton: TSpeedButton;
    Bevel1: TBevel;
    Bevel2: TBevel;
    CheckBox1: TCheckBox;
    DelButton: TSpeedButton;
    EditButton: TSpeedButton;
    EditButtonPanel: TPanel;
    LogGrid: TsWorksheetGrid;
    Panel1: TPanel;
    Panel2: TPanel;
    SpinEdit1: TSpinEdit;
    YearPanel: TPanel;
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
    procedure LogGridDblClick(Sender: TObject);
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

    procedure EditMotor;

    procedure OpenDatesList(const ASelectedDate: TDate);
  public
    procedure ViewUpdate;
  end;

var
  BuildLogForm: TBuildLogForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TBuildLogForm }

procedure TBuildLogForm.FormCreate(Sender: TObject);
begin
  SetToolPanels([
    ToolPanel
  ]);
  SetToolButtons([
    AddButton, DelButton, EditButton
  ]);

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
  ViewUpdate;
end;

procedure TBuildLogForm.LogGridDblClick(Sender: TObject);
begin
  if not BuildLog.IsSelected then Exit;
  EditMotor;
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
  DataBase.BuildListLoad(SelectedDate, SelectedDate, MainForm.UsedNameIDs,
                    CheckBox1.Checked, MotorIDs, NameIDs, OldMotors,
                    Tmp, MotorNames, MotorNums, RotorNums);
  BuildLog.Update(SelectedDate, MotorNames, MotorNums, RotorNums);
end;

procedure TBuildLogForm.OpenDatesList(const ASelectedDate: TDate);
begin
  LogGrid.Clear;
  DataBase.MonthAndDatesForBuildLogLoad(SpinEdit1.Value, MainForm.UsedNameIDs, Months, Dates);
  VSTDateList.Update(Months, Dates, ASelectedDate);
end;

procedure TBuildLogForm.ViewUpdate;
begin
  BuildLog.Unselect;
  OpenDatesList(SelectedDate);
end;

procedure TBuildLogForm.DelButtonClick(Sender: TObject);
begin
  if not Confirm('Удалить двигатель ' +  SYMBOL_BREAK +
                  MotorNames[BuildLog.SelectedIndex] +
                 ' № ' +
                  MotorNums[BuildLog.SelectedIndex] +
                 '?') then Exit;
  DataBase.Delete('MOTORLIST', 'MotorID', MotorIDs[BuildLog.SelectedIndex]);
  ViewUpdate;
end;

procedure TBuildLogForm.EditMotor;
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
    if BuildEditForm.ShowModal=mrOK then
    begin
      D:= BuildEditForm.DateTimePicker1.Date;
      BuildLog.Unselect;
      OpenDatesList(D);
    end;
  finally
    FreeAndNil(BuildEditForm);
  end;
end;

procedure TBuildLogForm.EditButtonClick(Sender: TObject);
begin
  EditMotor;
end;

procedure TBuildLogForm.CheckBox1Change(Sender: TObject);
begin
  ViewUpdate;
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

