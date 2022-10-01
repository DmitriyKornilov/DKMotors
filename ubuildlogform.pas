unit UBuildLogForm;

{$mode ObjFPC}{$H+}
interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, StdCtrls, VirtualTrees, DividerBevel, USQLite,
  rxctrls, DK_DateUtils, DK_Vector, DK_Matrix, DK_VSTTables,
  DK_Dialogs, DK_Const, UBuildAddForm, UBuildEditForm,
  SheetUtils, fpspreadsheetgrid;

type

  { TBuildLogForm }

  TBuildLogForm = class(TForm)
    AddButton: TSpeedButton;
    CheckBox1: TCheckBox;
    ChooseMotorNamesButton: TSpeedButton;
    DelButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    EditButton: TSpeedButton;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    EditButtonPanel: TPanel;
    DividerBevel1: TDividerBevel;
    CloseButton: TSpeedButton;
    Label1: TLabel;
    LogGrid: TsWorksheetGrid;
    MotorNamesButton: TRxSpeedButton;
    MotorNamesLabel: TLabel;
    MotorNamesPanel: TPanel;
    Panel1: TPanel;
    Panel2: TPanel;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    VT: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CheckBox1Change(Sender: TObject);
    procedure ChooseMotorNamesButtonClick(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LogGridMouseDown(Sender: TObject; Button: TMouseButton;
      {%H-}Shift: TShiftState; X, Y: Integer);
    procedure MotorNamesButtonClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure VTClick(Sender: TObject);

  private
    Months: TStrVector;
    Dates: TDateMatrix;
    StrDates: TStrMatrix;

    VST: TVSTCategoryRadioButtonTable;

    MotorBuildSheet: TMotorBuildSheet;

    MotorIDs, NameIDs, OldMotors: TIntVector;
    MotorNames, MotorNums, RotorNums: TStrVector;

    UsedNameIDs: TIntVector;
    UsedNames: TStrVector;

    SelectedIndex: Integer;
    procedure SelectionClear;
    procedure SelectLine(const ARow: Integer);

    procedure OpenDatesList(const ASelectDate: TDate);
    procedure OpenMotorsList;

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

  MotorBuildSheet:= TMotorBuildSheet.Create(LogGrid);

  VST:= TVSTCategoryRadioButtonTable.Create(VT);
  VST.SelectedFont.Style:= [fsBold];
  VST.CanRightMouseButtonUnselect:= False;
  VST.HeaderVisible:= False;
  VST.GridLinesVisible:= False;
  VST.AddColumn('Dates');

  SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, False, UsedNameIDs, UsedNames);

  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TBuildLogForm.ChooseMotorNamesButtonClick(Sender: TObject);
begin
  if SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, True, UsedNameIDs, UsedNames) then
    OpenMotorsList;
end;

procedure TBuildLogForm.FormDestroy(Sender: TObject);
begin
  if Assigned(MotorBuildSheet) then FreeAndNil(MotorBuildSheet);
  if Assigned(VST) then FreeAndNil(VST);
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

procedure TBuildLogForm.VTClick(Sender: TObject);
begin
  if VST.IsSelected then
    OpenMotorsList;
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

procedure TBuildLogForm.OpenDatesList(const ASelectDate: TDate);
var
  I1, I2: Integer;
begin
  LogGrid.Clear;
  VST.ValuesClear;

  if not SQLite.MonthAndDatesForBuildLogLoad(SpinEdit1.Value, Months, Dates) then Exit;

  StrDates:= MFormatDateTime('dd.mm.yyyy', Dates);
  VST.SetCategories(Months);
  VST.SetColumn('Dates', StrDates, taLeftJustify);
  VST.Draw;

  if MIndexOf(Dates, ASelectDate, I1, I2) then
  begin
    VST.Select(I1, I2);
    VST.Show(I1, I2);
  end
  else begin
    VST.Select(0, 0);
    VST.Show(0, 0);
  end;
  OpenMotorsList;
end;

procedure TBuildLogForm.OpenMotorsList;
var
  Tmp: TDateVector;
  D: TDate;
begin
  SelectionClear;
  if not VST.IsSelected then Exit;

  D:= Dates[VST.SelectedIndex1, VST.SelectedIndex2];
  SQLite.BuildListLoad(D, D, UsedNameIDs,
                    CheckBox1.Checked, MotorIDs, NameIDs, OldMotors,
                    Tmp, MotorNames, MotorNums, RotorNums);

  MotorBuildSheet.Draw(D, MotorNames, MotorNums, RotorNums);
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

  OpenDatesList(Dates[VST.SelectedIndex1, VST.SelectedIndex2]);
end;

procedure TBuildLogForm.EditButtonClick(Sender: TObject);
var
  BuildEditForm: TBuildEditForm;
  D: TDate;
begin
  BuildEditForm:= TBuildEditForm.Create(BuildLogForm);
  try
    BuildEditForm.DateTimePicker1.Date:= Dates[VST.SelectedIndex1, VST.SelectedIndex2];
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
    if VST.IsSelected then
      BuildAddForm.DateTimePicker1.Date:= Dates[VST.SelectedIndex1, VST.SelectedIndex2]
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

