unit UTestAddForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, DateTimePicker, DK_Vector, DK_Dialogs, DK_VSTTables, USQLite,
  DK_StrUtils, LCLType, VirtualTrees, DK_Const, DividerBevel, USheetUtils;

type

  { TTestAddForm }

  TTestAddForm = class(TForm)
    AddButton: TSpeedButton;
    ButtonPanel: TPanel;
    CancelButton: TSpeedButton;
    DateTimePicker1: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    Label6: TLabel;
    DelButton: TSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Memo1: TMemo;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel2: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    SaveButton: TSpeedButton;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure MotorNumEditKeyDown(Sender: TObject; var Key: Word;
      {%H-}Shift: TShiftState);
    procedure RadioButton2Click(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure VT1MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure VT2Exit(Sender: TObject);
    procedure VT2KeyUp(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure VT2MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
  private
    NameIDs: TIntVector;
    CanFormClose: Boolean;

    VSTViewTable, VSTTestTable: TVSTTable;

    TestMotorIDs: TIntVector;
    TestResults: TIntVector;
    TestNotes: TStrVector;
    TestMotorNames, TestMotorNums, TestResultsStr: TStrVector;

    ViewMotorIDs: TIntVector;
    ViewBuildDates, ViewMotorNums, ViewTests: TStrVector;

    procedure LoadMotorNames;

    procedure AddTest;
    procedure DelTest;

    procedure ShowTestList(const ANeedSelect: Boolean);

    procedure LoadMotors;

  end;

var
  TestAddForm: TTestAddForm;

implementation

{$R *.lfm}

{ TTestAddForm }

procedure TTestAddForm.LoadMotorNames;
begin
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);

  if VIsNil(NameIDs) then
    ShowInfo('Отсутствует список наименований двигателей!');
end;

procedure TTestAddForm.FormShow(Sender: TObject);
begin
  VSTViewTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTViewTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTViewTable.AddColumn('Дата сборки', 100);
  VSTViewTable.AddColumn('Номер', 100);
  VSTViewTable.AddColumn('Испытания',50);
  VSTViewTable.CanSelect:= True;
  VSTViewTable.Draw;

  VSTTestTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTestTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTTestTable.AddColumn('№ п/п', 60);
  VSTTestTable.AddColumn('Наименование', 220);
  VSTTestTable.AddColumn('Номер', 100);
  VSTTestTable.AddColumn('Результат', 100);
  VSTTestTable.AddColumn('Примечание');
  VSTTestTable.CanSelect:= True;
  VSTTestTable.Draw;

  DateTimePicker1.SetFocus;
end;

procedure TTestAddForm.MotorNameComboBoxChange(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TTestAddForm.MotorNumEditChange(Sender: TObject);
begin
  LoadMotors;
end;

procedure TTestAddForm.MotorNumEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if VIsNil(ViewMotorIDs) then Exit;
  if Key=VK_RETURN then
  begin
    if not VSTViewTable.IsSelected then
    begin
      VSTViewTable.Select(0);
      AddButton.Enabled:= True;
    end
    else AddTest;
  end;
end;

procedure TTestAddForm.RadioButton2Click(Sender: TObject);
begin
  Memo1.SetFocus;
end;

procedure TTestAddForm.SaveButtonClick(Sender: TObject);
begin
  CanFormClose:= False;
  if VIsNil(TestMotorIDs) then
  begin
    ShowInfo('Список испытанных двигателей не заполнен!');
    Exit;
  end;

  SQLite.MotorsInTestLogWrite(DateTimePicker1.Date, TestMotorIDs, TestResults, TestNotes);

  CanFormClose:= True;
  ModalResult:= mrOK;
end;

procedure TTestAddForm.VT1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  AddButton.Enabled:= VSTViewTable.IsSelected;
end;

procedure TTestAddForm.VT2Exit(Sender: TObject);
begin
  VSTTestTable.UnSelect;
  DelButton.Enabled:= False;
end;

procedure TTestAddForm.VT2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE then
    DelTest;
end;

procedure TTestAddForm.VT2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  DelButton.Enabled:= VSTTestTable.IsSelected;
end;

procedure TTestAddForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TTestAddForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
  ModalResult:= mrCancel;
end;

procedure TTestAddForm.DelButtonClick(Sender: TObject);
begin
  DelTest;
end;

procedure TTestAddForm.AddButtonClick(Sender: TObject);
begin
  AddTest;
end;

procedure TTestAddForm.FormCreate(Sender: TObject);
begin
  LoadMotorNames;
  TestMotorIDs:= nil;
  TestResults:= nil;
  TestNotes:= nil;

  VSTViewTable:= TVSTTable.Create(VT1);
  VSTTestTable:= TVSTTable.Create(VT2);

  CanFormClose:= True;
end;

procedure TTestAddForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTViewTable);
  FreeAndNil(VSTTestTable);
end;

procedure TTestAddForm.AddTest;
var
  x: Integer;
  Note: String;
begin
  if not VSTViewTable.IsSelected then Exit;

  x:= Ord(RadioButton2.Checked);
  Note:= STrim(Memo1.Text);

  VAppend(TestMotorIDs, ViewMotorIDs[VSTViewTable.SelectedIndex]);
  VAppend(TestResults, x);
  VAppend(TestNotes, Note);

  VAppend(TestMotorNames, MotorNameComboBox.Text);
  VAppend(TestMotorNums, ViewMotorNums[VSTViewTable.SelectedIndex]);
  if x=0 then
    VAppend(TestResultsStr, 'норма')
  else
    VAppend(TestResultsStr, 'брак');
  ShowTestList(False);

  MotorNumEdit.Text:= EmptyStr;
  Memo1.Lines.Clear;

  RadioButton1.Checked:= True;
  AddButton.Enabled:= False;
  MotorNumEdit.SetFocus;
end;

procedure TTestAddForm.DelTest;
var
  Ind: Integer;
begin
  if not VSTTestTable.IsSelected then Exit;
  Ind:= VSTTestTable.SelectedIndex;

  VDel(TestMotorIDs, Ind);
  VDel(TestResults, Ind);
  VDel(TestNotes, Ind);

  VDel(TestMotorNames, Ind);
  VDel(TestMotorNums, Ind);
  VDel(TestResultsStr, Ind);
  ShowTestList(True);
end;

procedure TTestAddForm.ShowTestList(const ANeedSelect: Boolean);
var
  Ind, LastInd: Integer;
begin
  Ind:= -1;
  if VSTTestTable.IsSelected then
    Ind:= VSTTestTable.SelectedIndex;

  VSTTestTable.SetColumn('№ п/п', VIntToStr(VOrder(Length(TestMotorNames))));
  VSTTestTable.SetColumn('Наименование', TestMotorNames, taLeftJustify);
  VSTTestTable.SetColumn('Номер', TestMotorNums);
  VSTTestTable.SetColumn('Результат', TestResultsStr);
  VSTTestTable.SetColumn('Примечание', TestNotes, taLeftJustify);
  VSTTestTable.Draw;

  if not VIsNil(TestMotorNames) then
  begin
    LastInd:= High(TestMotorNames);
    if (Ind<0) or (Ind>LastInd) then
      Ind:= LastInd;

    if ANeedSelect then
      VSTTestTable.Select(Ind)
    else
      VSTTestTable.Show(LastInd);
  end;

  DelButton.Enabled:= VSTTestTable.IsSelected;
end;

procedure TTestAddForm.LoadMotors;
var
  MotorNumberLike: String;
begin
  if VIsNil(NameIDs) then Exit;

  MotorNumberLike:= STrim(MotorNumEdit.Text);

  SQLite.TestChooseListLoad(NameIDs[MotorNameComboBox.ItemIndex], MotorNumberLike,
                      ViewMotorIDs, ViewMotorNums, ViewBuildDates, ViewTests);

  VSTViewTable.ValuesClear;
  VSTViewTable.SetColumn('Дата сборки', ViewBuildDates);
  VSTViewTable.SetColumn('Номер', ViewMotorNums);
  VSTViewTable.SetColumn('Испытания', ViewTests, taLeftJustify);
  VSTViewTable.Draw;
end;

end.

