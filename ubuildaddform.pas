unit UBuildAddForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  Buttons, DateTimePicker, DividerBevel, DK_Vector, USQLite, SheetUtils,
  DK_Dialogs, DK_StrUtils, DK_Const, LCLType, VirtualTrees, DK_VSTTables;

type

  { TBuildAddForm }

  TBuildAddForm = class(TForm)
    CancelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    SaveButton: TSpeedButton;
    OldMotorCheckBox: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    DelButton: TSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel1: TPanel;
    Panel2: TPanel;
    ButtonPanel: TPanel;
    RotorNumEdit: TEdit;
    AddButton: TSpeedButton;
    VT2: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure MotorNumEditKeyDown(Sender: TObject; var Key: Word;
      {%H-}Shift: TShiftState);
    procedure RotorNumEditKeyDown(Sender: TObject; var Key: Word;
      {%H-}Shift: TShiftState);
    procedure SaveButtonClick(Sender: TObject);
    procedure VT2Exit(Sender: TObject);
    procedure VT2KeyUp(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure VT2MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
  private
    NameIDs, OldMotors: TIntVector;

    MotorNameIDs: TIntVector;
    MotorNums, RotorNums, MotorNames: TStrVector;
    CanFormClose: Boolean;

    VSTTable: TVSTTable;

    procedure LoadMotorNames;
    procedure AddMotor;
    procedure DelMotor;

    procedure ShowBuildList(const ANeedSelect: Boolean);
  end;

var
  BuildAddForm: TBuildAddForm;

implementation

{$R *.lfm}

procedure TBuildAddForm.FormCreate(Sender: TObject);
begin
  LoadMotorNames;
  VSTTable:= TVSTTable.Create(VT2);
  CanFormClose:= True;
end;

procedure TBuildAddForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTTable);
end;

procedure TBuildAddForm.FormShow(Sender: TObject);
begin
  VSTTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTTable.AddColumn('??? ??/??', 60);
  VSTTable.AddColumn('????????????????????????', 220);
  VSTTable.AddColumn('??????????', 100);
  VSTTable.AddColumn('??????????');
  VSTTable.CanSelect:= True;
  VSTTable.Draw;
  DateTimePicker1.SetFocus;
end;

procedure TBuildAddForm.MotorNameComboBoxChange(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TBuildAddForm.MotorNumEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_RETURN then RotorNumEdit.SetFocus;
end;

procedure TBuildAddForm.AddButtonClick(Sender: TObject);
begin
  AddMotor;
end;

procedure TBuildAddForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
  ModalResult:= mrCancel;
end;

procedure TBuildAddForm.DelButtonClick(Sender: TObject);
begin
  DelMotor;
end;

procedure TBuildAddForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TBuildAddForm.RotorNumEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_RETURN then AddMotor;
end;

procedure TBuildAddForm.SaveButtonClick(Sender: TObject);
begin
  CanFormClose:= False;

  if MotorNameComboBox.Text=EmptyStr then
  begin
    ShowInfo('???? ?????????????? ???????????????????????? ??????????????????!');
    Exit;
  end;

  if VIsNil(MotorNums) then
  begin
    ShowInfo('???? ?????????????? ???????????? ????????????????????!');
    Exit;
  end;

  SQLite.MotorsInBuildLogWrite(DateTimePicker1.Date, MotorNameIDs, OldMotors,
                               MotorNums, RotorNums);

  CanFormClose:= True;
  ModalResult:= mrOK;
end;

procedure TBuildAddForm.VT2Exit(Sender: TObject);
begin
  VSTTable.UnSelect;
  DelButton.Enabled:= False;
end;

procedure TBuildAddForm.VT2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE then
    DelMotor;
end;

procedure TBuildAddForm.VT2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  DelButton.Enabled:= VSTTable.IsSelected;
end;

procedure TBuildAddForm.LoadMotorNames;
begin
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);

  if not VIsNil(NameIDs) then
    AddButton.Enabled:= True
  else
    ShowInfo('?????????????????????? ???????????? ???????????????????????? ????????????????????!');
end;

procedure TBuildAddForm.AddMotor;
var
  s1,s2: String;
  n: Integer;

  procedure CancelWrite;
  begin
    RotorNumEdit.Text:= EmptyStr;
    MotorNumEdit.Text:= EmptyStr;
    MotorNumEdit.SetFocus;
  end;

begin
  s1:= STrim(MotorNumEdit.Text);
  if s1=EmptyStr then
  begin
    ShowInfo('???? ???????????? ?????????? ??????????????????!');
    CancelWrite;
    Exit;
  end;

  //???????????????? ???? ?????????????? ?? ???????????? ?? ????????????
  n:= VIndexOf(MotorNums, s1);
  if n>=0 then
   if MotorNameIDs[n] = NameIDs[MotorNameComboBox.ItemIndex] then
  begin
    ShowInfo('??????????????????' + SYMBOL_BREAK +
              MotorNameComboBox.Text + ' ??? ' + s1 + SYMBOL_BREAK +
             '?????? ???????? ?? ????????????!');
    CancelWrite;
    Exit;
  end;

  //???????????????? ???? ?????????????? ?? ????????
  if SQLite.IsDuplicateMotorNumber(DateTimePicker1.Date,
                         NameIDs[MotorNameComboBox.ItemIndex], s1) then
  begin
    ShowInfo('??????????????????' + SYMBOL_BREAK +
              MotorNameComboBox.Text + ' ??? ' + s1 + SYMBOL_BREAK +
             '?????? ???????? ?? ????????!');
    CancelWrite;
    Exit;
  end;

  s2:= STrim(RotorNumEdit.Text);

  VAppend(MotorNames, MotorNameComboBox.Text);
  VAppend(MotorNums, s1);
  VAppend(RotorNums, s2);
  VAppend(OldMotors, Ord(OldMotorCheckBox.Checked));
  VAppend(MotorNameIDs, NameIDs[MotorNameComboBox.ItemIndex]);

  ShowBuildList(False);

  CancelWrite;
end;

procedure TBuildAddForm.DelMotor;
var
  Ind: Integer;
begin
  if not VSTTable.IsSelected then Exit;
  Ind:= VSTTable.SelectedIndex;
  VDel(MotorNames, Ind);
  VDel(MotorNums, Ind);
  VDel(RotorNums, Ind);
  VDel(OldMotors, Ind);
  VDel(MotorNameIDs, Ind);
  ShowBuildList(True);
end;

procedure TBuildAddForm.ShowBuildList(const ANeedSelect: Boolean);
var
  Ind, LastInd: Integer;
begin
  Ind:= -1;
  if VSTTable.IsSelected then
    Ind:= VSTTable.SelectedIndex;

  VSTTable.SetColumn('??? ??/??', VIntToStr(VOrder(Length(MotorNames))));
  VSTTable.SetColumn('????????????????????????', MotorNames, taLeftJustify);
  VSTTable.SetColumn('??????????', MotorNums);
  VSTTable.SetColumn('??????????', RotorNums);

  VSTTable.Draw;

  if not VIsNil(MotorNames) then
  begin
    LastInd:= High(MotorNames);
    if (Ind<0) or (Ind>LastInd) then
      Ind:= LastInd;

    if ANeedSelect then
      VSTTable.Select(Ind)
    else
      VSTTable.Show(LastInd);
  end;

  DelButton.Enabled:= VSTTable.IsSelected;

end;

end.

