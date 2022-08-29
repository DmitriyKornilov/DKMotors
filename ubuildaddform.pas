unit UBuildAddForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  Buttons, DateTimePicker, DividerBevel, DK_Vector, USQLite,
  DK_Dialogs, DK_StrUtils, DK_Const, LCLType;

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
    ListBox1: TListBox;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel1: TPanel;
    Panel2: TPanel;
    ButtonPanel: TPanel;
    RotorNumEdit: TEdit;
    AddButton: TSpeedButton;
    Splitter1: TSplitter;
    procedure AddButtonClick(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure ListBox1Exit(Sender: TObject);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure MotorNumEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RotorNumEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SaveButtonClick(Sender: TObject);
  private
    NameIDs, OldMotors: TIntVector;

    MotorNameIDs: TIntVector;
    MotorNums, RotorNums: TstrVector;
    CanFormClose: Boolean;

    procedure LoadMotorNames;
    procedure WriteMotor;
    procedure DeleteMotor;
  end;

var
  BuildAddForm: TBuildAddForm;

implementation

{$R *.lfm}

procedure TBuildAddForm.FormCreate(Sender: TObject);
begin
  LoadMotorNames;
  CanFormClose:= True;
end;

procedure TBuildAddForm.FormShow(Sender: TObject);
begin
  DateTimePicker1.SetFocus;
end;

procedure TBuildAddForm.ListBox1Click(Sender: TObject);
begin
  DelButton.Enabled:= ListBox1.ItemIndex>=0;
end;

procedure TBuildAddForm.ListBox1Exit(Sender: TObject);
begin
  DelButton.Enabled:= ListBox1.ItemIndex>=0;
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
  WriteMotor;
end;

procedure TBuildAddForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
  ModalResult:= mrCancel;
end;

procedure TBuildAddForm.DelButtonClick(Sender: TObject);
begin
  DeleteMotor;
end;

procedure TBuildAddForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TBuildAddForm.RotorNumEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_RETURN then WriteMotor;
end;

procedure TBuildAddForm.SaveButtonClick(Sender: TObject);
begin
  CanFormClose:= False;

  if MotorNameComboBox.Text=EmptyStr then
  begin
    ShowInfo('Не указано наименование двигателя!');
    Exit;
  end;

  if VIsNil(MotorNums) then
  begin
    ShowInfo('Не указаны номера двигателей!');
    Exit;
  end;

  //MotorsAdd(DateTimePicker1.Date, MotorNameIDs, OldMotors, MotorNums, RotorNums);
  SQLite.MotorsInBuildLogWrite(DateTimePicker1.Date, MotorNameIDs, OldMotors,
                               MotorNums, RotorNums);

  CanFormClose:= True;
  ModalResult:= mrOK;
end;

procedure TBuildAddForm.LoadMotorNames;
begin
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);

  if not VIsNil(NameIDs) then
    AddButton.Enabled:= True
  else
    ShowInfo('Отсутствует список наименований двигателей!');
end;

procedure TBuildAddForm.WriteMotor;
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
    ShowInfo('Не указан номер двигателя!');
    CancelWrite;
    Exit;
  end;

  //проверка на наличие в списке к записи
  n:= VIndexOf(MotorNums, s1);
  if n>=0 then
   if MotorNameIDs[n] = NameIDs[MotorNameComboBox.ItemIndex] then
  begin
    ShowInfo('Двигатель' + SYMBOL_BREAK +
              MotorNameComboBox.Text + ' № ' + s1 + SYMBOL_BREAK +
             'уже есть в списке!');
    CancelWrite;
    Exit;
  end;

  //проверка на наличие в базе
  if SQLite.IsDuplicateMotorNumber(DateTimePicker1.Date,
                         NameIDs[MotorNameComboBox.ItemIndex], s1) then
  begin
    ShowInfo('Двигатель' + SYMBOL_BREAK +
              MotorNameComboBox.Text + ' № ' + s1 + SYMBOL_BREAK +
             'уже есть в базе!');
    CancelWrite;
    Exit;
  end;

  s2:= STrim(RotorNumEdit.Text);

  VAppend(MotorNums, s1);
  VAppend(RotorNums, s2);
  VAppend(OldMotors, Ord(OldMotorCheckBox.Checked));
  VAppend(MotorNameIDs, NameIDs[MotorNameComboBox.ItemIndex]);
  s1:= MotorNameComboBox.Text + ' № ' + s1;
  if s2<>EmptyStr then
    s1:= s1 + ' (ротор №' + s2 + ')';
  if OldMotorCheckBox.Checked then
    S1:= s1 + ' - ' + OldMotorCheckBox.Caption;
  ListBox1.Items.Add(s1);
  ListBox1.ItemIndex:= ListBox1.Count-1;
  DelButton.Enabled:= True;

  CancelWrite;
end;

procedure TBuildAddForm.DeleteMotor;
var
  i: Integer;
begin
  i:= ListBox1.ItemIndex;
  VDel(MotorNums, i);
  VDel(RotorNums, i);
  VDel(OldMotors, i);
  VDel(MotorNameIDs, i);
  ListBox1.Items.Delete(i);
  if ListBox1.Items.Count>0 then
    ListBox1.ItemIndex:= ListBox1.Items.Count-1;
  DelButton.Enabled:= ListBox1.Items.Count>0;
end;

end.

