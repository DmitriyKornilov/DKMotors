unit UTestAddForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, DateTimePicker, DK_Vector, DK_Dialogs, DK_VSTTables, USQLite,
  DK_StrUtils, LCLType, VirtualTrees, DK_Const, DividerBevel, SheetUtils;

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
    ListBox1: TListBox;
    Memo1: TMemo;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel2: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    SaveButton: TSpeedButton;
    Splitter1: TSplitter;
    VT1: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure ListBox1Exit(Sender: TObject);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure MotorNumEditKeyDown(Sender: TObject; var Key: Word;
      {%H-}Shift: TShiftState);
    procedure RadioButton2Click(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure VT1MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
  private
    NameIDs: TIntVector;
    CanFormClose: Boolean;

    VSTTable: TVSTTable;

    TestMotorIDs: TIntVector;
    TestResults: TIntVector;
    TestNotes: TStrVector;

    ViewMotorIDs: TIntVector;
    ViewBuildDates, ViewMotorNums, ViewTests: TStrVector;

    procedure LoadMotorNames;

    procedure AddTest;
    procedure DelTest;

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
  VSTTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTTable.AddColumn('Дата сборки', 100);
  VSTTable.AddColumn('Номер', 100);
  VSTTable.AddColumn('Испытания',50);
  VSTTable.CanSelect:= True;
  VSTTable.Draw;

  DateTimePicker1.SetFocus;
end;

procedure TTestAddForm.ListBox1Click(Sender: TObject);
begin
  DelButton.Enabled:= ListBox1.ItemIndex>=0;
end;

procedure TTestAddForm.ListBox1Exit(Sender: TObject);
begin
  DelButton.Enabled:= ListBox1.ItemIndex>=0;
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
    if not VSTTable.IsSelected then
    begin
      VSTTable.Select(0);
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
  AddButton.Enabled:= VSTTable.IsSelected;
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

  VSTTable:= TVSTTable.Create(VT1);
  CanFormClose:= True;
end;

procedure TTestAddForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTTable);
end;

procedure TTestAddForm.AddTest;
var
  x: Integer;
  s, Note: String;
begin
  if not VSTTable.IsSelected then Exit;

  x:= Ord(RadioButton2.Checked);
  Note:= STrim(Memo1.Text);

  VAppend(TestMotorIDs, ViewMotorIDs[VSTTable.SelectedIndex]);
  VAppend(TestResults, x);
  VAppend(TestNotes, Note);

  s:= MotorNameComboBox.Text + ' № ' + ViewMotorNums[VSTTable.SelectedIndex];
  if x=0 then
    s:= s + ' -  норма'
  else
    s:= s + ' -  брак';
  if Note<>EmptyStr then
    s:= s + ' (' + Note + ')';

  ListBox1.Items.Add(s);
  ListBox1.ItemIndex:= ListBox1.Items.Count-1;

  MotorNumEdit.Text:= EmptyStr;
  Memo1.Lines.Clear;

  RadioButton1.Checked:= True;
  AddButton.Enabled:= False;
  DelButton.Enabled:= True;
  MotorNumEdit.SetFocus;
end;

procedure TTestAddForm.DelTest;
var
  i: Integer;
begin
  i:= ListBox1.ItemIndex;
  VDel(TestMotorIDs, i);
  VDel(TestResults, i);
  VDel(TestNotes, i);
  ListBox1.Items.Delete(i);
  if ListBox1.Items.Count>0 then
    ListBox1.ItemIndex:= ListBox1.Items.Count-1;
  DelButton.Enabled:= ListBox1.Items.Count>0;
end;

procedure TTestAddForm.LoadMotors;
var
  MotorNumberLike: String;
begin
  if VIsNil(NameIDs) then Exit;

  MotorNumberLike:= STrim(MotorNumEdit.Text);

  SQLite.TestChooseListLoad(NameIDs[MotorNameComboBox.ItemIndex], MotorNumberLike,
                      ViewMotorIDs, ViewMotorNums, ViewBuildDates, ViewTests);

  VSTTable.ValuesClear;
  VSTTable.SetColumn('Дата сборки', ViewBuildDates);
  VSTTable.SetColumn('Номер', ViewMotorNums);
  VSTTable.SetColumn('Испытания', ViewTests, taLeftJustify);
  VSTTable.Draw;
end;

end.

