unit UTestAddForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, DateTimePicker, LCLType, VirtualTrees, DividerBevel,
  //DK packages utils
  DK_VSTTables, DK_Vector, DK_Dialogs, DK_StrUtils, DK_Const, DK_CtrlUtils,
  DK_Filter,
  //Project utils
  UVars, USheets;

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
    FilterPanel: TPanel;
    Panel2: TPanel;
    NormRadioButton: TRadioButton;
    DefectRadioButton: TRadioButton;
    SaveButton: TSpeedButton;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure DefectRadioButtonClick(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure VT1MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure VT2Exit(Sender: TObject);
    procedure VT2KeyUp(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure VT2MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
  private
    FilterString: String;
    Filter: TDKFilter;

    NameIDs: TIntVector;

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

    procedure FilterMotors(const AFilterString: String);
    procedure LoadMotors;

    procedure MotorNumKeyDown(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
  public
    UsedNameID: Integer;
  end;

var
  TestAddForm: TTestAddForm;

implementation

{$R *.lfm}

{ TTestAddForm }

procedure TTestAddForm.FormCreate(Sender: TObject);
begin
  LoadMotorNames;
  TestMotorIDs:= nil;
  TestResults:= nil;
  TestNotes:= nil;

  VSTViewTable:= TVSTTable.Create(VT1);
  VSTTestTable:= TVSTTable.Create(VT2);

  FilterString:= EmptyStr;
  Filter:= DKFilterCreate('Номер двигателя', FilterPanel, @FilterMotors, -1, 300, @MotorNumKeyDown);
end;

procedure TTestAddForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTViewTable);
  FreeAndNil(VSTTestTable);
end;

procedure TTestAddForm.FormShow(Sender: TObject);
var
  Ind: Integer;
begin
  Images.ToButtons([SaveButton, CancelButton, DelButton, AddButton]);
  SetEventButtons([SaveButton, CancelButton, DelButton, AddButton]);

  if UsedNameID>0 then
  begin
    Ind:= VIndexOf(NameIDs, UsedNameID);
    if Ind>=0 then
      MotorNameComboBox.ItemIndex:= Ind;
  end;

  VSTViewTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTViewTable.AddColumn('Дата сборки', 100);
  VSTViewTable.AddColumn('Номер', 100);
  VSTViewTable.AddColumn('Испытания',50);
  VSTViewTable.CanSelect:= True;
  VSTViewTable.Draw;

  VSTTestTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTestTable.AddColumn('№ п/п', 60);
  VSTTestTable.AddColumn('Наименование', 220);
  VSTTestTable.AddColumn('Номер', 100);
  VSTTestTable.AddColumn('Результат', 100);
  VSTTestTable.AddColumn('Примечание');
  VSTTestTable.CanSelect:= True;
  VSTTestTable.Draw;

  DateTimePicker1.SetFocus;
end;

procedure TTestAddForm.LoadMotorNames;
begin
  DataBase.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);
  if VIsNil(NameIDs) then
    Inform('Отсутствует список наименований двигателей!');
end;

procedure TTestAddForm.MotorNameComboBoxChange(Sender: TObject);
begin
  Filter.Focus;
end;

procedure TTestAddForm.MotorNumKeyDown(Sender: TObject; var Key: Word;
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

procedure TTestAddForm.DefectRadioButtonClick(Sender: TObject);
begin
  Memo1.SetFocus;
end;

procedure TTestAddForm.CancelButtonClick(Sender: TObject);
begin
  ModalResult:= mrCancel;
end;

procedure TTestAddForm.SaveButtonClick(Sender: TObject);
begin
  if VIsNil(TestMotorIDs) then
  begin
    Inform('Список испытанных двигателей не заполнен!');
    Exit;
  end;

  if not DataBase.MotorsInTestLogWrite(DateTimePicker1.Date, TestMotorIDs, TestResults,
                                       TestNotes) then Exit;

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

procedure TTestAddForm.DelButtonClick(Sender: TObject);
begin
  DelTest;
end;

procedure TTestAddForm.AddButtonClick(Sender: TObject);
begin
  AddTest;
end;

procedure TTestAddForm.AddTest;
var
  x, MotorID: Integer;
  MotorNum, Note: String;

  procedure TestInfoClear;
  begin
    NormRadioButton.Checked:= True;
    AddButton.Enabled:= False;
    Memo1.Lines.Clear;
    Filter.Clear;
    Filter.Focus;
  end;

begin
  if not VSTViewTable.IsSelected then Exit;

  MotorID:= ViewMotorIDs[VSTViewTable.SelectedIndex];
  MotorNum:=  ViewMotorNums[VSTViewTable.SelectedIndex];

  if VIndexOf(TestMotorIDs, TestResults, MotorID, 0)>=0 then
  begin
    Inform('В списке уже есть успешные испытания ' +
           MotorNameComboBox.Text + ' № ' + MotorNum + '!');
    TestInfoClear;
    Exit;
  end;

  if DataBase.TestFineExists(MotorID) then
    if not Confirm('В базе данных уже есть успешные испытания ' +
                   MotorNameComboBox.Text + ' № ' + MotorNum +
                   '! Всё равно записать?' ) then
  begin
    TestInfoClear;
    Exit;
  end;

  x:= Ord(DefectRadioButton.Checked);
  Note:= STrim(Memo1.Text);

  VAppend(TestMotorIDs, MotorID);
  VAppend(TestResults, x);
  VAppend(TestNotes, Note);

  VAppend(TestMotorNames, MotorNameComboBox.Text);
  VAppend(TestMotorNums, MotorNum);
  if x=0 then
    VAppend(TestResultsStr, 'норма')
  else
    VAppend(TestResultsStr, 'брак');
  ShowTestList(False);

  TestInfoClear;
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

  VSTTestTable.ValuesClear;
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

procedure TTestAddForm.FilterMotors(const AFilterString: String);
begin
  FilterString:= AFilterString;
  LoadMotors;
end;

procedure TTestAddForm.LoadMotors;
begin
  if VIsNil(NameIDs) then Exit;

  DataBase.TestChooseListLoad(NameIDs[MotorNameComboBox.ItemIndex], STrim(FilterString),
                      ViewMotorIDs, ViewMotorNums, ViewBuildDates, ViewTests);

  VSTViewTable.ValuesClear;
  VSTViewTable.SetColumn('Дата сборки', ViewBuildDates);
  VSTViewTable.SetColumn('Номер', ViewMotorNums);
  VSTViewTable.SetColumn('Испытания', ViewTests, taLeftJustify);
  VSTViewTable.Draw;
end;

end.

