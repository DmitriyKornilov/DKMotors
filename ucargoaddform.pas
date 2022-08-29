unit UCargoAddForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  ButtonPanel, Buttons, DateTimePicker, DBUtils, DK_Vector, DK_Dialogs,
  DK_StrUtils, LCLType, CheckLst, DK_Const, Types;

type

  { TCargoAddForm }

  TCargoAddForm = class(TForm)
    AddButton: TSpeedButton;
    ButtonPanel1: TButtonPanel;
    DateTimePicker1: TDateTimePicker;
    DelButton: TSpeedButton;
    SeriesEdit: TEdit;
    FindButton: TSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    ListBox1: TListBox;
    ListBox2: TListBox;
    MotorNameComboBox: TComboBox;
    ReceiverNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel2: TPanel;
    Splitter1: TSplitter;
    procedure AddButtonClick(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);
    procedure CheckListBox1ClickCheck(Sender: TObject);
    procedure CheckListBox1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);

    procedure DateTimePicker1Change(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FindButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure ListBox1Exit(Sender: TObject);
    procedure ListBox2Click(Sender: TObject);
    procedure ListBox2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState
      );
    procedure ListBox2MouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure MotorNumEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure OKButtonClick(Sender: TObject);
    procedure ReceiverNameComboBoxChange(Sender: TObject);
    procedure SeriesEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    MotorID: Integer;
    MotorSeries: String;

    NameIDs: TIntVector;
    MotorNames: TStrVector;

    ReceiverIDs: TIntVector;
    ReceiverNames: TStrVector;

    MotorIDs: TIntVector;
    Series: TStrVector;

    ListMotorIDs: TIntVector;
    ListSeries: TStrVector;

    CanFormClose: Boolean;



    procedure LoadMotorNames;
    procedure LoadReceiverNames;

    procedure ChooseMotor;

    function FindMotors: Boolean;
    procedure AddMotor;
    procedure DelMotor;


  public

  end;

var
  CargoAddForm: TCargoAddForm;

implementation

{$R *.lfm}

{ TCargoAddForm }



procedure TCargoAddForm.LoadMotorNames;
begin
  GetKeyPickList('MOTORNAMES', 'NameID', 'MotorName',
                 NameIDs, MotorNames, True, 'NameID');
  VToStrings(MotorNames, MotorNameComboBox.Items);
  if not VIsNil(MotorNames) then
    MotorNameComboBox.ItemIndex:= 0
  else
    ShowInfo('Отсутствует список наименований двигателей!');
end;

procedure TCargoAddForm.LoadReceiverNames;
begin
  GetKeyPickList('CARGORECEIVERS', 'ReceiverID', 'ReceiverName',
                 ReceiverIDs, ReceiverNames, True, 'ReceiverName');
  VToStrings(ReceiverNames, ReceiverNameComboBox.Items);
  if not VIsNil(ReceiverNames) then
    ReceiverNameComboBox.ItemIndex:= 0
  else
    ShowInfo('Отсутствует список наименований грузополучателей!');
end;



procedure TCargoAddForm.FormShow(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TCargoAddForm.ListBox1Click(Sender: TObject);
begin
  DelButton.Enabled:= ListBox1.ItemIndex>=0;
end;

procedure TCargoAddForm.ListBox1Exit(Sender: TObject);
begin
  DelButton.Enabled:= ListBox1.ItemIndex>=0;
end;

procedure TCargoAddForm.ChooseMotor;
var
  i: Integer;
begin
  if ListBox2.Items.Count=0 then Exit;
  i:= ListBox2.ItemIndex;
  MotorID:= ListMotorIDs[i];
  MotorSeries:= ListSeries[i];
  SeriesEdit.Text:= MotorSeries;

end;

procedure TCargoAddForm.ListBox2Click(Sender: TObject);
begin
  ChooseMotor;
end;

procedure TCargoAddForm.ListBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key=VK_UP) or (Key=VK_DOWN) then ChooseMotor;
  if Key=VK_RETURN then
    SeriesEdit.SetFocus;
end;

procedure TCargoAddForm.ListBox2MouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
  ChooseMotor;
end;

procedure TCargoAddForm.MotorNameComboBoxChange(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TCargoAddForm.MotorNumEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_RETURN then
    FindMotors;
end;

procedure TCargoAddForm.OKButtonClick(Sender: TObject);
begin
  CanFormClose:= False;

  if ReceiverNameComboBox.Text=EmptyStr then
  begin
    ShowInfo('Не указано наименование грузополучателя!');
    Exit;
  end;

  if VIsNil(MotorIDs) then
  begin
    ShowInfo('Список отгруженных двигателей не заполнен!');
    Exit;
  end;

  WriteCargo(DateTimePicker1.Date, ReceiverIDs[ReceiverNameComboBox.ItemIndex],
             MotorIDs, Series);


  CanFormClose:= True;
end;

procedure TCargoAddForm.ReceiverNameComboBoxChange(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TCargoAddForm.SeriesEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_RETURN then AddMotor;
end;

procedure TCargoAddForm.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date:= Date;
  LoadMotorNames;
  LoadReceiverNames;
  CanFormClose:= True;
end;

procedure TCargoAddForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TCargoAddForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
end;

procedure TCargoAddForm.CheckListBox1ClickCheck(Sender: TObject);
begin

end;

procedure TCargoAddForm.CheckListBox1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

end;

//procedure TCargoAddForm.CheckListBox1ClickCheck(Sender: TObject);
//var
//  i, Ind: Integer;
//begin
//  if CheckListBox1.Items.Count=0 then Exit;
//  Ind:= CheckListBox1.ItemIndex;
//  for i:= 0 to CheckListBox1.Items.Count-1 do
//        CheckListBox1.Checked[i]:= i=Ind;
//  MotorID:= ListMotorIDs[Ind];
//  MotorSeries:= ListSeries[Ind];
//end;

//procedure TCargoAddForm.CheckListBox1KeyDown(Sender: TObject; var Key: Word;
//  Shift: TShiftState);
//var
//  i, Ind: Integer;
//begin
//  if CheckListBox1.Items.Count=0 then Exit;
//  if Key=VK_RETURN then
//  begin
//    Ind:= CheckListBox1.ItemIndex;
//    if CheckListBox1.Checked[Ind] then
//    begin
//      MotorID:= ListMotorIDs[Ind];
//      MotorSeries:= ListSeries[Ind];
//      AddMotor;
//    end
//    else begin
//      for i:= 0 to CheckListBox1.Items.Count-1 do
//        CheckListBox1.Checked[i]:= i=Ind;
//    end;
//  end;
//end;




procedure TCargoAddForm.AddButtonClick(Sender: TObject);
begin
  AddMotor;
end;

procedure TCargoAddForm.DateTimePicker1Change(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TCargoAddForm.DelButtonClick(Sender: TObject);
begin
  DelMotor;
end;

procedure TCargoAddForm.FindButtonClick(Sender: TObject);
begin
  FindMotors;
end;





function TCargoAddForm.FindMotors: Boolean;
var
  S: String;
  i: Integer;
  Mv: TIntVector;
  Sv: TStrVector;
begin
  if MotorNameComboBox.Text=EmptyStr then
  begin
    ShowInfo('Не указано наименование двигателя!');
    Exit;
  end;

  S:= STrim(MotorNumEdit.Text);
  if S=EmptyStr then
  begin
    ShowInfo('Не указан номер двигателя!');
    Exit;
  end;

  MotorID:= 0;
  MotorSeries:= EmptyStr;

  //проверка на существование движка в базе
  Result:= FindMotorsToCargo(NameIDs[MotorNameComboBox.ItemIndex], S, Mv, Sv);

  S:= MotorNameComboBox.Text + ' № ' + S;
  if Result then
  begin
    for i:= 0 to High(Mv) do
      if VIndexOf(MotorIDs, Mv[i])=-1 then
    begin
      VAppend(ListMotorIDs, Mv[i]);
      VAppend(ListSeries, Sv[i]);
    end;

    Result:= not VIsNil(ListMotorIDs);
    if Result then
    begin
      for i:= 0 to High(ListMotorIDs) do
        ListBox2.Items.Add(S + ' (' + ListSeries[i] + ')');
      ListBox2.ItemIndex:= 0;

      //for i:= 0 to High(ListMotorIDs) do
      //  CheckListBox1.Items.Add(S + ' (' + ListSeries[i] + ')');
      //CheckListBox1.Checked[0]:= True;
      //CheckListBox1.ItemIndex:= 0;
      MotorID:= ListMotorIDs[0];
      MotorSeries:= ListSeries[0];
      SeriesEdit.Text:= MotorSeries;
    end;
  end;
  ListBox2.Visible:= Result;
  Label5.Visible:= Result;
  SeriesEdit.Visible:= Result;
  //CheckListBox1.Visible:= Result;
  AddButton.Enabled:= Result;

  if not Result then
  begin
    MotorNumEdit.Text:= EmptyStr;
    ShowInfo('Двигатель ' + S  + ' не найден в списке неотгруженных!');
    MotorNumEdit.SetFocus;
  end
  else begin
    ListBox2.SetFocus;
  end;
end;

procedure TCargoAddForm.AddMotor;
var
  S: String;
begin
  VAppend(MotorIDs, MotorID);
  MotorSeries:= STrim(SeriesEdit.Text);
  VAppend(Series, MotorSeries);
  S:= MotorNameComboBox.Text + '   №   ' +
      STrim(MotorNumEdit.Text) +
      '   (партия № ' + MotorSeries + ')';


  ListBox1.Items.Add(S);
  ListBox1.ItemIndex:= ListBox1.Items.Count-1;
  DelButton.Enabled:= True;


  MotorID:= 0;
  MotorSeries:= EmptyStr;
  ListBox2.Items.Clear;
  ListBox2.Visible:= False;
  Label5.Visible:= False;
  SeriesEdit.Visible:= False;
  //CheckListBox1.Items.Clear;
  //CheckListBox1.Visible:= False;
  ListMotorIDs:= nil;
  ListSeries:= nil;
  MotorNumEdit.Text:= EmptyStr;
  AddButton.Enabled:= False;
  MotorNumEdit.SetFocus;
end;

procedure TCargoAddForm.DelMotor;
var
  i: Integer;
begin
  i:= ListBox1.ItemIndex;
  VDel(MotorIDs, i);
  VDel(Series, i);
  ListBox1.Items.Delete(i);
  if ListBox1.Items.Count>0 then
    ListBox1.ItemIndex:= ListBox1.Items.Count-1;
  DelButton.Enabled:= ListBox1.Items.Count>0;
end;



end.

