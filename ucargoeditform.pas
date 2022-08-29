unit UCargoEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, DateTimePicker, DK_Vector, DK_Dialogs, DK_VSTTables,
  DK_StrUtils, LCLType, VirtualTrees, DK_Const, DividerBevel, USQLite,
  SheetUtils;

type

  { TCargoEditForm }

  TCargoEditForm = class(TForm)
    AddButton: TSpeedButton;
    ButtonPanel: TPanel;
    CancelButton: TSpeedButton;
    DateTimePicker1: TDateTimePicker;
    DelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    SaveButton: TSpeedButton;
    SeriesEdit: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    ListBox1: TListBox;
    MotorNameComboBox: TComboBox;
    ReceiverNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel2: TPanel;
    Splitter1: TSplitter;
    VT1: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);


    procedure DateTimePicker1Change(Sender: TObject);
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
      Shift: TShiftState);
    procedure ReceiverNameComboBoxChange(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure SeriesEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure VT1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    NameIDs: TIntVector;
    MotorNames: TStrVector;

    ReceiverIDs: TIntVector;
    ReceiverNames: TStrVector;

    ViewMotorIDs: TIntVector;
    ViewMotorNums, ViewBuildDates, ViewSeries: TStrVector;

    MotorIDs: TIntVector;
    Series: TStrVector;

    WritedMotorIDs: TIntVector;

    CanFormClose: Boolean;

    VSTTable: TVSTTable;

    procedure LoadMotorNames;
    procedure LoadReceiverNames;

    procedure AddMotor;
    procedure DelMotor;

    procedure LoadMotors;
    procedure LoadCargo;
  public
    CargoID: Integer;
  end;

var
  CargoEditForm: TCargoEditForm;

implementation

{$R *.lfm}

{ TCargoEditForm }

procedure TCargoEditForm.LoadMotorNames;
begin
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);
  if VIsNil(NameIDs) then
    ShowInfo('Отсутствует список наименований двигателей!');
end;

procedure TCargoEditForm.LoadReceiverNames;
begin
  SQLite.ReceiverIDsAndNamesLoad(ReceiverNameComboBox, ReceiverIDs);
  if VIsNil(ReceiverIDs) then
    ShowInfo('Отсутствует список наименований грузополучателей!');
end;

procedure TCargoEditForm.FormShow(Sender: TObject);
begin
  if CargoID>0 then LoadCargo;

  VSTTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTTable.AddColumn('Дата сборки', 100);
  VSTTable.AddColumn('Номер', 100);
  VSTTable.AddColumn('Партия',50);
  VSTTable.CanSelect:= True;
  VSTTable.Draw;

  DateTimePicker1.SetFocus;
end;

procedure TCargoEditForm.ListBox1Click(Sender: TObject);
begin
  DelButton.Enabled:= ListBox1.ItemIndex>=0;
end;

procedure TCargoEditForm.ListBox1Exit(Sender: TObject);
begin
  DelButton.Enabled:= ListBox1.ItemIndex>=0;
end;

procedure TCargoEditForm.MotorNameComboBoxChange(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TCargoEditForm.MotorNumEditChange(Sender: TObject);
begin
  LoadMotors;
end;

procedure TCargoEditForm.MotorNumEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if VIsNil(ViewMotorIDs) then Exit;
  if Key=VK_RETURN then
  begin
    VSTTable.SelectIndex(0);
    AddButton.Enabled:= True;
    SeriesEdit.Text:= ViewSeries[VSTTable.SelectedIndex];
    SeriesEdit.SetFocus;
  end;
end;

procedure TCargoEditForm.ReceiverNameComboBoxChange(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TCargoEditForm.SaveButtonClick(Sender: TObject);
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

  if CargoID=0 then
  begin
    SQLite.CargoWrite(DateTimePicker1.Date, ReceiverIDs[ReceiverNameComboBox.ItemIndex],
               MotorIDs, Series);
    CargoID:= SQLite.LastWritedInt32ID('CARGOLIST');
  end
  else if CargoID>0 then
    SQLite.CargoUpdate(CargoID,
               DateTimePicker1.Date, ReceiverIDs[ReceiverNameComboBox.ItemIndex],
               MotorIDs, WritedMotorIDs, Series);

  CanFormClose:= True;
  ModalResult:= mrOK;
end;

procedure TCargoEditForm.SeriesEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if not VSTTable.IsSelected then Exit;
  if Key=VK_RETURN then AddMotor;
end;

procedure TCargoEditForm.VT1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  AddButton.Enabled:= VSTTable.IsSelected;
  if VSTTable.IsSelected then
    SeriesEdit.Text:= ViewSeries[VSTTable.SelectedIndex]
  else
    SeriesEdit.Text:= EmptyStr;
end;

procedure TCargoEditForm.FormCreate(Sender: TObject);
begin
  VSTTable:= TVSTTable.Create(VT1);
  DateTimePicker1.Date:= Date;
  LoadMotorNames;
  LoadReceiverNames;
  CanFormClose:= True;
end;

procedure TCargoEditForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTTable)
end;

procedure TCargoEditForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TCargoEditForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
  ModalResult:= mrCancel;
end;

procedure TCargoEditForm.AddButtonClick(Sender: TObject);
begin
  AddMotor;
end;

procedure TCargoEditForm.DateTimePicker1Change(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;

procedure TCargoEditForm.DelButtonClick(Sender: TObject);
begin
  DelMotor;
end;

procedure TCargoEditForm.AddMotor;
var
  S, MotorSeries: String;
  MotorID: Integer;
begin
  if not VSTTable.IsSelected then Exit;

  MotorID:= ViewMotorIDs[VSTTable.SelectedIndex];
  VAppend(MotorIDs, MotorID);

  MotorSeries:= STrim(SeriesEdit.Text);
  if MotorSeries=EmptyStr then
  begin
    ShowInfo('Не указан номер партии!');
    Exit;
  end;

  VAppend(Series, MotorSeries);
  S:= MotorNameComboBox.Text + '   №   ' +
      ViewMotorNums[VSTTable.SelectedIndex] +
      '   (партия № ' + MotorSeries + ')';

  ListBox1.Items.Add(S);
  ListBox1.ItemIndex:= ListBox1.Items.Count-1;
  DelButton.Enabled:= True;

  MotorNumEdit.Text:= EmptyStr;
  AddButton.Enabled:= False;
  MotorNumEdit.SetFocus;
end;

procedure TCargoEditForm.DelMotor;
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

procedure TCargoEditForm.LoadMotors;
var
  MotorNumberLike: String;
  i, n: Integer;
begin
  SeriesEdit.Text:= EmptyStr;
  AddButton.Enabled:= False;
  if VIsNil(NameIDs) then Exit;

  MotorNumberLike:= STrim(MotorNumEdit.Text);

  SQLite.CargoChooseListLoad(NameIDs[MotorNameComboBox.ItemIndex], MotorNumberLike,
                      ViewMotorIDs, ViewMotorNums, ViewBuildDates, ViewSeries);

  //убираем из списка просмотра двигатели, которые уже есть в списке этой отгрузки
  for i:=0 to High(MotorIDs) do
  begin
    n:= VIndexOf(ViewMotorIDs, MotorIDs[i]);
    if n>=0 then
    begin
      VDel(ViewMotorIDs, n);
      VDel(ViewMotorNums, n);
      VDel(ViewBuildDates, n);
      VDel(ViewSeries, n);
    end;
  end;

  VSTTable.ValuesClear;
  VSTTable.SetColumn('Дата сборки', ViewBuildDates);
  VSTTable.SetColumn('Номер', ViewMotorNums);
  VSTTable.SetColumn('Партия', ViewSeries{, taLeftJustify});
  VSTTable.Draw;
end;

procedure TCargoEditForm.LoadCargo;
var
  MNames, MNums: TStrVector;
  i: Integer;
  D: TDate;
  S: String;
begin
  if not SQLite.CargoMotorListLoad(CargoID,i, D, MotorIDs,
                                   MNames, MNums, Series) then Exit;

  DateTimePicker1.Date:= D;
  ReceiverNameComboBox.ItemIndex:= VIndexOf(ReceiverIDs, i);

  WritedMotorIDs:= VCut(MotorIDs);

  for i:= 0 to High(MotorIDs) do
  begin
    S:= MNames[i] + '   №   ' + MNums[i] + '   (партия № ' + Series[i] + ')';
    ListBox1.Items.Add(S);
  end;
  if ListBox1.Items.Count>0 then
    ListBox1.ItemIndex:= ListBox1.Items.Count-1;
  DelButton.Enabled:= ListBox1.Items.Count>0;
end;

end.

