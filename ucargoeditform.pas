unit UCargoEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, DateTimePicker, DK_Vector, DK_Dialogs, DK_VSTTables,
  DK_StrUtils, LCLType, VirtualTrees, DK_Const, DividerBevel, USQLite,
  USheetUtils;

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
    MotorNameComboBox: TComboBox;
    ReceiverNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel2: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);


    procedure DateTimePicker1Change(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure MotorNumEditKeyDown(Sender: TObject; var Key: Word;
      {%H-}Shift: TShiftState);
    procedure ReceiverNameComboBoxChange(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure SeriesEditKeyDown(Sender: TObject; var Key: Word;
      {%H-}Shift: TShiftState);
    procedure VT1MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure VT2Exit(Sender: TObject);
    procedure VT2KeyUp(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure VT2MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
  private
    NameIDs: TIntVector;
    ReceiverIDs: TIntVector;

    ViewMotorIDs: TIntVector;
    ViewMotorNums, ViewBuildDates, ViewSeries: TStrVector;

    MotorIDs: TIntVector;
    Series, MotorNames, MotorNums: TStrVector;

    WritedMotorIDs: TIntVector;

    CanFormClose: Boolean;

    VSTViewTable, VSTCargoTable: TVSTTable;

    procedure LoadMotorNames;
    procedure LoadReceiverNames;

    procedure AddMotor;
    procedure DelMotor;

    procedure LoadMotors;
    procedure LoadCargo;

    procedure ShowCargoList(const ANeedSelect: Boolean);
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
  VSTViewTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTViewTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTViewTable.AddColumn('Дата сборки', 100);
  VSTViewTable.AddColumn('Номер', 100);
  VSTViewTable.AddColumn('Партия',50);
  VSTViewTable.CanSelect:= True;
  VSTViewTable.Draw;

  VSTCargoTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTCargoTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTCargoTable.AddColumn('№ п/п', 60);
  VSTCargoTable.AddColumn('Наименование', 220);
  VSTCargoTable.AddColumn('Номер', 100);
  VSTCargoTable.AddColumn('Партия');
  VSTCargoTable.CanSelect:= True;
  VSTCargoTable.Draw;

  if CargoID>0 then LoadCargo;

  DateTimePicker1.SetFocus;
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
    VSTViewTable.Select(0);
    AddButton.Enabled:= True;
    SeriesEdit.Text:= ViewSeries[VSTViewTable.SelectedIndex];
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
  if not VSTViewTable.IsSelected then Exit;
  if Key=VK_RETURN then AddMotor;
end;

procedure TCargoEditForm.VT1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  AddButton.Enabled:= VSTViewTable.IsSelected;
  if VSTViewTable.IsSelected then
    SeriesEdit.Text:= ViewSeries[VSTViewTable.SelectedIndex]
  else
    SeriesEdit.Text:= EmptyStr;
end;

procedure TCargoEditForm.VT2Exit(Sender: TObject);
begin
  VSTCargoTable.UnSelect;
  DelButton.Enabled:= False;
end;

procedure TCargoEditForm.VT2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE then
    DelMotor;
end;

procedure TCargoEditForm.VT2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  DelButton.Enabled:= VSTCargoTable.IsSelected;
end;

procedure TCargoEditForm.FormCreate(Sender: TObject);
begin
  VSTViewTable:= TVSTTable.Create(VT1);
  VSTCargoTable:= TVSTTable.Create(VT2);
  DateTimePicker1.Date:= Date;
  LoadMotorNames;
  LoadReceiverNames;
  CanFormClose:= True;
end;

procedure TCargoEditForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTViewTable);
  FreeAndNil(VSTCargoTable);
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
  MotorSeries: String;
  MotorID: Integer;
begin
  if not VSTViewTable.IsSelected then Exit;

  MotorID:= ViewMotorIDs[VSTViewTable.SelectedIndex];
  VAppend(MotorIDs, MotorID);

  MotorSeries:= STrim(SeriesEdit.Text);
  if MotorSeries=EmptyStr then
  begin
    ShowInfo('Не указан номер партии!');
    Exit;
  end;

  VAppend(Series, MotorSeries);
  VAppend(MotorNames, MotorNameComboBox.Text);
  VAppend(MotorNums, ViewMotorNums[VSTViewTable.SelectedIndex]);

  ShowCargoList(False);

  MotorNumEdit.Text:= EmptyStr;
  AddButton.Enabled:= False;
  MotorNumEdit.SetFocus;
end;

procedure TCargoEditForm.DelMotor;
var
  Ind: Integer;
begin
  if not VSTCargoTable.IsSelected then Exit;
  Ind:= VSTCargoTable.SelectedIndex;

  VDel(MotorIDs, Ind);
  VDel(Series, Ind);
  VDel(MotorNames, Ind);
  VDel(MotorNums, Ind);
  ShowCargoList(True);
end;

procedure TCargoEditForm.ShowCargoList(const ANeedSelect: Boolean);
var
  Ind, LastInd: Integer;
begin
  Ind:= -1;
  if VSTCargoTable.IsSelected then
    Ind:= VSTCargoTable.SelectedIndex;

  VSTCargoTable.SetColumn('№ п/п', VIntToStr(VOrder(Length(MotorNames))));
  VSTCargoTable.SetColumn('Наименование', MotorNames, taLeftJustify);
  VSTCargoTable.SetColumn('Номер', MotorNums);
  VSTCargoTable.SetColumn('Партия', Series);
  VSTCargoTable.Draw;

  if not VIsNil(MotorNames) then
  begin
    LastInd:= High(MotorNames);
    if (Ind<0) or (Ind>LastInd) then
      Ind:= LastInd;

    if ANeedSelect then
      VSTCargoTable.Select(Ind)
    else
      VSTCargoTable.Show(LastInd);
  end;

  DelButton.Enabled:= VSTCargoTable.IsSelected;
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

  VSTViewTable.ValuesClear;
  VSTViewTable.SetColumn('Дата сборки', ViewBuildDates);
  VSTViewTable.SetColumn('Номер', ViewMotorNums);
  VSTViewTable.SetColumn('Партия', ViewSeries{, taLeftJustify});
  VSTViewTable.Draw;
end;

procedure TCargoEditForm.LoadCargo;
var
  i: Integer;
  D: TDate;
begin
  if not SQLite.CargoMotorListLoad(CargoID,i, D, MotorIDs,
                                   MotorNames, MotorNums, Series) then Exit;

  DateTimePicker1.Date:= D;
  ReceiverNameComboBox.ItemIndex:= VIndexOf(ReceiverIDs, i);

  WritedMotorIDs:= VCut(MotorIDs);

  ShowCargoList(False);
end;



end.

