unit URepairEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, DateTimePicker, VirtualTrees, DividerBevel, DateUtils,
  //DK packages utils
  DK_VSTTables, DK_Vector, DK_StrUtils, DK_Const, DK_Dialogs,
  //Project utils
  UDataBase, USheetUtils;

type

  { TRepairEditForm }

  TRepairEditForm = class(TForm)
    ButtonPanel: TPanel;
    CancelButton: TSpeedButton;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel2: TPanel;
    PassportCheckBox: TCheckBox;
    RecNoteMemo: TMemo;
    RepairPanel: TPanel;
    SaveButton: TSpeedButton;
    SendingCheckBox: TCheckBox;
    VT1: TVirtualStringTree;
    procedure CancelButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure SendingCheckBoxChange(Sender: TObject);
  private
    CanFormClose: Boolean;
    VSTTable: TVSTTable;

    NameIDs, RecIDs, MotorNameIDs: TIntVector;
    MotorNums, PlaceNames: TStrVector;
    BuildDates, RecDates: TDateVector;

    procedure LoadMotors;
    procedure LoadRepair;
  public
    RecID: Integer;
    MotorName, MotorNum: String;
  end;

var
  RepairEditForm: TRepairEditForm;

implementation

{$R *.lfm}

{ TRepairEditForm }

procedure TRepairEditForm.SendingCheckBoxChange(Sender: TObject);
begin
  DateTimePicker2.Enabled:= SendingCheckBox.Checked;
end;

procedure TRepairEditForm.LoadMotors;
var
  MotorNumberLike: String;
begin
  if VIsNil(NameIDs) then Exit;

  MotorNumberLike:= STrim(MotorNumEdit.Text);

  DataBase.MotorListToRepair(NameIDs[MotorNameComboBox.ItemIndex],
          MotorNumberLike, RecIDs, MotorNameIDs, BuildDates, RecDates, MotorNums, PlaceNames);

  VSTTable.ValuesClear;
  VSTTable.SetColumn('Дата сборки', VFormatDateTime('dd.mm.yyyy', BuildDates));
  VSTTable.SetColumn('Номер', MotorNums);
  VSTTable.SetColumn('Дата уведомления', VFormatDateTime('dd.mm.yyyy', RecDates));
  VSTTable.SetColumn('Предприятие', PlaceNames, taLeftJustify);
  VSTTable.Draw;
end;

procedure TRepairEditForm.LoadRepair;
var
  ArrivalDate, SendingDate: TDate;
  Passport: Integer;
  RepairNote: String;
begin
  MotorNameComboBox.ItemIndex:= MotorNameComboBox.Items.IndexOf(MotorName);
  MotorNameComboBox.Enabled:= False;
  MotorNumEdit.Text:= MotorNum;
  MotorNumEdit.Enabled:= False;

  DataBase.MotorListOnRepair(RecID, RecIDs, MotorNameIDs, BuildDates,
                           RecDates, MotorNums, PlaceNames);

  VSTTable.ValuesClear;
  VSTTable.SetColumn('Дата сборки', VFormatDateTime('dd.mm.yyyy', BuildDates));
  VSTTable.SetColumn('Номер', MotorNums);
  VSTTable.SetColumn('Дата уведомления', VFormatDateTime('dd.mm.yyyy', RecDates));
  VSTTable.SetColumn('Предприятие', PlaceNames, taLeftJustify);
  VSTTable.Draw;
  VSTTable.Select(0);
  VSTTable.CanUnselect:= False;

  if DataBase.RepairLoad(RecID, ArrivalDate, SendingDate, Passport, RepairNote) then
  begin
    if ArrivalDate>0 then
      DateTimePicker1.Date:= ArrivalDate
    else
      DateTimePicker1.Date:= Date;
    PassportCheckBox.Checked:= Passport>0;
    SendingCheckBox.Checked:= SendingDate>0;
    if SendingDate>0 then
      DateTimePicker2.Date:= SendingDate;
    RecNoteMemo.Text:= RepairNote;
  end;
end;

procedure TRepairEditForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
  ModalResult:= mrCancel;
end;

procedure TRepairEditForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TRepairEditForm.FormCreate(Sender: TObject);
begin
  DataBase.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);
  if VIsNil(NameIDs) then
    Inform('Отсутствует список наименований двигателей!');

  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= Date;
  VSTTable:= TVSTTable.Create(VT1);

  CanFormClose:= True;
end;

procedure TRepairEditForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTTable);
end;

procedure TRepairEditForm.FormShow(Sender: TObject);
begin
  VSTTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTable.AddColumn('Дата сборки', 120);
  VSTTable.AddColumn('Номер', 100);
  VSTTable.AddColumn('Дата уведомления', 160);
  VSTTable.AddColumn('Предприятие', 100);
  VSTTable.CanSelect:= True;
  VSTTable.Draw;

  if RecID>0 then LoadRepair;
end;

procedure TRepairEditForm.MotorNumEditChange(Sender: TObject);
begin
  LoadMotors;
end;

procedure TRepairEditForm.SaveButtonClick(Sender: TObject);
var
  ArrivalDate, SendingDate: TDate;
begin
  CanFormClose:= False;

  if not VSTTable.IsSelected then
  begin
    Inform('Не указан рекламационный электродвигатель!');
    Exit;
  end;

  ArrivalDate:= DateTimePicker1.Date;
  SendingDate:= 0;
  if SendingCheckBox.Checked then
  begin
    SendingDate:= DateTimePicker2.Date;
    if CompareDate(ArrivalDate, SendingDate)>0 then
    begin
      Inform('Дата убытия не может быть меньше даты прибытия!');
      Exit;
    end;
  end;
  if RecID=0 then
    RecID:= RecIDs[VSTTable.SelectedIndex];

  CanFormClose:= DataBase.RepairUpdate(RecID, ArrivalDate, SendingDate,
                                     Ord(PassportCheckBox.Checked),
                                     STrim(RecNoteMemo.Text));
  ModalResult:= mrOK;
end;

end.

