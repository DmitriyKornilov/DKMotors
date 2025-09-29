unit UReclamationEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons, DividerBevel,
  StdCtrls, DateTimePicker, VirtualTrees,
  //DK packages utils
  DK_Vector, DK_Dialogs, DK_VSTTables, DK_StrUtils, DK_Const, DK_CtrlUtils,
  //Project utils
  UVars, USheets;

type

  { TReclamationEditForm }

  TReclamationEditForm = class(TForm)
    ArrivalCheckBox: TCheckBox;
    DateTimePicker2: TDateTimePicker;
    DateTimePicker3: TDateTimePicker;
    RepairPanel: TPanel;
    ButtonPanel: TPanel;
    CancelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    Label10: TLabel;
    PassportCheckBox: TCheckBox;
    RecNoteMemo: TMemo;
    OpinionCheckBox: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    Label8: TLabel;
    Label9: TLabel;
    ReasonNameComboBox: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    DepartureEdit: TEdit;
    Label7: TLabel;
    MotorNameComboBox: TComboBox;
    PlaceNameComboBox: TComboBox;
    FactoryNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    MileageEdit: TEdit;
    Panel2: TPanel;
    DefectNameComboBox: TComboBox;
    SaveButton: TSpeedButton;
    SendingCheckBox: TCheckBox;
    VT1: TVirtualStringTree;
    procedure ArrivalCheckBoxChange(Sender: TObject);
    procedure CancelButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure SendingCheckBoxChange(Sender: TObject);
  private
    VSTTable: TVSTTable;

    MotorIDs: TIntVector;
    MotorNums, BuildDates, Shippings: TStrVector;

    NameIDs, PlaceIDs, FactoryIDs, DefectIDs, ReasonIDs: TIntVector;
    PlaceNames, FactoryNames, DefectNames, ReasonNames: TStrVector;

    procedure LoadNames;
    procedure LoadMotors;
    procedure LoadRec;
  public
    RecID: Integer;
  end;

var
  ReclamationEditForm: TReclamationEditForm;

implementation

{$R *.lfm}

{ TReclamationEditForm }

procedure TReclamationEditForm.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= Date;
  VSTTable:= TVSTTable.Create(VT1);
  LoadNames;
end;

procedure TReclamationEditForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTTable);
end;

procedure TReclamationEditForm.FormShow(Sender: TObject);
begin
  Images.ToButtons([SaveButton, CancelButton]);
  SetEventButtons([SaveButton, CancelButton]);

  VSTTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTable.AddColumn('Дата сборки', 100);
  VSTTable.AddColumn('Номер', 100);
  VSTTable.AddColumn('Отгружен',50);
  VSTTable.CanSelect:= True;
  VSTTable.Draw;

  if RecID>0 then LoadRec;
end;

procedure TReclamationEditForm.MotorNumEditChange(Sender: TObject);
begin
  LoadMotors;
end;

procedure TReclamationEditForm.CancelButtonClick(Sender: TObject);
begin
  ModalResult:= mrCancel;
end;

procedure TReclamationEditForm.SaveButtonClick(Sender: TObject);
var
  S: String;
  Mileage: Integer;
  ArrivalDate, SendingDate: TDate;
  IsOK: Boolean;
begin
  if not VSTTable.IsSelected then
  begin
    Inform('Не указан рекламационный электродвигатель!');
    Exit;
  end;

  if PlaceNameComboBox.ItemIndex=0 then
  begin
    Inform('Не указано предприятие (депо)!');
    Exit;
  end;

  S:= STrim(MileageEdit.Text);
  if S=EmptyStr then
    Mileage:= -1
  else if not TryStrToInt(S, Mileage) then
  begin
    Inform('Некорректное значение величины пробега!');
    Exit;
  end;

  if RecID=0 then
    IsOK:= DataBase.ReclamationWrite(DateTimePicker1.Date,
                          MotorIDs[VSTTable.SelectedIndex], Mileage,
                          PlaceIDs[PlaceNameComboBox.ItemIndex],
                          FactoryIDs[FactoryNameComboBox.ItemIndex],
                          DefectIDs[DefectNameComboBox.ItemIndex],
                          ReasonIDs[ReasonNameComboBox.ItemIndex],
                          Ord(OpinionCheckbox.Checked),
                          STrim(DepartureEdit.Text),
                          STrim(RecNoteMemo.Text))
  else begin
    ArrivalDate:= 0;
    if ArrivalCheckBox.Checked then
      ArrivalDate:= DateTimePicker2.Date;
    SendingDate:= 0;
    if SendingCheckBox.Checked then
      SendingDate:= DateTimePicker3.Date;
    IsOK:= DataBase.ReclamationUpdate(DateTimePicker1.Date,
                          ArrivalDate, SendingDate, RecID,
                          MotorIDs[VSTTable.SelectedIndex], Mileage,
                          PlaceIDs[PlaceNameComboBox.ItemIndex],
                          FactoryIDs[FactoryNameComboBox.ItemIndex],
                          DefectIDs[DefectNameComboBox.ItemIndex],
                          ReasonIDs[ReasonNameComboBox.ItemIndex],
                          Ord(OpinionCheckbox.Checked),
                          Ord(PassportCheckbox.Checked),
                          STrim(DepartureEdit.Text),
                          STrim(RecNoteMemo.Text));
  end;

  if not IsOK then Exit;
  ModalResult:= mrOK;
end;

procedure TReclamationEditForm.SendingCheckBoxChange(Sender: TObject);
begin
  DateTimePicker3.Enabled:= SendingCheckBox.Checked;
end;

procedure TReclamationEditForm.ArrivalCheckBoxChange(Sender: TObject);
begin
  DateTimePicker2.Enabled:= ArrivalCheckBox.Checked;
  PassportCheckBox.Enabled:= ArrivalCheckBox.Checked;
  SendingCheckBox.Enabled:= ArrivalCheckBox.Checked;
  if not ArrivalCheckBox.Checked then
  begin
    SendingCheckBox.Checked:= False;
    PassportCheckBox.Checked:= False;
  end;
end;

procedure TReclamationEditForm.LoadNames;
begin
  DataBase.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);
  if VIsNil(NameIDs) then
    Inform('Отсутствует список наименований двигателей!');

  DataBase.KeyPickList('RECLAMATIONPLACES', 'PlaceID', 'PlaceName',
                 PlaceIDs, PlaceNames, False, 'PlaceName');
  VToStrings(PlaceNames, PlaceNameComboBox.Items);
  PlaceNameComboBox.ItemIndex:= 0;

  DataBase.KeyPickList('RECLAMATIONFACTORIES', 'FactoryID', 'FactoryName',
                 FactoryIDs, FactoryNames, False, 'FactoryName');
  VToStrings(FactoryNames, FactoryNameComboBox.Items);
  FactoryNameComboBox.ItemIndex:= 0;

  DataBase.KeyPickList('RECLAMATIONDEFECTS', 'DefectID', 'DefectName',
                 DefectIDs, DefectNames, False, 'DefectName');
  VToStrings(DefectNames, DefectNameComboBox.Items);
  DefectNameComboBox.ItemIndex:= 0;

  DataBase.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
                 ReasonIDs, ReasonNames, False, 'ReasonName');
  VToStrings(ReasonNames, ReasonNameComboBox.Items);
  ReasonNameComboBox.ItemIndex:= 0;
end;

procedure TReclamationEditForm.LoadMotors;
var
  MotorNumberLike: String;
begin
  if VIsNil(NameIDs) then Exit;

  MotorNumberLike:= STrim(MotorNumEdit.Text);

  DataBase.ReclamationChooseListLoad(NameIDs[MotorNameComboBox.ItemIndex],
       MotorNumberLike, MotorIDs, MotorNums, BuildDates, Shippings);

  VSTTable.ValuesClear;
  VSTTable.SetColumn('Дата сборки', BuildDates);
  VSTTable.SetColumn('Номер', MotorNums);
  VSTTable.SetColumn('Отгружен', Shippings, taLeftJustify);
  VSTTable.Draw;
end;

procedure TReclamationEditForm.LoadRec;
var
  RecDate, ArrivalDate, SendingDate: TDate;
  NameID, MotorID, Mileage, PlaceID, FactoryID,
  DefectID, ReasonID, Opinion, Passport: Integer;
  MotorNum, Departure, RecNote: String;
begin
  if not DataBase.ReclamationLoad(RecID, RecDate, ArrivalDate, SendingDate,
                         NameID, MotorID, Mileage, PlaceID, FactoryID,
                         DefectID, ReasonID, Opinion, Passport,
                         MotorNum, Departure, RecNote) then Exit;

  RepairPanel.Visible:= True;

  DateTimePicker1.Date:= RecDate;
  MotorNameComboBox.ItemIndex:= VIndexOf(NameIDs, NameID);
  if Mileage>=0 then MileageEdit.Text:= IntToStr(Mileage);
  PlaceNameComboBox.ItemIndex:= VIndexOf(PlaceIDs, PlaceID);
  FactoryNameComboBox.ItemIndex:= VIndexOf(FactoryIDs, FactoryID);
  DepartureEdit.Text:= Departure;
  DefectNameComboBox.ItemIndex:= VIndexOf(DefectIDs, DefectID);
  ReasonNameComboBox.ItemIndex:= VIndexOf(ReasonIDs, ReasonID);
  OpinionCheckBox.Checked:= Opinion=1;
  RecNoteMemo.Text:= RecNote;

  if ArrivalDate>0 then
  begin
    ArrivalCheckBox.Checked:= True;
    DateTimePicker2.Date:= ArrivalDate;
  end
  else
    DateTimePicker2.Date:= Date;

  if SendingDate>0 then
  begin
    SendingCheckBox.Checked:= True;
    DateTimePicker3.Date:= SendingDate;
  end
  else
    DateTimePicker3.Date:= Date;

  PassportCheckBox.Checked:= Passport=1;

  MotorNumEdit.Text:= MotorNum;
  VSTTable.Select(VIndexOf(MotorIDs, MotorID));
end;

end.

