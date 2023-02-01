unit UReclamationEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, DateTimePicker, VirtualTrees, USQLite, DK_Vector, DK_Dialogs,
  DK_VSTTables, DK_StrUtils, USheetUtils, Buttons, DividerBevel;

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
  VSTTable:= TVSTTable.Create(VT1);
  LoadNames;
  CanFormClose:= True;
end;

procedure TReclamationEditForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTTable);
end;

procedure TReclamationEditForm.FormShow(Sender: TObject);
begin
  VSTTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  //VSTTable.SelectedFont.Style:= [fsBold];
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

procedure TReclamationEditForm.SaveButtonClick(Sender: TObject);
var
  S: String;
  Mileage: Integer;
  ArrivalDate, SendingDate: TDate;
begin
  CanFormClose:= False;

  if not VSTTable.IsSelected then
  begin
    ShowInfo('Не указан рекламационный электродвигатель');
    Exit;
  end;

  if PlaceNameComboBox.ItemIndex=0 then
  begin
    ShowInfo('Не указано предприятие (депо)!');
    Exit;
  end;

  S:= STrim(MileageEdit.Text);
  if S=EmptyStr then
    Mileage:= -1
  else if not TryStrToInt(S, Mileage) then
  begin
    ShowInfo('Некорректное значение величины пробега!');
    Exit;
  end;

  if RecID=0 then
    CanFormClose:= SQLite.ReclamationWrite(DateTimePicker1.Date,
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
    CanFormClose:= SQLite.ReclamationUpdate(DateTimePicker1.Date,
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

  ModalResult:= mrOK;
end;

procedure TReclamationEditForm.SendingCheckBoxChange(Sender: TObject);
begin
  DateTimePicker3.Enabled:= SendingCheckBox.Checked;
end;

procedure TReclamationEditForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
  ModalResult:= mrCancel;
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

procedure TReclamationEditForm.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TReclamationEditForm.LoadNames;
begin
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);
  if VIsNil(NameIDs) then
    ShowInfo('Отсутствует список наименований двигателей!');

  SQLite.KeyPickList('RECLAMATIONPLACES', 'PlaceID', 'PlaceName',
                 PlaceIDs, PlaceNames, False, 'PlaceName');
  VToStrings(PlaceNames, PlaceNameComboBox.Items);
  PlaceNameComboBox.ItemIndex:= 0;

  SQLite.KeyPickList('RECLAMATIONFACTORIES', 'FactoryID', 'FactoryName',
                 FactoryIDs, FactoryNames, False, 'FactoryName');
  VToStrings(FactoryNames, FactoryNameComboBox.Items);
  FactoryNameComboBox.ItemIndex:= 0;

  SQLite.KeyPickList('RECLAMATIONDEFECTS', 'DefectID', 'DefectName',
                 DefectIDs, DefectNames, False, 'DefectName');
  VToStrings(DefectNames, DefectNameComboBox.Items);
  DefectNameComboBox.ItemIndex:= 0;

  SQLite.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
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

  SQLite.ReclamationChooseListLoad(NameIDs[MotorNameComboBox.ItemIndex],
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
  if not SQLite.ReclamationLoad(RecID, RecDate, ArrivalDate, SendingDate,
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

