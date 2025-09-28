unit UBuildEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs,  ExtCtrls,
  StdCtrls, DateTimePicker, Buttons, DividerBevel,
  //DK packages utils
  DK_Vector, DK_Dialogs, DK_StrUtils,
  //Project utils
  UDataBase;

type

  { TBuildEditForm }

  TBuildEditForm = class(TForm)
    ButtonPanel: TPanel;
    CancelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    OldMotorCheckBox: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel2: TPanel;
    RotorNumEdit: TEdit;
    SaveButton: TSpeedButton;
    procedure CancelButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
  private
    CanFormClose: Boolean;
    NameIDs: TIntVector;
    procedure LoadMotorNames;
  public
    MotorID, OldNameID: Integer;
  end;

var
  BuildEditForm: TBuildEditForm;

implementation

{$R *.lfm}

procedure TBuildEditForm.FormCreate(Sender: TObject);
begin
  LoadMotorNames;
  CanFormClose:= True;
end;

procedure TBuildEditForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
  ModalResult:= mrCancel;
end;

procedure TBuildEditForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TBuildEditForm.FormShow(Sender: TObject);
begin
  MotorNameComboBox.ItemIndex:= VIndexOf(NameIDs, OldNameID);
end;

procedure TBuildEditForm.SaveButtonClick(Sender: TObject);
var
  NameID: Integer;
  MotorNum, RotorNum: String;
begin
  CanFormClose:= False;

  if MotorNameComboBox.Text=EmptyStr then
  begin
    Inform('Не указано наименование двигателя!');
    Exit;
  end;
  NameID:= NameIDs[MotorNameComboBox.ItemIndex];

  MotorNum:= STrim(MotorNumEdit.Text);
  if MotorNum=EmptyStr then
  begin
    Inform('Не указан номер двигателя!');
    Exit;
  end;

  RotorNum:= STrim(RotorNumEdit.Text);

  DataBase.MotorInBuildLogUpdate(MotorID, DateTimePicker1.Date, NameID,
                   Ord(OldMotorCheckbox.Checked), MotorNum, RotorNum);

  CanFormClose:= True;
  ModalResult:= mrOK;
end;

procedure TBuildEditForm.LoadMotorNames;
begin
  DataBase.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);

  if VIsNil(NameIDs) then
    Inform('Отсутствует список наименований двигателей!');
end;

end.

