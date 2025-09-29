unit UBuildEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs,  ExtCtrls,
  StdCtrls, DateTimePicker, Buttons, DividerBevel,
  //DK packages utils
  DK_Vector, DK_Dialogs, DK_StrUtils, DK_CtrlUtils,
  //Project utils
  UVars;

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
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
  private
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
end;

procedure TBuildEditForm.CancelButtonClick(Sender: TObject);
begin
  ModalResult:= mrCancel;
end;

procedure TBuildEditForm.FormShow(Sender: TObject);
begin
  Images.ToButtons([SaveButton, CancelButton]);
  SetEventButtons([SaveButton, CancelButton]);

  MotorNameComboBox.ItemIndex:= VIndexOf(NameIDs, OldNameID);
end;

procedure TBuildEditForm.SaveButtonClick(Sender: TObject);
var
  NameID: Integer;
  MotorNum, RotorNum: String;
begin
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

  if not DataBase.MotorInBuildLogUpdate(MotorID, DateTimePicker1.Date, NameID,
                   Ord(OldMotorCheckbox.Checked), MotorNum, RotorNum) then Exit;

  ModalResult:= mrOK;
end;

procedure TBuildEditForm.LoadMotorNames;
begin
  DataBase.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);

  if VIsNil(NameIDs) then
    Inform('Отсутствует список наименований двигателей!');
end;

end.

