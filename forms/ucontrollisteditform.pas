unit UControlListEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, VirtualTrees, DividerBevel,
  //DK packages utils
  DK_VSTTables, DK_Vector, DK_Dialogs, DK_StrUtils, DK_Const,
  //Project utils
  UDataBase, USheetUtils;

type

  { TControlListEditForm }

  TControlListEditForm = class(TForm)
    ButtonPanel: TPanel;
    CancelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel2: TPanel;
    NoteMemo: TMemo;
    RepairPanel: TPanel;
    SaveButton: TSpeedButton;
    VT1: TVirtualStringTree;
    procedure CancelButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
  private
    CanFormClose: Boolean;
    VSTTable: TVSTTable;

    NameIDs, MotorIDs: TIntVector;
    MotorNums, Series: TStrVector;

    procedure LoadMotors;
    procedure LoadControl;
  public
    MotorID: Integer;
    MotorName, MotorNum: String;
  end;

var
  ControlListEditForm: TControlListEditForm;

implementation

{$R *.lfm}

{ TControlListEditForm }

procedure TControlListEditForm.CancelButtonClick(Sender: TObject);
begin
  CanFormClose:= True;
  ModalResult:= mrCancel;
end;

procedure TControlListEditForm.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  CanClose:= CanFormClose;
end;

procedure TControlListEditForm.FormCreate(Sender: TObject);
begin
  DataBase.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);
  if VIsNil(NameIDs) then
    Inform('Отсутствует список наименований двигателей!');

  VSTTable:= TVSTTable.Create(VT1);

  CanFormClose:= True;
end;

procedure TControlListEditForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTTable);
end;

procedure TControlListEditForm.FormShow(Sender: TObject);
begin
  VSTTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTable.AddColumn('Номер', 200);
  VSTTable.AddColumn('Партия', 200);
  VSTTable.AutosizeColumnDisable;
  VSTTable.CanSelect:= True;
  VSTTable.Draw;

  if MotorID>0 then LoadControl;
end;

procedure TControlListEditForm.MotorNumEditChange(Sender: TObject);
begin
  LoadMotors;
end;

procedure TControlListEditForm.SaveButtonClick(Sender: TObject);
begin
  CanFormClose:= False;

  if not VSTTable.IsSelected then
  begin
    Inform('Не указан электродвигатель!');
    Exit;
  end;

  if MotorID=0 then
    MotorID:= MotorIDs[VSTTable.SelectedIndex];

  CanFormClose:= DataBase.ControlUpdate(MotorID, STrim(NoteMemo.Text));
  ModalResult:= mrOK;
end;

procedure TControlListEditForm.LoadMotors;
var
  MotorNumberLike: String;
begin
  if VIsNil(NameIDs) then Exit;

  MotorNumberLike:= STrim(MotorNumEdit.Text);

  DataBase.MotorListToControl(NameIDs[MotorNameComboBox.ItemIndex], MotorNumberLike,
                            MotorIDs, MotorNums, Series);

  VSTTable.ValuesClear;
  VSTTable.SetColumn('Номер', MotorNums);
  VSTTable.SetColumn('Партия', Series);
  VSTTable.Draw;
end;

procedure TControlListEditForm.LoadControl;
begin
  MotorNameComboBox.ItemIndex:= MotorNameComboBox.Items.IndexOf(MotorName);
  MotorNameComboBox.Enabled:= False;
  MotorNumEdit.Text:= MotorNum;
  MotorNumEdit.Enabled:= False;

  DataBase.MotorListOnControl(MotorID, MotorIDs, MotorNums, Series);
  VSTTable.ValuesClear;
  VSTTable.SetColumn('Номер', MotorNums);
  VSTTable.SetColumn('Партия', Series);
  VSTTable.Draw;
  VSTTable.Select(0);
  VSTTable.CanUnselect:= False;

  NoteMemo.Text:= DataBase.ControlNoteLoad(MotorID);
end;

end.

