unit UControlListEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, VirtualTrees, DividerBevel,
  //DK packages utils
  DK_VSTTables, DK_Vector, DK_Dialogs, DK_StrUtils, DK_Const, DK_CtrlUtils,
  DK_Filter,
  //Project utils
  UVars, USheets;

type

  { TControlListEditForm }

  TControlListEditForm = class(TForm)
    ButtonPanel: TPanel;
    CancelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    FilterPanel: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MotorNameComboBox: TComboBox;
    Panel2: TPanel;
    NoteMemo: TMemo;
    RepairPanel: TPanel;
    SaveButton: TSpeedButton;
    VT1: TVirtualStringTree;
    procedure CancelButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
  private
    FilterString: String;
    Filter: TDKFilter;

    VSTTable: TVSTTable;

    NameIDs, MotorIDs: TIntVector;
    MotorNums, Series: TStrVector;

    procedure FilterMotors(const AFilterString: String);
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

procedure TControlListEditForm.FormCreate(Sender: TObject);
begin
  DataBase.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs);
  if VIsNil(NameIDs) then
    Inform('Отсутствует список наименований двигателей!');

  VSTTable:= TVSTTable.Create(VT1);

  FilterString:= EmptyStr;
  Filter:= DKFilterCreate('Номер двигателя', FilterPanel, @FilterMotors, -1, 300);
end;

procedure TControlListEditForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(VSTTable);
end;

procedure TControlListEditForm.FormShow(Sender: TObject);
begin
  Images.ToButtons([SaveButton, CancelButton]);
  SetEventButtons([SaveButton, CancelButton]);

  VSTTable.HeaderBGColor:= COLOR_BACKGROUND_TITLE;
  VSTTable.AddColumn('Номер', 200);
  VSTTable.AddColumn('Партия', 200);
  VSTTable.AutosizeColumnDisable;
  VSTTable.CanSelect:= True;
  VSTTable.Draw;

  if MotorID>0 then LoadControl;
end;

procedure TControlListEditForm.CancelButtonClick(Sender: TObject);
begin
  ModalResult:= mrCancel;
end;

procedure TControlListEditForm.SaveButtonClick(Sender: TObject);
begin
  if not VSTTable.IsSelected then
  begin
    Inform('Не указан электродвигатель!');
    Exit;
  end;

  if MotorID=0 then
    MotorID:= MotorIDs[VSTTable.SelectedIndex];

  if not DataBase.ControlUpdate(MotorID, STrim(NoteMemo.Text)) then Exit;
  ModalResult:= mrOK;
end;

procedure TControlListEditForm.FilterMotors(const AFilterString: String);
begin
  FilterString:= AFilterString;
  LoadMotors;
end;

procedure TControlListEditForm.LoadMotors;
begin
  if not SEmpty(MotorNum) then Exit;
  if VIsNil(NameIDs) then Exit;

  DataBase.MotorListToControl(NameIDs[MotorNameComboBox.ItemIndex], STrim(FilterString),
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

  Filter.FilterString:= MotorNum;
  Filter.FilterEnabled:= False;

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

