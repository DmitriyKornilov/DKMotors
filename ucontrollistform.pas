unit UControlListForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  EditBtn, StdCtrls, VirtualTrees, DK_VSTTables, DK_SheetExporter,
  UCardForm, DK_Vector, USheetUtils, DK_StrUtils, USQLite, UControlListEditForm,
  BCButton, DK_Dialogs, DK_Const, UUtils;

type

  { TControlListForm }

  TControlListForm = class(TForm)
    AddButton: TSpeedButton;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    DelButton: TSpeedButton;
    EditButton: TSpeedButton;
    Label2: TLabel;
    MotorCardCheckBox: TCheckBox;
    MotorNumEdit: TEditButton;
    Panel3: TPanel;
    CardPanel: TPanel;
    Panel7: TPanel;
    ExportButton: TBCButton;
    Splitter2: TSplitter;
    ToolPanel: TPanel;
    VT1: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorCardCheckBoxChange(Sender: TObject);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure VT1DblClick(Sender: TObject);
  private
    CardForm: TCardForm;
    VSTMotorsTable: TVSTTable;

    MotorIDs: TIntVector;
    MotorNames, MotorNums, Series, Notes: TStrVector;

    procedure CreateMotorsTable;
    procedure SelectMotor;
    procedure OpenControlListEditForm(const AEditType: Byte);
  public
    procedure ViewUpdate;
  end;

var
  ControlListForm: TControlListForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TControlListForm }

procedure TControlListForm.FormCreate(Sender: TObject);
begin
  SetToolPanels([
    ToolPanel
  ]);
  SetToolButtons([
    AddButton, DelButton, EditButton
  ]);

  MainForm.SetNamesPanelsVisible(True, False);
  CardForm:= CreateCardForm(ControlListForm, CardPanel);
  CreateMotorsTable;
  MotorCardCheckBox.Checked:= False;
end;

procedure TControlListForm.DelButtonClick(Sender: TObject);
begin
  if not VSTMotorsTable.IsSelected then Exit;
  if not Confirm('Удалить данные о постановке на контроль?') then Exit;
  SQLite.ControlDelete(MotorIDs[VSTMotorsTable.SelectedIndex]);
  ViewUpdate;
end;

procedure TControlListForm.EditButtonClick(Sender: TObject);
begin
  OpenControlListEditForm(2);
end;

procedure TControlListForm.AddButtonClick(Sender: TObject);
begin
  OpenControlListEditForm(1);
end;

procedure TControlListForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTMotorsTable) then FreeAndNil(VSTMotorsTable);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TControlListForm.FormShow(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TControlListForm.MotorCardCheckBoxChange(Sender: TObject);
begin
  If MotorCardCheckBox.Checked then
  begin
    VT1.Align:= alCustom;
    Splitter2.Visible:= True;
    Splitter2.Align:= alTop;
    CardPanel.Visible:= True;
    Splitter2.Align:= alBottom;
    VT1.Align:= alClient;
  end
  else begin
    CardPanel.Visible:= False;
    Splitter2.Visible:= False;
  end;
end;

procedure TControlListForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TControlListForm.MotorNumEditChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TControlListForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TSheetsExporter;
  Sheet: TsWorksheet;
  ControlSheet: TControlSheet;
begin
  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    ControlSheet:= TControlSheet.Create(Sheet);
    try
      ControlSheet.Draw(MotorNames, MotorNums, Series, Notes);
    finally
      FreeAndNil(ControlSheet);
    end;
    Exporter.PageSettings(spoPortrait);

    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TControlListForm.VT1DblClick(Sender: TObject);
begin
  if not VSTMotorsTable.IsSelected then Exit;
  OpenControlListEditForm(2);
end;

procedure TControlListForm.CreateMotorsTable;
begin
  VSTMotorsTable:= TVSTTable.Create(VT1);
  VSTMotorsTable.OnSelect:= @SelectMotor;
  VSTMotorsTable.HeaderFont.Style:= [fsBold];
  VSTMotorsTable.AddColumn('№ п/п', 60);
  VSTMotorsTable.AddColumn('Наименование', 200);
  VSTMotorsTable.AddColumn('Номер', 100);
  VSTMotorsTable.AddColumn('Партия', 100);
  VSTMotorsTable.AddColumn('Примечание', 400);
  VSTMotorsTable.CanSelect:= True;
  VSTMotorsTable.Draw;
end;

procedure TControlListForm.SelectMotor;
var
  MotorID: Integer;
begin
  DelButton.Enabled:= VSTMotorsTable.IsSelected;
  EditButton.Enabled:= VSTMotorsTable.IsSelected;

  MotorID:= 0;
  if VSTMotorsTable.IsSelected then
    MotorID:= MotorIDs[VSTMotorsTable.SelectedIndex];
  CardForm.ShowCard(MotorID);
end;

procedure TControlListForm.OpenControlListEditForm(const AEditType: Byte);
var
  ControlListEditForm: TControlListEditForm;
begin
  ControlListEditForm:= TControlListEditForm.Create(ControlListForm);
  try
    ControlListEditForm.MotorID:= 0;
    if AEditType=2 then
    begin
      ControlListEditForm.MotorID:= MotorIDs[VSTMotorsTable.SelectedIndex];
      ControlListEditForm.MotorName:= MotorNames[VSTMotorsTable.SelectedIndex];
      ControlListEditForm.MotorNum:= MotorNums[VSTMotorsTable.SelectedIndex];
    end;
    if ControlListEditForm.ShowModal = mrOK then ViewUpdate;
  finally
    FreeAndNil(ControlListEditForm);
  end;
end;

procedure TControlListForm.ViewUpdate;
var
  MotorNumberLike: String;
begin
  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);
    MotorNumberLike:= STrim(MotorNumEdit.Text);

    SQLite.ControlListLoad(MainForm.UsedNameIDs, MotorNumberLike,
                           MotorIDs, MotorNames, MotorNums, Series, Notes);

    VSTMotorsTable.ValuesClear;
    VSTMotorsTable.SetColumn('№ п/п', VIntToStr(VOrder(Length(MotorIDs))));
    VSTMotorsTable.SetColumn('Наименование', MotorNames, taLeftJustify);
    VSTMotorsTable.SetColumn('Номер', MotorNums);
    VSTMotorsTable.SetColumn('Партия', Series);
    VSTMotorsTable.SetColumn('Примечание', Notes, taLeftJustify);
    VSTMotorsTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;
end;

end.

