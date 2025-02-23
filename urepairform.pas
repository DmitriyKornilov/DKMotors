unit URepairForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, EditBtn, VirtualTrees, UUtils, DK_VSTTables, USheetUtils,
  DK_Vector, USQLite, DK_StrUtils, DK_Dialogs, DK_Const,
  DK_SheetExporter, UCalendar, UCardForm, URepairEditForm, BCButton,
  DK_VSTTableTools;

type

  { TRepairForm }

  TRepairForm = class(TForm)
    AddButton: TSpeedButton;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    DelButton: TSpeedButton;
    EditButton: TSpeedButton;
    ExportButton: TBCButton;
    Label2: TLabel;
    MotorCardCheckBox: TCheckBox;
    MotorNumEdit: TEditButton;
    Panel1: TPanel;
    CardPanel: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel5: TPanel;
    LeftPanel: TPanel;
    Panel7: TPanel;
    MainPanel: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    ToolPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    VT3: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorCardCheckBoxChange(Sender: TObject);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure VT1DblClick(Sender: TObject);
  private
    CardForm: TCardForm;
    VSTMotorsTable: TVSTTable;
    VSTTypeList: TVSTStringList;
    VSTOrderList: TVSTStringList;
    RecIDs, MotorIDs: TIntVector;

    ArrivalDates, SendingDates: TDateVector;
    Passports, DayCounts: TIntVector;
    MotorNames, MotorNums, RepairNotes: TStrVector;

    procedure CreateMotorsTable;
    procedure SelectMotor;

    procedure CreateTypeList;
    procedure CreateOrderList;

    procedure OpenRepairEditForm(const AEditType: Byte);
  public
    procedure ViewUpdate;
  end;

var
  RepairForm: TRepairForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TRepairForm }

procedure TRepairForm.FormCreate(Sender: TObject);
begin
  SetToolPanels([
    ToolPanel
  ]);
  SetToolButtons([
    AddButton, DelButton, EditButton
  ]);

  MainForm.SetNamesPanelsVisible(True, False);
  CardForm:= CreateCardForm(RepairForm, CardPanel);
  CreateMotorsTable;
  CreateTypeList;
  CreateOrderList;
  MotorCardCheckBox.Checked:= False;
end;

procedure TRepairForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TSheetsExporter;
  Sheet: TsWorksheet;
  RepairSheet: TRepairSheet;
begin
  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    RepairSheet:= TRepairSheet.Create(Sheet);
    try
      RepairSheet.Draw(Passports, DayCounts, MotorNames, MotorNums, ArrivalDates, SendingDates);
    finally
      FreeAndNil(RepairSheet);
    end;
    Exporter.PageSettings(spoPortrait);

    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TRepairForm.AddButtonClick(Sender: TObject);
begin
  OpenRepairEditForm(1);
end;

procedure TRepairForm.DelButtonClick(Sender: TObject);
begin
  if not VSTMotorsTable.IsSelected then Exit;
  if not Confirm('Удалить данные о гарантийном ремонте?') then Exit;
  SQLite.RepairDelete(RecIDs[VSTMotorsTable.SelectedIndex]);
  ViewUpdate;
end;

procedure TRepairForm.EditButtonClick(Sender: TObject);
begin
  OpenRepairEditForm(2);
end;

procedure TRepairForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTMotorsTable) then FreeAndNil(VSTMotorsTable);
  if Assigned(VSTTypeList) then FreeAndNil(VSTTypeList);
  if Assigned(VSTOrderList) then FreeAndNil(VSTOrderList);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TRepairForm.FormShow(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TRepairForm.MotorCardCheckBoxChange(Sender: TObject);
begin
  if MotorCardCheckBox.Checked then
  begin
    Splitter2.Visible:= True;
    Splitter2.Align:= alTop;
    CardPanel.Visible:= True;
    Splitter2.Align:= alBottom;
  end
  else begin
    CardPanel.Visible:= False;
    Splitter2.Visible:= False;
  end;
end;

procedure TRepairForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TRepairForm.MotorNumEditChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TRepairForm.VT1DblClick(Sender: TObject);
begin
  if not VSTMotorsTable.IsSelected then Exit;
  OpenRepairEditForm(2);
end;

procedure TRepairForm.CreateMotorsTable;
begin
  VSTMotorsTable:= TVSTTable.Create(VT1);
  VSTMotorsTable.OnSelect:= @SelectMotor;
  VSTMotorsTable.HeaderFont.Style:= [fsBold];
  VSTMotorsTable.AddColumn('№ п/п', 50);
  VSTMotorsTable.AddColumn('Наименование', 200);
  VSTMotorsTable.AddColumn('Номер', 100);
  VSTMotorsTable.AddColumn('Наличие паспорта', 130);
  VSTMotorsTable.AddColumn('Дата прибытия', 100);
  VSTMotorsTable.AddColumn('Дата убытия', 100);
  VSTMotorsTable.AddColumn('Срок ремонта (рабочих дней)', 200);
  VSTMotorsTable.AddColumn('Примечание');
  VSTMotorsTable.CanSelect:= True;
  VSTMotorsTable.Draw;
end;

procedure TRepairForm.SelectMotor;
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

procedure TRepairForm.CreateTypeList;
var
  S: String;
  V: TStrVector;
begin
  S:= 'Отображать:';
  V:= VCreateStr([
    'все',
    'в ремонте',
    'отремонтированные'
  ]);
  VSTTypeList:= TVSTStringList.Create(VT2, S, @ViewUpdate);
  VSTTypeList.Update(V, 1);
end;

procedure TRepairForm.CreateOrderList;
var
  S: String;
  V: TStrVector;
begin
  S:= 'Упорядочить по:';
  V:= VCreateStr([
    'дате прибытия',
    'дате убытия',
    'номеру'
  ]);
  VSTOrderList:= TVSTStringList.Create(VT3, S, @ViewUpdate);
  VSTOrderList.Update(V);
end;

procedure TRepairForm.OpenRepairEditForm(const AEditType: Byte);
var
  RepairEditForm: TRepairEditForm;
begin
  RepairEditForm:= TRepairEditForm.Create(RepairForm);
  try
    RepairEditForm.RecID:= 0;
    if AEditType=2 then
    begin
      RepairEditForm.RecID:= RecIDs[VSTMotorsTable.SelectedIndex];
      RepairEditForm.MotorName:= MotorNames[VSTMotorsTable.SelectedIndex];
      RepairEditForm.MotorNum:= MotorNums[VSTMotorsTable.SelectedIndex];
    end;
    if RepairEditForm.ShowModal = mrOK then ViewUpdate;
  finally
    FreeAndNil(RepairEditForm);
  end;
end;

procedure TRepairForm.ViewUpdate;
var
  MotorNumberLike: String;
  i: Integer;
  PassStrs, DayCountsStrs, SendingDatesStrs, ArrivalDatesStrs: TStrVector;
begin
  if (not VSTTypeList.IsSelected) or (not VSTOrderList.IsSelected) then Exit;

  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);
    MotorNumberLike:= STrim(MotorNumEdit.Text);

    SQLite.RepairListLoad(MainForm.UsedNameIDs, MotorNumberLike,
                          VSTOrderList.SelectedIndex+1, VSTTypeList.SelectedIndex,
                          ArrivalDates, SendingDates,
                          RecIDs, MotorIDs, Passports, DayCounts,
                          MotorNames, MotorNums, RepairNotes);

    VDim(PassStrs{%H-}, Length(Passports));
    for i:= 0 to High(PassStrs) do
      if Passports[i]>0 then
        PassStrs[i]:= CHECK_SYMBOL;
    DayCountsStrs:= VIntToStr(DayCounts, True);
    ArrivalDatesStrs:= VFormatDateTime('dd.mm.yyyy', ArrivalDates, True);
    SendingDatesStrs:= VFormatDateTime('dd.mm.yyyy', SendingDates, True);

    VSTMotorsTable.ValuesClear;
    VSTMotorsTable.SetColumn('№ п/п', VIntToStr(VOrder(Length(MotorIDs))));
    VSTMotorsTable.SetColumn('Наименование', MotorNames, taLeftJustify);
    VSTMotorsTable.SetColumn('Номер', MotorNums);
    VSTMotorsTable.SetColumn('Наличие паспорта', PassStrs);
    VSTMotorsTable.SetColumn('Дата прибытия', ArrivalDatesStrs);
    VSTMotorsTable.SetColumn('Дата убытия', SendingDatesStrs);
    VSTMotorsTable.SetColumn('Срок ремонта (рабочих дней)', DayCountsStrs);
    VSTMotorsTable.SetColumn('Примечание', RepairNotes, taLeftJustify);
    VSTMotorsTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;
end;

end.

