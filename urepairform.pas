unit URepairForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, EditBtn, VirtualTrees, DividerBevel, DK_VSTTables, USheetUtils,
  rxctrls, DK_Vector, USQLite, DK_StrUtils, DK_Dialogs,
  DK_SheetExporter, UCalendar, UCardForm, URepairEditForm;

type

  { TRepairForm }

  TRepairForm = class(TForm)
    AddButton: TSpeedButton;
    DelButton: TSpeedButton;
    DividerBevel5: TDividerBevel;
    DividerBevel6: TDividerBevel;
    DividerBevel7: TDividerBevel;
    EditButton: TSpeedButton;
    ExportButton: TRxSpeedButton;
    Label2: TLabel;
    Label4: TLabel;
    Label5: TLabel;
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
  private
    CardForm: TCardForm;
    VSTMotorsTable: TVSTTable;
    VSTTypeTable: TVSTTable;
    VSTOrderTable: TVSTTable;
    RecIDs, MotorIDs: TIntVector;

    ArrivalDates, SendingDates: TDateVector;
    Passports, DayCounts: TIntVector;
    MotorNames, MotorNums, RepairNotes: TStrVector;



    procedure TypeSelect;
    procedure OrderSelect;
    procedure MotorSelect;

    procedure RepairEditFormOpen(const AEditType: Byte);
  public
    procedure ShowRepair;
  end;

var
  RepairForm: TRepairForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TRepairForm }

procedure TRepairForm.FormCreate(Sender: TObject);
var
  V: TStrVector;
  i: Integer;
begin
  MainForm.SetNamesPanelsVisible(True, False);
  CardForm:= CreateCardForm(RepairForm, CardPanel);

  VSTMotorsTable:= TVSTTable.Create(VT1);
  VSTMotorsTable.OnSelect:= @MotorSelect;
  VSTMotorsTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTMotorsTable.HeaderFont.Style:= [fsBold];
  VSTMotorsTable.AddColumn('№ п/п', 50);
  VSTMotorsTable.AddColumn('Наименование', 200);
  VSTMotorsTable.AddColumn('Номер', 100);
  VSTMotorsTable.AddColumn('Наличие паспорта', 130);
  VSTMotorsTable.AddColumn('Дата прибытия', 100);
  VSTMotorsTable.AddColumn('Дата убытия', 100);
  VSTMotorsTable.AddColumn('Срок ремонта (рабочих дней)', 200);
  //VSTMotorsTable.AutosizeColumnDisable;
  VSTMotorsTable.AddColumn('Примечание');
  VSTMotorsTable.CanSelect:= True;
  VSTMotorsTable.Draw;

  VSTTypeTable:= TVSTTable.Create(VT2);
  VSTTypeTable.OnSelect:= @TypeSelect;
  VSTTypeTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTTypeTable.HeaderVisible:= False;
  VSTTypeTable.GridLinesVisible:= False;
  VSTTypeTable.CanSelect:= True;
  VSTTypeTable.CanUnselect:= False;
  VSTTypeTable.AddColumn('Список');
  V:= VCreateStr(['все', 'в ремонте', 'отремонтированные']);
  i:= Round(VT2.DefaultNodeHeight*96/ScreenInfo.PixelsPerInchY);
  VT2.Height:= Length(V)*i;
  VSTTypeTable.SetColumn('Список', V, taLeftJustify);
  VSTTypeTable.Draw;

  VSTOrderTable:= TVSTTable.Create(VT3);
  VSTOrderTable.OnSelect:= @OrderSelect;
  VSTOrderTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTOrderTable.HeaderVisible:= False;
  VSTOrderTable.GridLinesVisible:= False;
  VSTOrderTable.CanSelect:= True;
  VSTOrderTable.CanUnselect:= False;
  VSTOrderTable.AddColumn('Список');
  V:= VCreateStr(['дате прибытия', 'номеру']);
  i:= Round(VT3.DefaultNodeHeight*96/ScreenInfo.PixelsPerInchY);
  VT3.Height:= Length(V)*i;
  VSTOrderTable.SetColumn('Список', V, taLeftJustify);
  VSTOrderTable.Draw;

  VSTTypeTable.Select(1);
  VSTOrderTable.Select(0);
  MotorCardCheckBox.Checked:= False;
end;

procedure TRepairForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TSheetExporter;
  Sheet: TsWorksheet;
  RepairSheet: TRepairSheet;
begin
  Exporter:= TSheetExporter.Create;
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
  RepairEditFormOpen(1);
end;

procedure TRepairForm.DelButtonClick(Sender: TObject);
begin
  if not VSTMotorsTable.IsSelected then Exit;
  if not Confirm('Удалить данные о гарантийном ремонте?') then Exit;
  SQLite.RepairDelete(RecIDs[VSTMotorsTable.SelectedIndex]);
  ShowRepair;
end;

procedure TRepairForm.EditButtonClick(Sender: TObject);
begin
  RepairEditFormOpen(2);
end;

procedure TRepairForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTMotorsTable) then FreeAndNil(VSTMotorsTable);
  if Assigned(VSTTypeTable) then FreeAndNil(VSTTypeTable);
  if Assigned(VSTOrderTable) then FreeAndNil(VSTOrderTable);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TRepairForm.FormShow(Sender: TObject);
begin
  ShowRepair;
end;

procedure TRepairForm.MotorCardCheckBoxChange(Sender: TObject);
begin
  if MotorCardCheckBox.Checked then
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

procedure TRepairForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TRepairForm.MotorNumEditChange(Sender: TObject);
begin
  ShowRepair;
end;

procedure TRepairForm.TypeSelect;
begin
  ShowRepair;
end;

procedure TRepairForm.OrderSelect;
begin
  ShowRepair;
end;

procedure TRepairForm.MotorSelect;
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

procedure TRepairForm.RepairEditFormOpen(const AEditType: Byte);
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
    if RepairEditForm.ShowModal = mrOK then ShowRepair;
  finally
    FreeAndNil(RepairEditForm);
  end;
end;

procedure TRepairForm.ShowRepair;
var
  MotorNumberLike: String;
  i: Integer;
  PassStrs, DayCountsStrs, SendingDatesStrs, ArrivalDatesStrs: TStrVector;
begin
  if (not VSTTypeTable.IsSelected) or (not VSTOrderTable.IsSelected) then Exit;

  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);
    MotorNumberLike:= STrim(MotorNumEdit.Text);

    SQLite.RepairListLoad(MainForm.UsedNameIDs, MotorNumberLike,
                          VSTOrderTable.SelectedIndex+1, VSTTypeTable.SelectedIndex,
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

