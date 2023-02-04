unit URepairForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, EditBtn, VirtualTrees, DividerBevel, DK_VSTTables, USheetUtils,
  rxctrls, DK_Vector, USQLite, DK_StrUtils,
  DK_SheetExporter, UCalendar, UCardForm;

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
    Panel6: TPanel;
    Panel7: TPanel;
    Panel9: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    TopToolsPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    VT3: TVirtualStringTree;
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
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
    MotorNames, MotorNums: TStrVector;



    procedure TypeSelect;
    procedure OrderSelect;
    procedure MotorSelect;

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
  v: TStrVector;
begin
  MainForm.SetNamesPanelsVisible(True, False);
  CardForm:= CrateCardForm(RepairForm, CardPanel);

  VSTMotorsTable:= TVSTTable.Create(VT1);
  VSTMotorsTable.OnSelect:= @MotorSelect;
  VSTMotorsTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTMotorsTable.HeaderFont.Style:= [fsBold];
  VSTMotorsTable.AddColumn('№ п/п', 60);
  VSTMotorsTable.AddColumn('Наименование', 250);
  VSTMotorsTable.AddColumn('Номер', 150);
  VSTMotorsTable.AddColumn('Наличие паспорта', 150);
  VSTMotorsTable.AddColumn('Дата прибытия', 150);
  VSTMotorsTable.AddColumn('Дата убытия', 150);
  VSTMotorsTable.AddColumn('Срок ремонта (рабочих дней)', 200);
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
  V:= VCreateStr(['Все', 'В ремонте', 'Отремонтированные']);
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
  V:= VCreateStr(['Дате прибытия', 'Номеру']);
  VSTOrderTable.SetColumn('Список', V, taLeftJustify);
  VSTOrderTable.Draw;

  VSTTypeTable.Select(1);
  VSTOrderTable.Select(0);
  MotorCardCheckBox.Checked:= False;

  ShowRepair;
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

procedure TRepairForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTMotorsTable) then FreeAndNil(VSTMotorsTable);
  if Assigned(VSTTypeTable) then FreeAndNil(VSTTypeTable);
  if Assigned(VSTOrderTable) then FreeAndNil(VSTOrderTable);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TRepairForm.MotorCardCheckBoxChange(Sender: TObject);
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
  MotorID:= 0;
  if VSTMotorsTable.IsSelected then
    MotorID:= MotorIDs[VSTMotorsTable.SelectedIndex];
  CardForm.ShowCard(MotorID);
end;



procedure TRepairForm.ShowRepair;
var
  MotorNumberLike: String;
  i: Integer;
  PassStrs, DayCountsStrs, SendingDatesStrs: TStrVector;
begin
  if (not VSTTypeTable.IsSelected) or (not VSTOrderTable.IsSelected) then Exit;

  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);
    MotorNumberLike:= STrim(MotorNumEdit.Text);

    SQLite.RepairListLoad(MainForm.UsedNameIDs, MotorNumberLike,
                          VSTOrderTable.SelectedIndex+1, VSTTypeTable.SelectedIndex,
                          ArrivalDates, SendingDates,
                          RecIDs, MotorIDs, Passports,
                          MotorNames, MotorNums);

    VDim(PassStrs{%H-}, Length(Passports));
    for i:= 0 to High(PassStrs) do
      if Passports[i]>0 then
        PassStrs[i]:= CHECK_SYMBOL;

    VDim(DayCounts{%H-}, Length(MotorIDs));
    for i:= 0 to High(DayCounts) do
      DayCounts[i]:= SQLite.LoadWorkDaysCountInPeriod(ArrivalDates[i], SendingDates[i]);

    VDim(DayCountsStrs{%H-}, Length(ArrivalDates));
    VDim(SendingDatesStrs{%H-}, Length(ArrivalDates));
    for i:= 0 to High(DayCountsStrs) do
    begin
      if SendingDates[i]>0 then
        SendingDatesStrs[i]:= FormatDateTime('dd.mm.yyyy', SendingDates[i]);
      if DayCounts[i]>0 then
        DayCountsStrs[i]:= IntToStr(DayCounts[i]);
    end;

    VSTMotorsTable.ValuesClear;
    VSTMotorsTable.SetColumn('№ п/п', VIntToStr(VOrder(Length(MotorIDs))));
    VSTMotorsTable.SetColumn('Наименование', MotorNames, taLeftJustify);
    VSTMotorsTable.SetColumn('Номер', MotorNums);
    VSTMotorsTable.SetColumn('Наличие паспорта', PassStrs);
    VSTMotorsTable.SetColumn('Дата прибытия', VFormatDateTime('dd.mm.yyyy', ArrivalDates));
    VSTMotorsTable.SetColumn('Дата убытия', SendingDatesStrs);
    VSTMotorsTable.SetColumn('Срок ремонта (рабочих дней)', DayCountsStrs);


    VSTMotorsTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;


end;

end.

