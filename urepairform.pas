unit URepairForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, EditBtn, VirtualTrees, DividerBevel, DK_VSTTables, USheetUtils,
  rxctrls, fpspreadsheetgrid, DK_Vector, USQLite, DK_StrUtils,
  DK_SheetExporter, fpspreadsheet, fpstypes, UCalendar;

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
    InfoGrid: TsWorksheetGrid;
    Label2: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    MoreInfoCheckBox: TCheckBox;
    MotorNumEdit: TEditButton;
    Panel1: TPanel;
    Panel10: TPanel;
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
    procedure MoreInfoCheckBoxChange(Sender: TObject);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
  private
    VSTMotorTable: TVSTTable;
    VSTTypeTable: TVSTTable;
    VSTOrderTable: TVSTTable;
    MotorInfoSheet: TMotorInfoSheet;
    RecIDs, MotorIDs: TIntVector;

    ArrivalDates, SendingDates: TDateVector;
    Passports, DayCounts: TIntVector;
    MotorNames, MotorNums: TStrVector;

    function CalcRepairWorkDaysCount(const ABeginDate, AEndDate: TDate): Integer;
    procedure InfoOpen;
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

  VSTMotorTable:= TVSTTable.Create(VT1);
  VSTMotorTable.OnSelect:= @MotorSelect;
  VSTMotorTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTMotorTable.HeaderFont.Style:= [fsBold];
  VSTMotorTable.AddColumn('№ п/п', 60);
  VSTMotorTable.AddColumn('Наименование', 250);
  VSTMotorTable.AddColumn('Номер', 150);
  VSTMotorTable.AddColumn('Наличие паспорта', 150);
  VSTMotorTable.AddColumn('Прибыл в ремонт', 150);
  VSTMotorTable.AddColumn('Убыл из ремонта', 150);
  VSTMotorTable.AddColumn('Срок ремонта (рабочих дней)', 200);
  VSTMotorTable.CanSelect:= True;
  VSTMotorTable.AutosizeColumnDisable;
  VSTMotorTable.Draw;

  MotorInfoSheet:= TMotorInfoSheet.Create(InfoGrid);

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
  if Assigned(VSTMotorTable) then FreeAndNil(VSTMotorTable);
  if Assigned(VSTTypeTable) then FreeAndNil(VSTTypeTable);
  if Assigned(VSTOrderTable) then FreeAndNil(VSTOrderTable);
  if Assigned(MotorInfoSheet) then FreeAndNil(MotorInfoSheet);
end;

procedure TRepairForm.MoreInfoCheckBoxChange(Sender: TObject);
begin
  If MoreInfoCheckBox.Checked then
  begin
    VT1.Align:= alCustom;
    Splitter2.Visible:= True;
    Splitter2.Align:= alTop;
    Panel10.Visible:= True;
    Splitter2.Align:= alBottom;
    VT1.Align:= alClient;
  end
  else begin
    Panel10.Visible:= False;
    Splitter2.Visible:= False;
  end;

  VSTMotorTable.CanSelect:= MoreInfoCheckBox.Checked;
end;

procedure TRepairForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TRepairForm.MotorNumEditChange(Sender: TObject);
begin
  ShowRepair;
end;



procedure TRepairForm.InfoOpen;
var
  MotorID: Integer;
  BuildDate, SendDate: TDate;
  MotorName, MotorNum, RotorNum, ReceiverName, Sers: String;
  TestDates, RecDates: TDateVector;
  TestResults, Mileages, Opinions: TIntVector;
  TestNotes, PlaceNames, FactoryNames, Departures,
  DefectNames, ReasonNames, RecNotes: TStrVector;
begin
  if VIsNil(MotorIDs) then Exit;
  if not VSTMotorTable.IsSelected then Exit;

  MotorID:= MotorIDs[VSTMotorTable.SelectedIndex];
  SQLite.MotorInfoLoad(MotorID, BuildDate, SendDate,
                MotorName, MotorNum, Sers, RotorNum, ReceiverName,
                TestDates, TestResults, TestNotes);

  SQLite.ReclamationListLoad(MotorID, RecDates, Mileages, Opinions,
                      PlaceNames, FactoryNames, Departures,
                      DefectNames, ReasonNames, RecNotes);


  MotorInfoSheet.Draw(BuildDate, SendDate, MotorName, MotorNum, Sers,
                      RotorNum, ReceiverName, TestDates, TestResults, TestNotes,
                      RecDates, Mileages, Opinions, PlaceNames, FactoryNames,
                      Departures, DefectNames, ReasonNames, RecNotes);
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
begin
  InfoGrid.Clear;
  if not VSTMotorTable.IsSelected then Exit;
  InfoOpen;
end;

function TRepairForm.CalcRepairWorkDaysCount(const ABeginDate, AEndDate: TDate): Integer;
var
  Calendar: TCalendar;
  D: TDate;
begin
  Result:= 0;
  if ABeginDate=0 then Exit;

  D:= AEndDate;
  if D=0 then D:= Date;
  Calendar:= SQLite.LoadCalendar(ABeginDate, D);
  try
    Result:= Calendar.WorkDaysCount;
  finally
    FreeAndNil(Calendar);
  end;
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
    InfoGrid.Clear;
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
      DayCounts[i]:= CalcRepairWorkDaysCount(ArrivalDates[i], SendingDates[i]);

    VDim(DayCountsStrs{%H-}, Length(ArrivalDates));
    VDim(SendingDatesStrs{%H-}, Length(ArrivalDates));
    for i:= 0 to High(DayCountsStrs) do
    begin
      if SendingDates[i]>0 then
        SendingDatesStrs[i]:= FormatDateTime('dd.mm.yyyy', SendingDates[i]);
      if DayCounts[i]>0 then
        DayCountsStrs[i]:= IntToStr(DayCounts[i]);
    end;

    VSTMotorTable.ValuesClear;
    VSTMotorTable.SetColumn('№ п/п', VIntToStr(VOrder(Length(MotorIDs))));
    VSTMotorTable.SetColumn('Наименование', MotorNames, taLeftJustify);
    VSTMotorTable.SetColumn('Номер', MotorNums);
    VSTMotorTable.SetColumn('Наличие паспорта', PassStrs);
    VSTMotorTable.SetColumn('Прибыл в ремонт', VFormatDateTime('dd.mm.yyyy', ArrivalDates));
    VSTMotorTable.SetColumn('Убыл из ремонта', SendingDatesStrs);
    VSTMotorTable.SetColumn('Срок ремонта (рабочих дней)', DayCountsStrs);


    VSTMotorTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;


end;

end.

