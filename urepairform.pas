unit URepairForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, EditBtn, VirtualTrees, DividerBevel, DK_VSTTables, SheetUtils,
  rxctrls, fpspreadsheetgrid, DK_Vector, USQLite, DK_StrUtils, DK_DateUtils;

type

  { TRepairForm }

  TRepairForm = class(TForm)
    AddButton: TSpeedButton;
    DelButton: TSpeedButton;
    DividerBevel5: TDividerBevel;
    DividerBevel6: TDividerBevel;
    DividerBevel7: TDividerBevel;
    DividerBevel9: TDividerBevel;
    EditButton: TSpeedButton;
    ExportButton: TRxSpeedButton;
    InfoGrid: TsWorksheetGrid;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    MotorNumEdit: TEditButton;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    RadioButton5: TRadioButton;
    Splitter2: TSplitter;
    TopToolsPanel: TPanel;
    VT1: TVirtualStringTree;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure RadioButton3Click(Sender: TObject);
    procedure RadioButton4Click(Sender: TObject);
    procedure RadioButton5Click(Sender: TObject);
    procedure VT1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    VSTTable: TVSTTable;
    MotorInfoSheet: TMotorInfoSheet;
    RecIDs, MotorIDs: TIntVector;

    procedure InfoOpen;
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
begin
  MainForm.SetNamesPanelsVisible(True, False);

  VSTTable:= TVSTTable.Create(VT1);
  VSTTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTTable.HeaderFont.Style:= [fsBold];
  VSTTable.AddColumn('№ п/п', 80);
  VSTTable.AddColumn('Наименование', 300);
  VSTTable.AddColumn('Номер', 200);
  VSTTable.AddColumn('Наличие паспорта', 200);
  VSTTable.AddColumn('Прибыл в ремонт', 200);
  VSTTable.AddColumn('Убыл из ремонта', 200);
  VSTTable.AddColumn('Срок ремонта (дней)', 200);
  VSTTable.CanSelect:= True;
  VSTTable.AutosizeColumnDisable;
  VSTTable.Draw;

  MotorInfoSheet:= TMotorInfoSheet.Create(InfoGrid);

  ShowRepair;
end;

procedure TRepairForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTTable) then FreeAndNil(VSTTable);
  if Assigned(MotorInfoSheet) then FreeAndNil(MotorInfoSheet);
end;

procedure TRepairForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TRepairForm.MotorNumEditChange(Sender: TObject);
begin
  ShowRepair;
end;

procedure TRepairForm.RadioButton1Click(Sender: TObject);
begin
  ShowRepair;
end;

procedure TRepairForm.RadioButton2Click(Sender: TObject);
begin
  ShowRepair;
end;

procedure TRepairForm.RadioButton3Click(Sender: TObject);
begin
  ShowRepair;
end;

procedure TRepairForm.RadioButton4Click(Sender: TObject);
begin
  ShowRepair;
end;

procedure TRepairForm.RadioButton5Click(Sender: TObject);
begin
  ShowRepair;
end;

procedure TRepairForm.VT1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  InfoGrid.Clear;
  if not VSTTable.IsSelected then Exit;
  InfoOpen;
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
  if not VSTTable.IsSelected then Exit;

  MotorID:= MotorIDs[VSTTable.SelectedIndex];
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

procedure TRepairForm.ShowRepair;
var
  MotorNumberLike: String;
  i, ListType, OrderType: Integer;
  ArrivalDates, SendingDates: TDateVector;
  Passports: TIntVector;
  MotorNames, MotorNums, PassStrs, DayCounts, SendingDatesStrs: TStrVector;
begin
  Screen.Cursor:= crHourGlass;
  try
    InfoGrid.Clear;
    MotorNumberLike:= STrim(MotorNumEdit.Text);

    ListType:= 0;
    if RadioButton1.Checked then
      ListType:= 1
    else if RadioButton2.Checked then
      ListType:= 2;

    if RadioButton4.Checked then
      OrderType:= 1
    else if RadioButton5.Checked then
      OrderType:= 2;


    SQLite.RepairListLoad(MainForm.UsedNameIDs, MotorNumberLike,
                          OrderType, ListType,
                          ArrivalDates, SendingDates,
                          RecIDs, MotorIDs, Passports,
                          MotorNames, MotorNums);

    VDim(PassStrs{%H-}, Length(Passports));
    for i:= 0 to High(PassStrs) do
      if Passports[i]>0 then
        PassStrs[i]:= CHECK_SYMBOL;

    VDim(DayCounts{%H-}, Length(ArrivalDates));
    VDim(SendingDatesStrs{%H-}, Length(ArrivalDates));
    for i:= 0 to High(DayCounts) do
    begin
      if SendingDates[i]>0 then
      begin
        SendingDatesStrs[i]:= FormatDateTime('dd.mm.yyyy', SendingDates[i]);
        DayCounts[i]:= IntToStr(DaysBetweenDates(ArrivalDates[i], SendingDates[i]) + 1)
      end
      else
        DayCounts[i]:= IntToStr(DaysBetweenDates(ArrivalDates[i], Date) + 1);
    end;

    VSTTable.ValuesClear;
    VSTTable.SetColumn('№ п/п', VIntToStr(VOrder(Length(MotorIDs))));
    VSTTable.SetColumn('Наименование', MotorNames, taLeftJustify);
    VSTTable.SetColumn('Номер', MotorNums);
    VSTTable.SetColumn('Наличие паспорта', PassStrs);
    VSTTable.SetColumn('Прибыл в ремонт', VFormatDateTime('dd.mm.yyyy', ArrivalDates));
    VSTTable.SetColumn('Убыл из ремонта', SendingDatesStrs);
    VSTTable.SetColumn('Срок ремонта (дней)', DayCounts);


    VSTTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;


end;

end.

