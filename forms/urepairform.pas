unit URepairForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, DividerBevel, VirtualTrees,
  //DK packages utils
  DK_VSTTables, DK_StrUtils, DK_Dialogs, DK_Const, DK_Vector, DK_SheetExporter,
  DK_VSTParamList, DK_CtrlUtils, DK_Filter,
  //Project utils
  UVars, USheets, UCalendar,
  //Forms
  UCardForm, URepairEditForm;

type

  { TRepairForm }

  TRepairForm = class(TForm)
    AddButton: TSpeedButton;
    DelButton: TSpeedButton;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    EditButton: TSpeedButton;
    ExportButton: TSpeedButton;
    MotorCardCheckBox: TCheckBox;
    Panel1: TPanel;
    CardPanel: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel5: TPanel;
    SettingClientPanel: TPanel;
    FilterPanel: TPanel;
    MainPanel: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    ToolPanel: TPanel;
    VT1: TVirtualStringTree;
    procedure AddButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorCardCheckBoxChange(Sender: TObject);
    procedure VT1DblClick(Sender: TObject);
  private
    CardForm: TCardForm;
    ParamList: TVSTParamList;
    VSTMotorsTable: TVSTTable;

    RecIDs, MotorIDs: TIntVector;
    FilterString: String;

    ArrivalDates, SendingDates: TDateVector;
    Passports, DayCounts: TIntVector;
    MotorNames, MotorNums, RepairNotes: TStrVector;

    procedure CreateParamList;

    procedure CreateMotorsTable;
    procedure SelectMotor;
    procedure FilterMotor(const AFilterString: String);

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
  MainForm.SetNamesPanelsVisible(True, False);
  CardForm:= CreateCardForm(RepairForm, CardPanel);
  CreateMotorsTable;
  CreateParamList;
  FilterString:= EmptyStr;
  DKFilterCreate('Поиск по номеру:', FilterPanel, @FilterMotor, 300, 500);
  MotorCardCheckBox.Checked:= False;
end;

procedure TRepairForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTMotorsTable) then FreeAndNil(VSTMotorsTable);
  if Assigned(ParamList) then FreeAndNil(ParamList);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TRepairForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);
  SetToolButtons([AddButton, DelButton, EditButton]);
  Images.ToButtons([ExportButton, AddButton, DelButton, EditButton]);
  ParamList.AutoHeight;
  ViewUpdate;
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
    RepairSheet:= TRepairSheet.Create(Sheet, nil, GridFont);
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
  DataBase.RepairDelete(RecIDs[VSTMotorsTable.SelectedIndex]);
  ViewUpdate;
end;

procedure TRepairForm.EditButtonClick(Sender: TObject);
begin
  OpenRepairEditForm(2);
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

procedure TRepairForm.VT1DblClick(Sender: TObject);
begin
  if not VSTMotorsTable.IsSelected then Exit;
  OpenRepairEditForm(2);
end;

procedure TRepairForm.CreateMotorsTable;
begin
  VSTMotorsTable:= TVSTTable.Create(VT1);
  VSTMotorsTable.SetSingleFont(GridFont);
  VSTMotorsTable.OnSelect:= @SelectMotor;
  VSTMotorsTable.HeaderFont.Style:= [fsBold];
  VSTMotorsTable.AddColumn('№ п/п', 50);
  VSTMotorsTable.AddColumn('Наименование', 200);
  VSTMotorsTable.AddColumn('Номер', 100);
  VSTMotorsTable.AddColumn('Наличие паспорта', 140);
  VSTMotorsTable.AddColumn('Дата прибытия', 120);
  VSTMotorsTable.AddColumn('Дата убытия', 120);
  VSTMotorsTable.AddColumn('Срок ремонта (рабочих дней)', 220);
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

procedure TRepairForm.FilterMotor(const AFilterString: String);
begin
  FilterString:= AFilterString;
  ViewUpdate;
end;

procedure TRepairForm.CreateParamList;
var
  S: String;
  V: TStrVector;
begin
  ParamList:= TVSTParamList.Create(SettingClientPanel);

  S:= 'Отображать:';
  V:= VCreateStr([
    'все',
    'в ремонте',
    'отремонтированные'
  ]);
  ParamList.AddStringList('TypeList', S, V, @ViewUpdate, 1);

  S:= 'Упорядочить по:';
  V:= VCreateStr([
    'дате прибытия',
    'дате убытия',
    'номеру'
  ]);
  ParamList.AddStringList('OrderList', S, V, @ViewUpdate);
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
  i: Integer;
  PassStrs, DayCountsStrs, SendingDatesStrs, ArrivalDatesStrs: TStrVector;
begin
  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);

    DataBase.RepairListLoad(MainForm.UsedNameIDs, STrim(FilterString),
                          ParamList.Selected['OrderList']+1, ParamList.Selected['TypeList'],
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

