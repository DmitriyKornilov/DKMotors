unit UReclamationForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons, StdCtrls,
  ComCtrls, VirtualTrees, fpspreadsheetgrid, DateTimePicker, DividerBevel,
  //DK packages utils
  DK_Vector, DK_DateUtils, DK_StrUtils, DK_Dialogs, DK_SheetExporter, DK_Zoom, DK_Const,
  DK_VSTParamList, DK_CtrlUtils, DK_Filter,
  //Project utils
  UVars, USheets,
  //Forms
  UCardForm, UReclamationEditForm, URepairEditForm, UControlListEditForm;

type

  { TReclamationForm }

  TReclamationForm = class(TForm)
    AddButton: TSpeedButton;
    Bevel1: TBevel;
    CardPanel: TPanel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    DividerBevel4: TDividerBevel;
    ExportButton: TSpeedButton;
    Label3: TLabel;
    SettingClientPanel: TPanel;
    LogGrid: TsWorksheetGrid;
    MainPanel: TPanel;
    Panel4: TPanel;
    RepairButton: TSpeedButton;
    MotorCardCheckBox: TCheckBox;
    Panel2: TPanel;
    DelButton: TSpeedButton;
    EditButton: TSpeedButton;
    Panel3: TPanel;
    Panel5: TPanel;
    FilterPanel: TPanel;
    ControlButton: TSpeedButton;
    ReportPeriodPanel: TPanel;
    Splitter0: TSplitter;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    ZoomPanel: TPanel;
    procedure AddButtonClick(Sender: TObject);
    procedure ControlButtonClick(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LogGridDblClick(Sender: TObject);
    procedure LogGridMouseDown(Sender: TObject; Button: TMouseButton;
      {%H-}Shift: TShiftState; X, Y: Integer);
    procedure MotorCardCheckBoxChange(Sender: TObject);
    procedure RepairButtonClick(Sender: TObject);
  private
    CardForm: TCardForm;
    ParamList: TVSTParamList;
    CanShow: Boolean;
    ZoomPercent: Integer;
    FilterString: String;
    SelectedIndex: Integer;
    ReclamationSheet: TReclamationSheet;

    RecDates, BuildDates, ArrivalDates, SendingDates: TDateVector;
    RecIDs, MotorIDs, Mileages, Opinions, ReasonColors, Passports: TIntVector;
    PlaceNames, FactoryNames, Departures: TStrVector;
    DefectNames, ReasonNames, RecNotes: TStrVector;
    MotorNames, MotorNums: TStrVector;

    ListReasonIDs: TIntVector;
    ListDefectIDs: TIntVector;

    procedure ClearSelection;
    procedure SelectLine(const ARow: Integer);

    procedure DeleteRaclamation;
    procedure OpenReclamationEditForm(const AEditType: Byte);

    procedure OpenRepairEditForm;
    procedure OpenControlEditForm;

    procedure LoadReclamation;
    procedure DrawReclamation(const AZoomPercent: Integer);
    procedure ExportReclamation;
    procedure FilterReclamation(const AFilterString: String);

    procedure CreateParamList;
  public
    procedure ViewUpdate;
  end;

var
  ReclamationForm: TReclamationForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TReclamationForm }

procedure TReclamationForm.FormCreate(Sender: TObject);
begin
  CanShow:= False;
  CardForm:= CreateCardForm(ReclamationForm, CardPanel);
  MotorCardCheckBox.Checked:= False;
  MainForm.SetNamesPanelsVisible(True, False);
  SelectedIndex:= -1;

  ZoomPercent:= 100;
  CreateZoomControls(50, 150, ZoomPercent, ZoomPanel, @DrawReclamation, True);

  ReclamationSheet:= TReclamationSheet.Create(LogGrid.Worksheet, LogGrid, GridFont);
  ReclamationSheet.Font.Size:= ReclamationSheet.Font.Size-1;

  CreateParamList;

  DateTimePicker2.Date:= FirstDayInYear(Date);
  DateTimePicker1.Date:= LastDayInYear(Date);

  FilterString:= EmptyStr;
  DKFilterCreate('Поиск по номеру:', FilterPanel, @FilterReclamation, 300, 500);

  CanShow:= True;
end;

procedure TReclamationForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CardForm) then FreeAndNil(CardForm);
  if Assigned(ReclamationSheet) then FReeAndNil(ReclamationSheet);
  if Assigned(ParamList) then FreeAndNil(ParamList);
end;

procedure TReclamationForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);
  SetToolButtons([
    AddButton, DelButton, EditButton, RepairButton, ControlButton
  ]);
  Images.ToButtons([
    ExportButton,
    AddButton, DelButton, EditButton, RepairButton, ControlButton
  ]);

  ParamList.AutoHeight;
  ViewUpdate;
end;

procedure TReclamationForm.LogGridDblClick(Sender: TObject);
begin
  if SelectedIndex<0 then Exit;
  OpenReclamationEditForm(2);
end;

procedure TReclamationForm.LogGridMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  R,C, MotorID: Integer;
begin
  if Button=mbRight then
    ClearSelection;
  if Button=mbLeft  then
  begin
    LogGrid.MouseToCell(X,Y,C{%H-},R{%H-});
    SelectLine(R);
  end;

  MotorID:= 0;
  if SelectedIndex>=0 then
    MotorID:= MotorIDs[SelectedIndex];
  CardForm.ShowCard(MotorID);
end;

procedure TReclamationForm.MotorCardCheckBoxChange(Sender: TObject);
begin
  if MotorCardCheckBox.Checked then
  begin
    Panel4.Align:= alCustom;
    Splitter1.Visible:= True;
    Splitter1.Align:= alTop;
    CardPanel.Visible:= True;
    Splitter1.Align:= alBottom;
    Panel4.Align:= alClient;
  end
  else begin
    CardPanel.Visible:= False;
    Splitter1.Visible:= False;
  end;
end;

procedure TReclamationForm.RepairButtonClick(Sender: TObject);
begin
  OpenRepairEditForm;
end;

procedure TReclamationForm.ClearSelection;
begin
  if SelectedIndex>-1 then
  begin
    ReclamationSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= -1;
  end;
  DelButton.Enabled:= False;
  EditButton.Enabled:= False;
  RepairButton.Enabled:= False;
  ControlButton.Enabled:= False;
end;

procedure TReclamationForm.SelectLine(const ARow: Integer);
const
  FirstRow = 2;
begin
  if VIsNil(RecDates) then Exit;
  if (ARow>=FirstRow) and (ARow<LogGrid.RowCount-1) then  //клик ниже заголовка
  begin
    if SelectedIndex>-1 then
      ReclamationSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= ARow - FirstRow;
    ReclamationSheet.DrawLine(SelectedIndex, True);
    DelButton.Enabled:= True;
    EditButton.Enabled:= True;
    RepairButton.Enabled:= True;
    ControlButton.Enabled:= True;
  end;
end;

procedure TReclamationForm.ViewUpdate;
begin
  if not CanShow then Exit;
  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);
    ClearSelection;
    LoadReclamation;
    DrawReclamation(ZoomPercent);
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TReclamationForm.LoadReclamation;
var
  UsedReasonIDs: TIntVector;
  UsedDefectIDs: TIntVector;
  OrderIndex: Integer;
begin
  OrderIndex:= 0;
  if ParamList.IsSelected['OrderList'] then
    OrderIndex:= ParamList.Selected['OrderList'];

  UsedReasonIDs:= VCut(ListReasonIDs, ParamList.Checkeds['ReasonList']);
  UsedDefectIDs:= VCut(ListDefectIDs, ParamList.Checkeds['DefectList']);

  DataBase.ReclamationListLoad(DateTimePicker2.Date, DateTimePicker1.Date,
                        MainForm.UsedNameIDs, UsedDefectIDs, UsedReasonIDs,
                        OrderIndex, STrim(FilterString),
                        RecDates, BuildDates, ArrivalDates, SendingDates,
                        RecIDs, MotorIDs, Mileages, Opinions,
                        ReasonColors, Passports,
                        PlaceNames, FactoryNames, Departures,
                        DefectNames, ReasonNames, RecNotes,
                        MotorNames, MotorNums);
end;

procedure TReclamationForm.DrawReclamation(const AZoomPercent: Integer);
begin
  ZoomPercent:= AZoomPercent;
  ReclamationSheet.Zoom(ZoomPercent);
  ReclamationSheet.Draw(RecDates, BuildDates, ArrivalDates, SendingDates,
                          Mileages, Opinions, ReasonColors,
                          PlaceNames, FactoryNames, Departures, DefectNames,
                          ReasonNames, RecNotes, MotorNames, MotorNums);
end;

procedure TReclamationForm.ExportReclamation;
var
  Drawer: TReclamationSheet;
  Sheet: TsWorksheet;
  Exporter: TSheetsExporter;
begin
  Exporter:= TSheetsExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    Drawer:= TReclamationSheet.Create(Sheet, nil, GridFont);
    Drawer.Font.Size:= Drawer.Font.Size-1;
    Drawer.Draw(RecDates, BuildDates, ArrivalDates, SendingDates,
                        Mileages, Opinions, ReasonColors,
                        PlaceNames, FactoryNames, Departures, DefectNames,
                        ReasonNames, RecNotes, MotorNames, MotorNums);
    Exporter.PageSettings(spoLandscape);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Drawer);
    FreeAndNil(Exporter);
  end;
end;

procedure TReclamationForm.FilterReclamation(const AFilterString: String);
begin
  FilterString:= AFilterString;
  ViewUpdate;
end;

procedure TReclamationForm.CreateParamList;
var
  S: String;
  V: TStrVector;
  B: TBoolVector;
  i: Integer;
begin
  ParamList:= TVSTParamList.Create(SettingClientPanel);

  S:= 'Упорядочить по:';
  V:= VCreateStr([
    'дате уведомления',
    'номеру двигателя',
    'дате сборки',
    'предприятию',
    'неисправному элементу',
    'причине неисправности'
  ]);
  ParamList.AddStringList('OrderList', S, V, @ViewUpdate);

  S:= 'Неисправный элемент:';
  DataBase.KeyPickList('RECLAMATIONDEFECTS', 'DefectID', 'DefectName',
                     ListDefectIDs, V, False, 'DefectName');
  for i:= 0 to High(V) do
    V[i]:= SFirstLower(V[i]);
  B:= VCreateBool(Length(V), True);
  ParamList.AddCheckList('DefectList', S, V, @ViewUpdate, B);

  S:= 'Причина неисправности:';
  DataBase.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
                     ListReasonIDs, V);
  for i:= 0 to High(V) do
    V[i]:= SFirstLower(V[i]);
  B:= VCreateBool(Length(V), True);
  ParamList.AddCheckList('ReasonList', S, V, @ViewUpdate, B);
end;

procedure TReclamationForm.DeleteRaclamation;
begin
  if SelectedIndex<0 then Exit;
  if not Confirm('Удалить рекламацию?') then Exit;
  DataBase.Delete('RECLAMATIONS', 'RecID', RecIDs[SelectedIndex]);
  ClearSelection;
  ViewUpdate;
end;

procedure TReclamationForm.OpenReclamationEditForm(const AEditType: Byte);
var
  ReclamationEditForm: TReclamationEditForm;
begin
  ReclamationEditForm:= TReclamationEditForm.Create(ReclamationForm);
  try
    ReclamationEditForm.RecID:= 0;
    if AEditType=2 then
      ReclamationEditForm.RecID:= RecIDs[SelectedIndex];
    if ReclamationEditForm.ShowModal = mrOK then ViewUpdate;
  finally
    FreeAndNil(ReclamationEditForm);
  end;
  ClearSelection;
end;

procedure TReclamationForm.OpenRepairEditForm;
var
  RepairEditForm: TRepairEditForm;
begin
  RepairEditForm:= TRepairEditForm.Create(ReclamationForm);
  try
    RepairEditForm.RecID:= RecIDs[SelectedIndex];
    RepairEditForm.MotorName:= MotorNames[SelectedIndex];
    RepairEditForm.MotorNum:= MotorNums[SelectedIndex];
    if RepairEditForm.ShowModal = mrOK then ViewUpdate;
  finally
    FreeAndNil(RepairEditForm);
  end;
end;

procedure TReclamationForm.OpenControlEditForm;
var
  ControlListEditForm: TControlListEditForm;
begin
  ControlListEditForm:= TControlListEditForm.Create(ReclamationForm);
  try
    ControlListEditForm.MotorID:= MotorIDs[SelectedIndex];
    ControlListEditForm.MotorName:= MotorNames[SelectedIndex];
    ControlListEditForm.MotorNum:= MotorNums[SelectedIndex];
    ControlListEditForm.ShowModal;
  finally
    FreeAndNil(ControlListEditForm);
  end;
end;

procedure TReclamationForm.DelButtonClick(Sender: TObject);
begin
  DeleteRaclamation;
end;

procedure TReclamationForm.EditButtonClick(Sender: TObject);
begin
  OpenReclamationEditForm(2);
end;

procedure TReclamationForm.AddButtonClick(Sender: TObject);
begin
  OpenReclamationEditForm(1);
end;

procedure TReclamationForm.ControlButtonClick(Sender: TObject);
begin
  OpenControlEditForm;
end;

procedure TReclamationForm.DateTimePicker1Change(Sender: TObject);
begin
  if not SettingClientPanel.Visible then Exit;
  ViewUpdate;
end;

procedure TReclamationForm.DateTimePicker2Change(Sender: TObject);
begin
  if not SettingClientPanel.Visible then Exit;
  ViewUpdate;
end;

procedure TReclamationForm.ExportButtonClick(Sender: TObject);
begin
  ExportReclamation;
end;

end.

