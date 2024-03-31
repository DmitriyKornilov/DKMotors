unit UReclamationForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, EditBtn, ComCtrls, VirtualTrees, fpspreadsheetgrid,
  DK_Vector, UUtils, USheetUtils, DK_DateUtils, UReclamationEditForm,
  DK_StrUtils, DK_Dialogs, DK_SheetExporter, USQLite, UCardForm, DK_VSTTables,
  URepairEditForm, DateTimePicker, UControlListEditForm, BCButton, DK_Zoom,
  DK_Const, DK_VSTTools;

type

  { TReclamationForm }

  TReclamationForm = class(TForm)
    AddButton: TSpeedButton;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Bevel5: TBevel;
    CardPanel: TPanel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    ExportButton: TBCButton;
    Label3: TLabel;
    LeftPanel: TPanel;
    LogGrid: TsWorksheetGrid;
    MainPanel: TPanel;
    Panel4: TPanel;
    RepairButton: TSpeedButton;
    MotorCardCheckBox: TCheckBox;
    Label2: TLabel;
    MotorNumEdit: TEditButton;
    Panel2: TPanel;
    DelButton: TSpeedButton;
    EditButton: TSpeedButton;
    Panel3: TPanel;
    Panel5: TPanel;
    Panel7: TPanel;
    ControlButton: TSpeedButton;
    ReportPeriodPanel: TPanel;
    Splitter0: TSplitter;
    Splitter1: TSplitter;
    ToolPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    VT3: TVirtualStringTree;
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
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure RepairButtonClick(Sender: TObject);
  private
    CardForm: TCardForm;
    VSTOrderList: TVSTStringList;
    VSTDefectList: TVSTCheckList;
    VSTReasonList: TVSTCheckTable;

    CanShow: Boolean;

    ZoomPercent: Integer;

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

    procedure CreateOrderList;
    procedure SelectOrder;

    procedure CreateDefectList;
    procedure CreateReasonList;
  public
    procedure ShowReclamation;
  end;

var
  ReclamationForm: TReclamationForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TReclamationForm }

procedure TReclamationForm.FormCreate(Sender: TObject);
begin
  SetToolPanels([
    ToolPanel
  ]);
  SetToolButtons([
    AddButton, DelButton, EditButton, RepairButton, ControlButton
  ]);

  CanShow:= False;
  CardForm:= CreateCardForm(ReclamationForm, CardPanel);
  MotorCardCheckBox.Checked:= False;
  MainForm.SetNamesPanelsVisible(True, False);
  SelectedIndex:= -1;

  ZoomPercent:= 100;
  CreateZoomControls(50, 150, ZoomPercent, ZoomPanel, @DrawReclamation, True);

  ReclamationSheet:= TReclamationSheet.Create(LogGrid.Worksheet, LogGrid);

  CreateOrderList;
  CreateDefectList;
  CreateReasonList;

  DateTimePicker2.Date:= FirstDayInYear(Date);
  DateTimePicker1.Date:= LastDayInYear(Date);

  CanShow:= True;
  ShowReclamation;
end;

procedure TReclamationForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CardForm) then FreeAndNil(CardForm);
  if Assigned(ReclamationSheet) then FReeAndNil(ReclamationSheet);
  if Assigned(VSTOrderList) then FreeAndNil(VSTOrderList);
  if Assigned(VSTDefectList) then FreeAndNil(VSTDefectList);
  if Assigned(VSTReasonList) then FreeAndNil(VSTReasonList);
end;

procedure TReclamationForm.FormShow(Sender: TObject);
begin
  ShowReclamation;
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

procedure TReclamationForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TReclamationForm.MotorNumEditChange(Sender: TObject);
begin
  ShowReclamation;
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

procedure TReclamationForm.ShowReclamation;
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
  if VSTOrderList.IsSelected then
    OrderIndex:= VSTOrderList.SelectedIndex;

  UsedReasonIDs:= VCut(ListReasonIDs, VSTReasonList.Selected);
  UsedDefectIDs:= VCut(ListDefectIDs, VSTDefectList.Selected);

  SQLite.ReclamationListLoad(DateTimePicker2.Date, DateTimePicker1.Date,
                        MainForm.UsedNameIDs, UsedDefectIDs, UsedReasonIDs,
                        OrderIndex, STrim(MotorNumEdit.Text),
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
    Drawer:= TReclamationSheet.Create(Sheet);
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

procedure TReclamationForm.CreateOrderList;
var
  S: String;
  V: TStrVector;
begin
  S:= 'Упорядочить по:';
  V:= VCreateStr([
    'дате уведомления',
    'номеру двигателя',
    'дате сборки',
    'предприятию',
    'неисправному элементу',
    'причине неисправности'
  ]);
  VSTOrderList:= TVSTStringList.Create(VT1, S, @SelectOrder);
  VSTOrderList.Update(V);
end;

procedure TReclamationForm.CreateDefectList;
var
  S: String;
  V: TStrVector;
  i: Integer;
begin
  S:= 'Неисправный элемент:';
  SQLite.KeyPickList('RECLAMATIONDEFECTS', 'DefectID', 'DefectName',
                     ListDefectIDs, V, False, 'DefectName');
  for i:= 0 to High(V) do
    V[i]:= SFirstLower(V[i]);
  VSTDefectList:= TVSTCheckList.Create(VT2, S, V, @ShowReclamation);
end;

procedure TReclamationForm.CreateReasonList;
var
  S: String;
  V: TStrVector;
  i: Integer;
begin
  S:= 'Причина неисправности:';
  SQLite.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
                     ListReasonIDs, V);
  for i:= 0 to High(V) do
    V[i]:= SFirstLower(V[i]);
  VSTReasonList:= TVSTCheckList.Create(VT3, S, V, @ShowReclamation);
end;

procedure TReclamationForm.SelectOrder;
begin
  ShowReclamation;
end;

procedure TReclamationForm.DeleteRaclamation;
begin
  if SelectedIndex<0 then Exit;
  if not Confirm('Удалить рекламацию?') then Exit;
  SQLite.Delete('RECLAMATIONS', 'RecID', RecIDs[SelectedIndex]);
  ClearSelection;
  ShowReclamation;
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
    if ReclamationEditForm.ShowModal = mrOK then ShowReclamation;
  finally
    FreeAndNil(ReclamationEditForm);
  end;
  ClearSelection;
  ShowReclamation;
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
    if RepairEditForm.ShowModal = mrOK then ShowReclamation;
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
  if not LeftPanel.Visible then Exit;
  ShowReclamation;
end;

procedure TReclamationForm.DateTimePicker2Change(Sender: TObject);
begin
  if not LeftPanel.Visible then Exit;
  ShowReclamation;
end;

procedure TReclamationForm.ExportButtonClick(Sender: TObject);
begin
  ExportReclamation;
end;

end.

