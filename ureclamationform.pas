unit UReclamationForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, EditBtn, ComCtrls, VirtualTrees, fpspreadsheetgrid, rxctrls,
  DK_Vector, DividerBevel, USheetUtils, DK_DateUtils, UReclamationEditForm,
  DK_StrUtils, DK_Dialogs, DK_SheetExporter, USQLite, UCardForm, DK_VSTTables,
  URepairEditForm, DateTimePicker;

type

  { TReclamationForm }

  TReclamationForm = class(TForm)
    AddButton: TSpeedButton;
    CardPanel: TPanel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label3: TLabel;
    Label5: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    LeftPanel: TPanel;
    LogGrid: TsWorksheetGrid;
    MainPanel: TPanel;
    Panel4: TPanel;
    Panel6: TPanel;
    Panel8: TPanel;
    RepairButton: TSpeedButton;
    MotorCardCheckBox: TCheckBox;
    DefectListButton: TRxSpeedButton;
    DividerBevel8: TDividerBevel;
    DividerBevel9: TDividerBevel;
    FactoryListButton: TRxSpeedButton;
    Label2: TLabel;
    MotorNumEdit: TEditButton;
    Panel2: TPanel;
    DelButton: TSpeedButton;
    DividerBevel5: TDividerBevel;
    DividerBevel7: TDividerBevel;
    EditButton: TSpeedButton;
    DividerBevel6: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Panel3: TPanel;
    Panel5: TPanel;
    Panel7: TPanel;
    PlaceListButton: TRxSpeedButton;
    ReasonListButton: TRxSpeedButton;
    ReportPeriodPanel: TPanel;
    Splitter0: TSplitter;
    Splitter1: TSplitter;
    TopToolsPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    VT3: TVirtualStringTree;
    ZoomCaptionLabel: TLabel;
    ZoomInButton: TSpeedButton;
    ZoomOutButton: TSpeedButton;
    ZoomPanel: TPanel;
    ZoomTrackBar: TTrackBar;
    ZoomValueLabel: TLabel;
    ZoomValuePanel: TPanel;
    procedure AddButtonClick(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FactoryListButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LogGridMouseDown(Sender: TObject; Button: TMouseButton;
      {%H-}Shift: TShiftState; X, Y: Integer);
    procedure MotorCardCheckBoxChange(Sender: TObject);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure PlaceListButtonClick(Sender: TObject);
    procedure DefectListButtonClick(Sender: TObject);
    procedure ReasonListButtonClick(Sender: TObject);
    procedure RepairButtonClick(Sender: TObject);
    procedure ZoomInButtonClick(Sender: TObject);
    procedure ZoomOutButtonClick(Sender: TObject);
    procedure ZoomTrackBarChange(Sender: TObject);
  private
    CardForm: TCardForm;
    VSTOrderList: TVSTTable;
    VSTDefectList: TVSTCheckTable;
    VSTReasonList: TVSTCheckTable;

    CanShow: Boolean;

    SelectedIndex: Integer;
    ReclamationSheet: TReclamationSheet;

    RecDates, BuildDates, ArrivalDates, SendingDates: TDateVector;
    RecIDs, MotorIDs, Mileages, Opinions, ReasonColors, Passports: TIntVector;
    PlaceNames, FactoryNames, Departures: TStrVector;
    DefectNames, ReasonNames, RecNotes: TStrVector;
    MotorNames, MotorNums: TStrVector;

    ListReasonIDs: TIntVector;
    ListDefectIDs: TIntVector;


    procedure SelectionClear;
    procedure SelectLine(const ARow: Integer);

    procedure DelRaclamation;
    procedure ReclamationEditFormOpen(const AEditType: Byte);

    procedure RepairEditFormOpen;

    procedure LoadReclamation;
    procedure DrawReclamation;
    procedure ExportReclamation;

    procedure CreateOrderList;
    procedure CreateDefectList;
    procedure CreateReasonList;

    procedure OrderSelect;
    procedure DefectSelect(const {%H-}AIndex: Integer; const {%H-}AChecked: Boolean);
    procedure ReasonSelect(const {%H-}AIndex: Integer; const {%H-}AChecked: Boolean);
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
  CanShow:= False;
  CardForm:= CreateCardForm(ReclamationForm, CardPanel);
  MotorCardCheckBox.Checked:= False;
  MainForm.SetNamesPanelsVisible(True, False);
  SelectedIndex:= -1;
  ReclamationSheet:= TReclamationSheet.Create(LogGrid.Worksheet, LogGrid);

  CreateOrderList;
  CreateDefectList;
  CreateReasonList;

  DateTimePicker2.Date:= FirstDayInYear(Date);
  CanShow:= True;
  DateTimePicker1.Date:= LastDayInYear(Date);
end;

procedure TReclamationForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CardForm) then FreeAndNil(CardForm);
  if Assigned(ReclamationSheet) then FReeAndNil(ReclamationSheet);
  if Assigned(VSTOrderList) then FreeAndNil(VSTOrderList);
end;

procedure TReclamationForm.FormShow(Sender: TObject);
begin
  ShowReclamation;
end;

procedure TReclamationForm.LogGridMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  R,C, MotorID: Integer;
begin
  if Button=mbRight then
    SelectionClear;
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

procedure TReclamationForm.PlaceListButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Предприятия (депо)',
    'RECLAMATIONPLACES', 'PlaceID', 'PlaceName', True, True) then
      ShowReclamation;
end;

procedure TReclamationForm.FactoryListButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Заводы',
    'RECLAMATIONFACTORIES', 'FactoryID', 'FactoryName', True, True) then
      ShowReclamation;
end;

procedure TReclamationForm.DefectListButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Неисправные элементы',
    'RECLAMATIONDEFECTS', 'DefectID', 'DefectName', True, True) then
      ShowReclamation;
end;

procedure TReclamationForm.ReasonListButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Причины возникновения неисправностей',
    'RECLAMATIONREASONS', 'ReasonID', 'ReasonName', True, True, 'ReasonColor') then
      ShowReclamation;
end;

procedure TReclamationForm.RepairButtonClick(Sender: TObject);
begin
  RepairEditFormOpen;
end;

procedure TReclamationForm.ZoomInButtonClick(Sender: TObject);
begin
  ZoomTrackBar.Position:= ZoomTrackBar.Position + 5;
end;

procedure TReclamationForm.ZoomOutButtonClick(Sender: TObject);
begin
  ZoomTrackBar.Position:= ZoomTrackBar.Position - 5;
end;

procedure TReclamationForm.ZoomTrackBarChange(Sender: TObject);
begin
  ZoomValueLabel.Caption:= IntToStr(ZoomTrackBar.Position) + ' %';
  DrawReclamation;
end;

procedure TReclamationForm.SelectionClear;
begin
  if SelectedIndex>-1 then
  begin
    ReclamationSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= -1;
  end;
  DelButton.Enabled:= False;
  EditButton.Enabled:= False;
  RepairButton.Enabled:= False;
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
  end;
end;

procedure TReclamationForm.ShowReclamation;
begin
  if not CanShow then Exit;
  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);
    SelectionClear;
    LoadReclamation;
    DrawReclamation;
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

procedure TReclamationForm.DrawReclamation;
begin
  ReclamationSheet.Zoom(ZoomTrackBar.Position);
  ReclamationSheet.Draw(RecDates, BuildDates, ArrivalDates, SendingDates,
                          Mileages, Opinions, ReasonColors,
                          PlaceNames, FactoryNames, Departures, DefectNames,
                          ReasonNames, RecNotes, MotorNames, MotorNums);
end;

procedure TReclamationForm.ExportReclamation;
var
  Drawer: TReclamationSheet;
  Sheet: TsWorksheet;
  Exporter: TSheetExporter;
begin
  Exporter:= TSheetExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    Drawer:= TReclamationSheet.Create(Sheet);
    try
      Drawer.Draw(RecDates, BuildDates, ArrivalDates, SendingDates,
                          Mileages, Opinions, ReasonColors,
                          PlaceNames, FactoryNames, Departures, DefectNames,
                          ReasonNames, RecNotes, MotorNames, MotorNums);
    finally
      FreeAndNil(Drawer);
    end;
    Exporter.PageSettings(spoLandscape);

    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TReclamationForm.CreateOrderList;
var
  V: TStrVector;
begin
  VSTOrderList:= TVSTTable.Create(VT1);
  VSTOrderList.OnSelect:= @OrderSelect;
  VSTOrderList.AutoHeight:= True;
  VSTOrderList.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTOrderList.HeaderVisible:= False;
  VSTOrderList.GridLinesVisible:= False;
  VSTOrderList.CanSelect:= True;
  VSTOrderList.CanUnselect:= False;
  VSTOrderList.AddColumn('Список');
  V:= VCreateStr(['дате уведомления', 'номеру двигателя', 'дате сборки',
                  'предприятию', 'неисправному элементу', 'причине неисправности']);
  VSTOrderList.SetColumn('Список', V, taLeftJustify);
  VSTOrderList.Draw;
  VSTOrderList.Select(0);
end;

procedure TReclamationForm.CreateDefectList;
var
  V: TStrVector;
  i: Integer;
begin
  VSTDefectList:= TVSTCheckTable.Create(VT2);
  VSTDefectList.OnSelect:= @DefectSelect;
  VSTDefectList.AutoHeight:= True;
  SQLite.KeyPickList('RECLAMATIONDEFECTS', 'DefectID', 'DefectName',
                     ListDefectIDs, V, False, 'DefectName');
  for i:= 0 to High(V) do
    V[i]:= SFirstLower(V[i]);

  VSTDefectList.GridLinesVisible:= False;
  VSTDefectList.HeaderVisible:= False;
  VSTDefectList.SelectedBGColor:= VT2.Color;
  VSTDefectList.AddColumn('Список');
  VSTDefectList.SetColumn('Список', V, taLeftJustify);
  VSTDefectList.Draw;
  VSTDefectList.CheckAll(True);
end;

procedure TReclamationForm.CreateReasonList;
var
  V: TStrVector;
  i: Integer;
begin
  VSTReasonList:= TVSTCheckTable.Create(VT3);
  VSTReasonList.OnSelect:= @ReasonSelect;
  VSTReasonList.AutoHeight:= True;
  SQLite.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
                     ListReasonIDs, V);
  for i:= 0 to High(V) do
    V[i]:= SFirstLower(V[i]);

  VSTReasonList.GridLinesVisible:= False;
  VSTReasonList.HeaderVisible:= False;
  VSTReasonList.SelectedBGColor:= VT3.Color;
  VSTReasonList.AddColumn('Список');
  VSTReasonList.SetColumn('Список', V, taLeftJustify);
  VSTReasonList.Draw;
  VSTReasonList.CheckAll(True);
end;

procedure TReclamationForm.OrderSelect;
begin
  ShowReclamation;
end;

procedure TReclamationForm.DefectSelect(const AIndex: Integer; const AChecked: Boolean);
begin
  ShowReclamation;
end;

procedure TReclamationForm.ReasonSelect(const AIndex: Integer; const AChecked: Boolean);
begin
  ShowReclamation;
end;

procedure TReclamationForm.DelRaclamation;
begin
  if SelectedIndex<0 then Exit;
  if not Confirm('Удалить рекламацию?') then Exit;
  SQLite.Delete('RECLAMATIONS', 'RecID', RecIDs[SelectedIndex]);
  SelectionClear;
  ShowReclamation;
end;

procedure TReclamationForm.ReclamationEditFormOpen(const AEditType: Byte);
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
  SelectionClear;
  ShowReclamation;
end;

procedure TReclamationForm.RepairEditFormOpen;
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

procedure TReclamationForm.DelButtonClick(Sender: TObject);
begin
  DelRaclamation;
end;

procedure TReclamationForm.EditButtonClick(Sender: TObject);
begin
  ReclamationEditFormOpen(2);
end;

procedure TReclamationForm.AddButtonClick(Sender: TObject);
begin
  ReclamationEditFormOpen(1);
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

