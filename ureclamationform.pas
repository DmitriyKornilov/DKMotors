unit UReclamationForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, Spin, EditBtn, fpspreadsheetgrid, rxctrls,
  DK_Vector, DividerBevel, SheetUtils, DK_DateUtils, UReclamationEditForm,
  DK_StrUtils, DK_Dialogs, DK_SheetExporter, fpstypes, USQLite;

type

  { TReclamationForm }

  TReclamationForm = class(TForm)
    AddButton: TSpeedButton;
    MotorRepairCheckBox: TCheckBox;
    DefectListButton: TRxSpeedButton;
    DividerBevel8: TDividerBevel;
    DividerBevel9: TDividerBevel;
    FactoryListButton: TRxSpeedButton;
    Label2: TLabel;
    LogGrid: TsWorksheetGrid;
    MotorNumEdit: TEditButton;
    Panel2: TPanel;
    Panel4: TPanel;
    DelButton: TSpeedButton;
    DividerBevel5: TDividerBevel;
    DividerBevel7: TDividerBevel;
    EditButton: TSpeedButton;
    DividerBevel6: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Panel1: TPanel;
    Panel3: TPanel;
    Panel5: TPanel;
    Panel7: TPanel;
    PlaceListButton: TRxSpeedButton;
    ReasonListButton: TRxSpeedButton;
    SpinEdit1: TSpinEdit;
    TopToolsPanel: TPanel;
    procedure AddButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FactoryListButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LogGridMouseDown(Sender: TObject; Button: TMouseButton;
      {%H-}Shift: TShiftState; X, Y: Integer);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure MotorRepairCheckBoxChange(Sender: TObject);
    procedure PlaceListButtonClick(Sender: TObject);
    procedure DefectListButtonClick(Sender: TObject);
    procedure ReasonListButtonClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    SelectedIndex: Integer;
    ReclamationSheet: TReclamationSheet;

    RecDates, BuildDates, ArrivalDates, SendingDates: TDateVector;
    RecIDs, MotorIDs, Mileages, Opinions, ReasonColors, Passports: TIntVector;
    PlaceNames, FactoryNames, Departures: TStrVector;
    DefectNames, ReasonNames, RecNotes: TStrVector;
    MotorNames, MotorNums: TStrVector;

    procedure ExportSheet;

    procedure SelectionClear;
    procedure SelectLine(const ARow: Integer);

    procedure DelRaclamation;
    procedure ReclamationEditFormOpen(const AEditType: Byte);
  public
    procedure ShowReclamation;
  end;

var
  ReclamationForm: TReclamationForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TReclamationForm }

procedure TReclamationForm.ExportSheet;
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(LogGrid);
  try
    //Exporter.SheetName:= 'Отчет';
    Exporter.PageSettings(spoLandscape, pfWidth);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TReclamationForm.FormCreate(Sender: TObject);
begin
  MainForm.SetNamesPanelsVisible(True, False);
  SelectedIndex:= -1;
  ReclamationSheet:= TReclamationSheet.Create(LogGrid);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TReclamationForm.FormDestroy(Sender: TObject);
begin
  if Assigned(ReclamationSheet) then FReeAndNil(ReclamationSheet);
end;

procedure TReclamationForm.FormShow(Sender: TObject);
begin
  ShowReclamation;
end;

procedure TReclamationForm.LogGridMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  R,C: Integer;
begin
  if Button=mbRight then
    SelectionClear;
  if Button=mbLeft  then
  begin
    LogGrid.MouseToCell(X,Y,C{%H-},R{%H-});
    SelectLine(R);
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

procedure TReclamationForm.MotorRepairCheckBoxChange(Sender: TObject);
begin
  SpinEdit1.Enabled:= not MotorRepairCheckBox.Checked;
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

procedure TReclamationForm.SpinEdit1Change(Sender: TObject);
begin
  ShowReclamation;
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
  end;
end;

procedure TReclamationForm.ShowReclamation;
var
  BeginDate, EndDate: TDate;
begin
  Screen.Cursor:= crHourGlass;
  try
    SelectionClear;
    LogGrid.Clear;
    BeginDate:= FirstDayInYear(SpinEdit1.Value);
    EndDate:= LastDayInYear(SpinEdit1.Value);

    SQLite.ReclamationListLoad(BeginDate, EndDate,
                        MainForm.UsedNameIDs,
                        STrim(MotorNumEdit.Text),
                        MotorRepairCheckBox.Checked,
                        RecDates, BuildDates, ArrivalDates, SendingDates,
                        RecIDs, MotorIDs, Mileages, Opinions,
                        ReasonColors, Passports,
                        PlaceNames, FactoryNames, Departures,
                        DefectNames, ReasonNames, RecNotes,
                        MotorNames, MotorNums);

    ReclamationSheet.Draw(RecDates, BuildDates, ArrivalDates, SendingDates,
                          Mileages, Opinions, ReasonColors, Passports,
                          PlaceNames, FactoryNames, Departures, DefectNames,
                          ReasonNames, RecNotes, MotorNames, MotorNums);



  finally
    Screen.Cursor:= crDefault;
  end;
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

procedure TReclamationForm.ExportButtonClick(Sender: TObject);
begin
  ExportSheet;
end;

end.

