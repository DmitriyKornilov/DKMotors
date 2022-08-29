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
    DefectListButton: TRxSpeedButton;
    DividerBevel8: TDividerBevel;
    DividerBevel9: TDividerBevel;
    FactoryListButton: TRxSpeedButton;
    Label2: TLabel;
    LogGrid: TsWorksheetGrid;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEditButton;
    Panel4: TPanel;
    DelButton: TSpeedButton;
    DividerBevel5: TDividerBevel;
    DividerBevel7: TDividerBevel;
    EditButton: TSpeedButton;
    CloseButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    DividerBevel6: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel5: TPanel;
    Panel7: TPanel;
    PlaceListButton: TRxSpeedButton;
    ReasonListButton: TRxSpeedButton;
    SpinEdit1: TSpinEdit;
    TopToolsPanel: TPanel;
    procedure AddButtonClick(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FactoryListButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LogGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure MotorNameComboBoxChange(Sender: TObject);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure PlaceListButtonClick(Sender: TObject);


    procedure DefectListButtonClick(Sender: TObject);
    procedure ReasonListButtonClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    SelectedIndex: Integer;
    ReclamationSheet: TReclamationSheet;

    RecDates, BuildDates, ArrivalDates, SendingDates: TDateVector;
    RecIDs, MotorIDs, Mileages, Opinions, ReasonColors, NameIDs, Passports: TIntVector;
    PlaceNames, FactoryNames, Departures: TStrVector;
    DefectNames, ReasonNames, RecNotes: TStrVector;
    MotorNames, MotorNums: TStrVector;

    procedure ExportSheet;

    procedure SelectionClear;
    procedure SelectLine(const ARow: Integer);

    procedure DataOpen;

    procedure DelRaclamation;

    procedure ReclamationEditFormOpen(const AEditType: Byte);


  public

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

procedure TReclamationForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton3.Down:= False;
  MainForm.ReclamationForm:= nil;
  CloseAction:= caFree;
end;

procedure TReclamationForm.FormCreate(Sender: TObject);
begin
  SelectedIndex:= -1;
  ReclamationSheet:= TReclamationSheet.Create(LogGrid);
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs, False);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TReclamationForm.FormDestroy(Sender: TObject);
begin
  if Assigned(ReclamationSheet) then FReeAndNil(ReclamationSheet);
end;

procedure TReclamationForm.FormShow(Sender: TObject);
begin
  DataOpen;
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

procedure TReclamationForm.MotorNameComboBoxChange(Sender: TObject);
begin
  DataOpen;
end;

procedure TReclamationForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TReclamationForm.MotorNumEditChange(Sender: TObject);
begin
  DataOpen;
end;

procedure TReclamationForm.PlaceListButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Предприятия (депо)',
    'RECLAMATIONPLACES', 'PlaceID', 'PlaceName', True, True) then
      DataOpen;
end;

procedure TReclamationForm.FactoryListButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Заводы',
    'RECLAMATIONFACTORIES', 'FactoryID', 'FactoryName', True, True) then
      DataOpen;
end;

procedure TReclamationForm.DefectListButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Заводы',
    'RECLAMATIONDEFECTS', 'DefectID', 'DefectName', True, True) then
      DataOpen;
end;

procedure TReclamationForm.ReasonListButtonClick(Sender: TObject);
begin
  if SQLite.EditList('Причины возникновения неисправностей',
    'RECLAMATIONREASONS', 'ReasonID', 'ReasonName', True, True, 'ReasonColor') then
      DataOpen;
end;

procedure TReclamationForm.SpinEdit1Change(Sender: TObject);
begin
  DataOpen;
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

procedure TReclamationForm.DataOpen;
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
                        NameIDs[MotorNameComboBox.ItemIndex],
                        STrim(MotorNumEdit.Text),
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
  DataOpen;
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
    if ReclamationEditForm.ShowModal = mrOK then DataOpen;
  finally
    FreeAndNil(ReclamationEditForm);
  end;
  SelectionClear;
  DataOpen;
end;

procedure TReclamationForm.CloseButtonClick(Sender: TObject);
begin
  Close;
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

