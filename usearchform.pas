unit USearchForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, rxctrls, DividerBevel, DBUtils, DK_Vector, LCLType, SheetUtils,
  fpspreadsheetgrid, DK_Dialogs, DK_StrUtils, DK_SheetExporter, fpstypes;

type

  { TSearchForm }

  TSearchForm = class(TForm)
    ListGrid: TsWorksheetGrid;
    CloseButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    DividerBevel5: TDividerBevel;
    ExportButton: TRxSpeedButton;
    FindButton: TSpeedButton;
    Label1: TLabel;
    LogGrid: TsWorksheetGrid;
    MotorNameComboBox: TComboBox;
    MotorNumEdit: TEdit;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Splitter1: TSplitter;
    procedure CloseButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FindButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);

    procedure ListGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure MotorNumEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    SelectedIndex: Integer;
    NameIDs, MotorIDs: TIntVector;
    MotorNames, Series: TStrVector;

    SeriesSheet: TSeriesSheet;
    MotorInfoSheet: TMotorInfoSheet;

    procedure LoadMotorNames;
    procedure FindMotor;
    procedure DataOpen;

    procedure SelectionClear;
    procedure SelectLine(const ARow: Integer);

    procedure ExportSheet;
  public

  end;

var
  SearchForm: TSearchForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TSearchForm }

procedure TSearchForm.SelectionClear;
begin
  if SelectedIndex>-1 then
  begin
    SeriesSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= -1;
  end;
  ExportButton.Enabled:= False;
  LogGrid.Clear;
end;

procedure TSearchForm.SelectLine(const ARow: Integer);
const
  FirstRow = 2;
begin
  if VIsNil(MotorNames) then Exit;
  if (ARow>=FirstRow) and (ARow<ListGrid.RowCount-1) then  //клик ниже заголовка
  begin
    if SelectedIndex>-1 then
      SeriesSheet.DrawLine(SelectedIndex, False);
    SelectedIndex:= ARow - FirstRow;
    SeriesSheet.DrawLine(SelectedIndex, True);
  end;
  ExportButton.Enabled:= SelectedIndex>-1;
  DataOpen;
end;

procedure TSearchForm.ExportSheet;
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(LogGrid);
  try
    //Exporter.SheetName:= 'Отчет';
    Exporter.PageSettings(spoLandscape, pfOnePage);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;

end;

procedure TSearchForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton6.Down:= False;
  MainForm.SearchForm:= nil;
  CloseAction:= caFree;
end;

procedure TSearchForm.FormCreate(Sender: TObject);
begin
  SelectedIndex:= -1;
  LoadMotorNames;
  SeriesSheet:= TSeriesSheet.Create(ListGrid);
  MotorInfoSheet:= TMotorInfoSheet.Create(LogGrid);
end;

procedure TSearchForm.FormDestroy(Sender: TObject);
begin
  if Assigned(SeriesSheet) then FreeAndNil(SeriesSheet);
  if Assigned(MotorInfoSheet) then FreeAndNil(MotorInfoSheet);
end;

procedure TSearchForm.FormShow(Sender: TObject);
begin
  MotorNumEdit.SetFocus;
end;



procedure TSearchForm.ListGridMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  R,C: Integer;
begin
  if Button=mbRight then
    SelectionClear;
  if Button=mbLeft  then
  begin
    ListGrid.MouseToCell(X,Y,C{%H-},R{%H-});
    SelectLine(R);
  end;
end;

procedure TSearchForm.MotorNumEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_RETURN then FindMotor;
end;

procedure TSearchForm.LoadMotorNames;
begin
  GetKeyPickList('MOTORNAMES', 'NameID', 'MotorName',
                 NameIDs, MotorNames, True, 'NameID');
  VToStrings(MotorNames, MotorNameComboBox.Items);
  if not VIsNil(MotorNames) then
    MotorNameComboBox.ItemIndex:= 0;
end;

procedure TSearchForm.FindMotor;
var
  S: String;
begin
  if MotorNameComboBox.Text=EmptyStr then
  begin
    ShowInfo('Не указано наименование двигателя!');
    Exit;
  end;

  S:= STrim(MotorNumEdit.Text);
  if S=EmptyStr then
  begin
    ShowInfo('Не указан номер двигателя!');
    Exit;
  end;

  SelectionClear;

  FindMotorsFromNameAndNum(NameIDs[MotorNameComboBox.ItemIndex], S,
                           MotorIDs, Series);
  SeriesSheet.Draw(Series);
end;

procedure TSearchForm.DataOpen;
var
  BuildDate, SendDate: TDate;
  MotorName, MotorNum, RotorNum, ReceiverName, Sers: String;
  TestDates: TDateVector;
  TestResults: TIntVector;
  TestNotes: TStrVector;
begin
  if SelectedIndex=-1 then Exit;
  MotorInfoLoad(MotorIDs[SelectedIndex], BuildDate, SendDate,
                MotorName, MotorNum, Sers, RotorNum, ReceiverName,
                TestDates, TestResults, TestNotes);
  MotorInfoSheet.Draw(BuildDate, SendDate, MotorName, MotorNum, Sers,
                      RotorNum, ReceiverName, TestDates, TestResults, TestNotes);
  LogGrid.Visible:= True;
end;

procedure TSearchForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TSearchForm.ExportButtonClick(Sender: TObject);
begin
  ExportSheet;
end;

procedure TSearchForm.FindButtonClick(Sender: TObject);
begin
  FindMotor;
end;

end.

