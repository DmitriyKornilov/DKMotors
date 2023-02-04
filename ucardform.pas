unit UCardForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, ComCtrls,
  StdCtrls, Buttons, rxctrls, fpspreadsheetgrid, USQLite, USheetUtils,
  DK_Vector, DK_SheetExporter;

type

  { TCardForm }

  TCardForm = class(TForm)
    CardPanel: TPanel;
    ExportButton: TRxSpeedButton;
    CardGrid: TsWorksheetGrid;
    Panel6: TPanel;
    TopPanel: TPanel;
    ZoomCaptionLabel: TLabel;
    ZoomInButton: TSpeedButton;
    ZoomOutButton: TSpeedButton;
    ZoomPanel: TPanel;
    ZoomTrackBar: TTrackBar;
    ZoomValueLabel: TLabel;
    ZoomValuePanel: TPanel;
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ZoomInButtonClick(Sender: TObject);
    procedure ZoomOutButtonClick(Sender: TObject);
    procedure ZoomTrackBarChange(Sender: TObject);
  private
    MotorCardSheet: TMotorCardSheet;
    MotorID: Integer;
    procedure DrawCard;
  public
    procedure ShowCard(const AMotorID: Integer);
  end;

  function CrateCardForm(AOwner: TComponent; AParent: TWinControl): TCardForm;

var
  CardForm: TCardForm;

implementation

{$R *.lfm}

function CrateCardForm(AOwner: TComponent; AParent: TWinControl): TCardForm;
begin
  Result:= TCardForm.Create(AOwner);
  Result.Parent:= AParent;
  Result.Left:= 0;
  Result.Top:= 0;
  Result.MakeFullyVisible();
  Result.Show;
end;

{ TCardForm }

procedure TCardForm.FormCreate(Sender: TObject);
begin
  MotorID:= 0;
  MotorCardSheet:= TMotorCardSheet.Create(CardGrid);
end;

procedure TCardForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(CardGrid);
  try
    Exporter.PageSettings(spoLandscape, pfOnePage);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TCardForm.FormDestroy(Sender: TObject);
begin
  if Assigned(MotorCardSheet) then FreeAndNil(MotorCardSheet);
end;

procedure TCardForm.ZoomInButtonClick(Sender: TObject);
begin
  ZoomTrackBar.Position:= ZoomTrackBar.Position + 5;
end;

procedure TCardForm.ZoomOutButtonClick(Sender: TObject);
begin
  ZoomTrackBar.Position:= ZoomTrackBar.Position - 5;
end;

procedure TCardForm.ZoomTrackBarChange(Sender: TObject);
begin
  DrawCard;
end;

procedure TCardForm.DrawCard;
var
  BuildDate, SendDate: TDate;
  MotorName, MotorNum, RotorNum, ReceiverName, Sers, ControlNote: String;
  TestDates, RecDates, ArrivalDates, SendingDates: TDateVector;
  TestResults, Mileages, Opinions, Passports, WorkDayCounts: TIntVector;
  TestNotes, PlaceNames, FactoryNames, Departures,
  DefectNames, ReasonNames, RecNotes, RepairNotes: TStrVector;
begin
  SQLite.MotorInfoLoad(MotorID, BuildDate, SendDate,
                MotorName, MotorNum, Sers, RotorNum, ReceiverName,
                TestDates, TestResults, TestNotes);
  ControlNote:= SQLite.ControlNoteLoad(MotorID);

  SQLite.ReclamationListLoad(MotorID, RecDates, Mileages, Opinions,
                      PlaceNames, FactoryNames, Departures,
                      DefectNames, ReasonNames, RecNotes);

  SQLite.RepairListForMotorIDLoad(MotorID, ArrivalDates, SendingDates,
                      Passports, WorkDayCounts, RepairNotes);

  MotorCardSheet.Zoom(ZoomTrackBar.Position);
  MotorCardSheet.Draw(BuildDate, SendDate, MotorName, MotorNum, Sers,
                      RotorNum, ReceiverName, ControlNote, TestDates, TestResults, TestNotes,
                      RecDates, Mileages, Opinions, PlaceNames, FactoryNames,
                      Departures, DefectNames, ReasonNames, RecNotes,
                      ArrivalDates, SendingDates, Passports, WorkDayCounts, RepairNotes);
end;

procedure TCardForm.ShowCard(const AMotorID: Integer);
begin
  CardGrid.Clear;
  MotorID:= AMotorID;
  if MotorID<=0 then Exit;
  DrawCard;
end;

end.

