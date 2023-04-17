unit UCardForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, ComCtrls,
  Buttons, rxctrls, fpspreadsheetgrid, USQLite, USheetUtils,
  DK_Vector, DK_SheetExporter, DK_Zoom;

type

  { TCardForm }

  TCardForm = class(TForm)
    CardPanel: TPanel;
    ExportButton: TRxSpeedButton;
    CardGrid: TsWorksheetGrid;
    TopPanel: TPanel;
    ZoomPanel: TPanel;
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    MotorCardSheet: TMotorCardSheet;
    MotorID: Integer;
    ZoomPercent: Integer;
    BuildDate, SendDate: TDate;
    MotorName, MotorNum, RotorNum, ReceiverName, Sers, ControlNote: String;
    TestDates, RecDates, ArrivalDates, SendingDates: TDateVector;
    TestResults, Mileages, Opinions, Passports, WorkDayCounts: TIntVector;
    TestNotes, PlaceNames, FactoryNames, Departures,
    DefectNames, ReasonNames, RecNotes, RepairNotes: TStrVector;

    procedure DrawCard(const AZoomPercent: Integer);
    procedure LoadCard;
    procedure ExportCard;
  public
    procedure ShowCard(const AMotorID: Integer);
  end;

  function CreateCardForm(AOwner: TComponent; AParent: TWinControl): TCardForm;

var
  CardForm: TCardForm;

implementation

{$R *.lfm}

function CreateCardForm(AOwner: TComponent; AParent: TWinControl): TCardForm;
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
  ZoomPercent:= 100;
  CreateZoomControls(50, 150, ZoomPercent, ZoomPanel, @DrawCard, True);
  MotorCardSheet:= TMotorCardSheet.Create(CardGrid.Worksheet, CardGrid);
end;

procedure TCardForm.ExportButtonClick(Sender: TObject);
begin
  ExportCard;
end;

procedure TCardForm.FormDestroy(Sender: TObject);
begin
  if Assigned(MotorCardSheet) then FreeAndNil(MotorCardSheet);
end;

procedure TCardForm.DrawCard(const AZoomPercent: Integer);
begin
  if MotorID<=0 then Exit;
  ZoomPercent:= AZoomPercent;
  MotorCardSheet.Zoom(ZoomPercent);
  MotorCardSheet.Draw(BuildDate, SendDate, MotorName, MotorNum, Sers,
                      RotorNum, ReceiverName, ControlNote, TestDates, TestResults, TestNotes,
                      RecDates, Mileages, Opinions, PlaceNames, FactoryNames,
                      Departures, DefectNames, ReasonNames, RecNotes,
                      ArrivalDates, SendingDates, Passports, WorkDayCounts, RepairNotes);
end;

procedure TCardForm.LoadCard;
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
end;

procedure TCardForm.ExportCard;
var
  Drawer: TMotorCardSheet;
  Sheet: TsWorksheet;
  Exporter: TSheetExporter;
begin
  Exporter:= TSheetExporter.Create;
  try
    Sheet:= Exporter.AddWorksheet('Лист1');
    Drawer:= TMotorCardSheet.Create(Sheet);
    try
      Drawer.Draw(BuildDate, SendDate, MotorName, MotorNum, Sers,
                RotorNum, ReceiverName, ControlNote, TestDates, TestResults, TestNotes,
                RecDates, Mileages, Opinions, PlaceNames, FactoryNames,
                Departures, DefectNames, ReasonNames, RecNotes,
                ArrivalDates, SendingDates, Passports, WorkDayCounts, RepairNotes);

    finally
      FreeAndNil(Drawer);
    end;
    Exporter.PageSettings(spoLandscape);

    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

procedure TCardForm.ShowCard(const AMotorID: Integer);
begin
  CardGrid.Clear;
  MotorID:= AMotorID;
  if MotorID<=0 then Exit;
  LoadCard;
  DrawCard(ZoomPercent);
end;

end.

