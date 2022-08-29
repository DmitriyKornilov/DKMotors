unit UStoreForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, DividerBevel, fpspreadsheetgrid, SheetUtils, USQLite, rxctrls,
  DK_Vector, DK_SheetExporter, fpstypes;

type

  { TStoreForm }

  TStoreForm = class(TForm)
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CloseButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    DividerBevel7: TDividerBevel;
    DividerBevel8: TDividerBevel;
    ExportButton: TRxSpeedButton;
    MotorNameComboBox: TComboBox;
    Panel2: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel1: TPanel;
    ReportGrid: TsWorksheetGrid;
    procedure CheckBox1Click(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MotorNameComboBoxChange(Sender: TObject);
  private
    NameIDs: TIntVector;
    StoreSheet: TStoreSheet;
    procedure DataOpen;
    procedure ExportSheet;

  public

  end;

var
  StoreForm: TStoreForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TStoreForm }

procedure TStoreForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton4.Down:= False;
  MainForm.StoreForm:= nil;
  CloseAction:= caFree;
end;

procedure TStoreForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TStoreForm.CheckBox1Click(Sender: TObject);
begin
  DataOpen;
end;

procedure TStoreForm.ExportButtonClick(Sender: TObject);
begin
  ExportSheet;
end;

procedure TStoreForm.FormCreate(Sender: TObject);
begin
  StoreSheet:= TStoreSheet.Create(ReportGrid);
  SQLite.NameIDsAndMotorNamesLoad(MotorNameComboBox, NameIDs, False);
end;

procedure TStoreForm.FormDestroy(Sender: TObject);
begin
  if Assigned(StoreSheet) then FreeAndNil(StoreSheet);
end;

procedure TStoreForm.FormShow(Sender: TObject);
begin
  DataOpen;
end;

procedure TStoreForm.MotorNameComboBoxChange(Sender: TObject);
begin
  Checkbox2.Visible:= MotorNameComboBox.ItemIndex=0;
  DataOpen;
end;

procedure TStoreForm.DataOpen;
var
  TotalMotorNames, MotorNames, MotorNums: TStrVector;
  TotalMotorCounts: TIntVector;
  TestDates: TDateVector;
begin
  Screen.Cursor:= crHourGlass;
  try

    SQLite.StoreListLoad(NameIDs[MotorNameComboBox.ItemIndex],
                       Checkbox1.Checked, Checkbox2.Checked,
                       TestDates, MotorNames, MotorNums);
    SQLite.StoreTotalLoad(NameIDs[MotorNameComboBox.ItemIndex],
                        TotalMotorNames, TotalMotorCounts);
    StoreSheet.Draw(TestDates, MotorNames, MotorNums,
                    TotalMotorNames, TotalMotorCounts);

  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TStoreForm.ExportSheet;
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(ReportGrid);
  try
    //Exporter.SheetName:= 'Отчет';
    Exporter.PageSettings(spoPortrait, pfWidth);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;



end.

