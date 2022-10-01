unit UStoreForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, Spin, DividerBevel, fpspreadsheetgrid, SheetUtils, USQLite, rxctrls,
  DK_Vector, DK_SheetExporter, fpstypes;

type

  { TStoreForm }

  TStoreForm = class(TForm)
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    ChooseMotorNamesButton: TSpeedButton;
    CloseButton: TSpeedButton;
    DividerBevel4: TDividerBevel;
    DividerBevel7: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    MotorNamesLabel: TLabel;
    MotorNamesPanel: TPanel;
    Panel2: TPanel;
    Panel5: TPanel;
    Panel1: TPanel;
    ReportGrid: TsWorksheetGrid;
    SpinEdit1: TSpinEdit;
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure ChooseMotorNamesButtonClick(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);

  private
    UsedNameIDs: TIntVector;
    UsedNames: TStrVector;
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

procedure TStoreForm.CheckBox2Click(Sender: TObject);
begin
  DataOpen;
end;

procedure TStoreForm.CheckBox3Click(Sender: TObject);
begin
  SpinEdit1.Enabled:= CheckBox3.Checked;
  DataOpen;
end;

procedure TStoreForm.ExportButtonClick(Sender: TObject);
begin
  ExportSheet;
end;

procedure TStoreForm.FormCreate(Sender: TObject);
begin
  StoreSheet:= TStoreSheet.Create(ReportGrid);

  SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, False, UsedNameIDs, UsedNames);

  Checkbox2.Visible:= Length(UsedNameIDs)<>1;
end;

procedure TStoreForm.ChooseMotorNamesButtonClick(Sender: TObject);
begin
 if SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, True, UsedNameIDs, UsedNames) then
    DataOpen;
  Checkbox2.Visible:= Length(UsedNameIDs)<>1;
end;

procedure TStoreForm.FormDestroy(Sender: TObject);
begin
  if Assigned(StoreSheet) then FreeAndNil(StoreSheet);
end;

procedure TStoreForm.FormShow(Sender: TObject);
begin
  DataOpen;
end;

procedure TStoreForm.SpinEdit1Change(Sender: TObject);
begin
  DataOpen;
end;

procedure TStoreForm.DataOpen;
var
  TotalMotorNames, MotorNames, MotorNums: TStrVector;
  TotalMotorCounts: TIntVector;
  TestDates: TDateVector;
  DeltaDays: Integer;
begin
  Screen.Cursor:= crHourGlass;
  try

    DeltaDays:= 0;
    if CheckBox3.Checked then
      DeltaDays:= SpinEdit1.Value;

    SQLite.StoreListLoad(UsedNameIDs, DeltaDays,
                       Checkbox1.Checked, Checkbox2.Checked,
                       TestDates, MotorNames, MotorNums);
    SQLite.StoreTotalLoad(UsedNameIDs, DeltaDays,
                        TotalMotorNames, TotalMotorCounts);
    StoreSheet.Draw(DeltaDays, TestDates, MotorNames, MotorNums,
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

