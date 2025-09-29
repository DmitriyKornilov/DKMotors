unit UStoreForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, Spin, DividerBevel, fpspreadsheetgrid,
  //DK packages utils
  DK_Vector, DK_SheetExporter, DK_CtrlUtils,
  //Project utils
  UVars, USheets;

type

  { TStoreForm }

  TStoreForm = class(TForm)
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    DividerBevel1: TDividerBevel;
    ExportButton: TSpeedButton;
    Label2: TLabel;
    Panel2: TPanel;
    Panel5: TPanel;
    ToolPanel: TPanel;
    ReportGrid: TsWorksheetGrid;
    SpinEdit1: TSpinEdit;
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    StoreSheet: TStoreSheet;
    procedure ExportSheet;
  public
    procedure ViewUpdate;
  end;

var
  StoreForm: TStoreForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TStoreForm }

procedure TStoreForm.FormCreate(Sender: TObject);
begin
  MainForm.SetNamesPanelsVisible(True, False);
  StoreSheet:= TStoreSheet.Create(ReportGrid.Worksheet, ReportGrid, GridFont);
  Checkbox2.Visible:= Length(MainForm.UsedNameIDs)<>1;
end;

procedure TStoreForm.FormDestroy(Sender: TObject);
begin
  if Assigned(StoreSheet) then FreeAndNil(StoreSheet);
end;

procedure TStoreForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);
  Images.ToButtons([ExportButton]);

  ViewUpdate;
end;

procedure TStoreForm.CheckBox1Click(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStoreForm.CheckBox2Click(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStoreForm.CheckBox3Click(Sender: TObject);
begin
  SpinEdit1.Enabled:= CheckBox3.Checked;
  ViewUpdate;
end;

procedure TStoreForm.ExportButtonClick(Sender: TObject);
begin
  ExportSheet;
end;

procedure TStoreForm.SpinEdit1Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TStoreForm.ViewUpdate;
var
  TotalMotorNames, MotorNames, MotorNums: TStrVector;
  TotalMotorCounts: TIntVector;
  TestDates: TDateVector;
  DeltaDays: Integer;
begin
  Checkbox2.Visible:= Length(MainForm.UsedNameIDs)<>1;

  Screen.Cursor:= crHourGlass;
  try

    DeltaDays:= 0;
    if CheckBox3.Checked then
      DeltaDays:= SpinEdit1.Value;

    DataBase.StoreListLoad(MainForm.UsedNameIDs, DeltaDays,
                       Checkbox1.Checked, Checkbox2.Checked,
                       TestDates, MotorNames, MotorNums);
    DataBase.StoreTotalLoad(MainForm.UsedNameIDs, DeltaDays,
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

