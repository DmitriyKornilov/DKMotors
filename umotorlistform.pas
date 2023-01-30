unit UMotorListForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, fpspreadsheetgrid, SheetUtils, USQLite, rxctrls,
  DividerBevel,  DK_Vector, DK_SheetExporter,
  FPSTypes, LCLType, EditBtn, Spin, DK_StrUtils, DK_DateUtils,
  VirtualTrees, DK_VSTTables;

type

  { TMotorListForm }

  TMotorListForm = class(TForm)
    CheckBox1: TCheckBox;
    MoreInfoCheckBox: TCheckBox;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    DividerBevel4: TDividerBevel;
    DividerBevel5: TDividerBevel;
    DividerBevel7: TDividerBevel;
    InfoGrid: TsWorksheetGrid;
    MotorNumEdit: TEditButton;
    ExportButton: TRxSpeedButton;
    Label2: TLabel;
    MotorShippedComboBox: TComboBox;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    SpinEdit1: TSpinEdit;
    Splitter2: TSplitter;
    VT1: TVirtualStringTree;
    procedure CheckBox1Change(Sender: TObject);
    procedure MoreInfoCheckBoxChange(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure MotorShippedComboBoxChange(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure VT1MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
  private
    VSTTable: TVSTTable;
    MotorInfoSheet: TMotorInfoSheet;
    MotorIDs: TIntVector;

    procedure InfoOpen;
  public
    procedure ShowMotorList;
  end;

var
  MotorListForm: TMotorListForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TMotorListForm }

procedure TMotorListForm.FormCreate(Sender: TObject);
begin
  MainForm.SetNamesPanelsVisible(True, False);

  VSTTable:= TVSTTable.Create(VT1);
  VSTTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTTable.HeaderFont.Style:= [fsBold];
  VSTTable.AddColumn('Дата сборки', 150);
  VSTTable.AddColumn('Наименование', 300);
  VSTTable.AddColumn('Номер', 150);
  VSTTable.AddColumn('Отгружен', 150);
  VSTTable.CanSelect:= True;
  VSTTable.Draw;

  MotorInfoSheet:= TMotorInfoSheet.Create(InfoGrid);

  SpinEdit1.Value:= YearOfDate(Date);
  ShowMotorList;
end;

procedure TMotorListForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTTable) then FreeAndNil(VSTTable);
  if Assigned(MotorInfoSheet) then FreeAndNil(MotorInfoSheet);
end;

procedure TMotorListForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TMotorListForm.MotorNumEditChange(Sender: TObject);
begin
  ShowMotorList;
end;

procedure TMotorListForm.MotorShippedComboBoxChange(Sender: TObject);
begin
  ShowMotorList;
end;

procedure TMotorListForm.SpinEdit1Change(Sender: TObject);
begin
  ShowMotorList;
end;

procedure TMotorListForm.VT1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  InfoGrid.Clear;
  if not VSTTable.IsSelected then Exit;
  InfoOpen;
end;

procedure TMotorListForm.ShowMotorList;
var
  MotorNumberLike: String;

  ABuildDates, AMotorNames, AMotorNums, AShippings: TStrVector;
begin
  Screen.Cursor:= crHourGlass;
  try
    InfoGrid.Clear;

    MotorNumberLike:= STrim(MotorNumEdit.Text);

    SQLite.MotorListLoad(SpinEdit1.Value,
                        MotorShippedComboBox.ItemIndex, MainForm.UsedNameIDs,
                        MotorNumberLike, Checkbox1.Checked,
                        MotorIDs, ABuildDates,
                        AMotorNames, AMotorNums, AShippings);

    VSTTable.ValuesClear;
    VSTTable.SetColumn('Дата сборки', ABuildDates);
    VSTTable.SetColumn('Наименование', AMotorNames);
    VSTTable.SetColumn('Номер', AMotorNums);
    VSTTable.SetColumn('Отгружен', AShippings, taLeftJustify);
    VSTTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TMotorListForm.InfoOpen;
var
  MotorID: Integer;
  BuildDate, SendDate: TDate;
  MotorName, MotorNum, RotorNum, ReceiverName, Sers: String;
  TestDates, RecDates: TDateVector;
  TestResults, Mileages, Opinions: TIntVector;
  TestNotes, PlaceNames, FactoryNames, Departures,
  DefectNames, ReasonNames, RecNotes: TStrVector;
begin
  if VIsNil(MotorIDs) then Exit;
  if not VSTTable.IsSelected then Exit;

  MotorID:= MotorIDs[VSTTable.SelectedIndex];
  SQLite.MotorInfoLoad(MotorID, BuildDate, SendDate,
                MotorName, MotorNum, Sers, RotorNum, ReceiverName,
                TestDates, TestResults, TestNotes);

  SQLite.ReclamationListLoad(MotorID, RecDates, Mileages, Opinions,
                      PlaceNames, FactoryNames, Departures,
                      DefectNames, ReasonNames, RecNotes);


  MotorInfoSheet.Draw(BuildDate, SendDate, MotorName, MotorNum, Sers,
                      RotorNum, ReceiverName, TestDates, TestResults, TestNotes,
                      RecDates, Mileages, Opinions, PlaceNames, FactoryNames,
                      Departures, DefectNames, ReasonNames, RecNotes);
end;

procedure TMotorListForm.CheckBox1Change(Sender: TObject);
begin
  ShowMotorList;
end;

procedure TMotorListForm.MoreInfoCheckBoxChange(Sender: TObject);
begin
  If MoreInfoCheckBox.Checked then
  begin
    VT1.Align:= alCustom;
    Splitter2.Visible:= True;
    Splitter2.Align:= alTop;
    Panel2.Visible:= True;
    Splitter2.Align:= alBottom;
    VT1.Align:= alClient;
  end
  else begin
    Panel2.Visible:= False;
    Splitter2.Visible:= False;
  end;

  VSTTable.CanSelect:= MoreInfoCheckBox.Checked;
end;

procedure TMotorListForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TGridExporter;
begin
  if not VSTTable.IsSelected then Exit;
  Exporter:= TGridExporter.Create(InfoGrid);
  try
    //Exporter.SheetName:= 'Отчет';
    Exporter.PageSettings(spoLandscape, pfOnePage);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

end.

