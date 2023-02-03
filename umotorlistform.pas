unit UMotorListForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, fpspreadsheetgrid, USheetUtils, USQLite, rxctrls,
  DividerBevel,  DK_Vector, DK_SheetExporter,
  FPSTypes, LCLType, EditBtn, Spin, DK_StrUtils, DK_DateUtils,
  VirtualTrees, DK_VSTTables, UCardForm;

type

  { TMotorListForm }

  TMotorListForm = class(TForm)
    CheckBox1: TCheckBox;
    Label4: TLabel;
    MoreInfoCheckBox: TCheckBox;
    DividerBevel3: TDividerBevel;
    DividerBevel4: TDividerBevel;
    DividerBevel5: TDividerBevel;
    DividerBevel7: TDividerBevel;
    MotorNumEdit: TEditButton;
    ExportButton: TRxSpeedButton;
    Label2: TLabel;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    CardPanel: TPanel;
    Panel5: TPanel;
    Panel7: TPanel;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    procedure CheckBox1Change(Sender: TObject);
    procedure MoreInfoCheckBoxChange(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    CardForm: TCardForm;
    VSTMotorsTable: TVSTTable;
    VSTTypeTable: TVSTTable;
    MotorIDs: TIntVector;

    procedure TypeSelect;
    procedure MotorSelect;
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
var
  V: TStrVector;
begin
  MainForm.SetNamesPanelsVisible(True, False);
  CardForm:= CrateCardForm(MotorListForm, CardPanel);

  VSTMotorsTable:= TVSTTable.Create(VT1);
  VSTMotorsTable.OnSelect:= @MotorSelect;
  VSTMotorsTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTMotorsTable.HeaderFont.Style:= [fsBold];
  VSTMotorsTable.AddColumn('Дата сборки', 150);
  VSTMotorsTable.AddColumn('Наименование', 300);
  VSTMotorsTable.AddColumn('Номер', 150);
  VSTMotorsTable.AddColumn('Отгружен', 150);
  VSTMotorsTable.CanSelect:= True;
  VSTMotorsTable.Draw;

  VSTTypeTable:= TVSTTable.Create(VT2);
  VSTTypeTable.OnSelect:= @TypeSelect;
  VSTTypeTable.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTTypeTable.HeaderVisible:= False;
  VSTTypeTable.GridLinesVisible:= False;
  VSTTypeTable.CanSelect:= True;
  VSTTypeTable.CanUnselect:= False;
  VSTTypeTable.AddColumn('Список');
  V:= VCreateStr(['Все', 'Отгруженные', 'Неотгруженные']);
  VSTTypeTable.SetColumn('Список', V, taLeftJustify);
  VSTTypeTable.Draw;

  SpinEdit1.Value:= YearOfDate(Date);

  VSTTypeTable.Select(0);

  ShowMotorList;
end;

procedure TMotorListForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTMotorsTable) then FreeAndNil(VSTMotorsTable);
  if Assigned(VSTTypeTable) then FreeAndNil(VSTTypeTable);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TMotorListForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TMotorListForm.MotorNumEditChange(Sender: TObject);
begin
  ShowMotorList;
end;

procedure TMotorListForm.SpinEdit1Change(Sender: TObject);
begin
  ShowMotorList;
end;

procedure TMotorListForm.ShowMotorList;
var
  MotorNumberLike: String;

  ABuildDates, AMotorNames, AMotorNums, AShippings: TStrVector;
begin
   if not VSTTypeTable.IsSelected then Exit;

  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);

    MotorNumberLike:= STrim(MotorNumEdit.Text);

    SQLite.MotorListLoad(SpinEdit1.Value,
                        VSTTypeTable.SelectedIndex, MainForm.UsedNameIDs,
                        MotorNumberLike, Checkbox1.Checked,
                        MotorIDs, ABuildDates,
                        AMotorNames, AMotorNums, AShippings);

    VSTMotorsTable.ValuesClear;
    VSTMotorsTable.SetColumn('Дата сборки', ABuildDates);
    VSTMotorsTable.SetColumn('Наименование', AMotorNames);
    VSTMotorsTable.SetColumn('Номер', AMotorNums);
    VSTMotorsTable.SetColumn('Отгружен', AShippings, taLeftJustify);
    VSTMotorsTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;
end;



procedure TMotorListForm.TypeSelect;
begin
  ShowMotorList;
end;

procedure TMotorListForm.MotorSelect;
var
  MotorID: Integer;
begin
  MotorID:= 0;
  if VSTMotorsTable.IsSelected then
    MotorID:= MotorIDs[VSTMotorsTable.SelectedIndex];
  CardForm.ShowCard(MotorID);
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
    CardPanel.Visible:= True;
    Splitter2.Align:= alBottom;
    VT1.Align:= alClient;
  end
  else begin
    CardPanel.Visible:= False;
    Splitter2.Visible:= False;
  end;

  VSTMotorsTable.CanSelect:= MoreInfoCheckBox.Checked;
end;

procedure TMotorListForm.ExportButtonClick(Sender: TObject);

begin

end;

end.

