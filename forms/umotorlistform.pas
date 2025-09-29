unit UMotorListForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, fpspreadsheetgrid, Spin, LCLType, DividerBevel, VirtualTrees,
  //DK packages utils
  DK_Vector, DK_StrUtils, DK_DateUtils, DK_VSTTables, DK_Const, DK_VSTTableTools,
  DK_CtrlUtils, DK_Filter,
  //Project utils
  UVars,
  //Forms
  UCardForm;

type

  { TMotorListForm }

  TMotorListForm = class(TForm)
    CheckBox1: TCheckBox;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    MoreInfoCheckBox: TCheckBox;
    LeftPanel: TPanel;
    MainPanel: TPanel;
    SpinEdit1: TSpinEdit;
    ToolPanel: TPanel;
    CardPanel: TPanel;
    FilterPanel: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    YearPanel: TPanel;
    procedure CheckBox1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MoreInfoCheckBoxChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    FilterString: String;

    CardForm: TCardForm;
    VSTMotorsTable: TVSTTable;
    VSTTypesList: TVSTStringList;
    MotorIDs: TIntVector;

    procedure CreateMotorsTable;
    procedure SelectMotor;
    procedure FilterMotor(const AFilterString: String);

    procedure CreateTypesList;
  public
    procedure ViewUpdate;
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
  CardForm:= CreateCardForm(MotorListForm, CardPanel);
  CreateMotorsTable;
  CreateTypesList;
  FilterString:= EmptyStr;
  DKFilterCreate('Поиск по номеру:', FilterPanel, @FilterMotor, 250, 500);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TMotorListForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTMotorsTable) then FreeAndNil(VSTMotorsTable);
  if Assigned(VSTTypesList) then FreeAndNil(VSTTypesList);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TMotorListForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);

  ViewUpdate;
end;

procedure TMotorListForm.SpinEdit1Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TMotorListForm.CreateMotorsTable;
begin
  VSTMotorsTable:= TVSTTable.Create(VT1);
  VSTMotorsTable.SetSingleFont(GridFont);
  VSTMotorsTable.OnSelect:= @SelectMotor;
  VSTMotorsTable.HeaderFont.Style:= [fsBold];
  VSTMotorsTable.HeaderHeight:= 25;
  VSTMotorsTable.AddColumn('Дата сборки', 100);
  VSTMotorsTable.AddColumn('Наименование', 200);
  VSTMotorsTable.AddColumn('Номер', 100);
  VSTMotorsTable.AddColumn('Отгружен');
  VSTMotorsTable.CanSelect:= True;
  VSTMotorsTable.Draw;
end;

procedure TMotorListForm.CreateTypesList;
var
  S: String;
  V: TStrVector;
begin
  S:= 'Отображать:';
  V:= VCreateStr([
    'все',
    'отгруженные',
    'неотгруженные'
  ]);
  VSTTypesList:= TVSTStringList.Create(VT2, S, @ViewUpdate);
  VSTTypesList.Update(V);
end;

procedure TMotorListForm.ViewUpdate;
var
  ABuildDates, AMotorNames, AMotorNums, AShippings: TStrVector;
begin
   if not VSTTypesList.IsSelected then Exit;

  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);
    DataBase.MotorListLoad(SpinEdit1.Value,
                        VSTTypesList.SelectedIndex, MainForm.UsedNameIDs,
                        STrim(FilterString), Checkbox1.Checked,
                        MotorIDs, ABuildDates,
                        AMotorNames, AMotorNums, AShippings);

    VSTMotorsTable.ValuesClear;
    VSTMotorsTable.SetColumn('Дата сборки', ABuildDates);
    VSTMotorsTable.SetColumn('Наименование', AMotorNames, taLeftJustify);
    VSTMotorsTable.SetColumn('Номер', AMotorNums);
    VSTMotorsTable.SetColumn('Отгружен', AShippings, taLeftJustify);
    VSTMotorsTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TMotorListForm.SelectMotor;
var
  MotorID: Integer;
begin
  MotorID:= 0;
  if VSTMotorsTable.IsSelected then
    MotorID:= MotorIDs[VSTMotorsTable.SelectedIndex];
  CardForm.ShowCard(MotorID);
end;

procedure TMotorListForm.FilterMotor(const AFilterString: String);
begin
  FilterString:= AFilterString;
  ViewUpdate;
end;

procedure TMotorListForm.CheckBox1Change(Sender: TObject);
begin
  ViewUpdate;
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

end.

