unit UMotorListForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, fpspreadsheetgrid, Spin, LCLType, DividerBevel, VirtualTrees,
  //DK packages utils
  DK_Vector, DK_StrUtils, DK_DateUtils, DK_VSTTables, DK_Const, DK_VSTParamList,
  DK_CtrlUtils, DK_Filter,
  //Project utils
  UVars,
  //Forms
  UCardForm;

type

  { TMotorListForm }

  TMotorListForm = class(TForm)
    OrderByNumCheckBox: TCheckBox;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    MoreInfoCheckBox: TCheckBox;
    SettingClientPanel: TPanel;
    MainPanel: TPanel;
    SpinEdit1: TSpinEdit;
    ToolPanel: TPanel;
    CardPanel: TPanel;
    FilterPanel: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    VT1: TVirtualStringTree;
    YearPanel: TPanel;
    procedure OrderByNumCheckBoxChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MoreInfoCheckBoxChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    FilterString: String;

    CardForm: TCardForm;
    MotorsTable: TVSTTable;
    ParamList: TVSTParamList;

    MotorIDs: TIntVector;

    procedure CreateMotorsTable;
    procedure SelectMotor;
    procedure FilterMotor(const AFilterString: String);

    procedure CreateParamList;
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
  CreateParamList;
  FilterString:= EmptyStr;
  DKFilterCreate('Поиск по номеру:', FilterPanel, @FilterMotor, 300, 500);
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TMotorListForm.FormDestroy(Sender: TObject);
begin
  if Assigned(MotorsTable) then FreeAndNil(MotorsTable);
  if Assigned(ParamList) then FreeAndNil(ParamList);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TMotorListForm.FormShow(Sender: TObject);
begin
  SetToolPanels([ToolPanel]);
  ParamList.AutoHeight;
  ViewUpdate;
end;

procedure TMotorListForm.SpinEdit1Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TMotorListForm.CreateMotorsTable;
begin
  MotorsTable:= TVSTTable.Create(VT1);
  MotorsTable.SetSingleFont(GridFont);
  MotorsTable.OnSelect:= @SelectMotor;
  MotorsTable.HeaderFont.Style:= [fsBold];
  MotorsTable.HeaderHeight:= 25;
  MotorsTable.AddColumn('Дата сборки', 100);
  MotorsTable.AddColumn('Наименование', 200);
  MotorsTable.AddColumn('Номер', 100);
  MotorsTable.AddColumn('Отгружен');
  MotorsTable.CanSelect:= True;
  MotorsTable.Draw;
end;

procedure TMotorListForm.CreateParamList;
var
  S: String;
  V: TStrVector;
begin
  ParamList:= TVSTParamList.Create(SettingClientPanel);

  S:= 'Отображать:';
  V:= VCreateStr([
    'все',
    'отгруженные',
    'неотгруженные'
  ]);

  ParamList.AddStringList('TypeList', S, V, @ViewUpdate);
end;

procedure TMotorListForm.ViewUpdate;
var
  ABuildDates, AMotorNames, AMotorNums, AShippings: TStrVector;
begin
  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);
    DataBase.MotorListLoad(SpinEdit1.Value,
                        ParamList.Selected['TypeList'], MainForm.UsedNameIDs,
                        STrim(FilterString), OrderByNumCheckBox.Checked,
                        MotorIDs, ABuildDates,
                        AMotorNames, AMotorNums, AShippings);

    MotorsTable.ValuesClear;
    MotorsTable.SetColumn('Дата сборки', ABuildDates);
    MotorsTable.SetColumn('Наименование', AMotorNames, taLeftJustify);
    MotorsTable.SetColumn('Номер', AMotorNums);
    MotorsTable.SetColumn('Отгружен', AShippings, taLeftJustify);
    MotorsTable.Draw;
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TMotorListForm.SelectMotor;
var
  MotorID: Integer;
begin
  MotorID:= 0;
  if MotorsTable.IsSelected then
    MotorID:= MotorIDs[MotorsTable.SelectedIndex];
  CardForm.ShowCard(MotorID);
end;

procedure TMotorListForm.FilterMotor(const AFilterString: String);
begin
  FilterString:= AFilterString;
  ViewUpdate;
end;

procedure TMotorListForm.OrderByNumCheckBoxChange(Sender: TObject);
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

  MotorsTable.CanSelect:= MoreInfoCheckBox.Checked;
end;

end.

