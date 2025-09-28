unit UMotorListForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls, EditBtn,
  Buttons, fpspreadsheetgrid, Spin, LCLType, VirtualTrees,
  //DK packages utils
  DK_Vector, DK_StrUtils, DK_DateUtils, DK_VSTTables, DK_Const, DK_VSTTableTools,
  //Project utils
  UDataBase, UUtils,
  //Forms
  UCardForm;

type

  { TMotorListForm }

  TMotorListForm = class(TForm)
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    CheckBox1: TCheckBox;
    MoreInfoCheckBox: TCheckBox;
    MotorNumEdit: TEditButton;
    Label2: TLabel;
    LeftPanel: TPanel;
    MainPanel: TPanel;
    SpinEdit1: TSpinEdit;
    ToolPanel: TPanel;
    CardPanel: TPanel;
    Panel7: TPanel;
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
    procedure MotorNumEditButtonClick(Sender: TObject);
    procedure MotorNumEditChange(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
  private
    CardForm: TCardForm;
    VSTMotorsTable: TVSTTable;
    VSTTypesList: TVSTStringList;
    MotorIDs: TIntVector;

    procedure CreateMotorsTable;
    procedure SelectMotor;

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
  SetToolPanels([
    ToolPanel
  ]);

  MainForm.SetNamesPanelsVisible(True, False);
  CardForm:= CreateCardForm(MotorListForm, CardPanel);
  CreateMotorsTable;
  CreateTypesList;
  SpinEdit1.Value:= YearOfDate(Date);
end;

procedure TMotorListForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTMotorsTable) then FreeAndNil(VSTMotorsTable);
  if Assigned(VSTTypesList) then FreeAndNil(VSTTypesList);
  if Assigned(CardForm) then FreeAndNil(CardForm);
end;

procedure TMotorListForm.MotorNumEditButtonClick(Sender: TObject);
begin
  MotorNumEdit.Text:= EmptyStr;
end;

procedure TMotorListForm.MotorNumEditChange(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TMotorListForm.SpinEdit1Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TMotorListForm.CreateMotorsTable;
begin
  VSTMotorsTable:= TVSTTable.Create(VT1);
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
  MotorNumberLike: String;

  ABuildDates, AMotorNames, AMotorNums, AShippings: TStrVector;
begin
   if not VSTTypesList.IsSelected then Exit;

  Screen.Cursor:= crHourGlass;
  try
    CardForm.ShowCard(0);

    MotorNumberLike:= STrim(MotorNumEdit.Text);

    DataBase.MotorListLoad(SpinEdit1.Value,
                        VSTTypesList.SelectedIndex, MainForm.UsedNameIDs,
                        MotorNumberLike, Checkbox1.Checked,
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

procedure TMotorListForm.CheckBox1Change(Sender: TObject);
begin
  ViewUpdate;
end;

procedure TMotorListForm.FormShow(Sender: TObject);
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

