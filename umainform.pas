unit UMainForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, Menus, DK_LCLStrRus, DK_Vector, DK_VSTTables,
  UBuildLogForm, UShipmentForm, UReclamationForm, UStoreForm, UAboutForm,
  UTestLogForm, UMotorListForm, UReportForm, UStatisticForm, URepairForm,
  UControlListForm, UCalendarForm, USQLite, BCButton, DK_Const, UUtils,
  DK_CtrlUtils, LCLType;

type

  { TMainForm }

  TMainForm = class(TForm)
    Bevel2: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    DictionaryButton: TBCButton;
    ManufactureButton: TBCButton;
    BuildLogMenuItem: TMenuItem;
    DictionaryMenu: TPopupMenu;
    DefectListMenuItem: TMenuItem;
    FactoryListMenuItem: TMenuItem;
    PlaceListMenuItem: TMenuItem;
    ReasonListMenuItem: TMenuItem;
    ReceiverNamesMenuItem: TMenuItem;
    MotorNamesMenuItem: TMenuItem;
    ProductMenu: TPopupMenu;
    ProductButton: TBCButton;
    MotorListMenuItem: TMenuItem;
    ManufactureMenu: TPopupMenu;
    RepairMenuItem: TMenuItem;
    Separator1: TMenuItem;
    Separator2: TMenuItem;
    Separator3: TMenuItem;
    ShipmentMenuItem: TMenuItem;
    ControlListMenuItem: TMenuItem;
    ReportButton: TBCButton;
    CalendarButton: TBCButton;
    StatisticMenuItem: TMenuItem;
    ReclamationMenuItem: TMenuItem;
    ChooseMotorNamesButton: TSpeedButton;
    ChooseRecieverNamesButton: TSpeedButton;
    ExitButton: TSpeedButton;
    ImageListEdit24: TImageList;
    ImageListCategory24: TImageList;
    ImageList16: TImageList;
    Label2: TLabel;
    Label3: TLabel;
    MainPanel: TPanel;
    MotorNamesLabel: TLabel;
    MotorNamesPanel: TPanel;
    RefreshButton: TSpeedButton;
    ClaimButton: TBCButton;
    ClaimMenu: TPopupMenu;
    StoreMenuItem: TMenuItem;
    TestLogMenuItem: TMenuItem;
    ToolPanel: TPanel;
    ReceiverNamesLabel: TLabel;
    ReceiverNamesPanel: TPanel;
    AboutButton: TSpeedButton;
    procedure AboutButtonClick(Sender: TObject);
    procedure BuildLogMenuItemClick(Sender: TObject);
    procedure ChooseMotorNamesButtonClick(Sender: TObject);
    procedure ChooseRecieverNamesButtonClick(Sender: TObject);
    procedure ClaimButtonClick(Sender: TObject);
    procedure ControlListMenuItemClick(Sender: TObject);
    procedure DefectListMenuItemClick(Sender: TObject);
    procedure DictionaryButtonClick(Sender: TObject);
    procedure ExitButtonClick(Sender: TObject);
    procedure FactoryListMenuItemClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Label2Click(Sender: TObject);
    procedure Label3Click(Sender: TObject);
    procedure ManufactureButtonClick(Sender: TObject);
    procedure MotorListMenuItemClick(Sender: TObject);
    procedure MotorNamesLabelClick(Sender: TObject);
    procedure MotorNamesMenuItemClick(Sender: TObject);
    procedure MotorNamesPanelClick(Sender: TObject);
    procedure PlaceListMenuItemClick(Sender: TObject);
    procedure ProductButtonClick(Sender: TObject);
    procedure ReasonListMenuItemClick(Sender: TObject);
    procedure ReceiverNamesLabelClick(Sender: TObject);
    procedure ReceiverNamesMenuItemClick(Sender: TObject);
    procedure ReceiverNamesPanelClick(Sender: TObject);
    procedure ReclamationMenuItemClick(Sender: TObject);
    procedure RefreshButtonClick(Sender: TObject);
    procedure RepairMenuItemClick(Sender: TObject);
    procedure ShipmentMenuItemClick(Sender: TObject);
    procedure ReportButtonClick(Sender: TObject);
    procedure CalendarButtonClick(Sender: TObject);
    procedure StatisticMenuItemClick(Sender: TObject);
    procedure StoreMenuItemClick(Sender: TObject);
    procedure TestLogMenuItemClick(Sender: TObject);
  private
    Category: Byte;

    BuildLogForm: TBuildLogForm;
    TestLogForm: TTestLogForm;
    ShipmentForm: TShipmentForm;
    ReclamationForm: TReclamationForm;
    StoreForm: TStoreForm;
    MotorListForm: TMotorListForm;
    ReportForm: TReportForm;
    StatisticForm: TStatisticForm;
    RepairForm: TRepairForm;
    ControlListForm: TControlListForm;
    CalendarForm: TCalendarForm;

    procedure ConnectDB;
    procedure CategoryChoose(const ACategory: Byte);
    procedure FreeForms;
    procedure SetFormPosition(AForm: TForm);
    procedure OpenBuildLogForm;
    procedure OpenTestForm;
    procedure OpenShipmentForm;
    procedure OpenReclamationForm;
    procedure OpenStoreForm;
    procedure OpenMotorListForm;
    procedure OpenReportForm;
    procedure OpenStatisticForm;
    procedure OpenRepairForm;
    procedure OpenControlListForm;
    procedure OpenCalendarForm;

    procedure ChangeUsedMotorList;
    procedure ChangeUsedReceiverList;

    procedure ShowData;

    procedure DictionaryChoose(const ADictionary: Byte);
  public
    UsedNameIDs: TIntVector;
    UsedNames: TStrVector;

    UsedReceiverIDs: TIntVector;
    UsedReceiverNames: TStrVector;

    procedure SetNamesPanelsVisible(const AMotorNamesPanelVisible,
                                          AReceiverNamesPanelVisible: Boolean);
  end;

var
  MainForm: TMainForm;


implementation

{$R *.lfm}

procedure TMainForm.FormCreate(Sender: TObject);
begin
  //for normal form maximizing
  Height:= 300;
  Width:= 500;

  SetToolPanels([
    ToolPanel, MotorNamesPanel, ReceiverNamesPanel
  ]);
  SetToolButtons([
    RefreshButton, AboutButton, ExitButton, ChooseMotorNamesButton, ChooseRecieverNamesButton
  ]);

  SQLite:= TSQLite.Create;
  ConnectDB;
  SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, False, UsedNameIDs, UsedNames);
  SQLite.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, False, UsedReceiverIDs, UsedReceiverNames);
  CategoryChoose(4);
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if Assigned(SQLite) then FreeAndNil(SQLite);
  FreeForms;
end;

procedure TMainForm.ChooseMotorNamesButtonClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TMainForm.Label2Click(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TMainForm.MotorNamesLabelClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TMainForm.MotorNamesMenuItemClick(Sender: TObject);
begin
  DictionaryChoose(1);
end;

procedure TMainForm.ReceiverNamesMenuItemClick(Sender: TObject);
begin
  DictionaryChoose(2);
end;

procedure TMainForm.DefectListMenuItemClick(Sender: TObject);
begin
  DictionaryChoose(3);
end;

procedure TMainForm.ReasonListMenuItemClick(Sender: TObject);
begin
  DictionaryChoose(4);
end;

procedure TMainForm.FactoryListMenuItemClick(Sender: TObject);
begin
  DictionaryChoose(5);
end;

procedure TMainForm.PlaceListMenuItemClick(Sender: TObject);
begin
  DictionaryChoose(6);
end;

procedure TMainForm.MotorNamesPanelClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TMainForm.ChooseRecieverNamesButtonClick(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.ManufactureButtonClick(Sender: TObject);
begin
  ControlPopupMenuShow(Sender, ManufactureMenu);
end;

procedure TMainForm.ProductButtonClick(Sender: TObject);
begin
  ControlPopupMenuShow(Sender, ProductMenu);
end;

procedure TMainForm.ClaimButtonClick(Sender: TObject);
begin
  ControlPopupMenuShow(Sender, ClaimMenu);
end;

procedure TMainForm.DictionaryButtonClick(Sender: TObject);
begin
  ControlPopupMenuShow(Sender, DictionaryMenu);
end;

procedure TMainForm.ControlListMenuItemClick(Sender: TObject);
begin
  if not Assigned(ControlListForm) then CategoryChoose(3);
end;

procedure TMainForm.Label3Click(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.MotorListMenuItemClick(Sender: TObject);
begin
  if not Assigned(MotorListForm) then CategoryChoose(4);
end;

procedure TMainForm.ReceiverNamesLabelClick(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.ReceiverNamesPanelClick(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.ReclamationMenuItemClick(Sender: TObject);
begin
  if not Assigned(ReclamationForm) then CategoryChoose(7);
end;

procedure TMainForm.RefreshButtonClick(Sender: TObject);
begin
  SQLite.Reconnect;
  CategoryChoose(Category);
end;

procedure TMainForm.ExitButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.AboutButtonClick(Sender: TObject);
var
  AboutForm: TAboutForm;
begin
  AboutForm:= TAboutForm.Create(MainForm);
  AboutForm.ShowModal;
  FreeAndNil(AboutForm);
end;

procedure TMainForm.BuildLogMenuItemClick(Sender: TObject);
begin
  if not Assigned(BuildLogForm) then CategoryChoose(1);
end;

procedure TMainForm.RepairMenuItemClick(Sender: TObject);
begin
  if not Assigned(RepairForm) then CategoryChoose(9);
end;

procedure TMainForm.ShipmentMenuItemClick(Sender: TObject);
begin
  if not Assigned(ShipmentForm) then CategoryChoose(6);
end;

procedure TMainForm.ReportButtonClick(Sender: TObject);
begin
  if not Assigned(ReportForm) then CategoryChoose(10);
end;

procedure TMainForm.CalendarButtonClick(Sender: TObject);
begin
  if not Assigned(CalendarForm) then CategoryChoose(11);
end;

procedure TMainForm.StatisticMenuItemClick(Sender: TObject);
begin
  if not Assigned(StatisticForm) then CategoryChoose(8);
end;

procedure TMainForm.StoreMenuItemClick(Sender: TObject);
begin
  if not Assigned(StoreForm) then CategoryChoose(5);
end;

procedure TMainForm.TestLogMenuItemClick(Sender: TObject);
begin
  if not Assigned(TestLogForm) then CategoryChoose(2);
end;

procedure TMainForm.ConnectDB;
var
  ProgPath, DBName, DDLName: String;
begin
  ProgPath:= ExtractFilePath(Application.ExeName);
  DBName:= ProgPath + 'DKMotors.db';
  DDLName:= ProgPath + 'DDL.sql';
  SQLite.Connect(DBName);
  SQLite.ExecuteScript(DDLName);
end;

procedure TMainForm.CategoryChoose(const ACategory: Byte);
begin
  Screen.Cursor:= crHourGlass;
  try
    FreeForms;
    Category:= ACategory;
    case Category of
      1: OpenBuildLogForm;
      2: OpenTestForm;
      3: OpenControlListForm;
      4: OpenMotorListForm;
      5: OpenStoreForm;
      6: OpenShipmentForm;
      7: OpenReclamationForm;
      8: OpenStatisticForm;
      9: OpenRepairForm;
      10: OpenReportForm;
      11: OpenCalendarForm;
    end;
  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TMainForm.FreeForms;
begin
  if Assigned(BuildLogForm) then FreeAndNil(BuildLogForm);
  if Assigned(TestLogForm) then FreeAndNil(TestLogForm);
  if Assigned(ShipmentForm) then FreeAndNil(ShipmentForm);
  if Assigned(ReclamationForm) then FreeAndNil(ReclamationForm);
  if Assigned(StoreForm) then FreeAndNil(StoreForm);
  if Assigned(ReportForm) then FreeAndNil(ReportForm);
  if Assigned(StatisticForm) then FreeAndNil(StatisticForm);
  if Assigned(MotorListForm) then FreeAndNil(MotorListForm);
  if Assigned(RepairForm) then FreeAndNil(RepairForm);
  if Assigned(ControlListForm) then FreeAndNil(ControlListForm);
  if Assigned(CalendarForm) then FreeAndNil(CalendarForm);
end;

procedure TMainForm.SetFormPosition(AForm: TForm);
begin
  AForm.Parent:= MainPanel;
  AForm.Left:= 0;
  AForm.Top:= 0;
  AForm.MakeFullyVisible();
end;

procedure TMainForm.OpenBuildLogForm;
begin
  BuildLogForm:= TBuildLogForm.Create(MainForm);
  SetFormPosition(TForm(BuildLogForm));
  BuildLogForm.Show;
end;

procedure TMainForm.OpenTestForm;
begin
  TestLogForm:= TTestLogForm.Create(MainForm);
  SetFormPosition(TForm(TestLogForm));
  TestLogForm.Show;
end;

procedure TMainForm.OpenShipmentForm;
begin
  ShipmentForm:= TShipmentForm.Create(MainForm);
  SetFormPosition(TForm(ShipmentForm));
  ShipmentForm.Show;
end;

procedure TMainForm.OpenReclamationForm;
begin
  ReclamationForm:= TReclamationForm.Create(MainForm);
  SetFormPosition(TForm(ReclamationForm));
  ReclamationForm.Show;
end;

procedure TMainForm.OpenStoreForm;
begin
  StoreForm:= TStoreForm.Create(MainForm);
  SetFormPosition(TForm(StoreForm));
  StoreForm.Show;
end;

procedure TMainForm.OpenMotorListForm;
begin
  MotorListForm:= TMotorListForm.Create(MainForm);
  SetFormPosition(TForm(MotorListForm));
  MotorListForm.Show;
end;

procedure TMainForm.OpenReportForm;
begin
  ReportForm:= TReportForm.Create(MainForm);
  SetFormPosition(TForm(ReportForm));
  ReportForm.Show;
end;

procedure TMainForm.OpenStatisticForm;
begin
  StatisticForm:= TStatisticForm.Create(MainForm);
  SetFormPosition(TForm(StatisticForm));
  StatisticForm.Show;
end;

procedure TMainForm.OpenRepairForm;
begin
  RepairForm:= TRepairForm.Create(MainForm);
  SetFormPosition(TForm(RepairForm));
  RepairForm.Show;
end;

procedure TMainForm.OpenControlListForm;
begin
  ControlListForm:= TControlListForm.Create(MainForm);
  SetFormPosition(TForm(ControlListForm));
  ControlListForm.Show;
end;

procedure TMainForm.OpenCalendarForm;
begin
  CalendarForm:= TCalendarForm.Create(MainForm);
  SetFormPosition(TForm(CalendarForm));
  CalendarForm.Show;
end;

procedure TMainForm.ChangeUsedMotorList;
begin
  if SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, True,
    UsedNameIDs, UsedNames) then ShowData;
end;

procedure TMainForm.ChangeUsedReceiverList;
begin
  if SQLite.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, True,
    UsedReceiverIDs, UsedReceiverNames) then ShowData;
end;

procedure TMainForm.ShowData;
begin
  case Category of
    1: BuildLogForm.ShowBuildLog;
    2: TestLogForm.ShowTestLog;
    3: ControlListForm.ShowControlList;
    4: MotorListForm.ShowMotorList;
    5: StoreForm.ShowStore;
    6: ShipmentForm.ShowShipment;
    7: ReclamationForm.ShowReclamation;
    8: StatisticForm.ShowStatistic;
    9: RepairForm.ShowRepair;
    10: ReportForm.ShowReport;
    //11: ;
  end;
end;

procedure TMainForm.DictionaryChoose(const ADictionary: Byte);
var
  IsOK: Boolean;
begin
  case ADictionary of
    1: IsOK:= SQLite.EditList('Наименования электродвигателей',
                              'MOTORNAMES', 'NameID', 'MotorName', False, True);
    2: IsOK:= SQLite.EditList('Грузополучатели',
                              'CARGORECEIVERS', 'ReceiverID', 'ReceiverName', True, True);
    3: IsOK:= SQLite.EditList('Элементы электродвигателей',
                              'RECLAMATIONDEFECTS', 'DefectID', 'DefectName', True, True);
    4: IsOK:= SQLite.EditTable('Причины возникновения неисправностей',
                              'RECLAMATIONREASONS', 'ReasonID',
                              ['ReasonName',  'ReasonColor'],
                              ['Причина',     'Цвет'       ],
                              [ctString,      ctColor      ],
                              [True,          True         ],
                              [300,           80           ],
                              [taLeftJustify, taCenter     ],
                              True, ['ReasonName']);
    5: IsOK:= SQLite.EditList('Заводы',
                              'RECLAMATIONFACTORIES', 'FactoryID', 'FactoryName', True, True);
    6: IsOK:= SQLite.EditList('Предприятия (депо)',
                              'RECLAMATIONPLACES', 'PlaceID', 'PlaceName', True, True);
  end;
  if IsOK then ShowData;
end;

procedure TMainForm.SetNamesPanelsVisible(const AMotorNamesPanelVisible,
  AReceiverNamesPanelVisible: Boolean);
begin
  MotorNamesPanel.Visible:= AMotorNamesPanelVisible;
  ReceiverNamesPanel.Visible:= AReceiverNamesPanelVisible;
end;

end.

