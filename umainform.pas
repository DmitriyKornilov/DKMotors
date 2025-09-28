unit UMainForm;

{$mode objfpc}{$H+}

//{$DEFINE DEBUG}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, Menus, LCLType, BCButton, SQLDB,
  //DK packages utils
  {$IFDEF DEBUG}
  DK_HeapTrace,
  {$ENDIF}
  DK_LCLStrRus, DK_Vector, DK_VSTTypes, DK_CtrlUtils, DK_Const,
  //Project utils
  UDataBase,
  //Forms
  UBuildLogForm, UShipmentForm, UReclamationForm, UStoreForm, UAboutForm,
  UTestLogForm, UMotorListForm, UReportForm, UStatisticForm, URepairForm,
  UControlListForm, UCalendarForm;

type

  { TMainForm }

  TMainForm = class(TForm)
    BaseDDLScript: TSQLScript;
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
    procedure FormShow(Sender: TObject);
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
    CategoryForm: TForm;

    procedure ConnectDB;

    procedure ViewUpdate;

    procedure ChangeUsedMotorList;
    procedure ChangeUsedReceiverList;

    procedure DictionaryChoose(const ADictionary: Byte);
  public
    UsedNameIDs: TIntVector;
    UsedNames: TStrVector;

    UsedReceiverIDs: TIntVector;
    UsedReceiverNames: TStrVector;

    procedure CategorySelect(const ACategory: Byte);

    procedure SetNamesPanelsVisible(const AMotorNamesPanelVisible,
                                          AReceiverNamesPanelVisible: Boolean);
  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

procedure TMainForm.FormCreate(Sender: TObject);
begin
  {$IFDEF DEBUG}
  HeapTraceOutputFile('trace.trc');
  {$ENDIF}

  Caption:= MAIN_CAPTION + ' - ' + PROJECT_NOTE;
  ConnectDB;

  DataBase.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, False, UsedNameIDs, UsedNames);
  DataBase.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, False, UsedReceiverIDs, UsedReceiverNames);

  CategorySelect(4);
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if Assigned(CategoryForm) then FreeAndNil(CategoryForm);
  if Assigned(DataBase) then FreeAndNil(DataBase);
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
  SetToolPanels([
    ToolPanel, MotorNamesPanel, ReceiverNamesPanel
  ]);
  SetToolButtons([
    RefreshButton, AboutButton, ExitButton, ChooseMotorNamesButton, ChooseRecieverNamesButton
  ]);
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
  if not Assigned(ControlListForm) then CategorySelect(3);
end;

procedure TMainForm.Label3Click(Sender: TObject);
begin
  ChangeUsedReceiverList;
end;

procedure TMainForm.MotorListMenuItemClick(Sender: TObject);
begin
  if not Assigned(MotorListForm) then CategorySelect(4);
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
  if not Assigned(ReclamationForm) then CategorySelect(7);
end;

procedure TMainForm.RefreshButtonClick(Sender: TObject);
begin
  DataBase.Reconnect;
  CategorySelect(Category);
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
  if not Assigned(BuildLogForm) then CategorySelect(1);
end;

procedure TMainForm.RepairMenuItemClick(Sender: TObject);
begin
  if not Assigned(RepairForm) then CategorySelect(9);
end;

procedure TMainForm.ShipmentMenuItemClick(Sender: TObject);
begin
  if not Assigned(ShipmentForm) then CategorySelect(6);
end;

procedure TMainForm.ReportButtonClick(Sender: TObject);
begin
  if not Assigned(ReportForm) then CategorySelect(10);
end;

procedure TMainForm.CalendarButtonClick(Sender: TObject);
begin
  if not Assigned(CalendarForm) then CategorySelect(11);
end;

procedure TMainForm.StatisticMenuItemClick(Sender: TObject);
begin
  if not Assigned(StatisticForm) then CategorySelect(8);
end;

procedure TMainForm.StoreMenuItemClick(Sender: TObject);
begin
  if not Assigned(StoreForm) then CategorySelect(5);
end;

procedure TMainForm.TestLogMenuItemClick(Sender: TObject);
begin
  if not Assigned(TestLogForm) then CategorySelect(2);
end;

procedure TMainForm.ConnectDB;
var
  DBPath, DBName{, DDLName}: String;
begin
  DBPath:= ExtractFilePath(Application.ExeName) + 'db' + DirectorySeparator;
  DBName:= DBPath + 'DKMotors.db';
  //DDLName:= DBPath + 'ddl.sql';

  DataBase:= TDataBase.Create;
  DataBase.Connect(DBName);

  //DataBase.ExecuteScript(DDLName);
  DataBase.ExecuteScript(BaseDDLScript);
end;

procedure TMainForm.ViewUpdate;
begin
  if not Assigned(CategoryForm) then Exit;

  case Category of
    1: (CategoryForm as TBuildLogForm).ViewUpdate;
    2: (CategoryForm as TTestLogForm).ViewUpdate;
    3: (CategoryForm as TControlListForm).ViewUpdate;
    4: (CategoryForm as TMotorListForm).ViewUpdate;
    5: (CategoryForm as TStoreForm).ViewUpdate;
    6: (CategoryForm as TShipmentForm).ViewUpdate;
    7: (CategoryForm as TReclamationForm).ViewUpdate;
    8: (CategoryForm as TStatisticForm).ViewUpdate;
    9: (CategoryForm as TRepairForm).ViewUpdate;
    10: (CategoryForm as TReportForm).ViewUpdate;
    //11: ;
  end;
end;

procedure TMainForm.CategorySelect(const ACategory: Byte);
begin
  if ACategory=Category then Exit;

  Screen.Cursor:= crHourGlass;
  try
    Category:= ACategory;

    if Assigned(CategoryForm) then FreeAndNil(CategoryForm);
    case Category of
      1: CategoryForm:= FormOnPanelCreate(TBuildLogForm, MainPanel);
      2: CategoryForm:= FormOnPanelCreate(TTestLogForm, MainPanel);
      3: CategoryForm:= FormOnPanelCreate(TControlListForm, MainPanel);
      4: CategoryForm:= FormOnPanelCreate(TMotorListForm, MainPanel);
      5: CategoryForm:= FormOnPanelCreate(TStoreForm, MainPanel);
      6: CategoryForm:= FormOnPanelCreate(TShipmentForm, MainPanel);
      7: CategoryForm:= FormOnPanelCreate(TReclamationForm, MainPanel);
      8: CategoryForm:= FormOnPanelCreate(TStatisticForm, MainPanel);
      9: CategoryForm:= FormOnPanelCreate(TRepairForm, MainPanel);
      10: CategoryForm:= FormOnPanelCreate(TReportForm, MainPanel);
      11: CategoryForm:= FormOnPanelCreate(TCalendarForm, MainPanel);
    end;

    if Assigned(CategoryForm) then
      CategoryForm.Show;

  finally
    Screen.Cursor:= crDefault;
  end;
end;

procedure TMainForm.ChangeUsedMotorList;
begin
  if DataBase.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, True,
    UsedNameIDs, UsedNames) then ViewUpdate;
end;

procedure TMainForm.ChangeUsedReceiverList;
begin
  if DataBase.ReceiverIDsAndNamesSelectedLoad(ReceiverNamesLabel, True,
    UsedReceiverIDs, UsedReceiverNames) then ViewUpdate;
end;

procedure TMainForm.DictionaryChoose(const ADictionary: Byte);
var
  IsOK: Boolean;
begin
  case ADictionary of
    1: IsOK:= DataBase.EditList('Наименования электродвигателей',
                              'MOTORNAMES', 'NameID', 'MotorName', False, True);
    2: IsOK:= DataBase.EditList('Грузополучатели',
                              'CARGORECEIVERS', 'ReceiverID', 'ReceiverName', True, True);
    3: IsOK:= DataBase.EditList('Элементы электродвигателей',
                              'RECLAMATIONDEFECTS', 'DefectID', 'DefectName', True, True);
    4: IsOK:= DataBase.EditTable('Причины возникновения неисправностей',
                              'RECLAMATIONREASONS', 'ReasonID',
                              ['ReasonName',  'ReasonColor'],
                              ['Причина',     'Цвет'       ],
                              [ctString,      ctColor      ],
                              [True,          True         ],
                              [300,           80           ],
                              [taLeftJustify, taCenter     ],
                              True, ['ReasonName']);
    5: IsOK:= DataBase.EditList('Заводы',
                              'RECLAMATIONFACTORIES', 'FactoryID', 'FactoryName', True, True);
    6: IsOK:= DataBase.EditList('Предприятия (депо)',
                              'RECLAMATIONPLACES', 'PlaceID', 'PlaceName', True, True);
  end;
  if IsOK then ViewUpdate;
end;

procedure TMainForm.SetNamesPanelsVisible(const AMotorNamesPanelVisible,
  AReceiverNamesPanelVisible: Boolean);
begin
  MotorNamesPanel.Visible:= AMotorNamesPanelVisible;
  ReceiverNamesPanel.Visible:= AReceiverNamesPanelVisible;
end;

end.

