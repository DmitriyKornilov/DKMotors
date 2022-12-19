unit UStatisticForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, rxctrls, DateTimePicker, DividerBevel, USQLite, DK_Vector,
  DK_DateUtils;

type

  { TStatisticForm }

  TStatisticForm = class(TForm)
    ChooseMotorNamesButton: TSpeedButton;
    CloseButton: TSpeedButton;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    DividerBevel4: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    MotorNamesLabel: TLabel;
    MotorNamesPanel: TPanel;
    NumberListCheckBox: TCheckBox;
    NumberListPanel: TPanel;
    OrderNumCheckBox: TCheckBox;
    Panel2: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    ReportPeriodPanel: TPanel;
    ToolPanel: TPanel;
    procedure ChooseMotorNamesButtonClick(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Label2Click(Sender: TObject);
    procedure MotorNamesLabelClick(Sender: TObject);
    procedure MotorNamesPanelClick(Sender: TObject);
  private
    UsedNameIDs: TIntVector;
    UsedNames: TStrVector;
    procedure ShowStatistic;
    procedure ChangeUsedMotorList;
  public

  end;

var
  StatisticForm: TStatisticForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TStatisticForm }

procedure TStatisticForm.FormClose(Sender: TObject;
  var CloseAction: TCloseAction);
begin
  MainForm.RxSpeedButton10.Down:= False;
  MainForm.StatisticForm:= nil;
  CloseAction:= caFree;
end;

procedure TStatisticForm.ChooseMotorNamesButtonClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TStatisticForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TStatisticForm.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= FirstDayInMonth(Date);
  SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, False, UsedNameIDs, UsedNames);
end;

procedure TStatisticForm.Label2Click(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TStatisticForm.MotorNamesLabelClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TStatisticForm.MotorNamesPanelClick(Sender: TObject);
begin
  ChangeUsedMotorList;
end;

procedure TStatisticForm.ShowStatistic;
begin

end;

procedure TStatisticForm.ChangeUsedMotorList;
begin
  if SQLite.NameIDsAndMotorNamesSelectedLoad(MotorNamesLabel, True, UsedNameIDs, UsedNames) then
    ShowStatistic;
end;

end.

