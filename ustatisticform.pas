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
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    ExportButton: TRxSpeedButton;
    Label1: TLabel;
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
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure ShowStatistic;
  end;

var
  StatisticForm: TStatisticForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TStatisticForm }

procedure TStatisticForm.FormCreate(Sender: TObject);
begin
  MainForm.SetNamesPanelsVisible(True, False);
  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= FirstDayInMonth(Date);
end;

procedure TStatisticForm.ShowStatistic;
begin

end;


end.

