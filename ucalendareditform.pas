unit UCalendarEditForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons, DateTimePicker, DividerBevel, DK_Const, DK_DateUtils, DateUtils,
  USQLite, DK_Vector;

type

  { TCalendarEditForm }

  TCalendarEditForm = class(TForm)
    ButtonPanel: TPanel;
    CancelButton: TSpeedButton;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    SaveButton: TSpeedButton;
    procedure CancelButtonClick(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
  private

  public
    Year: Integer;
    DayDate: TDate;
  end;

var
  CalendarEditForm: TCalendarEditForm;

implementation

{$R *.lfm}

{ TCalendarEditForm }

procedure TCalendarEditForm.CancelButtonClick(Sender: TObject);
begin
  ModalResult:= mrCancel;
end;

procedure TCalendarEditForm.ComboBox1Change(Sender: TObject);
begin
  Label4.Visible:= (ComboBox1.ItemIndex=2) or (ComboBox1.ItemIndex=3);
  ComboBox2.Visible:= Label4.Visible;
end;

procedure TCalendarEditForm.DateTimePicker1Change(Sender: TObject);
begin
  DateTimePicker2.MinDate:= DateTimePicker1.Date;
end;

procedure TCalendarEditForm.FormShow(Sender: TObject);
begin
  DateTimePicker1.MinDate:= FirstDayInYear(Year);
  DateTimePicker1.MaxDate:= LastDayInYear(Year);
  if SameDate(DayDate, NULDATE) then //новый
    DateTimePicker1.Date:= DateTimePicker1.MinDate
  else  //редактирование
    DateTimePicker1.Date:= DayDate;

  DateTimePicker2.Date:= DateTimePicker1.Date;
  DateTimePicker2.MinDate:= DateTimePicker1.Date;
  DateTimePicker2.MaxDate:= DateTimePicker1.MaxDate;
end;

procedure TCalendarEditForm.SaveButtonClick(Sender: TObject);
var
  i, N, Status, SwapDay: Integer;
  Dates: TDateVector;
begin
  Status:= ComboBox1.ItemIndex+1;
  SwapDay:= Ord(ComboBox2.Visible)*(ComboBox2.ItemIndex);
  N:= DaysBetweenDates(DateTimePicker1.Date, DateTimePicker2.Date);
  Dates:= nil;
  VDim(Dates, N+1);
  for i:= 0 to N do
    Dates[i]:= IncDay(DateTimePicker1.Date, i);
  SQLite.WriteCalendarSpecDays(Dates, Status, SwapDay);

  ModalResult:= mrOK;
end;

end.

