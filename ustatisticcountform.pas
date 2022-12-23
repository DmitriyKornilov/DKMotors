unit UStatisticCountForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  fpspreadsheetgrid, DividerBevel, DateTimePicker, DK_DateUtils, SheetUtils,
  DK_Vector, DK_Matrix, USQLite;

type

  { TStatisticCountForm }

  TStatisticCountForm = class(TForm)
    CheckBox1: TCheckBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DividerBevel1: TDividerBevel;
    DividerBevel2: TDividerBevel;
    Grid1: TsWorksheetGrid;
    Label1: TLabel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    ReportPeriodPanel: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure RadioButton3Click(Sender: TObject);
  private
    Sheet: TStatisticReclamationCountSheet;

    TitleReasonIDs: TIntVector;
    TitleReasonNames: TStrVector;

    procedure DrawData;
  public

  end;

var
  StatisticCountForm: TStatisticCountForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TStatisticCountForm }

procedure TStatisticCountForm.FormCreate(Sender: TObject);
begin
  SQLite.KeyPickList('RECLAMATIONREASONS', 'ReasonID', 'ReasonName',
                     TitleReasonIDs, TitleReasonNames, True {TODO: не указано тоже нужно учесть!});
  Sheet:= TStatisticReclamationCountSheet.Create(Grid1, Length(TitleReasonNames));

  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= FirstDayInYear(Date);
end;

procedure TStatisticCountForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(Sheet);
end;

procedure TStatisticCountForm.FormShow(Sender: TObject);
begin
  DrawData;
end;

procedure TStatisticCountForm.RadioButton1Click(Sender: TObject);
begin
  DrawData;
end;

procedure TStatisticCountForm.RadioButton2Click(Sender: TObject);
begin
  DrawData;
end;

procedure TStatisticCountForm.RadioButton3Click(Sender: TObject);
begin
  DrawData;
end;

procedure TStatisticCountForm.DrawData;
var
  BD, ED: TDate;
  ReportName, FirstColName, MotorNames: String;
  NameValues: TStrVector;
  CountValues: TIntVector;
  ReasonCountValues: TIntMatrix;
begin
  ED:= DateTimePicker1.Date;
  BD:= DateTimePicker2.Date;
  if BD>ED then Exit;

  ReportName:= EmptyStr;
  if RadioButton1.Checked then
  begin
    SQLite.ReclamationTotalWithReasonsLoad(BD, ED, MainForm.UsedNameIDs, TitleReasonIDs,
                   NameValues, CountValues, ReasonCountValues);
    FirstColName:= 'Наименование';
    ReportName:= 'общее количество';
  end
  else if RadioButton2.Checked then
  begin
    SQLite.ReclamationDefectsWithReasonsLoad(BD, ED, MainForm.UsedNameIDs, TitleReasonIDs,
                   NameValues, CountValues, ReasonCountValues);
    FirstColName:= 'Неисправный элемент';
    ReportName:= 'распределение по неисправным элементам';
  end
  else if RadioButton3.Checked then
  begin
    SQLite.ReclamationPlacesWithReasonsLoad(BD, ED, MainForm.UsedNameIDs, TitleReasonIDs,
                   NameValues, CountValues, ReasonCountValues);
    FirstColName:= 'Предприятие';
    ReportName:= 'распределение по предприятиям';
  end;

  MotorNames:= VVectorToStr(MainForm.UsedNames, ', ');
  Sheet.Draw(BD, ED, ReportName, MotorNames, FirstColName,
             NameValues, TitleReasonNames, CountValues, ReasonCountValues, True);
end;

end.

