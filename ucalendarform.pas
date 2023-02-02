unit UCalendarForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, VirtualTrees, DividerBevel, DateTimePicker, fpspreadsheetgrid,
  DK_VSTTables, DK_SheetExporter, DK_DateUtils, UCalendar, USQLite, DK_Vector,
  fpsTypes, LCLType, StdCtrls, ComCtrls, rxctrls, UCalendarEditForm, DK_Const,
  DateUtils, USheetUtils;

type

  { TCalendarForm }

  TCalendarForm = class(TForm)
    AddDateButton: TSpeedButton;
    CancelCopyButton: TSpeedButton;
    CopyDateButton: TSpeedButton;
    CopyEditPanel: TPanel;
    CopyPanel: TPanel;
    DateEditPanel: TPanel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    DateTimePicker3: TDateTimePicker;
    DayPanel: TPanel;
    DelCopyButton: TSpeedButton;
    DelDateButton: TSpeedButton;
    EditDateButton: TSpeedButton;
    EditPanel: TPanel;
    ExportButton: TRxSpeedButton;
    CalendarGrid: TsWorksheetGrid;
    BeforeCountLabel: TLabel;
    Label1: TLabel;
    Label9: TLabel;
    WorkCountLabel: TLabel;
    NotWorkCountLabel: TLabel;
    OffDayCountLabel: TLabel;
    HolidayCountLabel: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    EndDateLabel: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    TotalCountLabel: TLabel;
    MainPanel: TPanel;
    LeftPanel: TPanel;
    CalcPanel: TPanel;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    DividerBevel1: TDividerBevel;
    PeriodPanel: TPanel;
    SaveCopyButton: TSpeedButton;
    SpinEdit1: TSpinEdit;
    Splitter2: TSplitter;
    TopPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    YearSpinEdit: TSpinEdit;
    Splitter1: TSplitter;
    ZoomCaptionLabel: TLabel;
    ZoomInButton: TSpeedButton;
    ZoomOutButton: TSpeedButton;
    ZoomPanel: TPanel;
    ZoomTrackBar: TTrackBar;
    ZoomValueLabel: TLabel;
    ZoomValuePanel: TPanel;
    procedure AddDateButtonClick(Sender: TObject);
    procedure CalendarGridDblClick(Sender: TObject);
    procedure CalendarGridMouseDown(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure CancelCopyButtonClick(Sender: TObject);
    procedure CopyDateButtonClick(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure DateTimePicker3Change(Sender: TObject);
    procedure DelCopyButtonClick(Sender: TObject);
    procedure DelDateButtonClick(Sender: TObject);
    procedure EditDateButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SaveCopyButtonClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure VT1DblClick(Sender: TObject);
    procedure VT1KeyUp(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure VT1MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure VT2KeyUp(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure VT2MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure YearSpinEditChange(Sender: TObject);
    procedure ZoomInButtonClick(Sender: TObject);
    procedure ZoomOutButtonClick(Sender: TObject);
    procedure ZoomTrackBarChange(Sender: TObject);
  private
    YearCalendar: TCalendar;
    VSTDays: TVSTTable;
    VSTCopy: TVSTTable;
    CalendarSheet: TCalendarSheet;
    ColorVector: TColorVector;

    SelectedDates: TDateVector;
    SelectedStatusStr: String;
    SelectedStatus: Integer;

    SpecDays: TCalendarSpecDays;

    IsCopyDates: Boolean;

    procedure LoadColors;
    procedure LoadSpecDays(const ASelectedIndex: Integer=-1);
    procedure CalcCalendar;
    procedure DrawCalendar;
    procedure RefreshCalendar;

    procedure CalcAndShowInfo;
    procedure CalcAndShowDaysCount;
    procedure CalcAndShowWorkDaysEnd;


    procedure LoadCopyList(const ASelectedIndex: Integer=-1);

    procedure SetEditButtonsEnabled;

    procedure DelSpecDay;
    procedure DelCopyDay;

    procedure SelectDayInGrid(const ADate: TDate);
    procedure SelectDayInList(const ADate: TDate);
    procedure ChangeListSelection;

    procedure OpenCalendarEditForm(const ADate: TDate);

    procedure BeginCopy;
    procedure EndCopy;


  public

  end;

var
  CalendarForm: TCalendarForm;

implementation

uses UMainForm;

{$R *.lfm}

{ TCalendarForm }

procedure TCalendarForm.FormCreate(Sender: TObject);
var
  W1, W2: Integer;
begin
  MainForm.SetNamesPanelsVisible(False, False);

  DateTimePicker1.Date:= Date;
  DateTimePicker2.Date:= Date;
  DateTimePicker3.Date:= Date;

  LoadColors;

  W1:= 110;
  W2:= 200;
  LeftPanel.Width:= W1+W2;

  CalendarSheet:= TCalendarSheet.Create(CalendarGrid.Worksheet, CalendarGrid);

  VSTDays:= TVSTTable.Create(VT1);
  VSTDays.CanSelect:= True;
  VSTDays.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  //DefaultVSTTablesSettings(VSTDays);
  VSTDays.AddColumn('Дата', W1);
  VSTDays.AddColumn('Статус', W2);
  VSTDays.Draw;

  VSTCopy:= TVSTTable.Create(VT2);
  VSTCopy.CanSelect:= True;
  //DefaultVSTTablesSettings(VSTCopy);
  VSTCopy.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTCopy.AddColumn('Дата', W1);
  VSTCopy.AddColumn('Статус', W2);
  VSTCopy.Draw;

  YearCalendar:= TCalendar.Create;
  YearSpinEdit.Value:= YearOfDate(Date);

  IsCopyDates:= False;
end;

procedure TCalendarForm.FormDestroy(Sender: TObject);
begin
  FreeAndNil(CalendarSheet);
  FreeAndNil(YearCalendar);
  FreeAndNil(VSTDays);
  FreeAndNil(VSTCopy);
end;

procedure TCalendarForm.FormShow(Sender: TObject);
begin
  RefreshCalendar;
end;

procedure TCalendarForm.SaveCopyButtonClick(Sender: TObject);
begin
  EndCopy;
end;

procedure TCalendarForm.SpinEdit1Change(Sender: TObject);
begin
  CalcAndShowWorkDaysEnd;
end;

procedure TCalendarForm.VT1DblClick(Sender: TObject);
var
  DayDate: TDate;
begin
  DayDate:= NULDATE;
  if VSTDays.IsSelected then
    DayDate:= SpecDays.Dates[VSTDays.SelectedIndex];
  OpenCalendarEditForm(DayDate);
end;

procedure TCalendarForm.VT1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE then
    DelSpecDay
  else if Key in [VK_UP, VK_DOWN] then
    ChangeListSelection;
end;

procedure TCalendarForm.VT1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ChangeListSelection;
end;

procedure TCalendarForm.VT2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE then
    DelCopyDay;
end;

procedure TCalendarForm.VT2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  DelCopyButton.Enabled:= VSTCopy.IsSelected;
end;

procedure TCalendarForm.ChangeListSelection;
var
  DayDate: TDate;
begin
  DayDate:= NULDATE;
  if VSTDays.IsSelected then
    DayDate:= SpecDays.Dates[VSTDays.SelectedIndex];

  SelectDayInGrid(DayDate);
  SetEditButtonsEnabled;
end;

procedure TCalendarForm.YearSpinEditChange(Sender: TObject);
begin
  RefreshCalendar;
end;

procedure TCalendarForm.ZoomInButtonClick(Sender: TObject);
begin
  ZoomTrackBar.Position:= ZoomTrackBar.Position + 5;
end;

procedure TCalendarForm.ZoomOutButtonClick(Sender: TObject);
begin
  ZoomTrackBar.Position:= ZoomTrackBar.Position - 5;
end;

procedure TCalendarForm.ZoomTrackBarChange(Sender: TObject);
begin
  ZoomValueLabel.Caption:= IntToStr(ZoomTrackBar.Position) + ' %';
  DrawCalendar;
end;

procedure TCalendarForm.LoadColors;
begin
  ColorVector:= nil;
  VDim(ColorVector, 8);
  ColorVector[HOLIDEY_COLOR_INDEX]:= COLORS_CALENDAR[HOLIDEY_COLOR_INDEX];
  ColorVector[OFFDAY_COLOR_INDEX]:= COLORS_CALENDAR[OFFDAY_COLOR_INDEX];
  ColorVector[BEFORE_COLOR_INDEX]:= COLORS_CALENDAR[BEFORE_COLOR_INDEX];
  ColorVector[WEEKDAY_COLOR_INDEX]:= COLORS_CALENDAR[WEEKDAY_COLOR_INDEX];
  ColorVector[MONTHNAME_COLOR_INDEX]:= COLOR_CALENDAR_MONTHNAME;
  ColorVector[DAYNAME_COLOR_INDEX]:= COLOR_CALENDAR_DAYNAME;
  ColorVector[HIGHLIGHT_COLOR_INDEX]:= COLOR_BACKGROUND_SELECTED;
end;

procedure TCalendarForm.LoadSpecDays(const ASelectedIndex: Integer = -1);
var
  BD, ED: TDate;
  StrDates, StrStatuses: TStrVector;
begin
  BD:= FirstDayInYear(YearSpinEdit.Value);
  ED:= LastDayInYear(YearSpinEdit.Value);
  SpecDays:= SQLite.LoadCalendarSpecDays(BD, ED);

  CalendarSpecDaysToStr(SpecDays, StrDates, StrStatuses);
  VSTDays.SetColumn('Дата', StrDates);
  VSTDays.SetColumn('Статус', StrStatuses);
  VSTDays.Draw;

  if (ASelectedIndex>=0) and (ASelectedIndex<=High(StrDates)) then
    VSTDays.Select(ASelectedIndex);
end;

procedure TCalendarForm.CalcCalendar;
var
  BD, ED: TDate;
begin
  BD:= FirstDayInYear(YearSpinEdit.Value);
  ED:= LastDayInYear(YearSpinEdit.Value);
  YearCalendar.Calc(BD, ED, SpecDays);
end;

procedure TCalendarForm.DrawCalendar;
begin
  CalendarSheet.Zoom(ZoomTrackBar.Position);
  CalendarSheet.Draw(YearCalendar, SelectedDates);
  CalendarSheet.UpdateColors(ColorVector);
end;

procedure TCalendarForm.RefreshCalendar;
begin
  SelectedDates:= nil;
  LoadSpecDays;
  CalcCalendar;
  DrawCalendar;
  CalcAndShowInfo;
end;

procedure TCalendarForm.CalcAndShowInfo;
begin
  CalcAndShowDaysCount;
  CalcAndShowWorkDaysEnd;
end;

procedure TCalendarForm.CalcAndShowDaysCount;
var
  Calendar: TCalendar;
  BD, ED: TDate;
begin
  BD:= DateTimePicker2.Date;
  ED:= DateTimePicker1.Date;

  if CompareDate(BD, ED)>0 then
  begin
    TotalCountLabel.Caption:= '0';
    WorkCountLabel.Caption:= '0';
    NotWorkCountLabel.Caption:= '0';
    OffDayCountLabel.Caption:= '0';
    HolidayCountLabel.Caption:= '0';
    BeforeCountLabel.Caption:= '0';
    Exit;
  end;

  Calendar:= SQLite.LoadCalendar(BD, ED);
  try
    TotalCountLabel.Caption:= IntToStr(Calendar.DaysCount);
    WorkCountLabel.Caption:= IntToStr(Calendar.WorkDaysCount);
    NotWorkCountLabel.Caption:= IntToStr(Calendar.NotWorkDaysCount);
    OffDayCountLabel.Caption:= IntToStr(Calendar.OffDaysCount);
    HolidayCountLabel.Caption:= IntToStr(Calendar.HoliDaysCount);
    BeforeCountLabel.Caption:= IntToStr(Calendar.BeforeDaysCount);
  finally
    FreeAndNil(Calendar);
  end;
end;

procedure TCalendarForm.CalcAndShowWorkDaysEnd;
var
  BD, ED: TDate;
  N: Integer;
begin
  N:= SpinEdit1.Value;
  BD:= DateTimePicker3.Date;
  ED:= SQLite.LoadWorkDaysPeriodEndDate(BD, N);
  EndDateLabel.Caption:= FormatDateTime('dd.mm.yyyy', ED);
end;

procedure TCalendarForm.SetEditButtonsEnabled;
begin
  DelDateButton.Enabled:= VSTDays.IsSelected;
  EditDateButton.Enabled:= DelDateButton.Enabled;
  CopyDateButton.Enabled:= DelDateButton.Enabled;
end;

procedure TCalendarForm.DelSpecDay;
var
  Ind: Integer;
begin
  if not VSTDays.IsSelected then Exit;
  Ind:= VSTDays.SelectedIndex;
  SQLite.DeleteCalendarSpecDay(SpecDays.Dates[Ind]);

  LoadSpecDays(Ind);
  CalcCalendar;
  DrawCalendar;

  ChangeListSelection;
end;

procedure TCalendarForm.DelCopyDay;
var
  Ind: Integer;
begin
  Ind:= VSTCopy.SelectedIndex;
  VDel(SelectedDates, Ind);

  LoadCopyList(Ind);
  DrawCalendar;

  DelCopyButton.Enabled:= VSTCopy.IsSelected;
end;

procedure TCalendarForm.LoadCopyList(const ASelectedIndex: Integer = -1);
var
  SelStatuses, SelDates: TStrVector;
begin
  SelStatuses:= nil;
  SelDates:= VDateToStr(SelectedDates);
  VDim(SelStatuses, Length(SelectedDates), SelectedStatusStr);
  VSTCopy.SetColumn('Дата', SelDates);
  VSTCopy.SetColumn('Статус', SelStatuses);
  VSTCopy.Draw;

  if (ASelectedIndex>=0) and (ASelectedIndex<=High(SelectedDates)) then
    VSTCopy.Select(ASelectedIndex);

  SaveCopyButton.Enabled:= not VIsNil(SelectedDates);
end;

procedure TCalendarForm.SelectDayInGrid(const ADate: TDate);
var
  Ind: Integer;
begin
  if IsCopyDates then
  begin
    Ind:= VIndexOfDate(SelectedDates, ADate);
    if Ind>=0 then
      VDel(SelectedDates, Ind)
    else
      VAppend(SelectedDates, ADate);
    LoadCopyList;
  end
  else begin
    if not SameDate(ADate, NULDATE) then
      VDim(SelectedDates, 1, ADate)
    else
      SelectedDates:= nil;
  end;
  DrawCalendar;

  //VT1.SetFocus;
end;

procedure TCalendarForm.SelectDayInList(const ADate: TDate);
var
  Ind: Integer;
begin
  if IsCopyDates then Exit;

  Ind:= VIndexOfDate(SpecDays.Dates, ADate);
  if Ind>=0 then
    VSTDays.Select(Ind)
  else
    VSTDays.UnSelect;
  SetEditButtonsEnabled;
end;

procedure TCalendarForm.OpenCalendarEditForm(const ADate: TDate);
var
  CalendarEditForm: TCalendarEditForm;
begin
  CalendarEditForm:= TCalendarEditForm.Create(CalendarForm);
  try
    CalendarEditForm.DayDate:= ADate;
    CalendarEditForm.Year:= YearSpinEdit.Value;
    if CalendarEditForm.ShowModal=mrOK then
    begin
      RefreshCalendar;
    end
    else begin
      SelectedDates:= nil;
      DrawCalendar;
    end;

  finally
    FreeAndNil(CalendarEditForm);
  end;
end;

procedure TCalendarForm.BeginCopy;
begin
  IsCopyDates:= True;
  EditPanel.Align:= alBottom;
  EditPanel.Visible:= False;
  CopyPanel.Align:= alClient;
  CopyPanel.Visible:= True;

  SelectedStatus:= SpecDays.Statuses[VSTDays.SelectedIndex];
  SelectedStatusStr:= DayStatusToStr(SelectedStatus);

  SelectedDates:= nil;
  DrawCalendar;
  VSTDays.UnSelect;
  SetEditButtonsEnabled;
end;

procedure TCalendarForm.EndCopy;
begin
  if not VIsNil(SelectedDates) then
  begin
    if IsCopyDates then
    begin
      SQLite.WriteCalendarSpecDays(SelectedDates, SelectedStatus);
      RefreshCalendar;
    end
    else begin
      SelectedDates:= nil;
      DrawCalendar;
    end;
    VSTCopy.ValuesClear;
  end;

  CopyPanel.Align:= alBottom;
  CopyPanel.Visible:= False;
  EditPanel.Align:= alClient;
  EditPanel.Visible:= True;

  IsCopyDates:= False;
end;



procedure TCalendarForm.CopyDateButtonClick(Sender: TObject);
begin
  BeginCopy;
end;

procedure TCalendarForm.DateTimePicker1Change(Sender: TObject);
begin
  CalcAndShowDaysCount;
end;

procedure TCalendarForm.DateTimePicker2Change(Sender: TObject);
begin
  CalcAndShowDaysCount;
end;

procedure TCalendarForm.DateTimePicker3Change(Sender: TObject);
begin
  CalcAndShowWorkDaysEnd;
end;

procedure TCalendarForm.DelCopyButtonClick(Sender: TObject);
begin
  DelCopyDay;
end;

procedure TCalendarForm.DelDateButtonClick(Sender: TObject);
begin
  DelSpecDay;
end;

procedure TCalendarForm.EditDateButtonClick(Sender: TObject);
begin
  OpenCalendarEditForm(SpecDays.Dates[VSTDays.SelectedIndex]);
end;

procedure TCalendarForm.AddDateButtonClick(Sender: TObject);
begin
  OpenCalendarEditForm(NULDATE);
end;

procedure TCalendarForm.CalendarGridDblClick(Sender: TObject);
var
  DayDate: TDate;
begin
  if CalendarSheet.GridToDate(CalendarGrid.Row, CalendarGrid.Col, DayDate) then
    OpenCalendarEditForm(DayDate);
end;

procedure TCalendarForm.CalendarGridMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  R,C: Integer;
  DayDate: TDate;
begin
  R:= 0;
  C:= 0;
  if Button=mbLeft then
    (Sender as TsWorksheetGrid).MouseToCell(X,Y,C,R);

  CalendarSheet.GridToDate(R, C, DayDate);
  SelectDayInGrid(DayDate);
  SelectDayInList(DayDate);
end;

procedure TCalendarForm.CancelCopyButtonClick(Sender: TObject);
begin
  IsCopyDates:= False;
  EndCopy;
end;

procedure TCalendarForm.ExportButtonClick(Sender: TObject);
var
  Exporter: TGridExporter;
begin
  Exporter:= TGridExporter.Create(CalendarGrid);
  try
    Exporter.PageSettings(spoLandScape, pfOnePage);
    Exporter.Save('Выполнено!');
  finally
    FreeAndNil(Exporter);
  end;
end;

end.

