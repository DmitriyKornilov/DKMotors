unit UCalendarForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Spin,
  Buttons, VirtualTrees, DividerBevel, fpspreadsheetgrid,
  DK_VSTTables, DK_SheetExporter, DK_DateUtils, UCalendar, USQLite,
  DK_Vector, fpsTypes, LCLType, rxctrls, UCalendarEditForm, DK_Const,
  DateUtils, USheetUtils;

type

  { TCalendarForm }

  TCalendarForm = class(TForm)
    ExportButton: TRxSpeedButton;
    SaveCopyButton: TSpeedButton;
    AddDateButton: TSpeedButton;
    DelCopyButton: TSpeedButton;
    DelDateButton: TSpeedButton;
    CopyDateButton: TSpeedButton;
    CancelCopyButton: TSpeedButton;
    EditDateButton: TSpeedButton;
    CopyEditPanel: TPanel;
    DividerBevel1: TDividerBevel;
    CalendarGrid: TsWorksheetGrid;
    DateEditPanel: TPanel;
    CopyPanel: TPanel;
    TopPanel: TPanel;
    LeftPanel: TPanel;
    GridPanel: TPanel;
    EditPanel: TPanel;
    VT1: TVirtualStringTree;
    VT2: TVirtualStringTree;
    YearSpinEdit: TSpinEdit;
    Splitter1: TSplitter;
    procedure AddDateButtonClick(Sender: TObject);
    procedure CalendarGridDblClick(Sender: TObject);
    procedure CalendarGridMouseDown(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure CancelCopyButtonClick(Sender: TObject);
    procedure CopyDateButtonClick(Sender: TObject);
    procedure DelCopyButtonClick(Sender: TObject);
    procedure DelDateButtonClick(Sender: TObject);
    procedure EditDateButtonClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SaveCopyButtonClick(Sender: TObject);
    procedure VT1DblClick(Sender: TObject);
    procedure VT1KeyUp(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure VT1MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure VT2KeyUp(Sender: TObject; var Key: Word; {%H-}Shift: TShiftState);
    procedure VT2MouseUp(Sender: TObject; {%H-}Button: TMouseButton;
      {%H-}Shift: TShiftState; {%H-}X, {%H-}Y: Integer);
    procedure YearSpinEditChange(Sender: TObject);
  private
    YearCalendar: TCalendar;
    VSTDays: TVSTTable;
    VSTCopy: TVSTTable;
    CalendarSheet: TCalendarSheet;
    ColorVector: TColorVector;

    SelectedDates: TDateVector;
    SelectedStatusStr, SelectedSwapDayStr: String;
    SelectedStatus, SelectedSwapDay: Integer;

    SpecDays: TCalendarSpecDays;

    IsCopyDates: Boolean;

    procedure LoadColors;
    procedure LoadSpecDays(const ASelectedIndex: Integer=-1);
    procedure CalcCalendar;
    procedure DrawCalendar;
    procedure RefreshCalendar;

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
  W1, W2, W3: Integer;

  procedure CalcWidths;
  //const
  //  DeltaW = 55;
  //var
  //  FName: String;
  //  FSize: Single;
  //  x: Integer;
  begin
    //LoadFontFromControl(CalendarGrid, FName, FSize);
    //x:= Round(FSize);
    //if x<FSize then
    //  FSize:= x + 1
    //else
    //  FSize:= x;
    //
    //W1:= SWidth('00.00.0000', FName, FSize) + DeltaW;
    //W2:= SWidth('предпраздничный', FName, FSize) + DeltaW;
    //W3:= SWidth('Заменяемый день', FName, FSize) + DeltaW;
    W1:= 90;
    W2:= 150;
    W3:= 150;
  end;

begin
  MainForm.SetNamesPanelsVisible(False, False);

  LoadColors;

  //DefaultGridSettings(CalendarGrid);
  CalcWidths;
  LeftPanel.Width:= W1+W2+W3;

  CalendarSheet:= TCalendarSheet.Create(CalendarGrid.Worksheet, CalendarGrid);

  VSTDays:= TVSTTable.Create(VT1);
  VSTDays.CanSelect:= True;
  VSTDays.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  //DefaultVSTTablesSettings(VSTDays);
  VSTDays.AddColumn('Дата', W1);
  VSTDays.AddColumn('Статус', W2);
  VSTDays.AddColumn('Заменяемый день');
  VSTDays.Draw;

  VSTCopy:= TVSTTable.Create(VT2);
  VSTCopy.CanSelect:= True;
  //DefaultVSTTablesSettings(VSTCopy);
  VSTCopy.SelectedBGColor:= COLOR_BACKGROUND_SELECTED;
  VSTCopy.AddColumn('Дата', W1);
  VSTCopy.AddColumn('Статус', W2);
  VSTCopy.AddColumn('Заменяемый день');
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
  StrDates, StrStatuses, StrSwapDays: TStrVector;
begin
  BD:= FirstDayInYear(YearSpinEdit.Value);
  ED:= LastDayInYear(YearSpinEdit.Value);
  SpecDays:= SQLite.LoadCalendarSpecDays(BD, ED);

  CalendarSpecDaysToStr(SpecDays, StrDates, StrStatuses, StrSwapDays);
  VSTDays.SetColumn('Дата', StrDates);
  VSTDays.SetColumn('Статус', StrStatuses);
  VSTDays.SetColumn('Заменяемый день', StrSwapDays);
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
  CalendarSheet.Draw(YearCalendar, SelectedDates);
  CalendarSheet.UpdateColors(ColorVector);
end;

procedure TCalendarForm.RefreshCalendar;
begin
  SelectedDates:= nil;
  LoadSpecDays;
  CalcCalendar;
  DrawCalendar;
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
  SelStatuses, SelSwapdays, SelDates: TStrVector;
begin
  SelStatuses:= nil;
  SelSwapdays:= nil;
  SelDates:= VDateToStr(SelectedDates);
  VDim(SelStatuses, Length(SelectedDates), SelectedStatusStr);
  VDim(SelSwapdays, Length(SelectedDates), SelectedSwapDayStr);
  VSTCopy.SetColumn('Дата', SelDates);
  VSTCopy.SetColumn('Статус', SelStatuses);
  VSTCopy.SetColumn('Заменяемый день', SelSwapdays);
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
  SelectedSwapDay:= SpecDays.SwapDays[VSTDays.SelectedIndex];
  SelectedStatusStr:= DayStatusToStr(SelectedStatus);
  SelectedSwapDayStr:= SwapDayToStr(SelectedSwapDay);

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
      SQLite.WriteCalendarSpecDays(SelectedDates, SelectedStatus, SelectedSwapDay);
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

