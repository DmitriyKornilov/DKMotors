unit UCalendar;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DateUtils, Graphics, fpspreadsheet, fpspreadsheetgrid, fpstypes,
  DK_Vector, DK_Const, DK_DateUtils, DK_StrUtils, DK_SheetWriter, DK_SheetConst;

const
  {статус дня}
  DAY_STATUS_UNKNOWN = 0; //0 - неизвестный
  DAY_STATUS_HOLIDAY = 1; //1 - праздничный
  DAY_STATUS_OFFDAY  = 2; //2 - выходной
  DAY_STATUS_BEFORE  = 3; //3 - предпраздничный (рабочий с сокращенным временем)
  DAY_STATUS_WEEKDAY = 4; //4 - будний (обычный рабочий)

  DAY_STATUS_NAMES: array [0..4] of String = ('неизвестный',
                                              'праздничный',
                                              'выходной',
                                              'предпраздничный',
                                              'рабочий');

  COLOR_GRAY        = $00E0E0E0;
  COLOR_BLACK       = $00000000;
  COLOR_WHITE       = $00FFFFFF;
  COLOR_ORANGE      = $0097CBFF;
  COLOR_GREEN       = $00CCE3CC;

  COLOR_CALENDAR_MONTHNAME = COLOR_WHITE;//COLOR_BEIGE; //цвет ячейки с названием месяца
  COLOR_CALENDAR_DAYNAME   = COLOR_WHITE;//COLOR_GRAY;  //цвет ячеек с названиями дней недели
  COLOR_CALENDAR_QUARTER   = COLOR_GRAY;
  COLOR_CALENDAR_HALFYEAR  = COLOR_GREEN;
  COLOR_CALENDAR_YEAR      = COLOR_ORANGE;

  COLORS_CALENDAR: array [0..4] of Integer = (COLOR_BLACK,   //неизвестный
                                              COLOR_ORANGE,  //праздничный
                                              COLOR_GREEN,   //выходной
                                              COLOR_GRAY, //COLOR_VIOLET,  //предпраздничный
                                              COLOR_WHITE);  //рабочий

  MAIN_COLOR_INDEX      = TRANSPARENT_COLOR_INDEX;  //scTransparent
  HOLIDEY_COLOR_INDEX   = 1;
  OFFDAY_COLOR_INDEX    = 2;
  BEFORE_COLOR_INDEX    = 3;
  WEEKDAY_COLOR_INDEX   = 4;
  MONTHNAME_COLOR_INDEX = 5;
  DAYNAME_COLOR_INDEX   = 6;
  HIGHLIGHT_COLOR_INDEX = 7;
  QUARTER_COLOR_INDEX   = 8;
  HALFYEAR_COLOR_INDEX  = 9;
  YEAR_COLOR_INDEX      = 10;

  SHEET_FONT_NAME = 'Arial';
  SHEET_FONT_SIZE = 9;

  {преобразование целочисленного представления рабочих часов в фактическое дробное}
  //кол-во знаков после запятой в рабочих часах
  FRACTION_DIGITS_IN_WORKHOURS = 2;
  //10^FRACTION_DIGITS_IN_WORKHOURS - делитель целой части часов в дробную
  WORKHOURS_DENOMINATOR        = 100 ;
  //кол-во часов, на котрое уменьшается рабочая смена в предпраздничный день
  REDUCE_HOURS_COUNT_IN_BEFORE = 1;

type

  {Рабочие часы}
  TWorkHours = class (TObject)
    protected
      FTotal              : TIntVector; //вектор фактических часов (общее кол-во за смену)
      FNight              : TIntVector; //вектор ночных часов
      function GetSumTotal: Integer;
      function GetSumNight: Integer;
      procedure SetTotal(const AVector: TIntVector);
      procedure SetNight(const AVector: TIntVector);
    public
      constructor Create;
      destructor  Destroy; override;
      procedure Clear;
      procedure Add(const ATotal: Integer; const ANight: Integer = 0);
      procedure Copy(const ADestination: TWorkHours);
      property Total: TIntVector read FTotal write SetTotal;
      property Night: TIntVector read FNight write SetNight;
      property SumTotal: Integer read GetSumTotal;  //сумма отработанных часов
      property SumNight: Integer read GetSumNight;  //сумма ночных часов
  end;

  {Простой календарь}

  { TSimpleCalendar }

  TSimpleCalendar = class (TObject)
    protected
      FCalculated         : Boolean;     //рассчитан ли календарь
      FDates              : TDateVector; //вектор дат
      FDayStatuses        : TIntVector;  //статус дня: 2 - выходной 4 - будний (обычный рабочий)
      FDayNumsInWeek      : TIntVector;  //номер дня в неделе
      FWeekNumsInMonth    : TIntVector;  //номер недели в месяце
    private
      function GetIsWorkDay(const AIndex: Integer): Boolean;
      function GetWeekDaysCount: Integer;
      function GetOffDaysCount: Integer;
      function GetBeginDate: TDate;
      function GetDaysCount: Integer;
      function GetEndDate: TDate;
    public
      constructor Create;
      destructor  Destroy; override;
      procedure Clear;
      procedure Calc(const ABeginDate, AEndDate: TDate);
      property Calculated: Boolean read FCalculated;
      property BeginDate: TDate read GetBeginDate; //начальная дата
      property EndDate: TDate read GetEndDate;  //конечная дата
      property DaysCount: Integer read GetDaysCount; //кол-во дней
      property Dates: TDateVector read FDates;
      property DayStatuses: TIntVector read FDayStatuses;
      property DayNumsInWeek: TIntVector read FDayNumsInWeek;
      property WeekNumsInMonth: TIntVector read FWeekNumsInMonth;
      property WeekDaysCount: Integer read GetWeekDaysCount;//кол-во будних дней
      property OffDaysCount: Integer read GetOffDaysCount;  //кол-во выходных дней

      property IsWorkDay[const AIndex: Integer]: Boolean read GetIsWorkDay;
  end;

  {Корректировки производственного календаря}
    TCalendarSpecDays = record
      Dates   : TDateVector;
      Statuses: TIntVector;
    end;

const
    EmptyCalendarSpecDays: TCalendarSpecDays =
      (Dates   : nil;
       Statuses: nil;
      );

  function DayDateToStr(const ADate: TDate): String;
  function DayStatusToStr(const AStatus: Integer): String;
  procedure CalendarSpecDaysToStr(const ASpecDays: TCalendarSpecDays;
                                  out ADates, AStatuses: TStrVector);

type

  {Производственный календарь}
  TCalendar = class (TSimpleCalendar)
    protected
      FSwapDays : TIntVector;  //заменяемый день для особого дня со статусом "рабочий"
    private
      function GetHoliDaysCount: Integer;
      function GetBeforeDaysCount: Integer;
      function GetWorkDaysCount: Integer;
      function GetNotWorkDaysCount: Integer;
    public
      constructor Create;
      destructor  Destroy; override;
      procedure Clear;
      procedure Calc(const ABeginDate, AEndDate: TDate; const ASpecDays: TCalendarSpecDays);
      function Cut(const ABeginDate, AEndDate: TDate; var ACutCalendar: TCalendar): Boolean;
      function SumWorkHoursInt(const AHoursInWeek: Byte): Integer; //сумма рабочих часов в зависимости от кол-ва часов в неделю AHoursInWeek (40, 36, 24) в целочисленном формате
      function SumWorkHoursFrac(const AHoursInWeek: Byte): Double; //сумма рабочих часов в зависимости от кол-ва часов в неделю AHoursInWeek (40, 36, 24) в дробном формате
      property SwapDays: TIntVector read FSwapDays;
      property HoliDaysCount: Integer read GetHoliDaysCount; //кол-во праздничных дней
      property BeforeDaysCount: Integer read GetBeforeDaysCount;   //кол-во предпраздничных дней
      property WorkDaysCount: Integer read GetWorkDaysCount; //кол-во рабочих дней (будни+предпраздничные)
      property NotWorkDaysCount: Integer read GetNotWorkDaysCount; //кол-во нерабочих дней (выходные+праздничные)
  end;

  { TCalendarSheet }

  TCalendarSheet = class (TObject)
  private
  const
    RESUME_FIRST_ROW = 15;
    RESUME_FIRST_COL = 25;
    LEGEND_FIRST_ROW = 3;
    LEGEND_FIRST_COL = RESUME_FIRST_COL;
    MONTH_FIRST_ROWS:  array [1..12] of Byte = (3,3,3,12,12,12,21,21,21,30,30,30);
    MONTH_FIRST_COLS:  array [1..12] of Byte = (1,9,17,1,9,17,1,9,17,1,9,17);
    MONTH_RESUME_ROWS: array [1..12] of Byte = (RESUME_FIRST_ROW+4, RESUME_FIRST_ROW+5, RESUME_FIRST_ROW+6,
                                                RESUME_FIRST_ROW+8, RESUME_FIRST_ROW+9, RESUME_FIRST_ROW+10,
                                                RESUME_FIRST_ROW+13, RESUME_FIRST_ROW+14, RESUME_FIRST_ROW+15,
                                                RESUME_FIRST_ROW+17, RESUME_FIRST_ROW+18, RESUME_FIRST_ROW+19);
    QUARTER_RESUME_ROWS: array [1..4] of Byte = (RESUME_FIRST_ROW+7,
                                                 RESUME_FIRST_ROW+11,
                                                 RESUME_FIRST_ROW+16,
                                                 RESUME_FIRST_ROW+20);
    HALF_RESUME_ROWS: array [1..2] of Byte = (RESUME_FIRST_ROW+12,
                                              RESUME_FIRST_ROW+21);
    YEAR_RESUME_ROW = RESUME_FIRST_ROW+22;
  var
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;
    FCalendar: TCalendar;
    FYear: Word;
    FGridLineColor: TColor;
    FHighLightDays: TDateVector;
    FRowHeight: Integer;
    FColorVector: TColorVector;
    procedure DrawCaption;
    procedure DrawLegend;
    procedure DrawResumeLine(const R, C, AColorIndex: Integer; const ACalendar: TCalendar);
    procedure DrawResumeTableCaption;
    procedure DrawMonth(const AMonth: Byte);
    procedure DrawQuarter(const AQuarter: Byte);
    procedure DrawHalf(const AHalf: Byte);
    procedure DrawYear;
  public
    constructor Create(const AWorksheet: TsWorksheet; const AGrid: TsWorksheetGrid = nil);
    destructor  Destroy; override;
    procedure Zoom(const APercents: Integer);
    procedure Draw(const AYearCalendar: TCalendar; const AHighLightDays: TDateVector);
    procedure UpdateColors(const AColorVector: TColorVector);
    procedure ClearColors;
    function GridToMonth(const ARow, ACol: Integer; out ADayInWeek, AWeekInMonth, AMonth: Integer): Boolean;
    function GridToDate(const ARow, ACol: Integer; out ADate: TDate): Boolean;
    function DateToGrid(const ADate: TDate; out ARow, ACol: Integer): Boolean;
 end;


  function WeekHoursToWorkHoursInt(const AHoursInWeek: Byte): Integer;  //кол-во часов за смену

  {MonthGridToDate расчет даты дня, взятого из ячейки [ARow, ACol]
   таблицы месячного календаря 6х7 для месяца AMonth;
  false - если вне месяца}
  function MonthGridToDate(const AWeekInMonth{ARow}, ADayInWeek{ACol}, AMonth, AYear: Integer;
                           out ADate: TDate): Boolean;

implementation

{TWorkHours}

constructor TWorkHours.Create;
begin
  inherited Create;
  Clear;
end;

destructor TWorkHours.Destroy;
begin
  inherited Destroy;
end;

procedure TWorkHours.Clear;
begin
  FTotal:= nil;
  FNight:= nil;
end;

procedure TWorkHours.SetTotal(const AVector: TIntVector);
begin
  FTotal:= VCut(AVector);
end;

procedure TWorkHours.SetNight(const AVector: TIntVector);
begin
  FNight:= VCut(AVector);
end;

procedure TWorkHours.Add(const ATotal: Integer; const ANight: Integer = 0);
begin
  VAppend(FTotal, ATotal);
  VAppend(FNight, ANight);
end;

procedure TWorkHours.Copy(const ADestination: TWorkHours);
begin
  ADestination.Total:= VCut(FTotal);
  ADestination.Night:= VCut(FNight);
end;

function TWorkHours.GetSumTotal: Integer;
begin
  GetSumTotal:= VSum(FTotal);
end;

function TWorkHours.GetSumNight: Integer;
begin
  GetSumNight:= VSum(FNight);
end;

{TSimpleCalendar}

constructor TSimpleCalendar.Create;
begin
  inherited Create;
end;

destructor TSimpleCalendar.Destroy;
begin
  inherited Destroy;
end;

procedure TSimpleCalendar.Clear;
begin
  FCalculated:= False;
  FDates:= nil;
  FDayStatuses:= nil;
  FDayNumsInWeek:= nil;
  FWeekNumsInMonth:= nil;
end;

procedure TSimpleCalendar.Calc(const ABeginDate, AEndDate: TDate);
var
  i, DC: Integer;
begin
  Clear;
  DC:= DaysBetweenDates(ABeginDate, AEndDate)+1;
  VDim(FDates, DC);
  VDim(FDayNumsInWeek, DC);
  VDim(FWeekNumsInMonth, DC);
  VDim(FDayStatuses, DC);
  for i:= 0 to DC - 1 do
  begin
      FDates[i]:= IncDay(ABeginDate, i);
      FDayNumsInWeek[i]:= DayNumberInWeek(FDates[i]);
      FWeekNumsInMonth[i]:= WeekNumberInMonth(FDates[i]);
      if FDayNumsInWeek[i]<6 then
        FDayStatuses[i]:= DAY_STATUS_WEEKDAY
      else
        FDayStatuses[i]:= DAY_STATUS_OFFDAY;
  end;
  FCalculated:= True;
end;

function TSimpleCalendar.GetBeginDate: TDate;
begin
  if not VIsNil(FDates) then
    Result:= FDates[0]
  else
    Result:= 0;
end;

function TSimpleCalendar.GetDaysCount: Integer;
begin
  Result:= Length(FDates);
end;

function TSimpleCalendar.GetEndDate: TDate;
begin
  if not VIsNil(FDates) then
    Result:= FDates[High(FDates)]
  else
    Result:= 0;
end;

function TSimpleCalendar.GetIsWorkDay(const AIndex: Integer): Boolean;
begin
  Result:= (FDayStatuses[AIndex]=DAY_STATUS_WEEKDAY) or
           (FDayStatuses[AIndex]=DAY_STATUS_BEFORE);
end;

function TSimpleCalendar.GetWeekDaysCount: Integer;
begin
  Result:= VCountIf(FDayStatuses, DAY_STATUS_WEEKDAY);
end;

function TSimpleCalendar.GetOffDaysCount: Integer;
begin
  Result:= VCountIf(FDayStatuses, DAY_STATUS_OFFDAY);
end;

{TCalendarSpecDays utils}

function DayDateToStr(const ADate: TDate): String;
begin
  Result:= FormatDateTime('dd.mm.yyyy', ADate);
end;

function DayStatusToStr(const AStatus: Integer): String;
begin
  Result:= DAY_STATUS_NAMES[AStatus];
end;

procedure CalendarSpecDaysToStr(const ASpecDays: TCalendarSpecDays;
                            out ADates, AStatuses: TStrVector);
var
  i: Integer;
begin
  ADates:= nil;
  AStatuses:= nil;
  if VIsNil(ASpecDays.Dates) then Exit;

  VDim(ADates, Length(ASpecDays.Dates));
  VDim(AStatuses, Length(ASpecDays.Dates));
  for i:= 0 to High(ADates) do
  begin
    ADates[i]:= DayDateToStr(ASpecDays.Dates[i]);
    AStatuses[i]:= DayStatusToStr(ASpecDays.Statuses[i]);
  end;
end;

{TCalendar}

constructor TCalendar.Create;
begin
  inherited Create;
end;

destructor TCalendar.Destroy;
begin
  inherited Destroy;
end;

procedure TCalendar.Clear;
begin
  inherited Clear;
  FSwapDays:= nil;
end;

procedure TCalendar.Calc(const ABeginDate, AEndDate: TDate;
  const ASpecDays: TCalendarSpecDays);
var
  i, n: Integer;
begin
  Clear;
  inherited Calc(ABeginDate, AEndDate);
  FCalculated:= False;
  for i:= 0 to High(ASpecDays.Dates) do
  begin
    n:= DaysBetweenDates(BeginDate, ASpecDays.Dates[i]);
    //добавляем только даты, входящие в период календаря
    if (n>=0) and (n<DaysCount) then
      FDayStatuses[n]:= ASpecDays.Statuses[i];
  end;
  FCalculated:= True;
end;

function TCalendar.Cut(const ABeginDate, AEndDate: TDate;
                       var ACutCalendar: TCalendar): Boolean;
var
  BD, ED: TDate;
  I1, I2: Integer;
begin
  Result:= False;
  //проверяем, рассчиатн ли исходный календарь
  if not FCalculated then Exit;
  //проверяем, пересекаются ли периоды обрезки и исходного календаря, запоминаем период пересечения
  if not IsPeriodIntersect(ABeginDate, AEndDate, BeginDate, EndDate, BD, ED) then Exit;
  //определяем индексы среза
  I1:= DaysBetweenDates(BeginDate, BD);
  I2:= DaysBetweenDates(BeginDate, ED);
  //проверяем на существование календаря-среза, создаем при отсутствии
  if not Assigned(ACutCalendar) then
    ACutCalendar:= TCalendar.Create
  else
    ACutCalendar.Clear;
  //заполняем данные
  ACutCalendar.FDates:= VCut(FDates, I1, I2);
  ACutCalendar.FSwapDays:= VCut(FSwapDays, I1, I2);
  ACutCalendar.FDayStatuses:= VCut(FDayStatuses, I1, I2);
  ACutCalendar.FDayNumsInWeek:= VCut(FDayNumsInWeek, I1, I2);
  ACutCalendar.FWeekNumsInMonth:= VCut(FWeekNumsInMonth, I1, I2) ;
  ACutCalendar.FCalculated:= True;
  Result:= True;
end;

function TCalendar.SumWorkHoursInt(const AHoursInWeek: Byte): Integer; //сумма рабочих часов в зависимости от кол-ва часов в неделю AHoursInWeek (40, 36, 24)
var
  x: Integer;
begin
  Result:= 0;
  if AHoursInWeek>0 then
  begin
    x:= WeekHoursToWorkHoursInt(AHoursInWeek);
    Result:= x*WeekDaysCount +
            (x - WORKHOURS_DENOMINATOR*REDUCE_HOURS_COUNT_IN_BEFORE)*BeforeDaysCount;
  end;
end;

function TCalendar.SumWorkHoursFrac(const AHoursInWeek: Byte): Double; //сумма рабочих часов в зависимости от кол-ва часов в неделю AHoursInWeek (40, 36, 24) в дробном формате
begin
  Result:= 0;
  if AHoursInWeek>0 then
    Result:= AHoursInWeek*WorkDaysCount/5   -
             REDUCE_HOURS_COUNT_IN_BEFORE*BeforeDaysCount;
end;

function TCalendar.GetHoliDaysCount: Integer;
begin
  Result:= VCountIf(FDayStatuses, DAY_STATUS_HOLIDAY);
end;

function TCalendar.GetBeforeDaysCount: Integer;
begin
  Result:= VCountIf(FDayStatuses, DAY_STATUS_BEFORE);
end;

function TCalendar.GetWorkDaysCount: Integer;
begin
  Result:= WeekDaysCount + BeforeDaysCount;
end;

function TCalendar.GetNotWorkDaysCount: Integer;
begin
  Result:= HolidaysCount + OffDaysCount;
end;

{TCalendarSheet----------------------------------------------------------------}

procedure TCalendarSheet.DrawCaption;
begin
  FWriter.SetFont(FFontName, FFontSize+3, [fssBold], scBlack);
  FWriter.WriteText(1, 1, 1, FWriter.ColCount, 'ПРОИЗВОДСТВЕННЫЙ КАЛЕНДАРЬ НА ' +
               IntToStr(FYear) + ' ГОД');
end;

procedure TCalendarSheet.DrawLegend;
var
  R,C: Integer;

  procedure DrawLegendLine(const ARow, ACol, AColorIndex: Integer;
                           const ALegendValue: String);
  begin
    FWriter.WriteText(ARow, ACol, EmptyStr, cbtOuter);
    FWriter.AddCellBGColorIndex(ARow, ACol, AColorIndex);
    FWriter.WriteText(ARow, ACol+1, ARow, FWriter.ColCount, ALegendValue, cbtOuter);
    FWriter.DrawBorders(ARow, ACol, ARow, ACol+1, cbtAll);
  end;

begin
  FWriter.SetFont(FFontName, FFontSize, [], scBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  R:= LEGEND_FIRST_ROW;
  C:= LEGEND_FIRST_COL;
  DrawLegendLine(R, C, HOLIDEY_COLOR_INDEX, 'Нерабочий праздничный день');
  R:= R+1;
  DrawLegendLine(R, C, OFFDAY_COLOR_INDEX, 'Нерабочий выходной день');
  R:= R+1;
  DrawLegendLine(R, C, BEFORE_COLOR_INDEX, 'Рабочий предпраздничный (сокращенный) день');
  R:= R+1;
  DrawLegendLine(R, C, WEEKDAY_COLOR_INDEX, 'Рабочий день');
  FWriter.SetAlignmentDefault;
end;

procedure TCalendarSheet.DrawResumeTableCaption;
var
  R,C: Integer;
begin
  R:= RESUME_FIRST_ROW;
  C:= RESUME_FIRST_COL;
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  FWriter.WriteText(R, C, R+3, C, 'Период', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, MONTHNAME_COLOR_INDEX);
  C:= C+1;
  FWriter.WriteText(R, C, R, C+3, 'Количество дней', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, MONTHNAME_COLOR_INDEX);
  FWriter.WriteText(R+1, C, R+3, C, 'Кален-'+ SYMBOL_BREAK + 'дарных', cbtOuter);
  FWriter.AddCellBGColorIndex(R+1, C, MONTHNAME_COLOR_INDEX);
  C:= C+1;
  FWriter.WriteText(R+1, C, R+3, C, 'Рабочих', cbtOuter);
  FWriter.AddCellBGColorIndex(R+1, C, MONTHNAME_COLOR_INDEX);
  C:= C+1;
  FWriter.WriteText(R+1, C, R+3, C, 'Выход-'+ SYMBOL_BREAK + 'ных' , cbtOuter);
  FWriter.AddCellBGColorIndex(R+1, C, MONTHNAME_COLOR_INDEX);
  C:= C+1;
  FWriter.WriteText(R+1, C, R+3, C, 'Празд-'+ SYMBOL_BREAK + 'ничных', cbtOuter);
  FWriter.AddCellBGColorIndex(R+1, C, MONTHNAME_COLOR_INDEX);
  C:= C+1;
  FWriter.WriteText(R, C, R, C+2, 'Рабочее время (часов)', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, MONTHNAME_COLOR_INDEX);
  FWriter.WriteText(R+1, C, R+3, C, '40-часовая' + SYMBOL_BREAK + 'рабочая' + SYMBOL_BREAK + 'неделя', cbtOuter);
  FWriter.AddCellBGColorIndex(R+1, C, MONTHNAME_COLOR_INDEX);
  C:= C+1;
  FWriter.WriteText(R+1, C, R+3, C, '36-часовая' + SYMBOL_BREAK + 'рабочая' + SYMBOL_BREAK + 'неделя', cbtOuter);
  FWriter.AddCellBGColorIndex(R+1, C, MONTHNAME_COLOR_INDEX);
  C:= C+1;
  FWriter.WriteText(R+1, C, R+3, C, '24-часовая' + SYMBOL_BREAK + 'рабочая' + SYMBOL_BREAK + 'неделя', cbtOuter);
  FWriter.AddCellBGColorIndex(R+1, C, MONTHNAME_COLOR_INDEX);
  R:= RESUME_FIRST_ROW+4;
  C:= RESUME_FIRST_COL;
  FWriter.SetFont(FFontName, FFontSize, [], scBlack);
  FWriter.WriteText(R, C,   'Январь', cbtOuter);
  FWriter.WriteText(R+1, C, 'Февраль', cbtOuter);
  FWriter.WriteText(R+2, C, 'Март', cbtOuter);
  R:= R+3;
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  FWriter.WriteText(R, C,   'I КВАРТАЛ', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, QUARTER_COLOR_INDEX);
  R:= R+1;
  FWriter.SetFont(FFontName, FFontSize, [], scBlack);
  FWriter.WriteText(R, C,   'Апрель', cbtOuter);
  FWriter.WriteText(R+1, C, 'Май', cbtOuter);
  FWriter.WriteText(R+2, C, 'Июнь', cbtOuter);
  R:= R+3;
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  FWriter.WriteText(R, C,   'II КВАРТАЛ', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, QUARTER_COLOR_INDEX);
  R:= R+1;
  FWriter.WriteText(R, C,   'I ПОЛУГОДИЕ', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, HALFYEAR_COLOR_INDEX);
  R:= R+1;
  FWriter.SetFont(FFontName, FFontSize, [], scBlack);
  FWriter.WriteText(R, C,   'Июль', cbtOuter);
  FWriter.WriteText(R+1, C, 'Август', cbtOuter);
  FWriter.WriteText(R+2, C, 'Сентябрь', cbtOuter);
  R:= R+3;
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  FWriter.WriteText(R, C,   'III КВАРТАЛ', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, QUARTER_COLOR_INDEX);
  R:= R+1;
  FWriter.SetFont(FFontName, FFontSize, [], scBlack);
  FWriter.WriteText(R, C,   'Октябрь', cbtOuter);
  FWriter.WriteText(R+1, C, 'Ноябрь', cbtOuter);
  FWriter.WriteText(R+2, C, 'Декабрь', cbtOuter);
  R:= R+3;
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  FWriter.WriteText(R, C,   'IV КВАРТАЛ', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, QUARTER_COLOR_INDEX);
  R:= R+1;
  FWriter.WriteText(R, C,   'II ПОЛУГОДИЕ', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, HALFYEAR_COLOR_INDEX);
  R:= R+1;
  FWriter.WriteText(R, C,   IntToStr(FYear) + ' ГОД', cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, YEAR_COLOR_INDEX);

  R:= RESUME_FIRST_ROW;
  C:= RESUME_FIRST_COL;
  FWriter.DrawBorders(R, C, R+3, C+7, cbtAll);
  R:= RESUME_FIRST_ROW+4;
  C:= RESUME_FIRST_COL;
  FWriter.DrawBorders(R, C, R+18, C, cbtAll);
end;

procedure TCalendarSheet.DrawResumeLine(const R,C, AColorIndex: Integer; const ACalendar: TCalendar);
var
  i: Integer;
begin
  FWriter.WriteNumber(R,C,   ACalendar.DaysCount, cbtOuter);
  FWriter.WriteNumber(R,C+1, ACalendar.WorkDaysCount, cbtOuter);
  FWriter.WriteNumber(R,C+2, ACalendar.OffDaysCount, cbtOuter);
  FWriter.WriteNumber(R,C+3, ACalendar.HoliDaysCount, cbtOuter);

  FWriter.WriteNumber(R,C+4, ACalendar.SumWorkHoursFrac(40), FRACTION_DIGITS_IN_WORKHOURS, cbtOuter);
  FWriter.WriteNumber(R,C+5, ACalendar.SumWorkHoursFrac(36), FRACTION_DIGITS_IN_WORKHOURS, cbtOuter);
  FWriter.WriteNumber(R,C+6, ACalendar.SumWorkHoursFrac(24), FRACTION_DIGITS_IN_WORKHOURS, cbtOuter);

  if AColorIndex> 0 then
    for i:=0 to 6 do
      FWriter.AddCellBGColorIndex(R, C+i, AColorIndex);

  FWriter.DrawBorders(R, C, R, C+6, cbtAll);
end;

procedure TCalendarSheet.DrawYear;
var
  R,C: Integer;
begin
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  R:= YEAR_RESUME_ROW;
  C:= RESUME_FIRST_COL+1;
  DrawResumeLine(R,C, YEAR_COLOR_INDEX, FCalendar);
end;

procedure TCalendarSheet.DrawHalf(const AHalf: Byte);
var
  R,C: Integer;
  HalfCalendar: TCalendar;
  BD,ED: TDate;
begin
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  R:= HALF_RESUME_ROWS[AHalf];
  C:= RESUME_FIRST_COL+1;
  FirstLastDayInHalfYear(AHalf, FYear, BD, ED);
  HalfCalendar:= TCalendar.Create;
  try
    FCalendar.Cut(BD, ED, HalfCalendar);
    DrawResumeLine(R,C, HALFYEAR_COLOR_INDEX, HalfCalendar);
  finally
    FreeAndNil(HalfCalendar);
  end;
end;

procedure TCalendarSheet.DrawQuarter(const AQuarter: Byte);
var
  R,C: Integer;
  QuarterCalendar: TCalendar;
  BD,ED: TDate;
begin
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  R:= QUARTER_RESUME_ROWS[AQuarter];
  C:= RESUME_FIRST_COL+1;
  FirstLastDayInQuarter(AQuarter, FYear, BD, ED);
  QuarterCalendar:= TCalendar.Create;
  try
    FCalendar.Cut(BD, ED, QuarterCalendar);
    DrawResumeLine(R,C, QUARTER_COLOR_INDEX, QuarterCalendar);
  finally
    FreeAndNil(QuarterCalendar);
  end;
end;

procedure TCalendarSheet.DrawMonth(const AMonth: Byte);
var
  R,C, i,j: Integer;
  MonthCalendar: TCalendar;
begin
  FWriter.SetFont(FFontName, FFontSize, [fssBold], scBlack);
  R:= MONTH_FIRST_ROWS[AMonth];
  C:= MONTH_FIRST_COLS[AMonth];
  FWriter.WriteText(R,C,R,C+6, SUpper(MONTHSNOM[AMonth]), cbtOuter);
  FWriter.AddCellBGColorIndex(R, C, MONTHNAME_COLOR_INDEX);

  R:= R+1;
  for i:= 0 to 6 do
  begin
    C:= MONTH_FIRST_COLS[AMonth] + i;
    FWriter.WriteText(R, C, WEEKDAYSSHORT[i+1], cbtOuter);
    FWriter.AddCellBGColorIndex(R, C, DAYNAME_COLOR_INDEX);
  end;
  FWriter.SetFont(FFontName, FFontSize, [], scBlack);
  for i:= MONTH_FIRST_ROWS[AMonth] + 2 to MONTH_FIRST_ROWS[AMonth] + 7 do
    for j:= MONTH_FIRST_COLS[AMonth] to MONTH_FIRST_COLS[AMonth] + 6 do
      FWriter.WriteText(i, j, EmptyStr, cbtOuter);
  MonthCalendar:= TCalendar.Create;
  try
    FCalendar.Cut(FirstDayInMonth(AMonth, FYear), LastDayInMonth(AMonth, FYear), MonthCalendar);
    for i:=0 to MonthCalendar.DaysCount - 1 do
    begin
      R:= MONTH_FIRST_ROWS[AMonth] + MonthCalendar.WeekNumsInMonth[i] + 1;
      C:= MONTH_FIRST_COLS[AMonth] + MonthCalendar.DayNumsInWeek[i] - 1;
      FWriter.WriteNumber(R,C,i+1,cbtOuter);
      j:= VIndexOfDate(FHighLightDays, MonthCalendar.Dates[i]);
      if j>=0 then
        FWriter.AddCellBGColorIndex(R, C, HIGHLIGHT_COLOR_INDEX)
      else
        FWriter.AddCellBGColorIndex(R, C, MonthCalendar.DayStatuses[i]);
    end;
    R:= MONTH_RESUME_ROWS[AMonth];
    C:= RESUME_FIRST_COL+1;
    DrawResumeLine(R,C, 0, MonthCalendar);
  finally
    FreeAndNil(MonthCalendar);
  end;

  //устранение проблем с отображением границ
  R:= MONTH_FIRST_ROWS[AMonth];
  C:= MONTH_FIRST_COLS[AMonth];
  FWriter.DrawBorders(R, C, R+7, C+6, cbtAll);
end;

constructor TCalendarSheet.Create(const AWorksheet: TsWorksheet;
  const AGrid: TsWorksheetGrid);
var
  i, W: Integer;
  Widths: TIntVector;
begin
  FGridLineColor:= clBlack;
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;
  FRowHeight:= 25;

  Widths:= nil;
  VDim(Widths, 32);

  W:= 30;
  for i:= 0 to 23 do Widths[i]:= W;

  W:= 110;
  Widths[24]:= W;

  W:= 60;
  for i:= 25 to 28 do Widths[i]:= W;

  W:= 80;
  for i:= 29 to 31 do Widths[i]:= W;

  FWriter:= TSheetWriter.Create(Widths, AWorksheet, AGrid);
  FWriter.SetBordersColor(FGridLineColor);
end;

destructor TCalendarSheet.Destroy;
begin
  FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TCalendarSheet.Zoom(const APercents: Integer);
begin
  FWriter.SetZoom(APercents);
end;

procedure TCalendarSheet.Draw(const AYearCalendar: TCalendar;
  const AHighLightDays: TDateVector);
var
  i: Integer;
begin
  FCalendar:= AYearCalendar;
  FYear:= YearOfDate(FCalendar.BeginDate);
  FHighLightDays:= AHighLightDays;
  FWriter.BeginEdit;
  DrawCaption;
  DrawLegend;
  DrawResumeTableCaption;
  for i:=1 to 12 do DrawMonth(i);
  for i:=1 to 4  do DrawQuarter(i);
  for i:=1 to 2  do DrawHalf(i);
  DrawYear;
  FWriter.WriteText(13,25, EmptyStr);

  FWriter.SetRowHeight(2, FRowHeight-4);
  for i:= 3 to FWriter.RowCount-1 do
    FWriter.SetRowHeight(i, FRowHeight);

  FWriter.EndEdit;
end;

procedure TCalendarSheet.UpdateColors(const AColorVector: TColorVector);
begin
  FColorVector:= AColorVector;
  FWriter.ApplyBGColors(FColorVector);
end;

procedure TCalendarSheet.ClearColors;
begin
  FWriter.ClearBGColors;
end;

function TCalendarSheet.GridToMonth(const ARow, ACol: Integer;
                    out ADayInWeek, AWeekInMonth, AMonth: Integer): Boolean;

  function GetMonth(const AQuart: Byte): Boolean;
  var
    m: Integer;
  begin
    GetMonth:= True;
    m:= (AQuart-1)*3 + 1;
    if (ACol>=1) and (ACol<=7) then
    begin
      ADayInWeek:= ACol;
      AMonth:= m;
    end
    else if (ACol>=9) and (ACol<=15) then
    begin
      ADayInWeek:= ACol-8;
      AMonth:= m+1;
    end
    else if (ACol>=17) and (ACol<=23) then
    begin
      ADayInWeek:= ACol-16;
      AMonth:= m+2
    end
    else begin
      GetMonth:= False;
      Exit;
    end;
    case AQuart of
    1: AWeekInMonth:= ARow-4;
    2: AWeekInMonth:= ARow-13;
    3: AWeekInMonth:= ARow-22;
    4: AWeekInMonth:= ARow-31;
    end;
  end;

begin
  Result:= False;
  ADayInWeek:= 0;
  AWeekInMonth:= 0;
  AMonth:= 0;
  if (ARow>=5) and (ARow<=10) then Result:= GetMonth(1)
  else if (ARow>=14) and (ARow<=19) then Result:= GetMonth(2)
    else if (ARow>=23) and (ARow<=28) then Result:= GetMonth(3)
      else if (ARow>=32) and (ARow<=37) then Result:= GetMonth(4);
end;

function TCalendarSheet.GridToDate(const ARow, ACol: Integer; out ADate: TDate): Boolean;
var
  D,W,M: Integer;
begin
  Result:= False;
  ADate:= NULDATE;
  if GridToMonth(ARow, ACol, D,W,M) then
    Result:= MonthGridToDate(W,D,M,FYear, ADate);
end;

function TCalendarSheet.DateToGrid(const ADate: TDate; out ARow, ACol: Integer): Boolean;
var
  X: Integer;
begin
  ARow:= -1;
  ACol:= -1;
  Result:= False;
  if YearOfDate(ADate)<>FYear then Exit;

  X:= QuarterNumber(ADate);
  case X of
  1: ARow:= WeekNumberInMonth(ADate) + 4;
  2: ARow:= WeekNumberInMonth(ADate) + 13;
  3: ARow:= WeekNumberInMonth(ADate) + 22;
  4: ARow:= WeekNumberInMonth(ADate) + 31;
  end;

  X:= MonthNumberInQuarter(ADate);
  case X of
  1: ACol:= DayNumberInWeek(ADate);
  2: ACol:= DayNumberInWeek(ADate) + 8;
  3: ACol:= DayNumberInWeek(ADate) + 16;
  end;
  Result:= True;
end;


{Utils ------------------------------------------------------------------------}

function WeekHoursToWorkHoursInt(const AHoursInWeek: Byte): Integer;
begin
  Result:= Trunc(AHoursInWeek*WORKHOURS_DENOMINATOR/5);
end;

function MonthGridToDate(const AWeekInMonth, ADayInWeek, AMonth, AYear: Integer; out ADate: TDate): Boolean;
var
  i: Integer;
  D: TDate;
begin
  Result:= False;
  ADate:= NULDATE;
  for i:= 1 to DaysInAMonth(AYear, AMonth) do
  begin
    D:= EncodeDate(AYear, AMonth, i);
    if (DayNumberInWeek(D)=ADayInWeek) and (WeekNumberInMonth(D)=AWeekInMonth) then
    begin
      ADate:= D;
      Result:= True;
      break;
    end;
  end;
end;

end.

