unit USheetUtils;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DK_SheetWriter, fpstypes, fpspreadsheetgrid, fpspreadsheet,
  DK_Vector, DK_Matrix, DK_Fonts, Grids, Graphics, DK_SheetConst, DK_Const,
  DateUtils, DK_DateUtils;

const
  COLOR_BACKGROUND_TITLE = clBtnFace;
  COLOR_BACKGROUND_SELECTED   = $00FBDEBB;

  SHEET_FONT_NAME = 'Arial';//'Times New Roman';
  SHEET_FONT_SIZE = 9;

  CHECK_SYMBOL = '✓';

  PERCENT_FRAC_DIGITS = 1;

type

  { TSeriesSheet }

  TSeriesSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;

    Series: TStrVector;

    procedure DrawTitle;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure Draw(const ASeries: TStrVector);
  end;

  { TMotorBuildSheet }

  TMotorBuildSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;

    //BuildDates: TDateVector;
    BuildDate: TDate;
    MotorNames, MotorNums, RotorNums: TStrVector;

    procedure DrawTitle;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure Draw({const ABuildDates: TDateVector;}
                   const ABuildDate: TDate;
                   const AMotorNames, AMotorNums, ARotorNums: TStrVector);
    procedure DrawReport(const ABeginDate, AEndDate: TDate;
                   const ANeedNumberList: Boolean;
                   const ABuildDates: TDateVector;
                   const AMotorNames, AMotorNums, ARotorNums: TStrVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCount: TIntVector);

  end;

  { TReportShipmentSheet }

  TReportShipmentSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure DrawReport(const ASingleReceiver: Boolean;
                   const ANeedNumberList: Boolean;
                   const ABeginDate, AEndDate: TDate;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCount: TIntVector;
                   const ARecieverNames: TStrVector;
                   const ARecieverMotorNames: TStrMatrix;
                   const ARecieverMotorCounts: TIntMatrix;
                   const AListSendDates: TDateVector;
                   const AListMotorNames, AListMotorNums, AListReceiverNames: TStrVector);

  end;

  { TStatisticReclamationSheet }

  TStatisticReclamationSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;

    FGroupType: Integer;
    FSecondColumnType: Integer;
    FAdditionYearsCount: Integer;
    FShowAllYearsTotalCount: Boolean;
    FShowSecondColumnForTotalCount: Boolean;
    FShowSecondColumn: Boolean;
    FSeveralYears: Boolean;
    FUsedReasons: TBoolVector;
    FTotal: TIntVector;

    procedure DrawTop(var ARow: Integer; const ABeginDate, AEndDate: TDate;

                        const AReportName, AMotorNames: String);
    procedure DrawTableCaptionYear(var ACol: Integer; const ARow, AYear: Integer;
                            const AReasonNames: TStrVector);
    procedure DrawTableCaptionReason(var ACol: Integer; const ARow, AFirstYear, AIndex: Integer;
                            const AReasonName: String);
    procedure DrawTableCaption(var ARow: Integer;
                        const ABeginDate: TDate;
                        const AParamCaption: String;
                        const AReasonNames: TStrVector);

    procedure DrawCountColumn(const ARow, ACol: Integer;
                              const AValues: TIntVector;
                              const AResumeNeed: Boolean);
    procedure DrawAccumColumn(const ARow, ACol: Integer;
                              const AValues: TIntVector;
                              const AResumeNeed: Boolean);
    procedure DrawPercentColumn(const ARow, ACol: Integer;
                              const ATotal, AValues: TIntVector;
                              const AResumeNeed: Boolean);


    procedure DrawDataYear(const ARow, ACol: Integer; const ACounts: TIntMatrix3D;
                             const AResumeNeed: Boolean);
    procedure DrawDataReason(const ARow, ACol: Integer; const ACounts: TIntMatrix3D;
                             const AResumeNeed: Boolean);
    procedure DrawData(const ARow: Integer; const AParamNames: TStrVector;
                       const ACounts: TIntMatrix3D;  const AResumeNeed: Boolean);
    procedure CalcTotalForAllPeriods(const ACounts: TIntMatrix3D);
  public
    //AGroupType = 1 - группировка по годам, 2 - группировка по причинам неисправностей
    //ASecondColumnType = 1 - % от общего количества, 2 - накопление
    constructor Create(const AGrid: TsWorksheetGrid;
                       const AParamNameColumnWidth, AGroupType, ASecondColumnType, AAdditionYearsCount: Integer;
                       const AUsedReasons: TBoolVector;
                       const AShowSecondColumn: Boolean = True;
                       const AShowAllYearsTotalCount: Boolean = False);
    destructor  Destroy; override;
    procedure Zoom(const APercents: Integer);
    procedure Draw(const ABeginDate, AEndDate: TDate;
                   const AReportName, AMotorNames, AParamCaption: String;
                   const AParamNames, AReasonNames: TStrVector;
                   const ACounts: TIntMatrix3D;
                   const AResumeNeed: Boolean);

  end;

  { TStatisticReclamationCountSheet }

  //TStatisticReclamationCountSheet = class (TObject)
  //private
  //  FGrid: TsWorksheetGrid;
  //  FWriter: TSheetWriter;
  //  FFontName: String;
  //  FFontSize: Single;
  //  procedure DrawTitle(var ARow: Integer; const ABeginDate, AEndDate: TDate;
  //                      const AName, AMotorNames: String);
  //  procedure DrawData(const ARow: Integer;
  //              const AFirstColName: String;
  //              const ANameValues, ATitleReasonNames: TStrVector;
  //              const AUsedReasons: TBoolVector;
  //              const AReasonCountValues: TIntMatrix;
  //              const AResumeNeed: Boolean);
  //public
  //  constructor Create(const AGrid: TsWorksheetGrid; const AReasonsCount: Integer);
  //  destructor  Destroy; override;
  //  procedure Draw(const ABeginDate, AEndDate: TDate;
  //                 const AName, AMotorNames, AFirstColName: String;
  //                 const ANameValues, ATitleReasonNames: TStrVector;
  //                 const AUsedReasons: TBoolVector;
  //                 const AReasonCountValues: TIntMatrix;
  //                 const AResumeNeed: Boolean);
  //end;

  { TStatisticReclamationMonthSheet }

  //TStatisticReclamationMonthSheet = class (TObject)
  //private
  //  FGrid: TsWorksheetGrid;
  //  FWriter: TSheetWriter;
  //  FFontName: String;
  //  FFontSize: Single;
  //  procedure DrawTitle(var ARow: Integer; const AYears: TStrVector;
  //                      const AName, AMotorNames: String);
  //  procedure DrawData(const ARow: Integer;
  //              const AYears: TStrVector;
  //              const ACounts, AAccumulations: TIntMatrix;
  //              const AResumeNeed: Boolean);
  //public
  //  constructor Create(const AGrid: TsWorksheetGrid; const AYearsCount: Integer);
  //  destructor  Destroy; override;
  //  procedure Draw(const AYears: TStrVector;
  //                 const AName, AMotorNames: String;
  //                 const ACounts, AAccumulations: TIntMatrix;
  //                 const AResumeNeed: Boolean);
  //end;

  { TMotorTestSheet }

  TMotorTestSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;

    //TestDates: TDateVector;
    TestDate: TDate;
    MotorNames, MotorNums, TestNotes: TStrVector;
    TestResults: TIntVector;

    procedure DrawTitle;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure Draw(//const ATestDates: TDateVector;
                   const ATestDate: TDate;
                   const AMotorNames, AMotorNums, ATestNotes{, ASeries}: TStrVector;
                   const ATestResults, ATotalMotorCounts, ATotalFailCounts: TIntVector);
    procedure DrawReport(const ABeginDate, AEndDate: TDate;
                   const ANeedNumberList: Boolean;
                   const ATestDates: TDateVector;
                   const AMotorNames, AMotorNums, ATestNotes{, ASeries}: TStrVector;
                   const ATestResults: TIntVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCounts, ATotalFailCounts: TIntVector);

  end;

  { TStoreSheet }

  TStoreSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure Draw(const ADeltaDays: Integer;
                   const ATestDates: TDateVector;
                   const AMotorNames, AMotorNums{, ASeries}: TStrVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCounts: TIntVector);
  end;

  { TBeforeTestSheet }

  TBeforeTestSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure Draw(const ABuildDates: TDateVector;
                   const AMotorNames, AMotorNums: TStrVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCounts: TIntVector;
                   const ATestDates: TDateMatrix;
                   const ATestFails: TIntMatrix;
                   const ATestNotes: TStrMatrix);
  end;



  { TCargoSheet }

  TCargoSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;

  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure Draw(const ASendDate: TDate; const AReceiverName: String;
                   const AMotorNames: TStrVector; const AMotorsCount: TIntVector;
                   const AMotorNums, ASeries: TStrMatrix);
  end;

  { TRepairSheet }

  TRepairSheet = class (TObject)
  private
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;

  public
    constructor Create(const ASheet: TsWorksheet);
    destructor  Destroy; override;
    procedure Draw(const APassports, ADayCounts: TIntVector;
                   const AMotorNames, AMotorNums: TStrVector;
                   const AArrivalDates, ASendingDates: TDateVector);
  end;

  { TMotorCardSheet }

  TMotorCardSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure Zoom(const APercents: Integer);
    procedure Draw(const ABuildDate, ASendDate: TDate;
      const AMotorName, AMotorNum, ASeries, ARotorNum, AReceiverName, AControlNote: String;
      const ATestDates: TDateVector;
      const ATestResults: TIntVector;
      const ATestNotes: TStrVector;
      const ARecDates: TDateVector;
      const AMileages, AOpinions: TIntVector;
      const APlaceNames, AFactoryNames, ADepartures,
      ADefectNames, AReasonNames, ARecNotes: TStrVector);
  end;

  { TReclamationSheet }

  TReclamationSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;

    RecDates, BuildDates, ArrivalDates, SendingDates: TDateVector;
    Mileages, Opinions, ReasonColors{, Passports}: TIntVector;
    PlaceNames, FactoryNames, Departures: TStrVector;
    DefectNames, ReasonNames, RecNotes: TStrVector;
    MotorNames, MotorNums: TStrVector;

    procedure DrawTitle;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure Zoom(const APercents: Integer);
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure Draw(const ARecDates, ABuildDates, AArrivalDates, ASendingDates: TDateVector;
                   const AMileages, AOpinions, AReasonColors{, APassports}: TIntVector;
                   const APlaceNames, AFactoryNames, ADepartures,
                         ADefectNames, AReasonNames, ARecNotes,
                         AMotorNames, AMotorNums: TStrVector);
  end;

implementation

{ TRepairSheet }

constructor TRepairSheet.Create(const ASheet: TsWorksheet);
var
  ColWidths: TIntVector;
begin
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  ColWidths:= VCreateInt([
    60,   //№п/п
    300,  // Наименование
    150,  // Номер
    100,  // Наличие паспорта
    150,  // Прибыл в ремонт
    150,  // Убыл из ремонта
    130   // Срок ремонта (дней)
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, ASheet, nil);
end;

destructor TRepairSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TRepairSheet.Draw(const APassports, ADayCounts: TIntVector;
               const AMotorNames, AMotorNums: TStrVector;
               const AArrivalDates, ASendingDates: TDateVector);
var
  R, i: Integer;
  S: String;
begin
  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;

  R:= 1;
  FWriter.SetAlignment(haCenter, vaCenter);
  S:= 'Гарантийный ремонт электродвигателей';
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);

  R:=R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.WriteText(R, 1, '№п/п', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Наименование', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Номер', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Наличие паспорта', cbtOuter, True, True);
  FWriter.WriteText(R, 5, 'Прибыл в ремонт', cbtOuter, True, True);
  FWriter.WriteText(R, 6, 'Убыл из ремонта', cbtOuter, True, True);
  FWriter.WriteText(R, 7, 'Срок ремонта (рабочих дней)', cbtOuter, True, True);

  FWriter.SetFont(FFontName, FFontSize, [], clBlack);
  for i:= 0 to High(AMotorNums) do
  begin
    R:=R + 1;
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteNumber(R, 1, i+1, cbtOuter);
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
    S:= EmptyStr;
    if APassports[i]>0 then S:= CHECK_SYMBOL;
    FWriter.WriteText(R, 4, S, cbtOuter);
    FWriter.WriteDate(R, 5, AArrivalDates[i], cbtOuter);
    if ASendingDates[i]>0 then
      FWriter.WriteDate(R, 6, ASendingDates[i], cbtOuter)
    else
      FWriter.WriteText(R, 6, EmptyStr, cbtOuter);
    if ADayCounts[i]>0 then
      FWriter.WriteNumber(R, 7, ADayCounts[i], cbtOuter)
    else
      FWriter.WriteText(R, 7, EmptyStr, cbtOuter);
  end;

  FWriter.EndEdit;
end;

{ TStatisticReclamationSheet }

procedure TStatisticReclamationSheet.DrawTop(var ARow: Integer;
  const ABeginDate, AEndDate: TDate; const AReportName, AMotorNames: String);
var
  R, i: Integer;
  S: String;
begin

  FWriter.SetAlignment(haCenter, vaCenter);

  R:= ARow;
  S:= 'Отчет по рекламационным случаям электродвигателей';
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);

  if AMotorNames<>EmptyStr then
  begin
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.WriteText(R, 1, R, FWriter.ColCount, AMotorNames, cbtNone, True, True);
  end;

  R:= R + 1;
  S:= 'за ';
  if FSeveralYears then
  begin
    if SameDate(ABeginDate, AEndDate) then
      S:= S + FormatDateTime('dd ', ABeginDate) + MONTHSGEN[MonthOfDate(ABeginDate)]
    else
      S:= S + 'период с ' +
          FormatDateTime('dd ', ABeginDate) + MONTHSGEN[MonthOfDate(ABeginDate)] +
         ' по ' +
          FormatDateTime('dd ', AEndDate) + MONTHSGEN[MonthOfDate(AEndDate)];
    S:= S + ' ' + IntToStr(YearOfDate(ABeginDate)-FAdditionYearsCount);
    for i:= FAdditionYearsCount-1 downto 0 do
      S:= S + '/' + IntToStr(YearOfDate(ABeginDate)-i);
    S:= S + 'гг.'
  end
  else begin
    if SameDate(ABeginDate, AEndDate) then
      S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
    else
      S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
         ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);
  end;
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);

  if AReportName<>EmptyStr then
  begin
    R:= R + 1;
    S:= '(' + AReportName + ')';
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);
  end;

  //EmptyRow
  R:= R + 1;
  FWriter.WriteText(R, 1, R, FWriter.ColCount, EmptyStr, cbtNone);
  FWriter.SetRowHeight(R, 10);

  ARow:= R + 1;

end;

procedure TStatisticReclamationSheet.DrawTableCaptionYear(var ACol: Integer;
  const ARow, AYear: Integer; const AReasonNames: TStrVector);
var
  R1, R2, C1, C2, i, ColCount: Integer;
begin
  R1:= ARow;
  R2:= R1 + Ord(FSeveralYears);

  ColCount:= Ord(FUsedReasons[0]) * (1 + Ord(FShowSecondColumnForTotalCount)) +
             (1+Ord(FShowSecondColumn)) * VCountIf(FUsedReasons, True, 1);

  C1:= ACol + 1;
  C2:= ACol + ColCount;
  ACol:= C2;

  //Номер Года
  if FSeveralYears then
    FWriter.WriteNumber(R1, C1, R1, C2, AYear, cbtOuter);

  C2:= C1-1;
  //Общее количество за этот период
  if FUsedReasons[0] then
  begin
    C1:= C2 + 1;
    C2:= C1 + Ord(FShowSecondColumnForTotalCount);
    FWriter.WriteText(R2, C1, R2, C2, AReasonNames[0], cbtOuter, True, True);
  end;

  for i:= 1 to High(FUsedReasons) do
  begin
    if FUsedReasons[i] then
    begin
      C1:= C2 + 1;
      C2:= C1 + Ord(FShowSecondColumn);
      FWriter.WriteText(R2, C1, R2, C2, AReasonNames[i], cbtOuter, True, True);
    end;
  end;
end;

procedure TStatisticReclamationSheet.DrawTableCaptionReason(var ACol: Integer;
  const ARow, AFirstYear, AIndex: Integer; const AReasonName: String);
var
  R1, R2, C1, C2, i, ColCount: Integer;
begin
  R1:= ARow;
  R2:= R1 + Ord(FSeveralYears);

  if AIndex=0 then  //общее количество
    ColCount:= (FAdditionYearsCount + 1) * (1 + Ord(FShowSecondColumnForTotalCount))
  else //причина
    ColCount:= (FAdditionYearsCount + 1) * (1 + Ord(FShowSecondColumn));

  C1:= ACol + 1;
  C2:= ACol + ColCount;
  ACol:= C2;

  //Наименование причины неисправности
  FWriter.WriteText(R1, C1, R1, C2, AReasonName, cbtOuter, True, True);

  if not FSeveralYears then Exit;
  //Общее количество за этот период
  C2:= C1-1;
  for i:= 0 to FAdditionYearsCount do
  begin
    C1:= C2 + 1;
    if AIndex=0 then
      C2:= C1 + Ord(FShowSecondColumnForTotalCount)
    else
      C2:= C1 + Ord(FShowSecondColumn);
    FWriter.WriteNumber(R2, C1, R2, C2, AFirstYear+i, cbtOuter);
  end;
end;

procedure TStatisticReclamationSheet.DrawTableCaption(var ARow: Integer;
  const ABeginDate: TDate; const AParamCaption: String;
  const AReasonNames: TStrVector);
var
  R1, R2, C, i, FirstYear: Integer;
begin
  R1:= ARow;
  R2:= R1 + Ord(FSeveralYears);

  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);

  C:= 1;
  FWriter.WriteText(R1, C, R2, C, AParamCaption, cbtOuter, True, True);

  //общее количество за все периоды
  if FShowAllYearsTotalCount and FSeveralYears then
  begin
    C:= C + 1;
    FWriter.WriteText(R1, C, R2, C, 'Общее количество за все периоды', cbtOuter, True, True);
  end;

  FirstYear:= YearOfDate(ABeginDate) - FAdditionYearsCount;
  if FGroupType = 1 then  //группировка по годам
  begin
    //пробегаем по всем периодам
    for i:= 0 to FAdditionYearsCount do
      DrawTableCaptionYear(C, R1, FirstYear+i, AReasonNames);
  end
  else if FGroupType = 2 then //группировка по причинам неисправностей
  begin
    //пробегем по всем причинам
    for i:= 0 to High(AReasonNames) do
      if FUsedReasons[i] then
        DrawTableCaptionReason(C, R1, FirstYear, i, AReasonNames[i]);
  end;



  FWriter.DrawBorders(R1, 1, R2, FWriter.ColCount, cbtAll);

  ARow:= R2;
end;

procedure TStatisticReclamationSheet.DrawCountColumn(const ARow, ACol: Integer;
  const AValues: TIntVector; const AResumeNeed: Boolean);
var
  i, R: Integer;
begin
  R:= ARow;
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  for i:= 0 to High(AValues) do
  begin
    R:= R + 1;
    FWriter.WriteNumber(R, ACol, AValues[i], cbtOuter);
  end;
  if AResumeNeed then
  begin
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.WriteNumber(R, ACol, VSum(AValues), cbtOuter);
  end;
end;

procedure TStatisticReclamationSheet.DrawAccumColumn(const ARow, ACol: Integer;
  const AValues: TIntVector; const AResumeNeed: Boolean);
var
  Values: TIntVector;
  i, R: Integer;
begin
  Values:= nil;
  VDim(Values, Length(AValues));
  Values[0]:= AValues[0];
  for i:= 1 to High(Values) do
    Values[i]:= AValues[i] + Values[i-1];

  R:= ARow;
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  for i:= 0 to High(Values) do
  begin
    R:= R + 1;
    FWriter.WriteNumber(R, ACol, Values[i], cbtOuter);
  end;
  if AResumeNeed then
  begin
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.WriteNumber(R, ACol-1, R, ACol, Values[High(Values)], cbtOuter);
  end;
end;

procedure TStatisticReclamationSheet.DrawPercentColumn(const ARow,
  ACol: Integer; const ATotal, AValues: TIntVector; const AResumeNeed: Boolean);
var
  i, R, X: Integer;
begin
  R:= ARow;
  X:= VSum(ATotal);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  for i:= 0 to High(AValues) do
  begin
    R:= R + 1;
    if X=0 then
      FWriter.WriteNumber(R, ACol, 0, 1, cbtOuter, nfPercentage)
    else
      FWriter.WriteNumber(R, ACol, AValues[i]/X, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
  end;
  if AResumeNeed then
  begin
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    if X=0 then
      FWriter.WriteNumber(R, ACol, 0, 1, cbtOuter, nfPercentage)
    else
      FWriter.WriteNumber(R, ACol, VSum(AValues)/X, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
  end;
end;

procedure TStatisticReclamationSheet.DrawDataYear(const ARow, ACol: Integer;
  const ACounts: TIntMatrix3D; const AResumeNeed: Boolean);
var
  R, C, i, j: Integer;
begin
  R:= ARow;
  C:= ACol -1;
  for i:= 0 to FAdditionYearsCount do
  begin
    if FUsedReasons[0] then
    begin
      C:= C + 1;
      DrawCountColumn(R, C, ACounts[i, 0], AResumeNeed);
      if FShowSecondColumn then
      begin
        if FSecondColumnType=2{накопление} then
        begin
          C:= C + 1;
          DrawAccumColumn(R, C, ACounts[i, 0], AResumeNeed);
        end
        else if FSeveralYears and (FSecondColumnType=1{%}) then
        begin
          C:= C + 1;
          DrawPercentColumn(R, C, FTotal, ACounts[i, 0], AResumeNeed)
        end;
      end;
    end;
    for j:= 1 to High(FUsedReasons) do
    begin
      if FUsedReasons[j] then
      begin
        C:= C + 1;
        DrawCountColumn(R, C, ACounts[i, j], AResumeNeed);
        if FShowSecondColumn then
        begin
          C:= C + 1;
          if FSecondColumnType = 1 then
            DrawPercentColumn(R, C, ACounts[i, 0], ACounts[i, j], AResumeNeed)
          else  //2
            DrawAccumColumn(R, C, ACounts[i, j], AResumeNeed);
        end;
      end;
    end;
  end;
end;

procedure TStatisticReclamationSheet.DrawDataReason(const ARow, ACol: Integer;
  const ACounts: TIntMatrix3D; const AResumeNeed: Boolean);
var
  i,j, R, C: Integer;
begin
  FWriter.SetAlignment(haCenter, vaCenter);

  R:= ARow;
  C:= ACol - 1;

  if FUsedReasons[0] then
  begin
    for i:= 0 to FAdditionYearsCount do
    begin
      C:= C + 1;
      DrawCountColumn(R, C, ACounts[i, 0], AResumeNeed);
      if FShowSecondColumn then
      begin
        if FSecondColumnType=2{накопление} then
        begin
          C:= C + 1;
          DrawAccumColumn(R, C, ACounts[i, 0], AResumeNeed);
        end
        else if FSeveralYears and (FSecondColumnType=1{%}) then
        begin
          C:= C + 1;
          DrawPercentColumn(R, C, FTotal, ACounts[i, 0], AResumeNeed)
        end;
      end;
    end;
  end;

  for j:= 1 to High(FUsedReasons) do
  begin
    if not FUsedReasons[j] then
      continue;
    for i:= 0 to FAdditionYearsCount do
    begin
      C:= C + 1;
      DrawCountColumn(R, C, ACounts[i, j], AResumeNeed);
      if FShowSecondColumn then
      begin
        C:= C + 1;
        if FSecondColumnType = 1 then
          DrawPercentColumn(R, C, ACounts[i, 0], ACounts[i, j], AResumeNeed)
        else  //2
          DrawAccumColumn(R, C, ACounts[i, j], AResumeNeed);
      end;
    end;
  end;
end;

procedure TStatisticReclamationSheet.CalcTotalForAllPeriods(const ACounts: TIntMatrix3D);
var
  i: Integer;
begin
  FTotal:= nil;
  if MIsNil(ACounts) then Exit;

  VDim(FTotal, Length(ACounts[0,0]));
  for i:= 0 to High(ACounts) do
    FTotal:= VSum(FTotal, ACounts[i, 0]);
end;

procedure TStatisticReclamationSheet.DrawData(const ARow: Integer;
  const AParamNames: TStrVector; const ACounts: TIntMatrix3D;
  const AResumeNeed: Boolean);
var
  i, R, C: Integer;
  //Total: TIntVector;
begin
  //Total:= nil;

  R:= ARow;
  C:= 1;
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  for i:= 0 to High(AParamNames) do
  begin
    R:= R + 1;
    FWriter.WriteText(R, C, AParamNames[i], cbtOuter, True, True);
  end;
  if AResumeNeed then
  begin
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.WriteText(R, C, 'ИТОГО', cbtOuter, True, True);
  end;

  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  if FShowAllYearsTotalCount and FSeveralYears then //общее количество за все периоды
  begin
    //VDim(Total, Length(AParamNames));
    //for i:= 0 to High(ACounts) do
    //  Total:= VSum(Total, ACounts[i, 0]);
    R:= ARow;
    C:= C + 1;
    for i:= 0 to High(AParamNames) do
    begin
      R:= R + 1;
      FWriter.WriteNumber(R, C, FTotal[i], cbtOuter);
    end;
    if AResumeNeed then
    begin
      R:= R + 1;
      FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
      FWriter.WriteNumber(R, C, VSum(FTotal), cbtOuter);
    end;
  end;

  R:= ARow;
  C:= C + 1;
  if FGroupType=1 then
    DrawDataYear(R, C, ACounts, AResumeNeed)
  else if FGroupType=2 then
    DrawDataReason(R, C, ACounts, AResumeNeed)
end;

constructor TStatisticReclamationSheet.Create(const AGrid: TsWorksheetGrid;
  const AParamNameColumnWidth, AGroupType, ASecondColumnType, AAdditionYearsCount: Integer;
  const AUsedReasons: TBoolVector; const AShowSecondColumn: Boolean;
  const AShowAllYearsTotalCount: Boolean);
var
  ColWidths: TIntVector;

  procedure SetWidthsYear;
  var
    i, j, W1, W2: Integer;
  begin
    if FShowSecondColumnForTotalCount then
      W1:= 80
    else
      W1:= 100;
    if FShowSecondColumn then
      W2:= 80
    else
      W2:= 150;


    //пробегаем по всем периодам
    for i:= 0 to FAdditionYearsCount do
    begin
      if FUsedReasons[0] then
      begin
        VAppend(ColWidths, W1);
         //Общее количество за этот период
        if FShowSecondColumnForTotalCount then
          VAppend(ColWidths, W1);
      end;
      for j:= 1 to High(FUsedReasons) do
      begin
        if FUsedReasons[j] then
        begin
          VAppend(ColWidths, W2); //Не расследовано за этот период
          if FShowSecondColumn then
            VAppend(ColWidths, W2); //%
        end;
      end;
    end;
  end;

  procedure SetWidthsReason;
  var
    i, j, W1, W2: Integer;
  begin
    if FSeveralYears or FShowSecondColumnForTotalCount then
      W1:= 80
    else
      W1:= 100;
    if FShowSecondColumn or FSeveralYears then
      W2:= 80
    else
      W2:= 150;


    if FUsedReasons[0] then
    begin
      for j:= 0 to FAdditionYearsCount do
      begin
        VAppend(ColWidths, W1);
        if FShowSecondColumnForTotalCount then
          VAppend(ColWidths, W1);
      end;
    end;
    for i:= 1 to High(FUsedReasons) do
    begin
      if FUsedReasons[i] then
      begin
        for j:= 0 to FAdditionYearsCount do
        begin
          VAppend(ColWidths, W2);
          if FShowSecondColumn then
            VAppend(ColWidths, W2); //%
        end;
      end;
    end;
  end;

begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;

  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;



  FAdditionYearsCount:= AAdditionYearsCount;
  FShowAllYearsTotalCount:= AShowAllYearsTotalCount;
  FShowSecondColumn:= AShowSecondColumn;

  FSeveralYears:= FAdditionYearsCount>0;
  FUsedReasons:= AUsedReasons;
  FGroupType:= AGroupType;
  FSecondColumnType:= ASecondColumnType;

   FShowSecondColumnForTotalCount:= FShowSecondColumn and
     ((FSecondColumnType=2{накопление}) or (FSeveralYears and (FSecondColumnType=1{%})));


  ColWidths:= nil;
  VAppend(ColWidths, AParamNameColumnWidth); //наименование параметра распределения
  if FShowAllYearsTotalCount and FSeveralYears then
    VAppend(ColWidths, 120); //общее количество за все периоды

  if FGroupType = 1 then
    SetWidthsYear
  else if FGroupType = 2 then
    SetWidthsReason;

  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);
end;

destructor TStatisticReclamationSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TStatisticReclamationSheet.Zoom(const APercents: Integer);
begin
  FWriter.SetZoom(APercents);
end;

procedure TStatisticReclamationSheet.Draw(const ABeginDate, AEndDate: TDate;
  const AReportName, AMotorNames, AParamCaption: String; const AParamNames,
  AReasonNames: TStrVector; const ACounts: TIntMatrix3D;
  const AResumeNeed: Boolean);
var
  R, FixedRowsCount: Integer;
begin
  FGrid.Clear;
  if VIsNil(AParamNames) then Exit;
  CalcTotalForAllPeriods(ACounts);
  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;
  R:= 1;
  DrawTop(R, ABeginDate, AEndDate, AReportName, AMotorNames);
  DrawTableCaption(R, ABeginDate, AParamCaption, AReasonNames);
  FixedRowsCount:= R;
  DrawData(R, AParamNames, ACounts, AResumeNeed);
  FWriter.SetFrozenRows(FixedRowsCount);
  FWriter.EndEdit;

end;

{ TStatisticReclamationMonthSheet }

//procedure TStatisticReclamationMonthSheet.DrawTitle(var ARow: Integer;
//  const AYears: TStrVector; const AName, AMotorNames: String);
//var
//  R: Integer;
//  S: String;
//begin
//  FWriter.SetAlignment(haCenter, vaCenter);
//
//  R:= ARow;
//  S:= 'Отчет по рекламационным случаям электродвигателей';
//  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
//  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);
//
//  if AMotorNames<>EmptyStr then
//  begin
//    R:= R + 1;
//    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
//    FWriter.WriteText(R, 1, R, FWriter.ColCount, AMotorNames, cbtNone, True, True);
//  end;
//
//  if not VIsNil(AYears) then
//  begin
//    R:= R + 1;
//    S:= 'за ' + VVectorToStr(AYears, '/') + 'гг.';
//    FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
//    FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);
//  end;
//
//  if AName<>EmptyStr then
//  begin
//    R:= R + 1;
//    S:= '(' + AName + ')';
//    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
//    FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);
//  end;
//
//  ARow:= R + 1;
//end;
//
//procedure TStatisticReclamationMonthSheet.DrawData(const ARow: Integer;
//  const AYears: TStrVector; const ACounts, AAccumulations: TIntMatrix;
//  const AResumeNeed: Boolean);
//var
//  i, j, R, C: Integer;
//begin
//  R:= ARow;
//  FWriter.WriteText(R, 1, R, FWriter.ColCount, EmptyStr, cbtNone);
//  FWriter.SetRowHeight(R, 10);
//
//  FWriter.SetAlignment(haCenter, vaCenter);
//
//  R:= R + 1;
//  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
//  FWriter.WriteText(R, 1, R+1, 1, 'Месяц', cbtOuter, True, True);
//  C:= 1;
//  for i:= 0 to High(AYears) do
//  begin
//    C:= C + 1;
//    FWriter.WriteText(R, C, R, C+1, AYears[i], cbtOuter, True, True);
//    FWriter.WriteText(R+1, C, R+1, C, 'Количество', cbtOuter, True, True);
//    C:= C + 1;
//    FWriter.WriteText(R+1, C, R+1, C, 'Накопление', cbtOuter, True, True);
//  end;
//  C:= 1;
//  FWriter.DrawBorders(R, C, R+1, FWriter.ColCount, cbtAll);
//
//  R:= R + 1;
//
//  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
//  for i:= 1 to 12 do
//  begin
//    R:= R + 1;
//    C:= 1;
//    FWriter.WriteText(R, C, R, C, MONTHSNOM[i], cbtOuter, True, True);
//    for j:= 0 to High(AYears) do
//    begin
//      C:= C + 1;
//      FWriter.WriteNumber(R, C, ACounts[j,i-1], cbtOuter);
//      C:= C + 1;
//      FWriter.WriteNumber(R, C, AAccumulations[j,i-1], cbtOuter);
//    end;
//  end;
//
//  if AResumeNeed then
//  begin
//    R:= R + 1;
//    C:= 1;
//    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
//    FWriter.WriteText(R, C, 'ИТОГО', cbtOuter, True, True);
//    C:= 0;
//    for j:= 0 to High(AYears) do
//    begin
//      C:= C + 2;
//      FWriter.WriteNumber(R, C, R, C+1, AAccumulations[j,11], cbtOuter);
//    end;
//  end;
//end;
//
//constructor TStatisticReclamationMonthSheet.Create(
//  const AGrid: TsWorksheetGrid; const AYearsCount: Integer);
//var
//  ColWidths: TIntVector;
//begin
//  FGrid:= AGrid;
//  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
//  FGrid.MouseWheelOption:= mwGrid;
//  FGrid.ShowGridLines:= False;
//  FGrid.ShowHeaders:= False;
//  FGrid.SelectionPen.Style:= psClear;
//
//  FFontName:= SHEET_FONT_NAME;
//  FFontSize:= SHEET_FONT_SIZE;
//
//  ColWidths:= VCreateInt([
//    100 // Месяц
//  ]);
//
//  if AYearsCount>0 then
//    VReDim(ColWidths, Length(ColWidths) + 2*AYearsCount, 100);  // Кол-во|Накопление
//
//  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);
//
//end;
//
//destructor TStatisticReclamationMonthSheet.Destroy;
//begin
//  if Assigned(FWriter) then FreeAndNil(FWriter);
//  inherited Destroy;
//end;
//
//procedure TStatisticReclamationMonthSheet.Draw(const AYears: TStrVector;
//  const AName, AMotorNames: String; const ACounts, AAccumulations: TIntMatrix;
//  const AResumeNeed: Boolean);
//var
//  R: Integer;
//begin
//  FGrid.Clear;
//  if VIsNil(AYears) then Exit;
//  if MIsNil(ACounts) then Exit;
//
//  FWriter.BeginEdit;
//  FWriter.SetBackgroundClear;
//  R:= 1;
//  DrawTitle(R, AYears,AName, AMotorNames);
//  DrawData(R, AYears, ACounts, AAccumulations, AResumeNeed);
//
//  FWriter.EndEdit;
//end;

{ TStatisticReclamationCountSheet }

//procedure TStatisticReclamationCountSheet.DrawTitle(var ARow: Integer;
//  const ABeginDate, AEndDate: TDate; const AName, AMotorNames: String);
//var
//  R: Integer;
//  S: String;
//begin
//
//  FWriter.SetAlignment(haCenter, vaCenter);
//
//  R:= ARow;
//  S:= 'Отчет по рекламационным случаям электродвигателей';
//  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
//  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);
//
//  if AMotorNames<>EmptyStr then
//  begin
//    R:= R + 1;
//    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
//    FWriter.WriteText(R, 1, R, FWriter.ColCount, AMotorNames, cbtNone, True, True);
//  end;
//
//  R:= R + 1;
//  S:= 'за ';
//  if SameDate(ABeginDate, AEndDate) then
//    S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
//  else
//    S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
//       ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);
//  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
//  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);
//
//  if AName<>EmptyStr then
//  begin
//    R:= R + 1;
//    S:= '(' + AName + ')';
//    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
//    FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);
//  end;
//
//  ARow:= R + 1;
//end;
//
//procedure TStatisticReclamationCountSheet.DrawData(const ARow: Integer;
//  const AFirstColName: String;
//  const ANameValues, ATitleReasonNames: TStrVector;
//  const AUsedReasons: TBoolVector; const AReasonCountValues: TIntMatrix;
//  const AResumeNeed: Boolean);
//var
//  i, j, R, C, TotalCount, Count: Integer;
//  ShowTotal, ShowReason: Boolean;
//begin
//  R:= ARow;
//
//  ShowTotal:= AUsedReasons[0];
//  ShowReason:= VIsTrue(AUsedReasons, 1);
//
//  TotalCount:= VSum(AReasonCountValues[0]);
//
//  FWriter.WriteText(R, 1, R, FWriter.ColCount, EmptyStr, cbtNone);
//  FWriter.SetRowHeight(R, 10);
//
//  R:= R + 1;
//  C:= 1;
//  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
//  FWriter.SetAlignment(haCenter, vaCenter);
//  FWriter.WriteText(R, C, AFirstColName, cbtOuter, True, True);
//  //if ShowReason then
//  //  FWriter.WriteText(R, C, R+1, C, AFirstColName, cbtOuter, True, True)
//  //else
//  //  FWriter.WriteText(R, C, AFirstColName, cbtOuter, True, True);
//
//  if ShowTotal then
//  begin
//    C:= C + 1;
//    FWriter.WriteText(R, C, ATitleReasonNames[0], cbtOuter, True, True);
//    //if ShowReason then
//    //  FWriter.WriteText(R, C, R+1, C, ATitleReasonNames[0], cbtOuter, True, True)
//    //else
//    //  FWriter.WriteText(R, C, ATitleReasonNames[0], cbtOuter, True, True);
//  end;
//
//  if ShowReason then
//  begin
//    C:= Ord(ShowTotal);
//    for i:= 1 to High(ATitleReasonNames) do
//    begin
//      if AUsedReasons[i] then
//      begin
//        C:= C + 2;
//        FWriter.WriteText(R, C, R, C+1, ATitleReasonNames[i], cbtOuter, True, True);
//        //FWriter.WriteText(R+1, C, 'Количество', cbtOuter, True, True);
//        //FWriter.WriteText(R+1, C+1, '%', cbtOuter, True, True);
//      end;
//    end;
//    //R:= R + 1;
//  end;
//
//  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
//  for i:= 0 to High(ANameValues) do
//  begin
//    R:= R + 1;
//    FWriter.SetAlignment(haLeft, vaCenter);
//    C:= 1;
//    FWriter.WriteText(R, C, ANameValues[i], cbtOuter, True, True);
//    FWriter.SetAlignment(haCenter, vaCenter);
//    if ShowTotal then
//    begin
//      C:= C + 1;
//      FWriter.WriteNumber(R, C, AReasonCountValues[0, i], cbtOuter);
//    end;
//    C:= Ord(ShowTotal);
//    for j:= 1 to High(ATitleReasonNames) do
//    begin
//      if AUsedReasons[j] then
//      begin
//        C:= C + 2;
//        Count:= AReasonCountValues[j, i];
//        FWriter.WriteNumber(R, C, Count, cbtOuter);
//        FWriter.WriteNumber(R, C+1, Count/TotalCount, 1, cbtOuter, nfPercentage);
//      end;
//    end;
//  end;
//  if AResumeNeed then
//  begin
//    R:= R + 1;
//    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
//    FWriter.SetAlignment(haLeft, vaCenter);
//    C:= 1;
//    FWriter.WriteText(R, C, 'ИТОГО', cbtOuter, True, True);
//    FWriter.SetAlignment(haCenter, vaCenter);
//
//    if ShowTotal then
//    begin
//      C:= C + 1;
//      FWriter.WriteNumber(R, C, TotalCount, cbtOuter);
//    end;
//    C:= Ord(ShowTotal);
//    for i:= 1 to High(ATitleReasonNames) do
//    begin
//      if AUsedReasons[i] then
//      begin
//        C:= C + 2;
//        Count:= VSum(AReasonCountValues[i]);
//        FWriter.WriteNumber(R, C, Count, cbtOuter);
//        FWriter.WriteNumber(R, C+1, Count/TotalCount, 1, cbtOuter, nfPercentage);
//      end;
//    end;
//  end;
//end;
//
//constructor TStatisticReclamationCountSheet.Create(const AGrid: TsWorksheetGrid;
//  const AReasonsCount: Integer);
//var
//  ColWidths: TIntVector;
//begin
//  FGrid:= AGrid;
//  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
//  FGrid.MouseWheelOption:= mwGrid;
//  FGrid.ShowGridLines:= False;
//  FGrid.ShowHeaders:= False;
//  FGrid.SelectionPen.Style:= psClear;
//
//  FFontName:= SHEET_FONT_NAME;
//  FFontSize:= SHEET_FONT_SIZE;
//
//  ColWidths:= VCreateInt([
//    270 // Наименование двигателя
//
//  ]);
//
//  if AReasonsCount>0 then
//    VReDim(ColWidths, Length(ColWidths) + AReasonsCount, 80);
//
//  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);
//
//end;
//
//destructor TStatisticReclamationCountSheet.Destroy;
//begin
//  if Assigned(FWriter) then FreeAndNil(FWriter);
//  inherited Destroy;
//end;
//
//procedure TStatisticReclamationCountSheet.Draw(const ABeginDate, AEndDate: TDate;
//                   const AName, AMotorNames, AFirstColName: String;
//                   const ANameValues, ATitleReasonNames: TStrVector;
//                   const AUsedReasons: TBoolVector;
//                   const AReasonCountValues: TIntMatrix;
//                   const AResumeNeed: Boolean);
//var
//  R: Integer;
//begin
//  FGrid.Clear;
//  if VIsNil(ANameValues) then Exit;
//
//  FWriter.BeginEdit;
//  FWriter.SetBackgroundClear;
//  R:= 1;
//  DrawTitle(R, ABeginDate, AEndDate, AName, AMotorNames);
//  DrawData(R, AFirstColName, ANameValues, ATitleReasonNames, AUsedReasons,
//           AReasonCountValues, AResumeNeed);
//
//  FWriter.EndEdit;
//end;

{ TReportShipmentSheet }

constructor TReportShipmentSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;

  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  ColWidths:= VCreateInt([
    150, // № п/п - Дата отправки
    300, // Наименование двигателя
    150, // Номер двигателя
    300  // Грузополучатель
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TReportShipmentSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TReportShipmentSheet.DrawReport(const ASingleReceiver: Boolean;
                   const ANeedNumberList: Boolean;
                   const ABeginDate, AEndDate: TDate;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCount: TIntVector;
                   const ARecieverNames: TStrVector;
                   const ARecieverMotorNames: TStrMatrix;
                   const ARecieverMotorCounts: TIntMatrix;
                   const AListSendDates: TDateVector;
                   const AListMotorNames, AListMotorNums, AListReceiverNames: TStrVector);
var
  R, i, j, C: Integer;
  S: String;
begin
  FGrid.Clear;
  if VIsNil(ATotalMotorNames) then Exit;

  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;

  C:= 3 + Ord(ANeedNumberList);

  S:= 'Отчет по отгрузке электродвигателей за ';
  if SameDate(ABeginDate, AEndDate) then
    S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
  else
    S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
       ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);

  //Заголовок отчета
  R:= 1;
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);

  if not ASingleReceiver then
  begin

    FWriter.WriteText(R, 1, R, C, S, cbtNone, True, True);
    //Таблица сводная: наименование/количество
    R:= R + 1;
    FWriter.WriteText(R, 1, R, C, EmptyStr, cbtNone);
    //FWriter.SetRowHeight(R, 10);
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
    FWriter.WriteText(R, 1, R, C, 'Всего отгружено', cbtNone, True, True);
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
    FWriter.WriteText(R, 3, 'Количество', cbtOuter, True, True);
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
    for i:= 0 to High(ATotalMotorNames) do
    begin
      R:= R + 1;
      FWriter.SetAlignment(haLeft, vaCenter);
      FWriter.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
      FWriter.SetAlignment(haCenter, vaCenter);
      FWriter.WriteNumber(R, 3, ATotalMotorCount[i], cbtOuter);
    end;
  end
  else
    FWriter.WriteText(R, 1, R, 3, S, cbtNone, True, True);

  for i:= 0 to High(ARecieverNames) do
  begin
    R:= R + 1;
    FWriter.WriteText(R, 1, R, 3, EmptyStr, cbtNone);
    //FWriter.SetRowHeight(R, 15);
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
    FWriter.WriteText(R, 1, R, 3, ARecieverNames[i], cbtNone, True, True);
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
    FWriter.WriteText(R, 3, 'Количество', cbtOuter, True, True);
    for j:= 0 to High(ARecieverMotorNames[i]) do
    begin
      R:= R + 1;
      FWriter.SetFont(FFontName, FFontSize, [], clBlack);
      FWriter.SetAlignment(haLeft, vaCenter);
      FWriter.WriteText(R, 1, R, 2, ARecieverMotorNames[i,j], cbtOuter, True, True);
      FWriter.SetAlignment(haCenter, vaCenter);
      FWriter.WriteNumber(R, 3, ARecieverMotorCounts[i,j], cbtOuter);
    end;
  end;

  if ANeedNumberList then
  begin
    R:= R + 1;
    if not ASingleReceiver then
      FWriter.WriteText(R, 1, R, 4, EmptyStr, cbtNone)
    else
      FWriter.WriteText(R, 1, R, 3, EmptyStr, cbtNone);
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
    if not ASingleReceiver then
      FWriter.WriteText(R, 1, R, 4, 'Пономерной список', cbtNone, True, True)
    else
      FWriter.WriteText(R, 1, R, 3, 'Пономерной список', cbtNone, True, True);
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R, 1, 'Дата отгрузки', cbtOuter, True, True);
    FWriter.WriteText(R, 2, 'Наименование', cbtOuter, True, True);
    FWriter.WriteText(R, 3, 'Номер', cbtOuter, True, True);
    if not ASingleReceiver then
      FWriter.WriteText(R, 4, 'Грузополучатель', cbtOuter, True, True);


    FWriter.SetFont(FFontName, FFontSize, [], clBlack);

    for i:= 0 to High(AListSendDates) do
    begin
      R:= R + 1;
      FWriter.SetAlignment(haCenter, vaCenter);
      FWriter.WriteDate(R, 1, AListSendDates[i], cbtOuter);
      FWriter.SetAlignment(haLeft, vaCenter);
      FWriter.WriteText(R, 2, AListMotorNames[i], cbtOuter, True, True);
      FWriter.SetAlignment(haCenter, vaCenter);
      FWriter.WriteText(R, 3, AListMotorNums[i], cbtOuter, True, True);
      if not ASingleReceiver then
      begin
        FWriter.SetAlignment(haLeft, vaCenter);
        FWriter.WriteText(R, 4, AListReceiverNames[i], cbtOuter, True, True);
      end;
    end;
  end;




  FWriter.EndEdit;
end;

{ TBeforeTestSheet }

constructor TBeforeTestSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;

  //LoadFontFromControl(FGrid, FFontName, FFontSize);
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  ColWidths:= VCreateInt([
    150, // Дата сборки
    300, // Наименование двигателя
    200, // Номер двигателя
    300  // Примечание
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TBeforeTestSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TBeforeTestSheet.Draw(const ABuildDates: TDateVector;
  const AMotorNames, AMotorNums: TStrVector;
  const ATotalMotorNames: TStrVector; const ATotalMotorCounts: TIntVector;
  const ATestDates: TDateMatrix;
  const ATestFails: TIntMatrix;
  const ATestNotes: TStrMatrix);
var
  R, N, i, j, FrozenCount: Integer;
  S: String;
begin
  FGrid.Clear;
  if VIsNil(ABuildDates) then Exit;

  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;

  //S:= 'Список электродвигателей, не прошедших испытания, на ' +
  //    FormatDateTime('dd.mm.yyyy', Date);

  //Заголовок отчета
  R:= 1;
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, R, 4, 'На испытаниях', cbtNone, True, True);
  //Таблица сводная: наименование/количество
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Количество', cbtOuter, True, True);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  for i:= 0 to High(ATotalMotorNames) do
  begin
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteNumber(R, 3, ATotalMotorCounts[i], cbtOuter);
  end;
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, R, 2, 'ИТОГО', cbtOuter, True, True);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteNumber(R, 3, VSum(ATotalMotorCounts), cbtOuter);

  //Таблица пономерная
  R:= R + 2;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, 'Дата сборки', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Испытания', cbtOuter, True, True);

  FrozenCount:= R;


  for i:= 0 to High(ABuildDates) do
  begin
    R:= R + 1;
    N:= 0;
    if not VIsNil(ATestDates[i]) then
      N:= High(ATestDates[i]);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
    FWriter.WriteDate(R, 1, R+N, 1, ABuildDates[i], cbtOuter);
    FWriter.WriteText(R, 2, R+N, 2, AMotorNames[i], cbtOuter, True, True);
    FWriter.WriteText(R, 3, R+N, 3, AMotorNums[i], cbtOuter, True, True);
    FWriter.SetAlignment(haLeft, vaCenter);
    if VIsNil(ATestDates[i]) then
      FWriter.WriteText(R, 4, R+N, 4, 'Не проводились', cbtOuter, True, True)
    else begin
      for j:= 0 to High(ATestDates[i]) do
      begin
        if ATestFails[i][j]=0 then
        begin
          FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
          S:= ' - норма';
        end
        else begin
          FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clRed);
          S:= ' - брак';
        end;
        S:= FormatDateTime('dd.mm.yyyy', ATestDates[i][j]) + S;
        if ATestNotes[i][j]<>EmptyStr then
          S:= S + ' (' + ATestNotes[i][j] + ')';
        FWriter.WriteText(R+j, 4, S, cbtOuter, True, True);
      end;

    end;
    R:= R + N;
  end;

  FWriter.SetFrozenRows(FrozenCount);
  FWriter.SetRepeatedRows(FrozenCount,FrozenCount);

  FWriter.EndEdit;

end;

{ TMotorCardSheet }

constructor TMotorCardSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;

  //LoadFontFromControl(FGrid, FFontName, FFontSize);
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  ColWidths:= VCreateInt([
    130, // Дата уведомления
    130, // Пробег, км
    130, // Предприятие
    130, // Завод
    130, // Выезд/ФИО
    200, // Неисправный элемент
    200,  // Причина неисправности
    100,  //Особое мнение
    350  //Примечание
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TMotorCardSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TMotorCardSheet.Zoom(const APercents: Integer);
begin
  FWriter.SetZoom(APercents);
end;

procedure TMotorCardSheet.Draw(const ABuildDate, ASendDate: TDate;
      const AMotorName, AMotorNum, ASeries, ARotorNum, AReceiverName, AControlNote: String;
      const ATestDates: TDateVector;
      const ATestResults: TIntVector;
      const ATestNotes: TStrVector;
      const ARecDates: TDateVector;
      const AMileages, AOpinions: TIntVector;
      const APlaceNames, AFactoryNames, ADepartures,
      ADefectNames, AReasonNames, ARecNotes: TStrVector);
var
  R, i: Integer;
  S: String;
begin
  FGrid.Clear;
  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;

  FWriter.SetAlignment(haCenter, vaCenter);

  R:= 1;
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount,  'УЧЁТНАЯ КАРТОЧКА ЭЛЕКТРОДВИГАТЕЛЯ');

  R:= R + 1;
  FWriter.WriteText(R, 1,  EmptyStr);
  FWriter.SetRowHeight(R, 10);
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, 2,  'Наименование', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Номер', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Партия', cbtOuter, True, True);
  FWriter.WriteText(R, 5, 'Дата сборки', cbtOuter, True, True);
  FWriter.WriteText(R, 6, 'Ротор', cbtOuter, True, True);
  FWriter.WriteText(R, 7, R, FWriter.ColCount,  'Отгружен', cbtOuter, True, True);
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [], clBlack);
  FWriter.WriteText(R, 1, R, 2,  AMotorName, cbtOuter, True, True);
  FWriter.WriteText(R, 3, AMotorNum, cbtOuter, True, True);
  FWriter.WriteText(R, 4, ASeries, cbtOuter, True, True);
  FWriter.WriteDate(R, 5, ABuildDate, cbtOuter);
  FWriter.WriteText(R, 6, ARotorNum, cbtOuter, True, True);
  FWriter.SetAlignment(haLeft, vaCenter);
  S:= EmptyStr;
  if AReceiverName<>EmptyStr then
    S:= FormatDateTime('dd.mm.yyyy', ASendDate) + ' - ' + AReceiverName;
  FWriter.WriteText(R, 7, R, FWriter.ColCount,  S, cbtOuter, True, True);

  R:= R + 1;
  FWriter.WriteText(R, 1,  EmptyStr);
  FWriter.SetRowHeight(R, 10);
  R:= R + 1;
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.SetFont(FFontName, FFontSize+1, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount,  'Приемо-сдаточные испытания:');
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, 'Дата', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Результат', cbtOuter, True, True);
  FWriter.WriteText(R, 3, R, FWriter.ColCount, 'Примечание', cbtOuter, True, True);
  FWriter.SetFont(FFontName, FFontSize, [], clBlack);
  if VIsNil(ATestDates) then
  begin
    R:= R + 1;
    FWriter.WriteText(R, 1, EmptyStr, cbtOuter);
    FWriter.WriteText(R, 2, EmptyStr, cbtOuter, True, True);
    FWriter.WriteText(R, 3, R, FWriter.ColCount, EmptyStr, cbtOuter, True, True);
  end
  else begin
    for i:= 0 to High(ATestDates) do
    begin
      R:= R + 1;
      if ATestResults[i]=0 then
        S:= 'норма'
      else
        S:= 'брак';
      FWriter.SetAlignment(haCenter, vaCenter);
      FWriter.WriteDate(R, 1, ATestDates[i], cbtOuter);
      FWriter.WriteText(R, 2, S, cbtOuter, True, True);
      FWriter.SetAlignment(haLeft, vaCenter);
      FWriter.WriteText(R, 3, R, FWriter.ColCount, ATestNotes[i], cbtOuter, True, True);
    end;
  end;

  R:= R + 1;
  FWriter.WriteText(R, 1,  EmptyStr);
  FWriter.SetRowHeight(R, 10);
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize+1, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, R, FWriter.ColCount,  'Контроль:');
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount,  AControlNote, cbtOuter, True, True);

  R:= R + 1;
  FWriter.WriteText(R, 1,  EmptyStr);
  FWriter.SetRowHeight(R, 10);
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize+1, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, R, FWriter.ColCount,  'Рекламации:');
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, 'Дата уведомления', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Пробег, км', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Предприятие', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Завод', cbtOuter, True, True);
  FWriter.WriteText(R, 5, 'Выезд', cbtOuter, True, True);
  FWriter.WriteText(R, 6, 'Неисправный элемент', cbtOuter, True, True);
  FWriter.WriteText(R, 7, 'Причина неисправности', cbtOuter, True, True);
  FWriter.WriteText(R, 8, 'Особое мнение', cbtOuter, True, True);
  FWriter.WriteText(R, 9, 'Примечание', cbtOuter, True, True);


  FWriter.SetFont(FFontName, FFontSize, [], clBlack);
  if VIsNil(ARecDates) then
  begin
    R:= R + 1;
    for i:= 1 to 9 do
      FWriter.WriteText(R, i, EmptyStr, cbtOuter);
  end
  else begin
    for i:= 0 to High(ARecDates) do
    begin
      R:= R + 1;
      FWriter.SetAlignment(haCenter, vaTop);
      FWriter.WriteDate(R, 1, ARecDates[i], cbtOuter);
      if AMileages[i]>=0 then
        FWriter.WriteNumber(R, 2, AMileages[i], cbtOuter)
      else
        FWriter.WriteText(R, 2, EmptyStr, cbtOuter);
      FWriter.WriteText(R, 3, APlaceNames[i], cbtOuter, True, True);
      FWriter.WriteText(R, 4, AFactoryNames[i], cbtOuter, True, True);
      FWriter.WriteText(R, 5, ADepartures[i], cbtOuter, True, True);
      FWriter.WriteText(R, 6, ADefectNames[i], cbtOuter, True, True);
      FWriter.WriteText(R, 7, AReasonNames[i], cbtOuter, True, True);
      S:= EmptyStr;
      if AOpinions[i]=1 then S:= CHECK_SYMBOL;
      FWriter.WriteText(R, 8, S, cbtOuter);
      FWriter.SetAlignment(haLeft, vaTop);
      FWriter.WriteText(R, 9, ARecNotes[i], cbtOuter, True, True);
    end;
  end;


  R:= R + 1;
  FWriter.WriteText(R, 1,  EmptyStr);
  FWriter.SetRowHeight(R, 10);
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize+1, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, R, FWriter.ColCount,  'Гарантийный ремонт:');
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, 'Дата прибытия', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Наличие паспорта', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Дата убытия', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Срок ремонта (рабочих дней)', cbtOuter, True, True);
  FWriter.WriteText(R, 5, R, FWriter.ColCount,  'Примечание', cbtOuter, True, True);

  R:= R + 1;
  for i:= 1 to 9 do
    FWriter.WriteText(R, i, EmptyStr, cbtTop);



  FWriter.EndEdit;
end;

{ TSeriesSheet }

procedure TSeriesSheet.DrawTitle;
var
  R: Integer;
begin
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.SetBackground(COLOR_BACKGROUND_TITLE);

  R:= 1;
  FWriter.WriteText(R, 1, 'Номер партии', cbtOuter, True, True);
end;

constructor TSeriesSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;

  //LoadFontFromControl(FGrid, FFontName, FFontSize);
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  ColWidths:= VCreateInt([
    170 // Номер партии
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TSeriesSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TSeriesSheet.DrawLine(const AIndex: Integer; const ASelected: Boolean);
var
  R: Integer;
begin
  R:= AIndex + 2;

  FWriter.SetAlignment(haCenter, vaCenter);

  if ASelected then
  begin
    FWriter.SetBackground(COLOR_BACKGROUND_SELECTED);
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  end
  else begin
    FWriter.SetBackgroundClear;
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  end;

  FWriter.WriteText(R, 1, Series[AIndex], cbtOuter, True, True);
end;

procedure TSeriesSheet.Draw(const ASeries: TStrVector);
var
  i: Integer;
begin
  FGrid.Clear;

  FWriter.BeginEdit;

  Series:= ASeries;

  DrawTitle;

  if not VIsNil(Series) then
  begin
    for i:= 0 to High(Series) do
      DrawLine(i, False);
    FWriter.SetFrozenRows(1);
  end;
  FWriter.EndEdit;
end;

{ TStoreSheet }

constructor TStoreSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;

  //LoadFontFromControl(FGrid, FFontName, FFontSize);
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  ColWidths:= VCreateInt([
    150, // Дата сдачи (испытаний)
    300, // Наименование двигателя
    200//, // Номер двигателя
    //100  // Номер партии
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);
end;

destructor TStoreSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TStoreSheet.Draw(const ADeltaDays: Integer;
       const ATestDates: TDateVector;
       const AMotorNames, AMotorNums{, ASeries}: TStrVector;
       const ATotalMotorNames: TStrVector;
       const ATotalMotorCounts: TIntVector);
var
  R, i, FrozenCount: Integer;
  S: String;
begin
  FGrid.Clear;
  if VIsNil(ATestDates) then Exit;

  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;

  S:= 'Отчет по наличию электродвигателей на складе на ' +
      FormatDateTime('dd.mm.yyyy', Date);
  if ADeltaDays>0 then
    S:= S + ', хранящихся дольше ' + IntToStr(ADeltaDays) +
            ' дней (сданы ' +
            FormatDateTime('dd.mm.yyyy', IncDay(Date,-ADeltaDays)) +
            ' и ранее)';

  //Заголовок отчета
  R:= 1;
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, R, 3, S, cbtNone, True, True);
  //Таблица сводная: наименование/количество
  R:= R + 2;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Количество', cbtOuter, True, True);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  for i:= 0 to High(ATotalMotorNames) do
  begin
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteNumber(R, 3, ATotalMotorCounts[i], cbtOuter);
  end;
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, R, 2, 'ИТОГО', cbtOuter, True, True);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteNumber(R, 3, VSum(ATotalMotorCounts), cbtOuter);

  //Таблица пономерная
  R:= R + 2;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, 'Дата сдачи', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
  //FWriter.WriteText(R, 4, 'Номер партии', cbtOuter, True, True);

  FrozenCount:= R;

  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  for i:= 0 to High(ATestDates) do
  begin
    R:= R + 1;
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteDate(R, 1, ATestDates[i], cbtOuter);
    FWriter.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
    FWriter.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
    //FWriter.WriteText(R, 4, ASeries[i], cbtOuter, True, True);
  end;

  FWriter.SetFrozenRows(FrozenCount);

  FWriter.EndEdit;

end;

{ TMotorTestSheet }

procedure TMotorTestSheet.DrawTitle;
var
  R: Integer;
begin




  R:= 1;
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, 5, 'Журнал испытаний за ' +
                    FormatDateTime('dd.mm.yyyy', TestDate),
                    cbtBottom, True, True);
  R:= R+1;
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  //FWriter.SetBackground(COLOR_BACKGROUND_TITLE);
  //FWriter.WriteText(R, 1, 'Дата испытаний', cbtOuter, True, True);
  FWriter.WriteText(R, 1, '№ п/п', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
  //FWriter.WriteText(R, 4, 'Номер партии', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Результат испытаний', cbtOuter, True, True);
  FWriter.WriteText(R, 5, 'Примечание', cbtOuter, True, True);


end;

constructor TMotorTestSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;



  //LoadFontFromControl(FGrid, FFontName, FFontSize);
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  //FFontName:= Screen.SystemFont.Name;
  //FFontSize:= Screen.SystemFont.Size;

  ColWidths:= VCreateInt([
    150, // Дата испытаний
    300, // Наименование двигателя
    200, // Номер двигателя
    //150, // Номер партии
    200, // Результат (норма/брак)
    300  // Примечание
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TMotorTestSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TMotorTestSheet.DrawLine(const AIndex: Integer;
                                 const ASelected: Boolean);
var
  R: Integer;
  S: String;
begin
  R:= AIndex + 3;

  if ASelected then
    FWriter.SetBackground(COLOR_BACKGROUND_SELECTED)
  else
    FWriter.SetBackgroundClear;

  if TestResults[AIndex]=0 then
  begin
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
    S:= 'норма';
  end
  else begin
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clRed);
    S:= 'брак';
  end;

  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteNumber(R, 1, AIndex+1, cbtOuter);
  FWriter.WriteText(R, 2, MotorNames[AIndex], cbtOuter, True, True);
  FWriter.WriteText(R, 3, MotorNums[AIndex], cbtOuter, True, True);
  FWriter.WriteText(R, 4, S, cbtOuter, True, True);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 5, TestNotes[AIndex], cbtOuter, True, True);
end;

procedure TMotorTestSheet.Draw({const ATestDates: TDateVector; }
            const ATestDate: TDate;
            const AMotorNames, AMotorNums, ATestNotes{, ASeries}: TStrVector;
            const ATestResults, ATotalMotorCounts, ATotalFailCounts: TIntVector);
var
  i, N, K: Integer;
//  S: String;
begin
  FGrid.Clear;
  FWriter.BeginEdit;

  //TestDates:= ATestDates;
  TestDate:= ATestDate;
  MotorNames:= AMotorNames;
  MotorNums:= AMotorNums;
  TestResults:= ATestResults;
  TestNotes:= ATestNotes;
  //Series:= ASeries;

  DrawTitle;

  if not VIsNil(AMotorNames) then
  begin
    for i:= 0 to High(AMotorNames) do
      DrawLine(i, False);

    i:= Length(AMotorNames) + 3;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    N:= VSum(ATotalMotorCounts);
    K:= VSum(ATotalFailCounts);
    //S:= 'Норма: ' + IntToStr(N-K) + '            Брак: ' + IntToStr(K);
    FWriter.WriteText(i, 1,  'Норма: ' + IntToStr(N-K), cbtTop, True, True);
    FWriter.WriteText(i, 2,  'Брак: ' + IntToStr(K), cbtTop, True, True);
    FWriter.SetFrozenRows(2);
  end;


  FWriter.EndEdit;

end;

procedure TMotorTestSheet.DrawReport(const ABeginDate, AEndDate: TDate;
                   const ANeedNumberList: Boolean;
                   const ATestDates: TDateVector;
                   const AMotorNames, AMotorNums, ATestNotes{, ASeries}: TStrVector;
                   const ATestResults: TIntVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCounts, ATotalFailCounts: TIntVector);
var
  R, i, FrozenCount, N, K: Integer;
  S: String;
begin
  FGrid.Clear;
  if VIsNil(ATestDates) then Exit;

  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;

  S:= 'Отчет по испытанию электродвигателей за ';
  if SameDate(ABeginDate, AEndDate) then
    S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
  else
    S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
       ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);

  //Заголовок отчета
  R:= 1;
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  if ANeedNumberList then
    FWriter.WriteText(R, 1, R, 5, S, cbtNone, True, True)
  else
    FWriter.WriteText(R, 1, R, 4, S, cbtNone, True, True);
  //Таблица сводная: наименование/количество
  R:= R + 2;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Норма', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Брак', cbtOuter, True, True);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  FrozenCount:= R;
  for i:= 0 to High(ATotalMotorNames) do
  begin
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
    FWriter.SetAlignment(haCenter, vaCenter);
    N:= ATotalMotorCounts[i];
    K:= ATotalFailCounts[i];
    FWriter.WriteNumber(R, 3, N-K, cbtOuter);
    FWriter.WriteNumber(R, 4, K, cbtOuter);
  end;
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, R, 2, 'ИТОГО', cbtOuter, True, True);
  FWriter.SetAlignment(haCenter, vaCenter);
  N:= VSum(ATotalMotorCounts);
  K:= VSum(ATotalFailCounts);
  FWriter.WriteNumber(R, 3, N-K, cbtOuter);
  FWriter.WriteNumber(R, 4, K, cbtOuter);

  if ANeedNumberList then
  begin
    //Таблица пономерная
    R:= R + 2;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R, 1, 'Дата испытаний', cbtOuter, True, True);
    FWriter.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
    FWriter.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
    FWriter.WriteText(R, 4, 'Результат испытаний', cbtOuter, True, True);
    FWriter.WriteText(R, 5, 'Примечание', cbtOuter, True, True);
    FrozenCount:= R;

    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
    for i:= 0 to High(ATestDates) do
    begin
      R:= R + 1;
      if ATestResults[i]=0 then
      begin
        S:= 'норма';
        FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
      end
      else begin
        S:= 'брак';
        FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clRed);
      end;
      FWriter.SetAlignment(haCenter, vaCenter);
      FWriter.WriteDate(R, 1, ATestDates[i], cbtOuter);
      FWriter.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
      FWriter.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
      FWriter.WriteText(R, 4, S, cbtOuter, True, True);
      FWriter.SetAlignment(haLeft, vaCenter);
      FWriter.WriteText(R, 5, ATestNotes[i], cbtOuter, True, True);
    end;
  end;

  FWriter.SetFrozenRows(FrozenCount);

  FWriter.EndEdit;
end;

{ TMotorBuildSheet }

procedure TMotorBuildSheet.DrawTitle;
var
  R: Integer;
begin




  R:= 1;
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, 4, 'Журнал сборки за ' +
                    FormatDateTime('dd.mm.yyyy', BuildDate),
                    cbtBottom, True, True);

  R:= R+1;
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  //FWriter.SetBackground(COLOR_BACKGROUND_TITLE);
  //FWriter.WriteText(R, 1, 'Дата сборки', cbtOuter, True, True);
  FWriter.WriteText(R, 1, '№ п/п', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Номер ротора', cbtOuter, True, True);
end;

constructor TMotorBuildSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;


  //LoadFontFromControl(FGrid, FFontName, FFontSize);
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  ColWidths:= VCreateInt([
    150, // № п/п - Дата сборки
    300, // Наименование двигателя
    200, // Номер двигателя
    200  // Номер Ротора
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TMotorBuildSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TMotorBuildSheet.DrawLine(const AIndex: Integer;
                                  const ASelected: Boolean);
var
  R: Integer;
begin
  //R:= AIndex + 2;
  R:= AIndex + 3;

  FWriter.SetAlignment(haCenter, vaCenter);

  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  if ASelected then
    FWriter.SetBackground(COLOR_BACKGROUND_SELECTED)
  else
    FWriter.SetBackgroundClear;

  FWriter.WriteNumber(R, 1, AIndex+1, cbtOuter);
  FWriter.WriteText(R, 2, MotorNames[AIndex], cbtOuter, True, True);
  FWriter.WriteText(R, 3, MotorNums[AIndex], cbtOuter, True, True);
  FWriter.WriteText(R, 4, RotorNums[AIndex], cbtOuter, True, True);
end;

procedure TMotorBuildSheet.Draw({const ABuildDates: TDateVector;}
                        const ABuildDate: TDate;
                        const AMotorNames, AMotorNums, ARotorNums: TStrVector);
var
  i: Integer;
begin
  FGrid.Clear;

  FWriter.BeginEdit;

  //BuildDates:= ABuildDates;
  BuildDate:= ABuildDate;
  MotorNames:= AMotorNames;
  MotorNums:= AMotorNums;
  RotorNums:= ARotorNums;

  DrawTitle;

  if not VIsNil(AMotorNames) then
  begin
    for i:= 0 to High(AMotorNames) do
      DrawLine(i, False);
    FWriter.SetFrozenRows(2);
  end;


  FWriter.EndEdit;

end;

procedure TMotorBuildSheet.DrawReport(const ABeginDate, AEndDate: TDate;
                   const ANeedNumberList: Boolean;
                   const ABuildDates: TDateVector;
                   const AMotorNames, AMotorNums, ARotorNums: TStrVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCount: TIntVector);
var
  R, i, FrozenCount: Integer;
  S: String;
begin
  FGrid.Clear;
  if VIsNil(ABuildDates) then Exit;

  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;

  S:= 'Отчет по сборке электродвигателей за ';
  if SameDate(ABeginDate, AEndDate) then
    S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
  else
    S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
       ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);

  //Заголовок отчета
  R:= 1;
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  if ANeedNumberList then
    FWriter.WriteText(R, 1, R, 4, S, cbtNone, True, True)
  else
    FWriter.WriteText(R, 1, R, 3, S, cbtNone, True, True);
  //Таблица сводная: наименование/количество
  R:= R + 2;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Количество', cbtOuter, True, True);
  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  FrozenCount:= R;
  for i:= 0 to High(ATotalMotorNames) do
  begin
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteNumber(R, 3, ATotalMotorCount[i], cbtOuter);
  end;
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, R, 2, 'ИТОГО', cbtOuter, True, True);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteNumber(R, 3, VSum(ATotalMotorCount), cbtOuter);

  if ANeedNumberList then
  begin
    //Таблица пономерная
    R:= R + 2;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R, 1, 'Дата сборки', cbtOuter, True, True);
    FWriter.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
    FWriter.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
    FWriter.WriteText(R, 4, 'Номер ротора', cbtOuter, True, True);
    FrozenCount:= R;
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
    for i:= 0 to High(ABuildDates) do
    begin
      R:= R + 1;
      FWriter.WriteDate(R, 1, ABuildDates[i], cbtOuter);
      FWriter.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
      FWriter.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
      FWriter.WriteText(R, 4, ARotorNums[i], cbtOuter, True, True);
    end;
  end;

  FWriter.SetFrozenRows(FrozenCount);

  FWriter.EndEdit;
end;

{ TReclamationSheet }

procedure TReclamationSheet.DrawTitle;
var
  R: Integer;
begin
  R:= 1;
  FWriter.SetBackgroundClear;
  //FWriter.SetBackground(COLOR_BACKGROUND_TITLE);

  FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);

  FWriter.WriteText(R, 1, '№п/п', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Дата уведомления', cbtOuter, True, True);
  FWriter.WriteText(R, 3, 'Двигатель', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Номер', cbtOuter, True, True);
  FWriter.WriteText(R, 5, 'Дата сборки', cbtOuter, True, True);
  FWriter.WriteText(R, 6, 'Пробег, км', cbtOuter, True, True);
  FWriter.WriteText(R, 7, 'Предприятие', cbtOuter, True, True);
  FWriter.WriteText(R, 8, 'Завод', cbtOuter, True, True);
  FWriter.WriteText(R, 9, 'Выезд', cbtOuter, True, True);
  FWriter.WriteText(R, 10, 'Неисправный элемент', cbtOuter, True, True);
  FWriter.WriteText(R, 11, 'Причина неисправности', cbtOuter, True, True);
  FWriter.WriteText(R, 12, 'Особое мнение', cbtOuter, True, True);
  FWriter.WriteText(R, 13, 'Примечание', cbtOuter, True, True);
  FWriter.WriteText(R, 14, 'Прибыл в ремонт', cbtOuter, True, True);
  //FWriter.WriteText(R, 15, 'Наличие паспорта', cbtOuter, True, True);
  FWriter.WriteText(R, 15, 'Убыл из ремонта', cbtOuter, True, True);
  //FWriter.WriteText(R, 17, 'Срок, дней', cbtOuter, True, True);

end;

constructor TReclamationSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;

  //LoadFontFromControl(FGrid, FFontName, FFontSize);
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE - 1;


  ColWidths:= VCreateInt([
    45,  //№п/п
    100,  // Дата уведомления
    180, // Наименование двигателя
    65,  // Номер двигателя
    100,  // Дата сборки
    70,  // Пробег, км
    150, // Предприятие
    100,  // Завод
    100, // Выезд/ФИО
    120, // Неисправный элемент
    130, // Причина неисправности
    60,  // Особое мнение
    450, // Примечание
    70,  //Прибыл в ремонт
    //70,  // Наличие паспорта
    70   // Убыл из ремонта
    //50   //Срок
  ]);

  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TReclamationSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TReclamationSheet.Zoom(const APercents: Integer);
begin
  FWriter.SetZoom(APercents);
end;

procedure TReclamationSheet.DrawLine(const AIndex: Integer; const ASelected: Boolean);
var
  R: Integer;
  S: String;
begin
  R:= 2 + AIndex;


  FWriter.SetAlignment(haCenter, vaTop);
  FWriter.SetBackground(ReasonColors[AIndex]);
  FWriter.WriteNumber(R, 1, AIndex+1, cbtOuter);


  if ASelected then
  begin
    FWriter.SetBackground(COLOR_BACKGROUND_SELECTED);
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  end
  else begin
    FWriter.SetBackgroundClear;
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  end;


  FWriter.WriteDate(R, 2, RecDates[AIndex], cbtOuter);
  FWriter.SetAlignment(haLeft, vaTop);
  FWriter.WriteText(R, 3, MotorNames[AIndex], cbtOuter);
  FWriter.SetAlignment(haCenter, vaTop);
  FWriter.WriteText(R, 4, MotorNums[AIndex], cbtOuter);
  FWriter.WriteDate(R, 5, BuildDates[AIndex], cbtOuter);
  if Mileages[AIndex]>=0 then
    FWriter.WriteNumber(R, 6, Mileages[AIndex], cbtOuter)
  else
    FWriter.WriteText(R, 6, EmptyStr, cbtOuter);
  FWriter.WriteText(R, 7, PlaceNames[AIndex], cbtOuter, True, True);
  FWriter.WriteText(R, 8, FactoryNames[AIndex], cbtOuter, True, True);
  FWriter.WriteText(R, 9, Departures[AIndex], cbtOuter, True, True);
  FWriter.WriteText(R, 10, DefectNames[AIndex], cbtOuter, True, True);
  FWriter.WriteText(R, 11, ReasonNames[AIndex], cbtOuter, True, True);

  FWriter.SetFont(FFontName, FFontSize+3, [fsBold], clBlack);
  S:= EmptyStr;
  if Opinions[AIndex]=1 then S:= CHECK_SYMBOL;
  FWriter.WriteText(R, 12, S, cbtOuter);

  FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
  FWriter.SetAlignment(haLeft, vaTop);
  FWriter.WriteText(R, 13, RecNotes[AIndex], cbtOuter, True, True);

  FWriter.SetAlignment(haCenter, vaTop);
  if ArrivalDates[AIndex]=0 then
    FWriter.WriteText(R, 14, EmptyStr, cbtOuter, True, True)
  else
    FWriter.WriteDate(R, 14, ArrivalDates[AIndex], cbtOuter);

  //S:= EmptyStr;
  //if Passports[AIndex]=1 then S:= CHECK_SYMBOL;
  //FWriter.WriteText(R, 15, S, cbtOuter);

  if SendingDates[AIndex]=0 then
    FWriter.WriteText(R, 15, EmptyStr, cbtOuter, True, True)
  else
    FWriter.WriteDate(R, 15, SendingDates[AIndex], cbtOuter);

  //if ArrivalDates[AIndex]=0 then
  //  FWriter.WriteText(R, 17, EmptyStr, cbtOuter, True, True)
  //else begin
  //  if SendingDates[AIndex]>0 then
  //    D:= SendingDates[AIndex]
  //  else
  //    D:= Date;
  //  NumDays:= DaysBetweenDates(ArrivalDates[AIndex], D) + 1;
  //  FWriter.WriteNumber(R, 17, NumDays, cbtOuter);
  //end;

end;

procedure TReclamationSheet.Draw(const ARecDates, ABuildDates,
                   AArrivalDates, ASendingDates: TDateVector;
                   const AMileages, AOpinions, AReasonColors{, APassports}: TIntVector;
                   const APlaceNames, AFactoryNames, ADepartures,
                         ADefectNames, AReasonNames, ARecNotes,
                         AMotorNames, AMotorNums: TStrVector);
var
  i: Integer;
begin
  FGrid.Clear;
  FWriter.BeginEdit;

  RecDates:= ARecDates;
  BuildDates:= ABuildDates;
  Mileages:= AMileages;
  Opinions:= AOpinions;
  PlaceNames:= APlaceNames;
  FactoryNames:= AFactoryNames;
  Departures:= ADepartures;
  DefectNames:= ADefectNames;
  ReasonNames:= AReasonNames;
  RecNotes:= ARecNotes;
  MotorNames:= AMotorNames;
  MotorNums:= AMotorNums;
  ReasonColors:= AReasonColors;
  ArrivalDates:= AArrivalDates;
  SendingDates:= ASendingDates;
  //Passports:= APassports;

  DrawTitle;

  if not VIsNil(RecDates) then
  begin
    for i:= 0 to High(RecDates) do
      DrawLine(i, False);
    FWriter.SetBackgroundClear;
    for i:= 1 to 17 do
      FWriter.WriteText(2+Length(RecDates), i, EmptyStr, cbtTop);

    FWriter.SetFrozenRows(1);
    FWriter.SetRepeatedRows(1,1);
  end;


  FWriter.EndEdit;

end;

{ TCargoSheet }

constructor TCargoSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  FGrid:= AGrid;
  FGrid.DefaultRowHeight:= ROW_HEIGHT_DEFAULT;
  FGrid.MouseWheelOption:= mwGrid;
  FGrid.ShowGridLines:= False;
  FGrid.ShowHeaders:= False;
  FGrid.SelectionPen.Style:= psClear;

  //LoadFontFromControl(FGrid, FFontName, FFontSize);
  FFontName:= SHEET_FONT_NAME;
  FFontSize:= SHEET_FONT_SIZE;

  ColWidths:= VCreateInt([
    300, // Наименование
    150, // Количество      номера
    150, //                 номера
    100  // Номер партии
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TCargoSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TCargoSheet.Draw(const ASendDate: TDate; const AReceiverName: String;
                   const AMotorNames: TStrVector; const AMotorsCount: TIntVector;
                   const AMotorNums, ASeries: TStrMatrix);
var
  R, RR, i, j: Integer;
  S: String;
const
  FONT_SIZE_DELTA = 1;
begin
  FGrid.Clear;
  FWriter.BeginEdit;


  //Заголовок
  R:= 1;
  FWriter.SetFont(FFontName, FFontSize+3, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  S:= FormatDateTime('dd.mm.yyyy', ASendDate) +
      ' Отгрузка в ' + AReceiverName;
  FWriter.WriteText(R, 1, R, 4, S, cbtNone, True, True);

  //наименование + количество
  R:= R + 2;
  FWriter.SetFont(FFontName, FFontSize+FONT_SIZE_DELTA, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, 'Наименование двигателей', cbtOuter, True, True);
  FWriter.WriteText(R, 2, 'Количество', cbtOuter, True, True);

  FWriter.SetFont(FFontName, FFontSize+FONT_SIZE_DELTA, [{fsBold}], clBlack);
  for i:= 0 to High(AMotorNames) do
  begin
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.WriteText(R, 1, AMotorNames[i], cbtOuter, True, True);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteNumber(R, 2, AMotorsCount[i], cbtOuter);
  end;
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize+FONT_SIZE_DELTA, [fsBold], clBlack);
  FWriter.SetAlignment(haLeft, vaCenter);
  FWriter.WriteText(R, 1, 'ИТОГО', cbtOuter, True, True);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteNumber(R, 2, VSum(AMotorsCount), cbtOuter);

  //раскладка по номерам + партиям
  R:= R + 2;
  FWriter.SetFont(FFontName, FFontSize+FONT_SIZE_DELTA, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, 'Наименование двигателей', cbtOuter, True, True);
  FWriter.WriteText(R, 2, R, 3, 'Номера двигателей', cbtOuter, True, True);
  FWriter.WriteText(R, 4, 'Партия', cbtOuter, True, True);

  FWriter.SetFont(FFontName, FFontSize+FONT_SIZE_DELTA, [{fsBold}], clBlack);
  for i:= 0 to High(AMotorNames) do
  begin
    R:= R + 1;
    for j:= 0 to High(AMotorNums[i]) do
    begin
      RR:= R + j;
      //FWriter.SetAlignment(haCenter, vaTop);
      FWriter.SetAlignment(haLeft, vaTop);
      FWriter.WriteText(RR, 2, RR, 3,  AMotorNums[i][j], cbtOuter, True, True);
      FWriter.SetAlignment(haCenter, vaTop);
      FWriter.WriteText(RR, 4, ASeries[i][j], cbtOuter, True, True);
    end;
    FWriter.SetAlignment(haLeft, vaTop);
    FWriter.WriteText(R, 1, R+High(AMotorNums[i]), 1, AMotorNames[i], cbtOuter, True, True);
    R:= R + High(AMotorNums[i]);
  end;

  R:= R + 1;
  FWriter.WriteText(R, 1, R, 4, EmptyStr, cbtTop);

  FWriter.EndEdit;
end;

end.

