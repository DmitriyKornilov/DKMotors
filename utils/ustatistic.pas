unit UStatistic;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics,
  //DK packages utils
  DK_Vector, DK_Matrix, DK_SheetTypes, DK_Const, DK_Math;

const
  //кол-во стобцов
  COL_COUNT_BAR_FULL         = 150;                       // на которых могут быть bars
  COL_COUNT_BAR_USED_PERIOD  = COL_COUNT_BAR_FULL div 2;  // на которых фактически будут bars - для отчета за период
  COL_COUNT_BAR_USED_COMPAR  = COL_COUNT_BAR_FULL div 4;  // на которых фактически будут bars - для сопоставления по годам
  COL_COUNT_BAR_VALUE_PERIOD = COL_COUNT_BAR_FULL div 5;  // для надписи значения bar - для отчета за период
  COL_COUNT_BAR_VALUE_COMPAR = COL_COUNT_BAR_FULL div 10; // для надписи значения bar - для сопоставления по годам
  COL_COUNT_LEGEND_COLOR = 2; //для цветного прямоугольника легенды
  COL_COUNT_TOTAL = 1{ParamNamesColumn} + COL_COUNT_BAR_FULL; // всего

  COL_COUNT_CELL_FULL = COL_COUNT_BAR_FULL div 5;  // для целой ячейки (годы=нет, %=нет)
  COL_COUNT_CELL_HALF = COL_COUNT_CELL_FULL div 2; // для половины целой ячейки (годы=нет, %=да)
  COL_COUNT_CELL_YEAR = COL_COUNT_CELL_FULL div 5; // для ячейки (годы=да, %=нет)
  COL_COUNT_CELL_PART = COL_COUNT_CELL_YEAR div 2; // для ячейки (годы=да, %=да)

  //минимально возможное кол-во столбцов отчета-сопоставления, остальные справа делаем colwidth=0
  COL_COUNT_REPORT_MIN = COL_COUNT_CELL_YEAR*3{years}*5{reasons};

  COL_WIDTH_BAR_PERIOD = 6;
  COL_WIDTH_BAR_COMPAR = 18;
  COL_WIDTH_NAMES = 200; //defaul

  ROW_EMPTY_HEIGH = 5;
  PERCENT_FRAC_DIGITS = 1;

  //5 years, 5 reasons
  GRAPH_COLORS_DEFAULT: array [0..4] of TColor = (
    $00BD814F, //blue         // [0] : Не расследовано
    $004696F7, //orange       // [1] : Некачественные комплектующие
    $0059BB9B, //green        // [2] : Дефект сборки / изготовления
    $00A26480, //violet       // [3] : Нарушение условий эксплуатации
    $004D50C0  //red          // [4] : Электродвигатель исправен
  );

type

  { TStatSheet }

  TStatSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  private
    FReasonNeeds: TBoolVector;     //флаги использования критериев (причин) неисправности
    FReasonNames: TStrVector;      //список критериев (причин) неисправности
    FParamColName: String;         //наименование столбца параметров
    FParamNeeds: TBoolVector;      //флаги использования параметров
    FParamNames: TStrVector;       //список наименований параметров
    FCounts: TIntMatrix3D;         //количество рекламаций [period_index, reason_index, param_index]
    FSumCounts: TIntMatrix;        //сумма рекламаций по параметрам [period_index, param_index]
    FTotalCounts: TIntVector;      //общая сумма рекламаций по периодам [period_index]
    FReasonCounts: TIntMatrix;     //общая сумма рекламаций по критериям (причинам) [period_index, reason_index]

    procedure HorizBarDraw(const ARow: Integer;
                           const AColor: TColor;
                           const AValue,      //кол-во рекламаций для этого столбца
                                 AMaxValue,   //максимальное кол-во рекламаций в этой гистограмме - занимает все ABarUsedColCount
                                 ATotalValue, //общее кол-во рекламаций, от которого считаем %
                                 ABarUsedColCount,  //кол-во столбцов из COL_COUNT_BAR_FULL, на которых будет отрисовка самого bar
                                 ABarValueColCount: Integer; //кол-во столбцов из COL_COUNT_BAR_FULL, на которых будет отрисовка значения bar
                           const ANeedPercent: Boolean);

    procedure SimpleHistogramRowDraw(var ARow: Integer;
                           const AName: String;
                           const AValue, AMaxValue, ATotalCount,
                                 ABarUsedColCount, ABarValueColCount: Integer;
                           const ANeedPercent: Boolean);
    procedure SimpleHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntVector;
                                    const AMaxValue: Integer;
                                    const ANeeds: TBoolVector;
                                    const ATotalCount, ABarUsedColCount, ABarValueColCount: Integer;
                                    const ANeedPercent: Boolean;
                                    const ANeedSort: Boolean);

    procedure ComparisonHistogramRowDraw(var ARow: Integer;
                                    const AName: String;
                                    const AValues: TIntVector;
                                    const AMaxValue: Integer;
                                    const ATotalCounts: TIntVector;
                                    const ABarUsedColCount, ABarValueColCount: Integer;
                                    const ANeedPercent: Boolean);
    procedure ComparisonHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntMatrix;
                                    const AMaxValue: Integer;
                                    const ANeeds: TBoolVector;
                                    const ATotalCounts: TIntVector;
                                    const ABarUsedColCount, ABarValueColCount: Integer;
                                    const ACategories: TStrVector;
                                    const ANeedPercent: Boolean);

    procedure PeriodSumTableDraw(var ARow: Integer;
                            const AValueColumnTitle: String;
                            const ANames: TStrVector;
                            const AValues: TIntVector;
                            const ANeeds: TBoolVector;
                            const ATotalCount: Integer;
                            const ANeedPercent: Boolean;
                            const ANeedResume: Boolean);
    procedure ComparisonSumTableDraw(var ARow: Integer;
                            const AYear: Integer;
                            const ANames: TStrVector;
                            const AValues: TIntMatrix;
                            const ANeeds: TBoolVector;
                            const ATotalCounts: TIntVector;
                            const ANeedPercent: Boolean;
                            const ANeedResume: Boolean);


    procedure PeriodReasonTableDraw(var ARow: Integer;
                               const AParamSumCountForPercents: TIntVector;
                               const ATotalSumCountForPercent: Integer;
                               const ANeedPercent: Boolean;
                               const ANeedResume: Boolean);
    procedure ComparisonReasonTableDraw(var ARow: Integer;
                               const AYear: Integer;
                               const AParamSumCountForPercents: TIntMatrix;
                               const ATotalSumCountForPercent: TIntVector;
                               const ANeedPercent: Boolean;
                               const ANeedResume: Boolean);

    procedure ReportTitleDraw(var ARow: Integer; const AMotorNamesStr, APeriodStr: String);

    procedure PrepareData(const AParamColName: String;
                           const AReasonNeeds: TBoolVector;
                           const AReasonNames: TStrVector;
                           const AParamNeeds: TBoolVector;
                           const AParamNames: TStrVector;
                           const AClaimCounts: TIntMatrix3D;
                           const ASumType: Integer);
  public
    procedure PeriodDraw(const AParamColName, //наименование столбца параметров (напр., "Наименование электродвигателя")
                         APartTitle,          //наименование параметров (дат. падеж, мн. ч.) (напр., "наименованиям электродвигателей")
                         APartTitle2,         //наименование параметров (дат. падеж, ед. ч.) (напр., "наименованию электродвигателя")
                         AMotorNamesStr,      //строка с перечислением используемых типов(наименований) двигателей (напр., "7АЖ225М6У2 IМ2001, АЭВ71А2У2 IМ2003")
                         APeriodStr: String;  //строка с описанием периода, за который выводится отчет (напр., "с 01.01.2020 по 31.12.2020")
                   const AReasonNeeds: TBoolVector;  //флаги использования критериев (причин) неисправностей
                   const AReasonNames: TStrVector;   //вектор значений критериев (причин) неисправностей
                   const AParamNeeds: TBoolVector;   //флаги использования параметров
                   const AParamNames: TStrVector;    //вектор значений параметров
                   const AClaimCounts: TIntMatrix3D; //кол-во рекламаций
                   const ADataNeed: TBoolVector;     //включать данные
                                                     // [0] - общее количество по виду статистики
                                                     // [1] - общее количество по критериям неисправности
                                                     // [2] - количество по виду статистики и критериям неисправности
                   const ASumType: Integer;          // Подсчет суммы рекламаций
                                                     // =0 - по всем критериям
                                                     // =1 - только по включенным в отчёт критериям
                   const AHistogramNeed,             //выводить гистограммы
                         APercentNeed,               //выводить % от кол-ва
                         AAccumNeed,                 //общее кол-во по основному параметру считать с накоплением
                         ASortNeed: Boolean          //сортировать данные в гистограмме общего кол-ва по основному параметру
                   );

    procedure ComparisonDraw(const AYear: Integer;
                         const AParamColName, //наименование столбца параметров (напр., "Наименование электродвигателя")
                         APartTitle,          //наименование параметров (дат. падеж, мн. ч.) (напр., "наименованиям электродвигателей")
                         APartTitle2,         //наименование параметров (дат. падеж, ед. ч.) (напр., "наименованию электродвигателя")
                         AMotorNamesStr,      //строка с перечислением используемых типов(наименований) двигателей (напр., "7АЖ225М6У2 IМ2001, АЭВ71А2У2 IМ2003")
                         APeriodStr: String;  //строка с описанием периода, за который выводится отчет (напр., "с 01.01.2020 по 31.12.2020")
                   const AReasonNeeds: TBoolVector;  //флаги использования критериев (причин) неисправностей
                   const AReasonNames: TStrVector;   //вектор значений критериев (причин) неисправностей
                   const AParamNeeds: TBoolVector;   //флаги использования параметров
                   const AParamNames: TStrVector;    //вектор значений параметров
                   const AClaimCounts: TIntMatrix3D; //кол-во рекламаций
                   const ADataNeed: TBoolVector;     //включать данные
                                                     // [0] - общее количество по виду статистики
                                                     // [1] - общее количество по критериям неисправности
                                                     // [2] - количество по виду статистики и критериям неисправности
                   const ASumType: Integer;          // Подсчет суммы рекламаций
                                                     // =0 - по всем критериям
                                                     // =1 - только по включенным в отчёт критериям
                   const AHistogramNeed,             //выводить гистограммы
                         APercentNeed: Boolean       //выводить % от кол-ва
                   );
  end;



  //ClaimCountForReason - вектор данных (кол-во рекламаций), соответствующий ReasonNames
  // [0] : Не расследовано
  // [1] : Некачественные комплектующие
  // [2] : Дефект сборки / изготовления
  // [3] : Нарушение условий эксплуатации
  // [4] : Электродвигатель исправен
  // APeriodIndex - индекс периода
  //    0 - произвольный период (в этом случае это единственный индекс)
  //        или указанный период года, который нужно сопоставить с предыдущими
  //    1 - данные за указанный период 1-ого предшествующего года
  //    2 - данные за указанный период 2-ого предшествующего года
  //    3 - данные за указанный период 3-его предшествующего года
  //    4 - данные за указанный период 4-ого предшествующего года (max значение индекса)
  // AParamIndex - индекс параметра, соответствующий ParamNames
  //    - наименования двигателей
  //    - неисправные элементы двигателя
  //    - депо обнаружения неисправности
  //    - месяцы
  //    - пробеги локомотива
  // AClaimCounts - количество рекламаций
  function ClaimCountForReason(const APeriodIndex, AParamIndex: Integer;
                               const AClaimCounts: TIntMatrix3D): TIntVector;

  //ClaimCountSum - сумма рекламаций
  // - Index1 - индекс периода
  //      0 - данные за произвольный период (в этом случае это единственный индекс)
  //          или указанный период года, который нужно сопоставить с предыдущими
  //      1 - данные за указанный период 1-ого предшествующего года
  //      2 - данные за указанный период 2-ого предшествующего года
  //      3 - данные за указанный период 3-его предшествующего года
  //      4 - данные за указанный период 4-ого предшествующего года (max значение индекса)
  // - Index2 - индекс параметра, соответствующий ParamNames
  // AReasonNeeds - флаги используемых критериев (причин) неисправности
  function ClaimCountSum(const AClaimCounts: TIntMatrix3D;
                         const AReasonNeeds: TBoolVector): TIntMatrix;

  //ClaimReasonSum - сумма рекламаций
  // - Index1 - индекс периода
  //      0 - данные за произвольный период (в этом случае это единственный индекс)
  //          или указанный период года, который нужно сопоставить с предыдущими
  //      1 - данные за указанный период 1-ого предшествующего года
  //      2 - данные за указанный период 2-ого предшествующего года
  //      3 - данные за указанный период 3-его предшествующего года
  //      4 - данные за указанный период 4-ого предшествующего года (max значение индекса)
  // - Index2 - индекс критерия (причины) неисправности, соответствующий ReasonNames
  function ClaimReasonSum(const AClaimCounts: TIntMatrix3D): TIntMatrix;

  //ClaimTotalSum - общая сумма рекламаций
  // - Index1 - индекс периода
  //      0 - данные за произвольный период (в этом случае это единственный индекс)
  //          или указанный период года, который нужно сопоставить с предыдущими
  //      1 - данные за указанный период 1-ого предшествующего года
  //      2 - данные за указанный период 2-ого предшествующего года
  //      3 - данные за указанный период 3-его предшествующего года
  //      4 - данные за указанный период 4-ого предшествующего года (max значение индекса)
  // AClaimSumCounts - сумма рекламаций (результат ClaimCountSum)
  function ClaimTotalSum(const AClaimSumCounts: TIntMatrix): TIntVector;

implementation

function ClaimCountForReason(const APeriodIndex, AParamIndex: Integer;
                             const AClaimCounts: TIntMatrix3D): TIntVector;
begin
  Result:= nil;
  if High(AClaimCounts)<APeriodIndex then Exit;
  if High(AClaimCounts[APeriodIndex, 0])<AParamIndex then Exit;
  Result:= MRowGet(AClaimCounts[APeriodIndex], AParamIndex);
end;

function ClaimCountSum(const AClaimCounts: TIntMatrix3D;
                       const AReasonNeeds: TBoolVector): TIntMatrix;
var
  i, j: Integer;
begin
  Result:= nil;
  if MIsNil(AClaimCounts) then Exit;
  MDim(Result, Length(AClaimCounts), Length(AClaimCounts[0, 0]));
  for i:= 0 to High(AClaimCounts) do
    for j:= 0 to High(AClaimCounts[i, 0]) do
      Result[i, j]:= VSum(VCut(ClaimCountForReason(i, j, AClaimCounts), AReasonNeeds));
end;

function ClaimReasonSum(const AClaimCounts: TIntMatrix3D): TIntMatrix;
var
  i, j: Integer;
begin
  Result:= nil;
  if MIsNil(AClaimCounts) then Exit;
  MDim(Result, Length(AClaimCounts), Length(AClaimCounts[0]));
  for i:= 0 to High(AClaimCounts) do
    for j:= 0 to High(AClaimCounts[i]) do
      Result[i, j]:= VSum(AClaimCounts[i, j]);
end;

function ClaimTotalSum(const AClaimSumCounts: TIntMatrix): TIntVector;
var
  i: Integer;
begin
  Result:= nil;
  if MIsNil(AClaimSumCounts) then Exit;
  VDim(Result, Length(AClaimSumCounts));
  for i:= 0 to High(AClaimSumCounts) do
    Result[i]:= VSum(AClaimSumCounts[i]);
end;

{ TStatSheet }

function TStatSheet.SetWidths: TIntVector;
begin
  VDim(Result{%H-}, COL_COUNT_TOTAL+1, COL_WIDTH_BAR_PERIOD);
  Result[0]:= COL_WIDTH_NAMES;
  Result[High(Result)]:= 1;
end;

procedure TStatSheet.HorizBarDraw(const ARow: Integer;
                                  const AColor: TColor;
                                  const AValue, AMaxValue, ATotalValue,
                                        ABarUsedColCount, ABarValueColCount: Integer;
                                  const ANeedPercent: Boolean);
var
  R, C1, C2, ColorColCount: Integer;
  S: String;
  Percent: Double;
begin
  R:= ARow;

  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);

  ColorColCount:= PartRound(ABarUsedColCount*AValue, AMaxValue);
  if (ColorColCount=0) and (AValue>0) then
    ColorColCount:= 1;
  C1:= 1{ParamNames} + 1;
  C2:= C1 + ColorColCount - 1;

  if ColorColCount>0 then
  begin
    Writer.SetBackground(AColor);
    Writer.WriteText(R, C1, R, C2, EmptyStr);
    Writer.SetBackgroundClear;
  end;

  if AValue=0 then
  begin
    if not ANeedPercent then
      Writer.WriteNumber(R, C2+1, R, C2+ABarValueColCount, AValue, cbtLeft)
    else begin
      S:= '0  (0%  от  ' + IntToStr(ATotalValue) + ')';
      Writer.WriteText(R, C2+1, R, C2+ABarValueColCount, S, cbtLeft)
    end;
  end
  else begin
    if not ANeedPercent then
      Writer.WriteNumber(R, C2+1, R, C2+ABarValueColCount, AValue, cbtNone)
    else begin
      Percent:= Part(100*AValue, ATotalValue);
      S:= IntToStr(AValue) + '  (' + Format('%.1f', [Percent]) + '%  от  ' + IntToStr(ATotalValue) + ')';
      Writer.WriteText(R, C2+1, R, C2+ABarValueColCount, S, cbtNone);
    end;
  end;
end;

procedure TStatSheet.SimpleHistogramRowDraw(var ARow: Integer;
                                  const AName: String;
                                  const AValue, AMaxValue, ATotalCount,
                                        ABarUsedColCount, ABarValueColCount: Integer;
                                  const ANeedPercent: Boolean);
var
  R: Integer;
begin
  R:= ARow;
  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);

  //EmptyRow
  Writer.WriteText(R, 1, EmptyStr, cbtRight);
  Writer.RowHeight[R]:= ROW_EMPTY_HEIGH;

  //data row
  R:= R + 1;
  Writer.WriteText(R, 1, AName, cbtRight);
  HorizBarDraw(R, GRAPH_COLORS_DEFAULT[0], AValue, AMaxValue, ATotalCount,
               ABarUsedColCount, ABarValueColCount, ANeedPercent);

  //EmptyRow
  R:= R + 1;
  Writer.WriteText(R, 1, EmptyStr, cbtRight);
  Writer.RowHeight[R]:= ROW_EMPTY_HEIGH;

  ARow:= R + 1;
end;

procedure TStatSheet.SimpleHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntVector;
                                    const AMaxValue: Integer;
                                    const ANeeds: TBoolVector;
                                    const ATotalCount, ABarUsedColCount, ABarValueColCount: Integer;
                                    const ANeedPercent: Boolean;
                                    const ANeedSort: Boolean);
var
  R, i: Integer;
  Names: TStrVector;
  Values, Indexes: TIntVector;
  Needs: TBoolVector;
begin
  R:= ARow;

  if ANeedSort then
  begin
    VSort(AValues, Indexes, True{desc});
    Names:= VReplace(ANames, Indexes);
    Values:= VReplace(AValues, Indexes);
    Needs:= VReplace(ANeeds, Indexes);
  end
  else begin
    Names:= ANames;
    Values:= AValues;
    Needs:= ANeeds;
  end;

  for i:= 0 to High(Names) do
    if Needs[i] then
      SimpleHistogramRowDraw(R, Names[i], Values[i], AMaxValue, ATotalCount,
                             ABarUsedColCount, ABarValueColCount, ANeedPercent);

  ARow:= R;
end;

procedure TStatSheet.ComparisonHistogramRowDraw(var ARow: Integer;
                                    const AName: String;
                                    const AValues: TIntVector;
                                    const AMaxValue: Integer;
                                    const ATotalCounts: TIntVector;
                                    const ABarUsedColCount, ABarValueColCount: Integer;
                                    const ANeedPercent: Boolean);
var
  R, RR, i: Integer;
begin
  R:= ARow;
  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);

  //EmptyRow
  Writer.WriteText(R, 1, EmptyStr, cbtRight);
  Writer.RowHeight[R]:= ROW_EMPTY_HEIGH;

  //data row
  R:= R + 1;
  Writer.WriteText(R, 1, R+2*High(AValues), 1, AName, cbtRight);
  RR:= R;
  R:= R - 1;
  for i:= High(AValues) downto 0 do
  begin
    R:= R + 1;
    HorizBarDraw(R, GRAPH_COLORS_DEFAULT[High(AValues)-i], AValues[i], AMaxValue,
                 ATotalCounts[i], ABarUsedColCount, ABarValueColCount, ANeedPercent);
    if i>0 then//EmptyRow
    begin
      R:= R + 1;
      Writer.WriteText(R, 1, EmptyStr, cbtRight);
      Writer.RowHeight[R]:= (ROW_EMPTY_HEIGH div 2) + 1;
    end;
  end;
  if R<RR then R:= RR;

  //EmptyRow
  R:= R + 1;
  Writer.WriteText(R, 1, EmptyStr, cbtRight);
  Writer.RowHeight[R]:= ROW_EMPTY_HEIGH;

  R:= R + 1;
  Writer.WriteText(R, 1, EmptyStr, cbtNone);
  Writer.RowHeight[R]:= ROW_EMPTY_HEIGH;

  ARow:= R + 1;
end;

procedure TStatSheet.ComparisonHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntMatrix;
                                    const AMaxValue: Integer;
                                    const ANeeds: TBoolVector;
                                    const ATotalCounts: TIntVector;
                                    const ABarUsedColCount, ABarValueColCount: Integer;
                                    const ACategories: TStrVector;
                                    const ANeedPercent: Boolean);
var
  R, RR, i, C: Integer;
begin
  R:= ARow;
  RR:= ARow;

  for i:= 0 to High(ANames) do
    if ANeeds[i] then
      ComparisonHistogramRowDraw(R, ANames[i], MRowGet(AValues, i), AMaxValue,
                      ATotalCounts, ABarUsedColCount, ABarValueColCount, ANeedPercent);

  ARow:= R;

  C:= 3 + ABarUsedColCount + ABarValueColCount;
  Writer.SetAlignment(haLeft, vaCenter);
  for i:= 0 to High(ACategories) do
  begin
    RR:= RR + 1;
    Writer.SetBackground(GRAPH_COLORS_DEFAULT[i]);
    Writer.WriteText(RR, C, RR, C+COL_COUNT_LEGEND_COLOR, EmptyStr);
    Writer.SetBackgroundDefault;
    Writer.WriteText(RR, C+COL_COUNT_LEGEND_COLOR+2, RR, Writer.ColCount, ACategories[i]);
    //EmptyRow
    RR:= RR + 1;
  end;
end;

procedure TStatSheet.PeriodSumTableDraw(var ARow: Integer;
                                  const AValueColumnTitle: String;
                                  const ANames: TStrVector;
                                  const AValues: TIntVector;
                                  const ANeeds: TBoolVector;
                                  const ATotalCount: Integer;
                                  const ANeedPercent: Boolean;
                                  const ANeedResume: Boolean);
var
  R, C1, C2, i: Integer;
  Percent: Double;
begin
  R:= ARow;

  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);

  //заголовок
  C1:= 2;
  C2:= C1 + COL_COUNT_CELL_FULL - 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, FParamColName, cbtOuter, True, True);
  Writer.WriteText(R, C1, R, C2, AValueColumnTitle, cbtOuter, True, True);
  Writer.DrawBorders(R, C2+1, cbtLeft);

  //данные
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  for i:= 0 to High(ANames) do
  begin
    if not ANeeds[i] then continue;

    R:= R + 1;
    Writer.WriteText(R, 1, ANames[i], cbtOuter);
    if not ANeedPercent then
      Writer.WriteNumber(R, C1, R, C2, AValues[i], cbtOuter)
    else begin
      C1:= 2;
      C2:= C1 + COL_COUNT_CELL_HALF - 1;
      Writer.WriteNumber(R, C1, R, C2, AValues[i], cbtOuter);
      C1:= C2 + 1;
      C2:= C1 + COL_COUNT_CELL_HALF - 1;
      Percent:= Part(AValues[i], ATotalCount);
      Writer.WriteNumber(R, C1, R, C2, Percent, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  //итого
  if ANeedResume then
  begin
    R:= R + 1;
    C1:= 2;
    C2:= C1 + COL_COUNT_CELL_FULL - 1;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.WriteText(R, 1, 'ИТОГО', cbtOuter);
    Writer.WriteNumber(R, C1, R, C2, ATotalCount, cbtOuter);
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  ARow:= R;
end;

procedure TStatSheet.ComparisonSumTableDraw(var ARow: Integer;
                            const AYear: Integer;
                            const ANames: TStrVector;
                            const AValues: TIntMatrix;
                            const ANeeds: TBoolVector;
                            const ATotalCounts: TIntVector;
                            const ANeedPercent: Boolean;
                            const ANeedResume: Boolean);
var
  R, C1, C2, i, j, N: Integer;
  Percent: Double;
begin
  R:= ARow;

  N:= Length(ATotalCounts); //years count

  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);

  //заголовок
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, R+1, 1, FParamColName, cbtOuter, True, True);
  Writer.WriteText(R, 2, R, 1+COL_COUNT_CELL_YEAR*N, 'Количество рекламаций за период', cbtOuter, True, True);
  R:= R + 1;
  C2:= 1;
  for i:= N-1 downto 0 do
  begin
    C1:= C2 + 1;
    C2:= C1 + COL_COUNT_CELL_YEAR - 1;
    Writer.WriteNumber(R, C1, R, C2, AYear-i, cbtOuter);
  end;
  Writer.DrawBorders(R-1, C2+1, cbtLeft);
  Writer.DrawBorders(R, C2+1, cbtLeft);

  //данные
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  for j:= 0 to High(ANames) do
  begin
    if not ANeeds[j] then continue;

    R:= R + 1;
    Writer.WriteText(R, 1, ANames[j], cbtOuter);

    C2:= 1;
    for i:= N-1 downto 0 do
    begin
      C1:= C2 + 1;
      if not ANeedPercent then
      begin
        C2:= C1 + COL_COUNT_CELL_YEAR - 1;
        Writer.WriteNumber(R, C1, R, C2, AValues[i, j], cbtOuter);
      end
      else begin
        C2:= C1 + COL_COUNT_CELL_PART - 1;
        Writer.WriteNumber(R, C1, R, C2, AValues[i, j], cbtOuter);
        C1:= C2 + 1;
        C2:= C1 + COL_COUNT_CELL_PART - 1;
        Percent:= Part(AValues[i, j], ATotalCounts[i]);
        Writer.WriteNumber(R, C1, R, C2, Percent, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
      end;
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  //итого
  if ANeedResume then
  begin
    R:= R + 1;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.WriteText(R, 1, 'ИТОГО', cbtOuter);

    C2:= 1;
    for i:= N-1 downto 0 do
    begin
      C1:= C2 + 1;
      C2:= C1 + COL_COUNT_CELL_YEAR - 1;
      Writer.WriteNumber(R, C1, R, C2, ATotalCounts[i], cbtOuter);
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  ARow:= R;
end;

procedure TStatSheet.PeriodReasonTableDraw(var ARow: Integer;
                                     const AParamSumCountForPercents: TIntVector;
                                     const ATotalSumCountForPercent: Integer;
                                     const ANeedPercent: Boolean;
                                     const ANeedResume: Boolean);
var
  R, C1, C2, i, j, Value: Integer;
  Percent: Double;
begin
  R:= ARow;

  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);

  //заголовок
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, FParamColName, cbtOuter, True, True);
  C2:= 1;
  for i:= 0 to High(FReasonNames) do
  begin
    if not FReasonNeeds[i] then continue;
    C1:= C2 + 1;
    C2:= C1 + COL_COUNT_CELL_FULL - 1;
    Writer.WriteText(R, C1, R, C2, FReasonNames[i], cbtOuter, True, True);
  end;
  Writer.DrawBorders(R, C2+1, cbtLeft);

  //данные
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  for i:= 0 to High(FParamNames) do
  begin
    if not FParamNeeds[i] then continue;

    R:= R + 1;
    Writer.WriteText(R, 1, FParamNames[i], cbtOuter);
    C2:= 1;
    for j:= 0 to High(FReasonNames) do
    begin
      if not FReasonNeeds[j] then continue;
      C1:= C2 + 1;
      Value:= FCounts[0{period_index}, j{reason_index}, i{param_index}];
      if not ANeedPercent then
      begin
        C2:= C1 + COL_COUNT_CELL_FULL - 1;
        Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
      end
      else begin
        C2:= C1 + COL_COUNT_CELL_HALF - 1;
        Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
        C1:= C2 + 1;
        C2:= C1 + COL_COUNT_CELL_HALF - 1;
        Percent:= Part(Value, AParamSumCountForPercents[i]);
        Writer.WriteNumber(R, C1, R, C2, Percent, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
      end;
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  //итого
  if ANeedResume then
  begin
    R:= R + 1;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.WriteText(R, 1, 'ИТОГО', cbtOuter);
    C2:= 1;
    for i:= 0 to High(FReasonNames) do
    begin
      if not FReasonNeeds[i] then continue;
      C1:= C2 + 1;
      Value:= FReasonCounts[0{period_index}, i{reason_index}];
      if not ANeedPercent then
      begin
        C2:= C1 + COL_COUNT_CELL_FULL - 1;
        Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
      end
      else begin
        C2:= C1 + COL_COUNT_CELL_HALF - 1;
        Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
        C1:= C2 + 1;
        C2:= C1 + COL_COUNT_CELL_HALF - 1;
        Percent:= Part(Value, ATotalSumCountForPercent);
        Writer.WriteNumber(R, C1, R, C2, Percent, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
      end;
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  ARow:= R;
end;

procedure TStatSheet.ComparisonReasonTableDraw(var ARow: Integer;
                               const AYear: Integer;
                               const AParamSumCountForPercents: TIntMatrix;
                               const ATotalSumCountForPercent: TIntVector;
                               const ANeedPercent: Boolean;
                               const ANeedResume: Boolean);
var
  R, C1, C2, i, j, k, N, Value: Integer;
  Percent: Double;
begin
  R:= ARow;

  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);

  N:= Length(AParamSumCountForPercents); //years count

  //заголовок
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, R+1, 1, FParamColName, cbtOuter, True, True);
  C2:= 1;
  for i:= 0 to High(FReasonNames) do
  begin
    if not FReasonNeeds[i] then continue;
    C1:= C2 + 1;
    C2:= C1 + COL_COUNT_CELL_YEAR*N - 1;
    Writer.WriteText(R, C1, R, C2, FReasonNames[i], cbtOuter, True, True);
    C2:= C1 - 1;
    for j:= N-1 downto 0 do
    begin
      C1:= C2 + 1;
      C2:= C1 + COL_COUNT_CELL_YEAR - 1;
      Writer.WriteNumber(R+1, C1, R+1, C2, AYear-j, cbtOuter);
    end;
  end;
  Writer.DrawBorders(R, C2+1, cbtLeft);
  Writer.DrawBorders(R+1, C2+1, cbtLeft);
  R:= R + 1;

  //данные
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  for i:= 0 to High(FParamNames) do
  begin
    if not FParamNeeds[i] then continue;

    R:= R + 1;
    Writer.WriteText(R, 1, FParamNames[i], cbtOuter);
    C2:= 1;
    for j:= 0 to High(FReasonNames) do
    begin
      if not FReasonNeeds[j] then continue;

      for k:= N-1 downto 0 do
      begin
        C1:= C2 + 1;
        Value:= FCounts[k{period_index}, j{reason_index}, i{param_index}];
        if not ANeedPercent then
        begin
          C2:= C1 + COL_COUNT_CELL_YEAR - 1;
          Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
        end
        else begin
          C2:= C1 + COL_COUNT_CELL_PART - 1;
          Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
          C1:= C2 + 1;
          C2:= C1 + COL_COUNT_CELL_PART - 1;
          Percent:= Part(Value, AParamSumCountForPercents[k{period_index}, i{param_index}]);
          Writer.WriteNumber(R, C1, R, C2, Percent, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
        end;
      end;
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  //итого
  if ANeedResume then
  begin
    R:= R + 1;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.WriteText(R, 1, 'ИТОГО', cbtOuter);
    C2:= 1;
    for i:= 0 to High(FReasonNames) do
    begin
      if not FReasonNeeds[i] then continue;

      for j:= N-1 downto 0 do
      begin
        C1:= C2 + 1;
        Value:= FReasonCounts[j{period_index}, i{reason_index}];
        if not ANeedPercent then
        begin
          C2:= C1 + COL_COUNT_CELL_YEAR - 1;
          Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
        end
        else begin
          C2:= C1 + COL_COUNT_CELL_PART - 1;
          Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
          C1:= C2 + 1;
          C2:= C1 + COL_COUNT_CELL_PART - 1;
          Percent:= Part(Value, ATotalSumCountForPercent[j{period_index}]);
          Writer.WriteNumber(R, C1, R, C2, Percent, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
        end;
      end;
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  ARow:= R;
end;

procedure TStatSheet.ReportTitleDraw(var ARow: Integer; const AMotorNamesStr, APeriodStr: String);
var
  R: Integer;
  S: String;
begin
  R:= ARow;
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.SetFont(Font.Name, Font.Size+4, [fsBold], clBlack);
  S:= 'Отчет по рекламационным случаям электродвигателей';
  Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount, AMotorNamesStr, cbtNone, True, True);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size+4, [fsBold], clBlack);
  S:= 'за период ' + APeriodStr;
  Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
  ARow:= R;
end;

procedure TStatSheet.PeriodDraw(const AParamColName, APartTitle{дательный падеж}, APartTitle2,
                                AMotorNamesStr, APeriodStr: String;
                          const AReasonNeeds: TBoolVector;
                          const AReasonNames: TStrVector;
                          const AParamNeeds: TBoolVector;
                          const AParamNames: TStrVector;
                          const AClaimCounts: TIntMatrix3D;
                          const ADataNeed: TBoolVector;
                          const ASumType: Integer;
                          const AHistogramNeed, APercentNeed, AAccumNeed, ASortNeed: Boolean);
var
  R, k, Order, MaxValue, TotalCount: Integer;
  S: String;
  Needs: TBoolVector;
  Values: TIntVector;
  Names: TStrVector;

  procedure CountTotalForStatisticTypeDraw;
  begin
    Names:= FParamNames;
    Needs:= FParamNeeds;
    Values:= FSumCounts[0{period_index}];
    if AAccumNeed then
      Values:= VAccum(Values);
    MaxValue:= VMax(Values);
    TotalCount:= FTotalCounts[0{period_index}];

    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);

    R:= R + 1;
    Inc(Order);
    if AAccumNeed then
      S:= 'Накопление'
    else
      S:= 'Распределение';
    S:= IntToStr(Order) +
        ') ' + S + ' общего количества рекламационных случаев по ' +
        APartTitle;
    if APercentNeed then
      S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle + ' за период)';
    Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);

    R:= R + 2;
    if AAccumNeed then
      S:= 'Накопление количества' + SYMBOL_BREAK + 'рекламаций'
    else
      S:= 'Количество рекламаций' + SYMBOL_BREAK + 'за период';
    PeriodSumTableDraw(R, S, Names, Values, Needs, TotalCount, APercentNeed, not AAccumNeed);

    R:= R + 1;

    if not AHistogramNeed then Exit;
    R:= R + 1;
    SimpleHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCount,
                        COL_COUNT_BAR_USED_PERIOD, COL_COUNT_BAR_VALUE_PERIOD,
                        APercentNeed, ASortNeed{sort});
  end;

  procedure CountTotalForReasonDraw;
  var
    i: Integer;
    ParamSumCounts: TIntVector;
  begin
    TotalCount:= FTotalCounts[0{period_index}];

    //paragraph caption
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    R:= R + 1;
    Inc(Order);
    S:= IntToStr(Order) +
        ') Распределение общего количества рекламационных случаев по критериям неисправности';
    if APercentNeed then
      S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle +  ' за период)';
    Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
    //data grid
    R:= R + 2;
    ParamSumCounts:= nil;
    if APercentNeed then
      ParamSumCounts:= VCreateInt(Length(FParamNames), TotalCount);
    PeriodReasonTableDraw(R, ParamSumCounts, TotalCount, APercentNeed, True{строка итого});

    R:= R + 1;

    if not AHistogramNeed then Exit;

    //histogram
    R:= R + 1;
    Names:= FReasonNames;
    Needs:= FReasonNeeds;
    Values:= FReasonCounts[0{period_index}];
    MaxValue:= VMax(VCut(Values, Needs));
    SimpleHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCount,
                        COL_COUNT_BAR_USED_PERIOD, COL_COUNT_BAR_VALUE_PERIOD,
                        APercentNeed, False{no sort});

    //addition caption
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    Writer.WriteText(R, 1, R, Writer.ColCount, 'В том числе:', cbtNone, True, True);

    R:= R + 1;
    for i:=0 to High(FParamNames) do
    begin
      if not FParamNeeds[i] then continue;

      //histogram caption
      R:= R + 1;
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
      S:= FParamNames[i] + ' - количество рекламаций';
      if APercentNeed then
        S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle + ' за период)';
      Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
      //histogram
      R:= R + 1;
      //Names, Needs, MaxValue, TotalCount - рассчитаны выше
      Values:= ClaimCountForReason(0{period_index}, i{param_index}, FCounts);
      SimpleHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCount,
                          COL_COUNT_BAR_USED_PERIOD, COL_COUNT_BAR_VALUE_PERIOD,
                          APercentNeed, False{no sort});
    end;
  end;

  procedure CountForStatisticTypeAndReasonDraw;
  var
    i: Integer;
    ParamSumCounts: TIntVector;
  begin
    //paragraph caption
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    R:= R + 1;
    Inc(Order);
    S:= IntToStr(Order) +
        ') Распределение количества рекламационных случаев по ' +
        APartTitle + ' и критериям неисправности';
    if APercentNeed then
      S:= S + ' (с % от суммы рекламаций по ' + APartTitle2 + ')';
    Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
    //data grid
    R:= R + 2;
    ParamSumCounts:= FSumCounts[0{period_index}];
    PeriodReasonTableDraw(R, ParamSumCounts, 0{not need}, APercentNeed, False{без итого});

    R:= R + 1;

    if not AHistogramNeed then Exit;

    for i:=0 to High(FParamNames) do
    begin
      if not FParamNeeds[i] then continue;

      TotalCount:= FSumCounts[0{period_index}, i{param_index}];

      //histogram caption
      R:= R + 1;
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
      S:= FParamNames[i] + ' - количество рекламаций';
      if APercentNeed then
        S:= S + ' (с % от суммы рекламаций ' + FParamNames[i] + ' за период)';
      Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
      //histogram
      R:= R + 1;
      Names:= FReasonNames;
      Needs:= FReasonNeeds;
      Values:= ClaimCountForReason(0{period_index}, i{param_index}, FCounts);
      MaxValue:= VMax(VCut(Values, Needs));
      SimpleHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCount,
                          COL_COUNT_BAR_USED_PERIOD, COL_COUNT_BAR_VALUE_PERIOD,
                          APercentNeed, False{no sort});
    end;
  end;

begin
  if VIsAllFalse(ADataNeed) or
     VIsAllFalse(AReasonNeeds) or
     VIsAllFalse(AParamNeeds) then Exit;

  for k:= 2 to Writer.ColCount do
    Writer.ColWidth[k]:= COL_WIDTH_BAR_PERIOD;

  PrepareData(AParamColName, AReasonNeeds, AReasonNames, AParamNeeds, AParamNames,
              AClaimCounts, ASumType);

  Order:= 0;

  //drawing
  Writer.BeginEdit;
  Writer.SetBackgroundClear;
  R:= 1;
  ReportTitleDraw(R, AMotorNamesStr, APeriodStr);
  R:= R + 1;
  if ADataNeed[0] then //[0] - общее количество по виду статистики
    CountTotalForStatisticTypeDraw;
  if ADataNeed[1] then //[1] - общее количество по критериям неисправности
    CountTotalForReasonDraw;
  if ADataNeed[2] then //[2] - количество по виду статистики и критериям неисправности
    CountForStatisticTypeAndReasonDraw;

  Writer.EndEdit;
end;

procedure TStatSheet.PrepareData(const AParamColName: String;
                                 const AReasonNeeds: TBoolVector;
                                 const AReasonNames: TStrVector;
                                 const AParamNeeds: TBoolVector;
                                 const AParamNames: TStrVector;
                                 const AClaimCounts: TIntMatrix3D;
                                 const ASumType: Integer);
var
  Needs: TBoolVector;
begin
  //source data
  FParamColName:= AParamColName;
  FReasonNeeds:= AReasonNeeds;
  FReasonNames:= AReasonNames;
  FParamNeeds:= AParamNeeds;
  FParamNames:= AParamNames;
  FCounts:= AClaimCounts;

  //data calculation
  Writer.ColWidth[1]:= MMaxWidth([FReasonNames, FParamNames], Font) + 20;
  if ASumType = 0 then // =0 - по всем критериям
    Needs:= VCreateBool(Length(FReasonNames), True)
  else if ASumType = 1 then // =1 - только по включенным в отчёт критериям
    Needs:= FReasonNeeds;
  FSumCounts:= ClaimCountSum(FCounts, Needs); //сумма рекламаций по параметрам [period_index, param_index]
  FTotalCounts:= ClaimTotalSum(FSumCounts);   //общая сумма рекламаций по периодам [period_index]
  FReasonCounts:= ClaimReasonSum(FCounts);    //общая сумма рекламаций по критериям (причинам) [period_index, reason_index]

end;

procedure TStatSheet.ComparisonDraw(const AYear: Integer;
                                    const AParamColName, APartTitle, APartTitle2,
                                          AMotorNamesStr, APeriodStr: String;
                                    const AReasonNeeds: TBoolVector;
                                    const AReasonNames: TStrVector;
                                    const AParamNeeds: TBoolVector;
                                    const AParamNames: TStrVector;
                                    const AClaimCounts: TIntMatrix3D;
                                    const ADataNeed: TBoolVector;
                                    const ASumType: Integer;
                                    const AHistogramNeed, APercentNeed: Boolean);
var
  R, k, N, Order, MaxValue: Integer;
  S: String;
  Needs: TBoolVector;
  Values: TIntMatrix;
  Names, Categories: TStrVector;
  TotalCounts: TIntVector;

  procedure CountTotalForStatisticTypeDraw;
  begin
    Names:= FParamNames;
    Needs:= FParamNeeds;
    Values:= FSumCounts;
    MaxValue:= MMax(Values);
    TotalCounts:= FTotalCounts;

    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);

    R:= R + 1;
    Inc(Order);
    S:= IntToStr(Order) +
        ') Распределение общего количества рекламационных случаев по ' +
        APartTitle;
    if APercentNeed then
      S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle + ' за период)';
         //' за период = ' + IntToStr(TotalCount) + ')';
    Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);

    R:= R + 2;
    ComparisonSumTableDraw(R, AYear, Names, Values, Needs, TotalCounts, APercentNeed, True{итого});

    R:= R + 1;

    if not AHistogramNeed then Exit;
    Categories:= VIntToStr(VStep(AYear-High(Values), AYear, 1));
    R:= R + 1;
    ComparisonHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCounts,
                            COL_COUNT_BAR_USED_COMPAR, COL_COUNT_BAR_VALUE_COMPAR,
                            Categories, APercentNeed);
  end;

  procedure CountTotalForReasonDraw;
  var
    i, j: Integer;
    ParamSumCounts: TIntMatrix;
  begin
    TotalCounts:= FTotalCounts;

    //paragraph caption
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    R:= R + 1;
    Inc(Order);
    S:= IntToStr(Order) +
        ') Распределение общего количества рекламационных случаев по критериям неисправности';
    if APercentNeed then
      S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle +  ' за период)';
    Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
    //data grid
    R:= R + 2;
    ParamSumCounts:= nil;
    if APercentNeed then
      for i:= 0 to High(TotalCounts) do
        MAppend(ParamSumCounts, VCreateInt(Length(FParamNames), TotalCounts[i]));
    ComparisonReasonTableDraw(R, AYear, ParamSumCounts, TotalCounts, APercentNeed, True{строка итого});

    R:= R + 1;

    if not AHistogramNeed then Exit;

    //histogram
    Categories:= VIntToStr(VStep(AYear-High(Values), AYear, 1));
    R:= R + 1;
    Names:= FReasonNames;
    Needs:= FReasonNeeds;
    Values:= FReasonCounts;
    MaxValue:= MMax(MCut(Values, Needs));
    ComparisonHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCounts,
                            COL_COUNT_BAR_USED_COMPAR, COL_COUNT_BAR_VALUE_COMPAR,
                            Categories, APercentNeed);

    //addition caption
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    Writer.WriteText(R, 1, R, Writer.ColCount, 'В том числе:', cbtNone, True, True);

    R:= R + 1;
    for i:=0 to High(FParamNames) do
    begin
      if not FParamNeeds[i] then continue;

      //histogram caption
      R:= R + 1;
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
      S:= FParamNames[i] + ' - количество рекламаций';
      if APercentNeed then
        S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle + ' за период)';
      Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
      //histogram
      R:= R + 1;
      //Names, Needs, MaxValue, TotalCount - рассчитаны выше
      for j:= 0 to High(TotalCounts) do
        Values[j]:= ClaimCountForReason(j{period_index}, i{param_index}, FCounts);
      ComparisonHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCounts,
                              COL_COUNT_BAR_USED_COMPAR, COL_COUNT_BAR_VALUE_COMPAR,
                              Categories, APercentNeed);
    end;
  end;

  procedure CountForStatisticTypeAndReasonDraw;
  var
    i, j: Integer;
    ParamSumCounts: TIntMatrix;
  begin
    //paragraph caption
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    R:= R + 1;
    Inc(Order);
    S:= IntToStr(Order) +
        ') Распределение количества рекламационных случаев по ' +
        APartTitle + ' и критериям неисправности';
    if APercentNeed then
      S:= S + ' (с % от суммы рекламаций по ' + APartTitle2 + ')';
    Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
    //data grid
    R:= R + 2;
    ParamSumCounts:= FSumCounts;
    ComparisonReasonTableDraw(R, AYear, ParamSumCounts, nil{not need}, APercentNeed, False{без итого});

    R:= R + 1;

    if not AHistogramNeed then Exit;

    Categories:= VIntToStr(VStep(AYear-High(FCounts), AYear, 1));

    for i:=0 to High(FParamNames) do
    begin
      if not FParamNeeds[i] then continue;

      for j:= 0 to High(TotalCounts) do
        TotalCounts[j]:= FSumCounts[j{period_index}, i{param_index}];

      //histogram caption
      R:= R + 1;
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
      S:= FParamNames[i] + ' - количество рекламаций';
      if APercentNeed then
        S:= S + ' (с % от суммы рекламаций ' + FParamNames[i] + ' за период)';
      Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
      //histogram
      R:= R + 1;
      Names:= FReasonNames;
      Needs:= FReasonNeeds;
      //Values:= ClaimCountForReason(0{period_index}, i{param_index}, FCounts);
      for j:= 0 to High(TotalCounts) do
        Values[j]:= ClaimCountForReason(j{period_index}, i{param_index}, FCounts);
      MaxValue:= MMax(MCut(Values, Needs));
      ComparisonHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCounts,
                              COL_COUNT_BAR_USED_COMPAR, COL_COUNT_BAR_VALUE_COMPAR,
                              Categories, APercentNeed);
    end;
  end;

begin
  if VIsAllFalse(ADataNeed) or
     VIsAllFalse(AReasonNeeds) or
     VIsAllFalse(AParamNeeds) then Exit;

  for k:= 2 to Writer.ColCount do
    Writer.ColWidth[k]:= COL_WIDTH_BAR_COMPAR;

  //кол-во столбцов, занятых таблицами
  N:= COL_COUNT_CELL_YEAR*Length(AClaimCounts){years}*VCountIf(AReasonNeeds, True){reasons};
  //обнуляем ширину лишних столбцов
  if N>COL_COUNT_REPORT_MIN then
    N:= 1{ParamNameColumn} + N + 1
  else
    N:= 1{ParamNameColumn} + COL_COUNT_REPORT_MIN + 1;
  for k:= N to Writer.ColCount-1 do
    Writer.ColWidth[k]:= 0;
  if N <= Writer.ColCount-1 then  //последний столбец width=1 для норм отображения правой границы таблиц
    Writer.ColWidth[Writer.ColCount]:= 1;

  PrepareData(AParamColName, AReasonNeeds, AReasonNames, AParamNeeds, AParamNames,
              AClaimCounts, ASumType);

  Order:= 0;

  //drawing
  Writer.BeginEdit;
  Writer.SetBackgroundClear;
  R:= 1;
  ReportTitleDraw(R, AMotorNamesStr, APeriodStr);
  R:= R + 1;
  if ADataNeed[0] then //[0] - общее количество по виду статистики
    CountTotalForStatisticTypeDraw;
  if ADataNeed[1] then //[1] - общее количество по критериям неисправности
    CountTotalForReasonDraw;
  if ADataNeed[2] then //[2] - количество по виду статистики и критериям неисправности
    CountForStatisticTypeAndReasonDraw;

  Writer.EndEdit;
end;

end.

