unit UStatistic;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics,
  //DK packages utils
  DK_Vector, DK_Matrix, DK_SheetTypes, DK_Const, DK_Math;

const

  COL_COUNT_GRAPH = 120;
  COL_COUNT_TOTAL = 1{ParamNames} + COL_COUNT_GRAPH;
  COL_COUNT_USED  = COL_COUNT_GRAPH div 2;
  COL_COUNT_VALUE = 15;

  COL_WIDTH_GRAPH = 7;
  COL_WIDTH_NAMES = 150; //defaul

  ROW_EMPTY_HEIGH = 5;
  PERCENT_FRAC_DIGITS = 1;

  //REPORT_TITLE_MAIN = 'Отчет по рекламационным случаям электродвигателей';
  //
  //GRAPH_TITLE_MAIN   = 'Распределение количества рекламационных случаев по ';
  //GRAPH_TITLE_DEFECT = GRAPH_TITLE_MAIN + 'причинам возникновения неисправностей';

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

    FCellColCount: Integer;        //кол-во стобцов для целой ячейки (без %)
    FHalfCellColCount: Integer;    //кол-во столбцов для половины ячейки (если нужен столбец %)

    //FPercentNeed: Boolean;         //выводить % от кол-ва

    procedure HorizBarDraw(const ARow: Integer;
                           const AColor: TColor;
                           const AValue, AMaxValue: Integer;
                           const APercent: Double;
                           const ANeedPercent: Boolean);

    procedure HorizBarHistogramRowDraw(var ARow: Integer;
                           const AName: String;
                           const AValue, AMaxValue, ATotalCount: Integer;
                           const ANeedPercent: Boolean);
    procedure HorizBarHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntVector;
                                    const AMaxValue: Integer;
                                    const ANeeds: TBoolVector;
                                    const ATotalCount: Integer;
                                    const ANeedPercent: Boolean;
                                    const ANeedSort: Boolean);

    procedure HorizBarReasonHistogramRowDraw(var ARow: Integer;
                           const AName: String;
                           const AValues: TIntVector;
                           const AMaxValue, ATotalCount: Integer;
                           const ANeedPercent: Boolean);
    procedure HorizBarReasonHistogramDraw(var ARow: Integer;
                                    const ATotalCount: Integer;
                                    const ANeedPercent: Boolean);

    procedure SumTableDraw(var ARow: Integer;
                            const AValueColumnTitle: String;
                            const ANames: TStrVector;
                            const AValues: TIntVector;
                            const ANeeds: TBoolVector;
                            const ATotalCount: Integer;
                            const ANeedPercent: Boolean;
                            const ANeedResume: Boolean);
    procedure ReasonTableDraw(var ARow: Integer;
                               const AParamSumCountForPercents: TIntVector;
                               const ATotalSumCountForPercent: Integer;
                               const ANeedPercent: Boolean;
                               const ANeedResume: Boolean);


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
  VDim(Result{%H-}, COL_COUNT_TOTAL+1, COL_WIDTH_GRAPH);
  Result[0]:= COL_WIDTH_NAMES;
  Result[High(Result)]:= 1;
end;

procedure TStatSheet.HorizBarDraw(const ARow: Integer;
                                  const AColor: TColor;
                                  const AValue, AMaxValue: Integer;
                                  const APercent: Double;
                                  const ANeedPercent: Boolean);
var
  R, C1, C2, ColorColCount: Integer;
  S: String;
begin
  R:= ARow;

  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);

  ColorColCount:= PartRound(COL_COUNT_USED*AValue, AMaxValue);
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
      Writer.WriteNumber(R, C2+1, R, C2+COL_COUNT_VALUE, AValue, cbtLeft)
    else begin
      S:= '0  (0%)';
      Writer.WriteText(R, C2+1, R, C2+COL_COUNT_VALUE, S, cbtLeft)
    end;
  end
  else begin
    if not ANeedPercent then
      Writer.WriteNumber(R, C2+1, R, C2+COL_COUNT_VALUE, AValue, cbtNone)
    else begin
      S:= IntToStr(AValue) + '  (' + Format('%.1f', [APercent]) + '%)';
      Writer.WriteText(R, C2+1, R, C2+COL_COUNT_VALUE, S, cbtNone);
    end;
  end;
end;

procedure TStatSheet.HorizBarHistogramRowDraw(var ARow: Integer;
                                  const AName: String;
                                  const AValue, AMaxValue, ATotalCount: Integer;
                                  const ANeedPercent: Boolean);
var
  R: Integer;
  Percent: Double;
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
  Percent:= Part(100*AValue, ATotalCount);
  HorizBarDraw(R, GRAPH_COLORS_DEFAULT[0], AValue, AMaxValue, Percent, ANeedPercent);

  //EmptyRow
  R:= R + 1;
  Writer.WriteText(R, 1, EmptyStr, cbtRight);
  Writer.RowHeight[R]:= ROW_EMPTY_HEIGH;

  ARow:= R + 1;
end;

procedure TStatSheet.HorizBarHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntVector;
                                    const AMaxValue: Integer;
                                    const ANeeds: TBoolVector;
                                    const ATotalCount: Integer;
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
      HorizBarHistogramRowDraw(R, Names[i], Values[i], AMaxValue, ATotalCount, ANeedPercent);

  ARow:= R;
end;

procedure TStatSheet.HorizBarReasonHistogramRowDraw(var ARow: Integer;
                           const AName: String;
                           const AValues: TIntVector;
                           const AMaxValue, ATotalCount: Integer;
                           const ANeedPercent: Boolean);
var
  R, RR, i: Integer;
  Percent: Double;
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
  Writer.WriteText(R, 1, R+4, 1, AName, cbtRight);
  RR:= R;
  R:= R - 1;
  for i:= 0 to 4 do
  begin
    if not FReasonNeeds[i] then continue;
    R:= R + 1;
    Percent:= Part(100*AValues[i], ATotalCount);
    HorizBarDraw(R, GRAPH_COLORS_DEFAULT[i], AValues[i], AMaxValue, Percent, ANeedPercent);
  end;
  if R<RR then R:= RR;

  //EmptyRow
  R:= R + 1;
  Writer.WriteText(R, 1, EmptyStr, cbtRight);
  Writer.RowHeight[R]:= ROW_EMPTY_HEIGH;

  ARow:= R + 1;
end;

procedure TStatSheet.HorizBarReasonHistogramDraw(var ARow: Integer;
                                                 const ATotalCount: Integer;
                                                 const ANeedPercent: Boolean);
var
  R, i, MaxValue: Integer;
  Values: TIntVector;
begin
  R:= ARow;

  MaxValue:= MMax(FCounts[0{period_index}]);

  for i:= 0 to High(FParamNames) do
  begin
    Values:= ClaimCountForReason(0{period_index}, i{param_index}, FCounts);
    HorizBarReasonHistogramRowDraw(R, FParamNames[i], Values, MaxValue, ATotalCount, ANeedPercent);
  end;

  ARow:= R;
end;

procedure TStatSheet.SumTableDraw(var ARow: Integer;
                                  const AValueColumnTitle: String;
                                  const ANames: TStrVector;
                                  const AValues: TIntVector;
                                  const ANeeds: TBoolVector;
                                  const ATotalCount: Integer;
                                  const ANeedPercent: Boolean;
                                  const ANeedResume: Boolean);
var
  R, C1, C2, i{, Value}: Integer;
  Percent: Double;
begin
  R:= ARow;

  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);

  //заголовок
  C1:= 2;
  C2:= C1 + FCellColCount - 1;
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
    //Value:= FSumCounts[0{period_index}, i{param_index}];
    if not ANeedPercent then
      Writer.WriteNumber(R, C1, R, C2, AValues[i], cbtOuter)
    else begin
      C1:= 2;
      C2:= C1 + FHalfCellColCount - 1;
      Writer.WriteNumber(R, C1, R, C2, AValues[i], cbtOuter);
      C1:= C2 + 1;
      C2:= C1 + FHalfCellColCount - 1;
      //Percent:= Part(Value, FTotalCounts[0{period_index}]);
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
    C2:= C1 + FCellColCount - 1;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.WriteText(R, 1, 'ИТОГО', cbtOuter);
    //Writer.WriteNumber(R, C1, R, C2, FTotalCounts[0{period_index}], cbtOuter);
    Writer.WriteNumber(R, C1, R, C2, ATotalCount, cbtOuter);
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  ARow:= R;
end;

procedure TStatSheet.ReasonTableDraw(var ARow: Integer;
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
    C2:= C1 + FCellColCount - 1;
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
        C2:= C1 + FCellColCount - 1;
        Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
      end
      else begin
        C2:= C1 + FHalfCellColCount - 1;
        Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
        C1:= C2 + 1;
        C2:= C1 + FHalfCellColCount - 1;
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
        C2:= C1 + FCellColCount - 1;
        Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
      end
      else begin
        C2:= C1 + FHalfCellColCount - 1;
        Writer.WriteNumber(R, C1, R, C2, Value, cbtOuter);
        C1:= C2 + 1;
        C2:= C1 + FHalfCellColCount - 1;
        Percent:= Part(Value, ATotalSumCountForPercent);
        Writer.WriteNumber(R, C1, R, C2, Percent, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
      end;
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

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
  R, Order, MaxValue, TotalCount: Integer;
  S: String;
  Needs: TBoolVector;
  Values: TIntVector;
  Names: TStrVector;

  procedure ReportTitleDraw;
  begin
    R:= 1;
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
  end;

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

    R:= R + 2;
    Inc(Order);
    if AAccumNeed then
      S:= 'Накопление'
    else
      S:= 'Распределение';
    S:= IntToStr(Order) +
        ') ' + S + ' общего количества рекламационных случаев по ' +
        APartTitle;
    if APercentNeed then
      S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle +
         ' за период = ' + IntToStr(TotalCount) + ')';
    Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);

    R:= R + 2;
    if AAccumNeed then
      S:= 'Накопление количества' + SYMBOL_BREAK + 'рекламаций'
    else
      S:= 'Количество рекламаций' + SYMBOL_BREAK + 'за период';
    SumTableDraw(R, S, Names, Values, Needs, TotalCount, APercentNeed, not AAccumNeed);

    if not AHistogramNeed then Exit;
    R:= R + 2;
    HorizBarHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCount, APercentNeed, ASortNeed{sort});
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
    R:= R + 2;
    Inc(Order);
    S:= IntToStr(Order) +
        ') Распределение общего количества рекламационных случаев по критериям неисправности';
    if APercentNeed then
      S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle +
         ' за период = ' + IntToStr(TotalCount) + ')';
    Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
    //data grid
    R:= R + 2;
    ParamSumCounts:= nil;
    if APercentNeed then
      ParamSumCounts:= VCreateInt(Length(FParamNames), TotalCount);
    ReasonTableDraw(R, ParamSumCounts, TotalCount, APercentNeed, True{строка итого});

    if not AHistogramNeed then Exit;

    //histogram
    R:= R + 2;
    Names:= FReasonNames;
    Needs:= FReasonNeeds;
    Values:= FReasonCounts[0{period_index}];
    MaxValue:= VMax(VCut(Values, Needs));
    HorizBarHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCount, APercentNeed, False{no sort});

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
        S:= S + ' (с % от суммы рекламаций по всем ' + APartTitle +
            ' за период = ' + IntToStr(TotalCount) + ')';
      Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
      //histogram
      R:= R + 1;
      //Names, Needs, MaxValue, TotalCount - рассчитаны выше
      Values:= ClaimCountForReason(0{period_index}, i{param_index}, FCounts);
      HorizBarHistogramDraw(R, Names, Values, MaxValue, Needs, TotalCount, APercentNeed, False{no sort});
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
    R:= R + 2;
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
    ReasonTableDraw(R, ParamSumCounts, 0{not need}, APercentNeed, False{без итого});

    if not AHistogramNeed then Exit;

    R:= R + 1;
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
        S:= S + ' (с % от суммы рекламаций ' + FParamNames[i] +
            ' за период = ' + IntToStr(TotalCount) + ')';
      Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
      //histogram
      R:= R + 1;
      Names:= FReasonNames;
      Needs:= FReasonNeeds;
      Values:= ClaimCountForReason(0{period_index}, i{param_index}, FCounts);
      MaxValue:= VMax(VCut(Values, Needs));
      HorizBarHistogramDraw(R, Names, Values, MaxValue, Needs,TotalCount, APercentNeed, False{no sort});
    end;
  end;

begin
  if VIsAllFalse(ADataNeed) or
     VIsAllFalse(AReasonNeeds) or
     VIsAllFalse(AParamNeeds) then Exit;

  //source data
  FParamColName:= AParamColName;
  FReasonNeeds:= AReasonNeeds;
  FReasonNames:= AReasonNames;
  FParamNeeds:= AParamNeeds;
  FParamNames:= AParamNames;
  FCounts:= AClaimCounts;
  //FPercentNeed:= APercentNeed;
  Order:= 0;

  //data calculation
  Writer.ColWidth[1]:= MMaxWidth([FReasonNames, FParamNames], Font) + 20;
  if ASumType = 0 then // =0 - по всем критериям
    Needs:= VCreateBool(Length(FReasonNames), True)
  else if ASumType = 1 then // =1 - только по включенным в отчёт критериям
    Needs:= FReasonNeeds;
  FSumCounts:= ClaimCountSum(FCounts, Needs); //сумма рекламаций по параметрам [period_index, param_index]
  FTotalCounts:= ClaimTotalSum(FSumCounts);   //общая сумма рекламаций по периодам [period_index]
  FReasonCounts:= ClaimReasonSum(FCounts);    //общая сумма рекламаций по критериям (причинам) [period_index, reason_index]
  FCellColCount:= COL_COUNT_GRAPH div 5;
  FHalfCellColCount:= FCellColCount div 2;

  //drawing
  Writer.BeginEdit;
  Writer.SetBackgroundClear;
  ReportTitleDraw;
  if ADataNeed[0] then //[0] - общее количество по виду статистики
    CountTotalForStatisticTypeDraw;
  if ADataNeed[1] then //[1] - общее количество по критериям неисправности
    CountTotalForReasonDraw;
  if ADataNeed[2] then //[2] - количество по виду статистики и критериям неисправности
    CountForStatisticTypeAndReasonDraw;

  Writer.EndEdit;
end;

end.

