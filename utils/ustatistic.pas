unit UStatistic;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics, fpstypes, {Math,}
  //DK packages utils
  DK_Vector, DK_Matrix, DK_SheetWriter, DK_SheetTypes, DK_StrUtils, DK_Const;

const

  COL_COUNT_GRAPH = 120;
  COL_COUNT_TOTAL = 1{ParamNames} + COL_COUNT_GRAPH;
  COL_COUNT_USED  = COL_COUNT_GRAPH div 2;
  COL_COUNT_VALUE = 10;

  COL_WIDTH_GRAPH = 7;
  COL_WIDTH_NAMES = 150; //defaul

  ROW_EMPTY_HEIGH = 5;

  //REPORT_TITLE_MAIN = 'Отчет по рекламационным случаям электродвигателей';
  //
  //GRAPH_TITLE_MAIN   = 'Распределение количества рекламационных случаев по ';
  //GRAPH_TITLE_DEFECT = GRAPH_TITLE_MAIN + 'причинам возникновения неисправностей';

  //5 years, 5 reasons
  GRAPH_COLORS_DEFAULT: array [0..4] of TColor = (
    $00BD814F, //blue         // [0] : Не расследовано
    $004696F7, //orange       // [1] : Некачественные комплектующие
    $004D50C0, //red          // [2] : Дефект сборки / изготовления
    $00A26480, //violet       // [3] : Нарушение условий эксплуатации
    $0059BB9B  //green        // [4] : Электродвигатель исправен
  );

type

  { TStatSheet }

  TStatSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  private
    FReasonNeeds: TBoolVector;     //флаги отображения критериев (причин) неисправности
    FReasonNames: TStrVector;      //список критериев (причин) неисправности
    FParamColName: String;         //наименование столбца параметров
    FParamNames: TStrVector;       //список наименований параметров
    FCounts: TIntMatrix3D;         //количество рекламаций [period_index, reason_index, param_index]
    FSumCounts: TIntMatrix;        //сумма рекламаций по параметрам [period_index, param_index]
    FTotalCounts: TIntVector;      //общая сумма рекламаций по периодам [period_index]
    FReasonCounts: TIntMatrix;     //общая сумма рекламаций по критериям (причинам) [period_index, reason_index]

    procedure HorizBarDraw(var ARow: Integer;
                              const AName: String;
                              const AValue, AMaxValue: Integer);
    procedure HorizBarHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntVector;
                                    const ANeeds: TBoolVector);
    procedure HorizBarReasonHistogramDraw(var ARow: Integer;
                                    const ANeeds: TBoolVector);

    procedure SumTableDraw(var ARow: Integer);
    procedure ReasonTableDraw(var ARow: Integer);


  public
    procedure Draw(const AParamColName, APartTitle{дательный падеж}, AMotorNamesStr, APeriodStr: String;
                   const AReasonNeeds: TBoolVector;
                   const AReasonNames, AParamNames: TStrVector;
                   const AClaimCounts: TIntMatrix3D;
                   const AHistogramNeed: Boolean);
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
  function ClaimCountSum(const AClaimCounts: TIntMatrix3D): TIntMatrix;

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
  if High(AClaimCounts[APeriodIndex])<AParamIndex then Exit;
  Result:= MRowGet(AClaimCounts[APeriodIndex], AParamIndex);
end;

function ClaimCountSum(const AClaimCounts: TIntMatrix3D): TIntMatrix;
var
  i, j: Integer;
begin
  Result:= nil;
  if MIsNil(AClaimCounts) then Exit;
  MDim(Result, Length(AClaimCounts), Length(AClaimCounts[0, 0]));
  for i:= 0 to High(AClaimCounts) do
    for j:= 0 to High(AClaimCounts[i, 0]) do
      Result[i, j]:= VSum(ClaimCountForReason(i, j, AClaimCounts));
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
      Result[i, j]:= VSum(AClaimCounts[i, j])
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
  VDim(Result{%H-}, COL_COUNT_TOTAL, COL_WIDTH_GRAPH);
  Result[0]:= COL_WIDTH_NAMES;
end;

procedure TStatSheet.HorizBarDraw(var ARow: Integer; const AName: String;
                                  const AValue, AMaxValue: Integer);
var
  R, C1, C2, ColorColCount: Integer;
begin
  if AMaxValue=0 then
    ColorColCount:= 0
  else
    ColorColCount:= Round(COL_COUNT_USED*AValue/AMaxValue);

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

  Writer.SetBackground($00BD814F);
  C1:= 1{ParamNames} + 1;
  C2:= C1 + ColorColCount - 1;
  Writer.WriteText(R, C1, R, C2, EmptyStr);
  Writer.SetBackgroundClear;

  Writer.SetAlignment(haLeft, vaCenter);
  if AValue=0 then
    Writer.WriteNumber(R, C2+1, R, {Writer.ColCount} C2+COL_COUNT_VALUE, AValue, cbtLeft)
  else
    Writer.WriteNumber(R, C2+1, R, {Writer.ColCount} C2+COL_COUNT_VALUE, AValue, cbtNone);
  Writer.SetAlignment(haCenter, vaCenter);

  //EmptyRow
  R:= R + 1;
  Writer.WriteText(R, 1, EmptyStr, cbtRight);
  Writer.RowHeight[R]:= ROW_EMPTY_HEIGH;

  ARow:= R + 1;
end;

procedure TStatSheet.HorizBarHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntVector;
                                    const ANeeds: TBoolVector);
var
  R, i, MaxValue: Integer;
  Names: TStrVector;
  Values, Indexes: TIntVector;
  Needs: TBoolVector;
begin
  R:= ARow;

  MaxValue:= VMax(AValues);

  VSort(AValues, Indexes, True{DESC});
  Names:= VReplace(ANames, Indexes);
  Values:= VReplace(AValues, Indexes);
  Needs:= VReplace(ANeeds, Indexes);

  for i:= 0 to High(Names) do
    if Needs[i] then
      HorizBarDraw(R, Names[i], Values[i], MaxValue);

  ARow:= R;
end;

procedure TStatSheet.HorizBarReasonHistogramDraw(var ARow: Integer;
                                                 const ANeeds: TBoolVector);
var
  i, j, R, MaxValue: Integer;
begin
  R:= ARow;

  MaxValue:= MMax(FCounts[0{period_index}]);

  ARow:= R;
end;

procedure TStatSheet.SumTableDraw(var ARow: Integer);
var
  R, C1, C2, i: Integer;
begin
  R:= ARow;
  C1:= 2;
  C2:= C1 + (COL_COUNT_GRAPH div 5) - 1;

  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);

  //заголовок
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, FParamColName, cbtOuter, True, True);
  Writer.WriteText(R, C1, R, C2, 'Количество рекламаций' + SYMBOL_BREAK + 'за период', cbtOuter, True, True);
  Writer.DrawBorders(R, C2+1, cbtLeft);

  //данные
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  for i:= 0 to High(FParamNames) do
  begin
    R:= R + 1;
    Writer.WriteText(R, 1, FParamNames[i], cbtOuter);
    Writer.WriteNumber(R, C1, R, C2, FSumCounts[0{period_index}, i{param_index}], cbtOuter);
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  //итого
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, 'ИТОГО', cbtOuter);
  Writer.WriteNumber(R, C1, R, C2, FTotalCounts[0{period_index}], cbtOuter);
  Writer.DrawBorders(R, C2+1, cbtLeft);

  ARow:= R;
end;

procedure TStatSheet.ReasonTableDraw(var ARow: Integer);
var
  R, C1, C2, i, j: Integer;
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
    C2:= C1 + (COL_COUNT_GRAPH div 5) - 1;
    Writer.WriteText(R, C1, R, C2, FReasonNames[i], cbtOuter, True, True);
  end;
  Writer.DrawBorders(R, C2+1, cbtLeft);

  //данные
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  for i:= 0 to High(FParamNames) do
  begin
    R:= R + 1;
    Writer.WriteText(R, 1, FParamNames[i], cbtOuter);
    C2:= 1;
    for j:= 0 to High(FReasonNames) do
    begin
      if not FReasonNeeds[j] then continue;
      C1:= C2 + 1;
      C2:= C1 + (COL_COUNT_GRAPH div 5) - 1;
      Writer.WriteNumber(R, C1, R, C2, FCounts[0{period_index}, j{reason_index}, i{param_index}], cbtOuter);
    end;
    Writer.DrawBorders(R, C2+1, cbtLeft);
  end;

  //итого
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, 'ИТОГО', cbtOuter);
  C2:= 1;
  for i:= 0 to High(FReasonNames) do
  begin
    if not FReasonNeeds[i] then continue;
    C1:= C2 + 1;
    C2:= C1 + (COL_COUNT_GRAPH div 5) - 1;
    Writer.WriteNumber(R, C1, R, C2, FReasonCounts[0{period_index}, i{reason_index}], cbtOuter);
  end;
  Writer.DrawBorders(R, C2+1, cbtLeft);


  ARow:= R;
end;

procedure TStatSheet.Draw(const AParamColName, APartTitle{дательный падеж}, AMotorNamesStr, APeriodStr: String;
                          const AReasonNeeds: TBoolVector;
                          const AReasonNames, AParamNames: TStrVector;
                          const AClaimCounts: TIntMatrix3D;
                          const AHistogramNeed: Boolean);
var
  R, Order: Integer;
  S: String;
  Needs: TBoolVector;
begin
  Writer.ColWidth[1]:= MMaxWidth([AReasonNames, AParamNames], Font);
  FParamColName:= AParamColName;
  FReasonNeeds:= AReasonNeeds;
  FReasonNames:= AReasonNames;
  FParamNames:= AParamNames;
  FCounts:= AClaimCounts;
  FSumCounts:= ClaimCountSum(FCounts);
  FTotalCounts:= ClaimTotalSum(FSumCounts);
  FReasonCounts:= ClaimReasonSum(FCounts);
  Order:= 0;

  Writer.BeginEdit;

  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);

  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  S:= 'Отчет по рекламационным случаям электродвигателей';
  Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount, AMotorNamesStr, cbtNone, True, True);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  S:= 'за период ' + APeriodStr;
  Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);

  Writer.SetAlignment(haLeft, vaCenter);
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Inc(Order);
  S:= IntToStr(Order) +
      ') Распределение общего количества рекламационных случаев по ' +
      APartTitle;
  Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
  R:= R + 1;
  SumTableDraw(R);
  if AHistogramNeed then
  begin
    R:= R + 2;
    Needs:= VCreateBool(Length(FParamNames), True);
    HorizBarHistogramDraw(R, FParamNames, FSumCounts[0{period_index}], Needs);
  end;

  Writer.SetAlignment(haLeft, vaCenter);
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Inc(Order);
  S:= IntToStr(Order) +  ') Распределение количества рекламационных случаев по критериям неисправности';
  Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);
  R:= R + 1;
  ReasonTableDraw(R);
  if AHistogramNeed then
  begin
    R:= R + 2;
    Needs:= FReasonNeeds;
    HorizBarHistogramDraw(R, FReasonNames, FReasonCounts[0{period_index}], Needs);

    R:= R + 2;
    HorizBarReasonHistogramDraw(R, Needs);
  end;




  Writer.EndEdit;
end;

end.

