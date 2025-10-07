unit UStatSheets;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics, fpstypes, fpspreadsheetgrid, fpspreadsheet,
  DateUtils,
  //DK packages utils
  DK_Vector, DK_Matrix, DK_StatPlotter, DK_SheetWriter, DK_SheetUtils,
  DK_SheetConst, DK_Graph, DK_Math, DK_StrUtils, DK_Const, DK_DateUtils,
  DK_PPI,
  //Project utils
  USheets;

type

  { TStatisticSheet }

  TStatisticSheet = class (TObject)
  private
    const
      PARAM_COLUMN_MIN_WIDTH = 200;
      MARGIN_COLUMN_WIDTH = 10;

      HORIZ_HIST_DATAROW_HEIGHT = 60;
      HORIZ_HIST_ADDITION_HEIGHT = 50;

      EMPTY_ROW_HEIGHT = 25;

      DIFFERENT_COLORS_FOR_EACH_REASON = False;
    var
      FWriter: TSheetWriter;
      FFontName: String;
      FFontSize: Integer;
      FUsedReasonsCount: Integer;
      FUsedReasons: TBoolVector;
      FShowPercentColumn: Boolean;
      FShowGraphics: Boolean;
      FGraphWidth: Integer;
      FCalcTotalCountForUsedParamsOnly: Boolean;
      FBeginDate, FEndDate: TDate;
      FMotorNames: String;
      FParamNames, FReasonNames: TStrVector;

    function CalcHistWidth(const AAlignment: TAlignment): Integer;
    function CalcHorizHistHeight(const ARowCount: Integer; const ADataInRowCount: Integer = 1): Integer;
    function CalcVertHistHeight: Integer;

    procedure DrawEmptyRow(var ARow: Integer; const AHeight: Integer = EMPTY_ROW_HEIGHT);

    procedure DrawLinesForTotalCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AAlignment: TAlignment);
    procedure DrawGraphForTotalCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);
    procedure DrawLinesForReasonCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AAlignment: TAlignment);
    procedure DrawGraphForReasonCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);
    procedure DrawGraphForSumReasonCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const ACounts: TIntMatrix;
                                       const AUsedParams: TBoolVector;
                                       const ANeedOrder: Boolean;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);

    procedure DrawHistOnSheet(var ARow: Integer;
                              const AHeight: Integer;
                              const AAlignment: TAlignment;
                              const APlotter: TStatPlotterCustom);

    procedure DrawVertHistForDataVector(var ARow: Integer;
                                    const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                    const AGraphWidth, AGraphHeight: Integer;
                                    const AParamNames: TStrVector;
                                    const ADataValues: TIntVector;
                                    const AAlignment: TAlignment;
                                    const AOneColorPerTick: Boolean = False;
                                    const AColorIndexes: TIntVector = nil);
    procedure DrawVertHistForDataMatrix(var ARow: Integer;
                                    const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                    const AGraphWidth, AGraphHeight: Integer;
                                    const AParamNames, ALegend: TStrVector;
                                    const ADataValues: TIntMatrix;
                                    const AAlignment: TAlignment;
                                    const AOneColorPerTick: Boolean = False;
                                    const AColorIndexes: TIntVector = nil);
    procedure DrawHorizHistForDataVector(var ARow: Integer;
                                    const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                    const AGraphWidth, AGraphHeight: Integer;
                                    const AParamNames: TStrVector;
                                    const ADataValues: TIntVector;
                                    const AAlignment: TAlignment;
                                    const AOneColorPerTick: Boolean = False;
                                    const AColorIndexes: TIntVector = nil);
    procedure DrawHorizHistForDataMatrix(var ARow: Integer;
                                    const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                    const AGraphWidth, AGraphHeight: Integer;
                                    const AParamNames, ALegend: TStrVector;
                                    const ADataValues: TIntMatrix;
                                    const AAlignment: TAlignment;
                                    const AOneColorPerTick: Boolean = False;
                                    const AColorIndexes: TIntVector = nil);

    procedure DrawLinesForVector (var ARow: Integer;
                                   const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                   const AParamNames: TStrVector;
                                   const AParamValues: TIntVector;
                                   const AAlignment: TAlignment;
                                   const AColorIndexes: TIntVector = nil);
    procedure DrawGraphForVector(var ARow: Integer;
                                   const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                   const AParamNames: TStrVector;
                                   const AParamValues: TIntVector;
                                   const AHistType: TDirectionType;
                                   const AAlignment: TAlignment;
                                   const AOneColorPerTick: Boolean = False;
                                   const AColorIndexes: TIntVector = nil);
    procedure DrawLinesForMatrix(var ARow: Integer;
                                   const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                   const AParamNames, ALegend: TStrVector;
                                   const AParamValues: TIntMatrix;
                                   const AAlignment: TAlignment;
                                   const AColorIndexes: TIntVector = nil);
    procedure DrawGraphForMatrix(var ARow: Integer;
                                   const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                   const AParamNames, ALegend: TStrVector;
                                   const AParamValues: TIntMatrix;
                                   const AHistType: TDirectionType;
                                   const AAlignment: TAlignment;
                                   const AOneColorPerTick: Boolean = False;
                                   const AColorIndexes: TIntVector = nil);

    function CalcSumCountForReason(const AValues: TIntVector;
                                   const AUsedParams: TBoolVector): Integer;
    function CalcSumCountForParam(const ACounts: TIntMatrix;
                              const AUsedParams: TBoolVector;
                              const AParamIndex: Integer): Integer;
    function CalcSumCountForParam(const ACounts: TIntMatrix3D;
                              const AUsedParams: TBoolVector;
                              const AParamIndex: Integer): Integer;
    function CalcTotalCount(const ACounts: TIntMatrix;
                            const AUsedParams: TBoolVector): Integer;
    function CalcTotalCount(const ACounts: TIntMatrix3D;
                            const AUsedParams: TBoolVector): Integer;
    function CalcUsedParams(const ACounts: TIntMatrix;
                            const AShowParamIfAllValuesIsZero: Boolean): TBoolVector;
    function CalcUsedParams(const ACounts: TIntMatrix3D;
                            const AShowParamIfAllValuesIsZero: Boolean): TBoolVector;

    function CalcAccumCounts(const ACounts: TIntMatrix): TIntMatrix;
    function CalcAccumCounts(const ACounts: TIntMatrix3D): TIntMatrix3D;

    procedure MouseWheel(Sender: TObject; {%H-}Shift: TShiftState;
                         {%H-}WheelDelta: Integer; {%H-}MousePos: TPoint;
                         var {%H-}Handled: Boolean);
  public
    constructor Create(const AWorksheet: TsWorksheet;
                       const AGrid: TsWorksheetGrid;
                       const AFont: TFont; //if nil -> SHEET_FONT_NAME, SHEET_FONT_SIZE
                       const AColWidths: TIntVector;
                       const AUsedReasonsCount: Integer;
                       const AUsedReasons: TBoolVector;
                       const AShowPercentColumn: Boolean = False);
    destructor  Destroy; override;
    procedure Zoom(const APercents: Integer);
  end;

  { TStatisticSinglePeriodSheet }

  TStatisticSinglePeriodSheet = class (TStatisticSheet)
  private
    const
      TOTAL_WIDTH = 1080;//900;
    var
      FCounts: TIntMatrix;
      FShowLinePercentColumn: Boolean;

    procedure DrawReportTitle(var ARow: Integer; const ATitle: String);

    procedure DrawSinglePeriodDataTableTop(var ARow: Integer;
                                 const ATableTitle, AParamColumnCaption,
                                       ASumCountColumnCaption: String);

    procedure DrawSinglePeriodDataTableValues(var ARow: Integer;
                                 const ACounts: TIntMatrix;
                                 const AUsedParams: TBoolVector;
                                 const ATotalCount: Integer;
                                 const AShowResumeRow: Boolean;
                                 const AParamColumnHorAlignment: TsHorAlignment);

    procedure DrawParamsReport(const ATableTitlePart, AParamColumnCaption: String);

    procedure DrawReport; virtual; abstract;
  public
    constructor Create(const AWorksheet: TsWorksheet;
                       const AGrid: TsWorksheetGrid;
                       const AFont: TFont; //if nil -> SHEET_FONT_NAME, SHEET_FONT_SIZE
                       const AUsedReasons: TBoolVector;
                       const AShowPercentColumn: Boolean = False;
                       const AShowLinePercentColumn: Boolean = False);
    procedure Draw(const ABeginDate, AEndDate: TDate;
                   const AMotorNames: String;
                   const AParamNames, AReasonNames: TStrVector;
                   const ACounts: TIntMatrix;
                   const ACalcTotalCountForUsedParamsOnly,
                         AShowGraphics: Boolean);
  end;

  { TStatisticSinglePeriodAtMotorNamesSheet }

  TStatisticSinglePeriodAtMotorNamesSheet = class (TStatisticSinglePeriodSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSinglePeriodAtDefectNamesSheet }

  TStatisticSinglePeriodAtDefectNamesSheet = class (TStatisticSinglePeriodSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSinglePeriodAtPlaceNamesSheet }

  TStatisticSinglePeriodAtPlaceNamesSheet = class (TStatisticSinglePeriodSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSinglePeriodAtMonthNamesSheet }

  TStatisticSinglePeriodAtMonthNamesSheet = class (TStatisticSinglePeriodSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSinglePeriodAtMonthNamesSumSheet }

  TStatisticSinglePeriodAtMonthNamesSumSheet = class (TStatisticSinglePeriodSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSinglePeriodAtMileagesSheet }

  TStatisticSinglePeriodAtMileagesSheet = class (TStatisticSinglePeriodSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSeveralPeriodsSheet }

  TStatisticSeveralPeriodsSheet = class (TStatisticSheet)
  private
    FCounts: TIntMatrix3D;
    FSumCounts: TIntMatrix;
    FAdditionYearCount: Integer;

    function CalcSumCountMatrix(const ACounts: TIntMatrix3D): TIntMatrix;

    procedure DrawReportTitle(var ARow: Integer; const ATitle: String);

    procedure DrawSinglePeriodDataTableTop(var ARow: Integer;
                                 const ATableTitle, AParamColumnCaption,
                                       ASumCountColumnCaption: String);
    procedure DrawSinglePeriodDataTableValues(var ARow: Integer;
                                 const ACounts: TIntMatrix;
                                 const AUsedParams: TBoolVector;
                                 const ATotalCount: Integer;
                                 const AShowResumeRow: Boolean;
                                 const AParamColumnHorAlignment: TsHorAlignment);

    procedure DrawSeveralPeriodsDataTableTop(var ARow: Integer;
                               const ATableTitle, AParamColumnCaption,
                                     ASumCountColumnCaption: String);
    procedure DrawSeveralPeriodsDataTableValues(var ARow: Integer;
                                 const ACounts: TIntMatrix3D;
                                 const AUsedParams: TBoolVector;
                                 const ATotalCounts: TIntVector;
                                 const AShowResumeRow: Boolean;
                                 const AParamColumnHorAlignment: TsHorAlignment);

    procedure DrawLinesForTotalCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix3D;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AAlignment: TAlignment);
    procedure DrawGraphForTotalCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix3D;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);
    procedure DrawGraphForSumReasonCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const ACounts: TIntMatrix3D;
                                       const AUsedParams: TBoolVector;
                                       const ANeedOrder: Boolean;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);

    procedure DrawLinesForParamCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix3D;
                                       const AUsedParams: TBoolVector;
                                       const AAlignment: TAlignment);
    procedure DrawGraphsForParamCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix3D;
                                       const AUsedParams: TBoolVector;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);

    procedure DrawParamsReport(const ATableTitlePart, AParamColumnCaption: String);

    procedure DrawReport; virtual; abstract;
  public
    constructor Create(const AWorksheet: TsWorksheet;
                       const AGrid: TsWorksheetGrid;
                       const AFont: TFont; //if nil -> SHEET_FONT_NAME, SHEET_FONT_SIZE
                       const AAdditionYearCount: Integer;
                       const AUsedReasons: TBoolVector;
                       const AShowPercentColumn: Boolean = False);
    procedure Draw(const ABeginDate, AEndDate: TDate;
                   const AMotorNames: String;
                   const AParamNames, AReasonNames: TStrVector;
                   const ACounts: TIntMatrix3D;
                   const ACalcTotalCountForUsedParamsOnly,
                         AShowGraphics: Boolean);
  end;

  { TStatisticSeveralPeriodsAtMotorNamesSheet }

  TStatisticSeveralPeriodsAtMotorNamesSheet = class (TStatisticSeveralPeriodsSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSeveralPeriodsAtDefectNamesSheet }

  TStatisticSeveralPeriodsAtDefectNamesSheet = class (TStatisticSeveralPeriodsSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSeveralPeriodsAtPlaceNamesSheet }

  TStatisticSeveralPeriodsAtPlaceNamesSheet = class (TStatisticSeveralPeriodsSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSeveralPeriodsAtMonthNamesSheet }

  TStatisticSeveralPeriodsAtMonthNamesSheet = class (TStatisticSeveralPeriodsSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSeveralPeriodsAtMonthNamesSumSheet }

  TStatisticSeveralPeriodsAtMonthNamesSumSheet = class (TStatisticSeveralPeriodsSheet)
  private
    procedure DrawReport; override;
  end;

  { TStatisticSeveralPeriodsAtMileagesSheet }

  TStatisticSeveralPeriodsAtMileagesSheet = class (TStatisticSeveralPeriodsSheet)
  private
    procedure DrawReport; override;
  end;

  { Utils }

  function LinesForDataVector(const AWidth, AHeight, AFontSize: Integer;
               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
               const AParamNames: TStrVector;
               const ADataValues: TIntVector;
               const AColorIndexes: TIntVector = nil): TStatPlotterLine;
  function LinesForDataMatrix(const AWidth, AHeight, AFontSize: Integer;
               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
               const AParamNames, ALegend: TStrVector;
               const ADataValues: TIntMatrix;
               const AColorIndexes: TIntVector = nil): TStatPlotterLine;

  function VertHistForDataVector(const AWidth, AHeight, AFontSize: Integer;
               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
               const AParamNames: TStrVector;
               const ADataValues: TIntVector;
               const AOneColorPerTick: Boolean = False;
               const AColorIndexes: TIntVector = nil): TStatPlotterVertHist;
  function VertHistForDataMatrix(const AWidth, AHeight, AFontSize: Integer;
               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
               const AParamNames, ALegend: TStrVector;
               const ADataValues: TIntMatrix;
               const AOneColorPerTick: Boolean = False;
               const AColorIndexes: TIntVector = nil): TStatPlotterVertHist;

  function HorizHistForDataVector(const AWidth, AHeight, AFontSize: Integer;
               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
               const AParamNames: TStrVector;
               const ADataValues: TIntVector;
               const AOneColorPerTick: Boolean = False;
               const AColorIndexes: TIntVector = nil): TStatPlotterHorizHist;
  function HorizHistForDataMatrix(const AWidth, AHeight, AFontSize: Integer;
               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
               const AParamNames, ALegend: TStrVector;
               const ADataValues: TIntMatrix;
               const AOneColorPerTick: Boolean = False;
               const AColorIndexes: TIntVector = nil): TStatPlotterHorizHist;


implementation

{ TStatisticSheet }

function TStatisticSheet.CalcHistWidth(const AAlignment: TAlignment): Integer;
begin
  if AAlignment=taCenter then
    Result:= FGraphWidth
  else begin
    if FWriter.HasGrid then
      Result:= (FGraphWidth - MARGIN_COLUMN_WIDTH) div 2
    else
      Result:= (FGraphWidth - 8*MARGIN_COLUMN_WIDTH) div 2;
  end;
  //Result:= FWriter.ApplyZoom(WidthFromDefaultToScreen(Result));
  //Result:= WidthFromDefaultToScreen(Result);
end;

function TStatisticSheet.CalcHorizHistHeight(const ARowCount: Integer;
  const ADataInRowCount: Integer): Integer;
var
  H: Integer;
begin
  H:= Round(ADataInRowCount*HORIZ_HIST_DATAROW_HEIGHT*0.45);
  H:= Max(H, HORIZ_HIST_DATAROW_HEIGHT);
  Result:= ARowCount*H + HORIZ_HIST_ADDITION_HEIGHT;
  Result:= HeightFromDefaultToScreen(FWriter.ApplyZoom(Result));
end;

function TStatisticSheet.CalcVertHistHeight: Integer;
begin
  Result:= Round(0.4*FGraphWidth);
  //Result:= HeightFromDefaultToScreen(Result);
end;

procedure TStatisticSheet.DrawEmptyRow(var ARow: Integer; const AHeight: Integer = EMPTY_ROW_HEIGHT);
begin
  Inc(ARow);
  FWriter.WriteText(ARow, 1, EmptyStr);
  FWriter.RowHeight[ARow]:= AHeight;
end;

procedure TStatisticSheet.DrawLinesForTotalCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AAlignment: TAlignment);
var
  ParamNames: TStrVector;
  ParamValues: TIntVector;
  UsedParams: TBoolVector;
begin
  if not FUsedReasons[0] then Exit;

  if VIsNil(ASortIndexes) then
  begin
    UsedParams:= AUsedParams;
    ParamNames:= AParamNames;
    ParamValues:= AParamValues[0];
  end
  else begin
    UsedParams:= VReplace(AUsedParams, ASortIndexes);
    ParamNames:= VReplace(AParamNames, ASortIndexes);
    ParamValues:= VReplace(AParamValues[0], ASortIndexes);
  end;
  ParamNames:= VCut(ParamNames, UsedParams);
  ParamValues:= VCut(ParamValues, UsedParams);

  DrawLinesForVector(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, ParamValues, AAlignment, ASortIndexes);
end;

procedure TStatisticSheet.DrawGraphForTotalCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);
var
  ParamNames: TStrVector;
  ParamValues, ColorIndexes: TIntVector;
  UsedParams: TBoolVector;
begin
  if not FUsedReasons[0] then Exit;

  if VIsNil(ASortIndexes) then
  begin
    UsedParams:= AUsedParams;
    ParamNames:= AParamNames;
    ParamValues:= AParamValues[0];
  end
  else begin
    UsedParams:= VReplace(AUsedParams, ASortIndexes);
    ParamNames:= VReplace(AParamNames, ASortIndexes);
    ParamValues:= VReplace(AParamValues[0], ASortIndexes);
  end;
  ParamNames:= VCut(ParamNames, UsedParams);
  ParamValues:= VCut(ParamValues, UsedParams);

  ColorIndexes:= nil;
  if AOneColorPerTick then
    ColorIndexes:= ASortIndexes;
  DrawGraphForVector(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, ParamValues, AHistType, AAlignment,
                     AOneColorPerTick, ColorIndexes);
end;

procedure TStatisticSheet.DrawLinesForReasonCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AAlignment: TAlignment);
var
  i: Integer;
  ParamNames: TStrVector;
  ParamValues: TIntVector;
  UsedParams: TBoolVector;
  Legend: TStrVector;
  DataMatrix: TIntMatrix;
begin
  if VCountIf(FUsedReasons, True, 1)<2 then Exit;

  if VIsNil(ASortIndexes) then
  begin
    UsedParams:= AUsedParams;
    ParamNames:= AParamNames;
  end
  else begin
    UsedParams:= VReplace(AUsedParams, ASortIndexes);
    ParamNames:= VReplace(AParamNames, ASortIndexes);
  end;

  ParamNames:= VCut(ParamNames, UsedParams);
  if Length(ParamNames)<2 then Exit;

  Legend:= nil;
  DataMatrix:= nil;
  for i:= 1 to High(FUsedReasons) do
  begin
    if not FUsedReasons[i] then continue;
    VAppend(Legend, FReasonNames[i]);
    if VIsNil(ASortIndexes) then
      ParamValues:= AParamValues[i]
    else
      ParamValues:= VReplace(AParamValues[i], ASortIndexes);
    ParamValues:= VCut(ParamValues, UsedParams);
    MAppend(DataMatrix, ParamValues);
  end;

  DrawLinesForMatrix(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, Legend, DataMatrix, AAlignment, ASortIndexes);
end;

procedure TStatisticSheet.DrawGraphForReasonCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);
var
  i: Integer;
  ParamNames: TStrVector;
  ParamValues, ColorIndexes: TIntVector;
  UsedParams: TBoolVector;
  Legend: TStrVector;
  DataMatrix: TIntMatrix;
begin
  if VCountIf(FUsedReasons, True, 1)<2 then Exit;

  if VIsNil(ASortIndexes) then
  begin
    UsedParams:= AUsedParams;
    ParamNames:= AParamNames;
  end
  else begin
    UsedParams:= VReplace(AUsedParams, ASortIndexes);
    ParamNames:= VReplace(AParamNames, ASortIndexes);
  end;

  ParamNames:= VCut(ParamNames, UsedParams);
  if Length(ParamNames)<2 then Exit;

  Legend:= nil;
  DataMatrix:= nil;
  for i:= 1 to High(FUsedReasons) do
  begin
    if not FUsedReasons[i] then continue;
    VAppend(Legend, FReasonNames[i]);
    if VIsNil(ASortIndexes) then
      ParamValues:= AParamValues[i]
    else
      ParamValues:= VReplace(AParamValues[i], ASortIndexes);
    ParamValues:= VCut(ParamValues, UsedParams);
    MAppend(DataMatrix, ParamValues);
  end;

  ColorIndexes:= nil;
  if AOneColorPerTick then
    ColorIndexes:= ASortIndexes;
  DrawGraphForMatrix(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, Legend, DataMatrix, AHistType, AAlignment,
                     AOneColorPerTick, ColorIndexes);
end;

procedure TStatisticSheet.DrawGraphForSumReasonCounts(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const ACounts: TIntMatrix;
                                       const AUsedParams: TBoolVector;
                                       const ANeedOrder: Boolean;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);
var
  i: Integer;
  ParamNames: TStrVector;
  ParamValues, Indexes, ColorIndexes: TIntVector;
begin
  ParamNames:= VCut(FReasonNames, FUsedReasons, 1);
  if Length(ParamNames)<2 then Exit;
  ParamValues:= nil;
  for i:= 1 to High(ACounts) do
    if FUsedReasons[i] then
      VAppend(ParamValues, CalcSumCountForReason(ACounts[i], AUsedParams));

  Indexes:= nil;
  if ANeedOrder then
  begin
    VSort(ParamValues, Indexes);
    ParamNames:= VReplace(ParamNames, Indexes);
    ParamValues:= VReplace(ParamValues, Indexes);
  end;

  ColorIndexes:= nil;
  if AOneColorPerTick then
    ColorIndexes:= Indexes;

  DrawGraphForVector(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, ParamValues, AHistType, AAlignment,
                     AOneColorPerTick, ColorIndexes);
end;

procedure TStatisticSheet.DrawHistOnSheet(var ARow: Integer;
                              const AHeight: Integer;
                              const AAlignment: TAlignment;
                              const APlotter: TStatPlotterCustom);
var
  i, R, ColNum, RowCount, RowNum, RowHeight: Integer;
  ScaleX, ScaleY, OffsetX, OffsetY: Double;
  Stream: TMemoryStream;

  procedure CalcParams;
  begin
    ScaleX:= 96/ScreenInfo.PixelsPerInchX;
    ScaleY:= 96/ScreenInfo.PixelsPerInchY;
    OffsetX:= 0;
    OffsetY:= 0;
    ColNum:= 2;
    if AAlignment=taRightJustify then
    begin
      ColNum:= FWriter.ColCount;
      OffsetX:= - PixelToMillimeter(CalcHistWidth(AAlignment));
    end;
  end;

begin
  if not Assigned(APlotter) then Exit;

  RowNum:= ARow + 1;

  R:= ARow;
  RowHeight:= ROW_HEIGHT_DEFAULT; //EMPTY_ROW_HEIGHT
  RowCount:= AHeight div RowHeight;
  for i:= 1 to RowCount do
    DrawEmptyRow(R, RowHeight);
  if (AHeight mod RowHeight)>0 then
    DrawEmptyRow(R, AHeight - RowCount*RowHeight);
  ARow:= R;

  CalcParams;
  Stream:= TMemoryStream.Create;
  try
    APlotter.PNG.SaveToStream(Stream);
    FWriter.WriteImage(RowNum, ColNum, Stream, OffsetX, OffsetY, ScaleX, ScaleY);
    if FWriter.HasGrid then
      FWriter.Grid.Refresh;
  finally
    FreeAndNil(Stream);
  end;
end;

procedure TStatisticSheet.DrawVertHistForDataVector(var ARow: Integer;
                    const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                    const AGraphWidth, AGraphHeight: Integer;
                    const AParamNames: TStrVector;
                    const ADataValues: TIntVector;
                    const AAlignment: TAlignment;
                    const AOneColorPerTick: Boolean = False;
                    const AColorIndexes: TIntVector = nil);
var
  Plotter: TStatPlotterVertHist;
begin
  Plotter:= VertHistForDataVector(AGraphWidth, AGraphHeight, FFontSize, FFontName,
                                  AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                                  AParamNames, ADataValues,
                                  AOneColorPerTick, AColorIndexes);
  try
    DrawHistOnSheet(ARow, AGraphHeight, AAlignment, Plotter);
  finally
    FreeAndNil(Plotter);
  end;
end;

procedure TStatisticSheet.DrawVertHistForDataMatrix(var ARow: Integer;
                    const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                    const AGraphWidth, AGraphHeight: Integer;
                    const AParamNames, ALegend: TStrVector;
                    const ADataValues: TIntMatrix;
                    const AAlignment: TAlignment;
                    const AOneColorPerTick: Boolean = False;
                    const AColorIndexes: TIntVector = nil);
var
  Plotter: TStatPlotterVertHist;
begin
  Plotter:= VertHistForDataMatrix(AGraphWidth, AGraphHeight, FFontSize, FFontName,
                                  AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                                  AParamNames, ALegend, ADataValues,
                                  AOneColorPerTick, AColorIndexes);
  try
    DrawHistOnSheet(ARow, AGraphHeight, AAlignment, Plotter);
  finally
    FreeAndNil(Plotter);
  end;
end;

procedure TStatisticSheet.DrawHorizHistForDataVector(var ARow: Integer;
                  const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                  const AGraphWidth, AGraphHeight: Integer;
                  const AParamNames: TStrVector;
                  const ADataValues: TIntVector;
                  const AAlignment: TAlignment;
                  const AOneColorPerTick: Boolean = False;
                  const AColorIndexes: TIntVector = nil);
var
  Plotter: TStatPlotterHorizHist;
begin
  Plotter:= HorizHistForDataVector(AGraphWidth, AGraphHeight, FFontSize, FFontName,
                                   AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                                   AParamNames, ADataValues,
                                   AOneColorPerTick, AColorIndexes);
  try
    DrawHistOnSheet(ARow, AGraphHeight, AAlignment, Plotter);
  finally
    FreeAndNil(Plotter);
  end;
end;

procedure TStatisticSheet.DrawHorizHistForDataMatrix(var ARow: Integer;
                  const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                  const AGraphWidth, AGraphHeight: Integer;
                  const AParamNames, ALegend: TStrVector;
                  const ADataValues: TIntMatrix;
                  const AAlignment: TAlignment;
                  const AOneColorPerTick: Boolean = False;
                  const AColorIndexes: TIntVector = nil);
var
  Plotter: TStatPlotterHorizHist;
begin
  Plotter:= HorizHistForDataMatrix(AGraphWidth, AGraphHeight, FFontSize, FFontName,
                                  AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                                  AParamNames, ALegend, ADataValues,
                                  AOneColorPerTick, AColorIndexes);
  try
    DrawHistOnSheet(ARow, AGraphHeight, AAlignment, Plotter);
  finally
    FreeAndNil(Plotter);
  end;
end;

procedure TStatisticSheet.DrawLinesForVector(var ARow: Integer;
                                   const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                   const AParamNames: TStrVector;
                                   const AParamValues: TIntVector;
                                   const AAlignment: TAlignment;
                                   const AColorIndexes: TIntVector = nil);
var
  GraphWidth, GraphHeight: Integer;
  Plotter: TStatPlotterLine;
begin
  GraphWidth:= CalcHistWidth(AAlignment);
  GraphHeight:= CalcVertHistHeight;

  Plotter:= LinesForDataVector(GraphWidth, GraphHeight, FFontSize, FFontName,
                                  AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                                  AParamNames, AParamValues, AColorIndexes);
  try
    DrawHistOnSheet(ARow, GraphHeight, AAlignment, Plotter);
  finally
    FreeAndNil(Plotter);
  end;

  DrawEmptyRow(ARow);
end;

procedure TStatisticSheet.DrawGraphForVector(var ARow: Integer;
                   const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                   const AParamNames: TStrVector;
                   const AParamValues: TIntVector;
                   const AHistType: TDirectionType;
                   const AAlignment: TAlignment;
                   const AOneColorPerTick: Boolean = False;
                   const AColorIndexes: TIntVector = nil);
var
  DataRowCount: Integer;
  GraphWidth, GraphHeight: Integer;
begin
  DataRowCount:= Length(AParamNames);
  if DataRowCount<2 then Exit;

  GraphWidth:= CalcHistWidth(AAlignment);
  if AHistType=dtHorizontal then
  begin
    GraphHeight:= CalcHorizHistHeight(DataRowCount);
    DrawHorizHistForDataVector(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                              GraphWidth, GraphHeight,
                              AParamNames, AParamValues, AAlignment,
                              AOneColorPerTick, AColorIndexes);
  end
  else begin  //dtVertical
    GraphHeight:= CalcVertHistHeight;
    DrawVertHistForDataVector(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                              GraphWidth, GraphHeight,
                              AParamNames, AParamValues, AAlignment,
                              AOneColorPerTick, AColorIndexes);
  end;

  DrawEmptyRow(ARow);
end;

procedure TStatisticSheet.DrawLinesForMatrix(var ARow: Integer;
                                   const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                   const AParamNames, ALegend: TStrVector;
                                   const AParamValues: TIntMatrix;
                                   const AAlignment: TAlignment;
                                   const AColorIndexes: TIntVector = nil);
var
  GraphWidth, GraphHeight: Integer;
  Plotter: TStatPlotterLine;
begin
  GraphWidth:= CalcHistWidth(AAlignment);
  GraphHeight:= CalcVertHistHeight;

  Plotter:= LinesForDataMatrix(GraphWidth, GraphHeight, FFontSize, FFontName,
                                  AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                                  AParamNames, ALegend, AParamValues, AColorIndexes);
  try
    DrawHistOnSheet(ARow, GraphHeight, AAlignment, Plotter);
  finally
    FreeAndNil(Plotter);
  end;

  DrawEmptyRow(ARow);
end;

procedure TStatisticSheet.DrawGraphForMatrix(var ARow: Integer;
                   const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                   const AParamNames, ALegend: TStrVector;
                   const AParamValues: TIntMatrix;
                   const AHistType: TDirectionType;
                   const AAlignment: TAlignment;
                   const AOneColorPerTick: Boolean = False;
                   const AColorIndexes: TIntVector = nil);
var
  DataRowCount, DataInRowCount: Integer;
  GraphWidth, GraphHeight: Integer;
begin
  DataInRowCount:= Length(AParamValues);
  if DataInRowCount<2 then Exit;

  DataRowCount:= Length(AParamNames);
  if DataRowCount<2 then Exit;

  GraphWidth:= CalcHistWidth(AAlignment);
  if AHistType=dtHorizontal then
  begin
    GraphHeight:= CalcHorizHistHeight(DataRowCount, DataInRowCount);
    DrawHorizHistForDataMatrix(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                              GraphWidth, GraphHeight,
                              AParamNames, ALegend, AParamValues, AAlignment,
                              AOneColorPerTick, AColorIndexes);
  end
  else begin  //dtVertical
    GraphHeight:= CalcVertHistHeight;
    DrawVertHistForDataMatrix(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                              GraphWidth, GraphHeight,
                              AParamNames, ALegend, AParamValues, AAlignment,
                              AOneColorPerTick, AColorIndexes);
  end;

  DrawEmptyRow(ARow);
end;

function TStatisticSheet.CalcSumCountForReason(const AValues: TIntVector;
  const AUsedParams: TBoolVector): Integer;
begin
   if FCalcTotalCountForUsedParamsOnly then
    Result:= VSumIf(AValues, AUsedParams, True)
  else
    Result:= VSum(AValues);
end;

function TStatisticSheet.CalcSumCountForParam(const ACounts: TIntMatrix;
  const AUsedParams: TBoolVector; const AParamIndex: Integer): Integer;
var
  i: Integer;
begin
  Result:= 0;
  if not AUsedParams[AParamIndex] then Exit;
  Result:= ACounts[0, AParamIndex];
  if (not FCalcTotalCountForUsedParamsOnly) or (Result=0) then Exit;

  Result:= 0;
  for i:= 1 to High(FUsedReasons) do
    if FUsedReasons[i] then
      Result:= Result + ACounts[i, AParamIndex];
end;

function TStatisticSheet.CalcSumCountForParam(const ACounts: TIntMatrix3D;
  const AUsedParams: TBoolVector; const AParamIndex: Integer): Integer;
var
  i, k: Integer;
begin
  Result:= 0;
  if not AUsedParams[AParamIndex] then Exit;

  for k:=0 to High(ACounts) do
    Result:= Result + ACounts[k, 0, AParamIndex];
  if (not FCalcTotalCountForUsedParamsOnly) or (Result=0) then Exit;

  Result:= 0;
  for k:=0 to High(ACounts) do
    for i:= 1 to High(FUsedReasons) do
      if FUsedReasons[i] then
        Result:= Result + ACounts[k, i, AParamIndex];
end;

function TStatisticSheet.CalcTotalCount(const ACounts: TIntMatrix;
  const AUsedParams: TBoolVector): Integer;
var
  i: Integer;
begin
  if FCalcTotalCountForUsedParamsOnly then
  begin
    Result:= 0;
    for i:=0 to High(AUsedParams) do
      Result:= Result + CalcSumCountForParam(ACounts, AUsedParams, i);
  end
  else
    Result:= VSum(ACounts[0]);
end;

function TStatisticSheet.CalcTotalCount(const ACounts: TIntMatrix3D;
  const AUsedParams: TBoolVector): Integer;
var
  i: Integer;
begin
  Result:= 0;
  for i:= 0 to High(ACounts) do
    Result:= Result + CalcTotalCount(ACounts[i], AUsedParams);
end;

function TStatisticSheet.CalcUsedParams(const ACounts: TIntMatrix;
  const AShowParamIfAllValuesIsZero: Boolean): TBoolVector;
var
  i,j,n: Integer;
begin
  VDim(Result{%H-}, Length(FParamNames), True);
  if AShowParamIfAllValuesIsZero then Exit;

  //проверка на ненулевые данные для каждого параметра
  for i:= 0 to High(FParamNames) do
  begin
    n:= ACounts[0, i];
    if n>0 then
    begin
      n:= 0;
      for j:= 1 to High(FUsedReasons) do
        if FUsedReasons[j] then
          n:= n + ACounts[j, i];
    end;
    Result[i]:= n>0;
  end;
end;

function TStatisticSheet.CalcUsedParams(const ACounts: TIntMatrix3D;
  const AShowParamIfAllValuesIsZero: Boolean): TBoolVector;
var
  i,j,k,n: Integer;
begin
  VDim(Result{%H-}, Length(FParamNames), True);
  if AShowParamIfAllValuesIsZero then Exit;

  //проверка на ненулевые данные для каждого параметра
  for i:= 0 to High(FParamNames) do
  begin
    n:= 0;
    for k:= 0 to High(ACounts) do
      n:= n + ACounts[k, 0, i];
    if n>0 then
    begin
      n:= 0;
      for k:= 0 to High(ACounts) do
        for j:= 1 to High(FUsedReasons) do
          if FUsedReasons[j] then
            n:= n + ACounts[k, j, i];
    end;
    Result[i]:= n>0;
  end;
end;

function TStatisticSheet.CalcAccumCounts(const ACounts: TIntMatrix): TIntMatrix;
var
  i,j,k: Integer;
begin
  Result:= MCut(ACounts);
  k:= High(Result[0]);
  for i:=0 to High(Result) do
    for j:=1 to k do
      Result[i,j]:= Result[i,j] + Result[i,j-1];
end;

function TStatisticSheet.CalcAccumCounts(const ACounts: TIntMatrix3D): TIntMatrix3D;
var
  i: Integer;
begin
  Result:= MCut(ACounts);
  for i:= 0 to High(ACounts) do
    Result[i]:= CalcAccumCounts(ACounts[i]);
end;

procedure TStatisticSheet.MouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
  (Sender as TsWorksheetGrid).Invalidate;
end;

constructor TStatisticSheet.Create(const AWorksheet: TsWorksheet;
                                   const AGrid: TsWorksheetGrid;
                                   const AFont: TFont; //if nil -> SHEET_FONT_NAME, SHEET_FONT_SIZE
                                   const AColWidths: TIntVector;
                                   const AUsedReasonsCount: Integer;
                                   const AUsedReasons: TBoolVector;
                                   const AShowPercentColumn: Boolean = False);
begin
  if Assigned(AFont) then
  begin
    FFontName:= AFont.Name;
    FFontSize:= AFont.Size;
  end
  else begin
    FFontName:= SHEET_FONT_NAME;
    FFontSize:= SHEET_FONT_SIZE;
  end;

  FUsedReasons:= AUsedReasons;
  FShowPercentColumn:= AShowPercentColumn;
  FUsedReasonsCount:= AUsedReasonsCount;
  FShowGraphics:= False;

  FWriter:= TSheetWriter.Create(AColWidths, AWorksheet, AGrid);
  if FWriter.HasGrid then
    FWriter.Grid.OnMouseWheel:= @MouseWheel;

  FGraphWidth:= FWriter.ColsWidth(2, FWriter.ColCount-1);
end;

destructor TStatisticSheet.Destroy;
begin
  FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TStatisticSheet.Zoom(const APercents: Integer);
begin
  FWriter.SetZoom(APercents);
end;

{ TStatisticSinglePeriodSheet }

procedure TStatisticSinglePeriodSheet.DrawReportTitle(var ARow: Integer; const ATitle: String);
var
  R: Integer;
  S: String;
begin
  FWriter.SetAlignment(haCenter, vaCenter);

  R:= ARow;
  S:= ATitle;
  //S:= 'Отчет по рекламационным случаям электродвигателей';
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);

  if FMotorNames<>EmptyStr then
  begin
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.WriteText(R, 1, R, FWriter.ColCount, FMotorNames, cbtNone, True, True);
  end;

  R:= R + 1;
  S:= 'за ';
  if SameDate(FBeginDate, FEndDate) then
    S:= S + FormatDateTime('dd.mm.yyyy', FBeginDate)
  else
    S:= S + 'период с ' +
        FormatDateTime('dd.mm.yyyy', FBeginDate) + ' по ' +
        FormatDateTime('dd.mm.yyyy', FEndDate);
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);

  //EmptyRow
  R:= R + 1;
  FWriter.WriteText(R, 1, R, FWriter.ColCount, EmptyStr, cbtNone);
  FWriter.RowHeight[R]:= 10;

  ARow:= R + 1;
end;

procedure TStatisticSinglePeriodSheet.DrawSinglePeriodDataTableTop(
  var ARow: Integer; const ATableTitle, AParamColumnCaption,
  ASumCountColumnCaption: String);
var
  i, R1, R2, C1, C2, Count: Integer;

  procedure DrawCaption(const ACaption: String; const AIsLast: Boolean);
  begin
    C1:= C2 + 1;
    C2:= C1 + Ord(FShowPercentColumn) + Ord(FShowLinePercentColumn) + Ord(AIsLast);
    FWriter.WriteText(R1, C1, R1, C2, ACaption, cbtOuter, True, True);
  end;

begin
  R1:= ARow;

  if not SEmpty(ATableTitle) then
  begin
    FWriter.SetFont(FFontName, FFontSize+1, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R1, 1, R1, FWriter.ColCount, ATableTitle, cbtNone, True, True);
    Inc(R1);
  end;

  R2:= R1;
  C1:= 1;
  C2:= 2;
  FWriter.SetFont(FFontName, FFontSize-1, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R1, C1, R2, C2, AParamColumnCaption, cbtOuter, True, True);

  Count:= 0;
  for i:= 0 to High(FReasonNames) do
  begin
    if not FUsedReasons[i] then continue;
    Inc(Count);
    if i=0 then
      DrawCaption(ASumCountColumnCaption, Count=FUsedReasonsCount)
    else
      DrawCaption(FReasonNames[i], Count=FUsedReasonsCount);
  end;

  FWriter.DrawBorders(R1, 1, R2, FWriter.ColCount, cbtAll);
  ARow:= R2;
end;

procedure TStatisticSinglePeriodSheet.DrawSinglePeriodDataTableValues(var ARow: Integer;
                                 const ACounts: TIntMatrix;
                                 const AUsedParams: TBoolVector;
                                 const ATotalCount: Integer;
                                 const AShowResumeRow: Boolean;
                                 const AParamColumnHorAlignment: TsHorAlignment);
var
  i, R: Integer;

  procedure DrawParamName(var ARowNum, AColNum: Integer; const AName: String);
  var
    RR, C1, C2: Integer;
  begin
    RR:= ARowNum;
    C1:= AColNum;
    C2:= C1 + 1;
    FWriter.SetAlignment(AParamColumnHorAlignment, vaCenter);
    FWriter.WriteText(RR, C1, RR, C2, AName, cbtOuter, True, True);
    ARowNum:= RR;
    AColNum:= C2;
  end;

  procedure DrawParamValues(var ARowNum, AColNum: Integer;
                            const AValue, ALineTotalValue: Integer; const AIsLast: Boolean);
  var
    RR, C1, C2: Integer;
    DD: Double;
  begin
    RR:= ARowNum;
    C1:= AColNum + 1;
    C2:= C1;
    if (not FShowPercentColumn) and (not FShowLinePercentColumn) and AIsLast then
      C2:= C2 + 1;
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteNumber(RR, C1, RR, C2, AValue, cbtOuter);
    if FShowPercentColumn then
    begin
      if FShowPercentColumn and FShowLinePercentColumn then
        FWriter.SetBackground($00E9F5E8);
      C1:= C2 + 1;
      C2:= C1 + Ord(AIsLast)*Ord(not FShowLinePercentColumn);
      DD:= Part(AValue, ATotalCount);
      FWriter.WriteNumber(RR, C1, RR, C2, DD, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
    end;
    if FShowLinePercentColumn then
    begin
      if FShowPercentColumn and FShowLinePercentColumn then
        FWriter.SetBackground($00E1F8FF);
      C1:= C2 + 1;
      C2:= C1 + Ord(AIsLast);
      if ALineTotalValue>0 then
      begin
        DD:= Part(AValue, ALineTotalValue);
        FWriter.WriteNumber(RR, C1, RR, C2, DD, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
      end
      else
        FWriter.WriteText(RR, C1, RR, C2, '–', cbtOuter);
    end;
    FWriter.SetBackgroundClear;
    ARowNum:= RR;
    AColNum:= C2;
  end;

  procedure DrawDataRow(var ARowNum: Integer; const AParamName: String; const AIndex: Integer);
  var
    k, ReasonCount, RR, C, DataValue, LineTotalValue: Integer;
  begin
    if not AUsedParams[AIndex] then Exit;
    FWriter.SetFont(FFontName, FFontSize-1, [], clBlack);
    RR:= ARowNum + 1;
    C:= 1;
    DrawParamName(RR, C, AParamName);
    LineTotalValue:= CalcSumCountForParam(ACounts, AUsedParams, AIndex);
    ReasonCount:= 0;
    for k:= 0 to High(FUsedReasons) do
    begin
      if not FUsedReasons[k] then continue;
      Inc(ReasonCount);

      if k=0 then
        DataValue:= LineTotalValue
      else
        DataValue:= ACounts[k,AIndex];

      DrawParamValues(RR, C, DataValue, LineTotalValue, ReasonCount=FUsedReasonsCount);
    end;
    ARowNum:= RR;
  end;

  procedure DrawResumeRow(var ARowNum: Integer);
  var
    k, ReasonCount, RR, C, DataValue: Integer;
  begin
    if not AShowResumeRow then Exit;
    if (VCountIf(AUsedParams, True)<=1) and FCalcTotalCountForUsedParamsOnly then Exit;
    FWriter.SetFont(FFontName, FFontSize-1, [fsBold], clBlack);
    RR:= ARowNum + 1;
    C:= 1;
    DrawParamName(RR, C, 'ИТОГО');
    ReasonCount:= 0;
    for k:= 0 to High(FUsedReasons) do
    begin
      if not FUsedReasons[k] then continue;
      Inc(ReasonCount);
      if k=0 then
        DataValue:= ATotalCount
      else
        DataValue:= CalcSumCountForReason(ACounts[k], AUsedParams);
      DrawParamValues(RR, C, DataValue, 0, ReasonCount=FUsedReasonsCount);
    end;
    ARowNum:= RR;
  end;

begin
  R:= ARow;
  for i:= 0 to High(FParamNames) do
    DrawDataRow(R, FParamNames[i], i);
  DrawResumeRow(R);
  DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  ARow:= R;
end;

procedure TStatisticSinglePeriodSheet.DrawParamsReport(const ATableTitlePart,
  AParamColumnCaption: String);
var
  R: Integer;
  UsedParams: TBoolVector;
  Indexes: TIntVector;
  TotalCount: Integer;
begin
  R:= 1;
  DrawReportTitle(R, 'Отчет по рекламационным случаям электродвигателей');
  DrawSinglePeriodDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                   ATableTitlePart,
                   AParamColumnCaption, 'Всего за период');

  UsedParams:= CalcUsedParams(FCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)=0 then Exit;

  TotalCount:= CalcTotalCount(FCounts, UsedParams); //от этого считаем процент
  DrawSinglePeriodDataTableValues(R, FCounts, UsedParams, TotalCount, True, haLeft);

  if FShowGraphics then
  begin
    VSort(FCounts[0], Indexes); //в гистограммах сортировка по сумме случаев за период

    DrawGraphForTotalCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                ATableTitlePart,
                            EmptyStr, EmptyStr,
                            FParamNames, FCounts, Indexes, UsedParams,
                            dtHorizontal, taCenter);
    DrawGraphForSumReasonCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей',
                                EmptyStr, EmptyStr,
                                FCounts, UsedParams, True{сортировка},
                                dtHorizontal, taCenter,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});
    DrawGraphForReasonCounts(R, 'Распределение количества рекламационных случаев по ' +
                                ATableTitlePart + ' и '+
                                'причинам возникновения неисправностей',
                             EmptyStr, EmptyStr,
                             FParamNames, FCounts, Indexes, UsedParams,
                             dtHorizontal, taCenter);
  end;
end;

constructor TStatisticSinglePeriodSheet.Create(const AWorksheet: TsWorksheet;
                                        const AGrid: TsWorksheetGrid;
                                        const AFont: TFont; //if nil -> SHEET_FONT_NAME, SHEET_FONT_SIZE
                                        const AUsedReasons: TBoolVector;
                                        const AShowPercentColumn: Boolean = False;
                                        const AShowLinePercentColumn: Boolean = False);
var
  ColWidths: TIntVector;
  ColCount, W, UsedReasonsCount: Integer;
begin
  UsedReasonsCount:= VCountIf(AUsedReasons, True);

  ColCount:= UsedReasonsCount * (1 + Ord(AShowPercentColumn) + Ord(AShowLinePercentColumn));
  W:= (TOTAL_WIDTH - PARAM_COLUMN_MIN_WIDTH) div ColCount;
  VDim(ColWidths{%H-}, ColCount+1{param names column}+2{graph margin columns}, W);
  ColWidths[0]:= MARGIN_COLUMN_WIDTH;
  ColWidths[High(ColWidths)]:= MARGIN_COLUMN_WIDTH;
  ColWidths[High(ColWidths)-1]:= W - MARGIN_COLUMN_WIDTH;
  W:= TOTAL_WIDTH - W*ColCount - MARGIN_COLUMN_WIDTH;
  ColWidths[1]:= W;

  inherited Create(AWorksheet, AGrid, AFont, ColWidths,
                   UsedReasonsCount, AUsedReasons, AShowPercentColumn);

  FShowLinePercentColumn:= AShowLinePercentColumn;
end;

procedure TStatisticSinglePeriodSheet.Draw(const ABeginDate, AEndDate: TDate;
  const AMotorNames: String;
  const AParamNames, AReasonNames: TStrVector;
  const ACounts: TIntMatrix;
  const ACalcTotalCountForUsedParamsOnly,
        AShowGraphics: Boolean);
begin
  if (FUsedReasonsCount=0) or VIsNil(AParamNames) or MIsNil(ACounts) then Exit;

  FBeginDate:= ABeginDate;
  FEndDate:= AEndDate;
  FMotorNames:= AMotorNames;
  FParamNames:= AParamNames;
  FReasonNames:= AReasonNames;
  FCounts:= ACounts;
  FCalcTotalCountForUsedParamsOnly:= ACalcTotalCountForUsedParamsOnly;
  FShowGraphics:= AShowGraphics;
  FWriter.BeginEdit;
  DrawReport;
  FWriter.EndEdit;
end;

{ TStatisticSinglePeriodAtMotorNamesSheet }

procedure TStatisticSinglePeriodAtMotorNamesSheet.DrawReport;
begin
  DrawParamsReport('наименованиям электродвигателей', 'Наименование');
end;

{ TStatisticSinglePeriodAtDefectNamesSheet }

procedure TStatisticSinglePeriodAtDefectNamesSheet.DrawReport;
begin
  DrawParamsReport('неисправным элементам', 'Неисправный элемент');
end;

{ TStatisticSinglePeriodAtPlaceNamesSheet }

procedure TStatisticSinglePeriodAtPlaceNamesSheet.DrawReport;
begin
  DrawParamsReport('предприятиям', 'Предприятие');
end;

{ TStatisticSinglePeriodAtMonthNamesSheet }

procedure TStatisticSinglePeriodAtMonthNamesSheet.DrawReport;
var
  R: Integer;
  UsedParams: TBoolVector;
  TotalCount: Integer;
begin
  R:= 1;
  DrawReportTitle(R, 'Отчет по рекламационным случаям электродвигателей');
  DrawSinglePeriodDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                   'месяцам',
                   'Месяц', 'Всего за месяц');

  UsedParams:= CalcUsedParams(FCounts, True {показывать параметры, где все нули});
  if VCountIf(UsedParams, True)=0 then Exit;

  TotalCount:= CalcTotalCount(FCounts, UsedParams); //от этого считаем процент
  DrawSinglePeriodDataTableValues(R, FCounts, UsedParams, TotalCount, True {показать итого}, haCenter);

  if FShowGraphics then
  begin
    DrawGraphForTotalCounts(R, 'Распределение общего количества рекламационных случаев по '+
                            'месяцам',
                            EmptyStr, EmptyStr,
                            FParamNames, FCounts, nil {нет сортировки}, UsedParams,
                            dtVertical, taCenter);

    DrawGraphForSumReasonCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей',
                                EmptyStr, EmptyStr,
                                FCounts, UsedParams, True{сортировка},
                                dtHorizontal, taCenter,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});

    DrawGraphForReasonCounts(R, 'Распределение количества рекламационных случаев по ' +
                                'месяцам' +
                                ' и причинам возникновения неисправностей',
                             EmptyStr, EmptyStr,
                             FParamNames, FCounts, nil {нет сортировки}, UsedParams,
                             dtVertical, taCenter);
    DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  end;
end;

{ TStatisticSinglePeriodAtMonthNamesSumSheet }

procedure TStatisticSinglePeriodAtMonthNamesSumSheet.DrawReport;
var
  R: Integer;
  UsedParams: TBoolVector;
  TotalCount: Integer;
  AccumCounts: TIntMatrix;
begin
  R:= 1;
  DrawReportTitle(R, 'Отчет по рекламационным случаям электродвигателей');

  UsedParams:= CalcUsedParams(FCounts, True {показывать параметры, где все нули});
  if VCountIf(UsedParams, True)=0 then Exit;

  TotalCount:= CalcTotalCount(FCounts, UsedParams); //от этого считаем процент

  AccumCounts:= CalcAccumCounts(FCounts);
  DrawSinglePeriodDataTableTop(R, 'Накопление рекламационных случаев по месяцам ',
                  'Месяц', 'Количество с накоплением');
  DrawSinglePeriodDataTableValues(R, AccumCounts, UsedParams, TotalCount, False {не показывать итого}, haCenter);

  if FShowGraphics then
  begin
    DrawLinesForTotalCounts(R, 'Накопление общего количества рекламационных случаев по '+
                            'месяцам',
                            EmptyStr, EmptyStr,
                            FParamNames, AccumCounts, nil {нет сортировки}, UsedParams,
                            taCenter);
    DrawGraphForSumReasonCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей',
                                EmptyStr, EmptyStr,
                                FCounts, UsedParams, True{сортировка},
                                dtHorizontal, taCenter,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});

    DrawLinesForReasonCounts(R, 'Накопление количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей',
                             EmptyStr, EmptyStr,
                             FParamNames, AccumCounts, nil {нет сортировки}, UsedParams,
                             taCenter);

    DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  end;
end;

{ TStatisticSinglePeriodAtMileagesSheet }

procedure TStatisticSinglePeriodAtMileagesSheet.DrawReport;
var
  R: Integer;
  UsedParams: TBoolVector;
  TotalCount, n, i, j: Integer;
begin
  R:= 1;
  DrawReportTitle(R, 'Отчет по рекламационным случаям электродвигателей');
  DrawSinglePeriodDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                   'величине пробега локомотива',
                   'Пробег локомотива, тыс.км', 'Всего за период');

  UsedParams:= CalcUsedParams(FCounts, True {показывать параметры, где все нули});
  i:= VIndexOf(FParamNames, 'Не указан');
  n:= FCounts[0, i];
  if n>0 then
  begin
    n:= 0;
    for j:= 1 to High(FUsedReasons) do
      if FUsedReasons[j] then
        n:= n + FCounts[j, i];
  end;
  UsedParams[i]:= n>0;


  if VCountIf(UsedParams, True)=0 then Exit;

  TotalCount:= CalcTotalCount(FCounts, UsedParams); //от этого считаем процент
  DrawSinglePeriodDataTableValues(R, FCounts, UsedParams, TotalCount, True, haCenter);

  if FShowGraphics then
  begin
    DrawGraphForTotalCounts(R, 'Распределение общего количества рекламационных случаев по '+
                            'величине пробега локомотива',
                            EmptyStr, 'Пробег локомотива, тыс.км',
                            FParamNames, FCounts, nil {нет сортировки}, UsedParams,
                            dtVertical, taCenter);

    DrawGraphForSumReasonCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей',
                                EmptyStr, EmptyStr,
                                FCounts, UsedParams, True{сортировка},
                                dtHorizontal, taCenter,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});

    DrawGraphForReasonCounts(R, 'Распределение количества рекламационных случаев по ' +
                                'величине пробега локомотива' +
                                ' и причинам возникновения неисправностей',
                             EmptyStr, 'Пробег локомотива, тыс.км',
                             FParamNames, FCounts, nil {нет сортировки}, UsedParams,
                             dtVertical, taCenter);
  end;
end;

{ TStatisticSeveralPeriodsSheet }

function TStatisticSeveralPeriodsSheet.CalcSumCountMatrix(const ACounts: TIntMatrix3D): TIntMatrix;
var
  n, k, i, j, Size1, Size2: Integer;
begin
  Result:= nil;
  if MIsNil(ACounts) then Exit;

  Size1:= Length(ACounts[0]);
  if Size1=0 then Exit;
  Size2:= Length(ACounts[0, 0]);
  if Size2=0 then Exit;
  MDim(Result, Size1, Size2);

  for j:= 1 to Size1-1 do
    for i:= 0 to Size2-1 do
  begin
    n:= 0;
    for k:= 0 to High(ACounts) do
      n:= n + ACounts[k, j, i];
    Result[j, i]:= n;
  end;

  for i:= 0 to Size2-1 do
  begin
    n:= 0;
    for j:= 1 to Size1-1 do
      n:= n + Result[j, i];
    Result[0, i]:= n;
  end;
end;

procedure TStatisticSeveralPeriodsSheet.DrawReportTitle(var ARow: Integer;
  const ATitle: String);
var
  R, i: Integer;
  S: String;
begin

  FWriter.SetAlignment(haCenter, vaCenter);

  R:= ARow;
  S:= ATitle;//'Отчет по рекламационным случаям электродвигателей';
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);

  if not SEmpty(FMotorNames) then
  begin
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.WriteText(R, 1, R, FWriter.ColCount, FMotorNames, cbtNone, True, True);
  end;

  R:= R + 1;
  S:= 'за ';

  if SameDate(FBeginDate, FEndDate) then
    S:= S + FormatDateTime('dd ', FBeginDate) + MONTHSGEN[MonthOfDate(FBeginDate)]
  else
    S:= S + 'период с ' +
        FormatDateTime('dd ', FBeginDate) + MONTHSGEN[MonthOfDate(FBeginDate)] +
       ' по ' +
        FormatDateTime('dd ', FEndDate) + MONTHSGEN[MonthOfDate(FEndDate)];
  S:= S + ' ' + IntToStr(YearOfDate(FBeginDate)-FAdditionYearCount);
  for i:= FAdditionYearCount-1 downto 0 do
    S:= S + '/' + IntToStr(YearOfDate(FBeginDate)-i);
  S:= S + 'гг.';
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);

  DrawEmptyRow(R);

  ARow:= R + 1;
end;

procedure TStatisticSeveralPeriodsSheet.DrawSinglePeriodDataTableTop(
  var ARow: Integer; const ATableTitle, AParamColumnCaption,
  ASumCountColumnCaption: String);
var
  i, R1, R2, C1, C2, Count: Integer;

  procedure DrawCaption(const ACaption: String; const AIsLast: Boolean);
  begin
    C1:= C2 + 1;
    //C2:= C1 + Ord(FShowPercentColumn) + Ord(AIsLast);
    C2:= C1 + FAdditionYearCount*(1+Ord(FShowPercentColumn)) + Ord(FShowPercentColumn) + Ord(AIsLast);
    FWriter.WriteText(R1, C1, R1, C2, ACaption, cbtOuter, True, True);
  end;

begin
  R1:= ARow;

  if not SEmpty(ATableTitle) then
  begin
    FWriter.SetFont(FFontName, FFontSize+1, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R1, 1, R1, FWriter.ColCount, ATableTitle, cbtNone, True, True);
    Inc(R1);
  end;

  R2:= R1;
  C1:= 1;
  C2:= 2;
  FWriter.SetFont(FFontName, FFontSize-1, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R1, C1, R2, C2, AParamColumnCaption, cbtOuter, True, True);

  Count:= 0;
  for i:= 0 to High(FReasonNames) do
  begin
    if not FUsedReasons[i] then continue;
    Inc(Count);
    if i=0 then
      DrawCaption(ASumCountColumnCaption, Count=FUsedReasonsCount)
    else
      DrawCaption(FReasonNames[i], Count=FUsedReasonsCount);
  end;

  FWriter.DrawBorders(R1, 1, R2, FWriter.ColCount, cbtAll);
  ARow:= R2;
end;

procedure TStatisticSeveralPeriodsSheet.DrawSinglePeriodDataTableValues(var ARow: Integer;
                                 const ACounts: TIntMatrix;
                                 const AUsedParams: TBoolVector;
                                 const ATotalCount: Integer;
                                 const AShowResumeRow: Boolean;
                                 const AParamColumnHorAlignment: TsHorAlignment);
var
  i, R: Integer;

  procedure DrawParamName(var ARowNum, AColNum: Integer; const AName: String);
  var
    RR, C1, C2: Integer;
  begin
    RR:= ARowNum;
    C1:= AColNum;
    C2:= C1 + 1;
    FWriter.SetAlignment(AParamColumnHorAlignment, vaCenter);
    FWriter.WriteText(RR, C1, RR, C2, AName, cbtOuter, True, True);
    ARowNum:= RR;
    AColNum:= C2;
  end;

  procedure DrawParamValues(var ARowNum, AColNum: Integer;
                            const AValue: Integer; const AIsLast: Boolean);
  var
    RR, C1, C2: Integer;
    DD: Double;
  begin
    RR:= ARowNum;
    C1:= AColNum + 1;
    C2:= C1;
    C2:= C1 + FAdditionYearCount;
    if (not FShowPercentColumn) and AIsLast then
      C2:= C2 + 1;
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteNumber(RR, C1, RR, C2, AValue, cbtOuter);
    if FShowPercentColumn then
    begin
      C1:= C2 + 1;
      C2:= C1 + FAdditionYearCount + Ord(AIsLast);
      DD:= Part(AValue, ATotalCount);
      FWriter.WriteNumber(RR, C1, RR, C2, DD, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
    end;
    ARowNum:= RR;
    AColNum:= C2;
  end;

  procedure DrawDataRow(var ARowNum: Integer; const AParamName: String; const AIndex: Integer);
  var
    k, ReasonCount, RR, C, DataValue: Integer;
  begin
    if not AUsedParams[AIndex] then Exit;
    FWriter.SetFont(FFontName, FFontSize-1, [], clBlack);
    RR:= ARowNum + 1;
    C:= 1;
    DrawParamName(RR, C, AParamName);
    ReasonCount:= 0;
    for k:= 0 to High(FUsedReasons) do
    begin
      if not FUsedReasons[k] then continue;
      Inc(ReasonCount);

      if k=0 then
        DataValue:= CalcSumCountForParam(ACounts, AUsedParams, AIndex)
      else
        DataValue:= ACounts[k,AIndex];

      DrawParamValues(RR, C, DataValue, ReasonCount=FUsedReasonsCount);
    end;
    ARowNum:= RR;
  end;

  procedure DrawResumeRow(var ARowNum: Integer);
  var
    k, ReasonCount, RR, C, DataValue: Integer;
  begin
    if not AShowResumeRow then Exit;
    if (VCountIf(AUsedParams, True)<=1) and FCalcTotalCountForUsedParamsOnly then Exit;
    FWriter.SetFont(FFontName, FFontSize-1, [fsBold], clBlack);
    RR:= ARowNum + 1;
    C:= 1;
    DrawParamName(RR, C, 'ИТОГО');
    ReasonCount:= 0;
    for k:= 0 to High(FUsedReasons) do
    begin
      if not FUsedReasons[k] then continue;
      Inc(ReasonCount);
      if k=0 then
        DataValue:= ATotalCount
      else
        DataValue:= CalcSumCountForReason(ACounts[k], AUsedParams);
      DrawParamValues(RR, C, DataValue, ReasonCount=FUsedReasonsCount);
    end;
    ARowNum:= RR;
  end;

begin
  R:= ARow;
  for i:= 0 to High(FParamNames) do
    DrawDataRow(R, FParamNames[i], i);
  DrawResumeRow(R);
  DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  ARow:= R;
end;

procedure TStatisticSeveralPeriodsSheet.DrawSeveralPeriodsDataTableTop(var ARow: Integer;
  const ATableTitle, AParamColumnCaption, ASumCountColumnCaption: String);
var
  i, R1, R2, C1, C2, Count: Integer;

  procedure DrawCaption(const ACaption: String; const AIsLast: Boolean);
  var
    j: Integer;
  begin
    C1:= C2 + 1;
    C2:= C1 + (1+Ord(FShowPercentColumn))*(FAdditionYearCount+1) - 1 + Ord(AIsLast);
    FWriter.WriteText(R1, C1, R1, C2, ACaption, cbtOuter, True, True);
    C2:= C1-1;
    for j:= 0 to FAdditionYearCount do
    begin
      C1:= C2 + 1;
      C2:= C1 + Ord(FShowPercentColumn);
      if (j=FAdditionYearCount) and AIsLast then
        C2:= C2 + 1;
      FWriter.WriteNumber(R2, C1, R2, C2, YearOfDate(FBeginDate)-FAdditionYearCount+j);
    end;
  end;

begin
  R1:= ARow;

  if not SEmpty(ATableTitle) then
  begin
    FWriter.SetFont(FFontName, FFontSize+1, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R1, 1, R1, FWriter.ColCount, ATableTitle, cbtNone, True, True);
    Inc(R1);
  end;

  R2:= R1 + 1;
  C1:= 1;
  C2:= 2;
  FWriter.SetFont(FFontName, FFontSize-1, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R1, C1, R2, C2, AParamColumnCaption, cbtOuter, True, True);

  Count:= 0;
  for i:= 0 to High(FReasonNames) do
  begin
    if not FUsedReasons[i] then continue;
    Inc(Count);
    if i=0 then
      DrawCaption(ASumCountColumnCaption, Count=FUsedReasonsCount)
    else
      DrawCaption(FReasonNames[i], Count=FUsedReasonsCount);
  end;

  FWriter.DrawBorders(R1, 1, R2, FWriter.ColCount, cbtAll);
  ARow:= R2;
end;

procedure TStatisticSeveralPeriodsSheet.DrawSeveralPeriodsDataTableValues(var ARow: Integer;
                                 const ACounts: TIntMatrix3D;
                                 const AUsedParams: TBoolVector;
                                 const ATotalCounts: TIntVector;
                                 const AShowResumeRow: Boolean;
                                 const AParamColumnHorAlignment: TsHorAlignment);
var
  i, R: Integer;

  procedure DrawParamName(var ARowNum, AColNum: Integer; const AName: String);
  var
    RR, C1, C2: Integer;
  begin
    RR:= ARowNum;
    C1:= AColNum;
    C2:= C1 + 1;
    FWriter.SetAlignment(AParamColumnHorAlignment, vaCenter);
    FWriter.WriteText(RR, C1, RR, C2, AName, cbtOuter, True, True);
    ARowNum:= RR;
    AColNum:= C2;
  end;

  procedure DrawParamValues(var ARowNum, AColNum: Integer;
                            const AValue, APeriodIndex: Integer; const AIsLast: Boolean);
  var
    RR, C1, C2: Integer;
    DD: Double;
  begin
    RR:= ARowNum;
    C1:= AColNum + 1;
    C2:= C1;
    if (not FShowPercentColumn) and AIsLast then
      C2:= C2 + 1;
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteNumber(RR, C1, RR, C2, AValue, cbtOuter);
    if FShowPercentColumn then
    begin
      C1:= C2 + 1;
      C2:= C1 + Ord(AIsLast);
      DD:= Part(AValue, ATotalCounts[APeriodIndex]);
      FWriter.WriteNumber(RR, C1, RR, C2, DD, PERCENT_FRAC_DIGITS, cbtOuter, nfPercentage);
    end;
    ARowNum:= RR;
    AColNum:= C2;
  end;

  procedure DrawDataRow(var ARowNum: Integer; const AParamName: String; const AIndex: Integer);
  var
    j, k, n, ReasonCount, RR, C, DataValue: Integer;
  begin
    if not AUsedParams[AIndex] then Exit;
    FWriter.SetFont(FFontName, FFontSize-1, [], clBlack);
    RR:= ARowNum + 1;
    C:= 1;
    DrawParamName(RR, C, AParamName);
    ReasonCount:= 0;
    for j:= 0 to High(FUsedReasons) do
    begin
      if not FUsedReasons[j] then continue;
      Inc(ReasonCount);
      n:= High(ACounts);
      for k:= 0 to n do
      begin
        if j=0 then
          DataValue:= CalcSumCountForParam(ACounts[k], AUsedParams, AIndex)
        else
          DataValue:= ACounts[k, j, AIndex];
        DrawParamValues(RR, C, DataValue, k, (ReasonCount=FUsedReasonsCount) and (k=n));
      end;
    end;
    ARowNum:= RR;
  end;

  procedure DrawResumeRow(var ARowNum: Integer);
  var
    j, k, n, ReasonCount, RR, C, DataValue: Integer;
  begin
    if not AShowResumeRow then Exit;
    if (VCountIf(AUsedParams, True)<=1) and FCalcTotalCountForUsedParamsOnly then Exit;
    FWriter.SetFont(FFontName, FFontSize-1, [fsBold], clBlack);
    RR:= ARowNum + 1;
    C:= 1;
    DrawParamName(RR, C, 'ИТОГО');
    ReasonCount:= 0;
    for j:= 0 to High(FUsedReasons) do
    begin
      if not FUsedReasons[j] then continue;
      Inc(ReasonCount);
      n:= High(ACounts);
      for k:= 0 to n do
      begin
        if j=0 then
          DataValue:= ATotalCounts[k]
        else
          DataValue:= CalcSumCountForReason(ACounts[k,j], AUsedParams);
        DrawParamValues(RR, C, DataValue, k, (ReasonCount=FUsedReasonsCount) and (k=n));
      end;
    end;
    ARowNum:= RR;
  end;

begin
  R:= ARow;
  for i:= 0 to High(FParamNames) do
    DrawDataRow(R, FParamNames[i], i);
  DrawResumeRow(R);
  DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  ARow:= R;
end;

procedure TStatisticSeveralPeriodsSheet.DrawLinesForTotalCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix3D;
                                       const ASortIndexes: TIntVector;
                                       const AUsedParams: TBoolVector;
                                       const AAlignment: TAlignment);
var
  i, j, n: Integer;
  ParamNames: TStrVector;
  ParamValues: TIntVector;
  UsedParams: TBoolVector;
  Legend: TStrVector;
  DataMatrix: TIntMatrix;
begin

  if VIsNil(ASortIndexes) then
  begin
    UsedParams:= AUsedParams;
    ParamNames:= AParamNames;
  end
  else begin
    UsedParams:= VReplace(AUsedParams, ASortIndexes);
    ParamNames:= VReplace(AParamNames, ASortIndexes);
  end;

  ParamNames:= VCut(ParamNames, UsedParams);
  if Length(ParamNames)<2 then Exit;

  Legend:= nil;
  n:= YearOf(FBeginDate);
  for i:= FAdditionYearCount downto 0 do
    VAppend(Legend, IntToStr(n-i));

  DataMatrix:= nil;
  for i:= 0 to FAdditionYearCount do
  begin
    ParamValues:= nil;
    for j:= 0 to High(ParamNames) do
    begin
      if not VIsNil(ASortIndexes) then
        n:= ASortIndexes[j]
      else
        n:= j;
      VAppend(ParamValues, CalcSumCountForParam(AParamValues[i], AUsedParams{ no sort!!}, n));
    end;
    MAppend(DataMatrix, ParamValues);
  end;

  DrawLinesForMatrix(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, Legend, DataMatrix, AAlignment, ASortIndexes);
end;

procedure TStatisticSeveralPeriodsSheet.DrawGraphForTotalCountsInTime(var ARow: Integer;
                           const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                           const AParamNames: TStrVector;
                           const AParamValues: TIntMatrix3D;
                           const ASortIndexes: TIntVector;
                           const AUsedParams: TBoolVector;
                           const AHistType: TDirectionType;
                           const AAlignment: TAlignment;
                           const AOneColorPerTick: Boolean = False);
var
  i, j, n: Integer;
  ParamNames: TStrVector;
  ParamValues, ColorIndexes: TIntVector;
  UsedParams: TBoolVector;
  Legend: TStrVector;
  DataMatrix: TIntMatrix;
begin

  if VIsNil(ASortIndexes) then
  begin
    UsedParams:= AUsedParams;
    ParamNames:= AParamNames;
  end
  else begin
    UsedParams:= VReplace(AUsedParams, ASortIndexes);
    ParamNames:= VReplace(AParamNames, ASortIndexes);
  end;

  ParamNames:= VCut(ParamNames, UsedParams);
  if Length(ParamNames)<2 then Exit;

  Legend:= nil;
  n:= YearOf(FBeginDate);
  for i:= FAdditionYearCount downto 0 do
    VAppend(Legend, IntToStr(n-i));

  DataMatrix:= nil;
  for i:= 0 to FAdditionYearCount do
  begin
    ParamValues:= nil;
    for j:= 0 to High(ParamNames) do
    begin
      if not VIsNil(ASortIndexes) then
        n:= ASortIndexes[j]
      else
        n:= j;
      VAppend(ParamValues, CalcSumCountForParam(AParamValues[i], AUsedParams{ no sort!!}, n));
    end;
    MAppend(DataMatrix, ParamValues);
  end;

  ColorIndexes:= nil;
  if AOneColorPerTick then
    ColorIndexes:= ASortIndexes;

  DrawGraphForMatrix(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, Legend, DataMatrix, AHistType, AAlignment,
                     AOneColorPerTick, ColorIndexes);
end;

procedure TStatisticSeveralPeriodsSheet.DrawGraphForSumReasonCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const ACounts: TIntMatrix3D;
                                       const AUsedParams: TBoolVector;
                                       const ANeedOrder: Boolean;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);
var
  i, j: Integer;
  ParamNames: TStrVector;
  ParamValues, Indexes, ColorIndexes: TIntVector;
  UsedReasons: TBoolVector;
  Legend: TStrVector;
  DataMatrix: TIntMatrix;
begin
  ParamNames:= VCut(FReasonNames, 1);
  UsedReasons:= VCut(FUsedReasons, 1);

  Indexes:= nil;
  if ANeedOrder then
  begin
    ParamValues:= nil;
    for i:= 1 to High(FReasonNames) do
      VAppend(ParamValues, CalcSumCountForReason(FSumCounts[i], AUsedParams));
    VSort(ParamValues, Indexes);
    ParamNames:= VReplace(ParamNames, Indexes);
    UsedReasons:= VReplace(UsedReasons, Indexes);
  end;
  ParamNames:= VCut(ParamNames, UsedReasons);
  if Length(ParamNames)<2 then Exit;


  Legend:= nil;
  j:= YearOf(FBeginDate);
  for i:= FAdditionYearCount downto 0 do
    VAppend(Legend, IntToStr(j-i));

  DataMatrix:= nil;
  for i:= 0 to FAdditionYearCount do
  begin
    ParamValues:= nil;
    for j:= 1 to High(FReasonNames) do
      VAppend(ParamValues, CalcSumCountForReason(ACounts[i,j], AUsedParams));
    if ANeedOrder then
      ParamValues:= VReplace(ParamValues, Indexes);
    ParamValues:= VCut(ParamValues, UsedReasons);
    MAppend(DataMatrix, ParamValues);
  end;


  ColorIndexes:= nil;
  if AOneColorPerTick then
    ColorIndexes:= Indexes;

  DrawGraphForMatrix(ARow, AGraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, Legend, DataMatrix, AHistType, AAlignment,
                     AOneColorPerTick, ColorIndexes);
end;

procedure TStatisticSeveralPeriodsSheet.DrawLinesForParamCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix3D;
                                       const AUsedParams: TBoolVector;
                                       const AAlignment: TAlignment);
var
  i, j: Integer;
  ParamNames: TStrVector;
  Legend: TStrVector;
  DataMatrix: TIntMatrix;
  GraphTitle, S: String;
  UsedYears: TBoolVector;
begin
  ParamNames:= VCut(AParamNames, AUsedParams);

  i:= YearOf(FBeginDate);
  S:= IntToStr(i-FAdditionYearCount) + '-' + IntToStr(i) + 'гг.';

  Legend:= nil;
  j:= YearOf(FBeginDate);
  for i:= FAdditionYearCount downto 0 do
    VAppend(Legend, IntToStr(j-i));
  VDim(UsedYears{%H-}, FAdditionYearCount+1, True);


  for i:= 2 to High(FUsedReasons) do
  begin
    if not FUsedReasons[i] then continue;

    DataMatrix:= MIndex1(AParamValues, i, UsedYears, AUsedParams);
    GraphTitle:= AGraphTitle + ' (' + SLower(FReasonNames[i]) + ') за ' + S;
    DrawLinesForMatrix(ARow, GraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, Legend, DataMatrix, AAlignment);
  end;
end;

procedure TStatisticSeveralPeriodsSheet.DrawGraphsForParamCountsInTime(var ARow: Integer;
                                       const AGraphTitle, AVertAxisTitle, AHorizAxisTitle: String;
                                       const AParamNames: TStrVector;
                                       const AParamValues: TIntMatrix3D;
                                       const AUsedParams: TBoolVector;
                                       const AHistType: TDirectionType;
                                       const AAlignment: TAlignment;
                                       const AOneColorPerTick: Boolean = False);
var
  i, j: Integer;
  ParamNames: TStrVector;
  Legend: TStrVector;
  DataMatrix: TIntMatrix;
  GraphTitle, S: String;
  UsedYears: TBoolVector;
begin
  ParamNames:= VCut(AParamNames, AUsedParams);

  i:= YearOf(FBeginDate);
  S:= IntToStr(i-FAdditionYearCount) + '-' + IntToStr(i) + 'гг.';

  Legend:= nil;
  j:= YearOf(FBeginDate);
  for i:= FAdditionYearCount downto 0 do
    VAppend(Legend, IntToStr(j-i));
  VDim(UsedYears{%H-}, FAdditionYearCount+1, True);


  for i:= 2 to High(FUsedReasons) do
  begin
    if not FUsedReasons[i] then continue;

    DataMatrix:= MIndex1(AParamValues, i, UsedYears, AUsedParams);
    GraphTitle:= AGraphTitle + ' (' + SLower(FReasonNames[i]) + ') за ' + S;
    DrawGraphForMatrix(ARow, GraphTitle, AVertAxisTitle, AHorizAxisTitle,
                     ParamNames, Legend, DataMatrix, AHistType, AAlignment,
                     AOneColorPerTick);
  end;
end;

procedure TStatisticSeveralPeriodsSheet.DrawParamsReport(
   const ATableTitlePart, AParamColumnCaption: String);
var
  R, R1, R2, i: Integer;
  UsedParams: TBoolVector;
  Indexes: TIntVector;
  TotalCount: Integer;
  TotalCounts: TIntVector;
  S: String;
begin
  R:= 1;
  DrawReportTitle(R, 'Отчет по рекламационным случаям электродвигателей');

  VSort(FSumCounts[0], Indexes); //в гистограммах сортировка по сумме случаев за период


  UsedParams:= CalcUsedParams(FSumCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)>0 then
  begin
    i:= YearOf(FBeginDate);
    S:= IntToStr(i-FAdditionYearCount) + '-' + IntToStr(i) + 'гг.';
    DrawSinglePeriodDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                   ATableTitlePart, AParamColumnCaption, 'Всего за ' + S);
    TotalCount:= CalcTotalCount(FSumCounts, UsedParams); //от этого считаем процент
    DrawSinglePeriodDataTableValues(R, FSumCounts, UsedParams, TotalCount, True, haLeft);

    if FShowGraphics then
    begin
      R1:= R;
      DrawGraphForReasonCounts(R, 'Распределение количества рекламационных случаев по ' +
                                ATableTitlePart + ' и '+
                                'причинам возникновения неисправностей за ' + S,
                             EmptyStr, EmptyStr,
                             FParamNames, FSumCounts, Indexes, UsedParams,
                             dtHorizontal, taRightJustify);
      R2:= R;
      R:= R1;
      DrawGraphForTotalCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                ATableTitlePart + ' за ' + S,
                            EmptyStr, EmptyStr,
                            FParamNames, FSumCounts, Indexes, UsedParams,
                            dtHorizontal, taLeftJustify);

      if FWriter.HasGrid and ((R-R1)>30) then
      begin
        DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
        //DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
      end;

      DrawGraphForSumReasonCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                                EmptyStr, EmptyStr,
                                FSumCounts, UsedParams, True{сортировка},
                                dtHorizontal, taLeftJustify,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});
      R1:= Min(R, R2);
      R2:= Max(R, R2);

      for i:= R1 to R2+2 do
      begin
        R:= i-1;
        DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
      end;

      if FWriter.HasGrid then
      begin
        DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
        //DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
      end;
    end
    else //do not show graphics
      DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  end;

  UsedParams:= CalcUsedParams(FCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)=0 then Exit;

  VDim(TotalCounts{%H-}, Length(FCounts));
  for i:= 0 to High(TotalCounts) do
    TotalCounts[i]:= CalcTotalCount(FCounts[i], UsedParams); //от этого считаем процент в периоде

  DrawSeveralPeriodsDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                                     ATableTitlePart + ' и временным периодам',
                                     AParamColumnCaption, 'Всего за период');
  DrawSeveralPeriodsDataTableValues(R, FCounts, UsedParams, TotalCounts, True, haLeft);

  if FShowGraphics then
  begin
    R1:= R;
    DrawGraphForTotalCountsInTime(R,'Распределение количества рекламационных случаев по ' +
                                    ATableTitlePart + ' за ' + S,
                                  EmptyStr, EmptyStr,
                                  FParamNames, FCounts, Indexes, UsedParams,
                                  dtHorizontal, taLeftJustify);
    R2:= R;
    if FWriter.HasGrid and ((R2-R1)>30) then
    begin
      DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
      //DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
    end;
    DrawGraphForSumReasonCountsInTime(R, 'Распределение количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                                EmptyStr, EmptyStr,
                                FCounts, UsedParams, True{сортировка},
                                dtHorizontal, taLeftJustify,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});

    R:= R1;
    DrawGraphsForParamCountsInTime(R, 'Распределение количества рекламационных случаев по ' +
                                       ATableTitlePart, EmptyStr, EmptyStr,
                                   FParamNames, FCounts, UsedParams,
                                   dtHorizontal, taRightJustify);
  end;
end;

constructor TStatisticSeveralPeriodsSheet.Create(const AWorksheet: TsWorksheet;
                       const AGrid: TsWorksheetGrid;
                       const AFont: TFont; //if nil -> SHEET_FONT_NAME, SHEET_FONT_SIZE
                       const AAdditionYearCount: Integer;
                       const AUsedReasons: TBoolVector;
                       const AShowPercentColumn: Boolean = False);
var
  ColWidths: TIntVector;
  ColCount, W, UsedReasonsCount, TotalWidth: Integer;
begin
  case AAdditionYearCount of
  1: TotalWidth:= 1500 {+ 300*Ord(AShowPercentColumn)};
  2: TotalWidth:= 1800 + 200*Ord(AShowPercentColumn);
  3: TotalWidth:= 1500 + 1000*Ord(AShowPercentColumn);
  4: TotalWidth:= 1600 + 1400*Ord(AShowPercentColumn);
  end;

  UsedReasonsCount:= VCountIf(AUsedReasons, True);
  ColCount:= UsedReasonsCount * (1 + Ord(AShowPercentColumn)) * (AAdditionYearCount + 1);
  W:= (TotalWidth - PARAM_COLUMN_MIN_WIDTH) div ColCount;

  VDim(ColWidths{%H-}, ColCount+1{param names column}+2{graph margin columns}, W);
  ColWidths[0]:= MARGIN_COLUMN_WIDTH;
  ColWidths[High(ColWidths)]:= MARGIN_COLUMN_WIDTH;
  ColWidths[High(ColWidths)-1]:= W - MARGIN_COLUMN_WIDTH;
  W:= TotalWidth - W*ColCount - MARGIN_COLUMN_WIDTH;
  ColWidths[1]:= W;

  inherited Create(AWorksheet, AGrid, AFont, ColWidths,
                   UsedReasonsCount, AUsedReasons, AShowPercentColumn);
  FAdditionYearCount:= AAdditionYearCount;
end;

procedure TStatisticSeveralPeriodsSheet.Draw(const ABeginDate, AEndDate: TDate;
                   const AMotorNames: String;
                   const AParamNames, AReasonNames: TStrVector;
                   const ACounts: TIntMatrix3D;
                   const ACalcTotalCountForUsedParamsOnly,
                         AShowGraphics: Boolean);
begin
  if (FUsedReasonsCount=0) or VIsNil(AParamNames) or MIsNil(ACounts) then Exit;

  FBeginDate:= ABeginDate;
  FEndDate:= AEndDate;
  FMotorNames:= AMotorNames;
  FParamNames:= AParamNames;
  FReasonNames:= AReasonNames;
  FCounts:= ACounts;
  FSumCounts:= CalcSumCountMatrix(ACounts);
  FCalcTotalCountForUsedParamsOnly:= ACalcTotalCountForUsedParamsOnly;
  FShowGraphics:= AShowGraphics;
  FWriter.BeginEdit;
  DrawReport;
  FWriter.EndEdit;
end;

{ TStatisticSeveralPeriodsAtMotorNamesSheet }

procedure TStatisticSeveralPeriodsAtMotorNamesSheet.DrawReport;
begin
  DrawParamsReport('наименованиям электродвигателей', 'Наименование');
end;

{ TStatisticSeveralPeriodsAtDefectNamesSheet }

procedure TStatisticSeveralPeriodsAtDefectNamesSheet.DrawReport;
begin
  DrawParamsReport('неисправным элементам', 'Неисправный элемент');
end;

{ TStatisticSeveralPeriodsAtPlaceNamesSheet }

procedure TStatisticSeveralPeriodsAtPlaceNamesSheet.DrawReport;
begin
  DrawParamsReport('предприятиям', 'Предприятие');
end;

{ TStatisticSeveralPeriodsAtMonthNamesSheet }

procedure TStatisticSeveralPeriodsAtMonthNamesSheet.DrawReport;
var
  R, R1, R2, i: Integer;
  UsedParams: TBoolVector;
  TotalCount: Integer;
  TotalCounts: TIntVector;
  S: String;
begin
  R:= 1;
  DrawReportTitle(R, 'Отчет по рекламационным случаям электродвигателей');

  UsedParams:= CalcUsedParams(FSumCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)>0 then
  begin
    i:= YearOf(FBeginDate);
    S:= IntToStr(i-FAdditionYearCount) + '-' + IntToStr(i) + 'гг.';
    DrawSinglePeriodDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                   'месяцам', 'Месяц', 'Всего за ' + S);
    TotalCount:= CalcTotalCount(FSumCounts, UsedParams); //от этого считаем процент
    DrawSinglePeriodDataTableValues(R, FSumCounts, UsedParams, TotalCount, True, haCenter);

    if FShowGraphics then
    begin
      R1:= R;
      DrawGraphForReasonCounts(R, 'Распределение количества рекламационных случаев по ' +
                                ' месяцам и '+
                                'причинам возникновения неисправностей за ' + S,
                             EmptyStr, EmptyStr,
                             FParamNames, FSumCounts, nil{no sort}, UsedParams,
                             dtVertical, taRightJustify);
      R2:= R;
      R:= R1;
      DrawGraphForTotalCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'месяцам за ' + S,
                            EmptyStr, EmptyStr,
                            FParamNames, FSumCounts, nil{no sort}, UsedParams,
                            dtVertical, taLeftJustify);
      DrawGraphForSumReasonCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                                EmptyStr, EmptyStr,
                                FSumCounts, UsedParams, True{сортировка},
                                dtHorizontal, taLeftJustify,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});
      R1:= Min(R, R2);
      R2:= Max(R, R2);

      for i:= R1 to R2+2 do
      begin
        R:= i-1;
        DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
      end;
    end
    else //do not show graphics
      DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  end;

  UsedParams:= CalcUsedParams(FCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)=0 then Exit;

  VDim(TotalCounts{%H-}, Length(FCounts));
  for i:= 0 to High(TotalCounts) do
    TotalCounts[i]:= CalcTotalCount(FCounts[i], UsedParams); //от этого считаем процент в периоде

  DrawSeveralPeriodsDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                                     'временным периодам',
                                     'Месяц', 'Всего за период');
  DrawSeveralPeriodsDataTableValues(R, FCounts, UsedParams, TotalCounts, True, haCenter);

  if FShowGraphics then
  begin
    R1:= R;
    DrawGraphForTotalCountsInTime(R,'Распределение количества рекламационных случаев по ' +
                                    'месяцам за ' + S,
                                  EmptyStr, EmptyStr,
                                  FParamNames, FCounts, nil{no sort}, UsedParams,
                                  dtVertical, taLeftJustify);

    DrawGraphForSumReasonCountsInTime(R, 'Распределение количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                                EmptyStr, EmptyStr,
                                FCounts, UsedParams, True{сортировка},
                                dtHorizontal, taLeftJustify,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});

    R2:= R;
    R:= R1;
    DrawGraphsForParamCountsInTime(R, 'Распределение количества рекламационных случаев по ' +
                                       'месяцам', EmptyStr, EmptyStr,
                                   FParamNames, FCounts, UsedParams,
                                   dtVertical, taRightJustify);

    R:= Max(R, R2);
    DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  end;
end;

{ TStatisticSeveralPeriodsAtMonthNamesSumSheet }

procedure TStatisticSeveralPeriodsAtMonthNamesSumSheet.DrawReport;
var
  R, R1, R2, i: Integer;
  UsedParams: TBoolVector;
  TotalCount: Integer;
  TotalCounts: TIntVector;
  S: String;
  AccumCounts: TIntMatrix3D;
  SumAccumCounts: TIntMatrix;
begin
  R:= 1;
  DrawReportTitle(R, 'Отчет по рекламационным случаям электродвигателей');

  SumAccumCounts:= CalcAccumCounts(FSumCounts);
  AccumCounts:= CalcAccumCounts(FCounts);

  UsedParams:= CalcUsedParams(SumAccumCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)>0 then
  begin
    i:= YearOf(FBeginDate);
    S:= IntToStr(i-FAdditionYearCount) + '-' + IntToStr(i) + 'гг.';
    DrawSinglePeriodDataTableTop(R, 'Накопление количества рекламационных случаев по ' +
                   'месяцам', 'Месяц', 'Всего за ' + S);
    TotalCount:= CalcTotalCount(SumAccumCounts, UsedParams); //от этого считаем процент
    DrawSinglePeriodDataTableValues(R, SumAccumCounts, UsedParams, TotalCount, False, haCenter);

    if FShowGraphics then
    begin
      R1:= R;
      DrawLinesForReasonCounts(R, 'Накопление количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                             EmptyStr, EmptyStr,
                             FParamNames, SumAccumCounts, nil {нет сортировки}, UsedParams,
                             taRightJustify);

      R2:= R;
      R:= R1;
      DrawLinesForTotalCounts(R, 'Накопление общего количества рекламационных случаев по '+
                            'месяцам за ' + S,
                            EmptyStr, EmptyStr,
                            FParamNames, SumAccumCounts, nil {нет сортировки}, UsedParams,
                            taLeftJustify);

      DrawGraphForSumReasonCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                                EmptyStr, EmptyStr,
                                FSumCounts, UsedParams, True{сортировка},
                                dtHorizontal, taLeftJustify,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});

      R1:= Min(R, R2);
      R2:= Max(R, R2);

      for i:= R1 to R2+2 do
      begin
        R:= i-1;
        DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
      end;
    end
    else //do not show graphics
      DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  end;

  UsedParams:= CalcUsedParams(AccumCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)=0 then Exit;

  VDim(TotalCounts{%H-}, Length(FCounts));
  for i:= 0 to High(TotalCounts) do
    TotalCounts[i]:= CalcTotalCount(AccumCounts[i], UsedParams); //от этого считаем процент в периоде

  DrawSeveralPeriodsDataTableTop(R, 'Распределение накопленного количества рекламационных случаев по ' +
                                     'временным периодам',
                                     'Месяц', 'Всего за период');
  DrawSeveralPeriodsDataTableValues(R, AccumCounts, UsedParams, TotalCounts, False, haCenter);

  if FShowGraphics then
  begin
    R1:= R;
    DrawLinesForTotalCountsInTime(R,'Накопление количества рекламационных случаев по ' +
                                    'месяцам за ' + S,
                                  EmptyStr, EmptyStr,
                                  FParamNames, AccumCounts, nil{no sort}, UsedParams,
                                  taLeftJustify);
    DrawGraphForSumReasonCountsInTime(R, 'Распределение количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                                EmptyStr, EmptyStr,
                                FCounts, UsedParams, True{сортировка},
                                dtHorizontal, taLeftJustify,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});

    R:= R1;
    DrawLinesForParamCountsInTime(R, 'Накопление количества рекламационных случаев по ' +
                                       'месяцам', EmptyStr, EmptyStr,
                                   FParamNames, AccumCounts, UsedParams,
                                   taRightJustify);
    DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  end;
end;

{ TStatisticSeveralPeriodsAtMileagesSheet }

procedure TStatisticSeveralPeriodsAtMileagesSheet.DrawReport;
var
  R, R1, R2, i: Integer;
  UsedParams: TBoolVector;
  TotalCount: Integer;
  TotalCounts: TIntVector;
  S: String;
begin
  R:= 1;
  DrawReportTitle(R, 'Отчет по рекламационным случаям электродвигателей');

  UsedParams:= CalcUsedParams(FSumCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)>0 then
  begin
    i:= YearOf(FBeginDate);
    S:= IntToStr(i-FAdditionYearCount) + '-' + IntToStr(i) + 'гг.';
    DrawSinglePeriodDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                                 'величине пробега локомотива',
                                 'Пробег локомотива, тыс.км', 'Всего за ' + S);
    TotalCount:= CalcTotalCount(FSumCounts, UsedParams); //от этого считаем процент
    DrawSinglePeriodDataTableValues(R, FSumCounts, UsedParams, TotalCount, True, haCenter);

    if FShowGraphics then
    begin
      R1:= R;
      DrawGraphForReasonCounts(R, 'Распределение количества рекламационных случаев по ' +
                                'величине пробега локомотива и '+
                                'причинам возникновения неисправностей за ' + S,
                             EmptyStr, 'Пробег локомотива, тыс.км',
                             FParamNames, FSumCounts, nil{no sort}, UsedParams,
                             dtVertical, taRightJustify);
      R2:= R;
      R:= R1;
      DrawGraphForTotalCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'месяцам за ' + S,
                            EmptyStr, 'Пробег локомотива, тыс.км',
                            FParamNames, FSumCounts, nil{no sort}, UsedParams,
                            dtVertical, taLeftJustify);
      DrawGraphForSumReasonCounts(R, 'Распределение общего количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                                EmptyStr, EmptyStr,
                                FSumCounts, UsedParams, True{сортировка},
                                dtHorizontal, taLeftJustify,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});
      R1:= Min(R, R2);
      R2:= Max(R, R2);

      for i:= R1 to R2+2 do
      begin
        R:= i-1;
        DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
      end;
    end
    else //do not show graphics
      DrawEmptyRow(R, ROW_HEIGHT_DEFAULT);
  end;

  UsedParams:= CalcUsedParams(FCounts, False {не показывать параметры, где все нули});
  if VCountIf(UsedParams, True)=0 then Exit;

  VDim(TotalCounts{%H-}, Length(FCounts));
  for i:= 0 to High(TotalCounts) do
    TotalCounts[i]:= CalcTotalCount(FCounts[i], UsedParams); //от этого считаем процент в периоде

  DrawSeveralPeriodsDataTableTop(R, 'Распределение количества рекламационных случаев по ' +
                                     'величине пробега локомотива',
                                     'Пробег локомотива, тыс.км', 'Всего за период');
  DrawSeveralPeriodsDataTableValues(R, FCounts, UsedParams, TotalCounts, True, haCenter);


  if FShowGraphics then
  begin
    R1:= R;
    DrawGraphForTotalCountsInTime(R,'Распределение количества рекламационных случаев по ' +
                                    'величине пробега локомотива за ' + S,
                                  EmptyStr, 'Пробег локомотива, тыс.км',
                                  FParamNames, FCounts, nil{no sort}, UsedParams,
                                  dtVertical, taLeftJustify);

    DrawGraphForSumReasonCountsInTime(R, 'Распределение количества рекламационных случаев по ' +
                                'причинам возникновения неисправностей за ' + S,
                                EmptyStr, EmptyStr,
                                FCounts, UsedParams, True{сортировка},
                                dtHorizontal, taLeftJustify,
                                DIFFERENT_COLORS_FOR_EACH_REASON {отдельный цвет для каждой причины});


    R:= R1;
    DrawGraphsForParamCountsInTime(R, 'Распределение количества рекламационных случаев по ' +
                                       'величине пробега локомотива',
                                       EmptyStr, 'Пробег локомотива, тыс.км',
                                   FParamNames, FCounts, UsedParams,
                                   dtVertical, taRightJustify);
  end;
end;

{ Utils }

function LinesForDataVector(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames: TStrVector;
                               const ADataValues: TIntVector;
                               const AColorIndexes: TIntVector = nil): TStatPlotterLine;
begin
  Result:= TStatPlotterLine.Create(AWidth, AHeight);
  try
    Result.LineWidth:= 2;
    Result.PointRadius:= 4;
    Result.FrameColor:= clBlack;
    Result.FrameWidth:= 1;
    Result.DataFrameColor:= clGray;

    Result.TitleFont.Name:= AFontName;
    Result.TitleFont.Size:= AFontSize + 1;
    Result.TitleFont.Style:= [fsBold];
    Result.Title:= ATitle;

    Result.XTicksFont.Name:= AFontName;
    Result.XTicksFont.Size:= AFontSize;
    Result.XTicks:= AParamNames;

    Result.YTicksFont.Name:= AFontName;
    Result.YTicksFont.Size:= AFontSize;
    Result.YSeriesAdd(ADataValues);
    Result.XAxisTitle:= AHorizAxisTitle;

    Result.XAxisTitleFont.Name:= AFontName;
    Result.XAxisTitleFont.Size:= AFontSize;
    Result.YAxisType:= vatBoth;
    Result.YAxisTitle:= AVertAxisTitle;

    Result.YAxisTitleFont.Name:= AFontName;
    Result.YAxisTitleFont.Size:= AFontSize;
    Result.ColorIndexes:= AColorIndexes;
    Result.Calc;
  except
    FreeAndNil(Result);
  end;
end;

function LinesForDataMatrix(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames, ALegend: TStrVector;
                               const ADataValues: TIntMatrix;
                               const AColorIndexes: TIntVector = nil): TStatPlotterLine;
begin
  Result:= TStatPlotterLine.Create(AWidth, AHeight);
  try
    Result.LineWidth:= 2;
    Result.PointRadius:= 4;
    Result.FrameColor:= clBlack;
    Result.FrameWidth:= 1;
    Result.DataFrameColor:= clGray;

    Result.TitleFont.Name:= AFontName;
    Result.TitleFont.Size:= AFontSize + 1;
    Result.TitleFont.Style:= [fsBold];
    Result.Title:= ATitle;

    Result.LegendFont.Name:= AFontName;
    Result.LegendFont.Size:= AFontSize;
    Result.Legend:= ALegend;

    Result.XTicksFont.Name:= AFontName;
    Result.XTicksFont.Size:= AFontSize;
    Result.XTicks:= AParamNames;

    Result.YTicksFont.Name:= AFontName;
    Result.YTicksFont.Size:= AFontSize;
    Result.YSeries:= ADataValues;
    Result.YAxisType:= vatBoth;
    Result.XAxisTitle:= AHorizAxisTitle;

    Result.XAxisTitleFont.Name:= AFontName;
    Result.XAxisTitleFont.Size:= AFontSize;
    Result.YAxisTitle:= AVertAxisTitle;

    Result.YAxisTitleFont.Name:= AFontName;
    Result.YAxisTitleFont.Size:= AFontSize;
    Result.ColorIndexes:= AColorIndexes;
    Result.Calc;
  except
    FreeAndNil(Result);
  end;
end;

function VertHistForDataVector(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames: TStrVector;
                               const ADataValues: TIntVector;
                               const AOneColorPerTick: Boolean = False;
                               const AColorIndexes: TIntVector = nil): TStatPlotterVertHist;
begin
  Result:= TStatPlotterVertHist.Create(AWidth, AHeight);
  try
    Result.BarMaxWidthPercent:= 40;
    Result.FrameColor:= clBlack;
    Result.FrameWidth:= 1;
    Result.DataFrameColor:= clGray;

    Result.TitleFont.Name:= AFontName;
    Result.TitleFont.Size:= AFontSize + 1;
    Result.TitleFont.Style:= [fsBold];
    Result.Title:= ATitle;

    Result.XTicksFont.Name:= AFontName;
    Result.XTicksFont.Size:= AFontSize;
    Result.XTicks:= AParamNames;

    Result.YTicksFont.Name:= AFontName;
    Result.YTicksFont.Size:= AFontSize;
    Result.YSeriesAdd(ADataValues);
    Result.XAxisTitle:= AHorizAxisTitle;

    Result.XAxisTitleFont.Name:= AFontName;
    Result.XAxisTitleFont.Size:= AFontSize;
    Result.YAxisType:= vatBoth;
    Result.YAxisTitle:= AVertAxisTitle;

    Result.YAxisTitleFont.Name:= AFontName;
    Result.YAxisTitleFont.Size:= AFontSize;
    Result.OneColorPerTick:= AOneColorPerTick;
    Result.ColorIndexes:= AColorIndexes;
    Result.Calc;
  except
    FreeAndNil(Result);
  end;
end;

function VertHistForDataMatrix(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames, ALegend: TStrVector;
                               const ADataValues: TIntMatrix;
                               const AOneColorPerTick: Boolean = False;
                               const AColorIndexes: TIntVector = nil): TStatPlotterVertHist;
begin
  Result:= TStatPlotterVertHist.Create(AWidth, AHeight);
  try
    Result.BarMaxWidthPercent:= 20;
    Result.FrameColor:= clBlack;
    Result.FrameWidth:= 1;
    Result.DataFrameColor:= clGray;

    Result.TitleFont.Name:= AFontName;
    Result.TitleFont.Size:= AFontSize + 1;
    Result.TitleFont.Style:= [fsBold];
    Result.Title:= ATitle;

    Result.LegendFont.Name:= AFontName;
    Result.LegendFont.Size:= AFontSize;
    Result.Legend:= ALegend;

    Result.XTicksFont.Name:= AFontName;
    Result.XTicksFont.Size:= AFontSize;
    Result.XTicks:= AParamNames;

    Result.YTicksFont.Name:= AFontName;
    Result.YTicksFont.Size:= AFontSize;
    Result.YSeries:= ADataValues;
    Result.XAxisTitle:= AHorizAxisTitle;

    Result.XAxisTitleFont.Name:= AFontName;
    Result.XAxisTitleFont.Size:= AFontSize;
    Result.YAxisType:= vatBoth;
    Result.YAxisTitle:= AVertAxisTitle;

    Result.YAxisTitleFont.Name:= AFontName;
    Result.YAxisTitleFont.Size:= AFontSize;
    Result.OneColorPerTick:= AOneColorPerTick;
    Result.ColorIndexes:= AColorIndexes;
    Result.Calc;
  except
    FreeAndNil(Result);
  end;
end;

function HorizHistForDataVector(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames: TStrVector;
                               const ADataValues: TIntVector;
                               const AOneColorPerTick: Boolean = False;
                               const AColorIndexes: TIntVector = nil): TStatPlotterHorizHist;
begin
  Result:= TStatPlotterHorizHist.Create(AWidth, AHeight);
  try
    Result.BarMaxWidthPercent:= 40;
    Result.FrameColor:= clBlack;
    Result.FrameWidth:= 1;
    Result.DataFrameColor:= clGray;

    Result.TitleFont.Name:= AFontName;
    Result.TitleFont.Size:= AFontSize + 1;
    Result.TitleFont.Style:= [fsBold];
    Result.Title:= ATitle;

    Result.YTicksFont.Name:= AFontName;
    Result.YTicksFont.Size:= AFontSize;
    Result.YTicks:= AParamNames;

    Result.XTicksFont.Name:= AFontName;
    Result.XTicksFont.Size:= AFontSize;
    Result.XSeriesAdd(ADataValues);
    Result.XAxisType:= hatBoth;
    Result.XAxisTitle:= AHorizAxisTitle;

    Result.XAxisTitleFont.Name:= AFontName;
    Result.XAxisTitleFont.Size:= AFontSize;
    Result.YAxisTitle:= AVertAxisTitle;

    Result.YAxisTitleFont.Name:= AFontName;
    Result.YAxisTitleFont.Size:= AFontSize;
    Result.OneColorPerTick:= AOneColorPerTick;
    Result.ColorIndexes:= AColorIndexes;
    Result.Calc;
  except
    FreeAndNil(Result);
  end;
end;

function HorizHistForDataMatrix(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames, ALegend: TStrVector;
                               const ADataValues: TIntMatrix;
                               const AOneColorPerTick: Boolean = False;
                               const AColorIndexes: TIntVector = nil): TStatPlotterHorizHist;
begin
  Result:= TStatPlotterHorizHist.Create(AWidth, AHeight);
  try
    Result.BarMaxWidthPercent:= 40;
    Result.FrameColor:= clBlack;
    Result.FrameWidth:= 1;
    Result.DataFrameColor:= clGray;

    Result.TitleFont.Name:= AFontName;
    Result.TitleFont.Size:= AFontSize + 1;
    Result.TitleFont.Style:= [fsBold];
    Result.Title:= ATitle;

    Result.LegendFont.Name:= AFontName;
    Result.LegendFont.Size:= AFontSize;
    Result.Legend:= ALegend;

    Result.YTicksFont.Name:= AFontName;
    Result.YTicksFont.Size:= AFontSize;
    Result.YTicks:= AParamNames;

    Result.XTicksFont.Name:= AFontName;
    Result.XTicksFont.Size:= AFontSize;
    Result.XSeries:= ADataValues;
    Result.XAxisType:= hatBoth;
    Result.XAxisTitle:= AHorizAxisTitle;

    Result.XAxisTitleFont.Name:= AFontName;
    Result.XAxisTitleFont.Size:= AFontSize;
    Result.YAxisTitle:= AVertAxisTitle;

    Result.YAxisTitleFont.Name:= AFontName;
    Result.YAxisTitleFont.Size:= AFontSize;
    Result.OneColorPerTick:= AOneColorPerTick;
    Result.ColorIndexes:= AColorIndexes;
    Result.Calc;
  except
    FreeAndNil(Result);
  end;
end;

end.

