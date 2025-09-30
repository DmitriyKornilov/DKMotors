unit USheets;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics,  fpstypes, fpspreadsheetgrid, fpspreadsheet, DateUtils,
  //DK packages utils
  DK_Vector, DK_Matrix, DK_Fonts, DK_SheetConst, DK_Const, DK_SheetWriter,
  DK_DateUtils, DK_StatPlotter, DK_Math, DK_StrUtils, DK_Graph,
  DK_SheetUtils, DK_SheetTables, DK_Color, DK_SheetTypes;

const
  COLOR_BACKGROUND_TITLE = clBtnFace;

  SHEET_FONT_NAME = 'Arial';//'Times New Roman';
  SHEET_FONT_SIZE = 9;

  CHECK_SYMBOL = '✓';

  PERCENT_FRAC_DIGITS = 1;

type

  { TLogTable }

  TLogTable = class(TSheetTable)
  public
    constructor Create(const AGrid: TsWorksheetGrid; const AOnSelect: TSheetEvent; const AFont: TFont);
  end;

  { TCargoPackingSheet }

  TCargoPackingSheet = class(TLogTable)
  public
    constructor Create(const AGrid: TsWorksheetGrid; const AFont: TFont);
    procedure Update(const AMotorNames, AMotorNums, ATestDates: TStrVector);
  end;

  { TBuildLogTable }

  TBuildLogTable = class(TLogTable)
  public
    constructor Create(const AGrid: TsWorksheetGrid; const AOnSelect: TSheetEvent; const AFont: TFont);
    procedure Update(const ADate: TDate;
            const AMotorNames, AMotorNums, ARotorNums: TStrVector);
  end;

  { TTestLogTable }

  TTestLogTable = class(TLogTable)
  public
    constructor Create(const AGrid: TsWorksheetGrid; const AOnSelect: TSheetEvent; const AFont: TFont);
    procedure Update(const ADate: TDate; const ATestResults: TIntVector;
            const AMotorNames, AMotorNums, ATestNotes: TStrVector;
            const ATotalCount, AFailCount: Integer);
  end;

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

  { TMotorBuildSheet }

  TMotorBuildSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  private
    FBuildDate: TDate;
    FMotorNames, FMotorNums, FRotorNums: TStrVector;
    procedure DrawTitle;
  public
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure Draw(const ABuildDate: TDate;
                   const AMotorNames, AMotorNums, ARotorNums: TStrVector);
    procedure DrawReport(const ABeginDate, AEndDate: TDate;
                   const ANeedNumberList: Boolean;
                   const ABuildDates: TDateVector;
                   const AMotorNames, AMotorNums, ARotorNums: TStrVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCount: TIntVector);
  end;

  { TMotorTestSheet }

  TMotorTestSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  private
    FTestDate: TDate;
    FMotorNames, FMotorNums, FTestNotes: TStrVector;
    FTestResults: TIntVector;
    procedure DrawTitle;
  public
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure Draw(const ATestDate: TDate;
                   const AMotorNames, AMotorNums, ATestNotes: TStrVector;
                   const ATestResults, ATotalMotorCounts, ATotalFailCounts: TIntVector);
    procedure DrawReport(const ABeginDate, AEndDate: TDate;
                   const ANeedNumberList: Boolean;
                   const ATestDates: TDateVector;
                   const AMotorNames, AMotorNums, ATestNotes: TStrVector;
                   const ATestResults: TIntVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCounts, ATotalFailCounts: TIntVector);
  end;

  { TReportShipmentSheet }

  TReportShipmentSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  public
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

  { TStoreSheet }

  TStoreSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  public
    procedure Draw(const ADeltaDays: Integer;
                   const ATestDates: TDateVector;
                   const AMotorNames, AMotorNums: TStrVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCounts: TIntVector);
  end;

  { TBeforeTestSheet }

  TBeforeTestSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  public
    procedure Draw(const ABuildDates: TDateVector;
                   const AMotorNames, AMotorNums: TStrVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCounts: TIntVector;
                   const ATestDates: TDateMatrix;
                   const ATestFails: TIntMatrix;
                   const ATestNotes: TStrMatrix);
  end;

  { TCargoSheet }

  TCargoSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  public
    procedure Draw(const ASendDate: TDate; const AReceiverName: String;
                   const AMotorNames: TStrVector; const AMotorsCount: TIntVector;
                   const AMotorNums, ASeries: TStrMatrix);
  end;

  { TRepairSheet }

  TRepairSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  public
    procedure Draw(const APassports, ADayCounts: TIntVector;
                   const AMotorNames, AMotorNums: TStrVector;
                   const AArrivalDates, ASendingDates: TDateVector);
  end;

  { TControlSheet }

  TControlSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  public
    procedure Draw(const AMotorNames, AMotorNums, ASeries, ANotes: TStrVector);
  end;

  { TMotorCardSheet }

  TMotorCardSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  public
    procedure Draw(const ABuildDate, ASendDate: TDate;
      const AMotorName, AMotorNum, ASeries, ARotorNum, AReceiverName, AControlNote: String;
      const ATestDates: TDateVector;
      const ATestResults: TIntVector;
      const ATestNotes: TStrVector;
      const ARecDates: TDateVector;
      const AMileages, AOpinions: TIntVector;
      const APlaceNames, AFactoryNames, ADepartures,
      ADefectNames, AReasonNames, ARecNotes: TStrVector;
      const AArrivalDates, ASendingDates: TDateVector;
      const APassports, AWorkDayCounts: TIntVector;
      const ARepairNotes: TStrVector);
  end;

  { TReclamationSheet }

  TReclamationSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  private
    FRecDates, FBuildDates, FArrivalDates, FSendingDates: TDateVector;
    FMileages, FOpinions, FReasonColors: TIntVector;
    FPlaceNames, FFactoryNames, FDepartures: TStrVector;
    FDefectNames, FReasonNames, FRecNotes: TStrVector;
    FMotorNames, FMotorNums: TStrVector;

    procedure DrawTitle;
  public
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure Draw(const ARecDates, ABuildDates, AArrivalDates, ASendingDates: TDateVector;
                   const AMileages, AOpinions, AReasonColors: TIntVector;
                   const APlaceNames, AFactoryNames, ADepartures,
                         ADefectNames, AReasonNames, ARecNotes,
                         AMotorNames, AMotorNums: TStrVector);
  end;

  function LinesForDataVector(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames: TStrVector;
                               const ADataValues: TIntVector;
                               const AColorIndexes: TIntVector = nil): TStatPlotterLine;

  function VertHistForDataVector(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames: TStrVector;
                               const ADataValues: TIntVector;
                               const AOneColorPerTick: Boolean = False;
                               const AColorIndexes: TIntVector = nil): TStatPlotterVertHist;

  function LinesForDataMatrix(const AWidth, AHeight, AFontSize: Integer;
                               const AFontName, ATitle, AVertAxisTitle, AHorizAxisTitle: String;
                               const AParamNames, ALegend: TStrVector;
                               const ADataValues: TIntMatrix;
                               const AColorIndexes: TIntVector = nil): TStatPlotterLine;

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

{ TLogTable }

constructor TLogTable.Create(const AGrid: TsWorksheetGrid;
  const AOnSelect: TSheetEvent; const AFont: TFont);
begin
  inherited Create(AGrid);
  SetFontsName(AFont.Name);
  SetFontsSize(AFont.Size);
  RowBeforeFont.Size:= AFont.Size + 2;
  RowBeforeFont.Style:= [fsBold];
  HeaderFont.Style:= [fsBold];
  RowAfterFont.Style:= [fsBold];
  OnSelect:= AOnSelect;
end;

{ TCargoPackingSheet }

constructor TCargoPackingSheet.Create(const AGrid: TsWorksheetGrid; const AFont: TFont);
begin
  inherited Create(AGrid, nil, AFont);
  CanSelect:= False;
  AddColumn('№ п/п', 70);
  AddColumn('Наименование', 350);
  AddColumn('Заводской номер', 160);
  AddColumn('Дата изготовления', 160);
  AddToHeader(1, 1, '№ п/п');
  AddToHeader(1, 2, 'Наименование');
  AddToHeader(1, 3, 'Заводской номер');
  AddToHeader(1, 4, 'Дата изготовления');
end;

procedure TCargoPackingSheet.Update(const AMotorNames, AMotorNums,
  ATestDates: TStrVector);
begin
  SetColumnOrder('№ п/п');
  SetColumnString('Наименование', AMotorNames);
  SetColumnString('Заводской номер', AMotorNums);
  SetColumnString('Дата изготовления', ATestDates);
  Draw;
end;

{ TTestLogTable }

constructor TTestLogTable.Create(const AGrid: TsWorksheetGrid;
  const AOnSelect: TSheetEvent; const AFont: TFont);
begin
  inherited Create(AGrid, AOnSelect, AFont);
  AddColumn('№ п/п', 50);
  AddColumn('Наименование двигателя', 300);
  AddColumn('Номер двигателя', 150);
  AddColumn('Результат испытаний', 150);
  AddColumn('Примечание', 200, haLeft);
  AddToHeader(2, 1, '№ п/п');
  AddToHeader(2, 2, 'Наименование двигателя');
  AddToHeader(2, 3, 'Номер двигателя');
  AddToHeader(2, 4, 'Результат испытаний');
  AddToHeader(2, 5, 'Примечание');
  SetExtraFont('Результат испытаний', 'Результат испытаний', 'брак',
               ValuesFont.Name, ValuesFont.Size, ValuesFont.Style, clRed);
  SetExtraFont('Примечание', 'Результат испытаний', 'брак',
               ValuesFont.Name, ValuesFont.Size, ValuesFont.Style, clRed);
end;

procedure TTestLogTable.Update(const ADate: TDate; const ATestResults: TIntVector;
            const AMotorNames, AMotorNums, ATestNotes: TStrVector;
            const ATotalCount, AFailCount: Integer);
var
  S: String;
  TestResults: TStrVector;
begin
  TestResults:= VIntToStr(ATestResults);
  VChangeIf(TestResults, '0', 'норма');
  VChangeIf(TestResults, '1', 'брак');

  S:= 'Журнал испытаний за ' + FormatDateTime('dd.mm.yyyy', ADate);
  SetRowBefore(S, haLeft);
  SetColumnOrder('№ п/п');
  SetColumnString('Наименование двигателя', AMotorNames);
  SetColumnString('Номер двигателя', AMotorNums);
  SetColumnString('Результат испытаний', TestResults);
  SetColumnString('Примечание', ATestNotes);
  S:= 'Норма: ' + IntToStr(ATotalCount-AFailCount) + ' / ' +
      'Брак: ' + IntToStr(AFailCount);
  SetRowAfter(S, haLeft);
  Draw;
end;

{ TBuildLogTable }

constructor TBuildLogTable.Create(const AGrid: TsWorksheetGrid;
  const AOnSelect: TSheetEvent; const AFont: TFont);
begin
  inherited Create(AGrid, AOnSelect, AFont);
  AddColumn('№ п/п', 50);
  AddColumn('Наименование двигателя', 300);
  AddColumn('Номер двигателя', 150);
  AddColumn('Номер ротора', 150);
  AddToHeader(2, 1, '№ п/п');
  AddToHeader(2, 2, 'Наименование двигателя');
  AddToHeader(2, 3, 'Номер двигателя');
  AddToHeader(2, 4, 'Номер ротора');
end;

procedure TBuildLogTable.Update(const ADate: TDate;
                   const AMotorNames, AMotorNums, ARotorNums: TStrVector);
var
  S: String;
begin
  S:= 'Журнал сборки за ' + FormatDateTime('dd.mm.yyyy', ADate);
  SetRowBefore(S, haLeft);
  SetColumnOrder('№ п/п');
  SetColumnString('Наименование двигателя', AMotorNames);
  SetColumnString('Номер двигателя', AMotorNums);
  SetColumnString('Номер ротора', ARotorNums);
  Draw;
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

{ TStatisticSeveralPeriodsAtPlaceNamesSheet }

procedure TStatisticSeveralPeriodsAtPlaceNamesSheet.DrawReport;
begin
  DrawParamsReport('предприятиям', 'Предприятие');
end;

{ TStatisticSeveralPeriodsAtDefectNamesSheet }

procedure TStatisticSeveralPeriodsAtDefectNamesSheet.DrawReport;
begin
  DrawParamsReport('неисправным элементам', 'Неисправный элемент');
end;

{ TStatisticSeveralPeriodsAtMotorNamesSheet }

procedure TStatisticSeveralPeriodsAtMotorNamesSheet.DrawReport;
begin
  DrawParamsReport('наименованиям электродвигателей', 'Наименование');
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
end;

function TStatisticSheet.CalcHorizHistHeight(const ARowCount: Integer;
  const ADataInRowCount: Integer): Integer;
var
  H: Integer;
begin
  H:= Round(ADataInRowCount*HORIZ_HIST_DATAROW_HEIGHT*0.45);
  H:= Max(H, HORIZ_HIST_DATAROW_HEIGHT);
  Result:= ARowCount*H + HORIZ_HIST_ADDITION_HEIGHT;
end;

function TStatisticSheet.CalcVertHistHeight: Integer;
begin
  Result:= Round(0.4*FGraphWidth);
end;

procedure TStatisticSheet.DrawEmptyRow(var ARow: Integer; const AHeight: Integer = EMPTY_ROW_HEIGHT);
begin
  Inc(ARow);
  FWriter.WriteText(ARow, 1, EmptyStr);
  FWriter.SetRowHeight(ARow, AHeight);
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
    ScaleX:= 1;
    ScaleY:= 1;
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

{ TStatisticSinglePeriodAtPlaceNamesSheet }

procedure TStatisticSinglePeriodAtPlaceNamesSheet.DrawReport;
begin
  DrawParamsReport('предприятиям', 'Предприятие');
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

{ TStatisticSinglePeriodAtDefectNamesSheet }

procedure TStatisticSinglePeriodAtDefectNamesSheet.DrawReport;
begin
  DrawParamsReport('неисправным элементам', 'Неисправный элемент');
end;

{ TStatisticSinglePeriodAtMotorNamesSheet }

procedure TStatisticSinglePeriodAtMotorNamesSheet.DrawReport;
begin
  DrawParamsReport('наименованиям электродвигателей', 'Наименование');
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
  FWriter.SetRowHeight(R, 10);

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

{ TControlSheet }

function TControlSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    60,   //№п/п
    300,  // Наименование
    150,  // Номер
    100,  // Партия
    500   // Примечание
  ]);
end;

procedure TControlSheet.Draw(const AMotorNames, AMotorNums, ASeries, ANotes: TStrVector);
var
  R, i: Integer;
  S: String;
begin
  Writer.BeginEdit;
  Writer.SetBackgroundClear;

  R:= 1;
  Writer.SetAlignment(haCenter, vaCenter);
  S:= 'Электродвигатели на контроле';
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);

  R:=R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, '№п/п', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Наименование', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Номер', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Партия', cbtOuter, True, True);
  Writer.WriteText(R, 5, 'Примечание', cbtOuter, True, True);

  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  for i:= 0 to High(AMotorNums) do
  begin
    R:=R + 1;
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteNumber(R, 1, i+1, cbtOuter);
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
    Writer.WriteText(R, 4, ASeries[i], cbtOuter, True, True);
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.WriteText(R, 5, ANotes[i], cbtOuter, True, True);
  end;

  Writer.EndEdit;
end;

{ TRepairSheet }

function TRepairSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    60,   //№п/п
    300,  // Наименование
    150,  // Номер
    100,  // Наличие паспорта
    150,  // Прибыл в ремонт
    150,  // Убыл из ремонта
    130   // Срок ремонта (дней)
  ]);
end;

procedure TRepairSheet.Draw(const APassports, ADayCounts: TIntVector;
               const AMotorNames, AMotorNums: TStrVector;
               const AArrivalDates, ASendingDates: TDateVector);
var
  R, i: Integer;
  S: String;
begin
  Writer.BeginEdit;
  Writer.SetBackgroundClear;

  R:= 1;
  Writer.SetAlignment(haCenter, vaCenter);
  S:= 'Гарантийный ремонт электродвигателей';
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount, S, cbtNone, True, True);

  R:=R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, '№п/п', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Наименование', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Номер', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Наличие паспорта', cbtOuter, True, True);
  Writer.WriteText(R, 5, 'Прибыл в ремонт', cbtOuter, True, True);
  Writer.WriteText(R, 6, 'Убыл из ремонта', cbtOuter, True, True);
  Writer.WriteText(R, 7, 'Срок ремонта (рабочих дней)', cbtOuter, True, True);

  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  for i:= 0 to High(AMotorNums) do
  begin
    R:=R + 1;
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteNumber(R, 1, i+1, cbtOuter);
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
    S:= EmptyStr;
    if APassports[i]>0 then S:= CHECK_SYMBOL;
    Writer.WriteText(R, 4, S, cbtOuter);
    Writer.WriteDate(R, 5, AArrivalDates[i], cbtOuter);
    if ASendingDates[i]>0 then
      Writer.WriteDate(R, 6, ASendingDates[i], cbtOuter)
    else
      Writer.WriteText(R, 6, EmptyStr, cbtOuter);
    if ADayCounts[i]>0 then
      Writer.WriteNumber(R, 7, ADayCounts[i], cbtOuter)
    else
      Writer.WriteText(R, 7, EmptyStr, cbtOuter);
  end;

  Writer.EndEdit;
end;

{ TReportShipmentSheet }

function TReportShipmentSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    150, // № п/п - Дата отправки
    300, // Наименование двигателя
    150, // Номер двигателя
    300  // Грузополучатель
  ]);
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
  if VIsNil(ATotalMotorNames) then Exit;

  Writer.BeginEdit;
  Writer.SetBackgroundClear;

  C:= 3 + Ord(ANeedNumberList);

  S:= 'Отчет по отгрузке электродвигателей за ';
  if SameDate(ABeginDate, AEndDate) then
    S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
  else
    S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
       ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);

  //Заголовок отчета
  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);

  if not ASingleReceiver then
  begin

    Writer.WriteText(R, 1, R, C, S, cbtNone, True, True);
    //Таблица сводная: наименование/количество
    R:= R + 1;
    Writer.WriteText(R, 1, R, C, EmptyStr, cbtNone);
    //Writer.SetRowHeight(R, 10);
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    Writer.WriteText(R, 1, R, C, 'Всего отгружено', cbtNone, True, True);
    R:= R + 1;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
    Writer.WriteText(R, 3, 'Количество', cbtOuter, True, True);
    Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
    for i:= 0 to High(ATotalMotorNames) do
    begin
      R:= R + 1;
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
      Writer.SetAlignment(haCenter, vaCenter);
      Writer.WriteNumber(R, 3, ATotalMotorCount[i], cbtOuter);
    end;
  end
  else
    Writer.WriteText(R, 1, R, 3, S, cbtNone, True, True);

  for i:= 0 to High(ARecieverNames) do
  begin
    R:= R + 1;
    Writer.WriteText(R, 1, R, 3, EmptyStr, cbtNone);
    //Writer.SetRowHeight(R, 15);
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    Writer.WriteText(R, 1, R, 3, ARecieverNames[i], cbtNone, True, True);
    R:= R + 1;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
    Writer.WriteText(R, 3, 'Количество', cbtOuter, True, True);
    for j:= 0 to High(ARecieverMotorNames[i]) do
    begin
      R:= R + 1;
      Writer.SetFont(Font.Name, Font.Size, [], clBlack);
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.WriteText(R, 1, R, 2, ARecieverMotorNames[i,j], cbtOuter, True, True);
      Writer.SetAlignment(haCenter, vaCenter);
      Writer.WriteNumber(R, 3, ARecieverMotorCounts[i,j], cbtOuter);
    end;
  end;

  if ANeedNumberList then
  begin
    R:= R + 1;
    if not ASingleReceiver then
      Writer.WriteText(R, 1, R, 4, EmptyStr, cbtNone)
    else
      Writer.WriteText(R, 1, R, 3, EmptyStr, cbtNone);
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
    if not ASingleReceiver then
      Writer.WriteText(R, 1, R, 4, 'Пономерной список', cbtNone, True, True)
    else
      Writer.WriteText(R, 1, R, 3, 'Пономерной список', cbtNone, True, True);
    R:= R + 1;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteText(R, 1, 'Дата отгрузки', cbtOuter, True, True);
    Writer.WriteText(R, 2, 'Наименование', cbtOuter, True, True);
    Writer.WriteText(R, 3, 'Номер', cbtOuter, True, True);
    if not ASingleReceiver then
      Writer.WriteText(R, 4, 'Грузополучатель', cbtOuter, True, True);

    Writer.SetFont(Font.Name, Font.Size, [], clBlack);

    for i:= 0 to High(AListSendDates) do
    begin
      R:= R + 1;
      Writer.SetAlignment(haCenter, vaCenter);
      Writer.WriteDate(R, 1, AListSendDates[i], cbtOuter);
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.WriteText(R, 2, AListMotorNames[i], cbtOuter, True, True);
      Writer.SetAlignment(haCenter, vaCenter);
      Writer.WriteText(R, 3, AListMotorNums[i], cbtOuter, True, True);
      if not ASingleReceiver then
      begin
        Writer.SetAlignment(haLeft, vaCenter);
        Writer.WriteText(R, 4, AListReceiverNames[i], cbtOuter, True, True);
      end;
    end;
  end;

  Writer.EndEdit;
end;

{ TBeforeTestSheet }

function TBeforeTestSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    150, // Дата сборки
    300, // Наименование двигателя
    200, // Номер двигателя
    300  // Примечание
  ]);
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
  Writer.Clear;
  if VIsNil(ABuildDates) then Exit;

  Writer.BeginEdit;
  Writer.SetBackgroundClear;

  //Заголовок отчета
  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, 4, 'На испытаниях', cbtNone, True, True);
  //Таблица сводная: наименование/количество
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Количество', cbtOuter, True, True);
  Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
  for i:= 0 to High(ATotalMotorNames) do
  begin
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteNumber(R, 3, ATotalMotorCounts[i], cbtOuter);
  end;
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, 2, 'ИТОГО', cbtOuter, True, True);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteNumber(R, 3, VSum(ATotalMotorCounts), cbtOuter);

  //Таблица пономерная
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, 'Дата сборки', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Испытания', cbtOuter, True, True);

  FrozenCount:= R;

  for i:= 0 to High(ABuildDates) do
  begin
    R:= R + 1;
    N:= 0;
    if not VIsNil(ATestDates[i]) then
      N:= High(ATestDates[i]);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
    Writer.WriteDate(R, 1, R+N, 1, ABuildDates[i], cbtOuter);
    Writer.WriteText(R, 2, R+N, 2, AMotorNames[i], cbtOuter, True, True);
    Writer.WriteText(R, 3, R+N, 3, AMotorNums[i], cbtOuter, True, True);
    Writer.SetAlignment(haLeft, vaCenter);
    if VIsNil(ATestDates[i]) then
      Writer.WriteText(R, 4, R+N, 4, 'Не проводились', cbtOuter, True, True)
    else begin
      for j:= 0 to High(ATestDates[i]) do
      begin
        if ATestFails[i][j]=0 then
        begin
          Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
          S:= ' - норма';
        end
        else begin
          Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clRed);
          S:= ' - брак';
        end;
        S:= FormatDateTime('dd.mm.yyyy', ATestDates[i][j]) + S;
        if ATestNotes[i][j]<>EmptyStr then
          S:= S + ' (' + ATestNotes[i][j] + ')';
        Writer.WriteText(R+j, 4, S, cbtOuter, True, True);
      end;
    end;
    R:= R + N;
  end;

  Writer.SetFrozenRows(FrozenCount);
  Writer.SetRepeatedRows(FrozenCount,FrozenCount);
  if R>FrozenCount then
    Writer.DrawBorders(FrozenCount+1, 1, R, Writer.ColCount, cbtAll);

  Writer.EndEdit;
end;

{ TMotorCardSheet }

function TMotorCardSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    130, // Дата уведомления
    130, // Пробег, км
    130, // Предприятие
    130, // Завод
    130, // Выезд/ФИО
    130, // Неисправный элемент
    130, // Причина неисправности
    70,  //Особое мнение
    350  //Примечание
  ]);
end;

procedure TMotorCardSheet.Draw(const ABuildDate, ASendDate: TDate;
      const AMotorName, AMotorNum, ASeries, ARotorNum, AReceiverName, AControlNote: String;
      const ATestDates: TDateVector;
      const ATestResults: TIntVector;
      const ATestNotes: TStrVector;
      const ARecDates: TDateVector;
      const AMileages, AOpinions: TIntVector;
      const APlaceNames, AFactoryNames, ADepartures,
      ADefectNames, AReasonNames, ARecNotes: TStrVector;
      const AArrivalDates, ASendingDates: TDateVector;
      const APassports, AWorkDayCounts: TIntVector;
      const ARepairNotes: TStrVector);
var
  R, i: Integer;
  S: String;
begin
  Writer.BeginEdit;
  Writer.SetAlignment(haCenter, vaCenter);

  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount,  'УЧЁТНАЯ КАРТОЧКА ЭЛЕКТРОДВИГАТЕЛЯ');

  R:= R + 1;
  Writer.WriteText(R, 1,  EmptyStr);
  Writer.SetRowHeight(R, 10);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, 2,  'Наименование', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Номер', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Партия', cbtOuter, True, True);
  Writer.WriteText(R, 5, 'Дата сборки', cbtOuter, True, True);
  Writer.WriteText(R, 6, 'Ротор', cbtOuter, True, True);
  Writer.WriteText(R, 7, R, Writer.ColCount,  'Отгружен', cbtOuter, True, True);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  Writer.WriteText(R, 1, R, 2,  AMotorName, cbtOuter, True, True);
  Writer.WriteText(R, 3, AMotorNum, cbtOuter, True, True);
  Writer.WriteText(R, 4, ASeries, cbtOuter, True, True);
  Writer.WriteDate(R, 5, ABuildDate, cbtOuter);
  Writer.WriteText(R, 6, ARotorNum, cbtOuter, True, True);
  Writer.SetAlignment(haLeft, vaCenter);
  S:= EmptyStr;
  if AReceiverName<>EmptyStr then
    S:= FormatDateTime('dd.mm.yyyy', ASendDate) + ' - ' + AReceiverName;
  Writer.WriteText(R, 7, R, Writer.ColCount,  S, cbtOuter, True, True);

  R:= R + 1;
  Writer.WriteText(R, 1,  EmptyStr);
  Writer.SetRowHeight(R, 10);
  R:= R + 1;
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.SetFont(Font.Name, Font.Size+1, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount,  'Приемо-сдаточные испытания:');
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, 'Дата', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Результат', cbtOuter, True, True);
  Writer.WriteText(R, 3, R, Writer.ColCount, 'Примечание', cbtOuter, True, True);
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  if VIsNil(ATestDates) then
  begin
    R:= R + 1;
    Writer.WriteText(R, 1, EmptyStr, cbtOuter);
    Writer.WriteText(R, 2, EmptyStr, cbtOuter, True, True);
    Writer.WriteText(R, 3, R, Writer.ColCount, EmptyStr, cbtOuter, True, True);
  end
  else begin
    for i:= 0 to High(ATestDates) do
    begin
      R:= R + 1;
      if ATestResults[i]=0 then
        S:= 'норма'
      else
        S:= 'брак';
      Writer.SetAlignment(haCenter, vaCenter);
      Writer.WriteDate(R, 1, ATestDates[i], cbtOuter);
      Writer.WriteText(R, 2, S, cbtOuter, True, True);
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.WriteText(R, 3, R, Writer.ColCount, ATestNotes[i], cbtOuter, True, True);
    end;
  end;

  R:= R + 1;
  Writer.WriteText(R, 1,  EmptyStr);
  Writer.SetRowHeight(R, 10);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size+1, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, Writer.ColCount,  'Контроль:');
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount,  AControlNote, cbtOuter, True, True);

  R:= R + 1;
  Writer.WriteText(R, 1,  EmptyStr);
  Writer.SetRowHeight(R, 10);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size+1, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, Writer.ColCount,  'Рекламации:');
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, 'Дата уведомления', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Пробег, км', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Предприятие', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Завод', cbtOuter, True, True);
  Writer.WriteText(R, 5, 'Выезд', cbtOuter, True, True);
  Writer.WriteText(R, 6, 'Неисправный элемент', cbtOuter, True, True);
  Writer.WriteText(R, 7, 'Причина неисправности', cbtOuter, True, True);
  Writer.WriteText(R, 8, 'Особое мнение', cbtOuter, True, True);
  Writer.WriteText(R, 9, 'Примечание', cbtOuter, True, True);


  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  if VIsNil(ARecDates) then
  begin
    R:= R + 1;
    for i:= 1 to 9 do
      Writer.WriteText(R, i, EmptyStr, cbtOuter);
  end
  else begin
    for i:= 0 to High(ARecDates) do
    begin
      R:= R + 1;
      Writer.SetAlignment(haCenter, vaTop);
      Writer.WriteDate(R, 1, ARecDates[i], cbtOuter);
      if AMileages[i]>=0 then
        Writer.WriteNumber(R, 2, AMileages[i], cbtOuter)
      else
        Writer.WriteText(R, 2, EmptyStr, cbtOuter);
      Writer.WriteText(R, 3, APlaceNames[i], cbtOuter, True, True);
      Writer.WriteText(R, 4, AFactoryNames[i], cbtOuter, True, True);
      Writer.WriteText(R, 5, ADepartures[i], cbtOuter, True, True);
      Writer.WriteText(R, 6, ADefectNames[i], cbtOuter, True, True);
      Writer.WriteText(R, 7, AReasonNames[i], cbtOuter, True, True);
      S:= EmptyStr;
      if AOpinions[i]=1 then S:= CHECK_SYMBOL;
      Writer.WriteText(R, 8, S, cbtOuter);
      Writer.SetAlignment(haLeft, vaTop);
      Writer.WriteText(R, 9, ARecNotes[i], cbtOuter, True, True);
    end;
  end;

  R:= R + 1;
  Writer.WriteText(R, 1,  EmptyStr);
  Writer.SetRowHeight(R, 10);
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size+1, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, Writer.ColCount,  'Гарантийный ремонт:');
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, 'Дата прибытия', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Наличие паспорта', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Дата убытия', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Срок ремонта (рабочих дней)', cbtOuter, True, True);
  Writer.WriteText(R, 5, R, Writer.ColCount,  'Примечание', cbtOuter, True, True);
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  if VIsNil(AArrivalDates) then
  begin
    R:= R + 1;
    Writer.WriteText(R, 1, EmptyStr);
    Writer.WriteText(R, 2, EmptyStr);
    Writer.WriteText(R, 3, EmptyStr);
    Writer.WriteText(R, 4, EmptyStr);
    Writer.WriteText(R, 5, R, Writer.ColCount,  EmptyStr);
  end
  else begin
    for i:= 0 to High(AArrivalDates) do
    begin
      R:= R + 1;
      Writer.SetAlignment(haCenter, vaTop);
      Writer.WriteDate(R, 1, AArrivalDates[i], cbtOuter);
      S:= EmptyStr;
      if APassports[i]>0 then
        S:= CHECK_SYMBOL;
      Writer.WriteText(R, 2, S, cbtOuter);
      if ASendingDates[i]>0 then
        Writer.WriteDate(R, 3, ASendingDates[i], cbtOuter)
      else
        Writer.WriteText(R, 3, EmptyStr, cbtOuter);
      Writer.WriteNumber(R, 4, AWorkDayCounts[i], cbtOuter);
      Writer.SetAlignment(haLeft, vaTop);
      Writer.WriteText(R, 5, R, Writer.ColCount,  ARepairNotes[i], cbtOuter, True, True);
    end;
  end;

  Writer.EndEdit;
end;

{ TStoreSheet }

function TStoreSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    150, // Дата сдачи (испытаний)
    300, // Наименование двигателя
    200//, // Номер двигателя
    //100  // Номер партии
  ]);
end;

procedure TStoreSheet.Draw(const ADeltaDays: Integer;
       const ATestDates: TDateVector;
       const AMotorNames, AMotorNums: TStrVector;
       const ATotalMotorNames: TStrVector;
       const ATotalMotorCounts: TIntVector);
var
  R, i, FrozenCount: Integer;
  S: String;
begin
  if VIsNil(ATestDates) then Exit;

  Writer.BeginEdit;
  Writer.SetBackgroundClear;

  S:= 'Отчет по наличию электродвигателей на складе на ' +
      FormatDateTime('dd.mm.yyyy', Date);
  if ADeltaDays>0 then
    S:= S + ', хранящихся дольше ' + IntToStr(ADeltaDays) +
            ' дней (сданы ' +
            FormatDateTime('dd.mm.yyyy', IncDay(Date,-ADeltaDays)) +
            ' и ранее)';

  //Заголовок отчета
  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, R, 3, S, cbtNone, True, True);
  //Таблица сводная: наименование/количество
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Количество', cbtOuter, True, True);
  Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
  for i:= 0 to High(ATotalMotorNames) do
  begin
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteNumber(R, 3, ATotalMotorCounts[i], cbtOuter);
  end;
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, 2, 'ИТОГО', cbtOuter, True, True);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteNumber(R, 3, VSum(ATotalMotorCounts), cbtOuter);

  //Таблица пономерная
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, 'Дата сдачи', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
  //Writer.WriteText(R, 4, 'Номер партии', cbtOuter, True, True);

  FrozenCount:= R;

  Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
  for i:= 0 to High(ATestDates) do
  begin
    R:= R + 1;
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteDate(R, 1, ATestDates[i], cbtOuter);
    Writer.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
    Writer.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
    //Writer.WriteText(R, 4, ASeries[i], cbtOuter, True, True);
  end;

  Writer.SetFrozenRows(FrozenCount);

  Writer.EndEdit;
end;

{ TMotorTestSheet }

function TMotorTestSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    150, // Дата испытаний
    300, // Наименование двигателя
    200, // Номер двигателя
    //150, // Номер партии
    200, // Результат (норма/брак)
    300  // Примечание
  ]);
end;

procedure TMotorTestSheet.DrawTitle;
var
  R: Integer;
begin
  R:= 1;
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, 5, 'Журнал испытаний за ' +
                    FormatDateTime('dd.mm.yyyy', FTestDate),
                    cbtBottom, True, True);
  R:= R+1;
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  //Writer.SetBackground(COLOR_BACKGROUND_TITLE);
  Writer.WriteText(R, 1, '№ п/п', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
  //Writer.WriteText(R, 4, 'Номер партии', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Результат испытаний', cbtOuter, True, True);
  Writer.WriteText(R, 5, 'Примечание', cbtOuter, True, True);
end;

procedure TMotorTestSheet.DrawLine(const AIndex: Integer;
                                 const ASelected: Boolean);
var
  R: Integer;
  S: String;
begin
  R:= AIndex + 3;

  if ASelected then
    Writer.SetBackground(DefaultSelectionBGColor)
  else
    Writer.SetBackgroundClear;

  if FTestResults[AIndex]=0 then
  begin
    Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
    S:= 'норма';
  end
  else begin
    Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clRed);
    S:= 'брак';
  end;

  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteNumber(R, 1, AIndex+1, cbtOuter);
  Writer.WriteText(R, 2, FMotorNames[AIndex], cbtOuter, True, True);
  Writer.WriteText(R, 3, FMotorNums[AIndex], cbtOuter, True, True);
  Writer.WriteText(R, 4, S, cbtOuter, True, True);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 5, FTestNotes[AIndex], cbtOuter, True, True);
end;

procedure TMotorTestSheet.Draw(const ATestDate: TDate;
            const AMotorNames, AMotorNums, ATestNotes: TStrVector;
            const ATestResults, ATotalMotorCounts, ATotalFailCounts: TIntVector);
var
  i, N, K: Integer;
begin
  Writer.BeginEdit;

  FTestDate:= ATestDate;
  FMotorNames:= AMotorNames;
  FMotorNums:= AMotorNums;
  FTestResults:= ATestResults;
  FTestNotes:= ATestNotes;

  DrawTitle;

  if not VIsNil(AMotorNames) then
  begin
    for i:= 0 to High(AMotorNames) do
      DrawLine(i, False);

    i:= Length(AMotorNames) + 3;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    N:= VSum(ATotalMotorCounts);
    K:= VSum(ATotalFailCounts);
    Writer.WriteText(i, 1,  'Норма: ' + IntToStr(N-K), cbtTop, True, True);
    Writer.WriteText(i, 2,  'Брак: ' + IntToStr(K), cbtTop, True, True);
    Writer.SetFrozenRows(2);
  end;

  Writer.EndEdit;
end;

procedure TMotorTestSheet.DrawReport(const ABeginDate, AEndDate: TDate;
                   const ANeedNumberList: Boolean;
                   const ATestDates: TDateVector;
                   const AMotorNames, AMotorNums, ATestNotes: TStrVector;
                   const ATestResults: TIntVector;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCounts, ATotalFailCounts: TIntVector);
var
  R, i, FrozenCount, N, K: Integer;
  S: String;
begin
  if VIsNil(ATestDates) then Exit;

  Writer.BeginEdit;
  Writer.SetBackgroundClear;

  S:= 'Отчет по испытанию электродвигателей за ';
  if SameDate(ABeginDate, AEndDate) then
    S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
  else
    S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
       ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);

  //Заголовок отчета
  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  if ANeedNumberList then
    Writer.WriteText(R, 1, R, 5, S, cbtNone, True, True)
  else
    Writer.WriteText(R, 1, R, 4, S, cbtNone, True, True);
  //Таблица сводная: наименование/количество
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Норма', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Брак', cbtOuter, True, True);
  Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
  FrozenCount:= R;
  for i:= 0 to High(ATotalMotorNames) do
  begin
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
    Writer.SetAlignment(haCenter, vaCenter);
    N:= ATotalMotorCounts[i];
    K:= ATotalFailCounts[i];
    Writer.WriteNumber(R, 3, N-K, cbtOuter);
    Writer.WriteNumber(R, 4, K, cbtOuter);
  end;
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, 2, 'ИТОГО', cbtOuter, True, True);
  Writer.SetAlignment(haCenter, vaCenter);
  N:= VSum(ATotalMotorCounts);
  K:= VSum(ATotalFailCounts);
  Writer.WriteNumber(R, 3, N-K, cbtOuter);
  Writer.WriteNumber(R, 4, K, cbtOuter);

  if ANeedNumberList then
  begin
    //Таблица пономерная
    R:= R + 2;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteText(R, 1, 'Дата испытаний', cbtOuter, True, True);
    Writer.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
    Writer.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
    Writer.WriteText(R, 4, 'Результат испытаний', cbtOuter, True, True);
    Writer.WriteText(R, 5, 'Примечание', cbtOuter, True, True);
    FrozenCount:= R;

    Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
    for i:= 0 to High(ATestDates) do
    begin
      R:= R + 1;
      if ATestResults[i]=0 then
      begin
        S:= 'норма';
        Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
      end
      else begin
        S:= 'брак';
        Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clRed);
      end;
      Writer.SetAlignment(haCenter, vaCenter);
      Writer.WriteDate(R, 1, ATestDates[i], cbtOuter);
      Writer.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
      Writer.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
      Writer.WriteText(R, 4, S, cbtOuter, True, True);
      Writer.SetAlignment(haLeft, vaCenter);
      Writer.WriteText(R, 5, ATestNotes[i], cbtOuter, True, True);
    end;
  end;

  Writer.SetFrozenRows(FrozenCount);

  Writer.EndEdit;
end;

{ TMotorBuildSheet }

function TMotorBuildSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    150, // № п/п - Дата сборки
    300, // Наименование двигателя
    200, // Номер двигателя
    200  // Номер Ротора
  ]);
end;

procedure TMotorBuildSheet.DrawTitle;
var
  R: Integer;
begin
  R:= 1;
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, 4, 'Журнал сборки за ' +
                    FormatDateTime('dd.mm.yyyy', FBuildDate),
                    cbtBottom, True, True);

  R:= R+1;
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, '№ п/п', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Номер ротора', cbtOuter, True, True);
end;

procedure TMotorBuildSheet.DrawLine(const AIndex: Integer;
                                  const ASelected: Boolean);
var
  R: Integer;
begin
  R:= AIndex + 3;

  Writer.SetAlignment(haCenter, vaCenter);

  Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
  if ASelected then
    Writer.SetBackground(DefaultSelectionBGColor)
  else
    Writer.SetBackgroundClear;

  Writer.WriteNumber(R, 1, AIndex+1, cbtOuter);
  Writer.WriteText(R, 2, FMotorNames[AIndex], cbtOuter, True, True);
  Writer.WriteText(R, 3, FMotorNums[AIndex], cbtOuter, True, True);
  Writer.WriteText(R, 4, FRotorNums[AIndex], cbtOuter, True, True);
end;

procedure TMotorBuildSheet.Draw(const ABuildDate: TDate;
                        const AMotorNames, AMotorNums, ARotorNums: TStrVector);
var
  i: Integer;
begin
  Writer.BeginEdit;

  FBuildDate:= ABuildDate;
  FMotorNames:= AMotorNames;
  FMotorNums:= AMotorNums;
  FRotorNums:= ARotorNums;

  DrawTitle;

  if not VIsNil(AMotorNames) then
  begin
    for i:= 0 to High(AMotorNames) do
      DrawLine(i, False);
    Writer.SetFrozenRows(2);
  end;

  Writer.EndEdit;
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
  if VIsNil(ABuildDates) then Exit;

  Writer.BeginEdit;
  Writer.SetBackgroundClear;

  S:= 'Отчет по сборке электродвигателей за ';
  if SameDate(ABeginDate, AEndDate) then
    S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
  else
    S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
       ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);

  //Заголовок отчета
  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  if ANeedNumberList then
    Writer.WriteText(R, 1, R, 4, S, cbtNone, True, True)
  else
    Writer.WriteText(R, 1, R, 3, S, cbtNone, True, True);
  //Таблица сводная: наименование/количество
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, R, 2, 'Наименование', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Количество', cbtOuter, True, True);
  Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
  FrozenCount:= R;
  for i:= 0 to High(ATotalMotorNames) do
  begin
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.WriteText(R, 1, R, 2, ATotalMotorNames[i], cbtOuter, True, True);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteNumber(R, 3, ATotalMotorCount[i], cbtOuter);
  end;
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, 2, 'ИТОГО', cbtOuter, True, True);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteNumber(R, 3, VSum(ATotalMotorCount), cbtOuter);

  if ANeedNumberList then
  begin
    //Таблица пономерная
    R:= R + 2;
    Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteText(R, 1, 'Дата сборки', cbtOuter, True, True);
    Writer.WriteText(R, 2, 'Наименование двигателя', cbtOuter, True, True);
    Writer.WriteText(R, 3, 'Номер двигателя', cbtOuter, True, True);
    Writer.WriteText(R, 4, 'Номер ротора', cbtOuter, True, True);
    FrozenCount:= R;
    Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
    for i:= 0 to High(ABuildDates) do
    begin
      R:= R + 1;
      Writer.WriteDate(R, 1, ABuildDates[i], cbtOuter);
      Writer.WriteText(R, 2, AMotorNames[i], cbtOuter, True, True);
      Writer.WriteText(R, 3, AMotorNums[i], cbtOuter, True, True);
      Writer.WriteText(R, 4, ARotorNums[i], cbtOuter, True, True);
    end;
  end;

  Writer.SetFrozenRows(FrozenCount);

  Writer.EndEdit;
end;

{ TReclamationSheet }

function TReclamationSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    30,  //№п/п
    85,  // Дата уведомления
    160, // Наименование двигателя
    60,  // Номер двигателя
    80,  // Дата сборки
    80,  // Пробег, км
    110, // Предприятие
    110, // Завод
    110, // Выезд/ФИО
    120, // Неисправный элемент
    120, // Причина неисправности
    60,  // Особое мнение
    350, // Примечание
    80,  //Прибыл в ремонт
    80   // Убыл из ремонта
  ]);
end;

procedure TReclamationSheet.DrawTitle;
var
  R: Integer;
begin
  R:= 1;
  Writer.SetBackgroundClear;
  //Writer.SetBackground(COLOR_BACKGROUND_TITLE);

  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);

  Writer.WriteText(R, 1, '№ п/п', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Дата уведомления', cbtOuter, True, True);
  Writer.WriteText(R, 3, 'Двигатель', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Номер', cbtOuter, True, True);
  Writer.WriteText(R, 5, 'Дата сборки', cbtOuter, True, True);
  Writer.WriteText(R, 6, 'Пробег, км', cbtOuter, True, True);
  Writer.WriteText(R, 7, 'Предприятие', cbtOuter, True, True);
  Writer.WriteText(R, 8, 'Завод', cbtOuter, True, True);
  Writer.WriteText(R, 9, 'Выезд', cbtOuter, True, True);
  Writer.WriteText(R, 10, 'Неисправный элемент', cbtOuter, True, True);
  Writer.WriteText(R, 11, 'Причина неисправности', cbtOuter, True, True);
  Writer.WriteText(R, 12, 'Особое мнение', cbtOuter, True, True);
  Writer.WriteText(R, 13, 'Примечание', cbtOuter, True, True);
  Writer.WriteText(R, 14, 'Прибыл в ремонт', cbtOuter, True, True);
  Writer.WriteText(R, 15, 'Убыл из ремонта', cbtOuter, True, True);
end;

procedure TReclamationSheet.DrawLine(const AIndex: Integer; const ASelected: Boolean);
var
  R: Integer;
  S: String;
begin
  R:= 2 + AIndex;

  Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
  Writer.SetAlignment(haCenter, vaTop);
  Writer.SetBackground(FReasonColors[AIndex]);

  Writer.WriteNumber(R, 1, AIndex+1, cbtOuter);

  if ASelected then
    Writer.SetBackground(DefaultSelectionBGColor)
  else
    Writer.SetBackgroundClear;

  Writer.WriteDate(R, 2, FRecDates[AIndex], cbtOuter);
  Writer.SetAlignment(haLeft, vaTop);
  Writer.WriteText(R, 3, FMotorNames[AIndex], cbtOuter);
  Writer.SetAlignment(haCenter, vaTop);
  Writer.WriteText(R, 4, FMotorNums[AIndex], cbtOuter);
  Writer.WriteDate(R, 5, FBuildDates[AIndex], cbtOuter);
  if FMileages[AIndex]>=0 then
    Writer.WriteNumber(R, 6, FMileages[AIndex], cbtOuter)
  else
    Writer.WriteText(R, 6, EmptyStr, cbtOuter);
  Writer.WriteText(R, 7, FPlaceNames[AIndex], cbtOuter, True, True);
  Writer.WriteText(R, 8, FFactoryNames[AIndex], cbtOuter, True, True);
  Writer.WriteText(R, 9, FDepartures[AIndex], cbtOuter, True, True);
  Writer.WriteText(R, 10, FDefectNames[AIndex], cbtOuter, True, True);
  Writer.WriteText(R, 11, FReasonNames[AIndex], cbtOuter, True, True);

  Writer.SetFont(Font.Name, Font.Size+3, [fsBold], clBlack);
  S:= EmptyStr;
  if FOpinions[AIndex]=1 then S:= CHECK_SYMBOL;
  Writer.WriteText(R, 12, S, cbtOuter);

  Writer.SetFont(Font.Name, Font.Size, [{fsBold}], clBlack);
  Writer.SetAlignment(haLeft, vaTop);
  Writer.WriteText(R, 13, FRecNotes[AIndex], cbtOuter, True, True);

  Writer.SetAlignment(haCenter, vaTop);
  if FArrivalDates[AIndex]=0 then
    Writer.WriteText(R, 14, EmptyStr, cbtOuter, True, True)
  else
    Writer.WriteDate(R, 14, FArrivalDates[AIndex], cbtOuter);

  if FSendingDates[AIndex]=0 then
    Writer.WriteText(R, 15, EmptyStr, cbtOuter, True, True)
  else
    Writer.WriteDate(R, 15, FSendingDates[AIndex], cbtOuter);
end;

procedure TReclamationSheet.Draw(const ARecDates, ABuildDates,
                   AArrivalDates, ASendingDates: TDateVector;
                   const AMileages, AOpinions, AReasonColors: TIntVector;
                   const APlaceNames, AFactoryNames, ADepartures,
                         ADefectNames, AReasonNames, ARecNotes,
                         AMotorNames, AMotorNums: TStrVector);
var
  i: Integer;
begin
  Writer.BeginEdit;

  FRecDates:= ARecDates;
  FBuildDates:= ABuildDates;
  FMileages:= AMileages;
  FOpinions:= AOpinions;
  FPlaceNames:= APlaceNames;
  FFactoryNames:= AFactoryNames;
  FDepartures:= ADepartures;
  FDefectNames:= ADefectNames;
  FReasonNames:= AReasonNames;
  FRecNotes:= ARecNotes;
  FMotorNames:= AMotorNames;
  FMotorNums:= AMotorNums;
  FReasonColors:= AReasonColors;
  FArrivalDates:= AArrivalDates;
  FSendingDates:= ASendingDates;

  DrawTitle;

  if VIsNil(FRecDates) then Exit;

  for i:= 0 to High(FRecDates) do
    DrawLine(i, False);
  Writer.SetBackgroundClear;
  Writer.SetFrozenRows(1);
  Writer.SetRepeatedRows(1,1);

  Writer.EndEdit;
end;

{ TCargoSheet }

function TCargoSheet.SetWidths: TIntVector;
begin
  Result:= VCreateInt([
    300, // Наименование
    130, // Количество      номера
    100, //                 номера
    100  // Номер партии
  ]);
end;

procedure TCargoSheet.Draw(const ASendDate: TDate; const AReceiverName: String;
                   const AMotorNames: TStrVector; const AMotorsCount: TIntVector;
                   const AMotorNums, ASeries: TStrMatrix);
var
  R, RR, i, j: Integer;
  S: String;
const
  FONT_SIZE_DELTA = 0;
begin
  Writer.BeginEdit;

  //Заголовок
  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  S:= FormatDateTime('dd.mm.yyyy', ASendDate) +
      ' Отгрузка в ' + AReceiverName;
  Writer.WriteText(R, 1, R, 4, S, cbtNone, True, True);

  //наименование + количество
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size+FONT_SIZE_DELTA, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, 'Наименование двигателей', cbtOuter, True, True);
  Writer.WriteText(R, 2, 'Количество', cbtOuter, True, True);

  Writer.SetFont(Font.Name, Font.Size+FONT_SIZE_DELTA, [{fsBold}], clBlack);
  for i:= 0 to High(AMotorNames) do
  begin
    R:= R + 1;
    Writer.SetAlignment(haLeft, vaCenter);
    Writer.WriteText(R, 1, AMotorNames[i], cbtOuter, True, True);
    Writer.SetAlignment(haCenter, vaCenter);
    Writer.WriteNumber(R, 2, AMotorsCount[i], cbtOuter);
  end;
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size+FONT_SIZE_DELTA, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, 'ИТОГО', cbtOuter, True, True);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteNumber(R, 2, VSum(AMotorsCount), cbtOuter);

  //раскладка по номерам + партиям
  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size+FONT_SIZE_DELTA, [fsBold], clBlack);
  Writer.SetAlignment(haCenter, vaCenter);
  Writer.WriteText(R, 1, 'Наименование двигателей', cbtOuter, True, True);
  Writer.WriteText(R, 2, R, 3, 'Номера двигателей', cbtOuter, True, True);
  Writer.WriteText(R, 4, 'Партия', cbtOuter, True, True);

  Writer.SetFont(Font.Name, Font.Size+FONT_SIZE_DELTA, [{fsBold}], clBlack);
  for i:= 0 to High(AMotorNames) do
  begin
    R:= R + 1;
    for j:= 0 to High(AMotorNums[i]) do
    begin
      RR:= R + j;
      //Writer.SetAlignment(haCenter, vaTop);
      Writer.SetAlignment(haLeft, vaTop);
      Writer.WriteText(RR, 2, RR, 3,  AMotorNums[i][j], cbtOuter, True, True);
      Writer.SetAlignment(haCenter, vaTop);
      Writer.WriteText(RR, 4, ASeries[i][j], cbtOuter, True, True);
    end;
    Writer.SetAlignment(haLeft, vaTop);
    Writer.WriteText(R, 1, R+High(AMotorNums[i]), 1, AMotorNames[i], cbtOuter, True, True);
    R:= R + High(AMotorNums[i]);
  end;

  R:= R + 1;
  Writer.WriteText(R, 1, R, 4, EmptyStr, cbtTop);

  Writer.EndEdit;
end;

end.

