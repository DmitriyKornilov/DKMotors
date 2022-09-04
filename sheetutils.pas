unit SheetUtils;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DK_SheetWriter, fpstypes, fpspreadsheetgrid,
  DK_Vector, DK_Matrix, DK_Fonts, Grids, Graphics, DK_SheetConst, DK_Const,
  DateUtils{, Forms};

const
  COLOR_BACKGROUND_TITLE = clBtnFace;
  COLOR_BACKGROUND_SELECTED   = $00FBDEBB;

  SHEET_FONT_NAME = 'Arial';//'Times New Roman';
  SHEET_FONT_SIZE = 9;

  CHECK_SYMBOL = '✓';

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

  { TReportReclamationSheet }

  TReportReclamationSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;
  public
    constructor Create(const AGrid: TsWorksheetGrid; const AReasonsCount: Integer);
    destructor  Destroy; override;
    procedure DrawReport(const ASingleMotorName: Boolean;
                   const ABeginDate, AEndDate: TDate;
                   const ATotalMotorNames: TStrVector;
                   const ATotalMotorCount: TIntVector;
                   const ATitleReasonNames: TStrVector;
                   const ATotalMotorReasonCounts: TIntMatrix;
                   const APlaceNames: TStrVector;
                   const APlaceMotorCount: TIntVector;
                   const APlaceReasonCounts: TIntMatrix;
                   const ADefectNames: TStrVector;
                   const ADefectMotorCounts: TIntVector;
                   const ADefectReasonCounts: TIntMatrix);

  end;



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
    procedure Draw(const ATestDates: TDateVector;
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

  { TMotorInfoSheet }

  TMotorInfoSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure Draw(const ABuildDate, ASendDate: TDate;
      const AMotorName, AMotorNum, ASeries, ARotorNum, AReceiverName: String;
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
    Mileages, Opinions, ReasonColors, Passports: TIntVector;
    PlaceNames, FactoryNames, Departures: TStrVector;
    DefectNames, ReasonNames, RecNotes: TStrVector;
    MotorNames, MotorNums: TStrVector;

    procedure DrawTitle;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure DrawLine(const AIndex: Integer; const ASelected: Boolean);
    procedure Draw(const ARecDates, ABuildDates, AArrivalDates, ASendingDates: TDateVector;
                   const AMileages, AOpinions, AReasonColors, APassports: TIntVector;
                   const APlaceNames, AFactoryNames, ADepartures,
                         ADefectNames, AReasonNames, ARecNotes,
                         AMotorNames, AMotorNums: TStrVector);
  end;

implementation

{ TReportReclamationSheet }

constructor TReportReclamationSheet.Create(const AGrid: TsWorksheetGrid;
                                           const AReasonsCount: Integer);
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
    300, // Наименование двигателя
    200  // Количество
  ]);

  if AReasonsCount>0 then
    VReDim(ColWidths, Length(ColWidths) + AReasonsCount, 230);

  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TReportReclamationSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TReportReclamationSheet.DrawReport(const ASingleMotorName: Boolean;
   const ABeginDate, AEndDate: TDate; const ATotalMotorNames: TStrVector;
   const ATotalMotorCount: TIntVector;
   const ATitleReasonNames: TStrVector;
   const ATotalMotorReasonCounts: TIntMatrix;
   const APlaceNames: TStrVector;
   const APlaceMotorCount: TIntVector;
   const APlaceReasonCounts: TIntMatrix;
   const ADefectNames: TStrVector;
   const ADefectMotorCounts: TIntVector;
   const ADefectReasonCounts: TIntMatrix);
var
  S: String;
  R: Integer;

  procedure DrawPart(const APartName, AFirstColName: String;
    const ANameValues: TStrVector; const ACountValues: TIntVector;
    const AReasonCountValues: TIntMatrix;
    const AResumeNeed: Boolean);
  var
    i, j: Integer;
  begin
    R:= R + 1;
    FWriter.WriteText(R, 1, R, FWriter.ColCount, EmptyStr, cbtNone);
    //FWriter.SetRowHeight(R, 10);
    R:= R + 1;
    FWriter.SetAlignment(haLeft, vaCenter);
    FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
    FWriter.WriteText(R, 1, R, FWriter.ColCount, APartName, cbtNone, True, True);
    R:= R + 1;
    FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
    FWriter.SetAlignment(haCenter, vaCenter);
    FWriter.WriteText(R, 1, R+1, 1, AFirstColName, cbtOuter, True, True);
    FWriter.WriteText(R, 2, R+1, 2, 'Количество случаев', cbtOuter, True, True);
    FWriter.WriteText(R, 3, R, 3+High(ATitleReasonNames), 'В том числе по причине', cbtOuter, True, True);
    R:= R + 1;
    for i:= 0 to High(ATitleReasonNames) do
      FWriter.WriteText(R, 3+i, ATitleReasonNames[i], cbtOuter, True, True);

    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
    for i:= 0 to High(ANameValues) do
    begin
      R:= R + 1;
      FWriter.SetAlignment(haLeft, vaCenter);
      FWriter.WriteText(R, 1, ANameValues[i], cbtOuter, True, True);
      FWriter.SetAlignment(haCenter, vaCenter);
      FWriter.WriteNumber(R, 2, ACountValues[i], cbtOuter);
      for j:= 0 to High(ATitleReasonNames) do
        FWriter.WriteNumber(R, 3+j, AReasonCountValues[j, i], cbtOuter);
    end;
    if AResumeNeed then
    begin
      R:= R + 1;
      FWriter.SetFont(FFontName, FFontSize, [fsBold], clBlack);
      FWriter.SetAlignment(haLeft, vaCenter);
      FWriter.WriteText(R, 1, 'ИТОГО', cbtOuter, True, True);
      FWriter.SetAlignment(haCenter, vaCenter);
      FWriter.WriteNumber(R, 2, VSum(ACountValues), cbtOuter);
      for i:= 0 to High(ATitleReasonNames) do
        FWriter.WriteNumber(R, 3+i, VSum(AReasonCountValues[i]), cbtOuter);
    end;
  end;

begin
  FGrid.Clear;
  if VIsNil(ATotalMotorNames) then Exit;

  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;

  S:= 'Отчет по рекламационным случаям за ';
  if SameDate(ABeginDate, AEndDate) then
    S:= S + FormatDateTime('dd.mm.yyyy', ABeginDate)
  else
    S:= S + 'период с ' + FormatDateTime('dd.mm.yyyy', ABeginDate) +
       ' по ' + FormatDateTime('dd.mm.yyyy', AEndDate);

  //Заголовок отчета
  R:= 1;
  FWriter.SetFont(FFontName, FFontSize+2, [fsBold], clBlack);
  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.WriteText(R, 1, R, FWriter.ColCount, S, cbtNone, True, True);

  DrawPart('Общее количество рекламационных случаев', 'Электродвигатель', ATotalMotorNames, ATotalMotorCount,
            ATotalMotorReasonCounts, not ASingleMotorName);

  DrawPart('Распределение по неисправным элементам', 'Неисправный элемент',
           ADefectNames, ADefectMotorCounts, ADefectReasonCounts,  False);

  DrawPart('Распределение по предприятиям', 'Предприятие',
           APlaceNames, APlaceMotorCount, APlaceReasonCounts, False);



  FWriter.EndEdit;
end;

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
    100, // № п/п - Дата отправки
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
    100, // Дата сборки
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

{ TMotorInfoSheet }

constructor TMotorInfoSheet.Create(const AGrid: TsWorksheetGrid);
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
    80, // Пробег, км
    150, // Предприятие
    80, // Завод
    130, // Выезд/ФИО
    180, // Неисправный элемент
    200,  // Причина неисправности
    100,  //Особое мнение
    350  //Примечание
  ]);
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TMotorInfoSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TMotorInfoSheet.Draw(const ABuildDate, ASendDate: TDate;
      const AMotorName, AMotorNum, ASeries, ARotorNum, AReceiverName: String;
      const ATestDates: TDateVector;
      const ATestResults: TIntVector;
      const ATestNotes: TStrVector;
      const ARecDates: TDateVector;
      const AMileages, AOpinions: TIntVector;
      const APlaceNames, AFactoryNames, ADepartures,
      ADefectNames, AReasonNames, ARecNotes: TStrVector);
const
  FONTSIZE_DELTA = 2;
var
  R, i: Integer;
  S: String;

  procedure DrawLine(const ARow: Integer; const AItemName, AItemValue: String);
  begin
    FWriter.SetFont(FFontName, FFontSize+FONTSIZE_DELTA, [fsBold], clBlack);
    FWriter.WriteText(ARow, 1,{ ARow, 2,} AItemName, cbtNone, True, True);
    FWriter.SetFont(FFontName, FFontSize+FONTSIZE_DELTA, [{fsBold}], clBlack);
    FWriter.WriteText(ARow, 2, ARow, 9, AItemValue, cbtNone, True, True);
  end;



begin
  FGrid.Clear;
  FWriter.BeginEdit;
  FWriter.SetBackgroundClear;
  FWriter.SetAlignment(haLeft, vaCenter);

  R:= 1;
  S:= AMotorName + ' № ' + AMotorNum;
  if ASeries<>EmptyStr then
    S:= S + ' (' + ASeries + ')';
  DrawLine(R, 'Двигатель:', S);
  R:= R + 1;
  DrawLine(R, 'Дата сборки: ', FormatDateTime('dd.mm.yyyy', ABuildDate));
  R:= R + 1;
  DrawLine(R, 'Номер ротора:', ARotorNum);
  R:= R + 1;
  if AReceiverName=EmptyStr then
    S:= 'нет'
  else
    S:= FormatDateTime('dd.mm.yyyy', ASendDate) + ' - ' + AReceiverName;
  DrawLine(R, 'Отгружен:', S);
  R:= R + 1;

  //R:= R + 1;
  DrawLine(R, 'Испытания:', EmptyStr);
  FWriter.SetAlignment(haCenter, vaCenter);
  R:= R + 1;
  FWriter.SetFont(FFontName, FFontSize{+FONTSIZE_DELTA}, [fsBold], clBlack);
  FWriter.WriteText(R, 1, 'Дата', cbtOuter, True, True);
  FWriter.WriteText(R, 2, {R, 3,} 'Результат', cbtOuter, True, True);
  FWriter.WriteText(R, 3, R, 5, 'Примечание', cbtOuter, True, True);
  FWriter.WriteText(R, 6, EmptyStr, cbtLeft);

  FWriter.SetFont(FFontName, FFontSize{+FONTSIZE_DELTA}, [{fsBold}], clBlack);

  if VIsNil(ATestDates) then
  begin
    R:= R + 1;
    FWriter.WriteText(R, 1, EmptyStr, cbtOuter);
    FWriter.WriteText(R, 2, EmptyStr, cbtOuter, True, True);
    FWriter.WriteText(R, 3, R, 5, EmptyStr, cbtOuter, True, True);
    FWriter.WriteText(R, 6, EmptyStr, cbtLeft);
  end
  else begin
    for i:= 0 to High(ATestDates) do
    begin
      R:= R + 1;
      if ATestResults[i]=0 then
      begin
        FWriter.SetFont(FFontName, FFontSize{+FONTSIZE_DELTA}, [{fsBold}], clBlack);
        S:= 'норма';
      end
      else begin
        FWriter.SetFont(FFontName, FFontSize{+FONTSIZE_DELTA}, [{fsBold}], clRed);
        S:= 'брак';
      end;
      FWriter.WriteDate(R, 1, ATestDates[i], cbtOuter);
      FWriter.WriteText(R, 2, {R, 3,} S, cbtOuter, True, True);
      FWriter.WriteText(R, 3, R, 5, ATestNotes[i], cbtOuter, True, True);
      FWriter.WriteText(R, 6, EmptyStr, cbtLeft);
    end;
  end;

  R:= R + 1;
  FWriter.SetAlignment(haLeft, vaCenter);
  DrawLine(R, 'Рекламации:', EmptyStr);
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

  if VIsNil(ARecDates) then
  begin
    R:= R + 1;
    FWriter.SetBackgroundClear;
    FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
    for i:= 1 to 9 do
      FWriter.WriteText(R, i, EmptyStr, cbtOuter);
  end
  else begin
    for i:= 0 to High(ARecDates) do
    begin
      R:= R + 1;
      FWriter.SetBackgroundClear;
      FWriter.SetFont(FFontName, FFontSize, [{fsBold}], clBlack);
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
  FWriter.SetBackgroundClear;

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
    100, // Дата сдачи (испытаний)
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

procedure TStoreSheet.Draw(const ATestDates: TDateVector;
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
    100, // Дата испытаний
    300, // Наименование двигателя
    150, // Номер двигателя
    //150, // Номер партии
    150, // Результат (норма/брак)
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
    100, // № п/п - Дата сборки
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
  FWriter.WriteText(R, 15, 'Наличие паспорта', cbtOuter, True, True);
  FWriter.WriteText(R, 16, 'Убыл из ремонта', cbtOuter, True, True);

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
    40,  //№п/п
    90,  // Дата уведомления
    150, // Наименование двигателя
    60,  // Номер двигателя
    70,  // Дата сборки
    60,  // Пробег, км
    130, // Предприятие
    100,  // Завод
    100, // Выезд/ФИО
    120, // Неисправный элемент
    130, // Причина неисправности
    50,  // Особое мнение
    450, // Примечание
    70,  //Прибыл в ремонт
    60,  // Наличие паспорта
    70   // Убыл из ремонта
  ]);

  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);

end;

destructor TReclamationSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
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

  S:= EmptyStr;
  if Passports[AIndex]=1 then S:= CHECK_SYMBOL;
  FWriter.WriteText(R, 15, S, cbtOuter);

  if SendingDates[AIndex]=0 then
    FWriter.WriteText(R, 16, EmptyStr, cbtOuter, True, True)
  else
    FWriter.WriteDate(R, 16, SendingDates[AIndex], cbtOuter);




end;

procedure TReclamationSheet.Draw(const ARecDates, ABuildDates,
                   AArrivalDates, ASendingDates: TDateVector;
                   const AMileages, AOpinions, AReasonColors, APassports: TIntVector;
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
  Passports:= APassports;

  DrawTitle;

  if not VIsNil(RecDates) then
  begin
    for i:= 0 to High(RecDates) do
      DrawLine(i, False);
    FWriter.SetBackgroundClear;
    for i:= 1 to 16 do
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
    100, // Количество      номера
    90, //                 номера
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

