unit USheets;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics,  fpstypes, fpspreadsheetgrid, DateUtils,
  //DK packages utils
  DK_Vector, DK_Matrix, {DK_Fonts,} DK_SheetConst, DK_SheetWriter,
  DK_SheetTables, DK_Color, DK_SheetTypes;

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

implementation

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
  Writer.RowHeight[R]:= 10;
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
  Writer.RowHeight[R]:= 10;
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
  Writer.RowHeight[R]:= 10;
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size+1, [fsBold], clBlack);
  Writer.SetAlignment(haLeft, vaCenter);
  Writer.WriteText(R, 1, R, Writer.ColCount,  'Контроль:');
  R:= R + 1;
  Writer.SetFont(Font.Name, Font.Size, [], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount,  AControlNote, cbtOuter, True, True);

  R:= R + 1;
  Writer.WriteText(R, 1,  EmptyStr);
  Writer.RowHeight[R]:= 10;
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
  Writer.RowHeight[R]:= 10;
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

