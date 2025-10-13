unit UDataBase;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, StdCtrls, DateUtils,
  //DK packages utils
  DK_Vector, DK_Matrix, DK_Const, DK_StrUtils, DK_SQLite3, DK_SQLUtils, DK_DateUtils,
  //Project utils
  UCalendar, UStatistic;

type

  { TDataBase }

  TDataBase = class (TSQLite3)
  //private
    {function ReclamationReportLoad(const ATableName, AIDFieldName, ANameFieldName: String;
       const ABeginDate, AEndDate: TDate; const ANameIDs: TIntVector;
       out AExistsIDs: TIntVector; out ANames: TStrVector; out ACounts: TIntVector): Boolean;  }

    {function ReclamationReportWithReasonsLoad(const ATableName, AIDFieldName, ANameFieldName: String;
       const ABeginDate, AEndDate: TDate;
       const AAdditionYearsCount: Integer;
       //const AIsZeroIDNeed: Boolean;
       const ANameIDs, AReasonIDs: TIntVector;
       out AParamNames: TStrVector; out AParamCounts: TIntMatrix3D): Boolean; }
  public
    {КАЛЕНДАРЬ}
    //особые даты производственного календаря
    function LoadCalendarSpecDays(const ABeginDate, AEndDate: TDate): TCalendarSpecDays;
    function WriteCalendarSpecDays(const ADates: TDateVector; const AStatus: Integer): Boolean;
    procedure DeleteCalendarSpecDay(const ADate: TDate);
    //производственный календарь
    function LoadCalendar(const ABeginDate, AEndDate: TDate): TCalendar;
    //последняя дата срока, заданного в рабочих днях
    function LoadWorkDaysPeriodEndDate(const ABeginDate: TDate;
                              const AWorkDaysCount: Integer): TDate;
    //количество рабочих дней за период
    function LoadWorkDaysCountInPeriod(const ABeginDate, AEndDate: TDate): Integer;

    //справочники
    procedure NameIDsAndMotorNamesLoad(AComboBox: TComboBox;
                  out AIDs: TIntVector; const AKeyValueNotZero: Boolean = True);
    procedure ReceiverIDsAndNamesLoad(AComboBox: TComboBox;
                  out AIDs: TIntVector; const AKeyValueNotZero: Boolean = True);
    function NameIDsAndMotorNamesSelectedLoad(ALabel: TLabel;
                  const ANeedEdit: Boolean;
                  var ANameIDs: TIntVector; var AMotorNames: TStrVector): Boolean;
    function ReceiverIDsAndNamesSelectedLoad(ALabel: TLabel;
                  const ANeedEdit: Boolean;
                  var AReceiverIDs: TIntVector; var AReceiverNames: TStrVector): Boolean;

    //дерево месяцы-даты для журналов учета
    function MonthAndDatesForBuildLogLoad(const AYear: Word;
                  const AUsedNameIDs: TIntVector;
                  out AMonths: TStrVector; out ADates: TDateMatrix): Boolean;
    function MonthAndDatesForTestLogLoad(const AYear: Word;
                  const AUsedNameIDs: TIntVector;
                  out AMonths: TStrVector; out ADates: TDateMatrix): Boolean;

    //список двигателей
    function MotorListLoad(const ABuildYear, AShippedType: Integer;
                         const ANameIDs: TIntVector;
                         const ANumberLike: String;
                         const ANeedOrderByNumber: Boolean;
                         out AMotorIDs: TIntVector;
                         out ABuildDates, AMotorNames, AMotorNums,
                         AShippings: TStrVector): Boolean;
    procedure MotorInfoLoad(const AMotorID: Integer;
          out ABuildDate, ASendDate: TDate;
          out AMotorName, AMotorNum, ASeries, ARotorNum, AReceiverName: String;
          out ATestDates: TDateVector;
          out ATestResults: TIntVector;
          out ATestNotes: TStrVector);

    //сборка
    function BuildListLoad(const ABeginDate, AEndDate: TDate;
               const AUsedNameIDs: TIntVector; const ANeedOrderByNumber: Boolean;
               out AMotorIDs, ANameIDs, AOldMotors: TIntVector;
               out ABuildDates: TDateVector;
               out AMotorNames, AMotorNums, ARotorNums: TStrVector): Boolean;
    function BuildTotalLoad(const ABeginDate, AEndDate: TDate;
               const ANameIDs: TIntVector;
               out AMotorNames: TStrVector; out AMotorCounts: TIntVector): Boolean;
    function IsDuplicateMotorNumber(const ADate: TDate;
               const ANameID: Integer; const AMotorNum: String): Boolean;
    function MotorsInBuildLogWrite(const ABuildDate: TDate;
               const ANameIDs, AOldMotors: TIntVector;
               const AMotorNums, ARotorNums: TStrVector): Boolean;
    function MotorInBuildLogUpdate(const AMotorID: Integer;
                          const ABuildDate: TDate; const ANameID, AOldMotor: Integer;
                          const AMotorNum, ARotorNum: String): Boolean;

    //испытания
    function TestBeforeListLoad(const ANameIDs: TIntVector;
                          const ANeedOrderByNumber: Boolean;
                          out ABuildDates: TDateVector;
                          out AMotorNames, AMotorNums, ATotalMotorNames: TStrVector;
                          out ATotalMotorCounts: TIntVector;
                          out ATestDates: TDateMatrix;
                          out ATestFails: TIntMatrix;
                          out ATestNotes: TStrMatrix): Boolean;
    function TestListLoad(const ABeginBuildDate, AEndBuildDate: TDate;
               const ANameIDs: TIntVector; const ANeedOrderByNumber: Boolean;
               out ATestIDs, ATestResults: TIntVector;
               out ATestDates: TDateVector;
               out AMotorNames, AMotorNums, ATestNotes: TStrVector): Boolean;
    function TestTotalLoad(const ABeginBuildDate, AEndBuildDate: TDate;
                const ANameIDs: TIntVector;
                out AMotorNames: TStrVector;
                out AMotorCounts, AFailCounts: TIntVector): Boolean;
    function TestChooseListLoad(const ANameID: Integer;
                         const ANumberLike: String;
                         out AMotorIDs: TIntVector;
                         out AMotorNums, ABuildDates, ATests: TStrVector): Boolean;
    function MotorsInTestLogWrite(const ATestDate: TDate;
                           const AMotorIDs, ATestResults: TIntVector;
                           const ATestNotes: TStrVector): Boolean;
    function TestLastFineDatesStr(const AMotorIDs: TIntVector): TStrVector;

    //контроль
    function ControlNoteLoad(const AMotorID: Integer): String;
    function ControlListLoad(const ANameIDs: TIntVector; const ANumberLike: String;
          out AMotorIDs: TIntVector;
          out AMotorNames, AMotorNums, ASeries, ANotes: TStrVector): Boolean;
    function ControlUpdate(const AMotorID: Integer; const ANote: String): Boolean;
    function ControlDelete(const AMotorID: Integer): Boolean;
    function MotorListToControl(const ANameID: Integer; const ANumberLike: String;
                                out AMotorIDs: TIntVector;
                                out AMotorNums, ASeries: TStrVector): Boolean;
    function MotorListOnControl(const AMotorID: Integer;
                                out AMotorIDs: TIntVector;
                                out AMotorNums, ASeries: TStrVector): Boolean;

    //склад
    function StoreListLoad(const ANameIDs: TIntVector;
                    const ADeltaDays: Integer;
                    const ANeedOrderByNumber, ANeedOrderByName: Boolean;
                    out ATestDates: TDateVector;
                    out AMotorNames, AMotorNums: TStrVector): Boolean;
    function StoreTotalLoad(const ANameIDs: TIntVector;
                            const ADeltaDays: Integer;
                             out AMotorNames: TStrVector;
                             out AMotorCounts: TIntVector): Boolean;

    //отгрузки
    function ShipmentListLoad(const ANameIDs, AReceiverIDs: TIntVector;
                             const AYear: Word; out AMonths: TStrVector;
                          out AShipments: TStrMatrix;
                          out ACargoIDs: TIntMatrix): Boolean;
    function ShipmentMotorListLoad(const ABeginDate, AEndDate: TDate;
                const ANameIDs, AReceiverIDs: TIntVector;
                const ANeedOrderByNumber: Boolean;
                out ASendDates: TDateVector;
                out AMotorNames, AMotorNums, AReceiverNames: TStrVector): Boolean;
    function ShipmentTotalLoad(const ABeginDate, AEndDate: TDate;
                const ANameIDs, AReceiverIDs: TIntVector;
                out AMotorNames: TStrVector;
                out AMotorCounts: TIntVector): Boolean;
    function ShipmentRecieversTotalLoad(const ABeginDate, AEndDate: TDate;
                const ANameIDs, AReceiverIDs: TIntVector;
                out AReceiverNames: TStrVector;
                out AMotorNames: TStrMatrix;
                out AMotorCounts: TIntMatrix): Boolean;
    function CargoListLoad(const ANameIDs, AReceiverIDs: TIntVector;
                         const ABeginDate, AEndDate: TDate;
                         out ACargoIDs: TIntVector;
                         out ASendDates: TDateVector;
                         out AReceiverNames: TStrVector): Boolean;
    function CargoChooseListLoad(const ANameID: Integer;
                         const ANumberLike: String;
                         out AMotorIDs: TIntVector;
                         out AMotorNums, ABuildDates, ASeries: TStrVector): Boolean;
    function CargoMotorListLoad(const ACargoID: Integer;
                         out AReceiverID: Integer;
                         out ASendDate: TDate;
                         out AMotorIDs: TIntVector;
                         out AMotorNames, AMotorNums, ASeries: TStrVector): Boolean;
    function CargoPackingSheet(const ACargoID: Integer;
                         out AMotorNames, AMotorNums, ATestDates: TStrVector): Boolean;
    function CargoLoad(const ACargoID: Integer;
                     out ASendDate: TDate; out AReceiverName: String;
                     out AMotorNames: TStrVector; out AMotorCounts: TIntVector;
                     out AMotorNums, ASeries: TStrMatrix): Boolean;
    function CargoWrite(const ASendDate: TDate;
                      const AReceiverID: Integer;
                      const AMotorIDs: TIntVector;
                      const ASeries: TStrVector): Boolean;
    function CargoUpdate(const ACargoID: Integer;
                      const ASendDate: TDate;
                      const AReceiverID: Integer;
                      const AMotorIDs, AWritedMotorIDs: TIntVector;
                      const ASeries: TStrVector): Boolean;


    //рекламации
    function ReclamationListLoad(const AMotorID: Integer;
                          out ARecDates: TDateVector;
                          out AMileages, AOpinions: TIntVector;
                          out APlaceNames, AFactoryNames, ADepartures,
                          ADefectNames, AReasonNames,
                          ARecNotes: TStrVector): Boolean;
    function ReclamationListLoad(const ABeginDate, AEndDate: TDate;
          const ANameIDs, ADefectIDs, AReasonIDs: TIntVector;
          const AOrderIndex: Integer; const ANumberLike: String;
          out ARecDates, ABuildDates, AArrivalDates, ASendingDates: TDateVector;
          out ARecIDs, AMotorIDs, AMileages, AOpinions, AColors, APassports: TIntVector;
          out APlaceNames, AFactoryNames, ADepartures,
          ADefectNames, AReasonNames, ARecNotes,
          AMotorNames, AMotorNums: TStrVector): Boolean;
    function ReclamationChooseListLoad(const ANameID: Integer;
                         const ANumberLike: String;
                         out AMotorIDs: TIntVector;
                         out AMotorNums, ABuildDates, AShippings: TStrVector): Boolean;
    function ReclamationWrite(const ARecDate: TDate;
                      const AMotorID, AMileage, APlaceID, AFactoryID,
                            ADefectID, AReasonID, AOpinion: Integer;
                      const ADeparture, ARecNote: String): Boolean;
    function ReclamationUpdate(const ARecDate, AArrivalDate, ASendingDate: TDate;
                      const ARecID, AMotorID, AMileage, APlaceID, AFactoryID,
                            ADefectID, AReasonID, AOpinion, APassport: Integer;
                      const ADeparture, ARecNote: String): Boolean;
    function ReclamationLoad(const ARecID: Integer;
                      out ARecDate, AArrivalDate, ASendingDate: TDate;
                      out ANameID, AMotorID, AMileage, APlaceID, AFactoryID,
                          ADefectID, AReasonID, AOpinion, APassport: Integer;
                      out AMotorNum, ADeparture, ARecNote: String): Boolean;

    //ремонты
    function RepairListLoad(const ANameIDs: TIntVector; const ANumberLike: String;
          const AOrderType: Integer; //1 - по дате прибытия, 2- по дате убытия, 3 - по номеру
          const AListType: Integer; //0 - все, 1-в ремонте, 2 - отремонтированные
          out AArrivalDates, ASendingDates: TDateVector;
          out ARecIDs, AMotorIDs, APassports, AWorkDayCounts: TIntVector;
          out AMotorNames, AMotorNums, ARepairNotes: TStrVector): Boolean;
    function RepairListForMotorIDLoad(const AMotorID: Integer;
          out AArrivalDates, ASendingDates: TDateVector;
          out APassports, AWorkDayCounts: TIntVector;
          out ARepairNotes: TStrVector): Boolean;
    function RepairUpdate(const ARecID: Integer;
          const AArrivalDate, ASendingDate: TDate;
          const APassport: Integer; const ARepairNote: String): Boolean;
    function RepairLoad(const ARecID: Integer;
          out AArrivalDate, ASendingDate: TDate;
          out APassport: Integer; out ARepairNote: String): Boolean;
    procedure RepairDelete(const ARecID: Integer);
    function MotorListToRepair(const ANameID: Integer; const AMotorNumberLike: String;
          out ARecIDs, AMotorNameIDs: TIntVector; out ABuildDates, ARecDates: TDateVector;
          out AMotorNums, APlaceNames: TStrVector): Boolean;
    function MotorListOnRepair(const ARecID: Integer;
          out ARecIDs, AMotorNameIDs: TIntVector; out ABuildDates, ARecDates: TDateVector;
          out AMotorNums, APlaceNames: TStrVector): Boolean;


    {--- СТАТИСТИКА РЕКЛАМАЦИЙ ------}
    {//общее количество
    function ReclamationTotalCountLoad(const ABeginDate, AEndDate: TDate;
                const ANameIDs: TIntVector;
                out AExistsNameIDs: TIntVector;
                out AMotorNames: TStrVector;
                out AMotorCounts: TIntVector): Boolean;  }
    {function ReclamationMotorsWithReasonsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs, AReasonIDs: TIntVector;
                const AANEMAsSameName: Boolean;
                out AMotorNames: TStrVector;
                out AClaimCounts: TIntMatrix3D): Boolean; }
    function ReclamationByMotorNamesLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs: TIntVector;
                const AMotorTypeAsSingleName: Boolean;
                const ADoNotUseZeroClaimMotor: Boolean;
                var AMotorNames: TStrVector;
                out AMotorNeeds: TBoolVector;
                out AClaimCounts: TIntMatrix3D): Boolean;

    {function ReclamationDefectsWithReasonsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                //const AIsZeroIDNeed: Boolean;
                const ANameIDs, AReasonIDs: TIntVector;
                out ADefectNames: TStrVector;
                out AMotorCounts: TIntMatrix3D): Boolean;  }
    function ReclamationByDefectsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs: TIntVector;
                const ADoNotUseZeroClaimDefect: Boolean;
                out ADefectNames: TStrVector;
                out ADefectNeeds: TBoolVector;
                out AClaimCounts: TIntMatrix3D): Boolean;


    {function ReclamationPlacesWithReasonsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                //const AIsZeroIDNeed: Boolean;
                const ANameIDs, AReasonIDs: TIntVector;
                out APlaceNames: TStrVector;
                out AMotorCounts: TIntMatrix3D): Boolean;   }
    function ReclamationByPlacesLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs: TIntVector;
                const ADoNotUseZeroClaimPlace: Boolean;
                out APlaceNames: TStrVector;
                out APlaceNeeds: TBoolVector;
                out AClaimCounts: TIntMatrix3D): Boolean;


    {function ReclamationMonthsWithReasonsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const ANameIDs, AReasonIDs: TIntVector;
                out AMonthNames: TStrVector;
                out AMotorCounts{, AAccumCounts}: TIntMatrix3D): Boolean;  }
    function ReclamationByMonthsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs: TIntVector;
                const ADoNotUseZeroClaimMonth: Boolean;
                out AMonthNames: TStrVector;
                out AMonthNeeds: TBoolVector;
                out AClaimCounts: TIntMatrix3D): Boolean;

    function ReclamationMileagesWithReasonsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const ANameIDs, AReasonIDs: TIntVector;
                out AMileageNames: TStrVector;
                out AMotorCounts: TIntMatrix3D): Boolean;

  end;

implementation

{ TDataBase }

function TDataBase.LoadCalendarSpecDays(const ABeginDate, AEndDate: TDate): TCalendarSpecDays;
begin
  Result:= EmptyCalendarSpecDays;
  QSetQuery(FQuery);
  QSetSQL(
    'SELECT * FROM CALENDAR ' +
    'WHERE DayDate BETWEEN :BD AND :ED ' +
    'ORDER BY DayDate');
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(Result.Dates, QFieldDT('DayDate'));
      VAppend(Result.Statuses, QFieldInt('Status'));
      QNext;
    end;
  end;
  QClose;
end;

function TDataBase.WriteCalendarSpecDays(const ADates: TDateVector;
  const AStatus: Integer): Boolean;
var
  i: Integer;
begin
  Result:= False;

  QSetQuery(FQuery);
  try
    QSetSQL('INSERT OR REPLACE INTO CALENDAR (DayDate, Status) ' +
           'VALUES (:DayDate, :Status)');
    QParamInt('Status', AStatus);
    for i:= 0 to High(ADates) do
    begin
      QParamDT('DayDate', ADates[i]);
      QExec;
    end;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

procedure TDataBase.DeleteCalendarSpecDay(const ADate: TDate);
begin
  Delete('CALENDAR', 'DayDate', ADate);
end;

function TDataBase.LoadCalendar(const ABeginDate, AEndDate: TDate): TCalendar;
var
  SpecDays: TCalendarSpecDays;
begin
  SpecDays:= LoadCalendarSpecDays(ABeginDate, AEndDate);
  Result:= TCalendar.Create;
  Result.Calc(ABeginDate, AEndDate, SpecDays);
end;

function TDataBase.LoadWorkDaysPeriodEndDate(const ABeginDate: TDate;
  const AWorkDaysCount: Integer): TDate;
var
  Calendar: TCalendar;
  BD, ED: TDate;
  i, N: Integer;
begin
  Result:= ABeginDate;
  if AWorkDaysCount<1 then Exit;

  BD:= ABeginDate;
  N:= 0;
  while N<AWorkDaysCount do
  begin
    Calendar:= LoadCalendar(BD, IncDay(BD, AWorkDaysCount));
    try
      for i:=0 to Calendar.DaysCount-1 do
      begin
        ED:= Calendar.Dates[i];
        if Calendar.IsWorkDay[i] then
        begin
          N:= N + 1;
          if N=AWorkDaysCount then break;
        end;
      end;
      BD:= IncDay(ED, 1);
    finally
      FreeAndNil(Calendar);
    end;
  end;
  Result:= ED;
end;

function TDataBase.LoadWorkDaysCountInPeriod(const ABeginDate, AEndDate: TDate): Integer;
var
  Calendar: TCalendar;
  D: TDate;
begin
  Result:= 0;
  if ABeginDate=0 then Exit;

  D:= AEndDate;
  if D=0 then D:= Date;

  Calendar:= LoadCalendar(ABeginDate, D);
  try
    Result:= Calendar.WorkDaysCount;
  finally
    FreeAndNil(Calendar);
  end;
end;


function TDataBase.NameIDsAndMotorNamesSelectedLoad(ALabel: TLabel;
  const ANeedEdit: Boolean; var ANameIDs: TIntVector;
  var AMotorNames: TStrVector): Boolean;
begin
  Result:= LoadIDsAndNamesSelected(ALabel, ANeedEdit, ANameIDs, AMotorNames,
                          'Наименования электродвигателей',
                          'MOTORNAMES', 'NameID', 'MotorName', 'NameID',
                          True, 'ВСЕ НАИМЕНОВАНИЯ', False);
end;

function TDataBase.ReceiverIDsAndNamesSelectedLoad(ALabel: TLabel;
  const ANeedEdit: Boolean; var AReceiverIDs: TIntVector;
  var AReceiverNames: TStrVector): Boolean;
begin
  Result:= LoadIDsAndNamesSelected(ALabel, ANeedEdit, AReceiverIDs, AReceiverNames,
                          'Грузополучатели',
                          'CARGORECEIVERS', 'ReceiverID', 'ReceiverName', 'ReceiverName',
                          True, 'ВСЕ ГРУЗОПОЛУЧАТЕЛИ', False);
end;

procedure TDataBase.NameIDsAndMotorNamesLoad(AComboBox: TComboBox;
  out AIDs: TIntVector; const AKeyValueNotZero: Boolean = True);
begin
  KeyPickLoad(AComboBox, AIDs,
                  'MOTORNAMES', 'NameID', 'MotorName', 'NameID',
                  AKeyValueNotZero{, 'ВСЕ НАИМЕНОВАНИЯ'});
end;

procedure TDataBase.ReceiverIDsAndNamesLoad(AComboBox: TComboBox;
  out AIDs: TIntVector; const AKeyValueNotZero: Boolean = True);
begin
  KeyPickLoad(AComboBox, AIDs,
                  'CARGORECEIVERS', 'ReceiverID', 'ReceiverName', 'ReceiverName',
                  AKeyValueNotZero{, 'ВСЕ ГРУЗОПОЛУЧАТЕЛИ'});
end;

{function TDataBase.ReclamationReportWithReasonsLoad(const ATableName,
  AIDFieldName, ANameFieldName: String; const ABeginDate, AEndDate: TDate;
  const AAdditionYearsCount: Integer; {const AIsZeroIDNeed: Boolean; }
  const ANameIDs, AReasonIDs: TIntVector; out AParamNames: TStrVector; out
  AParamCounts: TIntMatrix3D): Boolean;
var
  ExistNameIDs: TIntVector;
  ParamCounts: TIntMatrix;
  BD, ED: TDate;
  i: Integer;

  procedure GetParamNamesSinglePeriod(const ABD, AED: TDate;
                          out AExistNames: TStrVector;
                          out AExistNameIDs: TIntVector);
  var
    WhereStr, KeyFieldName, PickFieldName: String;
  begin
    Result:= False;
    AExistNames:= nil;
    AExistNameIDs:= nil;

    KeyFieldName:= SqlEsc(AIDFieldName, False);
    PickFieldName:= SqlEsc(ANameFieldName, False);

    WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) ';
    if not VIsNil(ANameIDs) then
      WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

    //if not AIsZeroIDNeed then
    //  WhereStr:= WhereStr + 'AND (t3.' + KeyFieldName + '>0) ';

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t1.' + KeyFieldName + ') AS MotorCount, t1.' + KeyFieldName + ', t3.' + PickFieldName + ' ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      'INNER JOIN ' + SqlEsc(ATableName) + ' t3 ON (t1.' + KeyFieldName + '=t3.' + KeyFieldName + ') ' +
      WhereStr +
      'GROUP BY t1.' + KeyFieldName + ' ' +
      'ORDER BY t3.' + PickFieldName);
    QParamDT('BD', ABD);
    QParamDT('ED', AED);
    if not VIsNil(ANameIDs) then
      QParamsInt(ANameIDs);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        VAppend(AExistNameIDs, QFieldInt(AIDFieldName));
        VAppend(AExistNames, QFieldStr(ANameFieldName));
        QNext;
      end;
      Result:= True;
    end;
    QClose;
  end;

  procedure GetParamNamesAllPeriods(const ABD, AED: TDate;
                          const AAdditionYearsCount: Integer;
                          out AExistNames: TStrVector;
                          out AExistNameIDs: TIntVector);
  var
    ExistNames: TStrVector;
    ExistNameIDs: TIntVector;
    n: Integer;
  begin
    AExistNames:= nil;
    AExistNameIDs:= nil;
    for n:= AAdditionYearsCount downto 0 do
    begin
      GetParamNamesSinglePeriod(IncYear(ABD, -n), IncYear(AED, -n),
                                ExistNames, ExistNameIDs);
      AExistNames:= VUnion(AExistNames, ExistNames);
      AExistNameIDs:= VUnion(AExistNameIDs, ExistNameIDs);
    end;
  end;

  procedure GetDataForReason(const ABD, AED: TDate; const AReasonID: Integer; out ACounts: TIntVector);
  var
    KeyFieldName, PickFieldName: String;
    n: Integer;
    WhereStr: String;
  begin
    ACounts:= nil;
    VDim(ACounts, Length(ExistNameIDs), 0);

    KeyFieldName:= SqlEsc(AIDFieldName, False);
    PickFieldName:= SqlEsc(ANameFieldName, False);

    WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
               SqlIN('t2','NameID', Length(ANameIDs), 'NameID') + 'AND' +
               SqlIN('t1',AIDFieldName, Length(ExistNameIDs), AIDFieldName);

    if AReasonID>=0 then
      WhereStr:= WhereStr + ' AND (t1.ReasonID = :ReasonID) ';

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t1.' + KeyFieldName + ') As MotorCount, t1.' + KeyFieldName + ', t3.' + PickFieldName + ' ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      'INNER JOIN ' + SqlEsc(ATableName) + ' t3 ON (t1.' + KeyFieldName + '=t3.' + KeyFieldName + ') ' +
      WhereStr +
      'GROUP BY t1.' + KeyFieldName + ' ' +
      'ORDER BY t3.' + PickFieldName);
    QParamDT('BD', ABD);
    QParamDT('ED', AED);
    if AReasonID>=0 then
      QParamInt('ReasonID', AReasonID);
    QParamsInt(ANameIDs, 'NameID');
    QParamsInt(ExistNameIDs, AIDFieldName);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        n:= VIndexOf(ExistNameIDs, QFieldInt(AIDFieldName));
        if n>=0 then
          ACounts[n]:= QFieldInt('MotorCount');
        QNext;
      end;
    end;
    QClose;
  end;

  procedure GetDataForSinglePeriod(const ABD, AED: TDate; out ACounts: TIntMatrix);
  var
    Counts: TIntVector;
    n: Integer;
  begin
    ACounts:= nil;
    for n:=0 to High(AReasonIDs) do
    begin
      GetDataForReason(ABD, AED, AReasonIDs[n], Counts);
      MAppend(ACounts, Counts);
    end;
  end;

begin
  Result:= False;
  AParamNames:= nil;
  AParamCounts:= nil;

  if VIsNil(AReasonIDs) then Exit;

  //список наименований параметра распределения с существующими рекламациями, их общее количество за все годы
  GetParamNamesAllPeriods(ABeginDate, AEndDate, AAdditionYearsCount,
                          AParamNames, ExistNameIDs);
  //нет рекламаций в запрашиваемые периоды
  if VIsNil(ExistNameIDs) then Exit;

  //матрица данных для каждого периода
  for i:= AAdditionYearsCount downto 0 do
  begin
    BD:= IncYear(ABeginDate, -i);
    ED:= IncYear(AEndDate, -i);
    GetDataForSinglePeriod(BD, ED, ParamCounts);
    //итоговая матрица для всех периодов
    MAppend(AParamCounts, ParamCounts);
  end;

  Result:= True;
end; }


{function TDataBase.ReclamationReportLoad(const ATableName, AIDFieldName,
  ANameFieldName: String; const ABeginDate, AEndDate: TDate;
  const ANameIDs: TIntVector; out AExistsIDs: TIntVector;
  out ANames: TStrVector; out ACounts: TIntVector): Boolean;
var
  WhereStr, KeyFieldName, PickFieldName: String;
begin
  Result:= False;
  ANames:= nil;
  ACounts:= nil;
  AExistsIDs:= nil;

  KeyFieldName:= SqlEsc(AIDFieldName, False);
  PickFieldName:= SqlEsc(ANameFieldName, False);

  WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(t1.' + KeyFieldName + ') AS MotorCount, t1.' + KeyFieldName + ', t3.' + PickFieldName + ' ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN ' + SqlEsc(ATableName) + ' t3 ON (t1.' + KeyFieldName + '=t3.' + KeyFieldName + ') ' +
    WhereStr +
    'GROUP BY t1.' + KeyFieldName + ' ' +
    'ORDER BY t3.' + PickFieldName);
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AExistsIDs, QFieldInt(AIDFieldName));
      VAppend(ACounts, QFieldInt('MotorCount'));
      VAppend(ANames, QFieldStr(ANameFieldName));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;  }

function TDataBase.MonthAndDatesForBuildLogLoad(const AYear: Word;
  const AUsedNameIDs: TIntVector;
  out AMonths: TStrVector; out ADates: TDateMatrix): Boolean;
var
  i: Integer;
  BD, ED: TDate;
  Dates: TDateVector;
  WhereStr: String;
begin
  Result:= False;
  AMonths:= nil;
  ADates:= nil;

  WhereStr:= 'WHERE (BuildDate BETWEEN :BD AND :ED) ';
  if not VIsNil(AUsedNameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('','NameID', Length(AUsedNameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT DISTINCT BuildDate ' +
    'FROM MOTORLIST ' +
    WhereStr +
    'ORDER BY BuildDate DESC');

  if not VIsNil(AUsedNameIDs) then
    QParamsInt(AUsedNameIDs);

  for i:= 12 downto 1 do
  begin
    Dates:= nil;
    BD:= FirstDayInMonth(i, AYear);
    ED:= LastDayInMonth(i, AYear);
    QParamDT('BD', BD);
    QParamDT('ED', ED);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        VAppend(Dates, QFieldDT('BuildDate'));
        QNext;
      end;
    end;
    QClose;
    if not VIsNil(Dates) then
    begin
      VAppend(AMonths, FormatDateTime('mmmm yyyy', Dates[0]));
      MAppend(ADates, Dates);
    end;
  end;
  Result:= not VIsNil(AMonths);
end;

function TDataBase.MonthAndDatesForTestLogLoad(const AYear: Word;
  const AUsedNameIDs: TIntVector;
  out AMonths: TStrVector; out ADates: TDateMatrix): Boolean;
var
  i: Integer;
  BD, ED: TDate;
  Dates: TDateVector;
  WhereStr: String;
begin
  Result:= False;
  AMonths:= nil;
  ADates:= nil;

  WhereStr:= 'WHERE (t1.TestDate BETWEEN :BD AND :ED) ';
  if not VIsNil(AUsedNameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(AUsedNameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT DISTINCT t1.TestDate ' +
    'FROM MOTORTEST t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    WhereStr +
    'ORDER BY t1.TestDate DESC');

  if not VIsNil(AUsedNameIDs) then
    QParamsInt(AUsedNameIDs);

  for i:= 12 downto 1 do
  begin
    Dates:= nil;
    BD:= FirstDayInMonth(i, AYear);
    ED:= LastDayInMonth(i, AYear);
    QParamDT('BD', BD);
    QParamDT('ED', ED);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        VAppend(Dates, QFieldDT('TestDate'));
        QNext;
      end;
    end;
    QClose;
    if not VIsNil(Dates) then
    begin
      VAppend(AMonths, FormatDateTime('mmmm yyyy', Dates[0]));
      MAppend(ADates, Dates);
    end;
  end;
  Result:= not VIsNil(AMonths);
end;

function TDataBase.MotorListLoad(const ABuildYear, AShippedType: Integer;
  const ANameIDs: TIntVector; const ANumberLike: String;
  const ANeedOrderByNumber: Boolean; out AMotorIDs: TIntVector; out
  ABuildDates, AMotorNames, AMotorNums, AShippings: TStrVector): Boolean;
var
  OrderStr, WhereStr: String;
  SendDates, BuildDates: TDateVector;
  ReceiverNames: TStrVector;
  Indexes: TIntVector;
begin
  Result:= False;
  AMotorIDs:= nil;
  ABuildDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  AShippings:= nil;
  BuildDates:= nil;
  SendDates:= nil;
  ReceiverNames:= nil;

  if ANeedOrderByNumber then
    OrderStr:= 'ORDER BY t1.OldMotor, t1.MotorNum, t1.BuildDate '
  else
    OrderStr:= 'ORDER BY t1.OldMotor, t1.BuildDate, t1.MotorID ';

  if ANumberLike=EmptyStr then
    WhereStr:= 'WHERE (t1.BuildDate BETWEEN :BD AND :ED) AND (t1.OldMotor=0) '
  else
    WhereStr:= 'WHERE (UPPER(t1.MotorNum) LIKE :NumberLike) ';

  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(ANameIDs));

  if AShippedType=1 then
    WhereStr:= WhereStr + 'AND (t1.CargoID > 0) '
  else if AShippedType=2 then
    WhereStr:= WhereStr + 'AND (t1.CargoID = 0) ';

  QSetQuery(FQuery);
  QClose;
  QSetSQL(
    'SELECT t1.MotorID, t1.MotorNum, t1.BuildDate, ' +
           't2.MotorName, t3.SendDate, t4.ReceiverName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
    'INNER JOIN CARGOLIST t3 ON (t1.CargoID=t3.CargoID) ' +
    'INNER JOIN CARGORECEIVERS t4 ON (t3.ReceiverID=t4.ReceiverID) ' +
    WhereStr +
    OrderStr);
  if ANumberLike=EmptyStr then
  begin
    QParamDT('BD', FirstDayInYear(ABuildYear));
    QParamDT('ED', LastDayInYear(ABuildYear));
  end
  else
    QParamStr('NumberLike', SUpper(ANumberLike)+'%');

  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);

  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(BuildDates, QFieldDT('BuildDate'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(SendDates, QFieldDT('SendDate'));
      VAppend(ReceiverNames, QFieldStr('ReceiverName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;

  if ANeedOrderByNumber then
  begin
    VSortNum(AMotorNums, Indexes);

    AMotorIDs:= VReplace(AMotorIDs, Indexes);
    BuildDates:= VReplace(BuildDates, Indexes);
    AMotorNames:= VReplace(AMotorNames, Indexes);
    AMotorNums:= VReplace(AMotorNums, Indexes);
    SendDates:= VReplace(SendDates, Indexes);
    ReceiverNames:= VReplace(ReceiverNames, Indexes);
  end;

  ABuildDates:= VFormatDateTime('dd.mm.yyyy', BuildDates);
  AShippings:= VFormatDateTime('dd.mm.yyyy – ', SendDates);
  VChangeIf(AShippings, '30.12.1899 – ', EmptyStr);
  VChangeIf(ReceiverNames, '<не указан>', EmptyStr);
  AShippings:= VSum(AShippings, ReceiverNames);
end;

procedure TDataBase.MotorInfoLoad(const AMotorID: Integer; out ABuildDate,
  ASendDate: TDate; out AMotorName, AMotorNum, ASeries, ARotorNum,
  AReceiverName: String; out ATestDates: TDateVector; out
  ATestResults: TIntVector; out ATestNotes: TStrVector);
var
  CargoID: Integer;
begin
  ABuildDate:= 0;
  ASendDate:= 0;
  ARotorNum:= EmptyStr;
  AReceiverName:= EmptyStr;
  AMotorName:= EmptyStr;
  AMotorNum:= EmptyStr;
  ASeries:= EmptyStr;
  ATestDates:= nil;
  ATestResults:= nil;
  ATestNotes:= nil;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.BuildDate, t1.MotorNum, t1.RotorNum, ' +
           't1.CargoID, t1.Series, t2.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
    'WHERE t1.MotorID=:MotorID');
  QParamInt('MotorID', AMotorID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    ABuildDate:= QFieldDT('BuildDate');
    AMotorNum:= QFieldStr('MotorNum');
    ARotorNum:= QFieldStr('RotorNum');
    ASeries:= QFieldStr('Series');
    AMotorName:= QFieldStr('MotorName');
    CargoID:= QFieldInt('CargoID');
  end;
  QClose;

  if CargoID>0 then
  begin
    QSetSQL(
      'SELECT t1.SendDate, t2.ReceiverName ' +
      'FROM CARGOLIST t1  ' +
      'INNER JOIN CARGORECEIVERS t2 ON (t1.ReceiverID=t2.ReceiverID) ' +
      'WHERE CargoID=:CargoID');
    QParamInt('CargoID', CargoID);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      ASendDate:= QFieldDT('SendDate');
      AReceiverName:= QFieldStr('ReceiverName');
    end;
    QClose;
  end;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT TestDate, Fail, TestNote FROM MOTORTEST ' +
    'WHERE MotorID=:MotorID');
  QParamInt('MotorID', AMotorID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ATestDates, QFieldDT('TestDate'));
      VAppend(ATestResults, QFieldInt('Fail'));
      VAppend(ATestNotes, QFieldStr('TestNote'));
      QNext;
    end;
  end;
  QClose;
end;

function TDataBase.BuildListLoad(const ABeginDate, AEndDate: TDate;
                  const AUsedNameIDs: TIntVector;
                  const ANeedOrderByNumber: Boolean;
                  out AMotorIDs, ANameIDs, AOldMotors: TIntVector;
                  out ABuildDates: TDateVector;
                  out AMotorNames, AMotorNums, ARotorNums: TStrVector): Boolean;
var
  S, WhereStr, OrderStr: String;
  Indexes: TIntVector;
begin
  Result:= False;
  AMotorIDs:= nil;
  ANameIDs:= nil;
  ABuildDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ARotorNums:= nil;
  AOldMotors:= nil;

  S:= 'SELECT t1.MotorID, t1.MotorNum, t1.RotorNum, t1.BuildDate, ' +
             't1.NameID, t1.OldMotor, t2.MotorName ' +
      'FROM MOTORLIST t1 ' +
      'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ';

  WhereStr:= 'WHERE (t1.BuildDate BETWEEN :BD AND :ED) ';
  if not VIsNil(AUsedNameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(AUsedNameIDs));

  if ANeedOrderByNumber then
    OrderStr:= 'ORDER BY t1.MotorNum, t1.BuildDate'
  else
    OrderStr:= 'ORDER BY t1.BuildDate, t1.MotorID';

  S:= S + WhereStr + OrderStr;

  QSetQuery(FQuery);
  QClose;
  QSetSQL(S);
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  if not VIsNil(AUsedNameIDs) then
    QParamsInt(AUsedNameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AOldMotors, QFieldInt('OldMotor'));
      VAppend(ANameIDs, QFieldInt('NameID'));
      VAppend(ABuildDates, QFieldDT('BuildDate'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ARotorNums, QFieldStr('RotorNum'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;
  if not ANeedOrderByNumber then Exit;

  VSortNum(AMotorNums, Indexes);

  AMotorIDs:= VReplace(AMotorIDs, Indexes);
  AOldMotors:= VReplace(AOldMotors, Indexes);
  ANameIDs:= VReplace(ANameIDs, Indexes);
  ABuildDates:= VReplace(ABuildDates, Indexes);
  AMotorNames:= VReplace(AMotorNames, Indexes);
  AMotorNums:= VReplace(AMotorNums, Indexes);
  ARotorNums:= VReplace(ARotorNums, Indexes);
end;

function TDataBase.BuildTotalLoad(const ABeginDate, AEndDate: TDate;
  const ANameIDs: TIntVector; out AMotorNames: TStrVector; out
  AMotorCounts: TIntVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;

  WhereStr:= 'WHERE (t1.BuildDate BETWEEN :BD AND :ED) ';
  //if ANameID>0 then
  //  WhereStr:= WhereStr + ' AND (t1.NameID = :NameID) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(ANameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(t1.NameID) As MotorCount, t1.NameID, t2.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
    WhereStr +
    'GROUP BY t1.NameID ' +
    'ORDER BY t2.MotorName');
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  //if ANameID>0 then
  //  QParamInt('NameID', ANameID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorCounts, QFieldInt('MotorCount'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.IsDuplicateMotorNumber(const ADate: TDate;
  const ANameID: Integer; const AMotorNum: String): Boolean;
var
  BD, ED: TDate;
begin
  BD:= FirstDayInMonth(IncMonth(ADate, -1));
  ED:= LastDayInMonth(IncMonth(ADate, 1));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT MotorID FROM MOTORLIST ' +
    'WHERE (NameID=:NameID) AND (UPPER(MotorNum)=:MotorNum) AND ' +
          '(BuildDate BETWEEN :BD AND :ED)');
  QParamInt('NameID', ANameID);
  QParamStr('MotorNum', SUpper(AMotorNum));
  QParamDT('BD', BD);
  QParamDT('ED', ED);
  QOpen;
  Result:= not QIsEmpty;
  QClose;
end;

function TDataBase.MotorsInBuildLogWrite(const ABuildDate: TDate; const ANameIDs,
  AOldMotors: TIntVector; const AMotorNums, ARotorNums: TStrVector): Boolean;
var
  i: Integer;
  Series: String;
begin
  Result:= False;
  Series:= FormatDateTime('mm.yy', ABuildDate);
  QSetQuery(FQuery);
  try
    QSetSQL('INSERT INTO MOTORLIST ' +
           '(NameID, MotorNum, RotorNum, BuildDate, Series, OldMotor) ' +
           'VALUES (:NameID, :MotorNum, :RotorNum, :BuildDate, :Series, :OldMotor)');
    QParamDT('BuildDate', ABuildDate);
    QParamStr('Series', Series);
    for i:= 0 to High(AMotorNums) do
    begin
      QParamInt('NameID', ANameIDs[i]);
      QParamStr('MotorNum', AMotorNums[i]);
      QParamStr('RotorNum', ARotorNums[i]);
      QParamInt('OldMotor', AOldMotors[i]);
      QExec;
    end;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.MotorInBuildLogUpdate(const AMotorID: Integer;
  const ABuildDate: TDate; const ANameID, AOldMotor: Integer; const AMotorNum,
  ARotorNum: String): Boolean;
begin
  Result:= False;
  QSetQuery(FQuery);
  try
    QSetSQL('UPDATE MOTORLIST ' +
            'SET BuildDate=:BuildDate, NameID=:NameID, ' +
                'MotorNum=:MotorNum, RotorNum=:RotorNum, OldMotor=:OldMotor ' +
            'WHERE MotorID=:MotorID');
    QParamInt('MotorID', AMotorID);
    QParamInt('NameID', ANameID);
    QParamDT('BuildDate', ABuildDate);
    QParamStr('MotorNum', AMotorNum);
    QParamStr('RotorNum', ARotorNum);
    QParamInt('OldMotor', AOldMotor);
    QExec;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.TestBeforeListLoad(const ANameIDs: TIntVector;
  const ANeedOrderByNumber: Boolean; out ABuildDates: TDateVector; out
  AMotorNames, AMotorNums, ATotalMotorNames: TStrVector; out
  ATotalMotorCounts: TIntVector; out ATestDates: TDateMatrix; out
  ATestFails: TIntMatrix; out ATestNotes: TStrMatrix): Boolean;
var
  OrderStr, WhereStr: String;
  i: Integer;
  TestFails, MotorIDs: TIntVector;
  TestDates: TDateVector;
  TestNotes: TStrVector;
  Indexes: TIntVector;
begin
  Result:= False;

  ABuildDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ATotalMotorNames:= nil;
  ATotalMotorCounts:= nil;
  MotorIDs:= nil;
  ATestDates:= nil;
  ATestFails:= nil;
  ATestNotes:= nil;

  WhereStr:= 'WHERE (t1.OldMotor=0) AND ((t3.Fail IS NULL) OR ' +
                   '(t1.MotorID NOT IN (SELECT MotorID FROM MOTORTEST WHERE Fail=0))) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(ANameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(NameID) As MotorCount, MotorName ' +
    'FROM ( ' +
      'SELECT DISTINCT t1.MotorNum, t1.NameID, t2.MotorName ' +
      'FROM MOTORLIST t1 ' +
      'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
      'LEFT OUTER JOIN MOTORTEST t3 ON (t1.MotorID=t3.MotorID) ' +
      WhereStr +
      ') ' +
    'GROUP BY NameID ' +
    'ORDER BY MotorName '
  );

  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ATotalMotorCounts, QFieldInt('MotorCount'));
      VAppend(ATotalMotorNames, QFieldStr('MotorName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;

  if ANeedOrderByNumber then
    OrderStr:= 'ORDER BY t1.MotorNum, t1.BuildDate'
  else
    OrderStr:= 'ORDER BY t1.BuildDate, t1.MotorID';

  QSetSQL(
    'SELECT DISTINCT t1.MotorID, t1.BuildDate, t1.MotorNum, t2.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
    'LEFT OUTER JOIN MOTORTEST t3 ON (t1.MotorID=t3.MotorID) ' +
    WhereStr +
    OrderStr);
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(MotorIDs, QFieldInt('MotorID'));
      VAppend(ABuildDates, QFieldDT('BuildDate'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      QNext;
    end;
  end;
  QClose;

  if ANeedOrderByNumber then
  begin
    VSortNum(AMotorNums, Indexes);

    MotorIDs:= VReplace(MotorIDs, Indexes);
    ABuildDates:= VReplace(ABuildDates, Indexes);
    AMotorNames:= VReplace(AMotorNames, Indexes);
    AMotorNums:= VReplace(AMotorNums, Indexes);
  end;

  QSetSQL(
  'SELECT TestDate, Fail, TestNote ' +
  'FROM MOTORTEST ' +
  'WHERE MotorID = :MotorID ' +
  'ORDER BY TestDate');
  for i:= 0 to High(MotorIDs) do
  begin
    TestDates:= nil;
    TestFails:= nil;
    TestNotes:= nil;
    QParamInt('MotorID', MotorIDs[i]);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        VAppend(TestFails, QFieldInt('Fail'));
        VAppend(TestDates, QFieldDT('TestDate'));
        VAppend(TestNotes, QFieldStr('TestNote'));
        QNext;
      end;
    end;
    QClose;
    MAppend(ATestDates, TestDates);
    MAppend(ATestFails, TestFails);
    MAppend(ATestNotes, TestNotes);
  end;
end;

function TDataBase.TestListLoad(const ABeginBuildDate, AEndBuildDate: TDate;
  const ANameIDs: TIntVector; const ANeedOrderByNumber: Boolean; out ATestIDs,
  ATestResults: TIntVector; out ATestDates: TDateVector; out AMotorNames,
  AMotorNums, ATestNotes: TStrVector): Boolean;
var
  OrderStr: String;
  WhereStr: String;
  Indexes: TIntVector;
begin
  Result:= False;
  ATestIDs:= nil;
  ATestResults:= nil;
  ATestDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ATestNotes:= nil;

  WhereStr:= 'WHERE (t1.TestDate BETWEEN :BD AND :ED) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

  if ANeedOrderByNumber then
    OrderStr:= 'ORDER BY t2.MotorNum, t1.TestDate, t1.Fail DESC'
  else
    OrderStr:= 'ORDER BY t1.TestDate, t1.TestID, t1.Fail DESC';
  QSetQuery(FQuery);
  QClose;
  QSetSQL(
    'SELECT t1.TestID, t1.TestDate, t1.Fail, t1.TestNote, ' +
           't2.MotorNum, t2.Series, t3.MotorName ' +
    'FROM MOTORTEST t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    OrderStr);
  QParamDT('BD', ABeginBuildDate);
  QParamDT('ED', AEndBuildDate);
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ATestIDs, QFieldInt('TestID'));
      VAppend(ATestResults, QFieldInt('Fail'));
      VAppend(ATestDates, QFieldDT('TestDate'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ATestNotes, QFieldStr('TestNote'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;
  if not ANeedOrderByNumber then Exit;

  VSortNum(AMotorNums, Indexes);

  ATestIDs:= VReplace(ATestIDs, Indexes);
  ATestResults:= VReplace(ATestResults, Indexes);
  ATestDates:= VReplace(ATestDates, Indexes);
  AMotorNames:= VReplace(AMotorNames, Indexes);
  AMotorNums:= VReplace(AMotorNums, Indexes);
  ATestNotes:= VReplace(ATestNotes, Indexes);
end;

function TDataBase.TestTotalLoad(const ABeginBuildDate, AEndBuildDate: TDate;
  const ANameIDs: TIntVector; out AMotorNames: TStrVector; out AMotorCounts,
  AFailCounts: TIntVector): Boolean;
var
  n: Integer;
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;
  AFailCounts:= nil;

  WhereStr:= 'WHERE (t1.TestDate BETWEEN :BD AND :ED) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(t2.NameID) As MotorCount, t2.NameID, t3.MotorName ' +
    'FROM MOTORTEST t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    'GROUP BY t2.NameID ' +
    'ORDER BY t3.MotorName');
  QParamDT('BD', ABeginBuildDate);
  QParamDT('ED', AEndBuildDate);
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorCounts, QFieldInt('MotorCount'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  VDim(AFailCounts, Length(AMotorNames), 0);

  WhereStr:= 'WHERE (t1.TestDate BETWEEN :BD AND :ED) AND (t1.Fail=1) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

  QSetSQL(
    'SELECT COUNT(t2.NameID) As MotorCount, t2.NameID, t3.MotorName ' +
    'FROM MOTORTEST t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    'GROUP BY t2.NameID');
  QParamDT('BD', ABeginBuildDate);
  QParamDT('ED', AEndBuildDate);
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
       n:= VIndexOf(AMotorNames, QFieldStr('MotorName'));
       if n>=0 then
         AFailCounts[n]:= QFieldInt('MotorCount');
      QNext;
    end;
    Result:= True;
  end;
  QClose;

end;

function TDataBase.TestChooseListLoad(const ANameID: Integer;
  const ANumberLike: String; out AMotorIDs: TIntVector; out AMotorNums,
  ABuildDates, ATests: TStrVector): Boolean;
var
  BuildDates: TDateVector;
  i, Fail: Integer;
  TestDate: TDate;
  S, TestNote: String;
begin
  Result:= False;

  AMotorIDs:= nil;
  AMotorNums:= nil;
  ABuildDates:= nil;
  ATests:= nil;

  if SEmpty(ANumberLike) then Exit;

  BuildDates:= nil;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT MotorID, BuildDate, MotorNum ' +
    'FROM MOTORLIST ' +
    'WHERE (NameID = :NameID) AND (OldMotor=0) AND ' +
          '(UPPER(MotorNum) LIKE :NumberLike) ' +
    'ORDER BY BuildDate DESC');
  QParamInt('NameID', ANameID);
  QParamStr('NumberLike', SUpper(ANumberLike)+'%');
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(BuildDates, QFieldDT('BuildDate'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;

  ABuildDates:= VFormatDateTime('dd.mm.yyyy', BuildDates);
  for i:=0 to High(AMotorIDs) do
  begin
    S:= 'не проводились';
    QSetSQL(
      'SELECT TestDate, Fail, TestNote ' +
      'FROM MOTORTEST ' +
      'WHERE MotorID = :MotorID ' +
      'ORDER BY TestDate DESC ' +
      'LIMIT 1');
    QParamInt('MotorID', AMotorIDs[i]);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      Fail:= QFieldInt('Fail');
      TestDate:= QFieldDT('TestDate');
      TestNote:= QFieldStr('TestNote');
      S:= FormatDateTime('dd.mm.yyyy - ', TestDate);
      if Fail=0 then
        S:= S + 'норма '
      else
        S:= S + 'брак ';
      if TestNote<>EmptyStr then
        S:= S + '(' + TestNote + ')';
    end;
    QClose;
    VAppend(ATests, S);
  end;
end;

function TDataBase.MotorsInTestLogWrite(const ATestDate: TDate; const AMotorIDs,
  ATestResults: TIntVector; const ATestNotes: TStrVector): Boolean;
var
  i: Integer;
begin
  Result:= False;

  QSetQuery(FQuery);
  try
    QSetSQL('INSERT INTO MOTORTEST (MotorID, TestDate, Fail, TestNote) ' +
           'VALUES (:MotorID, :TestDate, :Fail, :TestNote)');
    QParamDT('TestDate', ATestDate);
    for i:= 0 to High(AMotorIDs) do
    begin
      QParamInt('MotorID', AMotorIDs[i]);
      QParamInt('Fail', ATestResults[i]);
      QParamStr('TestNote', ATestNotes[i]);
      QExec;
    end;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.TestLastFineDatesStr(const AMotorIDs: TIntVector): TStrVector;
var
  i: Integer;
begin
  VDim(Result{%H-}, Length(AMotorIDs), EmptyStr);

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT TestDate ' +
    'FROM MOTORTEST ' +
    'WHERE (Fail=0) AND (MotorID = :MotorID) ' +
    'ORDER BY TestDate DESC ' +
    'LIMIT 1'
  );

  for i:= 0 to High(AMotorIDs) do
  begin
    QParamInt('MotorID', AMotorIDs[i]);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      Result[i]:= FormatDateTime('dd.mm.yyyy', QFieldDT('TestDate'));
    end;
    QClose;
  end;
end;

function TDataBase.ControlNoteLoad(const AMotorID: Integer): String;
begin
  Result:= ValueStrByInt32ID('MOTORLIST', 'ControlNote', 'MotorID', AMotorID);
end;

function TDataBase.ControlListLoad(const ANameIDs: TIntVector;
  const ANumberLike: String; out AMotorIDs: TIntVector;
  out AMotorNames, AMotorNums, ASeries, ANotes: TStrVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;


  AMotorIDs:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ASeries:= nil;
  ANotes:= nil;

  WhereStr:= 'WHERE (t1.ControlNote IS NOT NULL) ';
  if ANumberLike<>EmptyStr then
    WhereStr:= WhereStr + ' AND (UPPER(t1.MotorNum) LIKE :NumberLike) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(ANameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.MotorID, t1.ControlNote, t1.MotorNum, t1.Series, t2.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
    WhereStr +
    'ORDER BY t1.MotorNum'
    );

  if ANumberLike<>EmptyStr then
    QParamStr('NumberLike', SUpper(ANumberLike)+'%');

  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);

  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ASeries, QFieldStr('Series'));
      VAppend(ANotes, QFieldStr('ControlNote'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.ControlUpdate(const AMotorID: Integer; const ANote: String): Boolean;
begin
  Result:= False;
  QSetQuery(FQuery);
  try
    QSetSQL(
      'UPDATE MOTORLIST ' +
      'SET ControlNote = :ControlNote ' +
      'WHERE MotorID = :MotorID'
      );
    QParamInt('MotorID', AMotorID);
    QParamStr('ControlNote', ANote);
    QExec;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.ControlDelete(const AMotorID: Integer): Boolean;
begin
  Result:= False;
  QSetQuery(FQuery);
  try
    QSetSQL(
      'UPDATE MOTORLIST ' +
      'SET ControlNote = NULL ' +
      'WHERE MotorID = :MotorID'
      );
    QParamInt('MotorID', AMotorID);
    QExec;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.MotorListToControl(const ANameID: Integer; const ANumberLike: String;
                                out AMotorIDs: TIntVector;
                                out AMotorNums, ASeries: TStrVector): Boolean;
begin
  Result:= False;

  AMotorIDs:= nil;
  AMotorNums:= nil;
  ASeries:= nil;
  if (ANumberLike=EmptyStr) or (ANameID<=0) then Exit;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT MotorID, MotorNum, Series, ControlNote ' +
    'FROM MOTORLIST ' +
    'WHERE (ControlNote IS NULL) AND ' +
          '(NameID = :NameID) AND ' +
          '(UPPER(MotorNum) LIKE :NumberLike) ' +
    'ORDER BY MotorNum'
    );
  QParamStr('NumberLike', SUpper(ANumberLike)+'%');
  QParamInt('NameID', ANameID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ASeries, QFieldStr('Series'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

end;

function TDataBase.MotorListOnControl(const AMotorID: Integer;
  out AMotorIDs: TIntVector; out AMotorNums, ASeries: TStrVector): Boolean;
begin
  Result:= False;

  AMotorIDs:= nil;
  AMotorNums:= nil;
  ASeries:= nil;

  if AMotorID<=0 then Exit;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT MotorID, MotorNum, Series, ControlNote ' +
    'FROM MOTORLIST ' +
    'WHERE MotorID = :MotorID ' +
    //'WHERE ControlNote IS NOT NULL ' +
    'ORDER BY MotorNum'
    );
  QParamInt('MotorID', AMotorID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ASeries, QFieldStr('Series'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.StoreListLoad(const ANameIDs: TIntVector; const ADeltaDays: Integer;
  const ANeedOrderByNumber, ANeedOrderByName: Boolean; out
  ATestDates: TDateVector; out AMotorNames, AMotorNums: TStrVector): Boolean;
var
  OrderStr: String;
  WhereStr: String;
  TestDates: TDateVector;
  MotorNames, MotorNums: TStrVector;
  i, n: Integer;
  Indexes: TIntVector;
begin
  Result:= False;
  ATestDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  TestDates:= nil;
  MotorNames:= nil;
  MotorNums:= nil;

  WhereStr:= 'WHERE (t1.Fail=0) AND (t2.CargoID=0) ';
  if ADeltaDays>0 then
    WhereStr:= WhereStr + 'AND (TestDate<=:BoundaryDate) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

  OrderStr:= 'ORDER BY ';
  if ANeedOrderByNumber or ANeedOrderByName then
  begin
    if ANeedOrderByName then
      OrderStr:= OrderStr + 't3.MotorName, ';
    if ANeedOrderByNumber then
      OrderStr:= OrderStr + 't2.MotorNum, ';
  end;
  OrderStr:= OrderStr + 't1.TestDate';

  QSetQuery(FQuery);
  QClose;
  QSetSQL(
    'SELECT t1.TestDate, t2.MotorNum, t3.MotorName ' +
    'FROM MOTORTEST t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    OrderStr);
  if ADeltaDays>0 then
    QParamDT('BoundaryDate', IncDay(Date, -ADeltaDays));
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(TestDates, QFieldDT('TestDate'));
      VAppend(MotorNames, QFieldStr('MotorName'));
      VAppend(MotorNums, QFieldStr('MotorNum'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;

  //записываем выходные вектора без возможных дублей (если норм тесты были несколько раз)
  for i:= High(TestDates) downto 0 do
  begin
    n:= VIndexOf(AMotorNums, MotorNums[i]);
    if (n>=0) and (MotorNames[i]=AMotorNames[n])  then //есть совпадение
    begin
      if TestDates[i]>ATestDates[n] then //записываем более позднюю дату испытаний
        ATestDates[n]:= TestDates[i];
    end
    else begin //совпадений нет
      VAppend(ATestDates, TestDates[i]);
      VAppend(AMotorNames, MotorNames[i]);
      VAppend(AMotorNums, MotorNums[i]);
    end;
  end;

  if ANeedOrderByNumber then
  begin
    VSortNum(AMotorNums, Indexes);
    ATestDates:= VReplace(ATestDates, Indexes);
    AMotorNames:= VReplace(AMotorNames, Indexes);
    AMotorNums:= VReplace(AMotorNums, Indexes);
  end
  else begin
    ATestDates:= VReverse(ATestDates);
    AMotorNames:= VReverse(AMotorNames);
    AMotorNums:= VReverse(AMotorNums);
  end;
end;

function TDataBase.StoreTotalLoad(const ANameIDs: TIntVector;
  const ADeltaDays: Integer;
  out AMotorNames: TStrVector; out AMotorCounts: TIntVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;

  WhereStr:= 'WHERE (t1.Fail=0) AND (t2.CargoID=0) ';
  if ADeltaDays>0 then
    WhereStr:= WhereStr + 'AND (TestDate<=:BoundaryDate) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(NameID) As MotorCount, MotorName ' +
    'FROM ' +
       '(SELECT DISTINCT t1.MotorID, t2.NameID, t3.MotorName ' +
        'FROM MOTORTEST t1 ' +
        'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
        'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
        WhereStr +
       ')' +
    'GROUP BY NameID ' +
    'ORDER BY MotorName');
  if ADeltaDays>0 then
    QParamDT('BoundaryDate', IncDay(Date, -ADeltaDays));
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorCounts, QFieldInt('MotorCount'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.ShipmentListLoad(const ANameIDs, AReceiverIDs: TIntVector;
  const AYear: Word; out AMonths: TStrVector; out AShipments: TStrMatrix; out
  ACargoIDs: TIntMatrix): Boolean;
var
  i,j: Integer;
  S: String;
  BD, ED: TDate;
  Shipments, ReceiverNames: TStrVector;
  SendDates: TDateVector;
  CargoIDs: TIntVector;
begin
  Result:= False;
  AMonths:= nil;
  AShipments:= nil;
  ACargoIDs:= nil;

  for i:= 12 downto 1 do
  begin
    Shipments:= nil;
    BD:= FirstDayInMonth(i, AYear);
    ED:= LastDayInMonth(i, AYear);
    CargoListLoad(ANameIDs, AReceiverIDs, BD, ED, CargoIDs, SendDates, ReceiverNames);
    if not VIsNil(SendDates) then
    begin
      for j:= 0 to High(SendDates) do
      begin
        S:= FormatDateTime('dd.mm.yyyy', SendDates[j]) + ' - ' + ReceiverNames[j];
        VAppend(Shipments, S);
      end;
      VAppend(AMonths, FormatDateTime('mmmm yyyy', SendDates[0]));
      MAppend(AShipments, Shipments);
      MAppend(ACargoIDs, CargoIDs);
    end;
  end;

  Result:= not VIsNil(AMonths);

end;

function TDataBase.ShipmentMotorListLoad(const ABeginDate, AEndDate: TDate;
  const ANameIDs, AReceiverIDs: TIntVector; const ANeedOrderByNumber: Boolean;
  out ASendDates: TDateVector; out AMotorNames, AMotorNums, AReceiverNames: TStrVector): Boolean;
var
  i: Integer;
  WhereStr, OrderStr: String;
  Indexes: TIntVector;
  Series: TStrVector;
begin
  Result:= False;
  ASendDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  AReceiverNames:= nil;
  Series:= nil;

  WhereStr:= 'WHERE (t1.CargoID>0) AND (t2.SendDate BETWEEN :BD AND :ED) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(ANameIDs), 'NameID');
  if not VIsNil(AReceiverIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','ReceiverID', Length(AReceiverIDs), 'ReceiverID');

  if ANeedOrderByNumber then
    OrderStr:= 'ORDER BY t1.MotorNum'
  else
    OrderStr:= 'ORDER BY t2.SendDate, t2.ReceiverID, t4.MotorName, t1.MotorNum';

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.MotorNum, t1.Series, t2.SendDate, t3.ReceiverName, t4.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN CARGOLIST t2 ON (t1.CargoID=t2.CargoID) ' +
    'INNER JOIN CARGORECEIVERS t3 ON (t2.ReceiverID=t3.ReceiverID) ' +
    'INNER JOIN MOTORNAMES t4 ON (t1.NameID=t4.NameID) ' +
    WhereStr +
    OrderStr);
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs, 'NameID');
  if not VIsNil(AReceiverIDs) then
    QParamsInt(AReceiverIDs, 'ReceiverID');
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ASendDates, QFieldDT('SendDate'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      //VAppend(AMotorNums, QFieldStr('MotorNum') + ' (' + QFieldStr('Series') + ')');
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(Series, QFieldStr('Series'));
      VAppend(AReceiverNames, QFieldStr('ReceiverName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;

  if ANeedOrderByNumber then
  begin
    VSortNum(AMotorNums, Indexes);
    ASendDates:= VReplace(ASendDates, Indexes);
    AMotorNames:= VReplace(AMotorNames, Indexes);
    AMotorNums:= VReplace(AMotorNums, Indexes);
    AReceiverNames:= VReplace(AReceiverNames, Indexes);
    Series:= VReplace(Series, Indexes);
  end;

  for i:= 0 to High(AMotorNums) do
    AMotorNums[i]:= AMotorNums[i] + ' (' + Series[i] + ')';
end;

function TDataBase.ShipmentTotalLoad(const ABeginDate, AEndDate: TDate;
  const ANameIDs, AReceiverIDs: TIntVector; out AMotorNames: TStrVector; out
  AMotorCounts: TIntVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;

  WhereStr:= 'WHERE (t1.CargoID>0) AND (t2.SendDate BETWEEN :BD AND :ED) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(ANameIDs), 'NameID');
  if not VIsNil(AReceiverIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','ReceiverID', Length(AReceiverIDs), 'ReceiverID');

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(t1.NameID) As MotorCount, t3.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN CARGOLIST t2 ON (t1.CargoID=t2.CargoID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t1.NameID=t3.NameID) ' +
    WhereStr +
    'GROUP BY t1.NameID ' +
    'ORDER BY t3.MotorName');
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  QParamsInt(ANameIDs, 'NameID');
  QParamsInt(AReceiverIDs, 'ReceiverID');
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorCounts, QFieldInt('MotorCount'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.ShipmentRecieversTotalLoad(const ABeginDate, AEndDate: TDate;
  const ANameIDs, AReceiverIDs: TIntVector; out AReceiverNames: TStrVector; out
  AMotorNames: TStrMatrix; out AMotorCounts: TIntMatrix): Boolean;
var
  WhereStr: String;
  ReceiverIDs, MotorCounts: TIntVector;
  MotorNames: TStrVector;
  i: Integer;

  procedure GetReceiverNames;
  begin
    WhereStr:= 'WHERE (t1.CargoID>0) AND (t2.SendDate BETWEEN :BD AND :ED) ';
    if not VIsNil(ANameIDs) then
      WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(ANameIDs), 'NameID');
    if not VIsNil(AReceiverIDs) then
      WhereStr:= WhereStr + 'AND' + SqlIN('t2','ReceiverID', Length(AReceiverIDs), 'ReceiverID');

    QSetSQL(
      'SELECT DISTINCT t2.ReceiverID, t3.ReceiverName ' +
      'FROM MOTORLIST t1 ' +
      'INNER JOIN CARGOLIST t2 ON (t1.CargoID=t2.CargoID) ' +
      'INNER JOIN CARGORECEIVERS t3 ON (t2.ReceiverID=t3.ReceiverID) ' +
      WhereStr +
      'ORDER BY t3.ReceiverName');
    QParamDT('BD', ABeginDate);
    QParamDT('ED', AEndDate);

    if not VIsNil(ANameIDs) then
      QParamsInt(ANameIDs, 'NameID');
    if not VIsNil(AReceiverIDs) then
      QParamsInt(AReceiverIDs, 'ReceiverID');
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        VAppend(ReceiverIDs, QFieldInt('ReceiverID'));
        VAppend(AReceiverNames, QFieldStr('ReceiverName'));
        QNext;
      end;
      Result:= True;
    end;
    QClose;
  end;

  procedure GetCargoTotal(const ID: Integer);
  begin
    MotorCounts:= nil;
    MotorNames:= nil;
    WhereStr:= 'WHERE (ReceiverID=:ReceiverID) AND (t3.SendDate BETWEEN :BD AND :ED) ';
    if not VIsNil(ANameIDs) then
      WhereStr:= WhereStr + 'AND' + SqlIN('t1','NameID', Length(ANameIDs));
    QSetSQL(
      'SELECT COUNT(t1.NameID) As MotorCount, t2.MotorName ' +
      'FROM MOTORLIST t1 ' +
      'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
      'INNER JOIN CARGOLIST t3 ON (t1.CargoID=t3.CargoID) ' +
      WhereStr +
      'GROUP BY t1.NameID ' +
      'ORDER BY t2.MotorName');
    QParamInt('ReceiverID', ID);
    QParamDT('BD', ABeginDate);
    QParamDT('ED', AEndDate);
    if not VIsNil(ANameIDs) then
      QParamsInt(ANameIDs);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        VAppend(MotorCounts, QFieldInt('MotorCount'));
        VAppend(MotorNames, QFieldStr('MotorName'));
        QNext;
      end;
    end;
    QClose;
  end;

begin
  Result:= False;
  AReceiverNames:= nil;
  AMotorNames:= nil;
  AMotorCounts:= nil;
  ReceiverIDs:= nil;

  QSetQuery(FQuery);

  //список грузополучателей
  GetReceiverNames;

  if not Result then Exit;

  //кол-во моторов по именам и грузополучателям
  for i:= 0 to High(ReceiverIDs) do
  begin
    GetCargoTotal(ReceiverIDs[i]);
    if not VIsNil(MotorNames) then
    begin
      MAppend(AMotorNames, VCut(MotorNames));
      MAppend(AMotorCounts, VCut(MotorCounts));
    end;
  end;
end;

function TDataBase.CargoListLoad(const ANameIDs, AReceiverIDs: TIntVector;
  const ABeginDate, AEndDate: TDate; out ACargoIDs: TIntVector; out
  ASendDates: TDateVector; out AReceiverNames: TStrVector): Boolean;
var
  i: Integer;
  WhereStr: String;
  CargoIDs: TIntVector;
  SendDates: TDateVector;
  ReceiverNames: TStrVector;

  function IsCargoHasMotorNames(const ACargoID: Integer): Boolean;
  var
    S: String;
  begin
    S:= 'WHERE (CargoID=:CargoID) ';
    S:= S + 'AND' + SqlIN('','NameID', Length(ANameIDs));
    QSetSQL(
      'SELECT MotorID FROM MOTORLIST ' +
      S
      );
    QParamInt('CargoID', ACargoID);
    QParamsInt(ANameIDs);
    QOpen;
    Result:= not QIsEmpty;
    QClose;
  end;

begin
  Result:= False;
  ACargoIDs:= nil;
  ASendDates:= nil;
  AReceiverNames:= nil;

  //WhereStr:= 'WHERE (t1.SendDate >= :BeginDate) AND (t1.SendDate <= :EndDate) ';
  //if not VIsNil(AReceiverIDs) then
  //  WhereStr:= WhereStr + 'AND' + SqlIN('t1','ReceiverID', Length(AReceiverIDs));
  //QSetQuery(FQuery);
  //QClose;
  //QSetSQL(
  //  'SELECT t1.CargoID, t1.SendDate, t2.ReceiverName ' +
  //  'FROM CARGOLIST t1 ' +
  //  'INNER JOIN CARGORECEIVERS t2 ON (t1.ReceiverID=t2.ReceiverID) ' +
  //  WhereStr +
  //  'ORDER BY t1.SendDate DESC');
  //QParamDT('BeginDate', ABeginDate);
  //QParamDT('EndDate', AEndDate);
  //if not VIsNil(AReceiverIDs) then
  //  QParamsInt(AReceiverIDs);
  //QOpen;
  //if not QIsEmpty then
  //begin
  //  QFirst;
  //  while not QEOF do
  //  begin
  //    VAppend(ASendDates, QFieldDT('SendDate'));
  //    VAppend(ACargoIDs, QFieldInt('CargoID'));
  //    VAppend(AReceiverNames, QFieldStr('ReceiverName'));
  //    QNext;
  //  end;
  //  Result:= True;
  //end;
  //QClose;

  CargoIDs:= nil;
  SendDates:= nil;
  ReceiverNames:= nil;
  WhereStr:= 'WHERE (t1.SendDate >= :BeginDate) AND (t1.SendDate <= :EndDate) ';
  if not VIsNil(AReceiverIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t1','ReceiverID', Length(AReceiverIDs));
  QSetQuery(FQuery);
  QClose;
  QSetSQL(
    'SELECT t1.CargoID, t1.SendDate, t2.ReceiverName ' +
    'FROM CARGOLIST t1 ' +
    'INNER JOIN CARGORECEIVERS t2 ON (t1.ReceiverID=t2.ReceiverID) ' +
    WhereStr +
    'ORDER BY t1.SendDate DESC');
  QParamDT('BeginDate', ABeginDate);
  QParamDT('EndDate', AEndDate);
  if not VIsNil(AReceiverIDs) then
    QParamsInt(AReceiverIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(SendDates, QFieldDT('SendDate'));
      VAppend(CargoIDs, QFieldInt('CargoID'));
      VAppend(ReceiverNames, QFieldStr('ReceiverName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if VIsNil(CargoIDs) then Exit;

  if VIsNil(ANameIDs) then
  begin
    Result:= True;
    ACargoIDs:= CargoIDs;
    ASendDates:= SendDates;
    AReceiverNames:= ReceiverNames;
    Exit;
  end;

  for i:= 0 to High(CargoIDs) do
  begin
    if not IsCargoHasMotorNames(CargoIDs[i]) then continue;
    VAppend(ASendDates, SendDates[i]);
    VAppend(ACargoIDs, CargoIDs[i]);
    VAppend(AReceiverNames, ReceiverNames[i]);
  end;

  Result:= not VIsNil(ACargoIDs);
end;

function TDataBase.CargoChooseListLoad(const ANameID: Integer;
  const ANumberLike: String; out AMotorIDs: TIntVector; out AMotorNums,
  ABuildDates, ASeries: TStrVector): Boolean;
var
  D: TDate;
begin
  Result:= False;

  AMotorIDs:= nil;
  AMotorNums:= nil;
  ABuildDates:= nil;
  ASeries:= nil;

  if ANumberLike=EmptyStr then Exit;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT MotorID, BuildDate, MotorNum, Series ' +
    'FROM MOTORLIST ' +
    'WHERE (NameID = :NameID) AND (CargoID=0) AND (OldMotor=0) AND ' +
          '(UPPER(MotorNum) LIKE :NumberLike)');
  QParamInt('NameID', ANameID);
  QParamStr('NumberLike', SUpper(ANumberLike)+'%');
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      D:= QFieldDT('BuildDate');
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ABuildDates, FormatDateTime('dd.mm.yyyy', D));
      VAppend(ASeries, QFieldStr('Series'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.CargoMotorListLoad(const ACargoID: Integer; out
  AReceiverID: Integer; out ASendDate: TDate; out AMotorIDs: TIntVector; out
  AMotorNames, AMotorNums, ASeries: TStrVector): Boolean;
begin
  Result:= False;
  AReceiverID:= 0;
  ASendDate:= 0;
  AMotorIDs:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ASeries:= nil;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT ReceiverID, SendDate FROM CARGOLIST ' +
    'WHERE CargoID = :CargoID');
  QParamInt('CargoID', ACargoID);
  QOpen;
  if QIsEmpty then
  begin
    QClose;
    Exit;
  end;
  QFirst;
  AReceiverID:= QFieldInt('ReceiverID');
  ASendDate:= QFieldDT('SendDate');
  QClose;

  QSetSQL(
    'SELECT t1.MotorID, t1.MotorNum, t1.Series, t2.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
    'WHERE t1.CargoID = :CargoID ' +
    'ORDER BY t2.MotorName, t1.MotorNum, t1.Series');
  QParamInt('CargoID', ACargoID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ASeries, QFieldStr('Series'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.CargoPackingSheet(const ACargoID: Integer; out AMotorNames,
  AMotorNums, ATestDates: TStrVector): Boolean;
var
  MotorIDs: TIntVector;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ATestDates:= nil;
  MotorIDs:= nil;

  if ACargoID=0 then Exit;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.MotorID, t1.MotorNum, t2.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
    'WHERE t1.CargoID = :CargoID ' +
    'ORDER BY t2.MotorName, t1.MotorNum'
  );
  QParamInt('CargoID', ACargoID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(MotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  ATestDates:= TestLastFineDatesStr(MotorIDs);
end;

function TDataBase.CargoLoad(const ACargoID: Integer; out ASendDate: TDate; out
  AReceiverName: String; out AMotorNames: TStrVector; out
  AMotorCounts: TIntVector; out AMotorNums, ASeries: TStrMatrix): Boolean;
var
  i, j: Integer;
  S: String;
  NameIDs: TIntVector;
  NameSeries, SeriesMotorNums: TStrVector;
begin
  Result:= False;
  ASendDate:= 0;
  AReceiverName:= EmptyStr;
  AMotorNames:= nil;
  AMotorCounts:= nil;
  AMotorNums:= nil;
  ASeries:= nil;
  NameIDs:= nil;

  QSetQuery(FQuery);
  QClose;

  //грузополучатель, дата отгрузки
  QSetSQL(
    'SELECT t1.SendDate, t2.ReceiverName ' +
    'FROM CARGOLIST t1 ' +
    'INNER JOIN CARGORECEIVERS t2 ON (t1.ReceiverID=t2.ReceiverID) ' +
    'WHERE t1.CargoID = :CargoID');
  QParamInt('CargoID', ACargoID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    ASendDate:= QFieldDT('SendDate');
    AReceiverName:= QFieldStr('ReceiverName');
    Result:= True;
  end;
  QClose;

  //перечень наименований двигателей, количество по наименованиям
  QSetSQL(
    'SELECT COUNT(t1.NameID) As MotorCount, t1.NameID, t2.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN MOTORNAMES t2 ON (t1.NameID=t2.NameID) ' +
    'WHERE t1.CargoID = :CargoID ' +
    'GROUP BY t1.NameID');
  QParamInt('CargoID', ACargoID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorCounts, QFieldInt('MotorCount'));
      VAppend(NameIDs, QFieldInt('NameID'));
      QNext;
    end;
  end
  else Result:= False;
  QClose;
  if VIsNil(NameIDs) then Exit;

  //номера партий по наименованиям
  QSetSQL(
    'SELECT DISTINCT Series ' +
    'FROM MOTORLIST ' +
    'WHERE (CargoID = :CargoID) AND (NameID = :NameID) ' +
    'ORDER BY Series');
  QParamInt('CargoID', ACargoID);
  for i:= 0 to High(NameIDs) do
  begin
    QParamInt('NameID', NameIDs[i]);
    QOpen;
    NameSeries:= nil;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        VAppend(NameSeries, QFieldStr('Series'));
        QNext;
      end;
    end;
    MAppend(ASeries, NameSeries);
    QClose;
  end;

  //номера двигателей по партиям
  QSetSQL(
    'SELECT MotorNum ' +
    'FROM MOTORLIST ' +
    'WHERE (CargoID = :CargoID) AND (Series = :Series) AND (NameID = :NameID) ' +
    'ORDER BY MotorNum');
  QParamInt('CargoID', ACargoID);
  for i:= 0 to High(NameIDs) do
  begin
    QParamInt('NameID', NameIDs[i]);
    SeriesMotorNums:= nil;
    for j:= 0 to High(ASeries[i]) do
    begin
      QParamStr('Series', ASeries[i][j]);
      QOpen;
      if not QIsEmpty then
      begin
        QFirst;
        S:= QFieldStr('MotorNum');
        QNext;
        while not QEOF do
        begin
          S:= S + ', ' + QFieldStr('MotorNum');
          QNext;
        end;
      end;
      VAppend(SeriesMotorNums, S);
      QClose;
    end;
    MAppend(AMotorNums, SeriesMotorNums);
  end;
end;

function TDataBase.CargoWrite(const ASendDate: TDate; const AReceiverID: Integer;
  const AMotorIDs: TIntVector; const ASeries: TStrVector): Boolean;
var
  i, CargoID: Integer;
begin
  Result:= False;

  QSetQuery(FQuery);
  try
    QSetSQL('INSERT INTO CARGOLIST (ReceiverID, SendDate) ' +
           'VALUES (:ReceiverID, :SendDate)');
    QParamInt('ReceiverID', AReceiverID);
    QParamDT('SendDate', ASendDate);
    QExec;
    CargoID:= LastWritedInt32ID('CARGOLIST');
    if CargoID>0 then
    begin
      QSetSQL(
          'UPDATE MOTORLIST ' +
          'SET CargoID = :CargoID, Series = :Series ' +
          'WHERE MotorID = :MotorID ');
      QParamInt('CargoID', CargoID);
      for i:= 0 to High(AMotorIDs) do
      begin
        QParamInt('MotorID', AMotorIDs[i]);
        QParamStr('Series', ASeries[i]);
        QExec;
      end;
      QCommit;
      Result:= True;
    end;
  except
    QRollback;
  end;
end;

function TDataBase.CargoUpdate(const ACargoID: Integer; const ASendDate: TDate;
  const AReceiverID: Integer; const AMotorIDs, AWritedMotorIDs: TIntVector;
  const ASeries: TStrVector): Boolean;
var
  i: Integer;
begin
  Result:= False;

  QSetQuery(FQuery);
  try
    if not VIsNil(AWritedMotorIDs) then
    begin
      QSetSQL(
            'UPDATE MOTORLIST ' +
            'SET CargoID = 0 ' +
            'WHERE MotorID = :MotorID ');
      for i:= 0 to High(AWritedMotorIDs) do
      begin
        QParamInt('MotorID', AWritedMotorIDs[i]);
        QExec;
      end;
    end;

    if not VIsNil(AMotorIDs) then
    begin
      QSetSQL(
            'UPDATE CARGOLIST ' +
            'SET ReceiverID = :ReceiverID, SendDate = :SendDate ' +
            'WHERE CargoID = :CargoID ');
      QParamInt('CargoID', ACargoID);
      QParamInt('ReceiverID', AReceiverID);
      QParamDT('SendDate', ASendDate);
      QExec;

      QSetSQL(
            'UPDATE MOTORLIST ' +
            'SET CargoID = :CargoID, Series = :Series ' +
            'WHERE MotorID = :MotorID ');
      QParamInt('CargoID', ACargoID);
      for i:= 0 to High(AMotorIDs) do
      begin
        QParamInt('MotorID', AMotorIDs[i]);
        QParamStr('Series', ASeries[i]);
        QExec;
      end;
    end;

    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.ReclamationListLoad(const AMotorID: Integer; out
  ARecDates: TDateVector; out AMileages, AOpinions: TIntVector; out
  APlaceNames, AFactoryNames, ADepartures, ADefectNames, AReasonNames,
  ARecNotes: TStrVector): Boolean;
begin
  Result:= False;

  ARecDates:= nil;
  AMileages:= nil;
  AOpinions:= nil;
  APlaceNames:= nil;
  AFactoryNames:= nil;
  ADepartures:= nil;
  ADefectNames:= nil;
  AReasonNames:= nil;
  ARecNotes:= nil;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.RecDate, t1.Mileage, t1.Opinion, t1.RecNote, t1.Departure,  ' +
           't2.PlaceName, t3.FactoryName, t4.DefectName, t5.ReasonName ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN RECLAMATIONPLACES t2 ON (t1.PlaceID=t2.PlaceID) ' +
    'INNER JOIN RECLAMATIONFACTORIES t3 ON (t1.FactoryID=t3.FactoryID) ' +
    'INNER JOIN RECLAMATIONDEFECTS t4 ON (t1.DefectID=t4.DefectID) ' +
    'INNER JOIN RECLAMATIONREASONS t5 ON (t1.ReasonID=t5.ReasonID) ' +
    'WHERE t1.MotorID=:MotorID ' +
    'ORDER BY t1.RecDate');
  QParamInt('MotorID', AMotorID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ARecDates, QFieldDT('RecDate'));
      VAppend(AMileages, QFieldInt('Mileage'));
      VAppend(AOpinions, QFieldInt('Opinion'));
      VAppend(APlaceNames, QFieldStr('PlaceName'));
      VAppend(AFactoryNames, QFieldStr('FactoryName'));
      VAppend(ADepartures, QFieldStr('Departure'));
      VAppend(ADefectNames, QFieldStr('DefectName'));
      VAppend(AReasonNames, QFieldStr('ReasonName'));
      VAppend(ARecNotes, QFieldStr('RecNote'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.ReclamationListLoad(const ABeginDate, AEndDate: TDate;
  const ANameIDs, ADefectIDs, AReasonIDs: TIntVector;
  const AOrderIndex: Integer; const ANumberLike: String;
  out ARecDates, ABuildDates, AArrivalDates, ASendingDates: TDateVector;
  out ARecIDs, AMotorIDs, AMileages, AOpinions, AColors, APassports: TIntVector;
  out APlaceNames, AFactoryNames, ADepartures, ADefectNames, AReasonNames,
  ARecNotes, AMotorNames, AMotorNums: TStrVector): Boolean;
var
  WhereStr, OrderStr: String;
  Indexes: TIntVector;
begin
  Result:= False;

  ARecDates:= nil;
  ARecIDs:= nil;
  AMotorIDs:= nil;
  AMileages:= nil;
  AOpinions:= nil;
  APlaceNames:= nil;
  AFactoryNames:= nil;
  ADepartures:= nil;
  ADefectNames:= nil;
  AReasonNames:= nil;
  ARecNotes:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ABuildDates:= nil;
  AColors:= nil;
  AArrivalDates:= nil;
  ASendingDates:= nil;
  APassports:= nil;

  if VIsNil(ADefectIDs) or VIsNil(AReasonIDs) then Exit;

  if ANumberLike<>EmptyStr then
    WhereStr:= 'WHERE (UPPER(t6.MotorNum) LIKE :NumberLike) '
   else
    WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t6','NameID', Length(ANameIDs), 'NameIDValue');

  WhereStr:= WhereStr + 'AND' + SqlIN('t1','DefectID', Length(ADefectIDs), 'DefectIDValue');
  WhereStr:= WhereStr + 'AND' + SqlIN('t1','ReasonID', Length(AReasonIDs), 'ReasonIDValue');

  OrderStr:= EmptyStr;
  case AOrderIndex of
  0: OrderStr:= 'ORDER BY t1.RecDate, t6.MotorNum';
  1: OrderStr:= 'ORDER BY t6.MotorNum, t1.RecDate';
  2: OrderStr:= 'ORDER BY t6.BuildDate, t1.RecDate';
  3: OrderStr:= 'ORDER BY t2.PlaceName, t1.RecDate';
  4: OrderStr:= 'ORDER BY t4.DefectName, t1.RecDate';
  5: OrderStr:= 'ORDER BY t5.ReasonName, t1.RecDate';
  end;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.RecID, t1.MotorID, t1.RecDate, t1.Mileage, ' +
           't1.Opinion, t1.RecNote, t1.Departure,  ' +
           't1.ArrivalDate, t1.SendingDate, t1.Passport, ' +
           't2.PlaceName, t3.FactoryName, t4.DefectName, t5.ReasonName, t5.ReasonColor, ' +
           't6.MotorNum, t6.BuildDate, t7.MotorName ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN RECLAMATIONPLACES t2 ON (t1.PlaceID=t2.PlaceID) ' +
    'INNER JOIN RECLAMATIONFACTORIES t3 ON (t1.FactoryID=t3.FactoryID) ' +
    'INNER JOIN RECLAMATIONDEFECTS t4 ON (t1.DefectID=t4.DefectID) ' +
    'INNER JOIN RECLAMATIONREASONS t5 ON (t1.ReasonID=t5.ReasonID) ' +
    'INNER JOIN MOTORLIST t6 ON (t1.MotorID=t6.MotorID) ' +
    'INNER JOIN MOTORNAMES t7 ON (t6.NameID=t7.NameID) ' +
    WhereStr +
    OrderStr);

  if ANumberLike<>EmptyStr then
    QParamStr('NumberLike', SUpper(ANumberLike)+'%')
  else begin
    QParamDT('BD', ABeginDate);
    QParamDT('ED', AEndDate);
  end;

  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs, 'NameIDValue');
  QParamsInt(ADefectIDs, 'DefectIDValue');
  QParamsInt(AReasonIDs, 'ReasonIDValue');

  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ARecDates, QFieldDT('RecDate'));
      VAppend(ARecIDs, QFieldInt('RecID'));
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMileages, QFieldInt('Mileage'));
      VAppend(AOpinions, QFieldInt('Opinion'));
      VAppend(APlaceNames, QFieldStr('PlaceName'));
      VAppend(AFactoryNames, QFieldStr('FactoryName'));
      VAppend(ADepartures, QFieldStr('Departure'));
      VAppend(ADefectNames, QFieldStr('DefectName'));
      VAppend(AReasonNames, QFieldStr('ReasonName'));
      VAppend(ARecNotes, QFieldStr('RecNote'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ABuildDates, QFieldDT('BuildDate'));
      VAppend(AColors, QFieldInt('ReasonColor'));
      VAppend(AArrivalDates, QFieldDT('ArrivalDate'));
      VAppend(ASendingDates, QFieldDT('SendingDate'));
      VAppend(APassports, QFieldInt('Passport'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;
  if AOrderIndex<>1{MotorNum} then Exit;

  VSortNum(AMotorNums, Indexes);

  ARecDates:= VReplace(ARecDates, Indexes);
  ARecIDs:= VReplace(ARecIDs, Indexes);
  AMotorIDs:= VReplace(AMotorIDs, Indexes);
  AMileages:= VReplace(AMileages, Indexes);
  AOpinions:= VReplace(AOpinions, Indexes);
  APlaceNames:= VReplace(APlaceNames, Indexes);
  AFactoryNames:= VReplace(AFactoryNames, Indexes);
  ADepartures:= VReplace(ADepartures, Indexes);
  ADefectNames:= VReplace(ADefectNames, Indexes);
  AReasonNames:= VReplace(AReasonNames, Indexes);
  ARecNotes:= VReplace(ARecNotes, Indexes);
  AMotorNames:= VReplace(AMotorNames, Indexes);
  AMotorNums:= VReplace(AMotorNums, Indexes);
  ABuildDates:= VReplace(ABuildDates, Indexes);
  AColors:= VReplace(AColors, Indexes);
  AArrivalDates:= VReplace(AArrivalDates, Indexes);
  ASendingDates:= VReplace(ASendingDates, Indexes);
  APassports:= VReplace(APassports, Indexes);
end;

function TDataBase.ReclamationChooseListLoad(const ANameID: Integer;
  const ANumberLike: String; out AMotorIDs: TIntVector; out AMotorNums,
  ABuildDates, AShippings: TStrVector): Boolean;
var
  SendDates, BuildDates: TDateVector;
  ReceiverNames: TStrVector;
begin
  Result:= False;

  AMotorIDs:= nil;
  AMotorNums:= nil;
  ABuildDates:= nil;
  AShippings:= nil;

  SendDates:= nil;
  BuildDates:= nil;
  ReceiverNames:= nil;

  if SEmpty(ANumberLike) then Exit;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.MotorID, t1.BuildDate, t1.MotorNum, ' +
           't2.SendDate, t3.ReceiverName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN CARGOLIST t2 ON (t1.CargoID=t2.CargoID) ' +
    'INNER JOIN CARGORECEIVERS t3 ON (t2.ReceiverID=t3.ReceiverID) ' +
    'WHERE (t1.NameID = :NameID) AND ' +
          '(UPPER(t1.MotorNum) LIKE :NumberLike)');
  QParamInt('NameID', ANameID);
  QParamStr('NumberLike', SUpper(ANumberLike)+'%');
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(BuildDates, QFieldDT('BuildDate'));
      VAppend(SendDates, QFieldDT('SendDate'));
      VAppend(ReceiverNames, QFieldStr('ReceiverName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if Result then
  begin
    ABuildDates:= VFormatDateTime('dd.mm.yyyy', BuildDates);
    AShippings:= VFormatDateTime('dd.mm.yyyy – ', SendDates);
    VChangeIf(AShippings, '30.12.1899 – ', EmptyStr);
    VChangeIf(ReceiverNames, '<не указан>', EmptyStr);
    AShippings:= VSum(AShippings, ReceiverNames);
  end;
end;

function TDataBase.ReclamationWrite(const ARecDate: TDate; const AMotorID,
  AMileage, APlaceID, AFactoryID, ADefectID, AReasonID, AOpinion: Integer;
  const ADeparture, ARecNote: String): Boolean;
begin
  Result:= False;
  QSetQuery(FQuery);
  try
    QSetSQL('INSERT INTO RECLAMATIONS ' +
            '(RecDate, MotorID, Mileage, PlaceID, FactoryID, ' +
            ' Departure, DefectID, ReasonID, Opinion, RecNote) ' +
            'VALUES (:RecDate, :MotorID, :Mileage, :PlaceID, :FactoryID, '+
                    ':Departure, :DefectID, :ReasonID, :Opinion, :RecNote)');
    QParamDT('RecDate', ARecDate);
    QParamInt('MotorID', AMotorID);
    QParamInt('Mileage', AMileage);
    QParamInt('PlaceID', APlaceID);
    QParamInt('FactoryID', AFactoryID);
    QParamStr('Departure', ADeparture);
    QParamInt('DefectID', ADefectID);
    QParamInt('ReasonID', AReasonID);
    QParamInt('Opinion', AOpinion);
    QParamStr('RecNote', ARecNote);
    QExec;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.ReclamationUpdate(const ARecDate, AArrivalDate,
  ASendingDate: TDate; const ARecID, AMotorID, AMileage, APlaceID, AFactoryID,
  ADefectID, AReasonID, AOpinion, APassport: Integer; const ADeparture,
  ARecNote: String): Boolean;
var
  S: String;
begin
  Result:= False;
  QSetQuery(FQuery);
  try
    S:= 'UPDATE RECLAMATIONS ' +
        'SET RecDate=:RecDate, MotorID=:MotorID, Mileage=:Mileage, ' +
            'PlaceID=:PlaceID, FactoryID=:FactoryID, ' +
            'Departure=:Departure, DefectID=:DefectID, ' +
            'ReasonID=:ReasonID, Opinion=:Opinion, RecNote=:RecNote, ' +
            'Passport=:Passport, ';
    if AArrivalDate>0 then
      S:= S + 'ArrivalDate=:ArrivalDate, '
    else
      S:= S + 'ArrivalDate=0, ';
    if ASendingDate>0 then
      S:= S + 'SendingDate=:SendingDate '
    else
      S:= S + 'SendingDate=0 ';
    S:= S + 'WHERE RecID = :RecID';

    QSetSQL(S);
    QParamInt('RecID', ARecID);
    QParamDT('RecDate', ARecDate);
    QParamInt('MotorID', AMotorID);
    QParamInt('Mileage', AMileage);
    QParamInt('PlaceID', APlaceID);
    QParamInt('FactoryID', AFactoryID);
    QParamStr('Departure', ADeparture);
    QParamInt('DefectID', ADefectID);
    QParamInt('ReasonID', AReasonID);
    QParamInt('Opinion', AOpinion);
    QParamStr('RecNote', ARecNote);
    if AArrivalDate>0 then
      QParamDT('ArrivalDate', AArrivalDate);
    if ASendingDate>0 then
      QParamDT('SendingDate', ASendingDate);
    QParamInt('Passport', APassport);
    QExec;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.ReclamationLoad(const ARecID: Integer; out ARecDate,
  AArrivalDate, ASendingDate: TDate; out ANameID, AMotorID, AMileage, APlaceID,
  AFactoryID, ADefectID, AReasonID, AOpinion, APassport: Integer; out
  AMotorNum, ADeparture, ARecNote: String): Boolean;
begin
  Result:= False;
  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.RecDate, t1.MotorID, t1.Mileage, t1.PlaceID, t1.FactoryID, ' +
           't1.Departure, t1.DefectID, t1.ReasonID, t1.Opinion, t1.RecNote, ' +
           't1.ArrivalDate, t1.SendingDate, t1.Passport, ' +
           't2.NameID, t2.MotorNum ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'WHERE t1.RecID = :RecID');
  QParamInt('RecID', ARecID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    ARecDate:= QFieldDT('RecDate');
    ANameID:= QFieldInt('NameID');
    AMotorID:= QFieldInt('MotorID');
    AMileage:= QFieldInt('Mileage');
    APlaceID:= QFieldInt('PlaceID');
    AFactoryID:= QFieldInt('FactoryID');
    ADeparture:= QFieldStr('Departure');
    ADefectID:= QFieldInt('DefectID');
    AReasonID:= QFieldInt('ReasonID');
    AOpinion:= QFieldInt('Opinion');
    ARecNote:= QFieldStr('RecNote');
    AMotorNum:= QFieldStr('MotorNum');
    AArrivalDate:= QFieldDT('ArrivalDate');
    ASendingDate:= QFieldDT('SendingDate');
    APassport:= QFieldInt('Passport');
    Result:= True;
  end;
  QClose;
end;

function TDataBase.RepairListLoad(const ANameIDs: TIntVector; const ANumberLike: String;
          const AOrderType: Integer; //1 - по дате прибытия, 2- по дате убытия, 3 - по номеру
          const AListType: Integer; //0 - все, 1-в ремонте, 2 - отремонтированные
          out AArrivalDates, ASendingDates: TDateVector;
          out ARecIDs, AMotorIDs, APassports, AWorkDayCounts: TIntVector;
          out AMotorNames, AMotorNums, ARepairNotes: TStrVector): Boolean;
var
  i: Integer;
  WhereStr, OrderStr: String;
  Indexes: TIntVector;
begin
  Result:= False;

  ARecIDs:= nil;
  AMotorIDs:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ARepairNotes:= nil;
  AArrivalDates:= nil;
  ASendingDates:= nil;
  APassports:= nil;
  AWorkDayCounts:= nil;

  if ANumberLike<>EmptyStr then
    WhereStr:= 'WHERE ((t1.ArrivalDate>0) AND (UPPER(t2.MotorNum) LIKE :NumberLike)) '
  else if AListType=1 then //в ремонте
    WhereStr:= 'WHERE ((t1.ArrivalDate>0) AND (t1.SendingDate=0)) '
  else if AListType=2 then //отремонтированные
    WhereStr:= 'WHERE ((t1.ArrivalDate>0) AND (t1.SendingDate>0)) '
  else //все
    WhereStr:= 'WHERE (t1.ArrivalDate>0) ';

  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

  if AOrderType=1 then
    OrderStr:= 'ORDER BY t1.ArrivalDate, t2.MotorNum'
  else if AOrderType=2 then
    OrderStr:= 'ORDER BY t1.SendingDate, t2.MotorNum'
  else if AOrderType=3 then
    OrderStr:= 'ORDER BY t2.MotorNum, t1.ArrivalDate';

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.RecID, t1.MotorID, t1.RecDate, t1.RepairNote, ' +
           't1.ArrivalDate, t1.SendingDate, t1.Passport, ' +
           't2.MotorNum, t3.MotorName ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    OrderStr);

  if ANumberLike<>EmptyStr then
    QParamStr('NumberLike', SUpper(ANumberLike)+'%');

  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);

  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ARecIDs, QFieldInt('RecID'));
      VAppend(AMotorIDs, QFieldInt('MotorID'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(ARepairNotes, QFieldStr('RepairNote'));
      VAppend(AArrivalDates, QFieldDT('ArrivalDate'));
      VAppend(ASendingDates, QFieldDT('SendingDate'));
      VAppend(APassports, QFieldInt('Passport'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;

  if AOrderType=3 then //MotorNum
  begin
    VSortNum(AMotorNums, Indexes);
    ARecIDs:= VReplace(ARecIDs, Indexes);
    AMotorIDs:= VReplace(AMotorIDs, Indexes);
    AMotorNames:= VReplace(AMotorNames, Indexes);
    AMotorNums:= VReplace(AMotorNums, Indexes);
    ARepairNotes:= VReplace(ARepairNotes, Indexes);
    AArrivalDates:= VReplace(AArrivalDates, Indexes);
    ASendingDates:= VReplace(ASendingDates, Indexes);
    APassports:= VReplace(APassports, Indexes);
  end;

  VDim(AWorkDayCounts{%H-}, Length(AArrivalDates));
    for i:= 0 to High(AWorkDayCounts) do
      AWorkDayCounts[i]:= LoadWorkDaysCountInPeriod(AArrivalDates[i], ASendingDates[i]);
end;

function TDataBase.RepairListForMotorIDLoad(const AMotorID: Integer;
          out AArrivalDates, ASendingDates: TDateVector;
          out APassports, AWorkDayCounts: TIntVector;
          out ARepairNotes: TStrVector): Boolean;
var
  i: Integer;
begin
  Result:= False;
  AArrivalDates:= nil;
  ASendingDates:= nil;
  APassports:= nil;
  AWorkDayCounts:= nil;
  ARepairNotes:= nil;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT RepairNote, ArrivalDate, SendingDate, Passport  ' +
    'FROM RECLAMATIONS ' +
    'WHERE (ArrivalDate > 0) AND (MotorID = :MotorID) ' +
    'ORDER BY ArrivalDate'
    );
  QParamInt('MotorID', AMotorID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AArrivalDates, QFieldDT('ArrivalDate'));
      VAppend(ASendingDates, QFieldDT('SendingDate'));
      VAppend(APassports, QFieldInt('Passport'));
      VAppend(ARepairNotes, QFieldStr('RepairNote'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;

  if not Result then Exit;
  VDim(AWorkDayCounts{%H-}, Length(AArrivalDates));
    for i:= 0 to High(AWorkDayCounts) do
      AWorkDayCounts[i]:= LoadWorkDaysCountInPeriod(AArrivalDates[i], ASendingDates[i]);
end;

function TDataBase.RepairUpdate(const ARecID: Integer; const AArrivalDate,
  ASendingDate: TDate; const APassport: Integer; const ARepairNote: String): Boolean;
begin
  Result:= False;
  QSetQuery(FQuery);
  try
    QSetSQL(
      'UPDATE RECLAMATIONS ' +
      'SET ArrivalDate = :ArrivalDate, SendingDate = :SendingDate, ' +
          'Passport = :Passport, RepairNote = :RepairNote ' +
      'WHERE RecID = :RecID'
      );
    QParamInt('RecID', ARecID);
    if AArrivalDate=0 then
      QParamInt('ArrivalDate', 0)
    else
      QParamDT('ArrivalDate', AArrivalDate);
    if ASendingDate=0 then
      QParamInt('SendingDate', 0)
    else
      QParamDT('SendingDate', ASendingDate);
    QParamInt('Passport', APassport);
    QParamStr('RepairNote', ARepairNote);
    QExec;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TDataBase.RepairLoad(const ARecID: Integer; out AArrivalDate,
  ASendingDate: TDate; out APassport: Integer; out ARepairNote: String): Boolean;
begin
  Result:= False;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT RepairNote, ArrivalDate, SendingDate, Passport  ' +
    'FROM RECLAMATIONS ' +
    'WHERE RecID = :RecID'
    );
  QParamInt('RecID', ARecID);

  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    AArrivalDate:= QFieldDT('ArrivalDate');
    ASendingDate:= QFieldDT('SendingDate');
    APassport:= QFieldInt('Passport');
    ARepairNote:= QFieldStr('RepairNote');
    Result:= True;
  end;
  QClose;
end;

procedure TDataBase.RepairDelete(const ARecID: Integer);
begin
  QSetQuery(FQuery);
  try
    QSetSQL(
      'UPDATE RECLAMATIONS ' +
      'SET ArrivalDate = 0, SendingDate = 0, ' +
          'Passport = 0, RepairNote = NULL ' +
      'WHERE RecID = :RecID'
      );
    QParamInt('RecID', ARecID);
    QExec;
    QCommit;
  except
    QRollback;
  end;
end;

function TDataBase.MotorListToRepair(const ANameID: Integer;
          const AMotorNumberLike: String;
          out ARecIDs, AMotorNameIDs: TIntVector;
          out ABuildDates, ARecDates: TDateVector;
          out AMotorNums, APlaceNames: TStrVector): Boolean;
begin
  Result:= False;
  ARecIDs:= nil;
  AMotorNameIDs:= nil;
  ABuildDates:= nil;
  ARecDates:= nil;
  AMotorNums:= nil;
  APlaceNames:= nil;
  if (AMotorNumberLike=EmptyStr) or (ANameID<=0) then Exit;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.RecID, t1.RecDate, t1.RepairNote, ' +
           't2.MotorNum, t2.BuildDate, t2.NameID, t3.PlaceName ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN RECLAMATIONPLACES t3 ON (t1.PlaceID=t3.PlaceID) ' +
    'WHERE (t1.ArrivalDate = 0) AND ' +
          '(t2.NameID = :NameID) AND ' +
          '(UPPER(t2.MotorNum) LIKE :NumberLike) ' +
    'ORDER BY t2.MotorNum'
    );
  QParamStr('NumberLike', SUpper(AMotorNumberLike)+'%');
  QParamInt('NameID', ANameID);

  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ARecIDs, QFieldInt('RecID'));
      VAppend(AMotorNameIDs, QFieldInt('NameID'));
      VAppend(ABuildDates, QFieldDT('BuildDate'));
      VAppend(ARecDates, QFieldDT('RecDate'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(APlaceNames, QFieldStr('PlaceName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TDataBase.MotorListOnRepair(const ARecID: Integer; out ARecIDs,
  AMotorNameIDs: TIntVector; out ABuildDates, ARecDates: TDateVector; out
  AMotorNums, APlaceNames: TStrVector): Boolean;
begin
  Result:= False;
  ARecIDs:= nil;
  AMotorNameIDs:= nil;
  ABuildDates:= nil;
  ARecDates:= nil;
  AMotorNums:= nil;
  APlaceNames:= nil;
  if ARecID<=0 then Exit;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.RecID, t1.RecDate, t1.RepairNote, ' +
           't2.MotorNum, t2.BuildDate, t2.NameID, t3.PlaceName ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN RECLAMATIONPLACES t3 ON (t1.PlaceID=t3.PlaceID) ' +
    'WHERE t1.RecID = :RecID ' +
    'ORDER BY t2.MotorNum'
    );
  QParamInt('RecID', ARecID);

  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ARecIDs, QFieldInt('RecID'));
      VAppend(AMotorNameIDs, QFieldInt('NameID'));
      VAppend(ABuildDates, QFieldDT('BuildDate'));
      VAppend(ARecDates, QFieldDT('RecDate'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      VAppend(APlaceNames, QFieldStr('PlaceName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

{function TDataBase.ReclamationMotorsWithReasonsLoad(const ABeginDate,
  AEndDate: TDate; const AAdditionYearsCount: Integer; const AUsedNameIDs,
  AReasonIDs: TIntVector; const AANEMAsSameName: Boolean;
  out AMotorNames: TStrVector; out AClaimCounts: TIntMatrix3D): Boolean;
var
  ExistsNameIDs: TIntVector;
  MotorCounts: TIntMatrix;
  BD, ED: TDate;
  i: Integer;

  procedure GetMotorNamesSinglePeriod(const ABD, AED: TDate;
                          out AExistNames: TStrVector;
                          out AExistNameIDs: TIntVector);
  begin
    AExistNames:= nil;
    AExistNameIDs:= nil;
    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t2.NameID) As MotorCount, t2.NameID, t3.MotorName ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
      'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
             SqlIN('t2','NameID', Length(AUsedNameIDs)) +
      'GROUP BY t2.NameID ' +
      'ORDER BY t3.MotorName');
    QParamDT('BD', ABD);
    QParamDT('ED', AED);
    QParamsInt(AUsedNameIDs);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        VAppend(AExistNames, QFieldStr('MotorName'));
        VAppend(AExistNameIDs, QFieldInt('NameID'));
        QNext;
      end;
    end;
    QClose;
  end;

  procedure GetMotorNamesAllPeriods(const ABD, AED: TDate;
                          const AAdditionYearsCount: Integer;
                          out AExistNames: TStrVector;
                          out AExistNameIDs: TIntVector);
  var
    ExistNames: TStrVector;
    ExistNameIDs: TIntVector;
    n: Integer;
  begin
    AExistNames:= nil;
    AExistNameIDs:= nil;
    for n:= AAdditionYearsCount downto 0 do
    begin
      GetMotorNamesSinglePeriod(IncYear(ABD, -n), IncYear(AED, -n),
                                ExistNames, ExistNameIDs);
      AExistNames:= VUnion(AExistNames, ExistNames);
      AExistNameIDs:= VUnion(AExistNameIDs, ExistNameIDs);
    end;
  end;

  procedure GetDataForReason(const ABD, AED: TDate; const AReasonID: Integer; out ACounts: TIntVector);
  var
    n: Integer;
    WhereStr: String;
  begin
    ACounts:= nil;
    VDim(ACounts, Length(ExistsNameIDs), 0);

    WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
               SqlIN('t2','NameID', Length(ExistsNameIDs));

    if AReasonID>=0 then
      WhereStr:= WhereStr + ' AND (t1.ReasonID = :ReasonID) ';

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t2.NameID) As MotorCount, t2.NameID, t3.MotorName ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
      WhereStr +
      'GROUP BY t2.NameID ' +
      'ORDER BY t3.MotorName');
    QParamDT('BD', ABD);
    QParamDT('ED', AED);
    QParamsInt(ExistsNameIDs);
    if AReasonID>=0 then
      QParamInt('ReasonID', AReasonID);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        n:= VIndexOf(ExistsNameIDs, QFieldInt('NameID'));
        if n>=0 then
          ACounts[n]:= QFieldInt('MotorCount');
        QNext;
      end;
    end;
    QClose;
  end;

  procedure GetDataForSinglePeriod(const ABD, AED: TDate; out ACounts: TIntMatrix);
  var
    Counts: TIntVector;
    n: Integer;
  begin
    ACounts:= nil;
    for n:=0 to High(AReasonIDs) do
    begin
      GetDataForReason(ABD, AED, AReasonIDs[n], Counts);
      MAppend(ACounts, Counts);
    end;
  end;

  procedure RecalcANEM;
  var
    n1, n2, i, j: Integer;
  begin
    n1:= VIndexOf(ExistsNameIDs, 1); //IM1001
    n2:= VIndexOf(ExistsNameIDs, 2); //IM1002
    if (n1<0) and (n2<0) then Exit;
    if n1>=0 then
    begin
      //есть IM1001 - меняем имя в списке
      AMotorNames[n1]:= 'АНЭМ225L4УХЛ2 IM1001/IM1002';
      if n2>=0 then  //есть еще IM1002 - пересчет
      begin
        for i:= 0 to High(AClaimCounts) do
          for j:= 0 to High(AClaimCounts[i]) do
        begin
          AClaimCounts[i, j, n1]:= AClaimCounts[i, j, n1] + AClaimCounts[i, j, n2];
          VDel(AClaimCounts[i, j], n2);
        end;
        VDel(AMotorNames, n2);
      end;
    end
    else //есть только IM1002 - только меняем имя в списке
      AMotorNames[n2]:= 'АНЭМ225L4УХЛ2 IM1001/IM1002';
  end;

begin
  Result:= False;
  AMotorNames:= nil;
  AClaimCounts:= nil;

  if VIsNil(AReasonIDs) then Exit;

  //список наименований двигателей с существующими рекламациями, их общее количество за все годы
  GetMotorNamesAllPeriods(ABeginDate, AEndDate, AAdditionYearsCount,
                          AMotorNames, ExistsNameIDs);
  //нет рекламаций в запрашиваемые периоды
  if VIsNil(ExistsNameIDs) then Exit;

  //матрица данных для каждого периода
  for i:= AAdditionYearsCount downto 0 do
  begin
    BD:= IncYear(ABeginDate, -i);
    ED:= IncYear(AEndDate, -i);
    GetDataForSinglePeriod(BD, ED, MotorCounts);
    //итоговая матрица для всех периодов
    MAppend(AClaimCounts, MotorCounts);
  end;

  if AANEMAsSameName then
    RecalcANEM;

  Result:= True;
end; }

function TDataBase.ReclamationByMotorNamesLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs: TIntVector;
                const AMotorTypeAsSingleName: Boolean;
                const ADoNotUseZeroClaimMotor: Boolean;
                var AMotorNames: TStrVector;
                out AMotorNeeds: TBoolVector;
                out AClaimCounts: TIntMatrix3D): Boolean;
var
  i, j, k: Integer;
  SumCounts: TIntMatrix;
  TypeIDs: TIntVector;

  procedure GetMotorTypes;
  var
    TypeID: Integer;
    TypeNames: TStrVector;
  begin
    TypeIDs:= nil;
    TypeNames:= nil;
    QSetQuery(FQuery);
    QSetSQL(
      'SELECT t1.TypeID, t2.MotorType ' +
      'FROM MOTORNAMES t1 ' +
      'INNER JOIN MOTORTYPES t2 ON (t1.TypeID=t2.TypeID) ' +
      'WHERE ' + SqlIN('t1','NameID', Length(AUsedNameIDs)) +
      'ORDER BY t1.NameID'
    );
    QParamsInt(AUsedNameIDs);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        TypeID:= QFieldInt('TypeID');
        if VIndexOf(TypeIDs, TypeID)<0 then
        begin
          VAppend(TypeIDs, TypeID);
          VAppend(TypeNames, QFieldStr('MotorType'));
        end;
        QNext;
      end;
    end;
    QClose;
    AMotorNames:= TypeNames;
  end;

  procedure GetData(const APeriodIndex: Integer);
  var
    MotorIndex, ReasonIndex: Integer;
    BD, ED: TDate;
  begin
    BD:= IncYear(ABeginDate, -APeriodIndex);
    ED:= IncYear(AEndDate, -APeriodIndex);

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t2.NameID) As ClaimCount, t2.NameID, t1.ReasonID, t3.TypeID, t4.MotorType ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
      'INNER JOIN MOTORTYPES t4 ON (t3.TypeID=t4.TypeID) ' +
      'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
             SqlIN('t2','NameID', Length(AUsedNameIDs)) +
      'GROUP BY t2.NameID, t1.ReasonID ' +
      'ORDER BY t2.NameID');
    QParamDT('BD', BD);
    QParamDT('ED', ED);
    QParamsInt(AUsedNameIDs);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        if AMotorTypeAsSingleName then
          MotorIndex:= VIndexOf(TypeIDs, QFieldInt('TypeID'))
        else
          MotorIndex:= VIndexOf(AUsedNameIDs, QFieldInt('NameID'));
        ReasonIndex:= QFieldInt('ReasonID');
        AClaimCounts[APeriodIndex, ReasonIndex, MotorIndex]:=
          AClaimCounts[APeriodIndex, ReasonIndex, MotorIndex] + QFieldInt('ClaimCount');
        QNext;
        Result:= True;
      end;
    end;
    QClose;
  end;

begin
  Result:= False;
  AClaimCounts:= nil;
  AMotorNeeds:= nil;
  if VIsNil(AUsedNameIDs) then Exit;
  if CompareDate(ABeginDate, AEndDate)>0 then Exit;

  if AMotorTypeAsSingleName then
    GetMotorTypes;

  MDim(AClaimCounts,
       AAdditionYearsCount+1, //кол-во периодов
       5,                     //кол-во критериев (ReasonNames)
       Length(AMotorNames),   //кол-во используемых наименований двигателей
       0                      //кол-во рекламаций
  );

  for i:= 0 to AAdditionYearsCount do
    GetData(i);

  AMotorNeeds:= VCreateBool(Length(AMotorNames), True);
  //ставим метки на двигатели, где нет рекламаций
  if ADoNotUseZeroClaimMotor then
  begin
    SumCounts:= ClaimCountSum(AClaimCounts, VCreateBool(5{кол-во критериев}, True));
    for j:= 0 to High(AMotorNames) do //params (motors)
    begin
      k:= 0;
      for i:= 0 to AAdditionYearsCount do  //years
        k:= k + SumCounts[i, j];
      if k=0 then
        AMotorNeeds[j]:= False;
    end;
  end;
end;

function TDataBase.ReclamationByDefectsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs: TIntVector;
                const ADoNotUseZeroClaimDefect: Boolean;
                out ADefectNames: TStrVector;
                out ADefectNeeds: TBoolVector;
                out AClaimCounts: TIntMatrix3D): Boolean;
var
  i, j, k: Integer;
  SumCounts: TIntMatrix;
  DefectIDs: TIntVector;

  procedure GetData(const APeriodIndex: Integer);
  var
    DefectIndex, ReasonIndex: Integer;
    BD, ED: TDate;
  begin
    BD:= IncYear(ABeginDate, -APeriodIndex);
    ED:= IncYear(AEndDate, -APeriodIndex);

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t1.DefectID) As ClaimCount, t1.DefectID, t1.ReasonID ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
      'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
             SqlIN('t2','NameID', Length(AUsedNameIDs)) +
      'GROUP BY t1.DefectID, t1.ReasonID ' +
      'ORDER BY t1.DefectID');
    QParamDT('BD', BD);
    QParamDT('ED', ED);
    QParamsInt(AUsedNameIDs);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        DefectIndex:= VIndexOf(DefectIDs, QFieldInt('DefectID'));
        ReasonIndex:= QFieldInt('ReasonID');
        AClaimCounts[APeriodIndex, ReasonIndex, DefectIndex]:=
          AClaimCounts[APeriodIndex, ReasonIndex, DefectIndex] + QFieldInt('ClaimCount');
        QNext;
        Result:= True;
      end;
    end;
    QClose;
  end;

begin
  Result:= False;
  AClaimCounts:= nil;
  ADefectNames:= nil;
  ADefectNeeds:= nil;
  if VIsNil(AUsedNameIDs) then Exit;
  if CompareDate(ABeginDate, AEndDate)>0 then Exit;

  KeyPickList('RECLAMATIONDEFECTS', 'DefectID', 'DefectName',
               DefectIDs, ADefectNames, True{ID>0}, 'DefectName');

  MDim(AClaimCounts,
       AAdditionYearsCount+1, //кол-во периодов
       5,                     //кол-во критериев (ReasonNames)
       Length(ADefectNames),   //кол-во используемых наименований двигателей
       0                      //кол-во рекламаций
  );

  for i:= 0 to AAdditionYearsCount do
    GetData(i);

  ADefectNeeds:= VCreateBool(Length(ADefectNames), True);
  //ставим метки на дефекты, где нет рекламаций
  if ADoNotUseZeroClaimDefect then
  begin
    SumCounts:= ClaimCountSum(AClaimCounts, VCreateBool(5{кол-во критериев}, True));
    for j:= 0 to High(ADefectNames) do //params (defects)
    begin
      k:= 0;
      for i:= 0 to AAdditionYearsCount do  //years
        k:= k + SumCounts[i, j];
      if k=0 then
        ADefectNeeds[j]:= False;
    end;
  end;
end;

function TDataBase.ReclamationByPlacesLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs: TIntVector;
                const ADoNotUseZeroClaimPlace: Boolean;
                out APlaceNames: TStrVector;
                out APlaceNeeds: TBoolVector;
                out AClaimCounts: TIntMatrix3D): Boolean;
var
  i, j, k: Integer;
  SumCounts: TIntMatrix;
  PlaceIDs: TIntVector;

  procedure GetData(const APeriodIndex: Integer);
  var
    PlaceIndex, ReasonIndex: Integer;
    BD, ED: TDate;
  begin
    BD:= IncYear(ABeginDate, -APeriodIndex);
    ED:= IncYear(AEndDate, -APeriodIndex);

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t1.PlaceID) As ClaimCount, t1.PlaceID, t1.ReasonID ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
      'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
             SqlIN('t2','NameID', Length(AUsedNameIDs)) +
      'GROUP BY t1.PlaceID, t1.ReasonID ' +
      'ORDER BY t1.PlaceID');
    QParamDT('BD', BD);
    QParamDT('ED', ED);
    QParamsInt(AUsedNameIDs);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        PlaceIndex:= VIndexOf(PlaceIDs, QFieldInt('PlaceID'));
        ReasonIndex:= QFieldInt('ReasonID');
        AClaimCounts[APeriodIndex, ReasonIndex, PlaceIndex]:=
          AClaimCounts[APeriodIndex, ReasonIndex, PlaceIndex] + QFieldInt('ClaimCount');
        QNext;
        Result:= True;
      end;
    end;
    QClose;
  end;

begin
  Result:= False;
  AClaimCounts:= nil;
  APlaceNames:= nil;
  APlaceNeeds:= nil;
  if VIsNil(AUsedNameIDs) then Exit;
  if CompareDate(ABeginDate, AEndDate)>0 then Exit;

  KeyPickList('RECLAMATIONPLACES', 'PlaceID', 'PlaceName',
               PlaceIDs, APlaceNames, True{ID>0}, 'PlaceName');

  MDim(AClaimCounts,
       AAdditionYearsCount+1, //кол-во периодов
       5,                     //кол-во критериев (ReasonNames)
       Length(APlaceNames),   //кол-во используемых наименований двигателей
       0                      //кол-во рекламаций
  );

  for i:= 0 to AAdditionYearsCount do
    GetData(i);

  APlaceNeeds:= VCreateBool(Length(APlaceNames), True);
  //ставим метки на предприятия, где нет рекламаций
  if ADoNotUseZeroClaimPlace then
  begin
    SumCounts:= ClaimCountSum(AClaimCounts, VCreateBool(5{кол-во критериев}, True));
    for j:= 0 to High(APlaceNames) do //params (places)
    begin
      k:= 0;
      for i:= 0 to AAdditionYearsCount do  //years
        k:= k + SumCounts[i, j];
      if k=0 then
        APlaceNeeds[j]:= False;
    end;
  end;
end;

function TDataBase.ReclamationByMonthsLoad(const ABeginDate, AEndDate: TDate;
                const AAdditionYearsCount: Integer;
                const AUsedNameIDs: TIntVector;
                const ADoNotUseZeroClaimMonth: Boolean;
                out AMonthNames: TStrVector;
                out AMonthNeeds: TBoolVector;
                out AClaimCounts: TIntMatrix3D): Boolean;
var
  i, j, k: Integer;
  SumCounts: TIntMatrix;
  BDs, EDs: TDateVector;

  procedure GetMonths;
  var
    n, M: Integer;
    BD, ED: TDate;
    S: String;
  begin
    BDs:= nil;
    EDs:= nil;
    M:= MonthsBetween(ABeginDate, AEndDate);
    for n:= 0 to M do
    begin
      if n=0 then
        BD:= ABeginDate
      else
        BD:= FirstDayInMonth(IncMonth(ABeginDate, n));
      if n=M then
        ED:= AEndDate
      else
        ED:= LastDayInMonth(BD);
      VAppend(BDs, BD);
      VAppend(EDs, ED);
      S:= SFirstUpper(MONTHSNOM[MonthOfDate(BD)]) + SYMBOL_SPACE + IntToStr(YearOfDate(BD));
      VAppend(AMonthNames, S);
    end;
  end;

  procedure GetData(const APeriodIndex: Integer);
  var
    MonthIndex, ReasonIndex: Integer;
    BD, ED: TDate;
  begin
    BD:= IncYear(ABeginDate, -APeriodIndex);
    ED:= IncYear(AEndDate, -APeriodIndex);

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT t1.RecDate, t1.ReasonID ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
      'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
             SqlIN('t2','NameID', Length(AUsedNameIDs)) +
      'ORDER BY t1.RecDate');
    QParamDT('BD', BD);
    QParamDT('ED', ED);
    QParamsInt(AUsedNameIDs);
    QOpen;
    if not QIsEmpty then
    begin
      QFirst;
      while not QEOF do
      begin
        MonthIndex:= VIndexOfDate(BDs, EDs, QFieldDT('RecDate'));
        ReasonIndex:= QFieldInt('ReasonID');
        AClaimCounts[APeriodIndex, ReasonIndex, MonthIndex]:=
          AClaimCounts[APeriodIndex, ReasonIndex, MonthIndex] + 1;
        QNext;
        Result:= True;
      end;
    end;
    QClose;
  end;

begin
  Result:= False;
  AClaimCounts:= nil;
  AMonthNames:= nil;
  AMonthNeeds:= nil;
  if VIsNil(AUsedNameIDs) then Exit;
  if CompareDate(ABeginDate, AEndDate)>0 then Exit;

  GetMonths;

  MDim(AClaimCounts,
       AAdditionYearsCount+1, //кол-во периодов
       5,                     //кол-во критериев (ReasonNames)
       Length(AMonthNames),   //кол-во используемых наименований двигателей
       0                      //кол-во рекламаций
  );

  for i:= 0 to AAdditionYearsCount do
    GetData(i);

  AMonthNeeds:= VCreateBool(Length(AMonthNames), True);
  //ставим метки на месяцы, где нет рекламаций
  if ADoNotUseZeroClaimMonth then
  begin
    SumCounts:= ClaimCountSum(AClaimCounts, VCreateBool(5{кол-во критериев}, True));
    for j:= 0 to High(AMonthNames) do //params (places)
    begin
      k:= 0;
      for i:= 0 to AAdditionYearsCount do  //years
        k:= k + SumCounts[i, j];
      if k=0 then
        AMonthNeeds[j]:= False;
    end;
  end;
end;

{function TDataBase.ReclamationTotalCountLoad(const ABeginDate, AEndDate: TDate;
  const ANameIDs: TIntVector;
  out AExistsNameIDs: TIntVector;
  out AMotorNames: TStrVector; out AMotorCounts: TIntVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;
  AExistsNameIDs:= nil;

  WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) ';
  if not VIsNil(ANameIDs) then
    WhereStr:= WhereStr + 'AND' + SqlIN('t2','NameID', Length(ANameIDs));

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(t2.NameID) As MotorCount, t2.NameID, t3.MotorName ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    'GROUP BY t2.NameID ' +
    'ORDER BY t3.MotorName');
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  if not VIsNil(ANameIDs) then
    QParamsInt(ANameIDs);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(AExistsNameIDs, QFieldInt('NameID'));
      VAppend(AMotorCounts, QFieldInt('MotorCount'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;  }

{function TDataBase.ReclamationDefectsWithReasonsLoad(const ABeginDate,
  AEndDate: TDate; const AAdditionYearsCount: Integer;
  {const AIsZeroIDNeed: Boolean;} const ANameIDs, AReasonIDs: TIntVector; out
  ADefectNames: TStrVector; out AMotorCounts: TIntMatrix3D): Boolean;
begin
  Result:= ReclamationReportWithReasonsLoad('RECLAMATIONDEFECTS',
             'DefectID', 'DefectName', ABeginDate, AEndDate, AAdditionYearsCount,
             {AIsZeroIDNeed,} ANameIDs, AReasonIDs,
             ADefectNames, AMotorCounts);
end; }

{function TDataBase.ReclamationPlacesWithReasonsLoad(const ABeginDate,
  AEndDate: TDate; const AAdditionYearsCount: Integer;
  {const AIsZeroIDNeed: Boolean;} const ANameIDs, AReasonIDs: TIntVector; out
  APlaceNames: TStrVector; out AMotorCounts: TIntMatrix3D): Boolean;
begin
  Result:= ReclamationReportWithReasonsLoad('RECLAMATIONPLACES',
             'PlaceID', 'PlaceName', ABeginDate, AEndDate, AAdditionYearsCount,
             {AIsZeroIDNeed,} ANameIDs, AReasonIDs,
             APlaceNames, AMotorCounts);
end; }

{function TDataBase.ReclamationMonthsWithReasonsLoad(const ABeginDate,
  AEndDate: TDate; const AAdditionYearsCount: Integer; const ANameIDs,
  AReasonIDs: TIntVector; out AMonthNames: TStrVector;
  out AMotorCounts{, AAccumCounts}: TIntMatrix3D): Boolean;
var
  i, FirstMonth,LastMonth, FirstYear, LastYear: Integer;
  MotorCounts{, AccumCounts}: TIntMatrix;

  procedure GetDataForMonth(const AMonth, AYear, AReasonID: Integer; out ACount: Integer);
  var
    WhereStr: String;
    BD, ED: TDate;

    function GetDate(const ADate: TDate): TDate;
    var
      D: Integer;
    begin
      D:= DayOfDate(ADate);
      if (AMonth=2) and (D=29) and (not IsLeapYear(AYear)) then D:= 28;
      Result:= EncodeDate(AYear, AMonth, D);
    end;

  begin
    if AMonth=FirstMonth then
      BD:= GetDate(ABeginDate)
    else
      BD:= FirstDayInMonth(AMonth, AYear);

    if AMonth=LastMonth then
      ED:= GetDate(AEndDate)
    else
      ED:= LastDayInMonth(AMonth, AYear);

    WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
               SqlIN('t2','NameID', Length(ANameIDs), 'NameID');
    if AReasonID>=0 then
      WhereStr:= WhereStr + ' AND (t1.ReasonID = :ReasonID) ';

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t1.RecID) As ACount ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      WhereStr
      );
    QParamDT('BD', BD);
    QParamDT('ED', ED);
    QParamsInt(ANameIDs, 'NameID');
    if AReasonID>=0 then
      QParamInt('ReasonID', AReasonID);
    QOpen;
    ACount:= 0;
    if not QIsEmpty then
    begin
      QFirst;
      ACount:= QFieldInt('ACount');
    end;
    QClose;
  end;

  procedure GetDataForReason(const AYear, AReasonID: Integer; out AYearCounts: TIntVector);
  var
    M, N: Integer;
  begin
    AYearCounts:= nil;
    VDim(AYearCounts, Length(AMonthNames));
    for M:= FirstMonth to LastMonth do
    begin
      GetDataForMonth(M, AYear, AReasonID, N);
      AYearCounts[M-FirstMonth]:= N;
    end;
  end;

  procedure GetDataForYear(const AYear: Integer; out AYearCounts: TIntMatrix);
  var
    k: Integer;
    YearCounts: TIntVector;
  begin
    AYearCounts:= nil;
    //количество для каждой из причин
    for k:=0 to High(AReasonIDs) do
    begin
      GetDataForReason(AYear, AReasonIDs[k], YearCounts);
      MAppend(AYearCounts, YearCounts);
    end;
  end;

begin
  Result:= False;
  AMotorCounts:= nil;
  //AAccumCounts:= nil;

  AMonthNames:= nil;
  FirstMonth:= MonthOfDate(ABeginDate);
  LastMonth:= MonthOfDate(AEndDate);
  for i:= FirstMonth to LastMonth do
    VAppend(AMonthNames, SFirstUpper(MONTHSNOM[i]));

  FirstYear:= YearOfDate(ABeginDate) - AAdditionYearsCount;
  LastYear:= YearOfDate(AEndDate);

  for i:= FirstYear to LastYear do
  begin
    GetDataForYear(i, MotorCounts{, AccumCounts});
    MAppend(AMotorCounts, MotorCounts);
    //MAppend(AAccumCounts, AccumCounts);
  end;

  Result:= True;
end;}

function TDataBase.ReclamationMileagesWithReasonsLoad(const ABeginDate,
  AEndDate: TDate; const AAdditionYearsCount: Integer; const ANameIDs,
  AReasonIDs: TIntVector; out AMileageNames: TStrVector; out
  AMotorCounts: TIntMatrix3D): Boolean;
var
  MileageMins, MileageMaxs: TIntVector;
  MotorCounts: TIntMatrix;
  i: Integer;
  BD, ED: TDate;

  procedure GetDataForMileage(const ABD, AED: TDate;
                              const AReasonID, AMileageMin, AMileageMax: Integer;
                              out ACount: Integer);
  var
    WhereStr: String;
  begin
    WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) AND ' +
                     '(t1.Mileage BETWEEN :MileageMin AND :MileageMax) AND ' +
               SqlIN('t2','NameID', Length(ANameIDs), 'NameID');
    if AReasonID>=0 then
      WhereStr:= WhereStr + ' AND (t1.ReasonID = :ReasonID) ';

    QSetQuery(FQuery);
    QSetSQL(
      'SELECT COUNT(t1.RecID) As ACount ' +
      'FROM RECLAMATIONS t1 ' +
      'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
      WhereStr
      );
    QParamDT('BD', ABD);
    QParamDT('ED', AED);
    QParamsInt(ANameIDs, 'NameID');

    QParamInt('MileageMin', AMileageMin);
    QParamInt('MileageMax', AMileageMax);
    if AReasonID>=0 then
      QParamInt('ReasonID', AReasonID);
    QOpen;
    ACount:= 0;
    if not QIsEmpty then
    begin
      QFirst;
      ACount:= QFieldInt('ACount');
    end;
    QClose;
  end;

  procedure GetDataForReason(const ABD, AED: TDate; const AReasonID: Integer; out ACounts: TIntVector);
  var
    n, Count, MileageMin, MileageMax: Integer;
  begin
    ACounts:= nil;
    VDim(ACounts, Length(MileageMins));
    for n:= 0 to High(ACounts) do
    begin
      if (n=0) or (n=High(ACounts)) then
        MileageMin:= MileageMins[n]
      else
        MileageMin:= MileageMins[n]+1;
      MileageMax:= MileageMaxs[n];

      GetDataForMileage(ABD, AED, AReasonID, MileageMin, MileageMax, Count);
      ACounts[n]:= Count;
    end;
  end;

  procedure GetDataForSinglePeriod(const ABD, AED: TDate; out ACounts: TIntMatrix);
  var
    Counts: TIntVector;
    n: Integer;
  begin
    ACounts:= nil;
    for n:=0 to High(AReasonIDs) do
    begin
      GetDataForReason(ABD, AED, AReasonIDs[n], Counts);
      MAppend(ACounts, Counts);
    end;
  end;

begin
  Result:= False;
  AMileageNames:= nil;
  AMotorCounts:= nil;

  MileageMins:= nil;
  MileageMaxs:= nil;

  MileageMins:= VCreateInt([0, 50, 100, 150, 200, 250, 300]);
  VDim(MileageMaxs, Length(MileageMins));
  VDim(AMileageNames, Length(MileageMins));
  for i:= 0 to High(MileageMins) do
  begin
    MileageMaxs[i]:= MileageMins[i]+50;
    AMileageNames[i]:= IntToStr(MileageMins[i]) + '-' + IntToStr(MileageMaxs[i]);

    MileageMaxs[i]:= MileageMaxs[i]*1000;
    MileageMins[i]:= MileageMins[i]*1000;
  end;
  VAppend(MileageMins, -1);
  VAppend(MileageMaxs, -1);
  VAppend(AMileageNames, 'Не указан');

  for i:= AAdditionYearsCount downto 0 do
  begin
    BD:= IncYear(ABeginDate, -i);
    ED:= IncYear(AEndDate, -i);
    GetDataForSinglePeriod(BD, ED, MotorCounts);
    //итоговая матрица для всех периодов
    MAppend(AMotorCounts, MotorCounts);
  end;

  Result:= True;
end;






end.

