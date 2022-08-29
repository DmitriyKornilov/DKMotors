unit USQLite;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, StdCtrls, DK_SQLite3, DK_SQLUtils, DK_DateUtils,
  DK_Vector, DK_Matrix, DK_StrUtils;

type

  { TSQLite }

  TSQLite = class (TSQLite3)
  private
    procedure IDsAndNamesLoad(AComboBox: TComboBox; out ANameIDs: TIntVector;
       const ATableName, AKeyFieldName, APickFieldName, AOrderFieldName: String;
       const AKeyValueNotZero: Boolean; const AZeroKeyPick: String = '');
    function MonthAndDatesForLogLoad(const ATableName, AFieldName: String;
       const AYear: Word; out AMonths: TStrVector; out ADates: TDateMatrix): Boolean;
    function ReclamationReportLoad(const ATableName, AIDFieldName, ANameFieldName: String;
       const ABeginDate, AEndDate: TDate; const ANameID: Integer;
       out ANames: TStrVector; out ACounts: TIntVector): Boolean;
  public
    //справочники
    procedure NameIDsAndMotorNamesLoad(AComboBox: TComboBox;
                  out AIDs: TIntVector; const AKeyValueNotZero: Boolean = True);
    procedure ReceiverIDsAndNamesLoad(AComboBox: TComboBox;
                  out AIDs: TIntVector; const AKeyValueNotZero: Boolean = True);


    //дерево месяцы-даты для журналов учета
    function MonthAndDatesForBuildLogLoad(const AYear: Word;
                  out AMonths: TStrVector; out ADates: TDateMatrix): Boolean;
    function MonthAndDatesForTestLogLoad(const AYear: Word;
                  out AMonths: TStrVector; out ADates: TDateMatrix): Boolean;

    //список двигателей
    function MotorListLoad(const ABuildYear, ANameID, AShippedType: Integer;
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
               const ANameID: Integer; const ANeedOrderByNumber: Boolean;
               out AMotorIDs, ANameIDs, AOldMotors: TIntVector;
               out ABuildDates: TDateVector;
               out AMotorNames, AMotorNums, ARotorNums: TStrVector): Boolean;
    function BuildTotalLoad(const ABeginDate, AEndDate: TDate;
               const ANameID: Integer;
               out AMotorNames: TStrVector; out AMotorCounts: TIntVector): Boolean;
    function IsDuplicateMotorNumber(const ADate: TDate;
               const ANameID: Integer; const AMotorNum: String): Boolean;
    function MotorsInBuildLogWrite(const ABuildDate: TDate;
               const ANameIDs, AOldMotors: TIntVector;
               const AMotorNums, ARotorNums: TStrVector): Boolean;
    procedure MotorInBuildLogUpdate(const AMotorID: Integer;
                          const ABuildDate: TDate; const ANameID, AOldMotor: Integer;
                          const AMotorNum, ARotorNum: String);

    //испытания
    function TestBeforeListLoad(const ANameID: Integer;
                          const ANeedOrderByNumber: Boolean;
                          out ABuildDates: TDateVector;
                          out AMotorNames, AMotorNums, ATotalMotorNames: TStrVector;
                          out ATotalMotorCounts: TIntVector;
                          out ATestDates: TDateMatrix;
                          out ATestFails: TIntMatrix;
                          out ATestNotes: TStrMatrix): Boolean;
    function TestListLoad(const ABeginBuildDate, AEndBuildDate: TDate;
               const ANameID: Integer; const ANeedOrderByNumber: Boolean;
               out ATestIDs, ATestResults: TIntVector;
               out ATestDates: TDateVector;
               out AMotorNames, AMotorNums, ATestNotes: TStrVector): Boolean;
    function TestTotalLoad(const ABeginBuildDate, AEndBuildDate: TDate;
                const ANameID: Integer;
                out AMotorNames: TStrVector;
                out AMotorCounts, AFailCounts: TIntVector): Boolean;
    function TestChooseListLoad(const ANameID: Integer;
                         const ANumberLike: String;
                         out AMotorIDs: TIntVector;
                         out AMotorNums, ABuildDates, ATests: TStrVector): Boolean;
    function MotorsInTestLogWrite(const ATestDate: TDate;
                           const AMotorIDs, ATestResults: TIntVector;
                           const ATestNotes: TStrVector): Boolean;

    //склад
    function StoreListLoad(const ANameID: Integer;
                    const ANeedOrderByNumber, ANeedOrderByName: Boolean;
                    out ATestDates: TDateVector;
                    out AMotorNames, AMotorNums: TStrVector): Boolean;
    function StoreTotalLoad(const ANameID: Integer;
                             out AMotorNames: TStrVector;
                             out AMotorCounts: TIntVector): Boolean;

    //отгрузки
    function ShipmentListLoad(const AYear: Word; out AMonths: TStrVector;
                          out AShipments: TStrMatrix;
                          out ACargoIDs: TIntMatrix): Boolean;
    function ShipmentMotorListLoad(const ABeginDate, AEndDate: TDate;
                const ANameID, AReceiverID: Integer;
                out ASendDates: TDateVector;
                out AMotorNames, AMotorNums, AReceiverNames: TStrVector): Boolean;
    function ShipmentTotalLoad(const ABeginDate, AEndDate: TDate;
                const ANameID: Integer;
                out AMotorNames: TStrVector;
                out AMotorCounts: TIntVector): Boolean;
    function ShipmentRecieversTotalLoad(const ABeginDate, AEndDate: TDate;
                const ANameID, AReceiverID: Integer;
                out AReceiverNames: TStrVector;
                out AMotorNames: TStrMatrix;
                out AMotorCounts: TIntMatrix): Boolean;
    function CargoListLoad(const ABeginDate, AEndDate: TDate;
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
          const ANameID: Integer; const ANumberLike: String;
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
    function ReclamationTotalLoad(const ABeginDate, AEndDate: TDate;
                const ANameID: Integer;
                out AMotorNames: TStrVector;
                out AMotorCounts: TIntVector): Boolean;
    function ReclamationPlacesLoad(const ABeginDate, AEndDate: TDate;
                const ANameID: Integer;
                out APlaceNames: TStrVector;
                out APlaceMotorCounts: TIntVector): Boolean;
    function ReclamationDefectsLoad(const ABeginDate, AEndDate: TDate;
                const ANameID: Integer;
                out ADefectNames: TStrVector;
                out ADefectMotorCounts: TIntVector): Boolean;
    function ReclamationReasonsLoad(const ABeginDate, AEndDate: TDate;
                const ANameID: Integer;
                out AReasonNames: TStrVector;
                out AReasonMotorCounts: TIntVector): Boolean;

  end;

var
  SQLite: TSQLite;

implementation

{ TSQLite }


procedure TSQLite.IDsAndNamesLoad(AComboBox: TComboBox; out ANameIDs: TIntVector;
   const ATableName, AKeyFieldName, APickFieldName, AOrderFieldName: String;
   const AKeyValueNotZero: Boolean; const AZeroKeyPick: String = '');

var
  MN: TStrVector;
  Ind: Integer;
begin
  KeyPickList(ATableName, AKeyFieldName, APickFieldName,
              ANameIDs, MN, AKeyValueNotZero, AOrderFieldName);
  if VIsNil(ANameIDs) then Exit;

  if not AKeyValueNotZero then
  begin
    Ind:= VIndexOf(ANameIDs, 0);
    if Ind>=0 then
      MN[Ind]:= AZeroKeyPick;
  end;

  VToStrings(MN, AComboBox.Items);
  AComboBox.ItemIndex:= 0;
end;

procedure TSQLite.NameIDsAndMotorNamesLoad(AComboBox: TComboBox;
  out AIDs: TIntVector; const AKeyValueNotZero: Boolean = True);
begin
  IDsAndNamesLoad(AComboBox, AIDs,
                  'MOTORNAMES', 'NameID', 'MotorName', 'NameID',
                  AKeyValueNotZero, 'ВСЕ НАИМЕНОВАНИЯ');
end;

procedure TSQLite.ReceiverIDsAndNamesLoad(AComboBox: TComboBox;
  out AIDs: TIntVector; const AKeyValueNotZero: Boolean = True);
begin
  IDsAndNamesLoad(AComboBox, AIDs,
                  'CARGORECEIVERS', 'ReceiverID', 'ReceiverName', 'ReceiverName',
                  AKeyValueNotZero, 'ВСЕ ГРУЗОПОЛУЧАТЕЛИ');
end;



function TSQLite.MonthAndDatesForLogLoad(const ATableName, AFieldName: String;
  const AYear: Word; out AMonths: TStrVector; out ADates: TDateMatrix): Boolean;
var
  i: Integer;
  BD, ED: TDate;
  Dates: TDateVector;
  FieldName: String;
begin
  Result:= False;
  AMonths:= nil;
  ADates:= nil;

  FieldName:= SqlEsc(AFieldName);
  QSetQuery(FQuery);
  QSetSQL(
    'SELECT DISTINCT' + FieldName +
    'FROM' + SqlEsc(ATableName) +
    'WHERE' +  FieldName + 'BETWEEN :BD AND :ED ' +
    'ORDER BY' + FieldName + 'DESC');
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
        VAppend(Dates, QFieldDT(AFieldName));
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

function TSQLite.ReclamationReportLoad(const ATableName, AIDFieldName,
  ANameFieldName: String; const ABeginDate, AEndDate: TDate;
  const ANameID: Integer; out ANames: TStrVector; out ACounts: TIntVector): Boolean;
var
  WhereStr, IDFieldName, NameFieldName: String;
begin
  Result:= False;
  ANames:= nil;
  ACounts:= nil;

  IDFieldName:= SqlEsc(AIDFieldName);
  NameFieldName:= SqlEsc(ANameFieldName);

  WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) ';
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t2.NameID = :NameID) ';

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(t1.' + IDFieldName + ') As MotorCount, t3.' + NameFieldName + ' ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN ' + SqlEsc(ATableName) + ' t3 ON (t1.' + IDFieldName + '=t3.' + IDFieldName + ') ' +
    WhereStr +
    'GROUP BY t1.' + IDFieldName + ' ' +
    'ORDER BY t3.' + NameFieldName);
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  if ANameID>0 then
    QParamInt('NameID', ANameID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ACounts, QFieldInt('MotorCount'));
      VAppend(ANames, QFieldStr(ANameFieldName));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TSQLite.MonthAndDatesForBuildLogLoad(const AYear: Word; out
  AMonths: TStrVector; out ADates: TDateMatrix): Boolean;
begin
  Result:= MonthAndDatesForLogLoad('MOTORLIST', 'BuildDate', AYear, AMonths, ADates);
end;

function TSQLite.MonthAndDatesForTestLogLoad(const AYear: Word; out
  AMonths: TStrVector; out ADates: TDateMatrix): Boolean;
begin
  Result:= MonthAndDatesForLogLoad('MOTORTEST', 'TestDate', AYear, AMonths, ADates);
end;

function TSQLite.MotorListLoad(const ABuildYear, ANameID,
  AShippedType: Integer; const ANumberLike: String;
  const ANeedOrderByNumber: Boolean; out AMotorIDs: TIntVector; out
  ABuildDates, AMotorNames, AMotorNums, AShippings: TStrVector): Boolean;
var
  OrderStr, WhereStr: String;
  SendDates, BuildDates: TDateVector;
  ReceiverNames: TStrVector;
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
    WhereStr:= 'WHERE (t1.BuildDate BETWEEN :BD AND :ED) '
  else
    WhereStr:= 'WHERE (UPPER(t1.MotorNum) LIKE :NumberLike) ';

  if ANameID>0 then
    WhereStr:= WhereStr + 'AND (t1.NameID = :NameID) ';

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
    QParamStr('NumberLike', ANumberLike+'%');

  if ANameID>0 then
    QParamInt('NameID', ANameID);

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

  if Result then
  begin
    ABuildDates:= VFormatDateTime('dd.mm.yyyy', BuildDates);
    AShippings:= VFormatDateTime('dd.mm.yyyy – ', SendDates);
    VChangeIf(AShippings, '30.12.1899 – ', EmptyStr);
    VChangeIf(ReceiverNames, '<не указан>', EmptyStr);
    AShippings:= VSum(AShippings, ReceiverNames);
  end;
end;

procedure TSQLite.MotorInfoLoad(const AMotorID: Integer; out ABuildDate,
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

function TSQLite.BuildListLoad(const ABeginDate, AEndDate: TDate;
                  const ANameID: Integer;
                  const ANeedOrderByNumber: Boolean;
                  out AMotorIDs, ANameIDs, AOldMotors: TIntVector;
                  out ABuildDates: TDateVector;
                  out AMotorNames, AMotorNums, ARotorNums: TStrVector): Boolean;
var
  S, WhereStr, OrderStr: String;

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
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t1.NameID = :NameID) ';

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
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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
end;

function TSQLite.BuildTotalLoad(const ABeginDate, AEndDate: TDate;
  const ANameID: Integer; out AMotorNames: TStrVector; out
  AMotorCounts: TIntVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;

  WhereStr:= 'WHERE (t1.BuildDate BETWEEN :BD AND :ED) ';
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t1.NameID = :NameID) ';

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
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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

function TSQLite.IsDuplicateMotorNumber(const ADate: TDate;
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

function TSQLite.MotorsInBuildLogWrite(const ABuildDate: TDate; const ANameIDs,
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

procedure TSQLite.MotorInBuildLogUpdate(const AMotorID: Integer;
  const ABuildDate: TDate; const ANameID, AOldMotor: Integer; const AMotorNum,
  ARotorNum: String);
begin
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
  except
    QRollback;
  end;
end;

function TSQLite.TestBeforeListLoad(const ANameID: Integer;
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
  if ANameID>0 then
    WhereStr:= WhereStr + 'AND (t1.NameID = :NameID) ';

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

  if ANameID>0 then
    QParamInt('NameID', ANameID);
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
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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

function TSQLite.TestListLoad(const ABeginBuildDate, AEndBuildDate: TDate;
  const ANameID: Integer; const ANeedOrderByNumber: Boolean; out ATestIDs,
  ATestResults: TIntVector; out ATestDates: TDateVector; out AMotorNames,
  AMotorNums, ATestNotes: TStrVector): Boolean;
var
  OrderStr: String;
  WhereStr: String;
begin
  Result:= False;
  ATestIDs:= nil;
  ATestResults:= nil;
  ATestDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  ATestNotes:= nil;

  WhereStr:= 'WHERE (t1.TestDate BETWEEN :BD AND :ED) ';
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t2.NameID = :NameID) ';

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
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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
end;

function TSQLite.TestTotalLoad(const ABeginBuildDate, AEndBuildDate: TDate;
  const ANameID: Integer; out AMotorNames: TStrVector; out AMotorCounts,
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
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t2.NameID = :NameID) ';

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
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t2.NameID = :NameID) ';

  QSetSQL(
    'SELECT COUNT(t2.NameID) As MotorCount, t2.NameID, t3.MotorName ' +
    'FROM MOTORTEST t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    'GROUP BY t2.NameID');
  QParamDT('BD', ABeginBuildDate);
  QParamDT('ED', AEndBuildDate);
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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

function TSQLite.TestChooseListLoad(const ANameID: Integer;
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

  if ANumberLike=EmptyStr then Exit;

  BuildDates:= nil;

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT MotorID, BuildDate, MotorNum ' +
    'FROM MOTORLIST ' +
    'WHERE (NameID = :NameID) AND ' +
          '(UPPER(MotorNum) LIKE :NumberLike)');
  QParamInt('NameID', ANameID);
  QParamStr('NumberLike', ANumberLike+'%');
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

function TSQLite.MotorsInTestLogWrite(const ATestDate: TDate; const AMotorIDs,
  ATestResults: TIntVector; const ATestNotes: TStrVector): Boolean;
var
  i: Integer;
begin
  Result:= False;
  if VIsNil(AMotorIDs) then Exit;
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

function TSQLite.StoreListLoad(const ANameID: Integer;
  const ANeedOrderByNumber, ANeedOrderByName: Boolean; out
  ATestDates: TDateVector; out AMotorNames, AMotorNums: TStrVector): Boolean;
var
  OrderStr: String;
  WhereStr: String;
begin
  Result:= False;
  ATestDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;

  WhereStr:= 'WHERE (t1.Fail=0) AND (t2.CargoID=0) ';
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t2.NameID = :NameID) ';

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
    'SELECT t1.TestDate, ' +
           't2.MotorNum, t2.Series, t3.MotorName ' +
    'FROM MOTORTEST t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    OrderStr);
  if ANameID>0 then
    QParamInt('NameID', ANameID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ATestDates, QFieldDT('TestDate'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TSQLite.StoreTotalLoad(const ANameID: Integer; out
  AMotorNames: TStrVector; out AMotorCounts: TIntVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;

  WhereStr:= 'WHERE (t1.Fail=0) AND (t2.CargoID=0) ';
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t2.NameID = :NameID) ';

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(t2.NameID) As MotorCount, t2.NameID, t3.MotorName ' +
    'FROM MOTORTEST t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    'GROUP BY t2.NameID ' +
    'ORDER BY t3.MotorName');
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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

function TSQLite.ShipmentListLoad(const AYear: Word; out AMonths: TStrVector;
  out AShipments: TStrMatrix; out ACargoIDs: TIntMatrix): Boolean;
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
    CargoListLoad(BD, ED, CargoIDs, SendDates, ReceiverNames);
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

function TSQLite.ShipmentMotorListLoad(const ABeginDate, AEndDate: TDate;
  const ANameID, AReceiverID: Integer; out ASendDates: TDateVector; out
  AMotorNames, AMotorNums, AReceiverNames: TStrVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  ASendDates:= nil;
  AMotorNames:= nil;
  AMotorNums:= nil;
  AReceiverNames:= nil;

  WhereStr:= 'WHERE (t1.CargoID>0) AND (t2.SendDate BETWEEN :BD AND :ED) ';
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t1.NameID = :NameID) ';
  if AReceiverID>0 then
    WhereStr:= WhereStr + ' AND (t2.ReceiverID = :ReceiverID) ';

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT t1.MotorNum, t1.Series, t2.SendDate, t3.ReceiverName, t4.MotorName ' +
    'FROM MOTORLIST t1 ' +
    'INNER JOIN CARGOLIST t2 ON (t1.CargoID=t2.CargoID) ' +
    'INNER JOIN CARGORECEIVERS t3 ON (t2.ReceiverID=t3.ReceiverID) ' +
    'INNER JOIN MOTORNAMES t4 ON (t1.NameID=t4.NameID) ' +
    WhereStr +
    'ORDER BY t2.SendDate, t4.MotorName, t1.MotorNum');
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  if ANameID>0 then
    QParamInt('NameID', ANameID);
  if AReceiverID>0 then
    QParamInt('ReceiverID', AReceiverID);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ASendDates, QFieldDT('SendDate'));
      VAppend(AMotorNames, QFieldStr('MotorName'));
      VAppend(AMotorNums, QFieldStr('MotorNum') + ' (' + QFieldStr('Series') + ')');
      VAppend(AReceiverNames, QFieldStr('ReceiverName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TSQLite.ShipmentTotalLoad(const ABeginDate, AEndDate: TDate;
  const ANameID: Integer; out AMotorNames: TStrVector; out
  AMotorCounts: TIntVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;

  WhereStr:= 'WHERE (t1.CargoID>0) AND (t2.SendDate BETWEEN :BD AND :ED) ';
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t1.NameID = :NameID) ';

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
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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

function TSQLite.ShipmentRecieversTotalLoad(const ABeginDate, AEndDate: TDate;
  const ANameID, AReceiverID: Integer; out AReceiverNames: TStrVector; out
  AMotorNames: TStrMatrix; out AMotorCounts: TIntMatrix): Boolean;
var
  WhereStr: String;
  ReceiverIDs, MotorCounts: TIntVector;
  MotorNames: TStrVector;
  i: Integer;

  procedure GetReceiverNames;
  begin
    WhereStr:= 'WHERE (t1.CargoID>0) AND (t2.SendDate BETWEEN :BD AND :ED) ';
    if ANameID>0 then
      WhereStr:= WhereStr + ' AND (t1.NameID = :NameID) ';
    if AReceiverID>0 then
      WhereStr:= WhereStr + ' AND (t2.ReceiverID = :ReceiverID) ';


    QSetSQL(
      'SELECT DISTINCT t2.ReceiverID, t3.ReceiverName ' +
      'FROM MOTORLIST t1 ' +
      'INNER JOIN CARGOLIST t2 ON (t1.CargoID=t2.CargoID) ' +
      'INNER JOIN CARGORECEIVERS t3 ON (t2.ReceiverID=t3.ReceiverID) ' +
      WhereStr +
      'ORDER BY t3.ReceiverName');
    QParamDT('BD', ABeginDate);
    QParamDT('ED', AEndDate);
    if ANameID>0 then
      QParamInt('NameID', ANameID);
    if AReceiverID>0 then
      QParamInt('ReceiverID', AReceiverID);
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
    if ANameID>0 then
      WhereStr:= WhereStr + ' AND (t1.NameID = :NameID) ';
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
    if ANameID>0 then
      QParamInt('NameID', ANameID);
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

function TSQLite.CargoListLoad(const ABeginDate, AEndDate: TDate; out
  ACargoIDs: TIntVector; out ASendDates: TDateVector; out
  AReceiverNames: TStrVector): Boolean;
begin
  Result:= False;
  ACargoIDs:= nil;
  ASendDates:= nil;
  AReceiverNames:= nil;

  QSetQuery(FQuery);
  QClose;
  QSetSQL(
    'SELECT t1.CargoID, t1.SendDate, t2.ReceiverName ' +
    'FROM CARGOLIST t1 ' +
    'INNER JOIN CARGORECEIVERS t2 ON (t1.ReceiverID=t2.ReceiverID) ' +
    'WHERE (t1.SendDate >= :BeginDate) AND (t1.SendDate <= :EndDate) ' +
    'ORDER BY t1.SendDate DESC');
  QParamDT('BeginDate', ABeginDate);
  QParamDT('EndDate', AEndDate);
  QOpen;
  if not QIsEmpty then
  begin
    QFirst;
    while not QEOF do
    begin
      VAppend(ASendDates, QFieldDT('SendDate'));
      VAppend(ACargoIDs, QFieldInt('CargoID'));
      VAppend(AReceiverNames, QFieldStr('ReceiverName'));
      QNext;
    end;
    Result:= True;
  end;
  QClose;
end;

function TSQLite.CargoChooseListLoad(const ANameID: Integer;
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
    'WHERE (NameID = :NameID) AND (CargoID=0) AND ' +
          '(UPPER(MotorNum) LIKE :NumberLike)');
  QParamInt('NameID', ANameID);
  QParamStr('NumberLike', ANumberLike+'%');
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

function TSQLite.CargoMotorListLoad(const ACargoID: Integer; out
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

function TSQLite.CargoLoad(const ACargoID: Integer; out ASendDate: TDate; out
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

function TSQLite.CargoWrite(const ASendDate: TDate; const AReceiverID: Integer;
  const AMotorIDs: TIntVector; const ASeries: TStrVector): Boolean;
var
  i, CargoID: Integer;
begin
  Result:= False;
  if VIsNil(AMotorIDs) then Exit;
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

function TSQLite.CargoUpdate(const ACargoID: Integer; const ASendDate: TDate;
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

function TSQLite.ReclamationListLoad(const AMotorID: Integer; out
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

function TSQLite.ReclamationListLoad(const ABeginDate, AEndDate: TDate;
  const ANameID: Integer; const ANumberLike: String; out ARecDates,
  ABuildDates, AArrivalDates, ASendingDates: TDateVector; out ARecIDs,
  AMotorIDs, AMileages, AOpinions, AColors, APassports: TIntVector; out
  APlaceNames, AFactoryNames, ADepartures, ADefectNames, AReasonNames,
  ARecNotes, AMotorNames, AMotorNums: TStrVector): Boolean;
var
  WhereStr: String;
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

  if ANumberLike=EmptyStr then
    WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) '
  else
    WhereStr:= 'WHERE (UPPER(t6.MotorNum) LIKE :NumberLike) ';

  if ANameID>0 then
    WhereStr:= WhereStr + 'AND (t6.NameID = :NameID) ';

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
    'ORDER BY t1.RecDate');

  if ANumberLike=EmptyStr then
  begin
    QParamDT('BD', ABeginDate);
    QParamDT('ED', AEndDate);
  end
  else
    QParamStr('NumberLike', ANumberLike+'%');

  if ANameID>0 then
    QParamInt('NameID', ANameID);

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
end;

function TSQLite.ReclamationChooseListLoad(const ANameID: Integer;
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
  QParamStr('NumberLike', ANumberLike+'%');
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

function TSQLite.ReclamationWrite(const ARecDate: TDate; const AMotorID,
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

function TSQLite.ReclamationUpdate(const ARecDate, AArrivalDate,
  ASendingDate: TDate; const ARecID, AMotorID, AMileage, APlaceID, AFactoryID,
  ADefectID, AReasonID, AOpinion, APassport: Integer; const ADeparture,
  ARecNote: String): Boolean;
begin
  Result:= False;
  QSetQuery(FQuery);
  try
    QSetSQL('UPDATE RECLAMATIONS ' +
            'SET RecDate=:RecDate, MotorID=:MotorID, Mileage=:Mileage, ' +
                'PlaceID=:PlaceID, FactoryID=:FactoryID, ' +
                'Departure=:Departure, DefectID=:DefectID, ' +
                'ReasonID=:ReasonID, Opinion=:Opinion, RecNote=:RecNote, ' +
                'ArrivalDate=:ArrivalDate, SendingDate=:SendingDate, Passport=:Passport ' +
            'WHERE RecID = :RecID');
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
    QParamDT('ArrivalDate', AArrivalDate);
    QParamDT('SendingDate', ASendingDate);
    QParamInt('Passport', APassport);
    QExec;
    QCommit;
    Result:= True;
  except
    QRollback;
  end;
end;

function TSQLite.ReclamationLoad(const ARecID: Integer; out ARecDate,
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

function TSQLite.ReclamationTotalLoad(const ABeginDate, AEndDate: TDate;
  const ANameID: Integer; out AMotorNames: TStrVector; out
  AMotorCounts: TIntVector): Boolean;
var
  WhereStr: String;
begin
  Result:= False;
  AMotorNames:= nil;
  AMotorCounts:= nil;

  WhereStr:= 'WHERE (t1.RecDate BETWEEN :BD AND :ED) ';
  if ANameID>0 then
    WhereStr:= WhereStr + ' AND (t2.NameID = :NameID) ';

  QSetQuery(FQuery);
  QSetSQL(
    'SELECT COUNT(t2.NameID) As MotorCount, t3.MotorName ' +
    'FROM RECLAMATIONS t1 ' +
    'INNER JOIN MOTORLIST t2 ON (t1.MotorID=t2.MotorID) ' +
    'INNER JOIN MOTORNAMES t3 ON (t2.NameID=t3.NameID) ' +
    WhereStr +
    'GROUP BY t2.NameID ' +
    'ORDER BY t3.MotorName');
  QParamDT('BD', ABeginDate);
  QParamDT('ED', AEndDate);
  if ANameID>0 then
    QParamInt('NameID', ANameID);
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

function TSQLite.ReclamationPlacesLoad(const ABeginDate, AEndDate: TDate;
  const ANameID: Integer; out APlaceNames: TStrVector; out
  APlaceMotorCounts: TIntVector): Boolean;
begin
  Result:= ReclamationReportLoad('RECLAMATIONPLACES',
             'PlaceID', 'PlaceName', ABeginDate, AEndDate, ANameID,
             APlaceNames, APlaceMotorCounts);
end;

function TSQLite.ReclamationDefectsLoad(const ABeginDate, AEndDate: TDate;
  const ANameID: Integer; out ADefectNames: TStrVector; out
  ADefectMotorCounts: TIntVector): Boolean;
begin
  Result:= ReclamationReportLoad('RECLAMATIONDEFECTS',
             'DefectID', 'DefectName', ABeginDate, AEndDate, ANameID,
             ADefectNames, ADefectMotorCounts);
end;

function TSQLite.ReclamationReasonsLoad(const ABeginDate, AEndDate: TDate;
  const ANameID: Integer; out AReasonNames: TStrVector; out
  AReasonMotorCounts: TIntVector): Boolean;
begin
  Result:= ReclamationReportLoad('RECLAMATIONREASONS',
             'ReasonID', 'ReasonName', ABeginDate, AEndDate, ANameID,
             AReasonNames, AReasonMotorCounts);
end;


end.

