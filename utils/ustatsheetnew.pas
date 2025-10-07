unit UStatSheetNew;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics, fpstypes, {Math,}
  //DK packages utils
  DK_Vector, DK_Matrix, DK_SheetWriter, DK_SheetTypes, DK_StrUtils, DK_Math;

const

  COL_COUNT_GRAPH = 120;
  COL_COUNT_TOTAL = 1{ParamNames} + COL_COUNT_GRAPH;
  COL_COUNT_USED  = COL_COUNT_GRAPH div 3;
  COL_COUNT_VALUE = 10;

  COL_WIDTH_GRAPH = 8;
  COL_WIDTH_NAMES = 150; //defaul

  ROW_EMPTY_HEIGH = 5;

  REPORT_TITLE_MAIN = 'Отчет по рекламационным случаям электродвигателей';

  GRAPH_TITLE_MAIN   = 'Распределение количества рекламационных случаев по ';
  GRAPH_TITLE_DEFECT = GRAPH_TITLE_MAIN + 'причинам возникновения неисправностей';

  //5 years, 5 reasons
  GRAPH_COLORS_DEFAULT: array [0..4] of TColor = (
    $00BD814F, //blue
    $004696F7, //orange
    $0059BB9B, //green
    $00A26480, //violet
    $004D50C0  //red
  );

type

  { TStatSheet }

  TStatSheet = class (TCustomSheet)
  protected
    function SetWidths: TIntVector; override;
  private
    procedure HorizBarDraw(var ARow: Integer;
                              const AName: String;
                              const AValue, AMaxValue: Integer);
    procedure HorizBarHistogramDraw(var ARow: Integer;
                                    const ANames: TStrVector;
                                    const AValues: TIntVector);
  public
    procedure Draw(const AParamTitle: String; //дательный падеж
                   const AReasonNames, AParamNames: TStrVector;
                   const AReasonValues, AParamValues: TIntVector);
  end;



implementation

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
                                    const AValues: TIntVector);
var
  R, i, MaxValue: Integer;
  Names: TStrVector;
  Values, Indexes: TIntVector;
begin
  R:= ARow;

  MaxValue:= VMax(AValues);

  VSort(AValues, Indexes, True{DESC});
  Names:= VReplace(ANames, Indexes);
  Values:= VReplace(AValues, Indexes);

  for i:= 0 to High(Names) do
    HorizBarDraw(R, Names[i], Values[i], MaxValue);



  ARow:= R;
end;

procedure TStatSheet.Draw(const AParamTitle: String; //дательный падеж
                          const AReasonNames, AParamNames: TStrVector;
                          const AReasonValues, AParamValues: TIntVector);
var
  R: Integer;
begin
  Writer.ColWidth[1]:= MMaxWidth([AReasonNames, AParamNames], Font);

  Writer.BeginEdit;

  Writer.SetBackgroundClear;
  Writer.SetAlignment(haCenter, vaCenter);

  R:= 1;
  Writer.SetFont(Font.Name, Font.Size+2, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount, REPORT_TITLE_MAIN, cbtNone, True, True);

  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount, GRAPH_TITLE_MAIN + AParamTitle, cbtNone, True, True);
  R:= R + 1;
  HorizBarHistogramDraw(R, AParamNames, AParamValues);

  R:= R + 2;
  Writer.SetFont(Font.Name, Font.Size, [fsBold], clBlack);
  Writer.WriteText(R, 1, R, Writer.ColCount, GRAPH_TITLE_DEFECT, cbtNone, True, True);
  R:= R + 1;
  HorizBarHistogramDraw(R, AReasonNames, AReasonValues);

  Writer.EndEdit;
end;

end.

