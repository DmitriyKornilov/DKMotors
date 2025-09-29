unit UVars;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics,
  //DK packages utils
  DK_Fonts,
  //Project utils
  UDataBase, UImages;

var
  Images: TImages;
  DataBase: TDataBase;
  GridFont: TFont;

  procedure GlobalVarCreate;
  procedure GlobalVarDestroy;
  procedure GlobalVarInit;

implementation

procedure GlobalVarCreate;
begin
  Images:= TImages.Create(nil);
  GridFont:= TFont.Create;
  DataBase:= TDataBase.Create;
end;

procedure GlobalVarDestroy;
begin
  FreeAndNil(Images);
  FreeAndNil(DataBase);
  FreeAndNil(GridFont);
end;

procedure GlobalVarInit;
begin
  GridFont.Name:= FontLikeToName(flTimes{flArial});
  GridFont.Size:= 10;
end;

initialization

  GlobalVarCreate;

finalization

  GlobalVarDestroy;

end.

