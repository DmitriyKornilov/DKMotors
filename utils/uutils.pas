unit UUtils;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Controls, DK_CtrlUtils;

 procedure SetToolPanels(const AControls: array of TControl);
 procedure SetFlatToolPanels(const AControls: array of TControl);
 procedure SetToolButtons(const AControls: array of TControl);

implementation

procedure SetToolPanels(const AControls: array of TControl);
var
  i: Integer;
begin
  for i:= 0 to High(AControls) do
    ControlHeight(AControls[i], TOOL_PANEL_HEIGHT_DEFAULT);
end;

procedure SetFlatToolPanels(const AControls: array of TControl);
var
  i: Integer;
begin
  for i:= 0 to High(AControls) do
    ControlHeight(AControls[i], TOOL_PANEL_HEIGHT_DEFAULT-4);
end;

procedure SetToolButtons(const AControls: array of TControl);
var
  i: Integer;
begin
  for i:= 0 to High(AControls) do
    ControlWidth(AControls[i], TOOL_BUTTON_WIDTH_DEFAULT);
end;

end.

