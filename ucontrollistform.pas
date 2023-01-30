unit UControlListForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  EditBtn, StdCtrls, rxctrls, VirtualTrees, DividerBevel, fpspreadsheetgrid;

type

  { TControlListForm }

  TControlListForm = class(TForm)
    AddButton: TSpeedButton;
    DelButton: TSpeedButton;
    DividerBevel5: TDividerBevel;
    DividerBevel6: TDividerBevel;
    DividerBevel7: TDividerBevel;
    EditButton: TSpeedButton;
    ExportButton: TRxSpeedButton;
    InfoGrid: TsWorksheetGrid;
    Label2: TLabel;
    MoreInfoCheckBox: TCheckBox;
    MotorNumEdit: TEditButton;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Splitter2: TSplitter;
    TopToolsPanel: TPanel;
    VT1: TVirtualStringTree;
  private

  public
    procedure ShowControlList;
  end;

var
  ControlListForm: TControlListForm;

implementation

{$R *.lfm}

{ TControlListForm }

procedure TControlListForm.ShowControlList;
begin

end;

end.

