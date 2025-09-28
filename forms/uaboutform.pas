unit UAboutForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, Buttons, LCLIntf,
  //DK packages utils
  DK_CtrlUtils;

const
  URL_PREFIX = 'https://';
  MAIL_PREFIX = 'mailto:';

  MAIN_CAPTION = 'DKMotors v1.2';

  PROJECT_NOTE = 'Учет производства, поставок и рекламационных случаев электродвигателей';
  PROJECT_URL  = 'www.github.com/DmitriyKornilov/DKMotors';

  AUTOR_NAME   = 'Дмитрий Корнилов';
  AUTOR_MAIL   = 'dakor2017@yandex.ru';

  COMPILER_VER = 'Free Pascal Compiler v3.2.2';
  COMPILER_URL = 'www.freepascal.org';
  LAZARUS_VER  = 'Lazarus IDE v4.2';
  LAZARUS_URL  = 'www.lazarus-ide.org';

  PNG_INFO     = 'PNG-файлы получены с ' +
                 'www.flaticon.com' +
                 ':';
  PNG_URL      = 'www.github.com/DmitriyKornilov/DKMotors/blob/main/png/Icons_info.pdf';

  PNG_URL_ADD1 = 'www.github.com/DmitriyKornilov/DKUIUtils/blob/main/source/png/icons_info.txt';
  PNG_URL_ADD2 = 'www.github.com/DmitriyKornilov/DKDatabase/blob/main/source/png/icons_info.txt';
  PNG_URL_ADD3 = 'www.github.com/DmitriyKornilov/DKDialogs/blob/main/source/png/icons_info.txt';
  PNG_URL_ADD4 = '';

type

  { TAboutForm }

  TAboutForm = class(TForm)
    AutorMailBeginLabel: TLabel;
    CompilerURLBeginLabel: TLabel;
    LazarusURLEndLabel: TLabel;
    LazarusURLLabel: TLabel;
    LazarusURLBeginLabel: TLabel;
    CompilerURLEndLabel: TLabel;
    CompilerURLLabel: TLabel;
    CompilerVerLabel: TLabel;
    LazarusVerLabel: TLabel;
    PNGURLAdd2Label: TLabel;
    PNGURLAdd3Label: TLabel;
    PNGURLAdd4Label: TLabel;
    PNGURLLabel: TLabel;
    PNGURLAdd1Label: TLabel;
    UsedLabel: TLabel;
    AutorNameLabel: TLabel;
    AutorMailEndLabel: TLabel;
    LazarusImage: TImage;
    AppNameLabel: TLabel;
    AutorLabel: TLabel;
    ProjectURLLabel: TLabel;
    ProjectNoteLabel: TLabel;
    ProjectLabel: TLabel;
    AutorMailLabel: TLabel;
    PNGInfoLabel: TLabel;
    procedure FormShow(Sender: TObject);
  private
    procedure DoMailTo(Sender: TObject);
    procedure DoOpenURL(Sender: TObject);
  public

  end;

var
  AboutForm: TAboutForm;

implementation

{$R *.lfm}

{ TAboutForm }

procedure TAboutForm.DoMailTo(Sender: TObject);
begin
  OpenURL(MAIL_PREFIX + (Sender as TLabel).Caption);
end;

procedure TAboutForm.DoOpenURL(Sender: TObject);
begin
  OpenURL(URL_PREFIX + (Sender as TLabel).Caption);
end;

procedure TAboutForm.FormShow(Sender: TObject);

  procedure SetLabel(const ALabel: TLabel;
                     const ACaption: String;
                     const AOnClick: TNotifyEvent = nil);
  begin
    ALabel.Caption:= ACaption;
    if Assigned(AOnClick) then
      ALabel.OnClick:= AOnClick;
  end;

begin
  SetLabel(AppNameLabel, MAIN_CAPTION);

  SetLabel(ProjectNoteLabel, PROJECT_NOTE);
  SetLabel(ProjectURLLabel, PROJECT_URL, @DoOpenURL);

  SetLabel(AutorNameLabel, AUTOR_NAME);
  SetLabel(AutorMailLabel, AUTOR_MAIL, @DoMailTo);

  SetLabel(CompilerVerLabel, COMPILER_VER);
  SetLabel(CompilerURLLabel, COMPILER_URL, @DoOpenURL);

  SetLabel(LazarusVerLabel, LAZARUS_VER);
  SetLabel(LazarusURLLabel, LAZARUS_URL, @DoOpenURL);

  SetLabel(PNGInfoLabel, PNG_INFO);
  SetLabel(PNGURLLabel, PNG_URL, @DoOpenURL);
  SetLabel(PNGURLAdd1Label, PNG_URL_ADD1, @DoOpenURL);
  SetLabel(PNGURLAdd2Label, PNG_URL_ADD2, @DoOpenURL);
  SetLabel(PNGURLAdd3Label, PNG_URL_ADD3, @DoOpenURL);
  SetLabel(PNGURLAdd4Label, PNG_URL_ADD4, @DoOpenURL);

  FormKeepMinSize(Self);
end;

end.

