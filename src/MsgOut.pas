unit MsgOut;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls;

type
  TMessageOutput = class(TForm)
    memOutput: TMemo;
    pnBot: TPanel;
    btnOK: TBitBtn;
    btnSaveFile: TButton;
    SaveDialog1: TSaveDialog;
    procedure FormCreate(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnSaveFileClick(Sender: TObject);
    //procedure ChangeLanguage;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MessageOutput: TMessageOutput;

implementation

uses
  Main;


{$R *.dfm}

procedure TMessageOutput.FormCreate(Sender: TObject);
begin
  Self.Caption := 'Output Message';
  memOutput.Text := '';
end;

procedure TMessageOutput.btnOKClick(Sender: TObject);
begin
  close;
end;

{procedure TMessageOutput.ChangeLanguage;
begin
  Self.Caption := sMsgOut_Caption;
  btnOK.Caption := sBtnOK;
end; }


procedure TMessageOutput.btnSaveFileClick(Sender: TObject);
//var
 // sPath, sPathFilename: string;
begin
 { sPath := ExtractFileDir(frmMain.ZipFName.Caption);
  if DirectoryExists(sPath) then
      SaveDialog1.InitialDir := sPath;

  //SaveDialog1.Filter := 'Text File(*.txt)|*.txt';

  if SaveDialog1.Execute then begin
      if SaveDialog1.FileName <> '' then begin
          sPathFilename := SaveDialog1.FileName;
        //  if Lowercase(ExtractFileExt(sPathFilename)) <> '.txt' then
         //     sPathFilename := sPathFilename + '.txt';

          memOutput.Lines.SaveToFile(sPathFilename);
      end;
  end; }
end;

end.
