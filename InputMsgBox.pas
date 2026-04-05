unit InputMsgBox;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons;

type
  TInputMsg = class(TForm)
    Memo1: TMemo;
    lblBuffer: TLabel;
    btnOK: TBitBtn;
    btnCancel: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  InputMsg: TInputMsg;

implementation

uses
  Main;
  
{$R *.dfm}

procedure TInputMsg.FormCreate(Sender: TObject);
begin
  Self.Caption := 'Add Comment';
  Memo1.Clear;
  lblBuffer.Caption := 'Input comment about this archive.'; 
  if Screen.PixelsPerInch <= 96 then begin
      Self.Width := 301;
      Self.Height := 276;
  end;
end;

procedure TInputMsg.btnOKClick(Sender: TObject);
begin
  Main.bCancelInputComment_OnInputMsgBox := False;
  Close;
end;

procedure TInputMsg.btnCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TInputMsg.FormActivate(Sender: TObject);
begin
  Main.bCancelInputComment_OnInputMsgBox := True;
end;

end.
