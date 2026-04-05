unit hrglass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Gauges;

type
  THourGlass = class(TForm)
    Animate1: TAnimate;
    lblOnProcess: TLabel;
    Gauge1: TGauge;
    btnCancel: TButton;
    Gauge2: TGauge;
    procedure FormCreate(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  HourGlass: THourGlass;

implementation

uses
  Main;
{$R *.dfm}

procedure THourGlass.FormCreate(Sender: TObject);
var
  idHandle : THANDLE;
  Menu : HMENU;
begin
  Self.Caption := XcProgramName;
  Self.Top := 0;
  Self.Left := 0;
  if FileExists(XsAppPath + 'Process.avi') then
      Animate1.FileName := XsAppPath + 'Process.avi';
  //Animate1.Active := True;
  idHandle := FindWindow(nil, PChar(Self.Caption));
  if (idHandle <> 0) then begin
    Menu := GetSystemMenu(idHandle, FALSE);
    if (Menu <> 0) then
    DeleteMenu(Menu, SC_CLOSE, MF_BYCOMMAND);
    //RemoveMenu(Menu, GetMenuItemCount(Menu) -1, MF_BYPOSITION Or MF_REMOVE);
  end;

  Self.Repaint;
end;

procedure THourGlass.btnCancelClick(Sender: TObject);
var
  sMsgForAdd: string;
begin
  sMsgForAdd := '';
  if self.Tag = 1 then // adding files
      sMsgForAdd := 'When adding files to archive, this will cause the ' +
      'target zip file corrupted.';

  if MessageDlg('Do you want to terminate the process? ' + sMsgForAdd,
  mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
      frmMain.ZipMaster1.Cancel := True;
      Close;
  end;
end;

end.
