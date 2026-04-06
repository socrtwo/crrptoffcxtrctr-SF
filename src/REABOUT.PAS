unit reabout;

interface

uses Windows, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, SysUtils, jpeg, StrUtils, ShellApi, Dialogs;

type
  TAboutBox = class(TForm)
    OKButton: TButton;
    lblProgNameVer: TLabel;
    lblCpyRightAuthor: TLabel;
    imgCcy: TImage;
    imgMainIcon: TImage;
    gbRegister: TGroupBox;
    lblRegisteredBy: TLabel;
    GroupBox1: TGroupBox;
    lblNotice: TLabel;
    lblWare: TLabel;
    lblRegisteredCompany: TLabel;
    btnRegisterNow: TButton;
    lblEmail: TLabel;
    lblHomePage: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure lblEmailClick(Sender: TObject);
    procedure lblHomePageClick(Sender: TObject);
  private
    { Private declarations }
    procedure ButtonRegisterNowClick(Sender: TObject);
  public
    { Public declarations }
  end;

//{$DEFINE IS_SHAREWARE} // Main, About Forms
  {$DEFINE SHAREWARE_FOR_S2SERVICES} // Main, reabout

var
  AboutBox: TAboutBox;

implementation

uses Main;



var
  bBombClose: boolean = True;

{$R *.dfm}
procedure TAboutBox.FormCreate(Sender: TObject);
{type
  TRegisteredUser = record

  sHeader: string[17];
  sRegUser: string[255];
end; }

//const cFileNameIs = 'reginfo.ini';
var
  MS: TMemoryStatus;
  MyRegister: TRegisteredUser;
  fReg: file of TRegisteredUser; //TLocation;

  sCopyrightAuthor, sProgNameVer, sContactEmail, sHomePage: string;
  sRegisteredCompanayName, sRegisteredUser, sUnRegistered: string;
begin
  beep;
  if lblWare.Parent = nil then begin // bomb!
  end;
  {$IFDEF IS_SHAREWARE}
  lblWare.Caption := 'This is shareware. This program is designed for www.s2services.com';
  {$ELSE}
  lblWare.Caption := 'This is freeware. This program is designed for www.s2services.com';
  {$ENDIF}
  // 'For personal, school or non-commercial use only.';

  if lblREgisteredCompany.Parent = nil then begin // bomb!
  end;
  lblREgisteredCompany.Caption := '';

  if lblRegisteredBy.Parent = nil then begin // bomb!
  end;
  lblRegisteredBy.Caption := 'This Copy is licensed to :' + #13#10;

  if lblNotice.Parent = nil then begin // bomb!
  end;

  {$IFDEF IS_SHAREWARE}
  lblNotice.Caption := 'Any questions, try below e-mail address.';
  {$ELSE}
  lblNotice.Caption := 'Any questions, try below e-mail address.';
  {$ENDIF}

  GlobalMemoryStatus(MS);
  //PhysMem.Caption := FormatFloat('#,###" KB"', MS.dwTotalPhys / 1024);
  //FreeRes.Caption := Format('%d %%', [MS.dwMemoryLoad]);

  FileMode := 2;  // read and write
  AssignFile(fReg, XsAppPath + XcIniRegFile);
  try
    Reset(fReg);  // open an old file and access read or write, error occurred when file doesn't exist
    //Rewrite(f);  // open a new file and access read or write
    //BlockWrite(f,buf,cnt,cnt);
    try
      Seek(fReg,0);
      Read(fReg,MyRegister);
      if MyRegister.sHeader = '[RegisteredBy]' then
          lblRegisteredBy.Caption := lblRegisteredBy.Caption +
          MyRegister.sRegUser;
          
    except

    end;

    CloseFile(fReg);
  except

  end;

  sProgNameVer := 'Corrupt Office 2007 Extractor v3.4.1';
  sCopyrightAuthor := 'Copyright (c) 2010 Ccy JC_HK';
  sContactEmail := 'contact@ccyjchk.com';
  sHomePage := 'http://www.ccyjchk.com/';
  sRegisteredUser := 'Registered User:';
  sUnRegistered := 'Unregistered';

 // Self.Caption := sProgNameVer + ' - About';
  Self.Caption := 'About';
  imgMainIcon.Picture.Icon := Application.Icon;

  if lblProgNameVer.Parent = nil then begin // bomb!
  end;
  lblProgNameVer.Caption := sProgNameVer;

  if lblCpyRightAuthor.Parent = nil then begin // bomb!
  end;
  lblCpyRightAuthor.Caption := sCopyrightAuthor;

  if lblEmail.Parent = nil then begin // bomb!
  end;
  lblEmail.Caption := sContactEmail;

  if lblHomePage.Parent = nil then begin // bomb!
  end;
  lblHomePage.Caption := sHomePage;
  
  {$IFDEF IS_SHAREWARE}
  if XbRegisteredUser then begin
      sRegisteredCompanayName := frmMain.ReadRegLocalMachine(XcSubKeyIs, XcRegCompanyName);
      lblRegisteredCompany.Caption := sRegisteredCompanayName;
      lblRegisteredBy.Font.Color := clBlack;
      lblRegisteredBy.Caption := sRegisteredUser + ' ' + frmMain.ReadRegLocalMachine(XcSubKeyIs, XcRegUserName);
  end
  else begin
      lblREgisteredCompany.Caption := '';
      lblRegisteredBy.Font.Color := clRed;
      lblRegisteredBy.Caption := sUnRegistered;
  end;
  btnRegisterNow.Enabled := not XbRegisteredUser;
  btnRegisterNow.Visible := not XbRegisteredUser;
  btnRegisterNow.OnClick := ButtonRegisterNowClick;
  {$ENDIF}


  {$IFDEF SHAREWARE_FOR_S2SERVICES}
  lblWare.Caption := 'This is shareware. This program is designed for www.s2services.com';
  lblRegisteredBy.Caption := sRegisteredUser + ' ' + frmMain.ReadRegLocalMachine(XcSubKeyIs, XcRegUserName);
  {$ENDIF}

  bBombClose := False; // is normal execution, this boolean can check outside cracking of TLabel
end;

procedure TAboutBox.OKButtonClick(Sender: TObject);
begin
  Close;
  //Release;
end;

procedure TAboutBox.FormActivate(Sender: TObject);
begin
  if bBombClose then // bomb!
      Close;
end;

procedure TAboutBox.ButtonRegisterNowClick(Sender: TObject);
var
  bRegistered: boolean;
begin
  {$IFDEF IS_SHAREWARE}
  bRegistered := False;
  with TOnRe.Create(Self) do begin
    try
      ShowModal;
      bRegistered := bSuccessfulRegistration;
    finally
      free;
    end;
  end;

  if bRegistered then
      Close;
  {$ENDIF}
end;

procedure TAboutBox.lblEmailClick(Sender: TObject);
var
  RetValue: integer;
  sEmail: string;
begin
  sEmail := 'contact@ccyjchk.com';
  if sEmail <> '' then begin
      sEmail := 'mailto:' + sEmail;
      RetValue := ShellExecute(handle, 'open', PChar(sEmail),
      '', '', SW_SHOW); // if success, return different value where depends on program

      if RetValue = 31 then begin // Has not the filename association?
          beep;
          MessageDlg('Failing to open e-mail client application.', mtError, [mbOK], 0);
      end;
  end
  else begin
    beep;
    MessageDlg('Invalid e-mail address.', mtInformation, [mbOK], 0);
  end;
end;

procedure TAboutBox.lblHomePageClick(Sender: TObject);
var
  RetValue: integer;
  sHomePage: string;
begin
  sHomePage := 'http://www.ccyjchk.com/';
  if sHomePage <> '' then begin
      RetValue := ShellExecute(handle, 'open', PChar(sHomePage),
      '', '', SW_SHOW); // if success, return different value where depends on program

      if RetValue = 31 then begin // Has not the filename association?
          beep;
          MessageDlg('Failing to open browser and connect to internet.', mtError, [mbOK], 0);
      end;
  end
  else begin
    beep;
    MessageDlg('Invalid URL.', mtInformation, [mbOK], 0);
  end;
end;

end.
