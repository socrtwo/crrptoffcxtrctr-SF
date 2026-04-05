unit settings;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShlObj, Registry, ComCtrls, Spin, Buttons, ExtCtrls,
  VirtualTrees;

type
  TfrmSettings = class(TForm)
    btnOK: TButton;
    PageControl1: TPageControl;
    tabMisc: TTabSheet;
    GroupBox1: TGroupBox;
    chkContextMenu: TCheckBox;
    chkAssociate: TCheckBox;
    cbLanguageFiles: TComboBox;
    FontDialog1: TFontDialog;
    tabMainList: TTabSheet;
    GroupBox5: TGroupBox;
    chkBrowseList: TCheckBox;
    GroupBox2: TGroupBox;
    rbLast: TRadioButton;
    rbNormal: TRadioButton;
    GroupBox6: TGroupBox;
    rbWinTemp: TRadioButton;
    rbCurrent: TRadioButton;
    rbTarget: TRadioButton;
    tabView: TTabSheet;
    GroupBox7: TGroupBox;
    chkName: TCheckBox;
    chkModified: TCheckBox;
    chkSize: TCheckBox;
    chkRatio: TCheckBox;
    chkPacked: TCheckBox;
    chkPath: TCheckBox;
    chkType: TCheckBox;
    chkCRC: TCheckBox;
    chkAttributes: TCheckBox;
    GroupBox4: TGroupBox;
    chkAutoSize: TCheckBox;
    GroupBox3: TGroupBox;
    edFontName: TEdit;
    btnFonts: TBitBtn;
    edFontSize: TEdit;
    GroupBox8: TGroupBox;
    chkGridLines: TCheckBox;
    chkRowSelect: TCheckBox;
    chkHotTrack: TCheckBox;
    chkFileToRecycleBin: TCheckBox;
    tabFolders: TTabSheet;
    GroupBox9: TGroupBox;
    rbLastOpen: TRadioButton;
    rbFavoriteOpen: TRadioButton;
    edFavoriteOpen: TEdit;
    btnFavoriteOpen: TBitBtn;
    GroupBox10: TGroupBox;
    rbLastAdd: TRadioButton;
    rbFavoriteAdd: TRadioButton;
    edFavoriteAdd: TEdit;
    btnFavoriteAdd: TBitBtn;
    GroupBox11: TGroupBox;
    rbLastExtract: TRadioButton;
    rbFavoriteExtract: TRadioButton;
    edFavoriteExtract: TEdit;
    btnFavoriteExtract: TBitBtn;
    chkTxtWithNotePad: TCheckBox;
    chkHiddenFiles: TCheckBox;
    chkSystemFiles: TCheckBox;
    tabProgLoc: TTabSheet;
    gbVirusScanner: TGroupBox;
    edVirusScanner: TEdit;
    btnVirusScanner: TBitBtn;
    lblScanner: TLabel;
    edPara: TEdit;
    lblPara: TLabel;
    OpenDialog1: TOpenDialog;
    pnBottom: TPanel;
    btnSettingsApply: TBitBtn;
    btnSettingsCancel: TBitBtn;
    btnSettingsOK: TBitBtn;
    rbFilenamePara: TRadioButton;
    rbParaFilename: TRadioButton;
    lblAsterisk: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure chkContextMenuMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnOKClick(Sender: TObject);
    procedure chkAssociateMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    function DeleteKeyOrValueInRegistry(sRootKeyIs: Cardinal; sSubKey: string;
      sValueName: string): boolean;
    procedure btnFontsClick(Sender: TObject);
    procedure chkContextMenuKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure chkAssociateKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure chkTypeClick(Sender: TObject);
    procedure chkCRCClick(Sender: TObject);
    procedure chkAttributesClick(Sender: TObject);
    procedure btnFavoriteOpenClick(Sender: TObject);
    procedure btnFavoriteAddClick(Sender: TObject);
    procedure btnFavoriteExtractClick(Sender: TObject);
    procedure btnVirusScannerClick(Sender: TObject);
    procedure btnSettingsApplyClick(Sender: TObject);
    procedure btnSettingsCancelClick(Sender: TObject);
    procedure btnSettingsOKClick(Sender: TObject);
    procedure chkAutoSizeClick(Sender: TObject);
    //procedure ResizeWidthOfMainListColumns;
    //procedure cbLanguageFilesClick(Sender: TObject);
  private
    { Private declarations }
    //procedure LoadLanguageFilesToMnu;
    //procedure ChangeLanguage;
    procedure SaveStatusOfThisForm;
    procedure SavingStatusOfThisForm(bSaveAll: boolean);
    procedure RestoreStatusOfThisForm;
    procedure AreArchivesAssociatedWithHaHaZip;
    procedure OnOffchkContextMenu;
    procedure OnOffchkAssociate;
    //procedure RegisterDLLorOCX(sPara: string; sFile:string);
  public
    { Public declarations }
  end;

var
  frmSettings: TfrmSettings;

implementation

uses
  Main, MsgOut, filectrl;

const
  cCheckContextMenu = 'CheckContextMenu';
  cRadioBtnMainScreen = 'RadioBtnMainScreen';
  cMainListFontName = 'MainListFontName';
  cMainListFontSize = 'MainListFontSize';
  cCheckAutoSize = 'CheckAutoSize';

  cCheckBrowseList = 'CheckBrowseList';
  cRadioBtnZipTemp = 'RadioBtnZipTemp';
  cCheckType = 'CheckType';
  cCheckCRC = 'CheckCRC';
  cCheckAttributes = 'CheckAttributes';
  cCheckGridLines = 'CheckGridLines';
  cCheckRowSelect = 'CheckRowSelect';
  cCheckHotTrack = 'CheckHotTrack';
  cCheckFileToRecycleBin = 'CheckFileToRecycleBin';
  cCheckFavoriteOpen = 'CheckFavoriteOpen';
  cCheckFavoriteAdd = 'CheckFavoriteAdd';
  cCheckFavoriteExtract = 'CheckFavoriteExtract';
  cEditFavoriteOpen = 'EditFavoriteOpen';
  cEditFavoriteAdd = 'EditFavoriteAdd';
  cEditFavoriteExtract = 'EditFavoriteExtract';
  cCheckTxtWithNotePad = 'CheckTxtWithNotePad';
  cCheckHiddenFiles = 'CheckHiddenFiles';
  cCheckSystemFiles = 'CheckSystemFiles';
  cCheckParaFilename = 'CheckParaFilename';
  cEditVirusScanner = 'EditVirusScanner';
  cEditPara = 'EditPara';
var
  AFont: TFont;
  bChange_chkAutoSize: boolean = False;
  bChange_chkColumns: boolean = False;
{$R *.dfm}

procedure TfrmSettings.FormCreate(Sender: TObject);
{var
  MyStatus: TLocation;
  f: file of TLocation; //TLocation;
  sFileNameIs: string; }
begin
  Self.Icon := Application.Icon;
  cbLanguageFiles.Text := '';
  PageControl1.ActivePageIndex := 0;

  if Screen.PixelsPerInch <= 96 then begin
      Self.Width := 417;
      Self.Height := 359;
  end;
  //LoadLanguageFilesToMnu;

  cbLanguageFiles.Items.Append('English(Preset)');

 { FileMode := 2;  // read and write
  sFileNameIs := 'ccyZip.ini';

  // open initial file
  AssignFile(f, XsAppPath + sFileNameIs);
  try
    Reset(f);  // open an old file and access read or write, error occurred when file doesn't exist
    //Rewrite(f);  // open a new file and access read or write
    //BlockWrite(f,buf,cnt,cnt);

    try
      Seek(f,9);
      Read(f,MyStatus);
      if MyStatus.InputH = '[ContextMenuOn]' then begin
          if MyStatus.SongName = 'True' then
              chkContextMenu.Checked := True
          else
              chkContextMenu.Checked := False;

      end;
    except

    end;
    CloseFile(f);
  except

  end; }

  AreArchivesAssociatedWithHaHaZip;
  RestoreStatusOfThisForm;
 { if XbNewLanguage then begin
      ChangeLanguage;
      cbLanguageFiles.Text := XsLanguageFilename;
  end; }
end;

procedure TfrmSettings.chkContextMenuMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  OnOffchkContextMenu;
end;

procedure TfrmSettings.btnOKClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSettings.chkAssociateMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  OnOffchkAssociate;
end;

function TfrmSettings.DeleteKeyOrValueInRegistry(sRootKeyIs: Cardinal;
 sSubKey: string; sValueName: string): boolean;
// if sValueName = '' means delete a key, otherwise delete a value name
var
  reg: TRegistry;
  Val: TStringList;
  i: integer;
  sCompValue: string;
begin
  Result := True;
  if sValueName = 'null' then
      sCompValue := ''
  else
      sCompValue := sValueName;

  if sSubKey <> '' then begin
      if copy(sSubKey, 1, 1) <> '\' then
          sSubKey := '\' + sSubKey;
          
  end;

  reg := TRegistry.Create;
  Val := TStringList.Create;
  try
    reg.RootKey := sRootKeyIs;
    if not reg.OpenKey(sSubKey, False) then

    else begin
        if sValueName = '' then begin
            if not reg.DeleteKey(sSubKey) then begin
                MessageDlg('Error found, failing to update registry.', mtError,
                [mbOK], 0);
                Result := False;
            end;
        end
        else begin
            reg.GetValueNames(Val);
            if Val.Count > 0 then begin
                for i := 0 to Val.Count-1 do begin
                  if Val.Strings[i] = sCompValue then begin
                      if not reg.DeleteValue(Val.Strings[i]) then begin
                          MessageDlg('Error found, failing to update registry.',
                          mtError, [mbOK], 0);
                          Result := False;
                      end;
                  end;
                end;
            end;
        end;

        reg.CloseKey;
    end;
  finally
    Val.Free;
    reg.Free;
  end;
end;

procedure TfrmSettings.SaveStatusOfThisForm;
var
  iRadio: shortint;
  sChecked: string;
  sZipTemp: string;
begin
  if rbLast.Checked then
      iRadio := rbLast.Tag
  else if rbNormal.Checked then
      iRadio := rbNormal.Tag;

  if rbWinTemp.Checked then begin
      sZipTemp := Main.XcZipTemp_WinTemp;
      frmMain.Set_ZipTempIs(sZipTemp); // update
  end
  else if rbCurrent.Checked then begin
      sZipTemp := Main.XcZipTemp_Current;
      frmMain.Set_ZipTempIs(sZipTemp);
  end
  else if rbTarget.Checked then begin
      sZipTemp := Main.XcZipTemp_Target;
      frmMain.Set_ZipTempIs(sZipTemp);
  end
  else
      sZipTemp := '';

  frmMain.WriteIniToReg(XcSubKeyIs, cRadioBtnMainScreen, inttostr(iRadio) );
  frmMain.WriteIniToReg(XcSubKeyIs, cMainListFontName, edFontName.Text);
  frmMain.WriteIniToReg(XcSubKeyIs, cMainListFontSize, edFontSize.Text);
  frmMain.WriteIniToReg(XcSubKeyIs, cRadioBtnZipTemp, sZipTemp);

  if AFont <> nil then
      frmMain.ListView1.Font.Assign(AFont);

  if chkAutoSize.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckAutoSize, sChecked);
end;

procedure TfrmSettings.RestoreStatusOfThisForm;
var
  sCheckContextMenu: string;
  iRadioMainPos: shortint;
  sRadioMainPos: string;
  sFontName, sFontSize: string;
  sCheckAutoSize: string;

  sCheckBrowseList, sZipTemp: string;
  sCheckType, sCheckCRC, sCheckAttributes, sCheckGridLines, sCheckRowSelect: string;
  sCheckHotTrack, sCheckFileToRecycleBin: string;
  sCheckFavoriteOpen, sCheckFavoriteAdd, sCheckFavoriteExtract: string;
  sCheckTxtWithNotePad, sCheckHiddenFiles, sCheckSystemFiles: string;
  sEditVirusScanner, sEditPara: string;
begin
  sCheckContextMenu := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckContextMenu);
  if sCheckContextMenu = '1' then
      chkContextMenu.Checked := True
  else
      chkContextMenu.Checked := False;

  sRadioMainPos := frmMain.ReadIniFromReg(XcSubKeyIs, cRadioBtnMainScreen);
  if sRadioMainPos <> '' then begin
      iRadioMainPos := strtoint(sRadioMainPos);
      if rbLast.Tag = iRadioMainPos then
          rbLast.Checked := True
      else if rbNormal.Tag = iRadioMainPos then
          rbNormal.Checked := True;
  end;

  sZipTemp := frmMain.ReadIniFromReg(XcSubKeyIs, cRadioBtnZipTemp);
  if sZipTemp = Main.XcZipTemp_WinTemp then
      rbWinTemp.Checked := True
  else if sZipTemp = Main.XcZipTemp_Current then
      rbCurrent.Checked := True
  else if sZipTemp = Main.XcZipTemp_Target then
      rbTarget.Checked := True;

  // ***
  sFontName := frmMain.ReadIniFromReg(XcSubKeyIs, cMainListFontName);
  if sFontName <> '' then
      edFontName.Text := sFontName;

  sFontSize := frmMain.ReadIniFromReg(XcSubKeyIs, cMainListFontSize);
  if sFontSize <> '' then
      edFontSize.Text := sFontSize;
  // ***
  sCheckAutoSize := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckAutoSize);
  chkAutoSize.Checked := (sCheckAutoSize = '1');

  sCheckBrowseList := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckBrowseList);
  if sCheckBrowseList = '0' then
      chkBrowseList.Checked := False
  else
      chkBrowseList.Checked := True;

  sCheckType := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckType);
  chkType.Checked := (sCheckType = '1');

  sCheckCRC := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckCRC);
  chkCRC.Checked := (sCheckCRC = '1');

  sCheckAttributes := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckAttributes);
  chkAttributes.Checked := (sCheckAttributes = '1');

  sCheckGridLines := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckGridLines);
  chkGridLines.Checked := (sCheckGridLines = '1');

  sCheckRowSelect := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckRowSelect);
  chkRowSelect.Checked := (sCheckRowSelect = '1');

  sCheckHotTrack := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckHotTrack);
  chkHotTrack.Checked := (sCheckHotTrack = '1');

  sCheckFileToRecycleBin := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckFileToRecycleBin);
  if sCheckFileToRecycleBin = '0' then
      chkFileToRecycleBin.Checked := False
  else
      chkFileToRecycleBin.Checked := True;

  sCheckFavoriteOpen := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckFavoriteOpen);
  rbFavoriteOpen.Checked := (sCheckFavoriteOpen = '1');
  sCheckFavoriteAdd := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckFavoriteAdd);
  rbFavoriteAdd.Checked := (sCheckFavoriteAdd = '1');
  sCheckFavoriteExtract := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckFavoriteExtract);
  rbFavoriteExtract.Checked := (sCheckFavoriteExtract = '1');
  edFavoriteOpen.Text := frmMain.ReadIniFromReg(XcSubKeyIs, cEditFavoriteOpen);
  edFavoriteAdd.Text := frmMain.ReadIniFromReg(XcSubKeyIs, cEditFavoriteAdd);
  edFavoriteExtract.Text := frmMain.ReadIniFromReg(XcSubKeyIs, cEditFavoriteExtract);

  sCheckTxtWithNotePad := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckTxtWithNotePad);
  if sCheckTxtWithNotePad = '0' then
      chkTxtWithNotePad.Checked := False
  else
      chkTxtWithNotePad.Checked := True;

  sCheckHiddenFiles := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckHiddenFiles);
  chkHiddenFiles.Checked := (sCheckHiddenFiles = '1');

  sCheckSystemFiles := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckSystemFiles);
  chkSystemFiles.Checked := (sCheckSystemFiles = '1');

  if frmMain.ReadIniFromReg(XcSubKeyIs, cCheckParaFilename) = '1' then
      rbParaFilename.Checked := True
  else 
      rbFilenamePara.Checked := True;

  sEditVirusScanner := frmMain.ReadIniFromReg(XcSubKeyIs, cEditVirusScanner);
  if sEditVirusScanner <> '' then
      edVirusScanner.Text := sEditVirusScanner;

  sEditPara := frmMain.ReadIniFromReg(XcSubKeyIs, cEditPara);
  if sEditPara <> '' then
      edPara.Text := sEditPara;
      
end;

procedure TfrmSettings.AreArchivesAssociatedWithHaHaZip;
var
  reg: TRegistry;
  sValueName: string;
  bHaHaZip_Is_Master: boolean;
  bKeyExist_HaHaZip: boolean;
begin
  sValueName := '';

  reg := TRegistry.Create;
  try
    reg.RootKey := HKEY_CLASSES_ROOT;

    bHaHaZip_Is_Master := False;
    if reg.OpenKey('.zip', False) then begin
        sValueName := reg.ReadString('');
        if sValueName = XcProgramName then
            bHaHaZip_Is_Master := True;
            
        reg.CloseKey;
    end;

    bKeyExist_HaHaZip := False;
    if reg.OpenKey(XcProgramName, False) then begin
        if reg.ReadString('') = 'Zip' then
            bKeyExist_HaHaZip := True;

        reg.CloseKey;
    end;

    if bHaHaZip_Is_Master and bKeyExist_HaHaZip then
        chkAssociate.Checked := True
    else
        chkAssociate.Checked := False;

  finally
    reg.free;
  end;
end;

// ********************************************************************* //
{procedure TfrmSettings.LoadLanguageFilesToMnu;
var
  sr: TSearchRec;
  iFileAttrs: Integer;
begin
  if not DirectoryExists(XsAppPath) then begin
      //sErrorMessage := 'Directory does not exist : ' + sPathIs;
      exit;
  end;
  iFileAttrs :=  faArchive; //$0000003F; //faAnyFile  faArchive;   //archives only
  if FindFirst(XsAppPath + '*.*', iFileAttrs, sr) = 0 then begin
      repeat
        if ((sr.Attr and iFileAttrs) = sr.Attr) and
        (sr.Name <> '.') and (sr.Name <> '..') then begin
            if AnsiLowercase(ExtractFileExt(sr.Name)) = '.lng' then begin
                cbLanguageFiles.Items.Append(sr.Name);
            end
        end;
      until FindNext(sr) <> 0;
      FindClose(sr);
  end;
end; }

{procedure TfrmSettings.cbLanguageFilesClick(Sender: TObject);
var
  iChangLang: shortint;
begin
  if cbLanguageFiles.Text = '' then begin
      beep;
      MessageDlg('Null string found', mtError, [mbOK], 0);
      exit;
  end;
  
  if cbLanguageFiles.Text <> 'English(Preset)' then begin
      if frmMain.DecryptLanguageFile(cbLanguageFiles.Text) then begin
          ChangeLanguage;
          frmMain.MainChangeLanguage;
          //GetDirectory.FolderChangeLanguage;
          XbNewLanguage := True;
      end
      else
          exit;
          
  end
  else if cbLanguageFiles.Text = 'English(Preset)' then 
      MessageDlg('It will take effect when HaHaZip starts on the next time.',
      mtInformation, [mbOK], 0);

  if cbLanguageFiles.Text = 'English(Preset)' then
      iChangLang := 0 // False
  else if cbLanguageFiles.Text <> '' then
      iChangLang := 1; // True

  if length(cbLanguageFiles.Text) <= 17 then
      frmMain.SaveSettings(nil, '[UseLanguage]', iChangLang, 0,
      cbLanguageFiles.Text, 10)
  else begin
      beep;
      MessageDlg('Error found, text of chosen item (' + cbLanguageFiles.Text +
      ') exceeds seventeen ' +
      'characters(bytes). This setting will not be saved to disk. Please alter ' +
      'the filename to shorter one and click on this item again.',
      mtError, [mbOK], 0);
  end;
end; }


{procedure TfrmSettings.ChangeLanguage;
begin
  Self.Caption := sSettings_Caption;
  chkContextMenu.Caption := sSettings_Show;
  chkAssociate.Caption := sSettings_Associate;
  lblLanguage.Caption := sSettings_Language + ' :';
  btnOK.Caption := sBtnOK;
end; }


procedure TfrmSettings.btnFontsClick(Sender: TObject);
begin
  if FontDialog1.Execute then begin
      AFont := FontDialog1.Font;
      edFontName.Text := AFont.Name;
      edFontSize.Text := inttostr(AFont.Size);
  end;
end;

procedure TfrmSettings.OnOffchkContextMenu;
const
  cRegSvr32Exe = 'regsvr32.exe';
  cUserToRegisterDll = 'UserToRegisterDll';
  cIsRegisteredDll= 'IsRegisteredDll';
var
  sPathRegSvrName: string;
  H: THandle;
  lpici: TCMInvokeCommandInfo;
  sCheck, sPara, sAppPathFilename: string;
  sCurUserName, sWho, sDll_IsRegisterd, sPath: string;
begin
  sPathRegSvrName := sSystemDirIs + cRegSvr32Exe;
  if AnsiPos(' ', sPathRegSvrName) <> 0 then
      sPathRegSvrName := ExtractShortPathName(sPathRegSvrName);

  if not FileExists(sPathRegSvrName) then begin
      beep;
      Raise Exception.Create('Error found, failing to find ' + cRegSvr32Exe);
      chkContextMenu.Checked := not chkContextMenu.Checked;
      exit;
  end;

  if not FileExists(XsAppPath + XcExeProgFilename) then begin
      beep;
      Raise Exception.Create('Error found, failing to find ' + XcExeProgFilename
      );
      chkContextMenu.Checked := not chkContextMenu.Checked;
      exit;
  end;

  with TRegistry.Create do begin // read App path from HKEY Local Machine
    try
      RootKey := HKEY_LOCAL_MACHINE;
      if OpenKey(XcSubKeyIs, False) then begin
          sPath := ReadString('RootDir');
          CloseKey;
      end;
    finally
      free;
    end;
  end;
  // If found path from HKEY Local Machine, use it
  if (sPath <> '') and FileExists(sPath + XcContextMenuProgName) then
      sAppPathFilename := sPath + XcContextMenuProgName // Other user is to register this file somewhere
  else // use App path
      sAppPathFilename := XsAppPath + XcContextMenuProgName;

  sAppPathFilename := ExtractShortPathName(sAppPathFilename);
  sCurUserName := AnsiLowercase(frmMain.UserID);
  if sCurUserName = '' then
      sCurUserName := AnsiLowercase(XcDefault);
      
  if chkContextMenu.Checked then begin
      with TRegistry.Create do // write path to HKey Current User
      try
        RootKey := HKEY_CURRENT_USER;
        if OpenKey(XcSubKeyIs, True) then begin
            WriteString('RootDir', XsAppPath);
            CloseKey;
        end;

        sWho := frmMain.ReadRegLocalMachine(XcSubKeyIs, XcInstalledBy);
        sDll_IsRegisterd := frmMain.ReadRegLocalMachine(XcSubKeyIs, cIsRegisteredDll);
        // Is User first time to register Dll? 
        if (sCurUserName = AnsiLowercase(sWho)) and (sDll_IsRegisterd = '') then begin
            RootKey := HKEY_LOCAL_MACHINE;
            if OpenKey(XcSubKeyIs, True) then begin
                WriteString('RootDir', XsAppPath);
                CloseKey;
            end;
        end
      finally
        Free;
      end;

      if (sCurUserName <> AnsiLowercase(sWho)) and ( (sDll_IsRegisterd = 'No') or
      (sDll_IsRegisterd = '') ) then begin
          beep;
          MessageDlg('Failing to enable context menu. Application installed ' +
          'by the User(' + sWho + ') who has first priority to change this ' +
          'option in the operating system. If the User already enabled ' +
          'this function and you do not have administrative priviledges, ' +
          'ignore this message. This function will be enabled on next ' +
          'logon.', mtError, [mbOK], 0);
          exit;
      end
      else begin
          //RegisterDLLorOCX('/s ', 'contmenu.dll');

          sPara := ' /s ' + sAppPathFilename;
          H := WinExec(PChar(sPathRegSvrName + sPara), lpici.nShow);
          if H < 32 then begin
              beep;
              MessageBox(lpici.hWnd, 'Error executing ' +
              cRegSvr32Exe + '.', 'Error', MB_ICONERROR or MB_OK);
              chkContextMenu.Checked := not chkContextMenu.Checked;
              exit
          end; 
          frmMain.WriteRegLocalMachine(XcSubKeyIs, cIsRegisteredDll, 'Yes');
      end;
  end
  else begin // want to unregister
      sWho := frmMain.ReadRegLocalMachine(XcSubKeyIs, XcInstalledBy);
      //sDll_IsRegisterd := frmMain.ReadRegLocalMachine(XcSubKeyIs, cIsRegisteredDll);
      if sCurUserName = AnsiLowercase(sWho) then begin
          sPara := ' /u /s ' + sAppPathFilename;
          H := WinExec(PChar(sPathRegSvrName + sPara), lpici.nShow);
          if H < 32 then begin
              beep;
              MessageBox(lpici.hWnd, 'Error executing ' +
              cRegSvr32Exe + '.', 'Error', MB_ICONERROR or MB_OK);
              chkContextMenu.Checked := not chkContextMenu.Checked;
              exit;
          end;
          frmMain.WriteRegLocalMachine(XcSubKeyIs, cIsRegisteredDll, 'No');
      end
      else begin
          beep;
          MessageDlg('Failing to disable context menu. Application installed ' +
          'by the User(' + sWho + ') who has first priority to change this ' +
          'option in the operating system.', mtError, [mbOK], 0);
          exit;
      end;
  end;

  if chkContextMenu.Checked then
      sCheck := '1'
  else
      sCheck := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckContextMenu, sCheck);
end;

procedure TfrmSettings.chkContextMenuKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 32 then begin
      chkContextMenu.Checked := not chkContextMenu.Checked;
      OnOffchkContextMenu;
      chkContextMenu.Checked := not chkContextMenu.Checked;
  end;
end;

procedure TfrmSettings.OnOffchkAssociate;
var
  sWho, sCurUserName: string;
  sRootK: Cardinal;
  sSubK, sSubKZip: string;
begin
  sCurUserName := frmMain.UserID;
  if sCurUserName = '' then
      sCurUserName := XcDefault;
      
  sWho := frmMain.ReadRegLocalMachine(XcSubKeyIs, XcInstalledBy);

  if ( bWinVista = False ) and ( AnsiLowercase(sCurUserName) <> AnsiLowercase(sWho) ) then begin
      beep;
      MessageDlg('Only the user(' + sWho + ') who firstly installed this ' +
      'application can change this option in the operating system.', mtError, [mbOK], 0);
      chkAssociate.Checked := not chkAssociate.Checked;
      exit;
  end;

  if chkAssociate.Checked then
      // below doing association
      frmMain.RegisterFileType('.zip', XcProgramName, 'Zip', XsAppPath +
      XcExeProgFilename, 0, True)
  else begin
      // Is deletion successful? expecially under NT, prevent duplicated error dialog
      // Super users can delete CLASSES_ROOT\SeeSeeWhy HaHaZip but can't
      // delete CLASSES_ROOT\.zip

      if bWinVista = True then begin
          sRootK := HKEY_CURRENT_USER;
          sSubK := '\Software\Classes\';
          sSubKZip := '\Software\Classes\.zip';
      end
      else begin
          sRootK := HKEY_CLASSES_ROOT;
          sSubK := '\';
          sSubKZip := '\.zip';
      end;

      if DeleteKeyOrValueInRegistry(sRootK, sSubK + XcProgramName, '') then begin
          DeleteKeyOrValueInRegistry(sRootK, sSubKZip, 'null');
          SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nil, nil);
      end;
  end;
end;

procedure TfrmSettings.chkAssociateKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 32 then begin
      chkAssociate.Checked := not chkAssociate.Checked;
      OnOffchkAssociate;
      chkAssociate.Checked := not chkAssociate.Checked;
  end;
end;

procedure TfrmSettings.chkTypeClick(Sender: TObject);
var
  sCheck: string;
begin
  bChange_chkColumns := True;
  frmMain.SetBool_CheckType(chkType.Checked);
  frmMain.SetBool_CntOf_ColumnChanged_OnMainLst(True);
  if chkType.Checked then
      sCheck := '1'
  else
      sCheck := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckType, sCheck);
end;

procedure TfrmSettings.chkCRCClick(Sender: TObject);
var
  sCheck: string;
begin
  bChange_chkColumns := True;
  frmMain.SetBool_CheckCRC(chkCRC.Checked);
  frmMain.SetBool_CntOf_ColumnChanged_OnMainLst(True);
  if chkCRC.Checked then
      sCheck := '1'
  else
      sCheck := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckCRC, sCheck);
end;

procedure TfrmSettings.chkAttributesClick(Sender: TObject);
var
  sCheck: string;
begin
  bChange_chkColumns := True;
  frmMain.SetBool_CheckAttributes(chkAttributes.Checked);
  frmMain.SetBool_CntOf_ColumnChanged_OnMainLst(True);
  if chkAttributes.Checked then
      sCheck := '1'
  else
      sCheck := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckAttributes, sCheck);
end;

procedure TfrmSettings.btnFavoriteOpenClick(Sender: TObject);
var
  sOutPath: string;
begin
  if not SelectDirectory('Select Folder...', '', sOutPath) then exit;
  edFavoriteOpen.Text := sOutPath;
end;

procedure TfrmSettings.btnFavoriteAddClick(Sender: TObject);
var
  sOutPath: string;
begin
  if not SelectDirectory('Select Folder...', '', sOutPath) then exit;
  edFavoriteAdd.Text := sOutPath;
end;

procedure TfrmSettings.btnFavoriteExtractClick(Sender: TObject);
var
  sOutPath: string;
begin
  if not SelectDirectory('Select Folder...', '', sOutPath) then exit;
  edFavoriteExtract.Text := sOutPath;
end;

procedure TfrmSettings.btnVirusScannerClick(Sender: TObject);
var
  sPathFName: string;
begin
  if OpenDialog1.Execute then begin
      sPathFName := OPenDialog1.FileName;
      edVirusScanner.Text := sPathFName;
  end;
end;

procedure TfrmSettings.btnSettingsApplyClick(Sender: TObject);
begin
  SavingStatusOfThisForm(False);
end;

procedure TfrmSettings.btnSettingsCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSettings.SavingStatusOfThisForm(bSaveAll: boolean);
var
  sCheck, sZipTemp: string;
  sCheckA, sCheckB, sCheckC: string;
  iRadio: shortint;
begin
  if tabMisc.Showing or bSaveAll then begin
      // ***
      frmMain.SetBool_CheckFileToRecycleBin(chkFileToRecycleBin.Checked);
      if chkFileToRecycleBin.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckFileToRecycleBin, sCheck);
      // ***
      frmMain.SetBool_CheckTxtWithNotePad(chkTxtWithNotePad.Checked);
      if chkTxtWithNotePad.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckTxtWithNotePad, sCheck);
      // ***
      if rbWinTemp.Checked then begin
          sZipTemp := Main.XcZipTemp_WinTemp;
          frmMain.Set_ZipTempIs(sZipTemp); // update
      end
      else if rbCurrent.Checked then begin
          sZipTemp := Main.XcZipTemp_Current;
          frmMain.Set_ZipTempIs(sZipTemp);
      end
      else if rbTarget.Checked then begin
          sZipTemp := Main.XcZipTemp_Target;
          frmMain.Set_ZipTempIs(sZipTemp);
      end
      else
          sZipTemp := '';

      frmMain.WriteIniToReg(XcSubKeyIs, cRadioBtnZipTemp, sZipTemp);
  end;

  if tabMainList.Showing or bSaveAll then begin
      if bChange_chkAutoSize then begin // Setting changed
          frmMain.SetStatusOfchkAutoSizeToVar(chkAutoSize.Checked);
          frmMain.ResizeWidthOfMainListColumns;
      end;

      if rbLast.Checked then
          iRadio := rbLast.Tag
      else if rbNormal.Checked then
          iRadio := rbNormal.Tag;

      frmMain.WriteIniToReg(XcSubKeyIs, cRadioBtnMainScreen, inttostr(iRadio) );
      frmMain.WriteIniToReg(XcSubKeyIs, cMainListFontName, edFontName.Text);
      frmMain.WriteIniToReg(XcSubKeyIs, cMainListFontSize, edFontSize.Text);
      if AFont <> nil then
          //frmMain.ListView1.Font.Assign(AFont);
          frmMain.VT.Font.Assign(AFont);

      if chkAutoSize.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckAutoSize, sCheck);

      if chkBrowseList.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckBrowseList, sCheck);

      if chkBrowseList.Checked then begin
          if not frmMain.PanMiddleLeft.Visible then begin
              frmMain.Splitter1.Visible := True; // must show this first
              frmMain.PanMiddleLeft.Visible := True;
          end;
      end
      else begin
          if frmMain.PanMiddleLeft.Visible then begin
              frmMain.Splitter1.Visible := False; // must show this first
              frmMain.PanMiddleLeft.Visible := False;
          end;
      end;
  end;

  if tabView.Showing or bSaveAll then begin
      if bChange_chkColumns then begin
          if frmMain.btnShowVirtual.Down then
              frmMain.TreeView1_Change
          else
              frmMain.Refresh_ItemsOnMainList;
              
      end;

      //frmMain.ListView1.GridLines := chkGridLines.Checked;
      if chkGridLines.Checked then begin
          frmMain.VT.TreeOptions.PaintOptions := frmMain.VT.TreeOptions.PaintOptions + [toShowHorzGridLines];
          frmMain.VT.TreeOptions.PaintOptions := frmMain.VT.TreeOptions.PaintOptions + [toShowVertGridLines];
      end
      else begin
          frmMain.VT.TreeOptions.PaintOptions := frmMain.VT.TreeOptions.PaintOptions - [toShowHorzGridLines];
          frmMain.VT.TreeOptions.PaintOptions := frmMain.VT.TreeOptions.PaintOptions - [toShowVertGridLines];
      end;

      if chkGridLines.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckGridLines, sCheck);

      //frmMain.ListView1.RowSelect := chkRowSelect.Checked;
      if chkRowSelect.Checked then
          frmMain.VT.TreeOptions.SelectionOptions := frmMain.VT.TreeOptions.SelectionOptions + [toFullRowSelect]
      else
          frmMain.VT.TreeOptions.SelectionOptions := frmMain.VT.TreeOptions.SelectionOptions - [toFullRowSelect];

      if chkRowSelect.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckRowSelect, sCheck);

      //frmMain.ListView1.HotTrack := chkHotTrack.Checked;
      if chkHotTrack.Checked then
          frmMain.VT.TreeOptions.PaintOptions := frmMain.VT.TreeOptions.PaintOptions + [toHotTrack]
      else
          frmMain.VT.TreeOptions.PaintOptions := frmMain.VT.TreeOptions.PaintOptions - [toHotTrack];

      if chkHotTrack.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckHotTrack, sCheck);

  end;

  if tabFolders.Showing or bSaveAll then begin
      frmMain.SetBool_CheckFavoriteOpen(rbFavoriteOpen.Checked);
      frmMain.SetBool_CheckFavoriteAdd(rbFavoriteAdd.Checked);
      frmMain.SetBool_CheckFavoriteExtract(rbFavoriteExtract.Checked);

      if rbFavoriteOpen.Checked then begin
          frmMain.SetStr_FavoriteOpen(edFavoriteOpen.Text);
          sCheckA := '1';
      end
      else
          sCheckA := '0';

      if rbFavoriteAdd.Checked then begin
          frmMain.SetStr_FavoriteAdd(edFavoriteAdd.Text);
          sCheckB := '1';
      end
      else
          sCheckB := '0';

      if rbFavoriteExtract.Checked then begin
          frmMain.SetStr_FavoriteExtract(edFavoriteExtract.Text);
          sCheckC := '1';
      end
      else
          sCheckC := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckFavoriteOpen, sCheckA);
      frmMain.WriteIniToReg(XcSubKeyIs, cCheckFavoriteAdd, sCheckB);
      frmMain.WriteIniToReg(XcSubKeyIs, cCheckFavoriteExtract, sCheckC);
      frmMain.WriteIniToReg(XcSubKeyIs, cEditFavoriteOpen, edFavoriteOpen.Text);
      frmMain.WriteIniToReg(XcSubKeyIs, cEditFavoriteAdd, edFavoriteAdd.Text);
      frmMain.WriteIniToReg(XcSubKeyIs, cEditFavoriteExtract, edFavoriteExtract.Text);

      frmMain.SetBool_CheckHiddenFiles(chkHiddenFiles.Checked);
      if chkHiddenFiles.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckHiddenFiles, sCheck);

      frmMain.SetBool_CheckSystemFiles(chkSystemFiles.Checked);
      if chkSystemFiles.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckSystemFiles, sCheck);

  end;

  if tabProgLoc.Showing or bSaveAll then begin
      frmMain.SetStr_VirusScannerProg(edVirusScanner.Text);
      frmMain.SetStr_VirusScannerPara(edPara.Text);
      frmMain.WriteIniToReg(XcSubKeyIs, cEditVirusScanner, edVirusScanner.Text);
      frmMain.WriteIniToReg(XcSubKeyIs, cEditPara, edPara.Text);

      frmMain.SetBool_RadioButtonParaFilename(rbParaFilename.Checked);
      if rbParaFilename.Checked then
          sCheck := '1'
      else
          sCheck := '0';

      frmMain.WriteIniToReg(XcSubKeyIs, cCheckParaFilename, sCheck);    
  end;
end;

procedure TfrmSettings.btnSettingsOKClick(Sender: TObject);
begin
  SavingStatusOfThisForm(True); // Save all changes
  Close;
end;

procedure TfrmSettings.chkAutoSizeClick(Sender: TObject);
begin
  bChange_chkAutoSize := True;
end;

{procedure TfrmSettings.RegisterDLLorOCX(sPara: string; sFile: string);
type
  TRegFunc = function : HResult; stdcall;
var
  ARegFunc : TRegFunc;
  aHandle  : THandle;
  ocxPath  : string;
begin
 try
  ocxPath := ExtractFilePath(Application.ExeName) + sFile;
  aHandle := LoadLibrary(PChar(ocxPath));
  if aHandle <> 0 then
  begin
    ARegFunc := GetProcAddress(aHandle,'DllRegisterServer');
    if Assigned(ARegFunc) then
    begin
      frmMain.ExecAndWait('regsvr32', sPara + '''''' + ocxPath + '''''');
    end;
    FreeLibrary(aHandle);
  end;
 except
  ShowMessage(Format('Unable to register %s', [ocxPath]));
 end;
end; }


end.
