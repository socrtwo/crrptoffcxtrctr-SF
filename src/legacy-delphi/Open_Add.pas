unit Open_Add;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IEComboBox, ExtCtrls, StdCtrls, ToolWin, ComCtrls, Buttons,
  IEListView, DirView, DriveView, IEDriveInfo, FileCtrl;

type
  TfrmDir = class(TForm)
    pnBottom: TPanel;
    PageControl1: TPageControl;
    tabAdd: TTabSheet;
    gbAddFiles: TGroupBox;
    lblCompressRatio: TLabel;
    btnAdd: TBitBtn;
    btnCancelAdd: TBitBtn;
    cbCompressRatio: TComboBox;
    chKeepAboveSet: TCheckBox;
    gb1: TGroupBox;
    Label1: TLabel;
    chAddWithDir: TCheckBox;
    chOverwriteWithNew: TCheckBox;
    btnComment: TButton;
    chPassword: TCheckBox;
    tabUnzip: TTabSheet;
    GroupBox1: TGroupBox;
    btnHideFolders: TSpeedButton;
    lblPath: TLabel;
    btnUnzipThis: TBitBtn;
    btnCancelUnzipThis: TBitBtn;
    edPath: TEdit;
    chIncludeHidden: TCheckBox;
    edWildcards: TEdit;
    lblWildcards: TLabel;
    btnAddWithWildcards: TBitBtn;
    chkExcludeWildcards: TCheckBox;
    pnMiddle: TPanel;
    Splitter1: TSplitter;
    pnTree: TPanel;
    DirView1: TDirView;
    pnTop: TPanel;
    CoolBar1: TCoolBar;
    chIncludeSystem: TCheckBox;
    DriveView1: TDriveView;
    CoolBar2: TCoolBar;
    IEDriveComboBox1: TIEDriveComboBox;
    ToolBar1: TToolBar;
    pnOnCoolBar: TPanel;
    btnUp1: TSpeedButton;
    btnRefresh1: TSpeedButton;
    btnHome: TSpeedButton;
    btnCollapse1: TSpeedButton;
    btnHiddenFiles: TSpeedButton;
    btnSystemFiles: TSpeedButton;
    FilterComboBox1: TFilterComboBox;
    cbHistoryPath: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure btnUp1Click(Sender: TObject);
    procedure btnRefresh1Click(Sender: TObject);
    procedure btnCollapse1Click(Sender: TObject);
    procedure DriveView1Change(Sender: TObject; Node: TTreeNode);
    procedure cbHistoryPathClick(Sender: TObject);
    procedure cbHistoryPathKeyPress(Sender: TObject; var Key: Char);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnHideFoldersClick(Sender: TObject);
    procedure edPathKeyPress(Sender: TObject; var Key: Char);
    procedure btnUnzipThisClick(Sender: TObject);
    procedure btnCommentClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnCancelAddClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FilterComboBox1Click(Sender: TObject);

    procedure SetFilterForListView(const NewFilter: string);
    procedure DirView1AddFile(Sender: TObject; var SearchRec: TSearchRec;
      var AddFile: Boolean);
    procedure DirView1Click(Sender: TObject);
    procedure DirView1ExecFile(Sender: TObject; Item: TListItem;
      var AllowExec: Boolean);
    procedure DirView1DblClick(Sender: TObject);
    procedure tabUnzipShow(Sender: TObject);
    procedure tabAddShow(Sender: TObject);
    procedure btnHiddenFilesClick(Sender: TObject);
    procedure btnSystemFilesClick(Sender: TObject);
    procedure edWildcardsChange(Sender: TObject);
    procedure btnAddWithWildcardsClick(Sender: TObject);
    procedure chkExcludeWildcardsClick(Sender: TObject);
    procedure DirView1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DriveView1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnHideFoldersMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure IEDriveComboBox1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure IEDriveComboBox1Change(Sender: TObject);
    procedure btnHomeClick(Sender: TObject);

  private
    { Private declarations }
    procedure HistoryPathToReg(const sPath: string);
    procedure OnOffchkKeepSettings;
    procedure SettingsBeforeAdd;
    procedure RestoreSettingsOfThisForm;
    procedure OnOffchkPassword;
    procedure OnOffchkAddWithDir;
    procedure OnOffchkOverwriteWithNew;
    procedure OnOffchkIncludeSystem;
    procedure OnOffchkIncludeHidden;
    procedure BtnAddFiles_ToArchive_Click(bWithWildcards: boolean);
    {How to detect an inserted or removed CD-ROM or mounted network drives:
     The message WM_DEVICECHANGE is not passed to child-windows (Driveview1),
     so the media change can only be detected by the applications window:}
    procedure WMDeviceChange(Var Message : TMessage); message WM_DEVICECHANGE;
  public
    { Public declarations }
  end;

var
  frmDir: TfrmDir;

implementation

uses
  Main, InputMsgBox, ZipMstr, StrUtils;

const
  cPreviousOpenDir = 'LastOpenDirectory';
  cPreviousAddDir = 'LastAddDirectory';
  cIndexOfCompressRatio = 'IndexOfCompressRatio';
  cCheckKeepAboveSet = 'CheckKeepAboveSet';
  cCheckAddWithDir = 'CheckAddWithDir';
  cCheckOverwriteWithNew = 'CheckOverwriteWithNew';
  cCheckPassword = 'CheckPassword';
  cCheckIncludeSystem = 'CheckIncludeSystem';
  cCheckIncludeHidden = 'CheckIncludeHidden';
  cCountOfHistoryPath = 10;
  cDirWidthOfShellTree = 'DirWidthOfShellTree';
  cSelfWidth = 'WidthOfOpenAddFrm';
  cSelfHeight = 'HeightOfOpenAddFrm';
var
  sSubKeyHistoryDir: string;
  sComment: string = '';
{$R *.dfm}

procedure TfrmDir.FormCreate(Sender: TObject);
var
  sOldPath: String;
  i: integer;
  sWidthOfShellTree: string;
  iWidthOfShellTree: integer;
  sSelfWidth, sSelfHeight: string;
  iSelfWidth, iSelfHeight: integer;
  StartDir : String;
begin
  // DriveView must be assigned a dir. Otherwise, DriveView will cause an error after removed USB drive. Dir not found
  IF Not Assigned(DriveView1.Selected) Then
  Begin
    GetDir(0, StartDir);
    DriveView1.Directory := StartDir;
  End;

  IF Assigned(DriveView1.Selected) Then
  DriveView1.Selected.Expand(False);

  Self.Icon := Application.Icon;
  edPath.Text := '';
  cbCompressRatio.Items.Insert(0, 'Maximum');
  cbCompressRatio.Items.Insert(1, 'Normal');
  cbCompressRatio.Items.Insert(2, 'Low');
  cbCompressRatio.Items.Insert(3, 'None');
  cbCompressRatio.ItemIndex := 0;

  if frmMain.btnAdvancedView.Down = False then
      btnUnzipThis.Caption := 'Extract text';

  RestoreSettingsOfThisForm;

  sSubKeyHistoryDir := XcSubKeyIs + '\OpenAddHistoryDir';

  cbHistoryPath.Clear;
  for i := 0 to cCountOfHistoryPath -1 do begin
    sOldPath := frmMain.ReadIniFromReg(sSubKeyHistoryDir, inttostr(i));
    if sOldPath <> '' then
        cbHistoryPath.Items.Append(sOldPath);

  end;

  sWidthOfShellTree := frmMain.ReadIniFromReg(XcSubKeyIs, cDirWidthOfShellTree);
  if sWidthOfShellTree <> '' then begin
      iWidthOfShellTree := strtoint(sWidthOfShellTree);
      if (iWidthOfShellTree > 200) and (iWidthOfShellTree < 400) then
          pnTree.Width := iWidthOfShellTree;

  end;

  sSelfWidth := frmMain.ReadIniFromReg(XcSubKeyIs, cSelfWidth);
  if sSelfWidth <> '' then begin
      iSelfWidth := strtoint(sSelfWidth);
      if iSelfWidth < 600 then
          iSelfWidth := 729;

      if (iSelfWidth > 0) and ( iSelfWidth <= Screen.Width ) then
      Self.Width := iSelfWidth;
  end;

  sSelfHeight := frmMain.ReadIniFromReg(XcSubKeyIs, cSelfHeight);
  if sSelfHeight <> '' then begin
      iSelfHeight := strtoint(sSelfHeight);
      if iSelfHeight < 500 then
          iSelfHeight := 610;

      if (iSelfHeight > 0) and ( iSelfHeight <= Screen.Height ) then
      Self.Height := iSelfHeight;
  end;
end;

procedure TfrmDir.WMDeviceChange(Var Message : TMessage);
const
  DBT_DEVICEARRIVAL	   = $8000;	// system detected a new device
  DBT_DEVICEREMOVECOMPLETE = $8004;     // device is gone

Begin
  Inherited;
  IF (Message.wParam = DBT_DEVICEARRIVAL) OR
     (Message.wParam = DBT_DEVICEREMOVECOMPLETE) Then
  Begin
    DriveView1.RefreshRootNodes(True, dsAll And Not dvdsFloppy Or dvdsRereadAllways);
    IEDriveComboBox1.ResetItems;
  End;
End; {WMDeviceChange}

procedure TfrmDir.btnUp1Click(Sender: TObject);
begin
  IF Assigned(DriveView1.Selected) And
     Assigned(DriveView1.Selected.Parent) Then
  DriveView1.Selected := DriveView1.Selected.Parent;
end;

procedure TfrmDir.btnRefresh1Click(Sender: TObject);
begin
  With DriveView1 Do
  Begin
    ValidateDirectory(RootNode(Selected));
    RefreshRootNodes(False, dsDisplayName Or dsImageIndex Or dvdsFloppy or dvdsRereadAllways);
    IF ShowDirSize Then
    RefreshDriveDirSize(GetDriveToNode(Selected));
  End;
  DirView1.Reload(True);
end;

procedure TfrmDir.btnCollapse1Click(Sender: TObject);
begin
  DriveView1.FullCollapse;
end;

procedure TfrmDir.DriveView1Change(Sender: TObject; Node: TTreeNode);
var
  sDir: string;
begin
  sDir := DriveView1.GetDirPathName(DriveView1.Selected);
  if DirectoryExists(sDir) then
      //edToPath.Text := sDir;
end;

procedure TfrmDir.HistoryPathToReg(const sPath: string);
var
  i: integer;
  sPathNameInList: string;
begin
  // Is count over limit?
  if cbHistoryPath.Items.Count > cCountOfHistoryPath then begin
      for i := cbHistoryPath.Items.Count -1 downto 0 do begin
        cbHistoryPath.Items.Delete(i);
        if cbHistoryPath.Items.Count <= cCountOfHistoryPath then // be my count
            break;

      end;
  end;
  // Is path already exist?
  for i := 0 to cbHistoryPath.Items.Count -1 do begin
    if AnsiLowercase(cbHistoryPath.Items.Strings[i]) = AnsiLowercase(sPath) then
        exit;

  end;
  // Is List not full?
  if cbHistoryPath.Items.Count < cCountOfHistoryPath then
      cbHistoryPath.Items.Insert(0, sPath)
  else begin // full
      cbHistoryPath.Items.Delete(cbHistoryPath.Items.Count -1); // remove first item
      cbHistoryPath.Items.Insert(0, sPath);
  end;
  // update to registry
  for i := 0 to cbHistoryPath.Items.Count - 1 do begin
    sPathNameInList := cbHistoryPath.Items.Strings[i];
    frmMain.WriteIniToReg(sSubKeyHistoryDir, inttostr(i), sPathNameInList);
  end;
end;

procedure TfrmDir.cbHistoryPathClick(Sender: TObject);
var
  sPath, sPathNameInList: string;
  i, iCurIndex: integer;
begin
  iCurIndex := cbHistoryPath.ItemIndex;
  if (iCurIndex = -1) and (cbHistoryPath.Text = '') then exit;
  if iCurIndex <> -1 then
      sPath := cbHistoryPath.Items.Strings[iCurIndex]
  else
      sPath := cbHistoryPath.Text;
      
  if DirectoryExists(sPath) then begin
      if (length(sPath) = 2) and (AnsiPos(':', sPath) > 0) then
          sPath := sPath + '\'; // IncludeTrailingPathDelimiter does not add '\' to drive(c:)
          
      HistoryPathToReg(sPath); // don't write last trailing
      DriveView1.Directory := sPath;
      //DriveView1.SetFocus;
  end
  else begin
      beep;
      MessageDlg('Directory does not exist.', mtError, [mbOK], 0);
      cbHistoryPath.Items.Delete(cbHistoryPath.ItemIndex);
      // update to registry
      for i := 0 to cbHistoryPath.Items.Count - 1 do begin
        sPathNameInList := cbHistoryPath.Items.Strings[i];
        frmMain.WriteIniToReg(sSubKeyHistoryDir, inttostr(i), sPathNameInList);
      end;
  end;
end;

procedure TfrmDir.cbHistoryPathKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) and (cbHistoryPath.Text <> '') then begin
      Key := #0;
      cbHistoryPathClick(nil);
  end;
end;

procedure TfrmDir.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  OnOffchkKeepSettings;
  if pnTree.Width > 200 then
      frmMain.WriteIniToReg(XcSubKeyIs, cDirWidthOfShellTree, inttostr(pnTree.Width) );

  frmMain.WriteIniToReg( XcSubKeyIs, cSelfWidth, inttostr(self.Width) );
  frmMain.WriteIniToReg( XcSubKeyIs, cSelfHeight, inttostr(self.Height) );
end;

procedure TfrmDir.btnHideFoldersClick(Sender: TObject);
begin
  btnHideFolders.Down := not btnHideFolders.Down;
  if btnHideFolders.Down then
      DirView1.ShowDirectories := False
  else
      DirView1.ShowDirectories := True;
      
end;

procedure TfrmDir.edPathKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key := #0;
      btnUnzipThis.Click;
  end;
end;

procedure TfrmDir.btnUnzipThisClick(Sender: TObject);
var
  sPathFilename: string;
  sPath: string;
begin
  sPathFilename := edPath.Text;
  if (sPathFilename <> '') and FileExists(sPathFilename) then begin
      Main.bCancelOpenZip := False;
      sPath := ExtractFilePath(sPathFilename);
      sPath := ExcludeTrailingPathDelimiter(sPath);
      HistoryPathToReg(sPath);
      frmMain.WriteIniToReg(XcSubKeyIs, cPreviousOpenDir, sPath);
  end
  else begin
      beep;
      Raise Exception.Create('File does not exist.');
  end;
  close;
end;

procedure TfrmDir.OnOffchkKeepSettings;
var
  sChecked: string[1];
  i: integer;
begin
  if chKeepAboveSet.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckKeepAboveSet, sChecked);

  if chKeepAboveSet.Checked then begin
      i := cbCompressRatio.ItemIndex;
      frmMain.WriteIniToReg(XcSubKeyIs, cIndexOfCompressRatio, inttostr(i));
      OnOffchkPassword;
      OnOffchkAddWithDir;
      OnOffchkOverwriteWithNew;
      OnOffchkIncludeSystem;
      OnOffchkIncludeHidden;
  end;
end;

procedure TfrmDir.btnCommentClick(Sender: TObject);
begin
  if frmMain.ZipFName.Caption = '' then begin
      beep;
      MessageDlg('Error, source zip filename is empty', mtError, [mbOK], 0);
      exit;
  end;

  with TInputMsg.Create(OWner) do begin
    try
      sComment := ''; // Reset to null
      if frmMain.ZipMaster1.ZipFileName <> '' then
          Memo1.Text := frmMain.ZipMaster1.ZipComment;
          
      Memo1.SelectAll;
      showmodal;
      if not Main.bCancelInputComment_OnInputMsgBox then  // not cancel
          sComment := Memo1.Text;
      
    finally
      Free;
    end;
  end;
end;

procedure TfrmDir.btnAddClick(Sender: TObject);
begin
  BtnAddFiles_ToArchive_Click(False);
end;

procedure TfrmDir.SettingsBeforeAdd;
var
  zmOpts: AddOpts;
begin
  zmOpts := [];
  with frmMain.ZipMaster1 do begin
    if chPassword.Checked then begin
        repeat
          ErrCode := 0; // reset prevents dead loop
          Password := GetAddPassword;
        until ErrCode <> 10104; //GE_WrongPassword
        
        if Password <> '' then
            zmOpts := zmOpts + [AddEncrypt];
          
    end;
  end;
             
  if chAddWithDir.Checked then begin
      //zmOpts := zmOpts + [AddRecurseDirs]; // Recursively find a specified filename
      zmOpts := zmOpts + [AddDirNames];
  end;

  if chIncludeSystem.Checked then // Include system files?
      frmMain.RadioSystem.ItemIndex := 1
  else
      frmMain.RadioSystem.ItemIndex := 0;

  if chIncludeHidden.Checked then // Include hidden files?
      frmMain.RadioHidden.ItemIndex := 1
  else
      frmMain.RadioHidden.ItemIndex := 0;

  if chOverwriteWithNew.Checked then begin
      zmOpts := zmOpts + [AddUpdate]; // must before ExtrUpdate
      zmOpts := zmOpts + [AddFreshen];
  end;

  zmOpts := zmOpts + [AddResetArchive];
  frmMain.ZipMaster1.AddOptions := zmOpts;
  
  // Compress ration
  if cbCompressRatio.ItemIndex = 0 then  // Maximum compress
      frmMain.ZipMaster1.AddCompLevel := 9
  else if cbCompressRatio.ItemIndex = 1 then // normal
      frmMain.ZipMaster1.AddCompLevel := 6
  else if cbCompressRatio.ItemIndex = 2 then // Low
      frmMain.ZipMaster1.AddCompLevel := 3
  else if cbCompressRatio.ItemIndex = 3 then // none
      frmMain.ZipMaster1.AddCompLevel := 0
  else
      frmMain.ZipMaster1.AddCompLevel := 6;

end;

procedure TfrmDir.btnCancelAddClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmDir.FormActivate(Sender: TObject);
begin
  Main.bCancelAddFiles := True;
  Main.bCancelOpenZip := True;
end;

procedure TfrmDir.RestoreSettingsOfThisForm;
var
  sIndexOfCompress: string;
  iIndexOfCompress: smallint;
  sStr: string;
  //sSelfTop, sSelfLeft: string;
  //sBtnViewStyle: string[1];
  bRestoreSettings: boolean;
begin
  sStr := '';
  sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckKeepAboveSet);
  if sStr = '1' then
      bRestoreSettings := True; // leave chKeepAboveSet to uncheck

  if bRestoreSettings then begin
      sIndexOfCompress := frmMain.ReadIniFromReg(XcSubKeyIs, cIndexOfCompressRatio);
      if sIndexOfCompress <> '' then begin
          iIndexOfCompress := strtoint(sIndexOfCompress);
          cbCompressRatio.ItemIndex := iIndexOfCompress;
      end;

      sStr := '';
      sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckAddWithDir);
      if sStr = '0' then
          chAddWithDir.Checked := False;

      sStr := '';
      sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckOverwriteWithNew);
      if sStr = '0' then
          chOverwriteWithNew.Checked := False;

      sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckPassword);
      if sStr = '1' then
          chPassword.Checked := True;

      sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckIncludeSystem);
      if sStr = '0' then
          chIncludeSystem.Checked := False;

      sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckIncludeHidden);
      if sStr = '0' then
          chIncludeHidden.Checked := False;
          
  end;

  { sSelfTop := frmMain.ReadIniFromReg(XcSubKeyIs, cSelfTop);
  if sSelfTop <> '' then
      Self.Top := strtoint(sSelfTop);

  sSelfLeft := frmMain.ReadIniFromReg(XcSubKeyIs, cSelfLeft);
  if sSelfLeft <> '' then
      Self.left := strtoint(sSelfLeft);

  sBtnViewStyle := frmMain.ReadIniFromReg(XcSubKeyIs, cViewStyleUpDown);
  if sBtnViewStyle = '1' then begin
      btnViewStyle.Down := True;
      btnViewStyle.Click;
  end; }
end;

procedure TfrmDir.FilterComboBox1Click(Sender: TObject);
var
  sFilter: string;
begin
  sFilter := '';
  if FilterComboBox1.ItemIndex = 0 then
      sFilter := '*.docx'
 { else if FilterComboBox1.ItemIndex = 1 then
      sFilter := '*.zip'
  else if FilterComboBox1.ItemIndex = 2 then
      sFilter := '*.exe' }
  else if FilterComboBox1.ItemIndex = 1 then
      sFilter := '*.*';

  if sFilter <> '' then
      DirView1.Mask := sFilter;
end;

procedure TfrmDir.SetFilterForListView(const NewFilter: string);
begin
  if NewFilter = '*.docx' then
      FilterComboBox1.ItemIndex := 0
 { else if NewFilter = '*.zip' then
      FilterComboBox1.ItemIndex := 1
  else if NewFilter = '*.exe' then
      FilterComboBox1.ItemIndex := 2 }
  else if NewFilter = '*.*' then
      FilterComboBox1.ItemIndex := 1;

  FilterComboBox1Click(nil);
end;

procedure TfrmDir.DirView1AddFile(Sender: TObject;
  var SearchRec: TSearchRec; var AddFile: Boolean);
begin
  if ((SearchRec.Attr and faDirectory) = 0) then begin // greater than 0 is Dir, equal to 0 is not Dir
      if FilterComboBox1.ItemIndex = 0 then begin // Filter is docx format
          if AnsiLowercase( ExtractFileExt(SearchRec.Name) ) <> '.docx' then
              AddFile := False;

      end;
     { else if FilterComboBox1.ItemIndex = 1 then begin // Filter is zip format
          if AnsiLowercase( ExtractFileExt(SearchRec.Name) ) <> '.zip' then
              AddFile := False;

      end
      else if FilterComboBox1.ItemIndex = 2 then begin
          if AnsiLowercase( ExtractFileExt(SearchRec.Name) ) <> '.exe' then
              AddFile := False;

      end; }
  end;
end;

procedure TfrmDir.OnOffchkPassword;
var
  sChecked: string[1];
begin
  if chPassword.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckPassword, sChecked);
end;

procedure TfrmDir.OnOffchkAddWithDir;
var
  sChecked: string;
begin
  if chAddWithDir.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckAddWithDir, sChecked);
end;

procedure TfrmDir.OnOffchkOverwriteWithNew;
var
  sChecked: string;
begin
  if chOverwriteWithNew.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckOverwriteWithNew, sChecked);
end;

procedure TfrmDir.OnOffchkIncludeSystem;
var
  sChecked: string;
begin
  if chIncludeSystem.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckIncludeSystem, sChecked);
end;

procedure TfrmDir.OnOffchkIncludeHidden;
var
  sChecked: string;
begin
  if chIncludeHidden.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckIncludeHidden, sChecked);
end;

procedure TfrmDir.DirView1Click(Sender: TObject);
var
  sPathFilename: string;
begin
  if tabUnzip.Showing then begin
      sPathFilename := DirView1.GetFullFileName(DirView1.Selected);
      if FileExists(sPathFilename) then
          edPath.Text := sPathFilename;
          
  end;
end;

procedure TfrmDir.DirView1ExecFile(Sender: TObject; Item: TListItem;
  var AllowExec: Boolean);
begin
  if FileExists(DirView1.GetFullFileName(Item)) then
      AllowExec := False;
      
end;

procedure TfrmDir.DirView1DblClick(Sender: TObject);
var
  sFileExt, sPathFilename: string;
begin
  if tabUnzip.Showing then begin
      if DirView1.Selected = nil then exit;

      sPathFilename := DirView1.GetFullFileName(DirView1.Selected);
      sFileExt := AnsiLowercase(ExtractFileExt(sPathFilename));
     // if (sFileExt = '.zip') or (sFileExt = '.exe') then begin
      if (sFileExt = '.docx') then begin
          edPath.Text := sPathFilename;
          btnUnzipThis.Click;
      end;
  end;
end;

procedure TfrmDir.tabUnzipShow(Sender: TObject);
var
  StartDir, sPreviouspath: string;
begin
  sPreviouspath := frmMain.ReadIniFromReg(XcSubKeyIs, cPreviousOpenDir);

  if Main.bFavorite_OpenFolder then
        DriveView1.Directory := Main.sFavorite_OpenFolder
  else if sPreviouspath <> '' then
      DriveView1.Directory := sPreviouspath
  else begin
      if not Assigned(DriveView1.Selected) then begin
          GetDir(0, StartDir);
          DriveView1.Directory := StartDir;
      end;
  end;

  IF Assigned(DriveView1.Selected) Then
      DriveView1.Selected.Expand(False);

end;

procedure TfrmDir.tabAddShow(Sender: TObject);
var
  StartDir, sPreviouspath: string;
begin
  sPreviouspath := frmMain.ReadIniFromReg(XcSubKeyIs, cPreviousAddDir);

  if Main.bFavorite_AddFolder then
      DriveView1.Directory := Main.sFavorite_AddFolder
  else if sPreviouspath <> '' then
      DriveView1.Directory := sPreviouspath
  else begin
      if not Assigned(DriveView1.Selected) then begin
          GetDir(0, StartDir);
          DriveView1.Directory := StartDir;
      end;
  end;

  IF Assigned(DriveView1.Selected) Then
      DriveView1.Selected.Expand(False);

end;

procedure TfrmDir.btnHiddenFilesClick(Sender: TObject);
begin
  if btnHiddenFiles.Down then
      DirView1.SelHidden := SelNo // don't show hidden files
  else
      DirView1.SelHidden := SelDontCare;

  btnRefresh1.Click;
end;

procedure TfrmDir.btnSystemFilesClick(Sender: TObject);
begin
  if btnSystemFiles.Down then
      DirView1.SelSysFile := SelNo // don't show system files
  else
      DirView1.SelSysFile := SelDontCare;

  btnRefresh1.Click;
end;

procedure TfrmDir.edWildcardsChange(Sender: TObject);
begin
  btnAddWithWildcards.Enabled := (edWildcards.Text <> '');
  btnAdd.Enabled := not btnAddWithWildcards.Enabled;
end;

procedure TfrmDir.btnAddWithWildcardsClick(Sender: TObject);
begin
  BtnAddFiles_ToArchive_Click(True);
end;

procedure TfrmDir.BtnAddFiles_ToArchive_Click(bWithWildcards: boolean);
var
  sPath, sPathFilename, sWildcard, sFileExt: string;
  i, j, iA, k, iRetPos: integer;
  sTemp: string;
  sArrayCatFileExt: array of string;
  bFound: boolean;
begin
  if DirView1.SelCount = 0 then begin
      beep;
      MessageDlg('Selected nothing. Please try again.', mtError, [mbOK], 0);
      exit;
  end;

  if bWithWildcards then begin
      if edWildcards.Text = '' then begin
          beep;
          MessageDlg('Wildcard is empty. Please correct it.', mtError, [mbOK], 0);
          exit;
      end;
      
      sTemp := edWildcards.Text;
      k := 0;
      repeat
        iRetPos := AnsiPos(',', sTemp);
        if iRetPos > 0 then begin
            SetLength(sArrayCatFileExt, k+1);
            sWildcard := copy(sTemp, 1, iRetPos-1); // wildcard without ","
            sWildcard := TrimRight(sWildcard); // remove unexpected last spaces

            if (copy(sWildcard, 1, 2) <> '*.') or (length(sWildcard) <= 2) then begin
                beep;
                MessageDlg('Invalid format of wildcard found(' + sWildcard +
                '). Please correct it.', mtError, [mbOK], 0);
                exit;
            end;

            sArrayCatFileExt[k] := AnsiLowercase(ExtractFileExt(sWildcard)); // File Ext only 
            sTemp := copy(sTemp, iRetPos+1, length(sTemp)-iRetPos);
            Inc(k);
        end
        else begin
            if sTemp <> '' then begin
                sTemp := TrimRight(sTemp); // remove unexpected last spaces
                if (copy(sTemp, 1, 2) <> '*.') or (length(sTemp) <= 2) then begin
                    beep;
                    MessageDlg('Invalid format of wildcard found(' + sTemp +
                    '). Please correct it.', mtError, [mbOK], 0);
                    exit;
                end;

                SetLength(sArrayCatFileExt, k+1);
                sArrayCatFileExt[k] := AnsiLowercase(ExtractFileExt(sTemp));
            end;
        end;
      until iRetPos = 0;
  end;

  if sComment <> '' then 
      Main.sComment := sComment;

  SettingsBeforeAdd; // initial all settings
  sErrorMessage := '';

  for i := 0 to DirView1.Items.Count -1 do begin
    sPathFilename := '';
    sPath := '';

    if DirView1.Items.Item[i].Selected then begin
        sPathFilename := DirView1.GetFullFileName(DirView1.Items[i]);
        if DirectoryExists(sPathFilename) then begin
            sPath := IncludeTrailingPathDelimiter(sPathFilename);
            frmMain.ReadFilesOrFolders(sPath, '*.*');
        end
        else if FileExists(sPathFilename) then begin
           { if bWithWildcards then begin
                sWildcardFileExt := ExtractFileExt(sWildcardIs);
                sFileExt := AnsiLowercase( ExtractFileExt(sPathFileName) );
                if sWildcardFileExt = sFileExt then
                    frmMain.lboFilesToZip.Items.Add(sPathFileName);

            end
            else }
                Main.slZipStrings.Append(sPathFileName);

        end
        else
            sErrorMessage := sErrorMessage + 'Unacceptable objects : ' +
            sPathFileName + #13+#10;

    end;
  end;

  //***
  if bWithWildcards then begin
      for j := Main.slZipStrings.Count -1 downto 0 do begin
        sFileExt := AnsiLowercase(ExtractFileExt(Main.slZipStrings.Strings[j]));
        bFound := False;
        for iA := Low(sArrayCatFileExt) to High(sArrayCatFileExt) do begin
          if sArrayCatFileExt[iA] = sFileExt then begin
              bFound := True;
              break;
          end;
        end;

        if chkExcludeWildcards.Checked = True then begin
            if bFound then
                Main.slZipStrings.Delete(j);

        end
        else begin // Add with wildcards
            if not bFound then
                Main.slZipStrings.Delete(j);
                
        end;

      end;
  end;
  //***

  if sErrorMessage <> '' then
      frmMain.PopUp_ErrorMessage('Below objects will not be added to archive. ' +
      'Probably those objects are not files or directories.');

  if DriveView1.Selected <> nil then
      sPath := DriveView1.GetDirPathName(DriveView1.Selected);

  if DirectoryExists(sPath) then begin
      sPath := ExcludeTrailingPathDelimiter(sPath);
      if length(sPath) > 2 then begin
          HistoryPathToReg(sPath);
          frmMain.WriteIniToReg(XcSubKeyIs, cPreviousAddDir, sPath);
      end;
  end;

  Main.bCancelAddFiles := False;
  Close;
end;

procedure TfrmDir.chkExcludeWildcardsClick(Sender: TObject);
begin
  if chkExcludeWildcards.Checked then
      lblWildcards.Caption := 'Exclude:'
  else
      lblWildcards.Caption := 'Add with wildcards:';
      
end;

procedure TfrmDir.DirView1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Shift = [ssCtrl]) and (Key = 85) then // u
      btnUp1.Click
  else if (Shift = [ssCtrl]) and (Key = 189) then // -
      btnCollapse1.Click
  else if (Shift = [ssCtrl]) and (Key = vk_Left) then
      btnRefresh1.Click
  else if (Shift = [ssCtrl]) and (Key = 72) then begin// h
      btnHiddenFiles.Down := not btnHiddenFiles.Down;
      btnHiddenFiles.Click;
  end
  else if (Shift = [ssCtrl]) and (Key = 83) then begin// s
      btnSystemFiles.Down := not btnSystemFiles.Down;
      btnSystemFiles.Click;
  end
end;

procedure TfrmDir.DriveView1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Shift = [ssCtrl]) and (Key = 85) then // u
      btnUp1.Click
  else if (Shift = [ssCtrl]) and (Key = 189) then // -
      btnCollapse1.Click
  else if (Shift = [ssCtrl]) and (Key = vk_Left) then
      btnRefresh1.Click
end;

procedure TfrmDir.btnHideFoldersMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  btnHideFolders.Down := not btnHideFolders.Down;
end;

procedure TfrmDir.IEDriveComboBox1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  cA: char;
  sDrive: string;
begin
  if Key = 13 then begin
      cA := IEDriveComboBox1.Drive;
      if length(cA) = 1 then
          sDrive := cA + ':';

      if DirectoryExists(sDrive) then begin
          DirView1.Path := sDrive;
          //ShellDir1.SetFocus;
      end;
  end;
end;

procedure TfrmDir.IEDriveComboBox1Change(Sender: TObject);
var
  cA: char;
  sDrive: string;
begin
  cA := IEDriveComboBox1.Drive;
  if length(cA) = 1 then
      sDrive := cA + ':';

  if not DirectoryExists(sDrive) and Main.bWin9598 then
      MessageDlg('Drive not ready or inaccessible.', mtInformation, [mbOK], 0);

end;

procedure TfrmDir.btnHomeClick(Sender: TObject);
begin
  DriveView1.Selected := DriveView1.RootNode(DriveView1.Selected); // := DirView1.
  DirView1.SetFocus;
end;

end.
