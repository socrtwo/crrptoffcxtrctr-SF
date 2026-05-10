unit Extract_To;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IEComboBox, ExtCtrls, StdCtrls, ToolWin, ComCtrls, Buttons,
  IEListView, DirView, DriveView, IEDriveInfo;

type
  TfrmToFolder = class(TForm)
    pnBottom: TPanel;
    gbExtract: TGroupBox;
    edToPath: TEdit;
    GroupBox2: TGroupBox;
    rbtnAllFiles: TRadioButton;
    rbSelectedFiles: TRadioButton;
    GroupBox3: TGroupBox;
    Label2: TLabel;
    chkWithDir: TCheckBox;
    chkAlwaysOverwrite: TCheckBox;
    btnOK: TBitBtn;
    btnCancel: TBitBtn;
    pnMiddle: TPanel;
    Splitter1: TSplitter;
    DirView1: TDirView;
    pnTree: TPanel;
    DriveView1: TDriveView;
    pnTop: TPanel;
    CoolBar1: TCoolBar;
    cbHistoryPath: TComboBox;
    CoolBar2: TCoolBar;
    IEDriveComboBox1: TIEDriveComboBox;
    ToolBar1: TToolBar;
    pnOnCoolBar: TPanel;
    btnUp1: TSpeedButton;
    btnCollapse1: TSpeedButton;
    btnRefresh1: TSpeedButton;
    btnHome: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure btnUp1Click(Sender: TObject);
    procedure btnRefresh1Click(Sender: TObject);
    procedure btnCollapse1Click(Sender: TObject);
    procedure DriveView1Change(Sender: TObject; Node: TTreeNode);
    procedure btnCancelClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure cbHistoryPathClick(Sender: TObject);
    procedure cbHistoryPathKeyPress(Sender: TObject; var Key: Char);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure edToPathChange(Sender: TObject);
    procedure DirView1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DriveView1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure IEDriveComboBox1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure IEDriveComboBox1Change(Sender: TObject);
    procedure btnHomeClick(Sender: TObject);
  private
    { Private declarations }
    procedure HistoryPathToReg(const sPath: string);
    {How to detect an inserted or removed CD-ROM or mounted network drives:
     The message WM_DEVICECHANGE is not passed to child-windows (Driveview1),
     so the media change can only be detected by the applications window:}
    procedure WMDeviceChange(Var Message : TMessage); message WM_DEVICECHANGE;
  public
    { Public declarations }
  end;

var
  frmToFolder: TfrmToFolder;

implementation

uses
  Main;

const
  cPreviousDir = 'LastExtractToDir';
  cCountOfHistoryPath = 10;
  cDirWidthOfShellTree = 'ExtractToWidthOfShellTree';
  cSelfWidth = 'WidthOfExtractFrm';
  cSelfHeight = 'HeightOfExtractFrm';
var
  sSubKeyHistoryDir: string;
{$R *.dfm}

procedure TfrmToFolder.FormCreate(Sender: TObject);
var
  StartDir, sOldPath, sPreviouspath: String;
  i: integer;
  sWidthOfShellTree: string;
  iWidthOfShellTree: integer;
  sSelfWidth, sSelfHeight: string;
  iSelfWidth, iSelfHeight: integer;
begin
  // DriveView must be assigned a dir. Otherwise, DriveView will cause an error after removed USB drive. Dir not found

  sSubKeyHistoryDir := XcSubKeyIs + '\ExtractToHistoryDir';

  cbHistoryPath.Clear;
  for i := 0 to cCountOfHistoryPath -1 do begin
    sOldPath := frmMain.ReadIniFromReg(sSubKeyHistoryDir, inttostr(i));
    if sOldPath <> '' then
        cbHistoryPath.Items.Append(sOldPath);

  end;

  sPreviouspath := frmMain.ReadIniFromReg(XcSubKeyIs, cPreviousDir);
  if Main.bFavorite_ExtractFolder then
      DriveView1.Directory := Main.sFavorite_ExtractFolder
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

  sWidthOfShellTree := frmMain.ReadIniFromReg(XcSubKeyIs, cDirWidthOfShellTree);
  if sWidthOfShellTree <> '' then begin
      iWidthOfShellTree := strtoint(sWidthOfShellTree);
      if (iWidthOfShellTree > 200) and (iWidthOfShellTree < 350) then
          pnTree.Width := iWidthOfShellTree;

  end;

  sSelfWidth := frmMain.ReadIniFromReg(XcSubKeyIs, cSelfWidth);
  if sSelfWidth <> '' then begin
      iSelfWidth := strtoint(sSelfWidth);
      if iSelfWidth < 500 then
          iSelfWidth := 581;

      if (iSelfWidth > 0) and ( iSelfWidth <= Screen.Width ) then
      Self.Width := iSelfWidth;
  end;

  sSelfHeight := frmMain.ReadIniFromReg(XcSubKeyIs, cSelfHeight);
  if sSelfHeight <> '' then begin
      iSelfHeight := strtoint(sSelfHeight);
      if iSelfHeight < 500 then
          iSelfHeight := 603;

      if (iSelfHeight > 0) and ( iSelfHeight <= Screen.Height ) then
      Self.Height := iSelfHeight;
  end;
end;

procedure TfrmToFolder.WMDeviceChange(Var Message : TMessage);
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

procedure TfrmToFolder.btnUp1Click(Sender: TObject);
begin
  IF Assigned(DriveView1.Selected) And
     Assigned(DriveView1.Selected.Parent) Then
  DriveView1.Selected := DriveView1.Selected.Parent;
end;

procedure TfrmToFolder.btnRefresh1Click(Sender: TObject);
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

procedure TfrmToFolder.btnCollapse1Click(Sender: TObject);
begin
  DriveView1.FullCollapse;
end;

procedure TfrmToFolder.DriveView1Change(Sender: TObject; Node: TTreeNode);
var
  sDir: string;
begin
  sDir := DriveView1.GetDirPathName(DriveView1.Selected);
  if DirectoryExists(sDir) then
      edToPath.Text := sDir;
end;

procedure TfrmToFolder.btnCancelClick(Sender: TObject);
begin
  //edToPath.Text := '';
  Close;
end;

procedure TfrmToFolder.btnOKClick(Sender: TObject);
var
  sPath: string;
begin
  sPath := edToPath.Text;
  if sPath <> '' then begin
      HistoryPathToReg(sPath);
      //sPath := IncludeTrailingPathDelimiter(sPath);
      frmMain.WriteIniToReg(XcSubKeyIs, cPreviousDir, sPath);
      Self.Tag := 1;
  end
  else begin
      beep;
      Raise Exception.Create('Failing to get folder name.');
      exit;
  end;
  close;
end;

procedure TfrmToFolder.HistoryPathToReg(const sPath: string);
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

procedure TfrmToFolder.cbHistoryPathClick(Sender: TObject);
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
      DriveView1.SetFocus;
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

procedure TfrmToFolder.cbHistoryPathKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) and (cbHistoryPath.Text <> '') then begin
      Key := #0;
      cbHistoryPathClick(nil);
  end;
end;

procedure TfrmToFolder.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if pnTree.Width > 200 then
      frmMain.WriteIniToReg(XcSubKeyIs, cDirWidthOfShellTree, inttostr(pnTree.Width) );

  frmMain.WriteIniToReg( XcSubKeyIs, cSelfWidth, inttostr(self.Width) );
  frmMain.WriteIniToReg( XcSubKeyIs, cSelfHeight, inttostr(self.Height) );
end;

procedure TfrmToFolder.edToPathChange(Sender: TObject);
var
  AFontStyles: TFontStyles;
begin
  if not DirectoryExists(edToPath.Text) then begin
      AFontStyles := [fsBold];
      edToPath.Font.Style := AFontStyles;
  end
  else begin
      AFontStyles := [];
      edToPath.Font.Style := AFontStyles;
  end;
end;

procedure TfrmToFolder.DirView1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Shift = [ssCtrl]) and (Key = 85) then // u
      btnUp1.Click
  else if (Shift = [ssCtrl]) and (Key = 189) then // -
      btnCollapse1.Click
  else if (Shift = [ssCtrl]) and (Key = vk_Left) then
      btnRefresh1.Click
end;

procedure TfrmToFolder.DriveView1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Shift = [ssCtrl]) and (Key = 85) then // u
      btnUp1.Click
  else if (Shift = [ssCtrl]) and (Key = 189) then // -
      btnCollapse1.Click
  else if (Shift = [ssCtrl]) and (Key = vk_Left) then
      btnRefresh1.Click
end;

procedure TfrmToFolder.IEDriveComboBox1KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
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

procedure TfrmToFolder.IEDriveComboBox1Change(Sender: TObject);
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

procedure TfrmToFolder.btnHomeClick(Sender: TObject);
//var
//  sPath: string;
//  iRetPos: integer;
begin
 { sPath := DirView1.Path;

  if sPath <> '' then begin
      iRetPos := AnsiPos('\', sPath);
      if iRetPos > 0 then begin
          sPath := copy(sPath , 1, iRetPos -1);
          DirView1.Path := sPath;

          DriveView1.Selected := DriveView1.RootNode(DriveView1.Selected); // := DirView1.
          DirView1.SetFocus;
      end;
  end; }

  DriveView1.Selected := DriveView1.RootNode(DriveView1.Selected); // := DirView1.
  DirView1.SetFocus;
end;

end.
