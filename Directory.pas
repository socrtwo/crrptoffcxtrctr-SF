unit Directory;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ShellCtrls, Buttons, FileCtrl, StrUtils,
  ExtCtrls, ToolWin;

type
  TGetFolderName = class(TForm)
    PageControl1: TPageControl;
    tabExtract: TTabSheet;
    pnShellList: TPanel;
    ShellListView1: TShellListView;
    Splitter1: TSplitter;
    CoolBar1: TCoolBar;
    pnCB: TPanel;
    dcDrive: TDriveComboBox;
    lblPath: TLabel;
    lblPreDrive: TLabel;
    lblGetDirName: TLabel;
    lblPathFilename: TLabel;
    cbHistoryPath: TComboBox;
    btnOK: TButton;
    btnCancel: TButton;
    gbBottom: TGroupBox;
    edToPath: TEdit;
    btnQuit: TBitBtn;
    btnExtract: TButton;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    chkWithDir: TCheckBox;
    chkAlwaysOverwrite: TCheckBox;
    GroupBox1: TGroupBox;
    rbtnAllFiles: TRadioButton;
    rbSelectedFiles: TRadioButton;
    pnTree: TPanel;
    stvDirectory: TShellTreeView;
    CoolBar2: TCoolBar;
    btnUp1: TSpeedButton;
    btnRefresh1: TSpeedButton;
    btnCollapse1: TSpeedButton;
    procedure btnOKClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure dcDriveChange(Sender: TObject);
    procedure cbHistoryPathClick(Sender: TObject);
    procedure cbHistoryPathDropDown(Sender: TObject);
    procedure btnExtractClick(Sender: TObject);
    procedure stvDirectoryMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnQuitClick(Sender: TObject);
    procedure ShellListView1Change(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure edToPathChange(Sender: TObject);
    procedure SetLastPath;
    procedure FormShow(Sender: TObject);
    procedure cbHistoryPathKeyPress(Sender: TObject; var Key: Char);
    procedure stvDirectoryKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
    procedure btnUp1Click(Sender: TObject);
    procedure btnCollapse1Click(Sender: TObject);
    procedure btnRefresh1Click(Sender: TObject);
  private
    { Private declarations }
    procedure HistoryPathToReg(const sPath: string);
    procedure SelectPath(sPath: string);
    function SelectRecursiveDir(const CurTreeNode: TTreeNode;
     sDir: string): TTreeNode;
    procedure SaveStatusOfThisForm;
    procedure RestoreStatusOfThisForm;
  public
    { Public declarations }
  end;

var
  GetFolderName: TGetFolderName;

implementation

uses
  Main;

const
  cPreviousDir = 'Last Directory';
  cCountOfHistoryPath = 10;

  cDirTop = 'DirTop';
  cDirLeft = 'DirLeft';
  cDirWidth = 'DirWidth';
  cDirHeight = 'DirHeight';
  cDirWindowState = 'DirWindowState';
  cDirWidthOfShellTree = 'DirWidthOfShellTree';
var
  sSubKeyHistoryDir: string = XcSubKeyIs + '\HistoryDir';
  sPreviouspath: string;
  bDescending: boolean = False;
{$R *.dfm}

procedure TGetFolderName.FormCreate(Sender: TObject);
var
  sOldPath: string;
  i: integer;
begin
  Self.Icon := Application.Icon;
  Self.Caption := XcProgramName + '...Select folder';
  lblPreDrive.Caption := '0';

  cbHistoryPath.Clear;
  for i := 0 to cCountOfHistoryPath -1 do begin
    sOldPath := frmMain.ReadIniFromReg(sSubKeyHistoryDir, inttostr(i));
    if sOldPath <> '' then
        cbHistoryPath.Items.Append(sOldPath);

  end;

  sPreviouspath := frmMain.ReadIniFromReg(XcSubKeyIs, cPreviousDir);
  //SelectPath(sPreviousPath);
  RestoreStatusOfThisForm;
end;

procedure TGetFolderName.btnOKClick(Sender: TObject);
var
  sPath: string;
begin
 { sPath := stvDirectory.SelectedFolder.PathName;
  if (sPath <> '') and DirectoryExists(sPath) then begin
      if copy(sPath, length(sPath), 1) <> '\' then
          sPath := sPath + '\';

      //frmMain.SetCurrentPath(sPath);
      lblPath.Caption := sPath;
      frmMain.WriteIniToReg(XcSubKeyIs, cPreviousDir, sPath);
      HistoryPathToReg(sPath);
  end
  else begin
      beep;
      Raise Exception.Create('Fail to get folder name.');
  end;
  close; }
end;

procedure TGetFolderName.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  SaveStatusOfThisForm;
end;

procedure TGetFolderName.dcDriveChange(Sender: TObject);
var
  sPathName: string;
begin
  if lblPreDrive.Caption = '0' then begin // first time run
      lblPreDrive.Caption := '1';
      exit;
  end;
  if lblPreDrive.Caption = dcDrive.Drive then exit;
  lblPreDrive.Caption := dcDrive.Drive;

  sPathName := dcDrive.Drive + ':\';
  if DirectoryExists(sPathName) then begin
      stvDirectory.Root := dcDrive.Drive + ':\';
      if stvDirectory.Showing then
          stvDirectory.SetFocus;
          
  end
  else begin
      beep;
      MessageDlg('Drive ' + sPathName + ' not ready.', mtError, [mbOK], 0);
  end;
end;

procedure TGetFolderName.HistoryPathToReg(const sPath: string);
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
    if lowercase(cbHistoryPath.Items.Strings[i]) = lowercase(sPath) then
        exit;

  end;
  // not full?
  if cbHistoryPath.Items.Count < cCountOfHistoryPath then
      cbHistoryPath.Items.Insert(0, sPath)
  else begin // full
      cbHistoryPath.Items.Delete(cbHistoryPath.Items.Count -1); 
      cbHistoryPath.Items.Insert(0, sPath);
  end;
  // update to registry
  for i := 0 to cbHistoryPath.Items.Count -1 do begin
    sPathNameInList := cbHistoryPath.Items.Strings[i];
    frmMain.WriteIniToReg(sSubKeyHistoryDir, inttostr(i), sPathNameInList);
  end;
end;

procedure TGetFolderName.SelectPath(sPath: string);
var
  i: integer;
  PLetter: PChar;
  iRetPos: integer;
  sTemp: string;

  sArrayDir: array of string;
  k: integer;
  sDirIs: string;
  ThisTreeNode: TTreeNode;
  chDriveIs: char;
  sDriveIs: string;
begin
  if sPath = '' then exit;
  sPath := IncludeTrailingPathDelimiter(sPath);
  sPath := lowercase(sPath);
  sDriveIs := ExtractFileDrive(sPath);
  if sDriveIs <> '' then begin // e.g local drive = "c:\" , remote = "\\unc\sharename" (may be has delimiter)
      sTemp := AnsiReplaceStr(sPath, sDriveIs, ''); // remove drive name
      if copy(sTemp, 1, 1) = '\' then
          sTemp := copy(sTemp, 2, length(sTemp)-1);

  end;

 { if length(sPath) > 3 then begin // Has drive and directory
      sTemp := copy(sPath, 4, length(sPath)-3); // remove drive name
  end; }

  k := 0;
  repeat
    iRetPos := Pos('\', sTemp);
    if iRetPos > 0 then begin
        SetLength(sArrayDir, k+1);
        sDirIs := copy(sTemp, 1, iRetPos-1); // dir without "\"
        sArrayDir[k] := sDirIs;
        sTemp := copy(sTemp, iRetPos+1, length(sTemp)-iRetPos);
        Inc(k);
    end;
  until iRetPos = 0;

  if length(sPath) > 3 then
      sPath := ExcludeTrailingPathDelimiter(sPath);

  if DirectoryExists(sPath) then begin
      PLetter := PChar(copy(sPath, 1, 1));
      chDriveIs := PLetter^;
      if chDriveIs = lowercase(dcDrive.Drive) then
          stvDirectory.Root := IncludeTrailingPathDelimiter(sDriveIs);

      dcDrive.Drive := PLetter^;

      if stvDirectory.Items[0] = nil then exit;
      ThisTreeNode := stvDirectory.Items[0];
      for i := Low(sArrayDir) to High(sArrayDir) do begin
        ThisTreeNode := SelectRecursiveDir(ThisTreeNode, sArrayDir[i]);
        if ThisTreeNode = nil then break;
      end;
  end;
end;

procedure TGetFolderName.cbHistoryPathClick(Sender: TObject);
var
  sPath, sPathNameInList: string;
  i, iCurIndex: integer;
begin
  iCurIndex := cbHistoryPath.ItemIndex;
  if iCurIndex = -1 then exit;
  sPath := cbHistoryPath.Items.Strings[iCurIndex];
  if DirectoryExists(sPath) then begin
      btnOK.Enabled := False;
      SelectPath(sPath);
      edToPath.Text := sPath;
      btnOK.Enabled := True;
  end
  else begin
      beep;
      MessageDlg('Directory does not exist.', mtError, [mbOK], 0);
      cbHistoryPath.DeleteSelected;
      // update to registry
      for i := cbHistoryPath.Items.Count - 1 downto 0 do begin
        sPathNameInList := cbHistoryPath.Items.Strings[i];
        frmMain.WriteIniToReg(sSubKeyHistoryDir, inttostr(i), sPathNameInList);
      end;
  end;
end;

function TGetFolderName.SelectRecursiveDir(const CurTreeNode: TTreeNode;
 sDir: string): TTreeNode;
var
  i: integer;
begin
  Result := nil;
  for i := 0 to CurTreeNode.Count -1 do begin
        lblGetDirname.Caption := '';
        lblGetDirName.Caption := CurTreeNode.Item[i].Text; // no trailing
        lblGetDirName.Caption := lowercase(lblGetDirName.Caption);
        if sDir = lblGetDirName.Caption then begin
            CurTreeNode.Item[i].Selected := True;
            CurTreeNode.Item[i].MakeVisible;
            CurTreeNode.Item[i].Expand(False);
            if stvDirectory.Showing then
                stvDirectory.SetFocus;
                
            Result := CurTreeNode.Item[i];
            break;
        end;
  end;
end;

procedure TGetFolderName.cbHistoryPathDropDown(Sender: TObject);
begin
  cbHistoryPath.Text := '';
  cbHistoryPath.DropDownCount := cbHistoryPath.Items.Count;
end;

procedure TGetFolderName.btnExtractClick(Sender: TObject);
var
  sPath: string;
begin
  sPath := edToPath.Text;
  if (sPath <> '') then begin //and DirectoryExists(sPath) then begin
      //sPath := IncludeTrailingPathDelimiter(sPath);
      lblPath.Caption := sPath;
      frmMain.WriteIniToReg(XcSubKeyIs, cPreviousDir, sPath);
      HistoryPathToReg(sPath);
      //Self.Tag := 0;
      Main.bCancelExtract_OnDirectory := False;
  end;
 { else begin
      beep;
      Raise Exception.Create('Failing to get folder name.');
  end; }
  Close;
end;

procedure TGetFolderName.stvDirectoryMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if tabExtract.Showing then
      edToPath.Text := stvDirectory.SelectedFolder.PathName;

end;

procedure TGetFolderName.btnQuitClick(Sender: TObject);
begin
  Close;
end;

procedure TGetFolderName.ShellListView1Change(Sender: TObject;
  Item: TListItem; Change: TItemChange);
//var
//  sPathFilename: string;
begin
  if tabExtract.Showing then begin
      if ShellListView1.SelectedFolder <> nil then begin
          if DirectoryExists(ShellListView1.SelectedFolder.PathName) then
              edToPath.Text := ShellListView1.SelectedFolder.PathName; // stvDirectory.SelectedFolder.PathName;
      end;
  end;
 { else if tabUnzip.Showing then begin
      if ShellListView1.SelectedFolder <> nil then
          sPathFilename := ShellListView1.SelectedFolder.PathName;

      leFile.Text := sPathFilename;
      lblPathFilename.Caption := leFile.Text;
  end; }
end;

procedure TGetFolderName.SaveStatusOfThisForm;
var
  sWindowStateIs: string;
begin
  if Self.WindowState = wsMaximized then
      sWindowStateIs := 'wsMaximized'
  else if Self.WindowState = wsNormal then begin
      frmMain.WriteIniToReg(XcSubKeyIs, cDirTop, inttostr(Self.Top));
      frmMain.WriteIniToReg(XcSubKeyIs, cDirLeft, inttostr(Self.Left));
      frmMain.WriteIniToReg(XcSubKeyIs, cDirWidth, inttostr(Self.Width));
      frmMain.WriteIniToReg(XcSubKeyIs, cDirHeight, inttostr(Self.Height));
  end;
  frmMain.WriteIniToReg(XcSubKeyIs, cDirWindowState, sWindowStateIs);
  frmMain.WriteIniToReg(XcSubKeyIs, cDirWidthOfShellTree, inttostr(stvDirectory.Width) );
end;

procedure TGetFolderName.RestoreStatusOfThisForm;
var
  iDirTop, iDirLeft, iDirWidth, iDirHeight: longint;
  sDirTop, sDirLeft, sDirWidth, sDirHeight: string;
  sWindowStateIs: string;
  iWidthOfShellTree: smallint;
  sWidthOfShellTree: string;
begin
  // vvv Window state vvv
  sWindowStateIs := frmMain.ReadIniFromReg(XcSubKeyIs, cDirWindowState);
  if sWindowStateIs = 'wsMaximized' then
      Self.WindowState := wsMaximized
  else begin
  sDirTop := frmMain.ReadIniFromReg(XcSubKeyIs, cDirTop);
  if sDirTop <> '' then begin
      iDirTop := strtoint(sDirTop);
      if (iDirTop > 0) and (iDirTop < Screen.Height) then
          Self.Top := iDirTop;

  end;
  sDirLeft := frmMain.ReadIniFromReg(XcSubKeyIs, cDirLeft);
  if sDirLeft <> '' then begin
      iDirLeft := strtoint(sDirLeft);
      if (iDirLeft > 0) and (iDirLeft < Screen.Width) then
          Self.Left := iDirLeft;

  end;
  sDirWidth := frmMain.ReadIniFromReg(XcSubKeyIs, cDirWidth);
  if sDirWidth <> '' then begin
      iDirWidth := strtoint(sDirWidth);
      if (iDirWidth > 0) and ( iDirWidth <= Screen.Width ) then
          Self.Width := iDirWidth;

  end;
  sDirHeight := frmMain.ReadIniFromReg(XcSubKeyIs, cDirHeight);
  if sDirHeight <> '' then begin
      iDirHeight := strtoint(sDirHeight);
      if (iDirHeight > 0) and ( iDirHeight <= Screen.Height) then
          Self.Height := iDirHeight;

  end;
  end;
  // ^^^
  sWidthOfShellTree := frmMain.ReadIniFromReg(XcSubKeyIs, cDirWidthOfShellTree);
  if sWidthOfShellTree <> '' then begin
      iWidthOfShellTree := strtoint(sWidthOfShellTree);
      if (iWidthOfShellTree > 10) and (iWidthOfShellTree <
      (Screen.Width div 2) ) then
          stvDirectory.Width := iWidthOfShellTree;

  end;
end;
procedure TGetFolderName.edToPathChange(Sender: TObject);
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

procedure TGetFolderName.SetLastPath;
begin
  SelectPath(sPreviouspath);
end;

procedure TGetFolderName.FormShow(Sender: TObject);
begin
  SetLastPath;
end;

procedure TGetFolderName.cbHistoryPathKeyPress(Sender: TObject;
  var Key: Char);
var
  sPath: string;
begin
  if Key <> Char(13) then exit;
  sPath := cbHistoryPath.Text;
  //sPath := IncludeTrailingPathDelimiter(sPath);
  if DirectoryExists(sPath) then begin
      btnOK.Enabled := False;
      SelectPath(sPath);
      edToPath.Text := sPath;
      btnOK.Enabled := True;
  end
  else begin
      beep;
      MessageDlg('Directory does not exist.', mtError, [mbOK], 0);
  end;
end;

procedure TGetFolderName.stvDirectoryKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if tabExtract.Showing then
      edToPath.Text := stvDirectory.SelectedFolder.PathName;
end;

procedure TGetFolderName.FormActivate(Sender: TObject);
begin
  Main.bCancelExtract_OnDirectory := True;
end;

procedure TGetFolderName.btnUp1Click(Sender: TObject);
var
  iSelected: integer;
begin
  if stvDirectory.Selected.Parent <> nil then begin
      stvDirectory.Selected := stvDirectory.Selected.Parent;
      iSelected := stvDirectory.Selected.AbsoluteIndex;
      edTopath.Text := stvDirectory.Folders[iSelected].PathName;
      stvDirectory.Selected.MakeVisible;
      stvDirectory.SetFocus;
  end;
end;

procedure TGetFolderName.btnCollapse1Click(Sender: TObject);
var
  iSelected: integer;
begin
  stvDirectory.FullCollapse;
  stvDirectory.TopItem.Expand(False);
  if stvDirectory.SelectionCount > 0 then begin
      iSelected := stvDirectory.Selected.AbsoluteIndex;
      edTopath.Text := stvDirectory.Folders[iSelected].PathName;
  end;
  stvDirectory.SetFocus;
end;

procedure TGetFolderName.btnRefresh1Click(Sender: TObject);
begin
  stvDirectory.Refresh(stvDirectory.Items[0]);
end;

end.
