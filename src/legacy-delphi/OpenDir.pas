unit OpenDir;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComCtrls, IEListView, DirView, FileCtrl,
  IEComboBox, DriveView, ExtCtrls, ToolWin, VirtualTrees,
  VirtualExplorerTree, MPCommonObjects, EasyListview,
  VirtualExplorerEasyListview, MPShellUtilities, MPCommonUtilities;

type
  TfrmOpenDir = class(TForm)
    pnTop: TPanel;
    CoolBar1: TCoolBar;
    cbHistoryPath: TComboBox;
    pnMiddle: TPanel;
    Splitter1: TSplitter;
    pnTree: TPanel;
    CoolBar2: TCoolBar;
    ToolBar1: TToolBar;
    pnOnCoolBar: TPanel;
    btnUp1: TSpeedButton;
    btnRefresh1: TSpeedButton;
    btnHome: TSpeedButton;
    btnCollapse1: TSpeedButton;
    btnHiddenFiles: TSpeedButton;
    btnSystemFiles: TSpeedButton;
    FilterComboBox1: TFilterComboBox;
    pnBottom: TPanel;
    PageControl1: TPageControl;
    tabUnzip: TTabSheet;
    GroupBox1: TGroupBox;
    btnHideFolders: TSpeedButton;
    lblPath: TLabel;
    btnUnzipThis: TBitBtn;
    btnCancelUnzipThis: TBitBtn;
    edPath: TEdit;
    VEasyList: TVirtualExplorerEasyListview;
    VTree: TVirtualExplorerTreeview;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure cbHistoryPathClick(Sender: TObject);
    procedure cbHistoryPathKeyPress(Sender: TObject; var Key: Char);
    procedure btnUp1Click(Sender: TObject);
    procedure btnRefresh1Click(Sender: TObject);
    procedure btnHomeClick(Sender: TObject);
    procedure btnUnzipThisClick(Sender: TObject);
    procedure VTreeClick(Sender: TObject);
    procedure VEasyListClick(Sender: TObject);
    procedure VEasyListEnumFolder(
      Sender: TCustomVirtualExplorerEasyListview; Namespace: TNamespace;
      var AllowAsChild: Boolean);
    procedure FilterComboBox1Click(Sender: TObject);
    procedure btnHiddenFilesClick(Sender: TObject);
    procedure btnCollapse1Click(Sender: TObject);
    procedure btnHideFoldersClick(Sender: TObject);
    procedure VEasyListDblClick(Sender: TCustomEasyListview;
      Button: TCommonMouseButton; MousePos: TPoint;
      ShiftState: TShiftState; var Handled: Boolean);
  private
    { Private declarations }
    procedure HistoryPathToReg(const sPath: string);
  public
    { Public declarations }
  end;

var
  frmOpenDir: TfrmOpenDir;

implementation

uses
  Main;

const
  cPreviousOpenDir = 'LastOpenDirectory';
  //cPreviousAddDir = 'LastAddDirectory';
  //cCheckKeepAboveSet = 'CheckKeepAboveSet';
  //cCheckAddWithDir = 'CheckAddWithDir';
  //cCheckOverwriteWithNew = 'CheckOverwriteWithNew';
  //cCheckPassword = 'CheckPassword';
  //cCheckIncludeSystem = 'CheckIncludeSystem';
  //cCheckIncludeHidden = 'CheckIncludeHidden';
  cCountOfHistoryPath = 10;
  cDirWidthOfShellTree = 'DirWidthOfShellTree';
  cSelfWidth = 'WidthOfOpenAddFrm';
  cSelfHeight = 'HeightOfOpenAddFrm';

var
  sSubKeyHistoryDir: string;

{$R *.dfm}

procedure TfrmOpenDir.FormCreate(Sender: TObject);
var
  sOldPath: String;
  i: integer;
  sWidthOfShellTree: string;
  iWidthOfShellTree: integer;
  sSelfWidth, sSelfHeight: string;
  iSelfWidth, iSelfHeight: integer;
  StartDir : String;
begin

  Self.Icon := Application.Icon;
  edPath.Text := '';


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


procedure TfrmOpenDir.FormActivate(Sender: TObject);
begin
  Main.bCancelOpenZip := True;
end;

procedure TfrmOpenDir.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if pnTree.Width > 200 then
      frmMain.WriteIniToReg(XcSubKeyIs, cDirWidthOfShellTree, inttostr(pnTree.Width) );

  frmMain.WriteIniToReg( XcSubKeyIs, cSelfWidth, inttostr(self.Width) );
  frmMain.WriteIniToReg( XcSubKeyIs, cSelfHeight, inttostr(self.Height) );

end;

procedure TfrmOpenDir.cbHistoryPathClick(Sender: TObject);
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
      VTree.BrowseTo(sPath);
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


procedure TfrmOpenDir.cbHistoryPathKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) and (cbHistoryPath.Text <> '') then begin
      Key := #0;
      cbHistoryPathClick(nil);
  end;
end;


procedure TfrmOpenDir.btnUp1Click(Sender: TObject);
var
 Node: PVirtualNode;
begin
  if not Assigned(VTree.FocusedNode) then
      exit;

  Node := VTree.FocusedNode;

  if VTree.GetNodeLevel(Node) > 0 then begin
      Node := Node.Parent;
      VTree.ClearSelection;
      VTree.FocusedNode := Node;
      VTree.Selected[Node] := True;

      VTree.Refresh;
      VTree.SetFocus;

  end;
end;

procedure TfrmOpenDir.btnRefresh1Click(Sender: TObject);
begin
  VTree.RefreshTree(True);
  VEasyList.Rebuild;
end;

procedure TfrmOpenDir.btnHomeClick(Sender: TObject);
var
  Node: PVirtualNode;
begin
  Node := VTree.GetFirst; // .TopNode is back to top depending on visible region

  //if VTree.GetNodeLevel(Node) = 0 then begin  // this is also work!!!
  if Assigned(Node) then begin
      VTree.ClearSelection;
      VTree.FocusedNode := Node;
      VTree.Selected[Node] := True;

      VTree.Refresh;
      VTree.SetFocus;
  end;

end;

procedure TfrmOpenDir.btnUnzipThisClick(Sender: TObject);
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

procedure TfrmOpenDir.VTreeClick(Sender: TObject);
var
  Node: PVirtualNode;
  NS: TNamespace;
  sFileExt, sPathFilename: string;
begin
 { Node := VTree.GetFirstSelected;
  if VTree.ValidateNamespace(Node, NS) then begin
      sPathFilename := NS.NameForParsing;

      sFileExt := AnsiLowercase(ExtractFileExt(sPathFilename));
      if not FileExists(sPathFilename) then
          exit
      else if (sFileExt <> '.docx') and (sFileExt <> '.pptx') and
      (sFileExt <> '.xlsx') then
          exit;



  end; }
end;

procedure TfrmOpenDir.HistoryPathToReg(const sPath: string);
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


procedure TfrmOpenDir.VEasyListClick(Sender: TObject);
var
  sPathFilename, sFileExt: string;
begin
  sPathFilename := VEasyList.SelectedPath;

  sFileExt := AnsiLowercase(ExtractFileExt(sPathFilename));
  if not FileExists(sPathFilename) then
      exit
  else if (sFileExt <> '.docx') and (sFileExt <> '.pptx') and
  (sFileExt <> '.xlsx') then
      exit;

  edPath.Text := sPathFilename;
end;

procedure TfrmOpenDir.VEasyListEnumFolder(
  Sender: TCustomVirtualExplorerEasyListview; Namespace: TNamespace;
  var AllowAsChild: Boolean);
var
  sExt: string;
begin
  if not Namespace.Folder then begin
      sExt := ExtractFileExt(Namespace.NameForParsing);
      sExt := Ansilowercase(sExt);
      if FilterComboBox1.ItemIndex = 0 then begin
          if (sExt <> '.docx') and (sExt <> '.pptx') and (sExt <> '.xlsx') then
              AllowAsChild := False;

      end
      else if FilterComboBox1.ItemIndex = 1 then begin
          if (sExt <> '.docx') then
              AllowAsChild := False;

      end
      else if FilterComboBox1.ItemIndex = 2 then begin
          if (sExt <> '.pptx') then
              AllowAsChild := False;

      end
      else if FilterComboBox1.ItemIndex = 3 then begin
          if (sExt <> '.xlsx') then
              AllowAsChild := False;

      end;
  end;
end;

procedure TfrmOpenDir.FilterComboBox1Click(Sender: TObject);
begin
  btnRefresh1.Click;
  VEasyList.Rebuild;
end;

procedure TfrmOpenDir.btnHiddenFilesClick(Sender: TObject);
begin
  if btnHiddenFiles.Down then
      VEasyList.FileObjects := VEasyList.FileObjects - [foHidden]
  else
      VEasyList.FileObjects := VEasyList.FileObjects + [foHidden];

end;

procedure TfrmOpenDir.btnCollapse1Click(Sender: TObject);
begin
  VTree.RebuildTree;
end;

procedure TfrmOpenDir.btnHideFoldersClick(Sender: TObject);
begin
  if btnHideFolders.Down then
      VEasyList.FileObjects := VEasyList.FileObjects - [foFolders]
  else
      VEasyList.FileObjects := VEasyList.FileObjects + [foFolders];

end;

procedure TfrmOpenDir.VEasyListDblClick(Sender: TCustomEasyListview;
  Button: TCommonMouseButton; MousePos: TPoint; ShiftState: TShiftState;
  var Handled: Boolean);
var
  sFileExt, sPathFilename: string;
begin
  if VEasyList.SelectedPath = '' then exit;

  sPathFilename := VEasyList.SelectedPath;
  sFileExt := AnsiLowercase(ExtractFileExt(sPathFilename));
  if (sFileExt = '.docx') or (sFileExt = '.pptx') or (sFileExt = '.xlsx') then begin
      edPath.Text := sPathFilename;
      btnUnzipThis.Click;
  end;
end;

end.
