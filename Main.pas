unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Gauges, ComCtrls, ExtCtrls, ImgList, Menus, ShellCtrls, ClipBrd,
  StdCtrls, FileCtrl, ToolWin, Buttons, ZipMstr, Registry, StrUtils, CommCtrl,
  DateUtils, IEListView, DirView, IEComboBox, IEDriveInfo, ComObj,
  Math, VirtualTrees, Contnrs, DriveView, LibXmlParser,
  VirtualExplorerTree, MPShellUtilities, pngimage, JvComponent, JvBaseDlg,
  JvBrowseFolder, JvExControls, JvArrowButton, Grids, JvImageList;


type
  TImagesOnSheet = record

  iSheetID: integer;
  iCol: integer;
  iToCol: integer;
  iRow: integer;
  iToRow: integer;
  iImageID: integer;
  iCx: integer;
  iCy: integer;
end;

type
  THeightOfRow = record

  iSheetID: integer;
  iIndex: integer;
  iHeight: integer;
end;

type
  TRegisteredUser = record

  sHeader: string[17];
  sRegUser: string[255];
end;

type
  //PVTPropEditData = ^TVTPropEditData;
  TVTPropEditData = class
  //FHeading: string;        // heading
  FCaption: TStringList;  // sub-headings
  FModified: TStringList;  // list of values in string form
  FSize: TStringList;
  FRatio: TStringList;
  FColPacked: TStringList;
  FFileType: TStringList;
  FCRC: TStringList;
  FAttributes: TStringList;
  FPath: TStringList;

  public
  { Public declarations }
  constructor Create;
  destructor Destroy;                                                              override;

  //property Heading: string          read FHeading     write FHeading ;
  property Caption: TStringList     read FCaption    write FCaption;
  property Modified: TStringList     read FModified      write FModified;
  property Size: TStringList     read FSize      write FSize;
  property Ratio: TStringList     read FRatio      write FRatio;
  property ColPacked: TStringList     read FColPacked      write FColPacked;
  property FileType: TStringList     read FFileType      write FFileType;
  property CRC: TStringList     read FCRC      write FCRC;
  property Attributes: TStringList     read FAttributes      write FAttributes;
  property Path: TStringList     read FPath      write FPath;
end;

type
  PEntry = ^TEntry;
  TEntry = record
  Value: array[0..9] of string;
end;


type
  PMyEntry = ^TMyEntry;
  TMyEntry = record
  Value: array[0..18278] of string;        // Max Col to ZZZ
  Cell_Loc: array[0..18278] of string;     // location of cells
end;


type
  TfrmMain = class(TForm)
    MainMenu1: TMainMenu;
    mnuFile1: TMenuItem;
    mnuSubNew1: TMenuItem;
    mnuSubOpen1: TMenuItem;
    mnuSubClose1: TMenuItem;
    N3: TMenuItem;
    mnuSubExit: TMenuItem;
    mnuActions: TMenuItem;
    mnuSubAdd: TMenuItem;
    mnuSubExtract: TMenuItem;
    mnuSubTest: TMenuItem;
    mnuSubSFX: TMenuItem;
    N4: TMenuItem;
    mnuSubViewComment: TMenuItem;
    mnuAddComment: TMenuItem;
    N7: TMenuItem;
    mnuSubSelectAll: TMenuItem;
    mnuSubInvertSelection: TMenuItem;
    mnuSettings1: TMenuItem;
    mnuSubPreferences: TMenuItem;
    mnuSubShowBrowse: TMenuItem;
    mnuDllVer: TMenuItem;
    mnuHelp: TMenuItem;
    smHelpFile: TMenuItem;
    smAbout: TMenuItem;
    OpenedFilesPopup: TPopupMenu;
    OpenDialog1: TOpenDialog;
    MainPopup: TPopupMenu;
    mnuZipDelete: TMenuItem;
    mnuDeleteAll: TMenuItem;
    Splitter1: TSplitter;
    SaveDialog: TSaveDialog;
    PopupMenu1: TPopupMenu;
    mnuOpenWith: TMenuItem;
    N1: TMenuItem;
    mnuDelete: TMenuItem;
    mnuAddFilesToZip: TMenuItem;
    mnuTest: TMenuItem;
    N5: TMenuItem;
    mnuViewComment: TMenuItem;
    pm1AddComment: TMenuItem;
    N2: TMenuItem;
    mnuSelectAll: TMenuItem;
    mnuInvertSelection: TMenuItem;
    N6: TMenuItem;
    mnuCloseAchive: TMenuItem;
    PnInvisible: TPanel;
    RadioVerboseOpt: TRadioGroup;
    RadioExtractWithDir: TRadioGroup;
    RadioOverwrite: TRadioGroup;
    RadioDirNames: TRadioGroup;
    RadioRecurse: TRadioGroup;
    lblExtractParaStr: TLabel;
    edtZipFileName: TEdit;
    edShadow: TEdit;
    mnuOpen: TMenuItem;
    mnuExtract: TMenuItem;
    mnuTerminate: TMenuItem;
    N8: TMenuItem;
    pmFavoriteDir: TPopupMenu;
    ZipFName: TLabel;
    N9: TMenuItem;
    mnuDiskSpan: TMenuItem;
    mnuCombine: TMenuItem;
    mnuSFXToZip: TMenuItem;
    mnuSubCopyArchive: TMenuItem;
    N10: TMenuItem;
    mnuSubDelete: TMenuItem;
    mnuSubOpen: TMenuItem;
    N11: TMenuItem;
    mnuSubOpenWith: TMenuItem;
    ilMenu: TImageList;
    ilTree: TImageList;
    CoolBarTop: TCoolBar;
    Panel4: TPanel;
    btnNew: TSpeedButton;

    btnClose: TSpeedButton;
    btnUnZip: TSpeedButton;
    btnSfxWizard: TSpeedButton;
    btnExit: TSpeedButton;
    mnuProperties: TMenuItem;
    mnuSubProperties: TMenuItem;
    mnuSubMoveArchive: TMenuItem;
    mnuSubRenameArchive: TMenuItem;
    mnuSubDeleteArchive: TMenuItem;
    N12: TMenuItem;
    mnuSubSort: TMenuItem;
    mnuSubByName: TMenuItem;
    mnuSubByModified: TMenuItem;
    mnuSubBySize: TMenuItem;
    mnuSubByRatio: TMenuItem;
    mnuSubByPacked: TMenuItem;
    mnuSubByFileType: TMenuItem;
    mnuSubByCRC: TMenuItem;
    mnuSubByPath: TMenuItem;
    mnuSubByAttributes: TMenuItem;
    mnuSubByDefault: TMenuItem;
    mnuRename: TMenuItem;
    mnuSubRename: TMenuItem;
    N13: TMenuItem;
    smHomePage: TMenuItem;
    mnuRepair: TMenuItem;
    RadioSystem: TRadioGroup;
    RadioHidden: TRadioGroup;
    mnuSubSearch: TMenuItem;
    mnuVirusScan: TMenuItem;
    ilSystem: TImageList;
    ZipMaster1: TZipMaster;
    ilSysIcons: TImageList;
    btnViewTxt: TSpeedButton;
    pnB: TPanel;
    pnBottom: TPanel;
    StatusBar1: TStatusBar;
    pnBottomLeft: TPanel;
    Gauge1: TGauge;
    Gauge2: TGauge;
    memOutput: TMemo;
    btnSaveTxtToFile: TSpeedButton;
    pnM: TPanel;
    pnListBack: TPanel;
    VT: TVirtualStringTree;
    ListView1: TListView;
    pnListTop: TPanel;
    btnComment: TSpeedButton;
    lblCurDir: TLabel;
    imgCurDir: TImage;
    PanMiddleLeft: TPanel;
    PanVirtualTree: TPanel;
    TreeView1: TTreeView;
    VTcoolBar: TCoolBar;
    Panel5: TPanel;
    btnVTup: TSpeedButton;
    btnVThome: TSpeedButton;
    btnVTcollapse: TSpeedButton;
    pnShell: TPanel;
    CoolBar1: TCoolBar;
    FilterComboBox1: TFilterComboBox;
    pnCB: TPanel;
    ToolBar2: TToolBar;
    pnOnCoolBar: TPanel;
    btnHome: TSpeedButton;
    btnUp1: TSpeedButton;
    btnRefresh1: TSpeedButton;

    Panel3: TPanel;
    Panel2: TPanel;
    Panel1: TPanel;
    PanMidLeftUp: TPanel;
    spCloseBrowsePanel: TSpeedButton;
    btnShowVirtual: TSpeedButton;
    btnShiftViewMode: TSpeedButton;
    btnAdvancedView: TSpeedButton;
    Splitter2: TSplitter;
    btnAbout: TButton;
    btnEditSelXML: TBitBtn;
    btnTest: TSpeedButton;
    btnRepair: TSpeedButton;
    OpenDlgSimple: TOpenDialog;
    lblOpenFileIs: TLabel;
    pmWord: TPopupMenu;
    mnuViewXMLContents: TMenuItem;
    btnHelp: TSpeedButton;
    btnAddFiles: TSpeedButton;
    btnHLImgFiles: TSpeedButton;
    btnRecoverIMAGES: TBitBtn;
    VETree: TVirtualExplorerTreeview;
    //JvBrowseFolder1: TJvBrowseFolder;
    pnWorksheets: TPanel;
    PgeCtr: TPageControl;
    JvBrowseFolder1: TJvBrowseForFolderDialog;
    btnFavoriteDir: TJvArrowButton;
    btnPopupOpen: TJvArrowButton;
    StringGrid1: TStringGrid;
    pnBack: TPanel;
    pmSG: TPopupMenu;
    pmCopyImgObj: TMenuItem;
    VTHide: TVirtualStringTree;
    btnSaveAllToCSV: TSpeedButton;
    pnPB: TPanel;
    lbPB: TLabel;
    GaugePB: TGauge;
    
    procedure SetZipFile(FName: String; bVerbose_List: boolean);
    procedure WriteIniToReg(sSubKey: string; sValueName: string; sValue: string);
    function ReadIniFromReg(sSubKey: string; sValueName: string): string;
    procedure WriteRegLocalMachine(sSubKey: string; sValueName: string; sValue: string);
    function ReadRegLocalMachine(sSubKey: string; sValueName: string): string;


    procedure SetStatusOfchkAutoSizeToVar(const bAutoSizeChecked: boolean);
    procedure ResizeWidthOfMainListColumns;
    procedure SetWidthOfMainListColumns;
    procedure RegisterFileType(cMyExt, cMyFileType, cMyDescription,
     ExeName: string; IcoIndex: integer; DoUpdate: boolean);
    procedure ReadFilesOrFolders(sPathIs: string; sWildcards: string);
    procedure BeforeExtracting(const sToPath: string; const bWithDir,
     bAlwaysOverwrite, bAllFiles: boolean);
    procedure StartUnzip(bVerbose: boolean; sZipFilename: string;
     sUnzipToFolder: string; sSelPathFNames: TStrings; bExtractSelected: boolean);
    procedure StartZip(sToPathZipName: string; sFiles: TStrings);

    procedure btnNewClick(Sender: TObject);
    procedure mnuAddFilesToZipClick(Sender: TObject);
    function ReadFileAndDelete(Sender: TObject; sPathIs: string): boolean;
    procedure KillSubFolderFiles(Sender: TObject; sPathIs: string);
    procedure FileOrFolder_ThroughContext_ToBeZipped(sPathZipName: string);
    procedure btnPopupOpenClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure mnuCloseAchiveClick(Sender: TObject);
    procedure RecentUnzipFilesToReg(const sPathFilenameIs: string);
    procedure mnuOpenWithClick(Sender: TObject);
    function GetPathFileName_FromSelected_Item(var bIsFolder, bIsEncrypted: boolean): string;
    procedure ExecuteFileClick(Sender: TObject; sExtractedPathFilename: string;
     bDoubleClick: boolean; sWithThisApp: string);
    procedure FileOpenedBy(sPahtFileNameIs: string);
    procedure mnuDeleteClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure mnuTestClick(Sender: TObject);

    Procedure ZipMaster1Message(Sender: TObject; ErrCode: Integer; Message: String);
    procedure SetStatusOfPopupUnzip(Sender: TObject);
    procedure btnAddFilesClick(Sender: TObject);
    procedure btnUnZipClick(Sender: TObject);
    procedure btnSfxWizardClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure mnuSubNew1Click(Sender: TObject);
    procedure mnuSubOpen1Click(Sender: TObject);
    procedure mnuSubClose1Click(Sender: TObject);
    procedure mnuSubExitClick(Sender: TObject);
    procedure mnuSubAddClick(Sender: TObject);
    procedure mnuSubExtractClick(Sender: TObject);
    procedure mnuSubTestClick(Sender: TObject);
    procedure mnuSubSFXClick(Sender: TObject);
    procedure mnuSelectAllClick(Sender: TObject);
    procedure mnuInvertSelectionClick(Sender: TObject);
    procedure mnuSubSelectAllClick(Sender: TObject);
    procedure mnuSubInvertSelectionClick(Sender: TObject);
    procedure mnuSubPreferencesClick(Sender: TObject);
    procedure mnuSubShowBrowseClick(Sender: TObject);
    procedure smHelpFileClick(Sender: TObject);
    procedure smAboutClick(Sender: TObject);

    procedure spCloseBrowsePanelClick(Sender: TObject);
    procedure btnUp1Click(Sender: TObject);
    procedure btnRefresh1Click(Sender: TObject);
    procedure FilterComboBox1Click(Sender: TObject);
    procedure FormCanResize(Sender: TObject; var NewWidth,
      NewHeight: Integer; var Resize: Boolean);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure FormResize(Sender: TObject);

    procedure mnuViewCommentClick(Sender: TObject);
    procedure mnuSubViewCommentClick(Sender: TObject);
    procedure pm1AddCommentClick(Sender: TObject);
    procedure mnuDllVerClick(Sender: TObject);
    procedure PopUp_ErrorMessage(sErr: string);
    procedure mnuOpenClick(Sender: TObject);
    procedure btnHomeClick(Sender: TObject);
    procedure WhatCompressionFile_WillBe_DirectlyExtracted(sFilename: string);
    procedure ZipFile_WillBe_DirectlyConvertToExe(sFilename: string);
    procedure Files_WantToBe_Compressed_W_WO_ShowMain(sOutputTo: string; sTheDirIs: string);
    procedure mnuExtractClick(Sender: TObject);
    procedure ZipMaster1PasswordError(Sender: TObject;
      IsZipAction: Boolean; var NewPassword: String; ForFile: String;
      var RepeatCount: Cardinal; var Action: TPasswordButton);
    procedure mnuTerminateClick(Sender: TObject);
    procedure mnuDiskSpanClick(Sender: TObject);
    procedure mnuCombineClick(Sender: TObject);
    procedure ListView1ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView1Compare(Sender: TObject; Item1, Item2: TListItem;
      Data: Integer; var Compare: Integer);
    procedure mnuSFXToZipClick(Sender: TObject);
    procedure mnuSubCopyArchiveClick(Sender: TObject);
    procedure mnuSubDeleteClick(Sender: TObject);
    procedure mnuSubOpenClick(Sender: TObject);
    procedure mnuSubOpenWithClick(Sender: TObject);
    procedure TreeView1Change(Sender: TObject; Node: TTreeNode);
    procedure btnShowVirtualClick(Sender: TObject);
    procedure btnShiftViewModeClick(Sender: TObject);
    procedure btnCommentClick(Sender: TObject);

    procedure Set_ZipTempIs(sStringIs: string);
    procedure btnFavoriteDirClick(Sender: TObject);
    procedure DelRegistryKey(const iRootKey: Cardinal; sSubKey: string);


    function ShowProperties(hWndOwner: HWND; const FileName: string): Boolean;
    procedure mnuPropertiesClick(Sender: TObject);
    procedure mnuSubPropertiesClick(Sender: TObject);
    function UserID: String;

    procedure SetBool_CheckType(bChecked: boolean);
    procedure SetBool_CheckCRC(bChecked: boolean);
    procedure SetBool_CheckAttributes(bChecked: boolean);
    procedure Refresh_ItemsOnMainList;
    procedure Refresh_ItemsOn_VirtualMainList(sPath: string;
     sArrayDirIs: array of string);
    procedure TreeView1_Change;
    procedure mnuSubMoveArchiveClick(Sender: TObject);
    function MoveFile(const sSource: string; const sTargetPath: string): boolean;
    procedure mnuSubRenameArchiveClick(Sender: TObject);
    procedure mnuSubDeleteArchiveClick(Sender: TObject);
    procedure SetBool_CheckFileToRecycleBin(bChecked: boolean);

    procedure SetBool_CheckFavoriteOpen(bChecked: boolean);
    procedure SetBool_CheckFavoriteAdd(bChecked: boolean);
    procedure SetBool_CheckFavoriteExtract(bChecked: boolean);
    procedure SetStr_FavoriteOpen(sPath: string);
    procedure SetStr_FavoriteAdd(sPath: string);
    procedure SetStr_FavoriteExtract(sPath: string);
    procedure SetStr_VirusScannerProg(sPathFName: string);
    procedure SetStr_VirusScannerPara(sStringIs: string);
    procedure SetBool_CntOf_ColumnChanged_OnMainLst(bChanged: boolean);
    procedure mnuSubByNameClick(Sender: TObject);
    procedure mnuSubByModifiedClick(Sender: TObject);
    procedure mnuSubBySizeClick(Sender: TObject);
    procedure mnuSubByRatioClick(Sender: TObject);
    procedure mnuSubByPackedClick(Sender: TObject);
    procedure mnuSubByFileTypeClick(Sender: TObject);
    procedure mnuSubByCRCClick(Sender: TObject);
    procedure mnuSubByAttributesClick(Sender: TObject);
    procedure mnuSubByPathClick(Sender: TObject);
    procedure mnuSubByDefaultClick(Sender: TObject);

    procedure SetBool_CheckTxtWithNotePad(bChecked: boolean);
    procedure SetBool_CheckHiddenFiles(bChecked: boolean);
    procedure SetBool_CheckSystemFiles(bChecked: boolean);
    procedure SetBool_RadioButtonParaFilename(bChecked: boolean);
    procedure mnuRenameClick(Sender: TObject);
    procedure mnuSubRenameClick(Sender: TObject);
    procedure mnuAddCommentClick(Sender: TObject);
    procedure smHomePageClick(Sender: TObject);
    procedure mnuRepairClick(Sender: TObject);
    function FreeDiskSpace(sDriveIs: string; var bNoErr: boolean): Int64;
    function Get_FileSize(sPathFName: string): integer;
    function FilterSysOrHidFile_And_Get_FileSize(sPathFName: string; var bSysOrHid_File_BeRemoved:
      boolean): integer;
    procedure ZipMaster1ItemProgress(Sender: TObject; Item: String;
      TotalSize: Cardinal; PerCent: Integer);
    procedure mnuSubSearchClick(Sender: TObject);
    function ChangeFolder_OnVirtualTree_After_SearchFile(sPathIs: string):
      boolean;
    procedure btnVTupClick(Sender: TObject);
    procedure btnVThomeClick(Sender: TObject);
    procedure btnVTcollapseClick(Sender: TObject);
    procedure mnuActionsClick(Sender: TObject);
    procedure mnuVirusScanClick(Sender: TObject);

    procedure Reading_Filenames_InFolder(sPathIs: string; sWildcards: string;
      var RetStrings: TStrings);
    procedure ZipMaster1Progress(Sender: TObject; ProgrType: ProgressType;
      Filename: String; FileSize: Cardinal);
    procedure VTGetNodeDataSize(Sender: TBaseVirtualTree;
      var NodeDataSize: Integer);
    procedure VTGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType;
      var CellText: WideString);
    procedure VTGetImageIndex(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Kind: TVTImageKind; Column: TColumnIndex; var Ghosted: Boolean;
      var ImageIndex: Integer);
    procedure VTHeaderClick(Sender: TVTHeader; Column: TColumnIndex;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure VTCompareNodes(Sender: TBaseVirtualTree; Node1,
      Node2: PVirtualNode; Column: TColumnIndex; var Result: Integer);

    procedure AddNodesToVirtualTree(TotalCnt: integer);
    procedure VTDblClick(Sender: TObject);
    procedure VTIncrementalSearch(Sender: TBaseVirtualTree;
      Node: PVirtualNode; const SearchText: WideString;
      var Result: Integer);
    procedure VTKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btnViewTxtClick(Sender: TObject);
    procedure btnSaveTxtToFileClick(Sender: TObject);
    procedure btnAdvancedViewClick(Sender: TObject);
    procedure btnAboutClick(Sender: TObject);
    procedure btnEditSelXMLClick(Sender: TObject);
    procedure btnTestClick(Sender: TObject);
    procedure btnRepairClick(Sender: TObject);
    procedure mnuViewXMLContentsClick(Sender: TObject);
    procedure AfterOpeningArchive_And_Do;
    //function ExecAndWait(const ExecuteFile, ParamString : string): boolean;

    procedure pmWordOnPopup(Sender: TObject);
    procedure btnHelpClick(Sender: TObject);
   { procedure VTBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellRect: TRect); }
    procedure btnHLImgFilesClick(Sender: TObject);
    procedure btnRecoverIMAGESClick(Sender: TObject);
    procedure VTBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect;
      var ContentRect: TRect);
    procedure VETreeEnumFolder(
      Sender: TCustomVirtualExplorerTree; Namespace: TNamespace;
      var AllowAsChild: Boolean);
    procedure VETreeClick(Sender: TObject);
    procedure VETreeFocusChanged(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Column: TColumnIndex);

    // ***
    //procedure ShellDir1AddFile(Sender: TObject; var SearchRec: TSearchRec;
    //  var AddFile: Boolean);
    //procedure ShellDir1Changing(Sender: TObject; Item: TListItem;
    //  Change: TItemChange; var AllowChange: Boolean);
    //procedure ShellDir1Click(Sender: TObject);
    //procedure ShellDir1DblClick(Sender: TObject);
    //procedure ShellDir1DDFileOperationExecuted(Sender: TObject;
    //  dwEffect: Integer; SourcePath, TargetPath: String);
    //procedure ShellDir1ExecFile(Sender: TObject; Item: TListItem;
    //  var AllowExec: Boolean);
    //procedure ShellDir1KeyUp(Sender: TObject; var Key: Word;
    //  Shift: TShiftState);
    //procedure cbShellDrive1Change(Sender: TObject);
    //procedure cbShellDrive1KeyUp(Sender: TObject; var Key: Word;
    //  Shift: TShiftState);

        {How to detect an inserted or removed CD-ROM or mounted network drives:
     The message WM_DEVICECHANGE is not passed to child-windows (Driveview1),
     so the media change can only be detected by the applications window:}
    //procedure WMDeviceChange(Var Message : TMessage); message WM_DEVICECHANGE;

    //***
    procedure Remove_StringGrids_And_Tabsheets;
    procedure Get_Drawing_Filename(const sSheetFile: string;
    const iSheetIndexIs: integer; const sRid: string);
    procedure Get_Images_And_Positions(const sDrawFileIs: string; const iSheetIndexIs: integer);
    function Get_Image_Filename_From_Rels(const sDrawFileIs: string;
             const sRidIs: string): string;
    procedure FreeBitmap;
    procedure pmCopyImgObjClick(Sender: TObject);
    procedure pnWorksheetsResize(Sender: TObject);
    procedure VTrAfterCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellRect: TRect);
    procedure VTrBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect;
      var ContentRect: TRect);
    procedure VTrFocusChanging(Sender: TBaseVirtualTree; OldNode,
      NewNode: PVirtualNode; OldColumn, NewColumn: TColumnIndex;
      var Allowed: Boolean);
    procedure VTrMeasureItem(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; var NodeHeight: Integer);
    procedure VTrMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure VTrMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnSaveAllToCSVClick(Sender: TObject);
    procedure memOutputMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure memOutputKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);


  private
    { Private declarations }
    //ThreadsRunning: Integer;
    //bFinished_Threads: boolean;
    //procedure ThreadDone(Sender: TObject);
    procedure SaveStatusOfProgram;
    procedure RestoreStatusOfProgram; // Restore Status of Program

    procedure Remove_TempFiles_UnderWinTemp;
    procedure OpenZipFile;
    procedure OpenedFileOnClick(Sender: TObject);
    procedure FavoriteDirOnClick(Sender: TObject);
    function Create_TempFolder_Under_WindowsTemp: string;
    procedure Initial_VT;
    procedure RestoreRecentUnzipFilenameToMenu;
    procedure CheckRegisteredUser;
    procedure CheckZipFilesAssociated;
    //procedure Set_MainList_ToUse_SystemIcons;
    procedure SaveRegUser(Sender: TObject; sHeader: string;
     sRegUser: string; iPointer: integer);

    procedure Zipping;

    procedure WMDropFiles(var Msg: TMessage); message WM_DropFiles;

    procedure AddNode_ToTreeView;
    procedure ChangeFolder_OnVirtualMainList_Click(sDir: string);
    procedure DeleteFiles_InArchive;
    procedure DeleteFiles_InArchive_OnVirtualMainList;
    procedure AddAllFiles_UnderDir_ToFSpecArgs_BeforeDel(sDir: string);
    function GetSelectedPath_FromVirtualTreeView: string;
    procedure Remove_DeletedNode_InVirtualTree(sArrayDirIs: array of string);
    procedure Extract_SelectedDir_OnVirtualMainList(sDir: string;
      var sPathFNames: TStrings);
    procedure ExtactFile_Before_ShowProperties(sFileName: string);
    procedure Who_Installed_Software_WriteTo_Registry;

    function GetInputOfFilename(sFilenameIs: string): string;

    procedure Sort_SubMenuClick(Sender: TObject);

    procedure Update_lblCurDir(sDirIs: string);
    function LevelOf_FreeDiskSpace_Extract(sDriveIs: string): integer;
    function Get_Predict_Compressed_FileSize(sPathFName: string): integer;
    procedure Set_MainListView_ToUse_SystemIcons;
    procedure Insert_HistoryOpenWithApp_ToMenu;
    procedure RecentOpenWithAppToReg(const sPathFilenameIs: string);
    procedure RecentOpenWithAppFromReg(var AFiles: TStrings);
    procedure mSubOpenWith_ThisApp_OnClick(Sender: TObject);
    procedure Update_FDataList;
    function MBToUnicode(sFrom: string): WideString;
    procedure SGDrawCell(Sender: TObject; ACol,
              ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure SGMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure SGMouseDown(Sender: TObject; Button: TMouseButton;
              Shift: TShiftState; X, Y: Integer);

    procedure VTrGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType;
      var CellText: WideString);
  public
    { Public declarations }
    TotalSize1, TotalProgress1, TotalSize2, TotalProgress2: Int64;
    function Get_SysIconIndex_Of_Given_FileExt(sWhatFileExt: string;
     var sFileTypeIs: string): integer;
    function GetAttributes(iFileAttri: Cardinal): string;
    procedure ClearColumnsImage_OnlvMain;
    function Get_ColIndex_ToSort_FromMenuSortBy: integer;

    procedure Output_Txt_Contents_Of_XML_From_Docx;
    procedure Output_Txt_Contents_Of_XML_From_Pptx;
    procedure Output_Txt_Contents_Of_XML_From_Xlsx;
    procedure Output_Txt_Contents_Of_XML_From_Xlsx_Speed;
    function Xlsx_To_CSV(const sSheetName: string; const sPFilenameIs: string): boolean;
    function Xlsx_To_CSV_Speed(iSheetIndex: integer; bSaveAll: boolean): boolean;
    procedure Get_SharedStrings_From_Xlsx;
    procedure Get_SheetNames_From_Xlsx;
    procedure Release_Memory_Of_Xlsx_File_Open;
  end;

//{$DEFINE IS_SHAREWARE} // Main, About Forms
//{$DEFINE IS_DEMO} // Main, EditXML Forms
  {$DEFINE SHAREWARE_FOR_S2SERVICES} // Main, reabout
  {$DEFINE CRC_CHECKSUM}             // Main

 { TLVColumnEx = packed record
    mask: UINT;
    fmt: Integer;
    cx: Integer;
    pszText: PAnsiChar;
    cchTextMax: Integer;
    iSubItem: Integer;
    iImage: integer; // New
    iOrder: integer; // New
  end; }



const
  XcProgramName = 'Corrupt Office 2007 Extractor v3.4';
  XcIniRegFile = 'reginfo.ini';
  XcExeProgFilename = 'hahazipw.exe';
  XcContextMenuProgName = 'contmenu.dll';
  XcIniHisFile = 'hisfile.ini';

  XcSubKeyIs = 'SOFTWARE\Ccy\Corrupt Office 2007 Extractor v3.4';
  XcNoComment = 'MakeNoComment';
  XcZipTemp_WinTemp = 'WinTemp';
  XcZipTemp_Current = 'Current';
  XcZipTemp_Target = 'Target';
  XcInstalledBy = 'Installed By';
  XcDefault = 'Default';
  XcContentsOfDocx = 'Contents of the ';

  // --- Begin important factors for information of registration

  XcRegUserName = 'RUN'; // Trial

  // --- End important factors for information of registration

var
  frmMain: TfrmMain;
  XmlParser: TXmlParser;
  DrawXmlParser: TXmlParser;
  ImagesOnSheet: array of TImagesOnSheet;
  AImage: array of TImage;
  AHeightOfRow: array of THeightOfRow;
  ABitmap, BBitmap, TeBitmap, ClipBitmap: TBitmap;

  XbOutputTxtContentsFile: boolean = False;
  XbOutputXmlFile: boolean = False;
  XbIsDoUnzip: boolean = False;
  XsExtractToThisPath: string = '';
  XslOfficeDoc: TStringList;
  XslSheetsFilenames: TStringList;
  XslSharedStr: TStringList;
  XslAlphabet: TStringList;
  XslRID: TStringList;
  AThisColW: array of integer;

  XsAppPath: string;
  XsWinPath: string;
  bExtractedFiles: boolean = False;
  sSystemDirIs: string;
  xExtractedFolder: string;
  XbHaveSelectedFiles: boolean = False; // determine user selected files through Context Menu or not
  chPreSearch: Char;
  iStartItemSearch: integer;
  //bCancelExtract_OnDirectory: boolean = True;
  bCancelInputComment_OnInputMsgBox: boolean = True;
  bCancelAddFiles: boolean = True;
  bCancelOpenZip: boolean = True;
  sErrorMessage: string = '';
  sComment: string = '';
  sZipTempFolder: string = '';

  bCheckType: boolean = False;
  bCheckCRC: boolean = False;
  bCheckAttributes: boolean = False;
  bDeletedFiles_To_RecycleBin: boolean = True;

  bFavorite_OpenFolder: boolean = False;
  bFavorite_AddFolder: boolean = False;
  bFavorite_ExtractFolder: boolean = False;
  sFavorite_OpenFolder: string;
  sFavorite_AddFolder: string;
  sFavorite_ExtractFolder: string;
  sVirusScanner_PathFName: string;
  sVirusScanner_Para: string;
  bCol_AddedOrRemoved_OnMainLst: boolean = False;
  bOpen_TextFile_WithNotePad: boolean = True;
  bDontShow_HiddenFiles: boolean = False;
  bDontShow_SystemFiles: boolean = False;
  bParaFilename: boolean = False;

  bWin9598: boolean = False;
  bWinVista: boolean = False;
  iColWidthName, iColWidthModified, iColWidthSize, iColWidthRatio, iColWidthPacked,
  iColWidthFileType, iColWidthCRC, iColWidthAttributes, iColWidthPath: integer;
  bCheckMainLstAutoSize: boolean = False;
  sPreExt: string;
  iPreImgI: integer;
  //ilMainList: TimageList;
  //AIcon: TIcon;
  slZipStrings: TStrings;
  iIndex_Ascend_SysIcons, iIndex_Descend_SysIcons: Cardinal;
  FDataList: TObjectList;
  ped : TVTPropEditData;

  mnuOpenWith_ThisApp: TMenuItem;
  mSubOpenWith_ThisApp: array of TMenuItem;



  // --- Begin important factors for information of registration


  XbRegisteredUser: boolean = False;

  // --- End important factors for information of registration



implementation

uses
  EditXML, ShellApi, OptAddFl, reabout, span, ShlObj, MsgOut, settings,
  hrglass, InputMsgBox, Open_Add, Extract_To, Repair, SearchBox, OpenDir,
  LbCipher, LbString;


const
  iFixedRowHeight: Shortint = 18;
  iRowHeight: Shortint = 15;
  cKeyNamechkAss = 'CheckAssociate';

  cMainTop = 'MainTop';
  cMainLeft = 'MainLeft';
  cMainWidth = 'MainWidth';
  cMainHeight = 'MainHeight';
  cMainWindowState = 'MainWindowState';

  cWidthName = 'WidthName';
  cWidthModified = 'WidthModified';
  cWidthSize = 'WidthSize';
  cWidthRatio = 'WidthRatio';
  cWidthPacked = 'WidthPacked';
  cWidthFileType = 'WidthFileType';
  cWidthCRC = 'WidthCRC';
  cWidthAttributes = 'WidthAttributes';
  cWidthPath = 'WidthPath';

  cWidthBrowseList = 'WidthBrowseList';
 // cUpDownBtnViewStyle = 'UpDownBtnViewStyle';
  cUpDownBtnAdvancedView = 'UpDownBtnAdvancedView';
  cHeightPnM = 'cHeightPnM';

  cCountOfRecentUnzipFiles = 8; // Lite version is 3. 10;
  cCountOfRecentOpenWithApp = 8; 
  cFavouriteFolders = '\FavouriteFolders';
  cMainMenuSort = 'MainMenuSort';
var

  bResizingCols: boolean = False;
  mnuNewMake: array of TMenuItem;
  mnuFavoriteDir: array of TMenuItem;
  sSubKeyRecentUnzip: string = XcSubKeyIs + '\HistoryUnzipFiles';
  sSubKeyRecentOpenWithApp: string = XcSubKeyIs + '\HistoryOpenWithApp';
  iCntFavorite: shortint = 3;
  ColumnToSort: integer;
  iPreImgIndex: Cardinal = 0;
  iStart_HOfRow: array of integer;

  bTerminatedProcess: boolean = False;
  ArrowUP, ArrowDown: HBitmap;

  FBufferSize: integer;       // for open xlsx file
  FBuffer: PChar;             // for open xlsx file
  CurStart: PChar;            // for open xlsx file
  //APrevColor: TColor;
{$R *.dfm}

{--------------------------------------------------------------------------
                              TVTPropEditData
   --------------------------------------------------------------------------}
constructor TVTPropEditData.Create;
begin
  inherited Create;
  FCaption := TStringList.Create;
  FModified := TStringList.Create;
  FSize := TStringList.Create;
  FRatio := TStringList.Create;
  FColPacked := TStringList.Create;
  FFileType := TStringList.Create;
  FCRC := TStringList.Create;
  FAttributes := TStringList.Create;
  FPath := TStringList.Create;

 // FCaptions.CommaText := s;  // string to list
 // FHeading := FCaptions[0];  // 1st element is caption
 // FCaptions.Delete(0);       // can now delete it
end;

destructor TVTPropEditData.Destroy;
begin
  FCaption.Free;
  FModified.Free;
  FSize.Free;
  FRatio.Free;
  FColPacked.Free;
  FFileType.Free;
  FCRC.Free;
  FAttributes.Free;
  FPath.Free;
  inherited Destroy;
end;


procedure TfrmMain.FormCreate(Sender: TObject);
var
  buffer: pchar;
  cRetWinLen, cdRetValue: Cardinal;
  OSInfo: TOSVersionInfo;
  iMajorVer: Shortint;
  iMinorVer: Shortint;
  StartDir : String;

  i: integer;
begin
  // DriveView must be assigned a dir. Otherwise, DriveView will cause an error after removed USB drive. Dir not found
 { IF Not Assigned(DriveView1.Selected) Then
  Begin
    GetDir(0, StartDir);
    DriveView1.Directory := StartDir;
  End;

  IF Assigned(DriveView1.Selected) Then
  DriveView1.Selected.Expand(False); }


  {$IFDEF IS_DEMO}
  Self.Caption := XcProgramName + ' - Demo';
  memOutput.ReadOnly := True;
  pmCopyImgObj.Enabled := False;  // pmSG popup menu
  {$ELSE}
  Self.Caption := XcProgramName;
  {$ENDIF}
  ZipFName.Caption := '';
  edShadow.Text := '';
  lblOpenFileIs.Caption := '';
  Update_lblCurDir('');
  memOutput.Clear;


  //lMainList := TImageList.Create(Self);
  //AIcon := TIcon.Create;
  mnuOpenWith_ThisApp := TMenuItem.Create(PopupMenu1);
  mnuOpenWith_ThisApp.Caption := 'Open With This';
  //pnSep.Caption := ''; // this panel provides successful Maximize of main screen


  ABitmap := TBitmap.Create;
  BBitmap := TBitmap.Create;
  TeBitmap := TBitmap.Create;
  ClipBitmap := TBitmap.Create;
  
  XslRID := TStringList.Create; // for worksheet name
  XslOfficeDoc := TStringList.Create; // for pptx
  XslSheetsFilenames := TStringList.Create;  // sheets filenames
  XslSharedStr := TStringList.Create; // for xlsx
  XslAlphabet := TStringList.Create; // store Excel Col header

  XslAlphabet.Append('-1'); // not use
  XslAlphabet.Append('A');
  XslAlphabet.Append('B');
  XslAlphabet.Append('C');
  XslAlphabet.Append('D');
  XslAlphabet.Append('E');
  XslAlphabet.Append('F');
  XslAlphabet.Append('G');
  XslAlphabet.Append('H');
  XslAlphabet.Append('I');
  XslAlphabet.Append('J');

  XslAlphabet.Append('K');
  XslAlphabet.Append('L');
  XslAlphabet.Append('M');
  XslAlphabet.Append('N');
  XslAlphabet.Append('O');
  XslAlphabet.Append('P');
  XslAlphabet.Append('Q');
  XslAlphabet.Append('R');
  XslAlphabet.Append('S');
  XslAlphabet.Append('T');

  XslAlphabet.Append('U');
  XslAlphabet.Append('V');
  XslAlphabet.Append('W');
  XslAlphabet.Append('X');
  XslAlphabet.Append('Y');
  XslAlphabet.Append('Z'); // 26



  slZipStrings := TStringList.Create; // for zipping data
  FDataList := TObjectList.Create;
  if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then // 95/98, ME
      VT.Font.Charset := GetDefFontCharSet; // VT uses wide string type, only correct charater set can bring proper Chinese words

  {$IFDEF IS_SHAREWARE}
  btnShowVirtual.Visible := True;
  btnShiftViewMode.Visible := True;
  {$ELSE}
  btnShowVirtual.Visible := False;
  btnShiftViewMode.Visible := False;
  {$ENDIF}

  ZipMaster1.AddCompLevel := 6;
  SetLength(mnuNewMake, cCountOfRecentUnzipFiles);

  //Set_ZipDllOptions;

  // get application path
  XsAppPath := ExtractFileDir(Application.ExeName);
  XsAppPath := IncludeTrailingPathDelimiter(XsAppPath);

  // get Windows path
  getmem(buffer,255);
  cRetWinLen := getwindowsdirectory(buffer,255);
  if cRetWinLen > 0 then begin
      XsWinPath := copy(buffer, 1, length(buffer));
      XsWinPath := IncludeTrailingPathDelimiter(XsWinPath);
  end
  else begin
      MessageDlg('Failing to get folder name of Windows. ' +
      XcProgramName + ' cannot get this folder name, it does not know where to ' +
      'create temporary folder under Windows. This problem is rare. If you still have ' +
      'problem, contact Author. Press OK to quit.', mtError, [mbOK], 0);
      Application.Terminate;
  end;
  freemem(buffer,255);

  // get directory of system
  getmem(buffer,255);
  cdRetValue := getsystemdirectory(buffer,255);
  if cdRetValue > 0 then begin
      sSystemDirIs := copy(buffer, 1, length(buffer));
      sSystemDirIs := IncludeTrailingPathDelimiter(sSystemDirIs);
  end
  else begin
      MessageDlg('Failing to get folder name of Windows System.',
      mtInformation, [mbOK], 0);
  end;
  freemem(buffer,255);

  OSInfo.dwOSVersionInfoSize := SizeOf(OSInfo);
  //Get the Windows version
  if GetVersionEx(OSInfo) then begin
      if OSInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS then  // 95/98, ME
          bWin9598 := True;


      iMajorVer := OSInfo.dwMajorVersion; // decision of NT4 or 2000
      iMinorVer := OSInfo.dwMinorVersion;  // Windows NT versions:
                                          // 3.51.1057    Windows NT 3.51
                                          // 4.0.1381     Windows NT 4.0
                                          // 5.0.2195     Windows 2000
                                          // 5.01.2600    Windows XP
                                          // 5.02.xxxx    Windows Server 2003
                                          // 6.           Windows Vista

     if iMajorVer >= 6 then begin // is Vista ?
         bWinVista := True;
         VT.Images := nil;
         VT.StateImages := ilSysIcons; // Sys images only show in this mode under Vista
     end;

  end;

  try
    ZipMaster1.Load_Zip_Dll;
    ZipMaster1.Load_Unz_Dll;
  except
    MessageDlg('Fatal error to load Dll file. Program will be terminated', mtError, [mbOK], 0);
    Application.Terminate;
  end;
  XmlParser := TXmlParser.Create;
  DrawXmlParser := TXmlParser.Create;
  DragAcceptFiles(Handle, True);
  RestoreStatusOfProgram;
  Initial_VT;
  RestoreRecentUnzipFilenameToMenu;
 // CheckRegisteredUser;
 // CheckZipFilesAssociated;
  //Set_MainList_ToUse_SystemIcons;
  Set_MainListView_ToUse_SystemIcons;
  SetStatusOfPopupUnzip(nil);
  Who_Installed_Software_WriteTo_Registry;

end;


procedure TfrmMain.SetZipFile(FName: String; bVerbose_List: boolean);
begin
   btnSaveTxtToFile.Enabled := False;
   btnSaveAllToCSV.Enabled := False;
   memOutput.Clear;

   ZipFName.Caption:=FName;
   if bVerbose_List then
       ZipMaster1.ZipFileName_VerboseList := FName
   else begin
       ZipMaster1.ZipFilename := FName;
      { ListView1.Items.BeginUpdate;
       ListView1.Items.Clear; // <= faster than ListView1.Clear
       ListView1.Items.EndUpdate; }
   end;
   ZipMaster1.Password := ''; // clear preserved passowrd
   SetCurrentDir(ExtractFileDir(FName));

end;

procedure TfrmMain.WriteIniToReg(sSubKey: string; sValueName: string; sValue: string);
begin
  with TRegistry.Create do
  try
    RootKey := HKEY_CURRENT_USER;
    if OpenKey(sSubKey, True) then begin
        WriteString(sValueName, sValue);
        CloseKey;
    end;
  finally
    Free;
  end;
end;

function TfrmMain.ReadIniFromReg(sSubKey: string; sValueName: string): string;
begin
  Result := '';

  with TRegistry.Create do
  try
    RootKey := HKEY_CURRENT_USER;
    if OpenKey(sSubKey, False) then begin
        Result := ReadString(sValueName);

        CloseKey;
    end;
  finally
    Free;
  end;
end;

procedure TfrmMain.WriteRegLocalMachine(sSubKey: string; sValueName: string; sValue: string);
begin
  with TRegistry.Create do
  try
    RootKey := HKEY_LOCAL_MACHINE;
    if OpenKey(sSubKey, True) then begin
        WriteString(sValueName, sValue);
        CloseKey;
    end;
  finally
    Free;
  end;
end;

function TfrmMain.ReadRegLocalMachine(sSubKey: string; sValueName: string): string;
begin
  Result := '';

  with TRegistry.Create do
  try
    RootKey := HKEY_LOCAL_MACHINE;
    if OpenKey(sSubKey, False) then begin
        Result := ReadString(sValueName);

        CloseKey;
    end;
  finally
    Free;
  end;
end;

procedure TfrmMain.SetStatusOfchkAutoSizeToVar(const bAutoSizeChecked: boolean);
begin
  bCheckMainLstAutoSize := bAutoSizeChecked;
end;

procedure TfrmMain.ResizeWidthOfMainListColumns;
var
  i: integer;
  iCntOfCols: shortint;
begin
  try
    bResizingCols := True; // trap
    iCntOfCols := VT.Header.Columns.Count;

    if bCheckMainLstAutoSize and (VT.TotalCount > 0) then begin
        Screen.Cursor := crHourGlass; // useful when over hundred thousand of nodes
        try
          VT.Header.AutoFitColumns(True);
        finally
          Screen.Cursor := crDefault;
        end;
    end
    else begin
        for i := iCntOfCols -1 downto 0 do begin
          if VT.Header.Columns[i].Text = 'Name' then
              VT.Header.Columns[i].Width := iColWidthName
          else if VT.Header.Columns[i].Text = 'Modified' then
              VT.Header.Columns[i].Width := iColWidthModified
          else if VT.Header.Columns[i].Text = 'Size' then
              VT.Header.Columns[i].Width := iColWidthSize
          else if VT.Header.Columns[i].Text = 'Ratio' then
              VT.Header.Columns[i].Width := iColWidthRatio
          else if VT.Header.Columns[i].Text = 'Packed' then
              VT.Header.Columns[i].Width := iColWidthPacked
          else if VT.Header.Columns[i].Text = 'File Type' then
              VT.Header.Columns[i].Width := iColWidthFileType
          else if VT.Header.Columns[i].Text = 'CRC' then
              VT.Header.Columns[i].Width := iColWidthCRC
          else if VT.Header.Columns[i].Text = 'Attributes' then
              VT.Header.Columns[i].Width := iColWidthAttributes
          else if VT.Header.Columns[i].Text = 'Path' then
              VT.Header.Columns[i].Width := iColWidthPath;

        end;
    end;
  finally
    bResizingCols := False;
  end;
end;

procedure TfrmMain.SetWidthOfMainListColumns;
var
  sWidthName, sWidthModified, sWidthSize, sWidthRatio, sWidthPacked,
  sWidthFileType, sWidthCRC, sWidthAttributes, sWidthPath: string;
  iWidthName, iWidthModified, iWidthSize, iWidthRatio, iWidthPacked,
  iWidthFileType, iWidthCRC, iWidthAttributes, iWidthPath: integer;
  i: integer;
begin
  iColWidthName := 160; iColWidthModified := 140;
  iColWidthSize := 100; iColWidthRatio := 80;
  iColWidthPacked := 90; iColWidthFileType := 100;
  iColWidthCRC := 80; iColWidthAttributes := 70;
  iColWidthPath := 100;
  // Width of columns on main list
  sWidthName := ReadIniFromReg(XcSubKeyIs, cWidthName);
  if sWidthName <> '' then begin
      iWidthName := strtoint(sWidthName);
      if (iWidthName > 10) and ( iWidthName <= VT.Width) then
          iColWidthName := iWidthName;

  end;

  sWidthModified := ReadIniFromReg(XcSubKeyIs, cWidthModified);
  if sWidthModified <> '' then begin
      iWidthModified := strtoint(sWidthModified);
      if (iWidthModified > 10) and ( iWidthModified <= VT.Width) then
          iColWidthModified := iWidthModified;

  end;
  sWidthSize := ReadIniFromReg(XcSubKeyIs, cWidthSize);
  if sWidthSize <> '' then begin
      iWidthSize := strtoint(sWidthSize);
      if (iWidthSize > 10) and ( iWidthSize <= VT.Width) then
          iColWidthSize := iWidthSize;

  end;
  sWidthRatio := ReadIniFromReg(XcSubKeyIs, cWidthRatio);
  if sWidthRatio <> '' then begin
      iWidthRatio := strtoint(sWidthRatio);
      if (iWidthRatio > 10) and ( iWidthRatio <= VT.Width) then
          iColWidthRatio := iWidthRatio;

  end;
  sWidthPacked := ReadIniFromReg(XcSubKeyIs, cWidthPacked);
  if sWidthPacked <> '' then begin
      iWidthPacked := strtoint(sWidthPacked);
      if (iWidthPacked > 10) and ( iWidthPacked <= VT.Width) then
          iColWidthPacked := iWidthPacked;

  end;
  sWidthFileType := ReadIniFromReg(XcSubKeyIs, cWidthFileType);
  if sWidthFileType <> '' then begin
      iWidthFileType := strtoint(sWidthFileType);
      if (iWidthFileType > 10) and ( iWidthFileType <= VT.Width) then
          iColWidthFileType := iWidthFileType;

  end;
  sWidthCRC := ReadIniFromReg(XcSubKeyIs, cWidthCRC);
  if sWidthCRC <> '' then begin
      iWidthCRC := strtoint(sWidthCRC);
      if (iWidthCRC > 5) and ( iWidthCRC <= VT.Width) then
          iColWidthCRC := iWidthCRC;

  end;
  sWidthAttributes := ReadIniFromReg(XcSubKeyIs, cWidthAttributes);
  if sWidthAttributes <> '' then begin
      iWidthAttributes := strtoint(sWidthAttributes);
      if (iWidthAttributes > 10) and ( iWidthAttributes <= VT.Width) then
          iColWidthAttributes := iWidthAttributes;

  end;
  sWidthPath := ReadIniFromReg(XcSubKeyIs, cWidthPath);
  if sWidthPath <> '' then begin
      iWidthPath := strtoint(sWidthPath);
      if (iWidthPath > 10) and ( iWidthPath <= (VT.Width)) then
          iColWidthPath := iWidthPath;

  end;

  for i := VT.Header.Columns.Count -1 downto 0 do begin
          if VT.Header.Columns[i].Text = 'Name' then
              VT.Header.Columns[i].Width := iColWidthName
          else if VT.Header.Columns[i].Text = 'Modified' then
              VT.Header.Columns[i].Width := iColWidthModified
          else if VT.Header.Columns[i].Text = 'Size' then
              VT.Header.Columns[i].Width := iColWidthSize
          else if VT.Header.Columns[i].Text = 'Ratio' then
              VT.Header.Columns[i].Width := iColWidthRatio
          else if VT.Header.Columns[i].Text = 'Packed' then
              VT.Header.Columns[i].Width := iColWidthPacked
          else if VT.Header.Columns[i].Text = 'File Type' then
              VT.Header.Columns[i].Width := iColWidthFileType
          else if VT.Header.Columns[i].Text = 'CRC' then
              VT.Header.Columns[i].Width := iColWidthCRC
          else if VT.Header.Columns[i].Text = 'Attributes' then
              VT.Header.Columns[i].Width := iColWidthAttributes
          else if VT.Header.Columns[i].Text = 'Path' then
              VT.Header.Columns[i].Width := iColWidthPath;

  end;
end;

procedure TfrmMain.RegisterFileType(cMyExt, cMyFileType, cMyDescription,
 ExeName: string; IcoIndex: integer; DoUpdate: boolean);
// cMyFileType is a value. e.g SeeSeeWhy HaHaZip
// cMyExt is a key. e.g .zip under Classes root
// ExeName is program name 
var
  Reg: TRegistry;
  bError_Found: boolean;
  wRetAns: Word;
  sSubWd: string;
begin
  bError_Found := False;
  sSubWd := '';
  Reg := TRegistry.Create;
  try
    if bWinVista = True then begin
        Reg.RootKey := HKEY_CURRENT_USER;
        sSubWd := 'Software\classes\';
    end
    else 
        Reg.RootKey := HKEY_CLASSES_ROOT;

    if Reg.OpenKey(sSubWd + cMyExt, True) then begin
        // Write my file type to it.
        // This adds HKEY_CLASSES_ROOT\.abc\(Default) = 'Project1.FileType'
        Reg.WriteString('', cMyFileType); // cMyFileType is a value. e.g SeeSeeWhy HaHaZip
        Reg.CloseKey;
    end
    else
        bError_Found := True;

    // Now create an association for that file type
    if Reg.OpenKey(sSubWd + cMyFileType, True) then begin // Using manifest will fail to open Registry
        // This adds HKEY_CLASSES_ROOT\Project1.FileType\(Default)
        //   = 'Project1 File'
        // This is what you see in the file type description for
        // the a file's properties.
        Reg.WriteString('', cMyDescription); // Using XP compatible mode will fail to write data to Registry
        Reg.CloseKey;    // Now write the default icon for my file type
    end
    else
        bError_Found := True;

    // This adds HKEY_CLASSES_ROOT\Project1.FileType\DefaultIcon
    //  \(Default) = 'Application Dir\Project1.exe,0'
    if Reg.OpenKey(sSubWd + cMyFileType + '\DefaultIcon', True) then begin
        Reg.WriteString('', ExeName + ',' + IntToStr(IcoIndex));
        Reg.CloseKey;
    end
    else
        bError_Found := True;

    // Now write the open action in explorer
    if Reg.OpenKey(sSubWd + cMyFileType + '\Shell\Open', True) then begin
        Reg.WriteString('', '&Open');
        Reg.CloseKey;
    end
    else
        bError_Found := True;

    // Write what application to open it with
    // This adds HKEY_CLASSES_ROOT\Project1.FileType\Shell\Open\Command
    //  (Default) = '"Application Dir\Project1.exe" "%1"'
    // Your application must scan the command line parameters
    // to see what file was passed to it.
    if Reg.OpenKey(sSubWd + cMyFileType + '\Shell\Open\Command', True) then begin
        Reg.WriteString('', '"' + ExeName + '" "%1"');
        Reg.CloseKey;
    end
    else
        bError_Found := True;

    // Finally, we want the Windows Explorer to realize we added
    // our file type by using the SHChangeNotify API.
    if DoUpdate then
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nil, nil);

  finally
    Reg.Free;
  end;

  if bError_Found then begin
      beep;
      with TRegistry.Create do

      try
        RootKey := HKEY_CURRENT_USER;
        wRetAns := MessageDlg('Failing to update registry. If under Windows NT, ' +
                   'only Administrator or Super Users have authorities to execute ' +
                   'this process. Do you still want to have same check next time?',
                   mtConfirmation, [mbYes, mbNo], 0); // = mrYes

        if OpenKey(XcSubKeyIs, True) then begin
            if wRetAns = mrYes then
                WriteString(cKeyNamechkAss, 'True')
            else
                WriteString(cKeyNamechkAss, 'False');

            CloseKey;
        end;
      finally
        Free;
      end;
  end;
end;

procedure TfrmMain.ReadFilesOrFolders(sPathIs: string; sWildcards: string);
var
  sr: TSearchRec;
  iFileAttrs: Integer;
  //iRetValue: Integer;
  iCountOfFiles: integer;
begin

  if not DirectoryExists(sPathIs) then Exit;

  if (length(sPathIs) = 3) and (copy(sPathIs, 2, 2) = ':\') and
  (AnsiLowercase( copy(sPathIs, 1, 1) ) <> 'a') then begin
      beep;
      if MessageDlg('Do you really want to add the whole drive ' +
      sPathIs + '? It is not a good idea to add the drive which stored ' +
      'many files. Especially you do not have enough free disk space for ' +
      'temporary files on drive, it will cause no response to application.',
      mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
          exit;

  end;

  if sPathIs <> '' then
      sPathIs := IncludeTrailingPathDelimiter(sPathIs);

  iFileAttrs := faAnyFile;  //faArchive;   //archives only
  if FindFirst(sPathIs + sWildcards, iFileAttrs, sr) = 0 then begin
      iCountOfFiles := 0;
      repeat
        // if ((sr.Attr and iFileAttrs) = sr.Attr) and // (SearchRec.Attr and faHidden) <> 0 mean True, = 0 means False
        if (sr.Name <> '.') and (sr.Name <> '..') then begin
            try
              if DirectoryExists(sPathIs + sr.Name) then begin
                  ReadFilesOrFolders(sPathIs + sr.Name, sWildcards);
              end
              else begin
                  slZipStrings.Append(sPathIs + sr.Name);
                  iCountOfFiles := iCountOfFiles + 1;
                  if (iCountOfFiles mod 30) = 0 then begin
                      iCountOfFiles := 0;
                      Application.ProcessMessages;
                  end;
              end
            except

            end;
        end;
      until FindNext(sr) <> 0;
      FindClose(sr);
  end

end;

procedure TfrmMain.BeforeExtracting(const sToPath: string; const bWithDir,
 bAlwaysOverwrite, bAllFiles: boolean);
var
  i, j, iCntCol: integer;
  sFileNameIs, sZipFolder, sExt, sTem: string;
  sCaptionIs, sPath, sXPath, sDoc, sAFile, sBFile: string;
  sStrings: TStrings;
  Node: PVirtualNode;
  Data: PEntry;
  sT, sErrorMessage: string;
begin
  if (ZipFName.Caption = '') or (sToPath = '') then begin
      beep;
       ShowMessage('Err: Failing to get archive name or target ' +
       'directory name.');
      exit;
  end;


  if bWithDir then
      RadioExtractWithDir.ItemIndex := 1
  else
      RadioExtractWithDir.ItemIndex := 0;

  if bAlwaysOverwrite then
      RadioOverwrite.ItemIndex := 1
  else
      RadioOverwrite.ItemIndex := 0;
      
  // Is the directory not exit?
  if not DirectoryExists(sToPath) then
      ForceDirectories(sToPath);  // create directory

  iCntCol := ListView1.Columns.Count;
  //ZipMaster1.PasswordReqCount := 3; // important if show password diaglog box
  if bAllFiles then begin
      StartUnzip(False, ZipFName.Caption, sToPath, nil, False);
  end
  else begin // selected files only

      XslOfficeDoc.Clear;
      XslSheetsFilenames.Clear;

      if btnShowVirtual.Down then begin // Virtual mode
          try
            sStrings := TStringList.Create;
            i := 0;
            repeat
              sZipFolder := '';
              //Is caption selected?;
              if ListView1.Items[i].Selected then begin
                  sCaptionIs := ListView1.Items[i].Caption; // File name
                  sCaptionIs := AnsiReplaceStr(sCaptionIs, '*', ''); // remove password char

                  if sCaptionIs = '..' then // skip up one level folder

                  else if (ListView1.Items[i].SubItems.Strings[0] = '') and
                  (ListView1.Items[i].SubItems.Strings[1] = '') and
                  btnShowVirtual.Down then begin // folder found
                      sPath := GetSelectedPath_FromVirtualTreeView;
                      sPath := sPath + sCaptionIs + '\';
                      Extract_SelectedDir_OnVirtualMainList(sPath, sStrings);
                  end
                  else begin // a file item
                      sZipFolder := ListView1.Items[i].SubItems.Strings[iCntCol -2]; // get path name
                      if sZipFolder <> '' then begin
                          if copy(sZipFolder, length(sZipFolder), 1) <> '\' then
                              sZipFolder := sZipFolder + '\';

                      end;

                      sStrings.Append(sZipFolder + sCaptionIs);
                  end;
              end;
              i := i + 1;
            until i >= ListView1.Items.Count;

            if sStrings <> nil then
                StartUnzip(False, ZipFName.Caption, sToPath, sStrings, True);

          finally
            sStrings.Free;
          end;
      end
      else begin // VT List View
          try
            sStrings := TStringList.Create;

            sExt := AnsiLowercase( ExtractFileExt(ZipFName.Caption) );

            if (sExt = '.docx') then begin // Word
                sAFile := 'word\document.xml';
                sBFile := 'word\_Rels\document.xml.rels';


                sStrings.Append(sAFile);
                sStrings.Append(sBFile);
                sXPath := IncludeTrailingPathDelimiter(sToPath);

                sDoc := sXPath + sAFile;
                if FileExists(sDoc) then begin
                    if DeleteFile(sDoc) = False then begin
                        ShowMessage('Err: Failing to delete the existed file: ' +
                        sDoc + '. Extracting has been terminated.');
                        exit;
                    end;
                end;

                sDoc := sXPath + sBFile;
                if FileExists(sDoc) then begin
                    if DeleteFile(sDoc) = False then begin
                        ShowMessage('Err: Failing to delete the existed file: ' +
                        sDoc + '. Extracting has been terminated.');
                        exit;
                    end;
                end;

            end
            else if (sExt = '.pptx') then begin // Powerpoint

                for i := 0 to ZipMaster1.ZipContents.Count -1 do begin
                  with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
                    sZipFolder := ExtractFilePath(FileName);
                    if (sZipFolder <> '') then
                        sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);

                    sFileNameIs := ExtractFileName(FileName);
                    sFileNameIs := AnsiReplaceStr(sFileNameIs, '*', ''); // remove password char

                    sTem := Ansilowercase(sZipFolder);

                    if sTem = 'ppt\slides\' then 
                        sStrings.Append(sZipFolder + sFileNameIs);

                  end;
                end;

                sErrorMessage := '';

                if sStrings.Count > 0 then begin
                    For j := 0 to sStrings.Count -1 do begin
                      sT := sStrings.Strings[j];

                      sXPath := IncludeTrailingPathDelimiter(sToPath);
                      sDoc := sXPath + sT;

                      if FileExists(sDoc) then begin
                          if DeleteFile(sDoc) = False then begin
                              sErrorMessage := sErrorMessage +
                              'Err: Failing to delete the existed file: ' +
                              sDoc + #13#10;
                          end;
                      end;
                    end;
                end;


                if sErrorMessage <> '' then begin
                    PopUp_ErrorMessage(sErrorMessage + 'Extracting has been terminated.');
                    exit;
                end
                else
                    XslOfficeDoc.Assign(sStrings);
                    
            end
            else if (sExt = '.xlsx') then begin // Excel

                for i := 0 to ZipMaster1.ZipContents.Count -1 do begin
                  with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
                    sZipFolder := ExtractFilePath(FileName);
                    if (sZipFolder <> '') then
                        sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);

                    sFileNameIs := ExtractFileName(FileName);
                    sFileNameIs := AnsiReplaceStr(sFileNameIs, '*', ''); // remove password char

                    sTem := Ansilowercase(sZipFolder);

                    if sTem = 'xl\worksheets\' then begin
                        sStrings.Append(sZipFolder + sFileNameIs);
                        XslSheetsFilenames.Append(sZipFolder + sFileNameIs);  // store sheetnames
                    end
                    else if sTem = 'xl\worksheets\_rels\' then
                        sStrings.Append(sZipFolder + sFileNameIs)
                    else if sTem = 'xl\media\' then
                        sStrings.Append(sZipFolder + sFileNameIs)
                    else if sTem = 'xl\drawings\' then
                        sStrings.Append(sZipFolder + sFileNameIs)
                    else if sTem = 'xl\drawings\_rels\' then
                        sStrings.Append(sZipFolder + sFileNameIs);


                  end;
                end;



                sAFile := 'xl\sharedStrings.xml';
                sStrings.Append(sAFile);

                sAFile := 'xl\workbook.xml';
                sStrings.Append(sAFile);

                sErrorMessage := '';

                if sStrings.Count > 0 then begin
                    For j := 0 to sStrings.Count -1 do begin
                      sT := sStrings.Strings[j];

                      sXPath := IncludeTrailingPathDelimiter(sToPath);
                      sDoc := sXPath + sT;

                      if FileExists(sDoc) then begin
                          if DeleteFile(sDoc) = False then begin
                              sErrorMessage := sErrorMessage +
                              'Err: Failing to delete the existed file: ' +
                              sDoc + #13#10;
                          end;
                      end;
                    end;
                end;


                if sErrorMessage <> '' then begin
                    PopUp_ErrorMessage(sErrorMessage + 'Extracting has been terminated.');
                    XslSheetsFilenames.Clear;
                    exit;
                end;


            end;

          {  Node := VT.GetFirstSelected;
            while Assigned(Node) do begin
              Data := VT.GetNodeData(Node);
              sFileNameIs := Data.Value[0]; // Name
              sFileNameIs := AnsiReplaceStr(sFileNameIs, '*', ''); // remove password char

              sZipFolder := Data.Value[8]; // Path
              if (sZipFolder <> '') then
                  sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);

              sStrings.Append(sZipFolder + sFileNameIs);

              Node := VT.GetNextSelected(Node);
            end; }

            StartUnzip(False, ZipFName.Caption, sToPath, sStrings, True);
            XbIsDoUnzip := True; // existed xml files are deleted and do unzip, this prevents to read old xml files
          finally
            sStrings.Free;
          end;
      end;
  end;
end;

procedure TfrmMain.StartUnzip(bVerbose: boolean; sZipFilename: string;
 sUnzipToFolder: string; sSelPathFNames: TStrings; bExtractSelected: boolean);
const
  cMinFreeSpace = 100000; // 100 kb
var
  sSaveDir, sWhatDrive: string;
  sStrings: TStrings;
  i, iCnt, iRetLevel, iNumFiles: integer;
  iFreeSpace, iFileSize, iTotalFileSize: Cardinal;
  bNoRetErr, bFree_DiskSpace_And_Try: boolean;
begin
  with ZipMaster1 Do Begin
    if (bVerbose = False) and ( Length(sUnzipToFolder) = 0 ) then begin
        ShowMessage('Err: Empty name of target directory');
        exit;
    end;
    sSaveDir := GetCurrentDir;
    SetCurrentDir(sUnzipToFolder);
   { if (bVerbose = False) and (ExcludeTrailingPathDelimiter(GetCurrentDir) <>
    ExcludeTrailingPathDelimiter(sUnzipToFolder) ) then begin
        ShowMessage('Error occurred when setting to new directory: ' +
        sUnzipToFolder);
        exit;
    end; }

    sErrorMessage := ''; // catch error message if found
    Verbose := bVerbose;
    ExtrBaseDir := sUnzipToFolder;
    ExtrOptions := [];
    if sZipTempFolder = XcZipTemp_WinTemp then
        TempDir := XsWinPath + 'Temp'
    else if sZipTempFolder = XcZipTemp_Current then
        TempDir := sSaveDir
    else if sZipTempFolder = XcZipTemp_Target then 
        TempDir := ExtractFileDir(sUnzipToFolder);

    //ZipMaster1.Password := ''; // important to clear previous password
    ZipMaster1.PasswordReqCount := 3; // important if show password diaglog box

    Gauge1.Visible := True;
    Gauge1.MinValue := 0;
    Gauge1.Progress := 0;
    Gauge2.Visible := True;
   { if bExtractSelected then
        Gauge1.MaxValue := ZipMaster1.FSpecArgs.Count
    else // all files }
        Gauge1.MaxValue := ZipMaster1.Count; 

    if bVerbose then begin
        if btnShowVirtual.Down then begin
            AddNode_ToTreeView;
            TreeView1.Selected := TreeView1.TopItem;
            TreeView1_Change;
        end
        else
            Refresh_ItemsOnMainList;

        btnComment.Visible := (ZipMaster1.ZipComment <> '');
    end
    else begin // extract to dir
      Screen.Cursor := crHourGlass;
      try
        //if  then
        //    ExtrOptions := ExtrOptions + [ExtrTest]; // Test files used this
        If RadioExtractWithDir.ItemIndex = 1 Then
            ExtrOptions := ExtrOptions + [ExtrDirNames];
        If RadioOverwrite.ItemIndex = 1 Then
            ExtrOptions := ExtrOptions + [ExtrOverwrite];

        mnuTerminate.Enabled := True; // Terminate Process

        sWhatDrive := ExtractFileDrive(sUnzipToFolder);
        if sWhatDrive <> '' then
            sWhatDrive := IncludeTrailingPathDelimiter(sWhatDrive);
            
            iRetLevel := LevelOf_FreeDiskSpace_Extract(sWhatDrive);
            if iRetLevel = 1 then begin // safe condition, do it once
                try
                  if bExtractSelected then begin // extract selected files only
                      ZipMaster1.FSpecArgs.Clear;
                      ZipMaster1.FSpecArgs.AddStrings(sSelPathFNames);
                      Gauge1.MaxValue := ZipMaster1.FSpecArgs.Count
                  end;
                  Extract;
                except
                  ShowMessage('Err: Error in Extraction. Fatal DLL Exception.');
                end;
            end
            else begin // -1 or 0. Control number of files to be extracted
                try
                  sStrings := TStringList.Create;
                  if bExtractSelected then begin // extract selected files only
                      iCnt := sSelPathFNames.Count -1;
                      Gauge1.MaxValue := sSelPathFNames.Count; // update
                  end
                  else
                      iCnt := ZipMaster1.Count -1;
                      
                  i := 0;
                  iNumFiles := 100;
                  repeat
                    iFreeSpace := FreeDiskSpace(sWhatDrive, bNoRetErr);
                    if bExtractSelected then begin
                        if (bNoRetErr) and (iFreeSpace <= cMinFreeSpace) then begin
                            beep;
                            ShowMessage('Err: Terminated. Not enough free disk space on target location. ' +
                            'Please make sure there is over 100 KB free.');
                            break;
                        end
                        else if (bNoRetErr) then begin
                            iNumFiles := iFreeSpace div cMinFreeSpace;
                            if iNumFiles <= 2 then
                                iNumFiles := 3;

                        end;
                    end;

                    bFree_DiskSpace_And_Try := False;
                    iTotalFileSize := 0;
                    sStrings.Clear;
                    for i := iCnt downto 0 do begin
                      if bExtractSelected then begin // extract selected files only
                          sStrings.Append(sSelPathFNames.Strings[i]);
                          if (i mod iNumFiles = 0) then begin
                              iCnt := i -1; // start from next
                              break;
                          end;
                      end
                      else begin // extract all files
                          iFileSize := ZipDirEntry(ZipMaster1.ZipContents[i]^).UncompressedSize;
                          iTotalFileSize := iTotalFileSize + iFileSize;
                          if bNoRetErr and (iTotalFileSize < iFreeSpace) then begin
                              if bExtractSelected then  // extract selected files only
                                  sStrings.Append(sSelPathFNames.Strings[i])
                              else // all files
                                  sStrings.Append(ZipDirEntry(ZipMaster1.ZipContents[i]^).FileName);

                          end
                          else begin // not enough free disk space
                              if sStrings.Count = 0 then begin
                                  beep;
                                  if MessageDlg('Interrupted. Not enough free disk space on target location. ' +
                                  #13#10 + 'Would you like to free more disk space on target location and ' +
                                  'retry again?', mtError, [mbYes, mbNo], 0) = mrYes then begin
                                      bFree_DiskSpace_And_Try := True; // important! don't quit loop when no files were extracting
                                     iCnt := i; // retry from current pos next time
                                  end;
                                  break;
                              end
                              else begin
                                  iCnt := i; // start from current pos next time
                                  break;
                              end;
                          end;
                      end;
                    end;

                    if sStrings.Count > 0 then begin
                        ZipMaster1.FSpecArgs.Clear;
                        ZipMaster1.FSpecArgs.AddStrings(sStrings);
                        try
                          Extract;
                        except
                          ShowMessage('Err: Error in Extraction. Fatal DLL Exception.');
                          i := 0; // quit loop
                        end;
                    end
                    else begin
                        if not bFree_DiskSpace_And_Try then
                            i := 0; // quit loop
                            
                    end;
                  until i <= 0;
                finally
                  sStrings.Free;
                end;
            end;

        mnuTerminate.Enabled := False; // Terminate Process
      finally
        Screen.Cursor := crDefault;
      end;
    end;
    SetCurrentDir(sSaveDir);
   { if (GetCurrentDir <> sSaveDir) then
        ShowMessage('Error occurred when setting to previous directory: ' +
        sSaveDir); }
    { If SuccessCnt = 1 Then
          IsOne := ' was'
      Else
          IsOne := 's were';
      ShowMessage(IntToStr(SuccessCnt) + ' file' + IsOne + ' extracted'); }
    Gauge1.Visible := False;
    Gauge1.Progress := 0;
    Gauge2.Visible := False;
  end;

  if sErrorMessage <> '' then
      PopUp_ErrorMessage('Error occurred when extracting file(s)');
      
end;

procedure TfrmMain.btnNewClick(Sender: TObject);
var
  sPathFilename, sFileExt: string;
  bCreateSuc: boolean;
begin
  if Sender <> nil then
      OpenDialog1.Filter := 'New zip file(.zip)|*.zip'
  else
      OpenDialog1.Filter := 'New zip file(.zip) or existed|*.zip';

  OpenDialog1.InitialDir := GetCurrentDir;

  repeat
    bCreateSuc := True;
    if OpenDialog1.Execute then begin
      btnClose.Click; // Close opened file first

      sPathFilename := OpenDialog1.FileName;
      sFileExt := AnsiLowercase( ExtractFileExt(sPathFilename) );
      if sFileExt <> '.zip' then
          sPathFilename := sPathFilename + '.zip';

      if (Sender <> nil) and FileExists(sPathFilename) then begin
          beep;
          MessageDlg('File already exists.', mtError, [mbOK], 0);
          bCreateSuc := False;
      end;
     { if FileExists(sPathFilename) then begin
          if MessageDlg(sPathFilename + #13+#10 + 'File already exists. Do you ' +
          'really want to overwrite?', mtConfirmation, [mbYes,mbNo], 0) = mrYes
          then begin
              if not DeleteFile(sPathFilename) then begin
                  beep;
                  MessageDlg(sPathFilename + #13+#10 + 'Failing to delete ' +
                  'above file. Please try again.', mtError, [mbOK], 0);
                  exit;
              end;
          end
          else
              exit;  // Don't use the new name
      end; }

      if bCreateSuc then begin
          if btnShowVirtual.Down then
              SetZipFile(sPathFilename, False)
          else
              SetZipFile(sPathFilename, True);

          mnuAddFilesToZipClick(nil);
      end;
    end;
  until bCreateSuc;

 { if VT.TotalCount = 0 then begin
      ZipFname.Caption := '';
      StatusBar1.Panels[0].Text := ZipFname.Caption;
      StatusBar1.Panels[1].Text := '';
      StatusBar1.Panels[2].Text := '';
  end; }
end;

procedure TfrmMain.mnuAddFilesToZipClick(Sender: TObject);
begin
  if ZipFName.Caption = '' then begin
      beep;
      MessageDlg('Error, target zip file not exist.', mtError, [mbOK], 0);
      exit;
  end;

  if (ZipFName.Caption <> '') then begin
      edtZipFileName.Text := ZipFName.Caption;
      if XbHaveSelectedFiles then begin // through context menu or drag-drop. Filenames on listbox
          with TOptionsOfPopupAdd.Create(self) do
          try
            showmodal;
          finally
            free;
          end;
          if bCancelAddFiles then
              exit;

      end
      else begin // process starts on Main screen.
          slZipStrings.Clear;

          with TfrmDir.Create(self) do
          try
            Caption := 'Add files';

            //ShellListView1.ObjectTypes := [otFolders, otNonFolders,
            //otHidden];
            tabUnzip.TabVisible := False;
            tabAdd.TabVisible := True;
            DirView1.MultiSelect := True;
            SetFilterForListView('*.*');
            btnHiddenFiles.Down := bDontShow_HiddenFiles;
            btnSystemFiles.Down := bDontShow_SystemFiles;
            if bDontShow_HiddenFiles then
                DirView1.SelHidden := SelNo // don't show hidden files
            else
                DirView1.SelHidden := SelDontCare;

            if bDontShow_SystemFiles then
                DirView1.SelSysFile := SelNo // don't show system files
            else
                DirView1.SelSysFile := SelDontCare;

            //ShellListView1.AutoNavigate := True;
            ShowModal; // get files that selected by user
          finally
            free;
          end;
      end;

      if (not bCancelAddFiles) and (slZipStrings.Count > 0) then begin
          Zipping;

          if bTerminatedProcess then begin // reload the archive
              bTerminatedProcess := False;
              if btnShowVirtual.Down then begin
                  SetZipFile(ZipFName.Caption, False);
                  StartUnzip(True, ZipFName.Caption, '', nil, False);
              end
              else
                  SetZipFile(ZipFName.Caption, True);

              //ListView1.SetFocus;
              VT.SetFocus;
          end
          else begin // refresh can do
              if btnShowVirtual.Down then begin
                  AddNode_ToTreeView;
                  TreeView1.Selected := TreeView1.TopItem;
                  TreeView1_Change;
              end
              else begin
                  Update_FDataList; // Must do this first
                  Refresh_ItemsOnMainList;
              end;
          end;

          // Are some files extracted to Windows\Temp\xxx ? True to remove it.
          if bExtractedFiles or (edShadow.Text <> '') then
              Remove_TempFiles_UnderWinTemp;

          bExtractedFiles := False;
      end
      else if (not bCancelAddFiles) and (slZipStrings.Count = 0) then begin
          beep;
          MessageDlg('No selected files found. If you selected folders, the ' +
          'folders may be empty.', mtInformation, [mbOK], 0);
      end
  end;
end;

procedure TfrmMain.StartZip(sToPathZipName: string; sFiles: TStrings);
const
  cMinFreeSpace = 100000; // 100 KB
var
  sWhatDrive: string;
  sStrings: TStrings;
  i, iCnt, iFileSize, iPredict: integer;
  iFreeSpace, iTotalFilesSize: Cardinal;
  bNoRetErr, bRemoveIt, bFree_DiskSpace_And_Try: boolean;
begin
  with ZipMaster1 do
    begin
      if sZipTempFolder = XcZipTemp_WinTemp then
          TempDir := XsWinPath + 'Temp'
      else if sZipTempFolder = XcZipTemp_Current then
          TempDir := GetCurrentDir
      else if sZipTempFolder = XcZipTemp_Target then
          TempDir := ExtractFileDir(sToPathZipName);

        { We want any DLL error messages to show over the top
          of the message form. }
        //AddOptions := [];
        //AddOptions := zmOpts;
       { Case AddForm.ZipAction Of       // Default is plain ADD.
            2: AddOptions := AddOptions + [AddUpdate]; // Update
            3: AddOptions := AddOptions + [AddFreshen]; // Freshen
            4: AddOptions := AddOptions + [AddMove]; // Move
        End;

        If AddForm.RecurseCB.Checked Then
            AddOptions := AddOptions + [AddRecurseDirs]; // we want recursion
        If AddForm.AtribOnlyCB.Checked Then
            AddOptions := AddOptions + [AddArchiveOnly]; // we want changed only
        If AddForm.AtribResetCB.Checked Then
            AddOptions := AddOptions + [AddResetArchive]; // we want reset
        If AddForm.DirnameCB.Checked Then
            AddOptions := AddOptions + [AddDirNames]; // we want dirnames
        If AddForm.DiskSpanCB.Checked Then
            AddOptions := AddOptions + [AddDiskSpan]; // we want diskspanning
        If AddForm.EncryptCB.Checked Then
        Begin
            AddOptions := AddOptions + [AddEncrypt]; // we want a password
            // GetAddPassword;
            // if Password = '' then
                  // The 2 password's entered by user didn't match.
                  // We'll give him one more try; if he still messes it
                  //  up, the DLL itself will prompt him one final time.
             //   GetAddPassword;
        End; }

        sWhatDrive := ExtractFileDrive(sToPathZipName);
        if sWhatDrive <> '' then
            sWhatDrive := IncludeTrailingPathDelimiter(sWhatDrive);

            sStrings := TStringList.Create;
            try
              Screen.Cursor := crHourGlass;
              mnuTerminate.Enabled := True; // Terminate Process
              iCnt := sFiles.Count -1;
              Gauge1.MaxValue := sFiles.Count; // update

              repeat
                bFree_DiskSpace_And_Try := False;
                iTotalFilesSize := 0;
                iFreeSpace := FreeDiskSpace(sWhatDrive, bNoRetErr);
                if (bNoRetErr) and (iFreeSpace <= cMinFreeSpace) then begin
                    beep;
                    if MessageDlg('Interrupted. Not enough free disk space on target location. ' +
                    #13#10 + 'Would you like to free more disk space on target location and ' +
                    'retry again?', mtError, [mbYes, mbNo], 0) <> mrYes then
                        break;
                        
                end;

                if iCnt > 0 then begin
                    Gauge2.MaxValue := iCnt; // must be greater than 0
                    Gauge2.Progress := iCnt;
                    if HourGlass.Showing then begin
                        HourGlass.Gauge2.MaxValue := iCnt;
                        HourGlass.Gauge2.Progress := iCnt;
                    end;
                end;
                
                sStrings.Clear;
                for i := iCnt downto 0 do begin
                  iFileSize := FilterSysOrHidFile_And_Get_FileSize(
                  sFiles.Strings[i], bRemoveIt);
                  if bRemoveIt then // user wants to remove system or hidden file, do it here
                      sFiles.Delete(i)
                  else begin
                      iTotalFilesSize := iTotalFilesSize + iFileSize;
                      if iTotalFilesSize < iFreeSpace then begin
                          sStrings.Append(sFiles.Strings[i]);
                          if Gauge2.Progress > 0 then begin
                              Gauge2.Progress := Gauge2.Progress -1;
                              Gauge2.Repaint;
                          end;

                          if HourGlass.Showing then begin
                              if HourGlass.Gauge2.Progress > 0 then
                                  HourGlass.Gauge2.Progress := HourGlass.Gauge2.Progress -1;
                                  
                          end;
                      end
                      else begin // detected not enough disk space, ignore this file
                          iCnt := i; // retry from current pos next time
                          if sStrings.Count = 0 then begin // No files can be added later
                              iPredict := Get_Predict_Compressed_FileSize(sFiles.Strings[i]);
                              if (iPredict <> -1) and (iPredict < iFreeSpace) then begin
                                  sStrings.Append(sFiles.Strings[i]);
                                  iCnt := i -1; // start from next pos
                              end
                              else if (iPredict <> -1) then begin // can get predict file size
                                  beep;
                                  if MessageDlg('Interrupted. Not enough free disk space on target location. ' +
                                  #13#10 + 'Would you like to free more disk space on target location and ' +
                                  'retry again?', mtError, [mbYes, mbNo], 0) = mrYes then begin
                                      bFree_DiskSpace_And_Try := True; // important! don't quit loop when no files were added
                                      iCnt := i; // retry from current pos next time
                                  end;
                              end
                              else begin
                                  beep;
                                  MessageDlg('Terminated. Not enough free disk space on target location. ',
                                  mtError, [mbOK], 0);
                              end;
                          end;
                          break;
                      end;
                  end;
                end;

                if sStrings.Count > 0 then begin
                    FSpecArgs.Clear;
                    FSpecArgs.Assign(sStrings); { specify filenames }
                    try
                      Add;
                    except
                      ShowMessage('Error occurred. Fatal DLL Exception.');
                      i := -1;
                    end;
                end
                else begin
                    if not bFree_DiskSpace_And_Try then
                        i := -1; // quit loop

                end;

              until i <= -1;
            finally
              sStrings.Free;
              Gauge2.Visible := False;
              mnuTerminate.Enabled := False; // Terminate Process
              Screen.Cursor := crDefault;
            end;


        if Main.sComment <> '' then begin // I put it here although created new zip
            ZipMaster1.ZipComment := sComment;
            ZipMaster1.List;
            Main.sComment := '';
        end;

        // important if user wants to test extraction after
        if ZipMaster1.Password <> '' then
            ZipMaster1.Password := '';
       { If SuccessCnt = 1 Then
            IsOne := ' was'
        Else
            IsOne := 's were';
        ShowMessage(IntToStr(SuccessCnt) + ' file' + IsOne + ' added'); }
    end;
end;

procedure TfrmMain.Remove_TempFiles_UnderWinTemp;
begin
  ReadFileAndDelete(nil, edShadow.Text);
  KillSubFolderFiles(nil, edShadow.Text);
  RemoveDir(edShadow.Text);
end;

function TfrmMain.ReadFileAndDelete(Sender: TObject; sPathIs: string): boolean;
var
  sr: TSearchRec;
  iFileAttrs: Integer;
  iRetValue: integer;
begin
  if not DirectoryExists(sPathIs) then begin
      ReadFileAndDelete := False;
      exit;
  end;
  iFileAttrs := $0000003F; //faAnyFile  faArchive;   //archives only
  if FindFirst(sPathIs + '*.*', iFileAttrs, sr) = 0 then begin
      repeat
        if ((sr.Attr and iFileAttrs) = sr.Attr) and
        (sr.Name <> '.') and (sr.Name <> '..') then begin
            try
              if deletefile(sPathIs + sr.Name) = False then begin
                 iRetValue := FileSetAttr(sPathIs + sr.Name, not faReadOnly and
                 not faSysFile and not faHidden and faArchive);
                 deletefile(sPathIs + sr.Name);
              end
            except

            end;
        end;
      until FindNext(sr) <> 0;
      FindClose(sr);
      ReadFileAndDelete := True;
  end
  else begin
      ReadFileAndDelete := True;
  end;
end;

procedure TfrmMain.KillSubFolderFiles(Sender: TObject; sPathIs: string);
var
  sr: TSearchRec;
  iFileAttrs: Integer;
  sNewpath: string;
begin
  if not DirectoryExists(sPathIs) then exit;

  iFileAttrs := faAnyFile;  // + faSysFile;   // faDirectory
  if FindFirst(sPathIs + '*.*', iFileAttrs, sr) = 0 then begin
      repeat
        if ((sr.Attr and iFileAttrs) = sr.Attr) and
        (sr.Name <> '.') and (sr.Name <> '..') then begin
            sNewpath := sPathIs + sr.Name + '\';
            ReadFileAndDelete(nil, sNewpath);
            KillSubFolderFiles(nil, sNewpath); // recursion
            RemoveDir(sNewpath);
        end;
      until FindNext(sr) <> 0;
      FindClose(sr);
  end;
end;

procedure TfrmMain.FileOrFolder_ThroughContext_ToBeZipped(sPathZipName: string);
var
  iRetReply: Word;
begin
  if FileExists(sPathZipName) then begin
      beep;
      iRetReply := MessageDlg('Zip file already exists. Do you want to overwrite?',
      mtConfirmation, [mbYes, mbNo, mbCancel], 0);
      if iRetReply = mrYes then begin
          if not DeleteFile(sPathZipName) then begin // failing to delete
              FileSetAttr(sPathZipName, not faReadOnly and not faSysFile and
              not faHidden and faArchive);

              if not DeleteFile(sPathZipName) then begin
                  beep;
                  MessageDlg('Failing to overwrite. Error occurred when trying ' +
                  'to delete it. Process will terminate.', mtError, [mbOK], 0);
                  exit;
              end;
          end;
      end
      else if iRetReply = mrCancel then
          exit;
          
  end;
  SetZipFile(sPathZipName, False);
  mnuAddFilesToZipClick(nil);
end;

procedure TfrmMain.btnPopupOpenClick(Sender: TObject);
begin
  OpenZipFile;
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
 { if ZipFName.Caption = '' then begin
      beep;
      exit;
  end; }
  mnuCloseAchiveClick(nil);
end;

procedure TfrmMain.OpenZipFile;
var
  //ObjectTypes: TShellObjectTypes;
  sPathFilename: string;
begin
 { GetDirectory.SetUpView(nil, 1);
  GetDirectory.ShowModal;}
{  ObjectTypes := [otFolders];//, otNonFolders];
  GetFolderName.stvDirectory.ObjectTypes := ObjectTypes;
  ObjectTypes := [otFolders, otNonFolders];
  GetFolderName.ShellListView1.ObjectTypes := ObjectTypes;
  GetFolderName.PageControl1.Visible := True;
  GetFolderName.tabExtract.TabVisible := False;
  GetFolderName.tabOpen.TabVisible := True;
  GetFolderName.ShowModal; }

  if btnAdvancedView.Down then begin
      with TfrmOpenDir.Create(nil) do
      try
        //Caption := 'Open file';
        //tabAdd.TabVisible := False;
        //tabUnzip.TabVisible := True;
        //DirView1.MultiSelect := False;
        //SetFilterForListView('.docx');
        //pnBottom.Height := 155;
        //ShellListView1.AutoNavigate := True;
        ShowModal; // get files that selected by user
        sPathFilename := edPath.Text;
      finally
        free;
      end;
  end
  else begin // Simple View on Main screen
      bCancelOpenZip := True;
      if OpenDlgSimple.Execute then begin
          sPathFilename := OpenDlgSimple.FileName;
          bCancelOpenZip := False;
      end;
  end;

  if not bCancelOpenZip then begin // not cancel
      if FileExists(sPathFilename) then begin
          if btnShowVirtual.Down then begin
              SetZipFile(sPathFilename, False);
              StartUnzip(True, sPathFilename, '', nil, False);
          end
          else
              SetZipFile(sPathFilename, True);

          Remove_StringGrids_And_Tabsheets;  // must do first before AfterOpeningArchive_And_Do
          AfterOpeningArchive_And_Do;

          // Are some files extracted to Windows\Temp\xxx ? True to remove it.
          if bExtractedFiles and (edShadow.Text <> '') then
              Remove_TempFiles_UnderWinTemp;

          RecentUnzipFilesToReg(sPathFilename);
          bExtractedFiles := False;
      end
      else begin
          beep;
          MessageDlg('File does not exist.', mtError, [mbOK], 0);
          exit;
      end;
  end;
end;

procedure TfrmMain.mnuCloseAchiveClick(Sender: TObject);
begin
  Screen.Cursor := crHourGlass;
  try
   { ListView1.Items.BeginUpdate;
    ListView1.Items.Clear;
    ListView1.Items.EndUpdate; }
    btnViewTxt.Enabled := False;
    btnTest.Enabled := False;
    //btnRepair.Enabled := False;
    btnRecoverIMAGES.Visible := False;

    Remove_StringGrids_And_Tabsheets;

    lblOpenFileIs.Caption := '';
    
    FDataList.Clear;
    VT.Clear;
    ClearColumnsImage_OnlvMain;
    TreeView1.Items.Clear;
    if btnShowVirtual.Down then
        Update_lblCurDir('');

    SetZipFile('', False);
    btnComment.Visible := False;
    SetStatusOfPopupUnzip(nil);

    btnSaveTxtToFile.Enabled := False;
    btnSaveAllToCSV.Enabled := False;
    memOutput.Clear;

    //StatusBar1.SimpleText := '';
    StatusBar1.Panels[0].Text := '';
    StatusBar1.Panels[1].Text := '';
    StatusBar1.Panels[2].Text := '';
    if bExtractedFiles and (edShadow.Text <> '') then
        Remove_TempFiles_UnderWinTemp;

  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmMain.RecentUnzipFilesToReg(const sPathFilenameIs: string);
var
  i: integer;
  sPathFilenameInMenu: string;
  iC: integer;
begin
  // Is count over limit?
  if OpenedFilesPopup.Items.Count > cCountOfRecentUnzipFiles then begin
      for i := OpenedFilesPopup.Items.Count -1 downto cCountOfRecentUnzipFiles
      do begin
        OpenedFilesPopup.Items.Delete(i);
        //if OpenedFilesPopup.Items.Count <= cCountOfRecentUnzipFiles then // be my count
        //    break;

      end;
  end;
  // Is path already exist?
  for i := 0 to OpenedFilesPopup.Items.Count -1 do begin
    if AnsiLowercase(OpenedFilesPopup.Items[i].Caption) =
    AnsiLowercase(sPathFilenameIs) then
        exit;

  end;
  // Full?
  if OpenedFilesPopup.Items.Count >= cCountOfRecentUnzipFiles then
      OpenedFilesPopup.Items.Delete(OpenedFilesPopup.Items.Count -1);

  //iC := OpenedFilesPopup.Items.Count;
  //if iC >= cCountOfRecentUnzipFiles then // exceed array value
  //    iC := cCountOfRecentUnzipFiles -1;
  iC := 0;
  if FileExists(sPathFilenameIs) then begin
      mnuNewMake[iC] := TMenuItem.Create(Owner);
      mnuNewMake[iC].Caption := sPathFilenameIs;
      OpenedFilesPopup.Items.Insert(iC,mnuNewMake[iC]);
      mnuNewMake[iC].OnClick := OpenedFileOnClick;
      mnuNewMake[iC].AutoCheck := True;
  end;

  // update to registry
  for i := OpenedFilesPopup.Items.Count - 1 downto 0 do begin
    sPathFilenameInMenu := OpenedFilesPopup.Items[i].Caption;
    frmMain.WriteIniToReg(sSubKeyRecentUnzip, inttostr(i), sPathFilenameInMenu);
  end;
end;

procedure TfrmMain.OpenedFileOnClick(Sender: TObject);
var
  i, j: shortint;
  iCountOfOldFiles: shortint;
  sGetOldPathFilename, sPathFilenameInMenu: string;
begin
  sGetOldPathFilename := '';
  iCountOfOldFiles := OpenedFilesPopup.Items.Count;
  for i := 0 to iCountOfOldFiles -1 do begin
    if OpenedFilesPopup.Items[i].Checked then begin
        sGetOldPathFilename := OpenedFilesPopup.Items[i].Caption;
        OpenedFilesPopup.Items[i].Checked := False;
        j := i;
        break;
    end;
  end;

  if sGetOldPathFilename <> '' then begin
      if not FileExists(sGetOldPathFilename) then begin
          beep;
          MessageDlg('File does not exist.', mtError, [mbOK], 0);
          OpenedFilesPopup.Items.Delete(j);
          // update to registry
          for i := OpenedFilesPopup.Items.Count - 1 downto 0 do begin
            sPathFilenameInMenu := OpenedFilesPopup.Items[i].Caption;
            if sPathFilenameInMenu <> '' then
                WriteIniToReg(sSubKeyRecentUnzip, inttostr(i),
                sPathFilenameInMenu);
          end;
          exit;
      end;

      if btnShowVirtual.Down then begin
          SetZipFile(sGetOldPathFilename, False);
          StartUnzip(True, sGetOldPathFilename, '', nil, False);
      end
      else
          SetZipFile(sGetOldPathFilename, True);

      Remove_StringGrids_And_Tabsheets; // must do first before AfterOpeningArchive_And_Do
      AfterOpeningArchive_And_Do;

      // Are some files extracted to Windows\Temp\xxx ? True to remove it.
      if bExtractedFiles and (edShadow.Text <> '') then
          Remove_TempFiles_UnderWinTemp;

      bExtractedFiles := False;
  end
  else begin
      beep;
      MessageDlg('Error found, empty string', mtError,
      [mbOK], 0);
  end;
end;

procedure TfrmMain.FavoriteDirOnClick(Sender: TObject);
var
  i, j, k: shortint;
  sOldPath: string;
begin
  sOldPath := '';
  for i := 0 to pmFavoriteDir.Items.Count -1 do begin
    if pmFavoriteDir.Items[i].Checked then begin
        sOldPath := pmFavoriteDir.Items[i].Caption;
        pmFavoriteDir.Items[i].Checked := False;
        j := i;
        break;
    end;
  end;

  if sOldPath <> '' then begin
      if not DirectoryExists(sOldPath) then begin
          beep;
          MessageDlg('Directory does not exist.', mtError, [mbOK], 0);
          pmFavoriteDir.Items[j].Caption := '';
          for k := j to pmFavoriteDir.Items.Count -1 do begin
            if k = (pmFavoriteDir.Items.Count -1) then begin // on last item
                pmFavoriteDir.Items[k].Caption := ''; // set last item to empty
                break;
            end;
            // move up
            pmFavoriteDir.Items[k].Caption := pmFavoriteDir.Items[k +1].Caption
          end;

          DelRegistryKey(HKEY_CURRENT_USER, XcSubKeyIs + cFavouriteFolders);
          // Save favourite folders to registry
          for k := 0 to pmFavoriteDir.Items.Count -1 do begin
            if pmFavoriteDir.Items[k].Caption = '' then break;
            WriteIniToReg(XcSubKeyIs + cFavouriteFolders, inttostr(k),
            pmFavoriteDir.Items[k].Caption)
          end;
          exit;
      end;
      sOldPath := IncludeTrailingPathDelimiter(sOLdPath);
      //ShellDir1.Path := sOldPath;
      VETree.BrowseTo(sOldPath);
  end
  else begin
      beep;
      MessageDlg('Error found, empty string', mtError,
      [mbOK], 0);
  end;
end;

procedure TfrmMain.mnuOpenWithClick(Sender: TObject);
var
  sPathFileName: string;
  bIsVirtualFolder, bIsEncrypted: boolean;
begin
  if VT.SelectedCount = 1 then begin
      sPathFileName := GetPathFileName_FromSelected_Item(bIsVirtualFolder, bIsEncrypted);
      if (sPathFileName <> '') and (bIsVirtualFolder = False) then
          ExecuteFileClick(nil, sPathFileName, False, '');
     { beep;
      MessageDlg('Nothing selected.', mtInformation, [mbOK], 0);
      exit; }
  end;

end;

function TfrmMain.GetPathFileName_FromSelected_Item(var bIsFolder, bIsEncrypted: boolean): string;
var
  i, iR, iCntCol: integer;
  sFileNameIs: string;
  sPathIs: string;
  sPathFileName: string;
  //sSubItemsText: string;
  Node: PVirtualNode;
  Data: PEntry;
begin
  Result := '';
  bIsFolder := False;
  if VT.SelectedCount = 0 then exit;
  //sFileNameIs := ListView1.Selected.Caption;
  Node := VT.GetFirstSelected;
 // while Assigned(Node) do begin
    Data := VT.GetNodeData(Node);
    sFileNameIs := Data.Value[0]; // Name
    sPathIs := Data.Value[8]; // Path
 // end;

  if AnsiPos('*', sFileNameIs) > 0 then
      bIsEncrypted := True
  else
      bIsEncrypted := False;
      
  sFileNameIs := AnsiReplaceStr(sFileNameIs, '*', ''); // remove password char
  if  sFileNameIs = '' then exit;

 { iR := ListView1.Selected.Index;

  if (ListView1.Items.Item[iR].SubItems.Strings[0] = '') and
  (ListView1.Items.Item[iR].SubItems.Strings[1] = '') and
  btnShowVirtual.Down then
      bIsFolder := True;

  iCntCol := ListView1.Columns.Count;
  sPathIs := ListView1.Items.Item[iR].SubItems.Strings[iCntCol -2]; }

  if sPathIs <> '' then
      sPathIs := IncludeTrailingPathDelimiter(sPathIs);

  sPathFileName := sPathIs + sFileNameIs;
  Result := sPathFileName;
end;

procedure TfrmMain.ExecuteFileClick(Sender: TObject;
 sExtractedPathFilename: string; bDoubleClick: boolean; sWithThisApp: string);
var
  sOpenFile, sFileExt, sPath_NotepadProg, sPathExeProgram, sPahtFileNameIs: string;
  RetValue, H: integer;
  bExtractOne: boolean;
  sStrings: TStrings;
begin
  // Is there no any extraced file? (= no temporary folder was created)
  if bExtractedFiles = False then begin
      edShadow.Text := Create_TempFolder_Under_WindowsTemp;
      if edShadow.Text = '' then
          exit
      else
          bExtractedFiles := True;
  end;

  sFileExt := AnsiLowercase( ExtractFileExt(sExtractedPathFilename) );
  if sFileExt = '.exe' then
      bExtractOne := False
  else
       bExtractOne := True;

  try
    sStrings := TStringList.Create;
    if bExtractOne then
        sStrings.Append(sExtractedPathFilename)
    else
        sStrings := nil;

    RadioExtractWithDir.ItemIndex := 1;
    RadioOverwrite.ItemIndex := 1;
    StartUnzip(False, ZipFName.Caption, edShadow.Text, sStrings, bExtractOne);
  finally
    sStrings.Free;
  end;
  sOpenFile := edShadow.Text + sExtractedPathFilename;
  // Is it not a double-click from ListView?
  if not bDoubleClick then begin
      if sWithThisApp <> '' then begin // using specified App
          sPathExeProgram := sWithThisApp;
          sPathExeProgram := ExtractShortPathName(sPathExeProgram);
          sPahtFileNameIs := '"' + sOpenFile + '"'; // When opening with specified appication + parameters, this will open long filename correctly
          H := ShellExecute(handle, 'open', PAnsiChar(sPathExeProgram),
               PAnsiChar(sPahtFileNameIs), '', SW_SHOW);
          if H = 31 then
              MessageDlg('Error occurred when trying to open file with '
              + 'this application.', mtError, [mbOK], 0);

      end
      else
          FileOpenedBy(sOpenFile); // right menu-click,open file with
          
      exit;
  end
  else begin // this is double-click from ListView
    if bOpen_TextFile_WithNotePad and (sFileExt = '.txt') then begin
        sPath_NotepadProg := XsWinPath + 'notepad.exe';
        sPath_NotepadProg := ExtractShortPathName(sPath_NotepadProg);
        sOpenFile := '"' + sOpenFile + '"'; // When opening with specified appication + parameters, this will open long filename correctly
        RetValue := ShellExecute(handle, 'open', PAnsiChar(sPath_NotepadProg),
                    PAnsiChar(sOpenFile), '', SW_SHOW);
        if RetValue = 31 then begin
            beep;
            MessageDlg('Failing to open text file.', mtError, [mbOK], 0);
            exit;
        end;
    end
    else begin
        //sOpenFile := '"' + sOpenFile + '"'; // No need if no parameters! Otherwise, Win95 will fail to open the file
        RetValue := ShellExecute(handle, 'open', PAnsiChar(sOpenFile),
        '', '', SW_SHOW);

        if RetValue = 31 then begin // Has not the filename association?
            beep;
            if MessageDlg('No application is associated with this ' +
            'sort of file. Do you want to open it with your choice?',
            mtConfirmation, [mbYes, mbNo], 0) = mrNo then
                exit
            else
                FileOpenedBy(sOpenFile); // try open file with
        end;
    end;
  end;
end;

function TfrmMain.Create_TempFolder_Under_WindowsTemp: string;
var
  i: integer;
  sTempFolder: string;
begin
  Result := '';
  // Is Windows' Temp not exist?
  if not DirectoryExists(XsWinPath + 'Temp\') then begin
      // Is Windows' Temp still not exist after making creation?
      if not CreateDir(XsWinPath + 'Temp\') then begin
          MessageDlg('Unexpected error, failing to create ' +
          'temporary folder under directory of Windows.', mtError,
          [mbOK], 0);
          exit;
      end;
  end;

  i := 1;
  repeat
    sTempFolder := XsWinPath + 'Temp\ziphaha' + inttostr(i) + '\';
    i := i + 1;
  until not DirectoryExists(sTempFolder); // want new dir

  if not ForceDirectories(sTempFolder) then begin // create dir
      MessageDlg('Unexpected error, failing to create a temporary folder ' +
      'under Windows Temp.', mtError, [mbOK], 0);
      exit;
  end
  else
      Result := sTempFolder;
end;

procedure TfrmMain.FileOpenedBy(sPahtFileNameIs: string);
var
  sPathExeProgram: string;
  H: Cardinal;
  AFiles: TStrings;
  i, iCnt: integer;
begin
  OpenDialog1.Filter := 'executable files (*.exe)|*.exe';
  if OpenDialog1.Execute then begin
      if OPenDialog1.FileName <> '' then begin
          sPathExeProgram := OPenDialog1.FileName;
          sPathExeProgram := ExtractShortPathName(sPathExeProgram);
          sPahtFileNameIs := '"' + sPahtFileNameIs + '"'; // When opening with specified appication + parameters, this will open long filename correctly
          H := ShellExecute(handle, 'open', PAnsiChar(sPathExeProgram),
               PAnsiChar(sPahtFileNameIs), '', SW_SHOW);
          if H = 31 then
              MessageDlg('Error occurred when trying to open file with '
              + 'this application.', mtError, [mbOK], 0);

          RecentOpenWithAppToReg(OPenDialog1.FileName);
          if mnuOpenWith_ThisApp.Tag = 1 then begin // already exist
              AFiles := TStringList.Create;
              try
                RecentOpenWithAppFromReg(AFiles);
                if AFiles = nil then exit;

                mnuOpenWith_ThisApp.Clear;
                iCnt := AFiles.Count;
                SetLength(mSubOpenWith_ThisApp, iCnt);
                for i := 0 to iCnt -1 do begin
                  mSubOpenWith_ThisApp[i] := TMenuItem.Create(Self);
                  mSubOpenWith_ThisApp[i].Caption := AFiles.Strings[i];
                  mSubOpenWith_ThisApp[i].OnClick := mSubOpenWith_ThisApp_OnClick;
                  mnuOpenWith_ThisApp.Insert(i, mSubOpenWith_ThisApp[i]);
                end;
              finally
                AFiles.Free;
              end;
          end;
      end
      else
          beep;
  end;
end;

procedure TfrmMain.mnuDeleteClick(Sender: TObject);
begin
  if VT.SelectedCount = 0 then begin
      beep;
      MessageDlg('Nothing selected.', mtInformation, [mbOK], 0);
      exit;
  end;

  if MessageDlg('Do you really want to delete selected file(s)?',
  mtConfirmation, [mbYes, mbCancel], 0) <> mrYes then
      exit;

 { if ListView1.SelCount = 0 then begin
      beep;
      MessageDlg('No selected file.', mtConfirmation, [mbOK], 0);
      exit;
  end; }

  if not FileExists(ZipFName.Caption) then begin
      ZipFName.Caption := '';
      beep;
      MessageDlg('Detected the opened zip file not exist!', mtError, [mbOK], 0);
      exit;
  end;

  if ZipMaster1.Count < 1 then begin
      MessageDlg('No opened zip file', mtError, [mbOK], 0);
      exit;
  end;

  if btnShowVirtual.Down then
      DeleteFiles_InArchive_OnVirtualMainList
  else
      DeleteFiles_InArchive;

end;

procedure TfrmMain.Initial_VT;
var
  k: integer;
  NewCol: TCollectionItem;
begin
 { ListView1.Column[0].Width := 180;
  ListView1.Column[1].Width := 140;
  ListView1.Column[2].Width := 60;
  ListView1.Column[3].Width := 48;
  ListView1.Column[4].Width := 70;
  ListView1.Column[5].Width := 300; }

      for k := VT.Header.Columns.Count -2 downto 5 do begin // remove extra columns
        VT.Header.Columns.Delete(k);
      end;
      if bCheckType then begin
          // insert to last column, new caption will replace old one.
          // If insert to ListView1.Columns.Count -1, Column index will be different. ListITem.SubItems.Add will depend on what col created early.
          NewCol := VT.Header.Columns.Insert(VT.Header.Columns.Count -1); // before Col Path, if u want last col, don't -1
          VT.Header.Columns[NewCol.Index].Text := 'File Type';
          VT.Header.Columns[NewCol.Index].Alignment := taRightJustify;
          VT.Header.Columns[NewCol.Index].MinWidth := 60;
          VT.Header.Columns[NewCol.Index].Width := iColWidthFileType;
      end;
      if bCheckCRC then begin
          NewCol := VT.Header.Columns.Insert(VT.Header.Columns.Count -1); // before Col Path, if u want last col, don't -1
          VT.Header.Columns[NewCol.Index].Text := 'CRC';
          VT.Header.Columns[NewCol.Index].Alignment := taRightJustify;
          VT.Header.Columns[NewCol.Index].MinWidth := 60;
          VT.Header.Columns[NewCol.Index].Width := iColWidthCRC;
      end;
      if bCheckAttributes then begin
          NewCol := VT.Header.Columns.Insert(VT.Header.Columns.Count -1); // before Col Path, if u want last col, don't -1
          VT.Header.Columns[NewCol.Index].Text := 'Attributes';
          VT.Header.Columns[NewCol.Index].Alignment := taRightJustify;
          VT.Header.Columns[NewCol.Index].MinWidth := 60;
          VT.Header.Columns[NewCol.Index].Width := iColWidthAttributes;
      end;
      SetBool_CntOf_ColumnChanged_OnMainLst(False);
     { ListView1.Column[ListView1.Columns.Count -1].Caption := 'Path'; // write back the caption Path
      ListView1.Column[ListView1.Columns.Count -1].Alignment := taRightJustify;
      ListView1.Column[ListView1.Columns.Count -1].Width := iColWidthPath; }
end;

procedure TfrmMain.RestoreStatusOfProgram;
var
  iMainTop, iMainLeft, iMainWidth, iMainHeight: longint;
  sMainTop, sMainLeft, sMainWidth, sMainHeight: string;
  sWindowStateIs, sUpDownBtnViewStyle, sUpDownBtnAdvancedView: string;

  bLastPos: boolean;
  sMainLstFontName, sMainLstFontSize: string;
  sWidthBrowseList, sHeightPnM: string;
  iWidthBrowseList, iHeightPnM: integer;
  bCheckBrowseList, bCheckGridLines, bCheckRowSelect, bCheckHotTrack: boolean;

  i: integer;
  sPath, sMainMenuSort: string;
begin
  with TfrmSettings.Create(Self) do
  try
    bLastPos := rbLast.Checked;
    sMainLstFontName := edFontName.Text;
    sMainLstFontSize := edFontSize.Text;
    bCheckMainLstAutoSize := chkAutoSize.Checked;
    bCheckBrowseList := chkBrowseList.Checked;
    bCheckGridLines := chkGridLines.Checked;
    bCheckRowSelect := chkRowSelect.Checked;
    bCheckHotTrack := chkHotTrack.Checked;
    
    SetBool_CheckType(chkType.Checked);
    SetBool_CheckCRC(chkCRC.Checked);
    SetBool_CheckAttributes(chkAttributes.Checked);
    SetBool_CheckFileToRecycleBin(chkFileToRecycleBin.Checked);
    
    SetBool_CheckFavoriteOpen(rbFavoriteOpen.Checked);
    SetBool_CheckFavoriteAdd(rbFavoriteAdd.Checked);
    SetBool_CheckFavoriteExtract(rbFavoriteExtract.Checked);
    SetStr_FavoriteOpen(edFavoriteOpen.Text);
    SetStr_FavoriteAdd(edFavoriteAdd.Text);
    SetStr_FavoriteExtract(edFavoriteExtract.Text);
    SetStr_VirusScannerProg(edVirusScanner.Text);
    SetStr_VirusScannerPara(edPara.Text);

    SetBool_CheckTxtWithNotePad(chkTxtWithNotePad.Checked);
    SetBool_CheckHiddenFiles(chkHiddenFiles.Checked);
    SetBool_CheckSystemFiles(chkSystemFiles.Checked);
    SetBool_RadioButtonParaFilename(rbParaFilename.Checked);
    
    if rbWinTemp.Checked then
        Set_ZipTempIs(XcZipTemp_WinTemp)
    else if rbCurrent.Checked then
        Set_ZipTempIs(XcZipTemp_Current)
    else if rbTarget.Checked then
        Set_ZipTempIs(XcZipTemp_Target);

  finally
    Free;
  end;

  if sMainLstFontName <> '' then
      //ListView1.Font.Name := sMainLstFontName;
      VT.Font.Name := sMainLstFontName;

  if sMainLstFontSize <> '' then
      //ListView1.Font.Size := strtoint(sMainLstFontSize);
      VT.Font.Size := strtoint(sMainLstFontSize);

  PanMiddleLeft.Visible := bCheckBrowseList;
  //ListView1.GridLines := bCheckGridLines;
  //ListView1.RowSelect := bCheckRowSelect;
  if bCheckGridLines then begin
      VT.TreeOptions.PaintOptions := VT.TreeOptions.PaintOptions + [toShowHorzGridLines];
      VT.TreeOptions.PaintOptions := VT.TreeOptions.PaintOptions + [toShowVertGridLines];
  end;

  if bCheckRowSelect then
      VT.TreeOptions.SelectionOptions := VT.TreeOptions.SelectionOptions + [toFullRowSelect];

  if bCheckHotTrack then
      VT.TreeOptions.PaintOptions := VT.TreeOptions.PaintOptions + [toHotTrack];

  // vvv Window state vvv
  sWindowStateIs := ReadIniFromReg(XcSubKeyIs, cMainWindowState);
  if (sWindowStateIs = 'wsMaximized') and bLastPos then
      Self.WindowState := wsMaximized
  else begin
      sMainTop := ReadIniFromReg(XcSubKeyIs, cMainTop);
      if sMainTop <> '' then begin
          iMainTop := strtoint(sMainTop);
          if (iMainTop > 0) and (iMainTop < Screen.Height) then
              Self.Top := iMainTop;

      end;
      sMainLeft := ReadIniFromReg(XcSubKeyIs, cMainLeft);
      if sMainLeft <> '' then begin
          iMainLeft := strtoint(sMainLeft);
          if (iMainLeft > 0) and (iMainLeft < Screen.Width) then
              Self.Left := iMainLeft;

      end;
      sMainWidth := ReadIniFromReg(XcSubKeyIs, cMainWidth);
      if sMainWidth <> '' then begin
          iMainWidth := strtoint(sMainWidth);
          if iMainWidth < 500 then // prevent Toolbar collapsed
              iMainWidth := 500;

          if (iMainWidth > 0) and ( iMainWidth <= Screen.Width ) then
              Self.Width := iMainWidth;

      end;
      sMainHeight := ReadIniFromReg(XcSubKeyIs, cMainHeight);
      if sMainHeight <> '' then begin
          iMainHeight := strtoint(sMainHeight);
          if (iMainHeight > 0) and ( iMainHeight <= Screen.Height) then
              Self.Height := iMainHeight;

      end;
  end;

  sWidthBrowseList := ReadIniFromReg(XcSubKeyIs, cWidthBrowseList);
  if sWidthBrowseList <> '' then begin
      iWidthBrowseList := strtoint(sWidthBrowseList);
      if (iWidthBrowseList > 30) and (iWidthBrowseList < self.Width) then
          PanMiddleLeft.Width := iWidthBrowseList;

  end
  else begin 
      if screen.PixelsPerInch = 96 then // small font
          PanMiddleLeft.Width := 198;
          
  end;

  sHeightPnM := ReadIniFromReg(XcSubKeyIs, cHeightPnM);
  if sHeightPnM <> '' then begin
      iHeightPnM := strtoint(sHeightPnM);
      if (iHeightPnM > 200) and ( (iHeightPnM - 100) < self.Height ) then
          pnM.Height := iHeightPnM;

  end;

 { sUpDownBtnViewStyle := ReadIniFromReg(XcSubKeyIs, cUpDownBtnViewStyle);
  if sUpDownBtnViewStyle = '1' then begin
      btnViewStyle.Down := True;
      btnViewStyle.Click;
  end
  else
      btnViewStyle.Click; }

 { sUpDownBtnAdvancedView := ReadIniFromReg(XcSubKeyIs, cUpDownBtnAdvancedView);
  if sUpDownBtnAdvancedView = '1' then begin
      btnAdvancedView.Down := True;
      btnAdvancedView.Click;
  end
  else
      btnAdvancedView.Click; }

  SetWidthOfMainListColumns;
    // restore favorite dir
  SetLength(mnuFavoriteDir, iCntFavorite);
  for i := 0 to iCntFavorite -1 do begin
    sPath := ReadIniFromReg(XcSubKeyIs + cFavouriteFolders, inttostr(i));
    if sPath = '' then break;
    mnuFavoriteDir[i] := TMenuItem.Create(Self);
    mnuFavoriteDir[i].Caption := sPath;
    pmFavoriteDir.Items.Insert(i,mnuFavoriteDir[i]);
    mnuFavoriteDir[i].OnClick := FavoriteDirOnClick;
    mnuFavoriteDir[i].AutoCheck := True;
  end;

  sMainMenuSort := ReadIniFromReg(XcSubKeyIs, cMainMenuSort);
  if sMainMenuSort <> '' then begin
      for i := 0 to mnuSubSort.Count -1 do begin
        if mnuSubSort.Items[i].Caption = sMainMenuSort then begin
            mnuSubSort.Items[i].Checked := True;
            break;
        end;
      end;
  end;
end;

procedure TfrmMain.RestoreRecentUnzipFilenameToMenu;
var
  i: integer;
  sPathFilename: string;
begin
  OpenedFilesPopup.Items.Clear;
  for i := 0 to cCountOfRecentUnzipFiles -1 do begin
    sPathFilename := frmMain.ReadIniFromReg(sSubKeyRecentUnzip, inttostr(i));
    if sPathFilename <> '' then begin
        mnuNewMake[i] := TMenuItem.Create(OpenedFilesPopup);
        mnuNewMake[i].Caption := sPathFilename;
        OpenedFilesPopup.Items.Insert(i,mnuNewMake[i]);
        mnuNewMake[i].OnClick := OpenedFileOnClick;
        mnuNewMake[i].AutoCheck := True
    end;
  end;
end;

procedure TfrmMain.CheckRegisteredUser;
var
  sRegisteredUser: string;
  fReg: file of TRegisteredUser;
  MyUser: TRegisteredUser;
begin
  // handle register's message
  FileMode := 2;
  AssignFile(fReg, XsAppPath + XcIniRegFile);
  try
    Reset(fReg);  // open an old file and access read or write, error occurred when file doesn't exist
    //Rewrite(f);  // open a new file and access read or write
    //BlockWrite(f,buf,cnt,cnt);
    try
      Seek(fReg,0);
      Read(fReg,MyUser);
      if MyUser.sHeader = '[RegisteredBy]' then

      else begin
          sRegisteredUser := InputBox(XcProgramName,
          'Please enter a name to register this product! It is free! :', 'User');
          sRegisteredUser := copy(sRegisteredUser, 1, 255);
          SaveRegUser(nil, '[RegisteredBy]', sRegisteredUser, 0);
      end;

    except
      On EInOutError do begin
          sRegisteredUser := InputBox(XcProgramName,
          'Please enter a name to register this product! It is free! :', 'User');
          sRegisteredUser := copy(sRegisteredUser, 1, 255);
          SaveRegUser(nil, '[RegisteredBy]', sRegisteredUser, 0);
      end;
    end;
    CloseFile(fReg);
  except
    on E: Exception do begin
        if E.Message = 'File not found' then begin
            sRegisteredUser := InputBox(XcProgramName,
            'Please enter a name to register this product! It is free! :', 'User');
            sRegisteredUser := copy(sRegisteredUser, 1, 255);
            SaveRegUser(nil, '[RegisteredBy]', sRegisteredUser, 0);
        end;
    end;
  end;
end;

procedure TfrmMain.CheckZipFilesAssociated;
var
  //Reg: TRegistry;
  sDefaultZipMaster: string;
  bCheckAssociate: boolean;
  RegIni: TRegIniFile;
  sGetKeyName: string;
  wRetAns: Word;
  bHaHaZip_KeyExist: boolean;
  sCurUserName, sWho: string;
begin
  sCurUserName := UserID;
  if sCurUserName = '' then
      sCurUserName := XcDefault;

  sWho := ReadRegLocalMachine(XcSubKeyIs, XcInstalledBy);
  if AnsiLowercase(sCurUserName) <> AnsiLowercase(sWho) then exit;

  bCheckAssociate := True; // assume check association every time program starts
  with TRegistry.Create do
  try // Read Initial Reg to need check file association or not
    RootKey := HKEY_CURRENT_USER;
    if OpenKey(XcSubKeyIs, False) then begin
        sGetKeyName := ReadString(cKeyNamechkAss);
        if sGetKeyName = 'False' then
            bCheckAssociate := False
        else
            bCheckAssociate := True;

        CloseKey;
    end;
  finally
    Free;
  end;

  // Does it need to check file association?
  if bCheckAssociate then begin
      RegIni := TRegIniFile.Create( '' );
      try // Does Class ".zip" already point to a Class HaHazip in Registry? 
        RegIni.RootKey := HKEY_CLASSES_ROOT;
        if RegIni.OpenKey('.zip',False) then begin
            sDefaultZipMaster := RegIni.ReadString('','','');
            RegIni.CloseKey;
        end;
      finally
        RegIni.Free;
      end;

      bHaHaZip_KeyExist := False;
      with TRegistry.Create do begin
        try // One more check! Is Class HaHaZip exist?
          RootKey := HKEY_CLASSES_ROOT;
          if OpenKey(XcProgramName, False) then begin
              if ReadString('') = 'Zip' then
                  bHaHaZip_KeyExist := True;
                  
              CloseKey;
          end;
        finally
          Free;
        end;
      end;

      if (sDefaultZipMaster<>XcProgramName) or (not bHaHaZip_KeyExist) then begin
          beep;
          wRetAns := MessageDlg(XcProgramName + ' is not associated with ' +
          'zip files. Do you want to change it now?' + #13#10 +
          'If pressing button No, this window will not show in next time.',
          mtConfirmation, [mbYes, mbNo], 0); //= mrYes then begin
          with TRegistry.Create do
            try
              RootKey := HKEY_CURRENT_USER;
              if OpenKey(XcSubKeyIs, True) then begin
                  if wRetAns = mrYes then begin
                      WriteString(cKeyNamechkAss, 'True');
                      // below doing association
                      RegisterFileType('.zip', XcProgramName, 'Zip',
                      XsAppPath + XcExeProgFilename, 0, True);
                  end
                  else
                      WriteString(cKeyNamechkAss, 'False');

                  CloseKey;
              end
            finally
              Free;
            end;
      end;
  end;
end;

{procedure TfrmMain.Set_MainList_ToUse_SystemIcons;
var
  FileInfo: TSHFileInfo;
  ImageListHandle: THandle;
begin
  ImageListHandle := SHGetFileInfo('C:\',
                           0,
                           FileInfo,
                           SizeOf(FileInfo),
                           SHGFI_SYSICONINDEX or SHGFI_SMALLICON);
  SendMessage(ListView1.Handle, LVM_SETIMAGELIST, LVSIL_SMALL, ImageListHandle);

 { ImageListHandle := SHGetFileInfo('C:\',
                           0,
                           FileInfo,
                           SizeOf(FileInfo),
                           SHGFI_SYSICONINDEX or SHGFI_LARGEICON);

  SendMessage(ListView1.Handle, LVM_SETIMAGELIST, LVSIL_NORMAL, ImageListHandle); }
//end; }

procedure TfrmMain.SaveRegUser(Sender: TObject; sHeader: string;
 sRegUser: string; iPointer: integer);
var
  fReg: file of TRegisteredUser;
  MyUser: TRegisteredUser;
begin
  MyUser.sHeader := sHeader;
  MyUser.sRegUser := sRegUser;

  //OldFileMode := FileMode;  //store Current file mode
  FileMode := 2;  // read and write
  AssignFile(fReg, XsAppPath + XcIniRegFile);
  try
    Rewrite(fReg);  // open an old file and access read or write, error occurred when file doesn't exist
  except on E: Exception do
    begin
      ShowMessage('Error occurred when saving the data to ' + XcIniRegFile +
      Chr(13) + Chr(10) + 'Error # '  + E.Message);
      exit;
    end;
  end;
  Seek(fReg,iPointer);
  Write(fReg,MyUser);
  CloseFile(fReg);
end;

procedure TfrmMain.mnuTestClick(Sender: TObject);
var
  sNumOfFiles: string;
  MO: TMessageOutput;
  PN: TPanel;
  LBL: TLabel;
begin
  if ZipFName.Caption = '' then begin
      beep;
      MessageDlg('Error, source zip filename is empty', mtError, [mbOK], 0);
      exit;
  end;

  if ZipMaster1.Count < 1 then begin
      beep;
      MessageDlg('Nothing to test', mtError, [mbOK], 0);
      exit;
  end;

  if Zipmaster1.Zipfilename = '' then
      exit;

  MessageOutput.memOutput.Clear;
 // MessageOutput.Show;
  //ZipMaster1Message(self, 0, 'Beginning test of ' + ZipMaster1.ZipFileName);
  with ZipMaster1 do begin
    Password := ''; // important to clear previous password
    PasswordReqCount := 3; // important if show password diaglog box
    FSpecArgs.Clear;
    ExtrOptions := ExtrOptions + [ExtrTest];
    FSpecArgs.Add('*.*');           // Test all the files in the .zip
    Verbose := False;
    Trace := False;
    Screen.Cursor := crHourGlass;
    try
      Extract;                        // This will really do a test
    finally
      Screen.Cursor := crDefault;
    end;
  end;

  with ZipMaster1 do begin
    If SuccessCnt = DirOnlyCount + Count Then begin
        if SuccessCnt = 1 then
            sNumOfFiles := 'file'
        else
            sNumOfFiles := 'files';
            
        ShowMessage('Total ' + IntToStr(DirOnlyCount + Count) +
        ' ' + sNumOfFiles + ' tested OK');
    end
    Else
        ShowMessage( 'Error, bad files or skipped test: Total = ' +
        IntToStr(DirOnlyCount + Count - SuccessCnt) );
  end;


  MO := TMessageOutput.Create(self);
  with MO do
  try
    MO.AutoSize := False;
    MO.Height := 540;
    MO.AutoSize := True;

    PN := TPanel.Create(MO);
    PN.Parent := MO;

    LBL := TLabel.Create(MO);
    LBL.Parent := PN;
    try
      PN.Align := alTop;
      PN.Visible := True;
      PN.BevelOuter := bvNone;
      PN.Height := 180;
      PN.Caption := '';

      LBL.AutoSize := False;
      LBL.Visible := True;
      LBL.WordWrap := True;
      LBL.Left := 0;
      LBL.Top := 0;
      LBL.Width := PN.Width -2;
      LBL.Height := PN.Height -2;
      LBL.Caption := 'Disclaimer:' + #10#13 + #10#13 +
                     'Docx, pptx or xlsx files are a collections of ' +
                     'conventionally zipped XML files and media like pictures ' +
                     'such as JPG images.  This program''s ''''Test'''' and ' +
                     '''''Repair'''' features will repair the zip members ' +
                     'where possible.  After repair, you should try to open the ' +
                     'file again in both Word and this program. However ' +
                     'experience has shown that repairing the zip file may not ' +
                     'improve the salvaging of any additional data or ' +
                     'formatting.  The native non-repaired text unzipping and ' +
                     'text extraction features of this program are likely to be ' +
                     'the best recovery possible';

      memOutput.Lines.Text := Messageoutput.memOutput.Lines.Text;
      ShowModal;
    finally
      LBL.Free;
      PN.Free;
    end;
  finally
    Free;
  end;
end;

// This is the "OnMessage" event handler
Procedure TfrmMain.ZipMaster1Message(Sender: TObject; ErrCode: Integer; Message: String);
const
  cCheckDirErr = 'checkdir error:';
Begin
  MessageOutput.memOutput.Lines.Append(Message);
  PostMessage(MessageOutput.memOutput.Handle, EM_SCROLLCARET, 0, 0);
  If (ErrCode > 0) And Not ZipMaster1.Unattended Then begin
      //ShowMessage('Error Msg: ' + Message);
      sErrorMessage := sErrorMessage + Message + #13+#10;
      if (ErrCode >= 11036) and (ErrCode <= 11038) then // DS_NoValidZip , DS_FirstInSet, DS_NotLastInSet
          btnClose.Click;

      //if ErrCode = 11032 then // DS_NoDiskSpace
      //    ZipMaster1.Cancel := True;
  end
  else if (ErrCode = 0) and ( copy(Message, 1, length(cCheckDirErr)) =
  cCheckDirErr ) then
      sErrorMessage := sErrorMessage + Message + #13+#10;

End;

procedure TfrmMain.SetStatusOfPopupUnzip(Sender: TObject);
begin
  // PopupMenu1
  mnuOpenWith.Enabled := False; // Open With
  mnuOpen.Enabled := False; // Open
  mnuRename.Enabled := False; // Rename
  mnuProperties.Enabled := False; // Properties
  mnuDelete.Enabled := False; // Delete
  mnuAddFilesToZip.Enabled := False; // Add
  mnuExtract.Enabled := False; // Extract
  mnuTest.Enabled := False; // Test
  mnuViewComment.Enabled := False; // View comment
  pm1AddComment.Enabled := False; // Add comment
  mnuSelectAll.Enabled := False; // Select All
  mnuInvertSelection.Enabled := False; // Invert Selection
  mnuCloseAchive.Enabled := False; // Close archive

  // MainMenu1
  mnuSubMoveArchive.Enabled := False; // Move archive
  mnuSubCopyArchive.Enabled := False; // Copy archive
  mnuSubRenameArchive.Enabled := False; // Rename archive
  mnuSubDeleteArchive.Enabled := False; // Delete archive
  mnuSubClose1.Enabled := False; // Close archive
  mnuSubOpenWith.Enabled := False; // Open With
  mnuSubOpen.Enabled := False; // Open
  mnuSubRename.Enabled := False; // Rename
  mnuSubSearch.Enabled := False; // Search
  mnuSubProperties.Enabled := False; // Properties
  mnuSubDelete.Enabled := False; // Delete
  mnuSubAdd.Enabled := False; // Add
  mnuSubExtract.Enabled := False; // Extract
  mnuSubTest.Enabled := False; // Test
  mnuSFXToZip.Enabled := False; // Convert Exe To Zip
  mnuDiskSpan.Enabled := False; // Disk Span
  mnuCombine.Enabled := False; // Combine Split Zip
  mnuVirusScan.Enabled := False; // Virus Scan
  mnuSubViewComment.Enabled := False; // View comment
  mnuAddComment.Enabled := False; // Add comment
  mnuSubSelectAll.Enabled := False; // Select All
  mnuSubInvertSelection.Enabled := False; // Invert Selection

  if (ZipFName.Caption <> '') then begin
      if VT.SelectedCount > 0 then begin
          if VT.SelectedCount = 1 then begin
             { if (ListView1.Selected.Caption = '..') then exit; // is up one level

              if (ListView1.Selected.SubItems.Strings[0] <> '') and
              (ListView1.Selected.SubItems.Strings[1] <> '') then begin }
                  // PopupMenu1
                  mnuOpenWith.Enabled := True; // Open With
                  mnuOpen.Enabled := True; // Open
                  mnuRename.Enabled := True; // Rename
                  mnuProperties.Enabled := True; // Properties
                  // MainMenu1
                  mnuSubOpenWith.Enabled := True; // Open With
                  mnuSubOpen.Enabled := True; // Open
                  mnuSubRename.Enabled := True; // Rename
                  mnuSubProperties.Enabled := True; // Properties
              //end;
          end;
          // PopupMenu1
          mnuDelete.Enabled := True; // Delete
          mnuExtract.Enabled := True; // Extract
          mnuInvertSelection.Enabled := True; // Invert Selection
          // MainMenu1
          mnuSubDelete.Enabled := True; // Delete
          mnuSubExtract.Enabled := True; // Extract
          mnuSubInvertSelection.Enabled := True; // Invert Selection
      end
      else begin
          // PopupMenu1
          mnuAddFilesToZip.Enabled := True; // Add
          mnuExtract.Enabled := True; // Extract
          mnuTest.Enabled := True; // Test
          mnuViewComment.Enabled := True; // View comment
          pm1AddComment.Enabled := True; // Add comment
          mnuSelectAll.Enabled := True; // Select All
          mnuCloseAchive.Enabled := True; // Close archive

          // MainMenu1
          mnuSubMoveArchive.Enabled := True; // Move archive
          mnuSubCopyArchive.Enabled := True; // Copy archive
          mnuSubRenameArchive.Enabled := True; // Rename archive
          mnuSubDeleteArchive.Enabled := True; // Delete archive
          mnuSubClose1.Enabled := True; // Close archive
          mnuSubSearch.Enabled := True; // Search
          mnuSubAdd.Enabled := True; // Add
          mnuSubExtract.Enabled := True; // Extract
          mnuSubTest.Enabled := True; // Test
          mnuSFXToZip.Enabled := True; // Convert Exe To Zip
          mnuDiskSpan.Enabled := True; // Disk Span
          mnuCombine.Enabled := True; // Combine Split Zip
          mnuVirusScan.Enabled := TRue; // Virus Scan
          mnuSubViewComment.Enabled := True; // View comment
          mnuAddComment.Enabled := True; // Add comment
          mnuSubSelectAll.Enabled := True; // Select All
      end;
  end;
  Insert_HistoryOpenWithApp_ToMenu;
end;

procedure TfrmMain.btnAddFilesClick(Sender: TObject);
begin
  mnuAddFilesToZipClick(nil);
end;

procedure TfrmMain.btnUnZipClick(Sender: TObject);
begin
  mnuExtractClick(nil);
end;

procedure TfrmMain.btnSfxWizardClick(Sender: TObject);
begin
  {

  with TdlgConvertToSFX.Create(self) do
  try
    ShowModal;
  finally
    Free;
  end;

  }
end;

procedure TfrmMain.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.mnuSubNew1Click(Sender: TObject);
begin
  btnNewClick(btnNew);
end;

procedure TfrmMain.mnuSubOpen1Click(Sender: TObject);
begin
  OpenZipFile;
end;

procedure TfrmMain.mnuSubClose1Click(Sender: TObject);
begin
  mnuCloseAchiveClick(nil);
end;

procedure TfrmMain.mnuSubExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.mnuSubAddClick(Sender: TObject);
begin
  btnAddFiles.Click;
end;

procedure TfrmMain.mnuSubExtractClick(Sender: TObject);
begin
  mnuExtractClick(nil);
end;

procedure TfrmMain.mnuSubTestClick(Sender: TObject);
begin
  mnuTestClick(nil);
end;

procedure TfrmMain.mnuSubSFXClick(Sender: TObject);
begin
  btnSFXWizard.Click;
end;

procedure TfrmMain.mnuSelectAllClick(Sender: TObject);
begin
  if (ZipFName.Caption <> '') then begin
      //ListView1.SelectAll;
      VT.SelectAll(False);
      VT.Refresh;
  end;
end;

procedure TfrmMain.mnuInvertSelectionClick(Sender: TObject);
//var
//  i: integer;
begin
  if (ZipFName.Caption <> '') then begin
      VT.InvertSelection(False);
      VT.Refresh;
     { for i := 0 to ListView1.Items.Count -1 do begin
        ListView1.Items[i].Selected := not ListView1.Items[i].Selected;
      end; }
  end;
end;

procedure TfrmMain.mnuSubSelectAllClick(Sender: TObject);
begin
  if (ZipFName.Caption <> '') then begin
      VT.SelectAll(False);
      VT.SetFocus;
  end;
end;

procedure TfrmMain.mnuSubInvertSelectionClick(Sender: TObject);
var
  i: integer;
begin
  if (ZipFName.Caption <> '') then begin
      VT.InvertSelection(False);
      VT.SetFocus;
     { for i := 0 to ListView1.Items.Count -1 do begin
        ListView1.Items[i].Selected := not ListView1.Items[i].Selected;
      end;
      ListView1.SetFocus; }
  end;
end;

procedure TfrmMain.mnuSubPreferencesClick(Sender: TObject);
begin
  with TfrmSettings.Create(Self) do
  try
    showmodal;
  finally
    free;
  end;
end;

procedure TfrmMain.mnuSubShowBrowseClick(Sender: TObject);
begin
  Splitter1.Visible := True; // must show this first
  PanMiddleLeft.Visible := True;
  if bResizingCols then exit;
  if Main.bCheckMainLstAutoSize then
      ResizeWidthOfMainListColumns;

end;

procedure TfrmMain.smHelpFileClick(Sender: TObject);
var
  RetValue: integer;
  sPathFilename: string;
begin
  sPathFilename := XsAppPath + 'hhzhelp.chm';
  if FileExists(sPathFilename) then begin
      RetValue := ShellExecute(handle, 'open', PChar(sPathFilename),
      '', '', SW_SHOW); // if success, return different value where depends on program

      if RetValue = 31 then begin // Has not the filename association?
          beep;
          MessageDlg('No application is associated with this kind of file extension.' +
          ' Failing to open file.', mtError, [mbOK], 0);
      end;
  end
  else begin
    beep;
    MessageDlg('Help file not found.', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmMain.smAboutClick(Sender: TObject);
begin
  with TAboutBox.Create(Self) do
  try
    ShowModal;
  finally
    free;
  end;
end;

procedure TfrmMain.spCloseBrowsePanelClick(Sender: TObject);
begin
  if btnShowVirtual.Down then begin
      beep;
      MessageDlg('Browse list cannot be closed when virtual list view mode is on.',
      mtInformation, [mbOK], 0);
      exit;
  end;

  PanMiddleLeft.Visible := False;
  Splitter1.Visible := False;
  if bResizingCols then exit;
  if Main.bCheckMainLstAutoSize then
      ResizeWidthOfMainListColumns;

end;

procedure TfrmMain.btnUp1Click(Sender: TObject);
var
 // sPath, sTemp: string;
 // iRetPos, iLastPos: integer;
 Node: PVirtualNode;
begin
 { sPath := ShellDir1.Path;

  if sPath <> '' then begin
      sTemp := sPath;
      iLastPos := 0;
      repeat
        iRetPos := AnsiPos('\', sTemp);
        if iRetPos > 0 then begin
            iLastPos := iLastPos + iRetPos;
            sTemp := copy(sTemp, iRetPos +1, length(sTemp) -iRetPos)
        end;
      until iRetPos = 0;
  end;

  if iLastPos > 0 then begin
      sPath := copy(sPath, 1, iLastPos -1);
      ShellDir1.Path := sPath;
  end;

  ShellDir1.SetFocus; }

  if not Assigned(VETree.FocusedNode) then
      exit;

  Node := VETree.FocusedNode;

  if VETree.GetNodeLevel(Node) > 0 then begin
      Node := Node.Parent;
      VETree.ClearSelection;
      VETree.FocusedNode := Node;
      VETree.Selected[Node] := True;

      VETree.Refresh;
      VETree.SetFocus;

  end;
end;

procedure TfrmMain.btnRefresh1Click(Sender: TObject);
begin
  VETree.RefreshTree(True);
end;

procedure TfrmMain.FilterComboBox1Click(Sender: TObject);
begin
 { if FilterComboBox1.ItemIndex = 0 then
      ShellDir1.Mask := '*.docx'
  else if FilterComboBox1.ItemIndex = 1 then
      ShellDir1.Mask := '*.zip'
  else if FilterComboBox1.ItemIndex = 2 then
      ShellDir1.Mask := '*.exe'
  else if FilterComboBox1.ItemIndex = 1 then
      ShellDir1.Mask := '*.*';

  btnRefresh1.Click; }
  //ShellDir1.SetFocus;

  btnRefresh1.Click;
end;


procedure TfrmMain.FormCanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
begin
  StatusBar1.Panels[0].Width := StatusBar1.Width - 220;
  StatusBar1.Panels[1].Width := StatusBar1.Width - StatusBar1.Panels[0].Width -
  110;
  StatusBar1.Panels[2].Width := StatusBar1.Panels[1].Width;
end;

procedure TfrmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
//var
 { Handle: HKey;
  Res: integer;
  Disposition: Integer; }
begin
  SaveStatusOfProgram;
 { with TRegistry.Create do
  try
    RootKey := HKEY_CURRENT_USER;
    if OpenKey(XcSubKeyIs, True) then begin
        WriteString('RootDir', XsAppPath);
        CloseKey;
    end;
  finally
    Free;
  end; }

 { Res := RegCreateKeyEx(HKEY_LOCAL_MACHINE, PChar(XcSubKeyIs), 0, '',
  REG_OPTION_NON_VOLATILE, KEY_READ or KEY_WRITE, nil, Handle, @Disposition);
  if Res = 0 then begin // Is creation of key success?
      Res := RegSetValueEx(Handle, Pchar('RootDir'), 0, REG_SZ,
      Pchar(XsAppPath), length(XsAppPath)+1); // try to write data
      RegCloseKey(Handle);
      //if Res <> 0 then
      //    ShowMessage('Error updating registry');

  end; }
end;

procedure TfrmMain.SaveStatusOfProgram;
var
  sWindowStateIs, sBtnViewStyle, sColName, sBtnAdvancedView: string;
  iWidthName, iWidthModified, iWidthSize, iWidthRatio: integer;
  iWidthPacked, iWidthFileType, iWidthCRC, iWidthAttributes, iWidthPath: integer;
  i, iWidthBrowseList, iHeightPnM: integer;
begin
  iWidthName := 0; iWidthModified := 0; iWidthSize := 0; iWidthRatio := 0;
  iWidthPacked := 0; iWidthFileType := 0; iWidthCRC := 0;
  iWidthAttributes := 0; iWidthPath := 0;
  for i := 0 to VT.Header.Columns.Count -1 do begin
    sColName := VT.Header.Columns[i].Text;

    if AnsiPos('Name', sColName) > 0 then
        iWidthName := VT.Header.Columns[i].Width
    else if AnsiPos('Modified', sColName) > 0 then
        iWidthModified := VT.Header.Columns[i].Width
    else if AnsiPos('Size', sColName) > 0 then
        iWidthSize := VT.Header.Columns[i].Width
    else if AnsiPos('Ratio', sColName) > 0 then
        iWidthRatio := VT.Header.Columns[i].Width
    else if AnsiPos('Packed', sColName) > 0 then
        iWidthPacked := VT.Header.Columns[i].Width
    else if AnsiPos('File Type', sColName) > 0 then
        iWidthFileType := VT.Header.Columns[i].Width
    else if AnsiPos('CRC', sColName) > 0 then
        iWidthCRC := VT.Header.Columns[i].Width
    else if AnsiPos('Attributes', sColName) > 0 then
        iWidthAttributes := VT.Header.Columns[i].Width
    else if AnsiPos('Path', sColName) > 0 then
        iWidthPath := VT.Header.Columns[i].Width;

  end;

  if Self.WindowState = wsMaximized then
      sWindowStateIs := 'wsMaximized'
  else if Self.WindowState = wsNormal then begin
      WriteIniToReg(XcSubKeyIs, cMainTop, inttostr(Self.Top));
      WriteIniToReg(XcSubKeyIs, cMainLeft, inttostr(Self.Left));
      WriteIniToReg(XcSubKeyIs, cMainWidth, inttostr(Self.Width));
      WriteIniToReg(XcSubKeyIs, cMainHeight, inttostr(Self.Height));
  end;
  WriteIniToReg(XcSubKeyIs, cMainWindowState, sWindowStateIs);
  if not bCheckMainLstAutoSize then begin
      if iWidthName > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthName, inttostr(iWidthName) );
      if iWidthModified > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthModified, inttostr(iWidthModified) );
      if iWidthSize > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthSize, inttostr(iWidthSize) );
      if iWidthRatio > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthRatio, inttostr(iWidthRatio) );
      if iWidthPacked > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthPacked, inttostr(iWidthPacked) );
      if iWidthFileType > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthFileType, inttostr(iWidthFileType) );
      if iWidthCRC > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthCRC, inttostr(iWidthCRC) );
      if iWidthAttributes > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthAttributes, inttostr(iWidthAttributes) );
      if iWidthPath > 0 then
          WriteIniToReg( XcSubKeyIs, cWidthPath, inttostr(iWidthPath) );
  end;

  if PanMiddleLeft.Visible then begin
      iWidthBrowseList := PanMiddleLeft.Width;
      if (iWidthBrowseList > 15) and (iWidthBrowseList < Screen.Width) then
          WriteIniToReg(XcSubKeyIs, cWidthBrowseList, inttostr(iWidthBrowseList));
          
  end;

  if pnM.Visible then begin
      iHeightPnM := pnM.Height;
      if (iHeightPnM > 200) and (pnB.Height > 60) then
          WriteIniToReg(XcSubKeyIs, cHeightPnM, inttostr(iHeightPnM));

  end;

 { if btnViewStyle.Down then
      sBtnViewStyle := '1'
  else
      sBtnViewStyle := '0';

  WriteIniToReg(XcSubKeyIs, cUpDownBtnViewStyle, sBtnViewStyle); }

  if btnAdvancedView.Down then
      sBtnAdvancedView := '1'
  else
      sBtnAdvancedView := '0';

  WriteIniToReg(XcSubKeyIs, cUpDownBtnAdvancedView, sBtnAdvancedView);

end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  // If there has extracted files in temp, try to remove
  if bExtractedFiles and (edShadow.Text <> '') then
      Remove_TempFiles_UnderWinTemp;

  ABitmap.Free;
  BBitmap.Free;
  TeBitmap.Free;
  ClipBitmap.Free;

  XslOfficeDoc.Free;
  XslSheetsFilenames.Free;
  XslSharedStr.Free;
  XslAlphabet.Free;
  XslRID.Free;

  slZipStrings.Free;
  FDataList.Free;
  XmlParser.Free;
  DrawXmlParser.Free;
  //ilSysIcons.BkColor := APrevColor;
  //AIcon.Free;
  //ilMainList.Free;
end;

procedure TfrmMain.FormPaint(Sender: TObject);
begin
  StatusBar1.Repaint;
end;

procedure TfrmMain.FormResize(Sender: TObject);
begin
  if Self.Width < 550 then // prevent Toolbar collapsed
      Self.Width := 550;

  if bResizingCols then exit;
  if Main.bCheckMainLstAutoSize then
      ResizeWidthOfMainListColumns;

end;

function TfrmMain.Get_SysIconIndex_Of_Given_FileExt(sWhatFileExt: string;
 var sFileTypeIs: string): integer;
var
  FileInfo: TSHFileInfo;
begin
  // Retrieve icon and typename for the directory: FILE_ATTRIBUTE_DIRECTORY, SHGFI_TYPENAME Or SHGFI_SYSICONINDEX
  // Retrieve icon and typename for the file: FILE_ATTRIBUTE_NORMAL, SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX
  SHGetFileInfo( PChar( sWhatFileExt ), FILE_ATTRIBUTE_NORMAL, FileInfo,
  SizeOf(FileInfo), SHGFI_SYSICONINDEX or SHGFI_TYPENAME or SHGFI_USEFILEATTRIBUTES ); // FILE_ATTRIBUTE_READONLY, Or SHGFI_SMALLICON Or SHGFI_OPENICON
  Result := FileInfo.iIcon;
  sFileTypeIs := FileInfo.szTypeName;
end;

procedure TfrmMain.Zipping;
begin
  if (edtZipFileName.Text = '') then begin
      beep;
      MessageDlg('No target zip filename.', mtError, [mbOK], 0);
      exit;
  end;

  if (slZipStrings.Count = 0) then begin
      beep;
      MessageDlg('No selected file.', mtInformation, [mbOK], 0);
      exit;
  end;

  if ExtractFileExt(edtZipFileName.Text) = '' then
      edtZipFileName.Text := edtZipFileName.Text + '.zip';

  Gauge1.MaxValue := slZipStrings.Count -1;
  Gauge1.Visible := True;
  Gauge2.Visible := True;
  MessageOutput.memOutput.Clear;
  sErrorMessage := '';
  StartZip(edtZipFileName.Text, slZipStrings);
  Gauge1.Progress := 0;
  Gauge1.Visible := False;

  if sErrorMessage <> '' then
      PopUp_ErrorMessage('Error occurred when adding file(s).');
end;

procedure TfrmMain.Refresh_ItemsOnMainList;
var
  k, iIndex: integer;
  sNumFiles: string;
  NewCol: TCollectionItem;
begin
  if ZipMaster1.ZipFileName = '' then exit;

  if bCol_AddedOrRemoved_OnMainLst then begin
      for k := VT.Header.Columns.Count -2 downto 5 do begin // remove extra columns
        VT.Header.Columns.Delete(k);
      end;
      if bCheckType then begin
          // insert to last column, new caption will replace old one.
          // If insert to ListView1.Columns.Count -1, Column index will be different. ListITem.SubItems.Add will depend on what col created early.
          NewCol := VT.Header.Columns.Insert(VT.Header.Columns.Count -1); // before Col Path, if u want last col, don't -1
          VT.Header.Columns[NewCol.Index].Text := 'File Type';
          VT.Header.Columns[NewCol.Index].Alignment := taRightJustify;
          VT.Header.Columns[NewCol.Index].MinWidth := 60;
          VT.Header.Columns[NewCol.Index].Width := iColWidthFileType;
      end;
      if bCheckCRC then begin
          NewCol := VT.Header.Columns.Insert(VT.Header.Columns.Count -1); // before Col Path, if u want last col, don't -1
          VT.Header.Columns[NewCol.Index].Text := 'CRC';
          VT.Header.Columns[NewCol.Index].Alignment := taRightJustify;
          VT.Header.Columns[NewCol.Index].MinWidth := 60;
          VT.Header.Columns[NewCol.Index].Width := iColWidthCRC;
      end;
      if bCheckAttributes then begin
          NewCol := VT.Header.Columns.Insert(VT.Header.Columns.Count -1); // before Col Path, if u want last col, don't -1
          VT.Header.Columns[NewCol.Index].Text := 'Attributes';
          VT.Header.Columns[NewCol.Index].Alignment := taRightJustify;
          VT.Header.Columns[NewCol.Index].MinWidth := 60;
          VT.Header.Columns[NewCol.Index].Width := iColWidthAttributes;
      end;
      SetBool_CntOf_ColumnChanged_OnMainLst(False);
     { ListView1.Column[ListView1.Columns.Count -1].Caption := 'Path'; // write back the caption Path
      ListView1.Column[ListView1.Columns.Count -1].Alignment := taRightJustify;
      ListView1.Column[ListView1.Columns.Count -1].Width := iColWidthPath; }
  end;

  Screen.Cursor := crHourGlass;
  try
    Gauge1.Visible := True;
    Gauge1.MinValue := 0;
    Gauge1.Progress := 0;
    Gauge1.MaxValue := ZipMaster1.ZipContents.Count;

    TreeView1.Items.Clear;
    //VT.BeginUpdate;
    VT.Clear; // <= faster than ListView1.Clear
    AddNodesToVirtualTree(FDataList.Count);
    VT.RootNodeCount := FDataList.Count;
  finally
    //VT.EndUpdate;
    Screen.Cursor := crDefault;
    Gauge1.Visible := False;
    Gauge1.Progress := 0;
  end;

  if ZipMaster1.Count > 1 then
      sNumFiles := 'Objects = '
  else
      sNumFiles := 'Object = ';

  StatusBar1.Panels[0].Text := ZipMaster1.ZipFileName;

  if btnAdvancedView.Down then
      StatusBar1.Panels[1].Text := sNumFiles + inttostr(ZipMaster1.Count)
  else
      StatusBar1.Panels[1].Text := '';
      
  StatusBar1.Panels[2].Text := 'Size = ' + inttostr(ZipMaster1.ZipFileSize
  div 1024) + ' KB' ;
  ClearColumnsImage_OnlvMain;
  iIndex := Get_ColIndex_ToSort_FromMenuSortBy;
  if iIndex <> -1 then
      //ListView1.OnColumnClick(ListView1, ListView1.Column[iIndex]);
      VT.OnHeaderClick(VT.Header, iIndex, mbLeft, [], 0, 0);

  if Main.bCheckMainLstAutoSize then
      ResizeWidthOfMainListColumns;

  SetStatusOfPopupUnzip(nil);
end;

procedure TfrmMain.mnuViewCommentClick(Sender: TObject);
begin
  if ZipFName.Caption = '' then begin
      beep;
      MessageDlg('Error, source zip filename is empty', mtError, [mbOK], 0);
      exit;
  end;

 { if ZipMaster1.Count < 1 then begin
      beep;
      MessageDlg('Nothing to test', mtError, [mbOK], 0);
      exit;
  end; }

  if Zipmaster1.Zipfilename = '' then
      exit;

  MessageOutput.memOutput.Clear;
  MessageOutput.Show;

  MessageOutput.memOutput.Text := ZipMaster1.ZipComment;
end;

procedure TfrmMain.mnuSubViewCommentClick(Sender: TObject);
begin
  mnuViewCommentClick(nil);
end;

procedure TfrmMain.pm1AddCommentClick(Sender: TObject);
begin
  if (ZipFName.Caption = '') or (Zipmaster1.Zipfilename = '') then begin
      beep;
      MessageDlg('Error, source zip filename is empty', mtError, [mbOK], 0);
      exit;
  end;

 { if ZipMaster1.Count < 1 then begin
      beep;
      MessageDlg('Nothing to test', mtError, [mbOK], 0);
      exit;
  end; }

  with TInputMsg.Create(OWner) do begin
    try
      Memo1.Text := ZipMaster1.ZipComment;
      Memo1.SelectAll;
      showmodal;
      if not bCancelInputComment_OnInputMsgBox then begin // not cancel
          ZipMaster1.ZipComment := Memo1.Text;
          ZipMaster1.List;
          btnComment.Visible := (ZipMaster1.ZipComment <> '');
      end;
    finally
      Free;
    end;
  end;
end;

procedure TfrmMain.mnuDllVerClick(Sender: TObject);
var
  sVer: string;
begin
  sVer := 'Zipdll.dll version is ' + IntToStr(ZipMaster1.ZipVers) + #13#10#13#10
          + 'Unzip.dll version is ' + IntToStr(ZipMaster1.UnzVers);
  MessageOutput.memOutput.Clear;
  MessageOutput.Show;
  MessageOutput.memOutput.Text := sVer;
end;

procedure TfrmMain.WMDropFiles(var Msg: TMessage);
var
  hDrop    : THandle;
  FileName : array[0..254] of Char;
  iFiles   : integer;
  i        : integer;
  sPath: string;
  sFilenameIs, sExt: string;
begin
  slZipStrings.Clear;
  
  hDrop  := Msg.WParam;
try
  iFiles := DragQueryFile(hDrop, $FFFFFFFF, FileName, 254);

  for i := 0 to iFiles - 1 do begin
    DragQueryFile(hDrop, i, FileName, 254);
    sFilenameIs := FileName;
    sFilenameIs := ExpandFilename(sFilenameIs);

    sExt := AnsiLowercase(ExtractFileExt(sFilenameIs));

    // Is first file docx file? Try to load to unzip
    if ( (sExt = '.docx') or (sExt = '.pptx') ) and (i = 0) and (iFiles = 1) and (ZipFName.Caption = '') then begin
        if btnShowVirtual.Down then begin
            SetZipFile(sFilenameIs, False);
            StartUnzip(True, sFilenameIs, '', nil, False);
        end
        else
            SetZipFile(sFilenameIs, True);

        //ListView1.SetFocus;
        Remove_StringGrids_And_Tabsheets; // must do first before AfterOpeningArchive_And_Do
        AfterOpeningArchive_And_Do;

        // Are some files extracted to Windows\Temp\xxx ? True to remove it.
        if bExtractedFiles and (edShadow.Text <> '') then
            Remove_TempFiles_UnderWinTemp;

        RecentUnzipFilesToReg(sFilenameIs);
        bExtractedFiles := False;
        exit;
    end
    else if (AnsiLowercase(ExtractFileExt(sFilenameIs)) <> '.docx') and
    (i = 0) and (iFiles = 1) and (ZipFName.Caption = '') then begin
        ShowMessage('Not docx file found.');

    end;
    // Files that will be zipped. Try to load to zip
  {  if (slZipStrings.IndexOf(sFilenameIs) = -1) then begin
        if DirectoryExists(sFilenameIs) then begin // Is it a directory name?
            sPath := sFilenameIs;
            if sPath <> '' then begin
                sPath := IncludeTrailingPathDelimiter(sPath);
                ReadFilesOrFolders(sPath, '*.*');
            end;
        end
        else // a filename only
            slZipStrings.Append(sFilenameIs);
    end; }
  end;

 { if slZipStrings.Count > 0 then begin // Files will be zipped
      if ZipFName.Caption = '' then begin  // no opened file?
          XbHaveSelectedFiles := True;
          frmMain.btnNewClick(nil);
          XbHaveSelectedFiles := False;
      end
      else begin // a file was opened
          XbHaveSelectedFiles := True;
          mnuAddFilesToZipClick(nil);
          XbHaveSelectedFiles := False;
      end;
  end
  else begin
      beep;
      MessageDlg('No selected files found. If you selected folders, the ' +
      'folders may be empty.', mtInformation, [mbOK], 0);
  end; }
finally
  DragFinish(hDrop);
end;
end;

procedure Tfrmmain.PopUp_ErrorMessage(sErr: string);
begin
  beep;
  MessageOutput.memOutput.Text := '';
  MessageOutput.memOutput.Text := sErr + #13+#10 + sErrorMessage;
  MessageOutput.ShowModal;
end;

procedure TfrmMain.mnuOpenClick(Sender: TObject);
begin
  if VT.SelectedCount = 0 then begin
      beep;
      MessageDlg('Nothing selected.', mtInformation, [mbOK], 0);
      exit;
  end;
  VTDblClick(nil);
end;

procedure TfrmMain.btnHomeClick(Sender: TObject);
var
  //sPath: string;
  //iRetPos: integer;
  Node: PVirtualNode;
begin
 { sPath := ShellDir1.Path;

  if sPath <> '' then begin
      iRetPos := AnsiPos('\', sPath);
      if iRetPos > 0 then begin
          sPath := copy(sPath , 1, iRetPos -1);
          ShellDir1.Path := sPath;
          ShellDir1.SetFocus;
      end;
  end; }


  Node := VETree.GetFirst; // .TopNode is back to top depending on visible region

  //if VETree.GetNodeLevel(Node) = 0 then begin  // this is also work!!!
  if Assigned(Node) then begin
      VETree.ClearSelection;
      VETree.FocusedNode := Node;
      VETree.Selected[Node] := True;

      VETree.Refresh;
      VETree.SetFocus;
  end;

end;

procedure TfrmMain.WhatCompressionFile_WillBe_DirectlyExtracted(sFilename: string);
var
  sFileExt, sExtractToPath: string;
  sModuleName: array[0..255] of char;
begin
  sFileExt := AnsiLowercase( ExtractFileExt(sFilename) );
  if (sFileExt = '.docx') or (sFileExt = '.pptx') or (sFileExt = '.xlsx') then begin
  //if sFileExt = '.zip' then begin
      SetZipFile(sFilename, False);

      //with TfrmToFolder.Create(Self) do begin
      //try
      //  rbtnAllFiles.Checked := True;
      //  rbSelectedFiles.Enabled := False;
      //  if bFavorite_ExtractFolder then
      //      DriveView1.Directory := sFavorite_ExtractFolder;

      //  ShowModal;
      //  if Tag = 1 then begin // not cancel
          {  try
              HourGlass.Gauge1.Visible := True;
              HourGlass.Gauge1.MaxValue := frmMain.ZipMaster1.Count;
              HourGlass.lblOnProcess.Caption := 'Extracting....Please wait!' +
              #13#10 + 'Total number of file(s) = ' + inttostr(
              frmMain.ZipMaster1.Count);
              if HourGlass.Animate1.FileName <> '' then
                  HourGlass.Animate1.Active := True;

              HourGlass.Tag := 0; // 0 is extracting files
              HourGlass.show;
              HourGlass.Animate1.Repaint;
              HourGlass.lblOnProcess.Repaint; }
              //sExtractToPath := edToPath.Text;
              //BeforeExtracting(sExtractToPath, chkWithDir.Checked,
              //chkAlwaysOverwrite.Checked, rbtnAllFiles.Checked);


              sExtractToPath := ExtractFileDir(Application.ExeName);
             { if sExtractToPath = '' then begin
                  GetModuleFilename(HInstance, sModuleName, SizeOf(sModuleName)); // Full pathname where this dll exists
                  sExtractToPath := ExtractFileDir(sModuleName);
              end; }



              XsExtractToThisPath := sExtractToPath; // know path to locate document.txt file
              BeforeExtracting(sExtractToPath, True, True, False);
           { finally
              HourGlass.Close;
            end; }
        //end;
      //finally
      //  free;
      //end;
      //end;
  end;
end;

procedure TfrmMain.Files_WantToBe_Compressed_W_WO_ShowMain(sOutputTo: string; sTheDirIs: string);
begin
  XbHaveSelectedFiles := True;
  if sOutputTo = '' then begin // not directly compress to a file
      if DirectoryExists(sTheDirIs) then
          SetCurrentDir(sTheDirIs);

      Self.Show;
      btnNewClick(nil);
  end
  else begin // directly zip to specified filename and quit
      try
        HourGlass.Gauge1.Visible := True;
        HourGlass.Gauge1.MaxValue := slZipStrings.Count;
        if HourGlass.Animate1.FileName <> '' then
            HourGlass.Animate1.Active := True;
            
        HourGlass.lblOnProcess.Caption := 'Ready for adding files to archive. ' +
        'Please wait until process done!';
        HourGlass.Tag := 1; // 1 is adding files
        HourGlass.show;
        HourGlass.lblOnProcess.Repaint;
        FileOrFolder_ThroughContext_ToBeZipped(sOutputTo);
        Application.ShowMainForm := False;
        Application.Terminate;
      finally
        HourGlass.Close;
      end;
  end;
  XbHaveSelectedFiles := False;
end;

procedure TfrmMain.mnuExtractClick(Sender: TObject);
var
  sPath: string;
begin
  if ZipFName.Caption = '' then begin
      beep;
      MessageDlg('There has not opened file.', mtError, [mbOK], 0);
      exit;
  end;

  with TfrmToFolder.Create(Self) do begin
  try
    if VT.SelectedCount > 0 then begin
        rbSelectedFiles.Enabled := True;
        rbSelectedFiles.Checked := True;
    end
    else begin
        rbtnAllFiles.Checked := True;
        rbSelectedFiles.Enabled := False;
    end;

    ShowModal;
    if Tag = 1 then begin // not cancel.
        sPath := edToPath.Text;
        BeforeExtracting(sPath, chkWithDir.Checked, chkAlwaysOverwrite.Checked,
        rbtnAllFiles.Checked);
    end;
  finally
    free;
  end;
  end;
end;

procedure TfrmMain.ZipMaster1PasswordError(Sender: TObject;
  IsZipAction: Boolean; var NewPassword: String; ForFile: String;
  var RepeatCount: Cardinal; var Action: TPasswordButton);
var
  PasswordBtn: TPasswordButton;
begin
  if RepeatCount <= 1 then // count down value of PasswordRegCount = 3
      MessageDlg('Wrong password! File will be skipped.', mtError, [mbOK], 0)
  else begin
      if RepeatCount = 2 then
          beep; // warning beep if wrong password found

      PasswordBtn := ZipMaster1.GetPassword(XcProgramName + '...Input Password',
      'Protected:   ' + ExtractFilename(ForFile),
      [pwbOK, pwbCancel, pwbCancelAll], NewPassword);

      ZipMaster1.Password := NewPassword; // preserve password
      if PasswordBtn = pwbCancelAll then
          ZipMaster1.PasswordReqCount := 0;

      //NewPassword := inputbox(XcProgramName + ' Input Password',
      //'Below file is protected by password. Please input.' + #13+#10+#13+#10 +
      //ExtractFilename(ForFile) + #13+#10, '');
  end;
end;

procedure TfrmMain.mnuTerminateClick(Sender: TObject);
var
  sMsgForAdd: string;
  bIsZipping: boolean;
begin
  sMsgForAdd := '';
  bIsZipping := ZipMaster1.ZipBusy;
  if bIsZipping then // adding files
      sMsgForAdd := 'When adding files to archive, this will cause the ' +
      'target zip file corrupted.';

  if MessageDlg('Do you want to terminate the process? ' + sMsgForAdd,
  mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
      ZipMaster1.Cancel := True;
      if bIsZipping then // when adding files to archive, below boolean can tell to reload archive file
          bTerminatedProcess := True;

  end;
end;

procedure TfrmMain.mnuDiskSpanClick(Sender: TObject);
begin
  with TDiskSpan.Create(Owner) do begin
    try
      showmodal;
    finally
      free;
    end;
  end;
end;

procedure TfrmMain.mnuCombineClick(Sender: TObject);
var
  sInFile, sOutPath, sFileExt: String;
  fd:                   String;
  len :                 LongInt;
  drivetype:            LongWord;
begin
  if not ZipMaster1.IsSpanned then begin
      beep;
      MessageDlg('Current opened file is not spanned.', mtInformation, [mbOK], 0);
      exit;
  end;

  if ZipMaster1.ZipFileName = '' then begin
      beep;
      MessageDlg('There has not opened file.', mtError, [mbOK], 0);
      exit;
  end;

  sInFile := ZipMaster1.ZipFileName;
  fd := ExtractFileDrive ( sInFile ) + '\';
  drivetype := GetDriveType( PChar( fd ) );
  len := 3;

  if (drivetype = DRIVE_FIXED) or (drivetype = DRIVE_REMOTE) then begin
      sFileExt := ExtractFileExt( sInFile );
      len := Length( sInFile ) - Length( sFileExt );
      if StrToIntDef( Copy( sInFile, len - 2, 3 ), -1 ) = -1 then begin
          ShowMessage( 'This is not a valid (last)part of a spanned archive' );
          Exit;
      end;
  end;

  if SelectDirectory('Save combined zip to...', '', sOutPath) then begin
      sOutPath := IncludeTrailingPathDelimiter(sOutPath);
      if (drivetype = DRIVE_FIXED) or (drivetype = DRIVE_REMOTE) then
          sOutPath := sOutPath + ExtractFileName( Copy( sInFile, 1, len - 3 ) +
          sFileExt )
      else
          sOutPath := sOutPath + ExtractFileName( sInFile );

      ZipMaster1.ReadSpan( sInFile, sOutPath );
  end;
end;

procedure TfrmMain.ListView1ColumnClick(Sender: TObject;
  Column: TListColumn);
var
  AColumn: THdItem;
  iImgIndex: Cardinal;
begin
  if ColumnToSort = Column.Index then // sorting same column
      iImgIndex := iPreImgIndex // get last image index
  else begin
      iImgIndex := 0;
      iPreImgIndex := 0; // reset
  end;

  ColumnToSort := Column.Index;
  (Sender as TCustomListView).AlphaSort;

  ClearColumnsImage_OnlvMain;

  if iImgIndex = iIndex_Ascend_SysIcons then
      AColumn.hbm := ArrowDown
  else if iImgIndex = iIndex_Descend_SysIcons then
      AColumn.hbm := ArrowUp
  else
      AColumn.hbm := ArrowDown;

  AColumn.mask := 0; // important! reset
  AColumn.fmt := 0; // important! reset
  with AColumn do begin
    Mask := HDI_FORMAT;
    Header_GetItem(GetDlgItem(ListView1.Handle, 0), Column.Index, AColumn);
    Mask := HDI_BITMAP or HDI_FORMAT;

    fmt := fmt or HDF_BITMAP;
    case Column.Alignment of
      taLeftJustify: fmt := fmt or HDF_BITMAP_ON_RIGHT;
      taCenter: fmt := fmt or HDF_BITMAP_ON_RIGHT;
      taRightJustify: fmt := fmt or HDF_BITMAP_ON_RIGHT;
    end;

    Header_SetItem(GetDlgItem(ListView1.Handle, 0), Column.Index, AColumn);
  end;

 { with AColumn do begin
    //pszText := PAnsiChar(Column.Caption);
    mask := mask or LVCF_IMAGE or LVCF_FMT;
    fmt := fmt or LVCFMT_IMAGE;
    fmt := fmt or LVCFMT_BITMAP_ON_RIGHT;
    case Column.Alignment of
      taLeftJustify: fmt := fmt or LVCFMT_LEFT;
      taCenter: fmt := fmt or LVCFMT_CENTER;
      taRightJustify: fmt := fmt or LVCFMT_RIGHT;
    end;
  end;

  if iImgIndex = iIndex_Ascend_SysIcons then
      AColumn.iImage := iIndex_Descend_SysIcons
  else if iImgIndex = iIndex_Descend_SysIcons then
      AColumn.iImage := iIndex_Ascend_SysIcons
  else
      AColumn.iImage := iIndex_Descend_SysIcons; }

  //SendMessage(ListView1.Handle, LVM_SETCOLUMN, Column.Index, Longint(@AColumn));
  iPreImgIndex := AColumn.hbm; // Preserve
end;

procedure TfrmMain.ListView1Compare(Sender: TObject; Item1,
  Item2: TListItem; Data: Integer; var Compare: Integer);
var
  ix, iV1, iV2, iImgIndex: Cardinal;
  Modified1, Modified2: TDateTime;
  sColumnCap: string;
begin
  iImgIndex := iPreImgIndex; // get last image index
  sColumnCap := ListView1.Columns[ColumnToSort].Caption;
  ix := ColumnToSort -1;

  if (Item1.SubItems.Strings[0] = '') and (Item1.SubItems.Strings[1] = '') and
  btnShowVirtual.Down then begin
      if (Item1.SubItems.Strings[0] = '') and (Item1.SubItems.Strings[1] = '') then
          Compare := 0
      else
          Compare := -1;

      exit;
  end;

  if AnsiPos('Name', sColumnCap) > 0 then begin// Name
      if iImgIndex = iIndex_Ascend_SysIcons then
          Compare := AnsiCompareStr(Item1.Caption, Item2.Caption)
      else if iImgIndex = iIndex_Descend_SysIcons then
          Compare := AnsiCompareStr(Item2.Caption,Item1.Caption)
      else
          Compare := AnsiCompareStr(Item1.Caption, Item2.Caption);

  end
  else if AnsiPos('Modified', sColumnCap) > 0 then begin // Modified
      if iImgIndex = iIndex_Ascend_SysIcons then begin
          Modified1 := StrToDateTime(Item1.SubItems[ix]);
          Modified2 := StrToDateTime(Item2.SubItems[ix]);
      end
      else if iImgIndex = iIndex_Descend_SysIcons then begin
          Modified1 := StrToDateTime(Item2.SubItems[ix]);
          Modified2 := StrToDateTime(Item1.SubItems[ix]);
      end
      else begin
          Modified1 := StrToDateTime(Item1.SubItems[ix]);
          Modified2 := StrToDateTime(Item2.SubItems[ix]);
      end;

      Compare := CompareDateTime(Modified1,Modified2);
  end
  else if AnsiPos('Size', sColumnCap) > 0 then begin //
      iV1 := round(strtoint(Item1.SubItems[ix]));
      iV2 := round(strtoint(Item2.SubItems[ix]));
      if iImgIndex = iIndex_Ascend_SysIcons then
          Compare := CompareValue(iV1, iV2)
      else if iImgIndex = iIndex_Descend_SysIcons then
          Compare := CompareValue(iV2, iV1)
      else
          Compare := CompareValue(iV1, iV2);

  end
  else if AnsiPos('Ratio', sColumnCap) > 0 then begin //
      iV1 := round(strtoint(AnsiReplaceStr(Item1.SubItems[ix], '% ', '')));
      iV2 := round(strtoint(AnsiReplaceStr(Item2.SubItems[ix], '% ', '')));
      if iImgIndex = iIndex_Ascend_SysIcons then
          Compare := CompareValue(iV1, iV2)
      else if iImgIndex = iIndex_Descend_SysIcons then
          Compare := CompareValue(iV2, iV1)
      else
          Compare := CompareValue(iV1, iV2);

  end
  else if AnsiPos('Packed', sColumnCap) > 0 then begin //
      iV1 := round(strtoint(Item1.SubItems[ix]));
      iV2 := round(strtoint(Item2.SubItems[ix]));
      if iImgIndex = iIndex_Ascend_SysIcons then
          Compare := CompareValue(iV1, iV2)
      else if iImgIndex = iIndex_Descend_SysIcons then
          Compare := CompareValue(iV2, iV1)
      else
          Compare := CompareValue(iV1, iV2)

  end
  else if AnsiPos('File Type', sColumnCap) > 0 then begin //
      if iImgIndex = iIndex_Ascend_SysIcons then
          Compare := AnsiCompareStr(Item2.SubItems[ix], Item1.SubItems[ix])
      else if iImgIndex = iIndex_Descend_SysIcons then
          Compare := AnsiCompareStr(Item1.SubItems[ix], Item2.SubItems[ix])
      else
          Compare := AnsiCompareStr(Item2.SubItems[ix], Item1.SubItems[ix]);

  end
  else if AnsiPos('CRC', sColumnCap) > 0 then begin //
      if iImgIndex = iIndex_Ascend_SysIcons then
          Compare := AnsiCompareStr(Item1.SubItems[ix], Item2.SubItems[ix])
      else if iImgIndex = iIndex_Descend_SysIcons then
          Compare := AnsiCompareStr(Item2.SubItems[ix], Item1.SubItems[ix])
      else
          Compare := AnsiCompareStr(Item1.SubItems[ix], Item2.SubItems[ix]);

  end
  else if AnsiPos('Attributes', sColumnCap) > 0 then begin //
      if iImgIndex = iIndex_Ascend_SysIcons then
          Compare := AnsiCompareStr(Item2.SubItems[ix], Item1.SubItems[ix])
      else if iImgIndex = iIndex_Descend_SysIcons then
          Compare := AnsiCompareStr(Item1.SubItems[ix], Item2.SubItems[ix])
      else
          Compare := AnsiCompareStr(Item2.SubItems[ix], Item1.SubItems[ix]);

  end
  else if AnsiPos('Path', sColumnCap) > 0 then begin //
      if iImgIndex = iIndex_Ascend_SysIcons then
          Compare := AnsiCompareStr(Item2.SubItems[ix], Item1.SubItems[ix])
      else if iImgIndex = iIndex_Descend_SysIcons then
          Compare := AnsiCompareStr(Item1.SubItems[ix], Item2.SubItems[ix])
      else
          Compare := AnsiCompareStr(Item2.SubItems[ix], Item1.SubItems[ix]);

  end;
 { else begin
      ix := ColumnToSort - 1;
      Compare := AnsiCompareStr(Item1.SubItems[ix],Item2.SubItems[ix]);
  end; }

end;

procedure TfrmMain.ZipFile_WillBe_DirectlyConvertToExe(sFilename: string);
begin
  {

  if not FileExists(sFilename) then begin
      beep;
      MessageDlg('Source file not found: ' + sFilename, mtError, [mbOK], 0);
      exit;
  end;

  with TdlgConvertToSFX.Create(self) do
  try
    edSource.Text := sFilename;
    edTarget.Text := ChangeFileExt(sFilename, '.exe');
    ShowModal;
  finally
    Free;
  end;

  }
end;

procedure TfrmMain.mnuSFXToZipClick(Sender: TObject);
var
  sFilename: string;
  iRet: integer;
begin
  if (ZipFName.Caption = '') then begin
      beep;
      MessageDlg('No opened file.', mtError, [mbOK], 0);
      exit;
  end;

  if AnsiLowercase( ExtractFileExt(ZipFName.Caption) ) <> '.exe' then begin
      beep;
      MessageDlg('Opened file is not executable file(.exe).', mtError, [mbOK], 0);
      exit;
  end;

  sFilename := ChangeFileExt(ZipFName.Caption, '.zip');
  If FileExists(sFilename) then begin
      if MessageDlg(sFilename + #13+#10 + 'File already exists. Do you want ' +
      'to overwrite?', mtConfirmation, [mbOK, mbCancel], 0) <> mrOK then
          exit;

  end;

  iRet := ZipMaster1.ConvertZIP;

  if iRet = 0 then // success
      MessageDlg(sFilename + #13+#10 + 'File has been successfully created.',
      mtInformation, [mbOK], 0)
  else begin
      beep;
      MessageDlg('Failing to convert. Error code: ' + inttostr(iRet), mtError,
      [mbOK], 0);
  end;


 { OpenDialog1.Filter := 'Open (.zip)|*.zip';
  OpenDialog1.InitialDir := GetCurrentDir;
  if OpenDialog1.Execute then begin
      frmZipSfxMain.ZipSFX1.SourceFile := OpenDialog1.FileName;
      sFilename := ChangeFileExt(frmZipsfxMain.ZipSFX1.SourceFile, '.zip');

      SaveDialog.Filter := 'Save to (.exe)|*.exe';
      SaveDialog.FileName := ExtractFilename(sFilename);
      if SaveDialog.Execute then begin
          If FileExists(SaveDialog.FileName) then begin
              if MessageDlg('File already exists. Do you want to overwrite?',
              mtConfirmation, [mbYes, mbCancel], 0) <> mrYes then
                  exit;

          end;

          frmZipSfxMain.ZipSFX1.TargetFile := SaveDialog.FileName;
          frmZipsfxMain.ZipSFX1.ConvertToZip;
          MessageDlg('ZIP archive '+ frmZipSfxMain.ZipSFX1.TargetFile+#13#10+
              'has been created.', mtInformation, [mbOK], 0);
      end;
  end; }
end;

procedure TfrmMain.mnuSubCopyArchiveClick(Sender: TObject);
var
  sOutPath, sFilename: string;
  iRet: integer;
begin
  if ZipMaster1.ZipFileName = '' then begin
      beep;
      MessageDlg('There has not opened file.', mtError, [mbOK], 0);
      exit;
  end;

  if SelectDirectory('Copy Archive to...', '', sOutPath) then begin
      sOutPath := IncludeTrailingPathDelimiter(sOutPath);
      sFilename := ExtractFilename(ZipMaster1.ZipFileName);

      if FileExists(sOutPath + sFilename) then begin
          if MessageDlg('File already exists. Do you want to overwrite?',
          mtConfirmation, [mbYes, mbNo], 0) <> mrYes then 
              exit;

      end;

      iRet := ZipMaster1.CopyFile(ZipMaster1.ZipFileName, sOutPath + sFilename);
      if iRet <> 0 then begin
          beep;
          if iRet = -4 then // Error setting date/time of destination
              MessageDlg('Archive has been copied successfully but failing to set ' +
              'date-time to it.', mtInformation, [mbOK], 0)
          else
              MessageDlg('Failing to copy file. Error code: ' + inttostr(iRet),
              mtError, [mbOK], 0);
              
      end;
  end;
end;

procedure TfrmMain.mnuSubDeleteClick(Sender: TObject);
begin
  mnuDeleteClick(nil);
end;

procedure TfrmMain.mnuSubOpenClick(Sender: TObject);
begin
  mnuOpenClick(nil);
end;

procedure TfrmMain.mnuSubOpenWithClick(Sender: TObject);
begin
  mnuOpenWithClick(nil);
end;

procedure TfrmMain.AddNode_ToTreeView;
var
  i, j, k, iZ, iRetPos, iCnt, iFrom: integer;
  sTemp, sDirIs, sPrePath: string;
  sArrayDir: array of string;
  UpTreeNode, NodeIs: TTreeNode;
  bFoundSameNode: boolean;
begin
  TreeView1.Items.Clear;
  UpTreeNode := nil;
  if TreeView1.TopItem = nil then begin // First time new node added
      UpTreeNode := TreeView1.Items.Add(nil, ExtractFilename(ZipFName.Caption));
      for i := Low(sArrayDir) to High(sArrayDir) do begin
        UpTreeNode := TreeView1.Items.AddChild(UpTreeNode, sArrayDir[i]);
      end;
      TreeView1.Items[0].Expand(False);
  end;

  sPrePath := '';
  for iZ := 0 to ZipMaster1.ZipContents.Count -1 do begin
  with ZipDirEntry(ZipMaster1.ZipContents[iZ]^) do begin

    sTemp := ExtractFilePath(FileName);
    if (sTemp <> '') and (sTemp <> sPrePath) then begin
       sPrePath := sTemp; 
       k := 0;
       repeat
         iRetPos := AnsiPos('\', sTemp);
         if iRetPos > 0 then begin
             SetLength(sArrayDir, k+1);
             sDirIs := copy(sTemp, 1, iRetPos-1); // dir without "\"
             sArrayDir[k] := sDirIs;
             sTemp := copy(sTemp, iRetPos+1, length(sTemp)-iRetPos);
             Inc(k);
         end;
       until iRetPos = 0;

       NodeIs := nil;
       UpTreeNode := nil;
       iFrom := 1; // 0 is filename on top item
       for i := Low(sArrayDir) to High(sArrayDir) do begin
         bFoundSameNode := False;
         for j := iFrom to TreeView1.Items.Count -1 do begin
           if TreeView1.Items.Item[j].Level = i +1 then begin // starting level is 1, so i +1
               if AnsiLowercase(TreeView1.Items.Item[j].Text) = AnsiLowercase(sArrayDir[i]) then begin
                   bFoundSameNode := True;
                   iFrom := j +1; // Start from next node
                   NodeIs := TreeView1.Items.Item[j]; // Preserve found node
                   break;
               end;
           end;
         end;

         if bFoundSameNode = False then begin
             if NodeIs <> nil then begin
                 UpTreeNode := NodeIs;
                 for iCnt := i to High(sArrayDir) do begin
                   UpTreeNode := TreeView1.Items.AddChild(UpTreeNode, sArrayDir[iCnt]);
                 end;
             end
             else begin
                 UpTreeNode := TreeView1.TopItem;
                 for iCnt := Low(sArrayDir) to High(sArrayDir) do begin
                   UpTreeNode := TreeView1.Items.AddChild(UpTreeNode, sArrayDir[iCnt]);
                 end;
             end;
         end;
       end;

    end;
  end;
  end;
end;

procedure TfrmMain.Refresh_ItemsOn_VirtualMainList(sPath: string;
 sArrayDirIs: array of string);
var
  i, j, k, iGaugeStart, iIndex: integer;
  ListItem: TListItem;
  sFilename, sFileExt, sNumFiles, sListPath, sListFilename, sThePath, sFile_Type: string;
  NewCol: TCollectionItem;
begin
  {$IFDEF IS_SHAREWARE}
  if ZipMaster1.ZipFileName = '' then exit;

  if bCol_AddedOrRemoved_OnMainLst then begin
      for k := ListView1.Columns.Count -2 downto 5 do begin // remove extra columns
        ListView1.Columns.Delete(k);
      end;
      if bCheckType then begin
          NewCol := ListView1.Columns.Insert(ListView1.Columns.Count);
          ListView1.Column[NewCol.Index -1].Caption := 'File Type';
          ListView1.Column[NewCol.Index -1].Alignment := taRightJustify;
          ListView1.Column[NewCol.Index -1].Width := self.iColWidthFileType;
      end;
      if bCheckCRC then begin
          NewCol := ListView1.Columns.Insert(ListView1.Columns.Count);
          ListView1.Column[NewCol.Index -1].Caption := 'CRC';
          ListView1.Column[NewCol.Index -1].Alignment := taRightJustify;
          ListView1.Column[NewCol.Index -1].Width := self.iColWidthCRC;
      end;
      if bCheckAttributes then begin
          NewCol := ListView1.Columns.Insert(ListView1.Columns.Count);
          ListView1.Column[NewCol.Index -1].Caption := 'Attributes';
          ListView1.Column[NewCol.Index -1].Alignment := taRightJustify;
          ListView1.Column[NewCol.Index -1].Width := self.iColWidthAttributes;
      end;
      SetBool_CntOf_ColumnChanged_OnMainLst(False);
      ListView1.Column[ListView1.Columns.Count -1].Caption := 'Path';
  end;

  // important to set autosize of main list columns is false, faster refresh
  if Main.bCheckMainLstAutoSize then begin
      for j := ListView1.Columns.Count -1 downto 0 do begin
        ListView1.Columns[j].AutoSize := False;
        if ListView1.Columns[j].Caption = 'Name' then
            ListView1.Columns[j].Width := self.iColWidthName
        else if ListView1.Columns[j].Caption = 'Modified' then
            ListView1.Columns[j].Width := self.iColWidthModified
        else if ListView1.Columns[j].Caption = 'Size' then
            ListView1.Columns[j].Width := self.iColWidthSize
        else if ListView1.Columns[j].Caption = 'Ratio' then
            ListView1.Columns[j].Width := self.iColWidthRatio
        else if ListView1.Columns[j].Caption = 'Packed' then
            ListView1.Columns[j].Width := self.iColWidthPacked
        else if ListView1.Columns[j].Caption = 'File Type' then
            ListView1.Columns[j].Width := self.iColWidthFileType
        else if ListView1.Columns[j].Caption = 'CRC' then
            ListView1.Columns[j].Width := self.iColWidthCRC
        else if ListView1.Columns[j].Caption = 'Attributes' then
            ListView1.Columns[j].Width := self.iColWidthAttributes
        else if ListView1.Columns[j].Caption = 'Path' then
            ListView1.Columns[j].Width := self.iColWidthPath;

    end;
  end;

  //Gauge1.Visible := True;
  //Gauge1.MinValue := 0;
  //Gauge1.Progress := 0;
  //Gauge1.MaxValue := ZipMaster1.Count + Length(sArrayDirIs);
  //iGaugeStart := 0;
  TreeView1.Cursor := crHourGlass;
  ListView1.Cursor := crHourGlass;

  ListView1.Items.BeginUpdate;
  ListView1.Items.Clear; // <= faster than ListView1.Clear
  ListView1.Items.EndUpdate;

  if Length(sArrayDirIs) > 0 then begin // folders found !
      for i := low(sArrayDirIs) to high(sArrayDirIs) do begin
        if sArrayDirIs[i] <> '' then begin
            ListItem := ListView1.Items.Add;
            ListItem.Caption := sArrayDirIs[i];
            if ListItem.Caption = '..' then  // up one level folder
                ListItem.ImageIndex := 5
            else // folder
                ListItem.ImageIndex := 6;

            //ListItem.ImageIndex := Get_SysIconIndex_Of_Given_FileExt(sFileExt);
            ListItem.SubItems.Add(''); // Modified
            ListItem.SubItems.Add(''); // Size
            ListItem.SubItems.Add(''); // Ratio
            ListItem.SubItems.Add(''); // Packed

            if bCheckType then  // Type
                ListITem.SubItems.Add('Folder');

            if bCheckCRC then  // CRC
                ListITem.SubItems.Add('');

            if bCheckAttributes then  // Attributes
                ListITem.SubItems.Add('');

            ListItem.SubItems.Add(''); // Path
        end;
        //Gauge1.Progress := i;
        //iGaugeStart := Gauge1.Progress;
      end;
  end;

  sPath := AnsiLowercase(sPath); // AnsiLowercase first
  for i := 0 to ZipMaster1.ZipContents.Count -1 do begin
    with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
      sListPath := ExtractFilePath(FileName);
      sListFilename := ExtractFileName(FileName);
      if (AnsiLowercase(sListPath) = sPath) and (sListFilename <> '') then begin
          ListItem := ListView1.Items.Add;
          if Encrypted then
              ListItem.Caption := sListFilename + '*' // Name
          else
              ListItem.Caption := sListFilename; // Name
              
          ListItem.ImageIndex := -1;
          sFileExt := AnsiLowercase( ExtractFileExt(FileName) );
          sFilename := AnsiLowercase(FileName);

          if sFilename = 'setup.exe' then
              ListItem.ImageIndex := 3
          else if (sFilename = 'readme.txt') or (sFilename = 'readme1st.txt') then
              ListItem.ImageIndex := 4
          else if sPreExt = sFileExt then
              ListItem.ImageIndex := iPreImgI
          else  begin
              iIndex := Get_SysIconIndex_Of_Given_FileExt(sFileExt, sFile_Type);
              ListItem.ImageIndex := iIndex;
            {  if sFileExt = '' then
                  iIndex := cbSystemIcons.Items.IndexOf('.')
              else
                  iIndex := cbSystemIcons.Items.IndexOf(sFileExt); // Using combo box to check existed icon
                  
              if iIndex = -1 then begin // Icon is not exist in ilSystem
                  iIndex := Get_SysIconIndex_Of_Given_FileExt(sFileExt, sFile_Type);
                  ilMainList.GetIcon(iIndex, AIcon);
                  ilSystem.InsertIcon(ilSystem.Count, AIcon); // insert new
                  if sFileExt = '' then
                      cbSystemIcons.Items.Append('.') // combo box does not allow append null string
                  else
                      cbSystemIcons.Items.Append(sFileExt); // update combo box
                      
                  ListItem.ImageIndex := ilSystem.Count -1;
              end
              else // existed icon
                  ListItem.ImageIndex := iIndex +8; // ilSystem has preset icons

              }
              sPreExt := sFileExt;
              iPreImgI := ListItem.ImageIndex;
          end;

          ListItem.SubItems.Add(FormatDateTime('ddddd  t',
          FileDateToDateTime(DateTime))); // Modified
          ListItem.SubItems.Add(IntToStr(UncompressedSize)); // Size
          if UncompressedSize <> 0 Then
              ListItem.SubItems.Add(IntToStr(Round((1 - (CompressedSize /
              UncompressedSize)) * 100)) + '% ') //Ratio
          else
              ListItem.SubItems.Add('0% ');

          sThePath := ExtractFilePath(FileName);
          ListItem.SubItems.Add(IntToStr(CompressedSize)); // Packed

          if bCheckType then begin // Type
              if sFile_Type <> '' then
                  ListItem.SubItems.Add(sFile_Type)
              else begin
                  sFile_Type := Uppercase( ExtractFileExt(Filename) );
                  sFile_Type := AnsiReplaceStr(sFile_Type, '.', ' ');
                  ListItem.SubItems.Add(sFile_Type);
              end;
          end;

          if bCheckCRC then  // CRC
              ListITem.SubItems.Add(inttohex(CRC32, 2));

          if bCheckAttributes then  // Attributes
              ListITem.SubItems.Add(GetAttributes(ExtFileAttrib));

          ListITem.SubItems.Add(sThePath); //Path

      end;

      if (ListView1.Items.Count mod 400) = 0 then begin
          ListView1.Repaint;
          Self.Repaint;
      end;
      //Gauge1.Progress := iGaugeStart + i;
    end;
  end;

  //Gauge1.Visible := False;
  //Gauge1.Progress := 0;
  TreeView1.Cursor := crDefault;
  ListView1.Cursor := crDefault;

  if ZipMaster1.Count > 1 then
      sNumFiles := 'Objects = '
  else
      sNumFiles := 'Object = ';

  StatusBar1.Panels[0].Text := ZipMaster1.ZipFileName;

  if btnAdvancedView.Down then
      StatusBar1.Panels[1].Text := sNumFiles + inttostr(ZipMaster1.Count)
  else
      StatusBar1.Panels[1].Text := '';
      
  StatusBar1.Panels[2].Text := 'Size = ' + inttostr(ZipMaster1.ZipFileSize
  div 1024) + ' KB' ;
  ClearColumnsImage_OnlvMain;
  iIndex := Get_ColIndex_ToSort_FromMenuSortBy;
  if iIndex <> -1 then
      ListView1.OnColumnClick(ListView1, ListView1.Column[iIndex]);

  if Main.bCheckMainLstAutoSize then
      ResizeWidthOfMainListColumns;
      
  SetStatusOfPopupUnzip(nil);
  {$ENDIF}
end;

procedure TfrmMain.TreeView1Change(Sender: TObject; Node: TTreeNode);
begin
  TreeView1_Change;
end;

procedure TfrmMain.ChangeFolder_OnVirtualMainList_Click(sDir: string);
var
  iIndex: integer;
  sPath: string;
  FirstChild, NextChild: TTreeNode;
  bFoundFolder: boolean;
begin
  sDir := AnsiLowercase(sDir);
  sPath := '';
  bFoundFolder := False;
  if TreeView1.SelectionCount > 0 then begin
      if sDir = AnsiLowercase('..') then begin // is up one level
          if TreeView1.Selected.Parent <> nil then
              iIndex := TreeView1.Selected.Parent.AbsoluteIndex
          else begin
              beep;
              MessageDlg('Unexpected error. Failing to get folder name of upper level.', mtError, [mbOK], 0);
              exit;
          end;
      end
      else begin // folder
          iIndex := TreeView1.Selected.AbsoluteIndex;

          FirstChild := TreeView1.Items.Item[iIndex].getFirstChild;
          if FirstChild <> nil then begin // get all folder names under
              if sDir = AnsiLowercase(FirstChild.Text) then begin // Found folder match with clicking object on Virtual List
                  bFoundFolder := True;
                  iIndex := FirstChild.AbsoluteIndex; // Change index number to selected folder name on Virtual List
              end
              else begin // try to find others
                  NextChild := FirstChild;
                  repeat
                    NextChild := TreeView1.Items.Item[iIndex].GetNextChild(NextChild);
                    if NextChild <> nil then begin
                        if sDir = AnsiLowercase(NextChild.Text) then begin
                            bFoundFolder := True;
                            iIndex := NextChild.AbsoluteIndex;
                        end;
                    end;
                  until NextChild = nil;
              end;
          end
          else begin
              beep;
              MessageDlg('Unexpected error. Failing to change folder.', mtError, [mbOK], 0);
              exit;
          end;

          if bFoundFolder = False then begin
              beep;
              MessageDlg('Error. Failing to change folder.', mtError, [mbOK], 0);
              exit;
          end;
      end;

      TreeView1.Items.Item[iIndex].Selected := True; // update selected tree
  end;
end;


procedure TfrmMain.btnShowVirtualClick(Sender: TObject);
begin
  {$IFDEF IS_SHAREWARE}
  if btnShowVirtual.Down then begin
      btnShiftViewMode.Visible := True;
      pnShell.Visible := False;
      if ZipFName.Caption <> '' then begin
          AddNode_ToTreeView;
          TreeView1.Selected := TreeView1.TopItem;
          TreeView1_Change;
      end
      else
          Update_lblCurDir('');
          
  end
  else begin
      btnShiftViewMode.Visible := False;
      pnShell.Visible := True;
      CoolBar1.Repaint;
      if ZipFName.Caption <> '' then
          Refresh_ItemsOnMainList;

      Update_lblCurDir(ShellDir1.Path);    
  end;
  {$ENDIF}
end;

procedure TfrmMain.DeleteFiles_InArchive;
var
  sFileName, sPath, sFileNoIs: string;
  Node: PVirtualNode;
  Data: PEntry;
begin
  sPath := '';
  sErrorMessage := '';
  ZipMaster1.FSpecArgs.Clear;

  Screen.Cursor := crHourGlass;
  try
    Node := VT.GetFirstSelected;
    while Assigned(Node) do begin
      Data := VT.GetNodeData(Node);
      sFileName := Data.Value[0]; // Name
      sFileName := AnsiReplaceStr(sFileName, '*', ''); // remove password char
      sPath := Data.Value[8]; // Path
      sPath := AnsiReplaceStr(sPath, '\', '/');
      if (sPath <> '') and ( copy(sPath, length(sPath), 1)
      <> '/' ) then
          sPath := sPath + '/';

      ZipMaster1.FSpecArgs.Add(sPath + sFileName);

      Node := VT.GetNextSelected(Node);
    end;

    if ZipMaster1.FSpecArgs.Count < 1 then
        sErrorMessage := sErrorMessage + 'No selected file found.' + #13+#10
    else begin // could get filenames to delete
        Screen.Cursor := crHourGlass;
        try
          try
            ZipMaster1.Delete;
          except
            sErrorMessage := sErrorMessage + 'Fatal error trying to ' +
            'delete file(s). Successful deletion = ' + inttostr(
            ZipMaster1.SuccessCnt) + ' file(s) only.' + #13+#10;
          end;
        finally
          Screen.Cursor := crDefault;
        end;
    end;

  finally
    Screen.Cursor := crDefault;

    if sErrorMessage <> '' then begin// Catch error messages
        PopUp_ErrorMessage('Error occurred when deleting file(s).');
        Update_FDataList; // Must do this first!
        Refresh_ItemsOnMainList;
    end
    else begin // Found no error message, directly delete nodes on VT
        Update_FDataList; // Must do this first!
        VT.DeleteSelectedNodes;
    end;

    if FDataList.Count > 1 then
        sFileNoIs := 'Objects = '
    else
        sFileNoIs := 'Object = ';

    if btnAdvancedView.Down then
        StatusBar1.Panels[1].Text := sFileNoIs + inttostr(FDataList.Count)
    else
        StatusBar1.Panels[1].Text := '';
        
    StatusBar1.Panels[2].Text := 'Size = ' + inttostr(ZipMaster1.ZipFileSize
    div 1024) + ' KB' ;
  end;
end;

procedure TfrmMain.DeleteFiles_InArchive_OnVirtualMainList;
var
  sCaptionIs, sZipFolder, sPath: string;
  i, k, iIndex, iCntCol: integer;
  sArrayDirs: array of string;
begin
try
  TreeView1.Enabled := False; // don't let user change status
  sZipFolder := '';
  sErrorMessage := '';
  ZipMaster1.FSpecArgs.Clear;

  iCntCol := ListView1.Columns.Count;
  i := 0;
  k := 0;
  Screen.Cursor := crHourGlass;
  try
    repeat
      sZipFolder := '';
      //Is caption selected?;
      if ListView1.Items[i].Selected then begin
          sCaptionIs := ListView1.Items[i].Caption; // File name
          sCaptionIs := AnsiReplaceStr(sCaptionIs, '*', ''); // remove password char

          if sCaptionIs = '..' then // skip up one level folder

          else if (ListView1.Items[i].SubItems.Strings[0] = '') and
          (ListView1.Items[i].SubItems.Strings[1] = '') and
          btnShowVirtual.Down then begin // folder found
                sPath := GetSelectedPath_FromVirtualTreeView;
                sPath := sPath + sCaptionIs + '\';
                AddAllFiles_UnderDir_ToFSpecArgs_BeforeDel(sPath);

                SetLength(sArrayDirs, k +1);
                sArrayDirs[k] := sCaptionIs; // preserve folder name for deleting node
                inc(k);
          end
          else begin // a file item
               sZipFolder := ListView1.Items[i].SubItems.Strings[iCntCol -2]; // get path name
               sZipFolder := AnsiReplaceStr(sZipFolder, '\', '/');
               if (sZipFolder <> '') and ( copy(sZipFolder, length(sZipFolder), 1)
               <> '/' ) then
                   sZipFolder := sZipFolder + '/';

               // store them up
               ZipMaster1.FSpecArgs.Add(sZipFolder + sCaptionIs)
          end;
      end;
      i := i + 1;

    until i >= ListView1.Items.Count;
  finally
    Screen.Cursor := crDefault;
  end;

  if ZipMaster1.FSpecArgs.Count < 1 then
      sErrorMessage := sErrorMessage + 'No selected file found.' + #13+#10
  else begin // got filenames and try to delete
      try
        ZipMaster1.Delete;
      except
        sErrorMessage := sErrorMessage + 'Fatal error trying to ' +
        'delete file(s). Successful deletion = ' + inttostr(
        ZipMaster1.SuccessCnt) + ' file(s) only.' + #13+#10;
      end;
  end;

  if Length(sArrayDirs) > 0 then
      Remove_DeletedNode_InVirtualTree(sArrayDirs);

  if sErrorMessage <> '' then begin// Catch error messages
      iIndex := TreeView1.Selected.AbsoluteIndex;
      AddNode_ToTreeView; // Reload items to Tree View from zip stream, it also can find out what items on virtual main list are failing to be deleted
      if iIndex <= TreeView1.Items.Count -1 then begin // go to previous selected folder
          TreeView1.Items.Item[iIndex].Selected := True;
          TreeView1.Items.Item[iIndex].Expand(False);
      end
      else begin // prevent error when previous selected folder not found
          TreeView1.Selected := TreeView1.TopItem;
          TreeView1.Selected.Expand(False);
      end;
      TreeView1_Change; // refresh TreeView and Virtual ListView
      PopUp_ErrorMessage('Error occurred when deleting file(s).');
  end
  else
      TreeView1_Change; // refresh TreeView and Virtual ListView
      
finally
   TreeView1.Enabled := True; // let user use it again
end;
end;

procedure TfrmMain.AddAllFiles_UnderDir_ToFSpecArgs_BeforeDel(sDir: string);
var
  i, iLen: integer;
  sPathFilename: string;
begin
  sDir := AnsiLowercase(sDir); // Format of sDir: aa\bb\ccc\
  iLen := length(sDir);
  for i := 0 to ZipMaster1.ZipContents.Count -1 do begin
    with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
      if sDir = AnsiLowercase(copy(Filename, 1, iLen)) then begin // is under sDir
          sPathFilename := Filename;
          sPathFilename := AnsiReplaceStr(sPathFilename, '\', '/');
          // store them up
          ZipMaster1.FSpecArgs.Add(sPathFilename)
      end;
    end;
  end;
end;

function TfrmMain.GetSelectedPath_FromVirtualTreeView: string;
var // e.g: xxx\yyy\zzzz\
  iIndex: integer;
  sPath: string;
  ParentNode: TTreeNode;
begin
  Result := '';
  sPath := '';
  iIndex := TreeView1.Selected.AbsoluteIndex;
  if not TreeView1.Items.Item[iIndex].IsFirstNode then // not zip filename
      sPath := TreeView1.Items.Item[iIndex].Text + '\';

      ParentNode := TreeView1.Items.Item[iIndex].Parent;
      // Get path name without first node(normally, first node is zip filename)
      if (ParentNode <> nil) and (ParentNode <> TreeView1.Items.GetFirstNode) then begin
          repeat
            sPath := ParentNode.Text + '\' + sPath;
            ParentNode := ParentNode.Parent;
            if ParentNode = TreeView1.Items.GetFirstNode then break;
          until ParentNode = nil;
      end;

      Result := sPath;
end;

procedure TfrmMain.Remove_DeletedNode_InVirtualTree(sArrayDirIs: array of string);
var
  i, iIndex: integer;
  FirstChild, NextChild: TTreeNode;
begin
  if Length(sArrayDirIs) > 0 then begin
      iIndex := TreeView1.Selected.AbsoluteIndex;

      for i := low(sArrayDirIs) to high(sArrayDirIs) do begin
        FirstChild := TreeView1.Items.Item[iIndex].getFirstChild;
        if FirstChild = nil then exit; // no sub-dir found
        NextChild := FirstChild;
        repeat
          if AnsiLowercase(sArrayDirIs[i]) = AnsiLowercase(NextChild.Text) then begin
              NextChild.Delete;
              break;
          end;
          NextChild := TreeView1.Items.Item[iIndex].GetNextChild(NextChild);
        until NextChild = nil;
      end;
  end;
end;

procedure TfrmMain.TreeView1_Change;
var
  k, iIndex: integer;
  sPath: string;
  sArrayDir: array of string;
  ParentNode, FirstChild, NextChild: TTreeNode;
begin
  sPath := '';
  if TreeView1.SelectionCount > 0 then begin
      TreeView1.Selected.SelectedIndex := 1;
          
      TreeView1.Selected.Expand(False);
      TreeView1.Repaint;
      iIndex := TreeView1.Selected.AbsoluteIndex;

      k := 0;
      if not TreeView1.Items.Item[iIndex].IsFirstNode then begin // not zip filename
          SetLength(sArrayDir, k+1);
          sArrayDir[k] := '..'; // add up one level folder
          inc(k);
      end;

      FirstChild := TreeView1.Items.Item[iIndex].getFirstChild;
      if FirstChild <> nil then begin // get all folder names under
          SetLength(sArrayDir, k+1);
          sArrayDir[k] := FirstChild.Text;
          inc(k);
          NextChild := FirstChild;
          repeat
            NextChild := TreeView1.Items.Item[iIndex].GetNextChild(NextChild);
            if NextChild <> nil then begin
                SetLength(sArrayDir, k+1);
                sArrayDir[k] := NextChild.Text; // add folder name
            end;
            inc(k);
          until NextChild = nil;
      end;

      if not TreeView1.Items.Item[iIndex].IsFirstNode then
          sPath := TreeView1.Items.Item[iIndex].Text + '\' + sPath;

      ParentNode := TreeView1.Items.Item[iIndex].Parent;
      // Get path name without first node(normally, first node is zip filename)
      if (ParentNode <> nil) and (ParentNode <> TreeView1.Items.GetFirstNode) then begin
          repeat
            sPath := ParentNode.Text + '\' + sPath;
            ParentNode := ParentNode.Parent;
            if ParentNode = TreeView1.Items.GetFirstNode then break;
          until ParentNode = nil;
      end;
      Refresh_ItemsOn_VirtualMainList(sPath, sArrayDir);
      Update_lblCurDir('\' + ExcludeTrailingPathDelimiter(sPath)); // update caption on header
  end;
end;

procedure TfrmMain.Extract_SelectedDir_OnVirtualMainList(sDir: string;
 var sPathFNames: TStrings);
var
  i, iLen: integer;
begin
  sDir := AnsiLowercase(sDir); // Format of sDir: aa\bb\ccc\
  iLen := length(sDir);
  for i := 0 to ZipMaster1.ZipContents.Count -1 do begin
    with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
      if sDir = AnsiLowercase(copy(Filename, 1, iLen)) then // is under sDir
          sPathFNames.Append(Filename);

    end;
  end;
end;

procedure TfrmMain.btnShiftViewModeClick(Sender: TObject);
var
  iIndex: integer;
  sPath: string;
  ParentNode: TTreeNode;
  Node: PVirtualNode;
  NS: TNamespace;
begin
  pnShell.Visible := not pnShell.Visible;
  if pnShell.Visible then begin
      Node := VETree.GetFirstSelected;
      if VETree.ValidateNamespace(Node, NS) then
          Update_lblCurDir(NS.NameForParsing);

      //Update_lblCurDir(ShellDir1.Path)
  end
  else begin
      if TreeView1.SelectionCount = 0 then begin
          Update_lblCurDir('');
          exit;
      end;

      iIndex := TreeView1.Selected.AbsoluteIndex;
      if TreeView1.Items.Item[iIndex].IsFirstNode then
          sPath := '\'
      else
          sPath :=  TreeView1.Items.Item[iIndex].Text;

      ParentNode := TreeView1.Items.Item[iIndex].Parent;
      // Get path name without first node(normally, first node is zip filename)
      if (ParentNode <> nil) and (ParentNode <> TreeView1.Items.GetFirstNode) then begin
          repeat
            sPath := ParentNode.Text + '\' + sPath;
            ParentNode := ParentNode.Parent;
            if ParentNode = TreeView1.Items.GetFirstNode then begin
                sPath := '\' + sPath;
                break;
            end;
          until ParentNode = nil;
      end;
      Update_lblCurDir(sPath);
  end;
end;

procedure TfrmMain.btnCommentClick(Sender: TObject);
begin
  mnuViewCommentClick(nil);
end;

procedure TfrmMain.Set_ZipTempIs(sStringIs: string);
begin
  sZipTempFolder := sStringIs;
end;


procedure TfrmMain.btnFavoriteDirClick(Sender: TObject);
var
  sSelectedFolder: string;
  i, iCnt: integer;
  Node: PVirtualNode;
  NS: TNamespace;
begin
  //if ShellDir1.Path = '' then exit;
  //sSelectedFolder := ShellDir1.Path;
  Node := VETree.GetFirstSelected;
  if VETree.ValidateNamespace(Node, NS) then begin
      if NS.Folder then
          sSelectedFolder := NS.NameForParsing
      else
          exit;
          
  end
  else
      exit;

  iCnt := pmFavoriteDir.Items.Count;
  for i := 0 to iCnt -1 do begin
    if AnsiLowercase(sSelectedFolder) = AnsiLowercase(pmFavoriteDir.Items[i].Caption) then
        exit;

  end;

  if iCnt < iCntFavorite then begin // not full, create new item 
      mnuFavoriteDir[iCnt] := TMenuItem.Create(Self);
      mnuFavoriteDir[iCnt].Caption := '';
      pmFavoriteDir.Items.Insert(iCnt, mnuFavoriteDir[iCnt]);
      mnuFavoriteDir[iCnt].OnClick := FavoriteDirOnClick;
      mnuFavoriteDir[iCnt].AutoCheck := True;
  end;

  for i := (pmFavoriteDir.Items.Count -1) downto 0 do begin // move down
    if i = 0 then break;
    pmFavoriteDir.Items[i].Caption := pmFavoriteDir.Items[i-1].Caption;
  end;
  
  pmFavoriteDir.Items[0].Caption := sSelectedFolder; // new to top item
  DelRegistryKey(HKEY_CURRENT_USER, XcSubKeyIs + cFavouriteFolders);
  // Save favourite folders to registry
  for i := 0 to pmFavoriteDir.Items.Count -1 do begin
    if pmFavoriteDir.Items[i].Caption = '' then break;
        WriteIniToReg(XcSubKeyIs + cFavouriteFolders, inttostr(i),
        pmFavoriteDir.Items[i].Caption)
  end;
end;

procedure TfrmMain.DelRegistryKey(const iRootKey: Cardinal; sSubKey: string);
begin
  with TRegistry.Create do
  try
    DeleteRegKey(sSubKey, iRootKey);
  finally
    Free;
  end;
end;


function TfrmMain.ShowProperties(hWndOwner: HWND; const FileName: string): Boolean;
var
  Info: TShellExecuteInfo;
begin 
  { Fill in the SHELLEXECUTEINFO structure } 
  with Info do 
  begin 
    cbSize := SizeOf(Info); 
    fMask := SEE_MASK_NOCLOSEPROCESS or 
             SEE_MASK_INVOKEIDLIST or 
             SEE_MASK_FLAG_NO_UI; 
    wnd  := hWndOwner; 
    lpVerb := 'properties'; 
    lpFile := pChar(FileName); 
    lpParameters := nil; 
    lpDirectory := nil; 
    nShow := 0; 
    hInstApp := 0; 
    lpIDList := nil; 
  end; 

  { Call Windows to display the properties dialog. } 
  Result := ShellExecuteEx(@Info);
end; 

procedure TfrmMain.mnuPropertiesClick(Sender: TObject);
var
  sPathFileName: string;
  bIsVirtualFolder, bIsEncrypted: boolean;
begin
  if VT.SelectedCount = 0 then begin
      beep;
      MessageDlg('Nothing selected.', mtInformation, [mbOK], 0);
      exit;
  end;
  
  sPathFileName := GetPathFileName_FromSelected_Item(bIsVirtualFolder, bIsEncrypted);
  if (sPathFileName <> '') and (bIsVirtualFolder = False) then
      ExtactFile_Before_ShowProperties(sPathFileName);

end;

procedure TfrmMain.ExtactFile_Before_ShowProperties(sFileName: string);
var
  sOpenFile: string;
  bExtractOne: boolean;
  sStrings: TStrings;
begin
  // Is there no any extraced file? (= no temporary folder was created)
  if bExtractedFiles = False then begin
      edShadow.Text := Create_TempFolder_Under_WindowsTemp;
      if edShadow.Text = '' then
          exit
      else
          bExtractedFiles := True;
  end;

  bExtractOne := True;
  try
    sStrings := TStringList.Create;
    sStrings.Append(sFileName);

    RadioExtractWithDir.ItemIndex := 1;
    RadioOverwrite.ItemIndex := 1;
    StartUnzip(False, ZipFName.Caption, edShadow.Text, sStrings, bExtractOne);
  finally
    sStrings.Free;
  end;

  sOpenFile := edShadow.Text + sFileName;
  ShowProperties(Self.Handle, sOpenFile);
end;

procedure TfrmMain.mnuSubPropertiesClick(Sender: TObject);
begin
  mnuPropertiesClick(nil);
end;

procedure TfrmMain.Who_Installed_Software_WriteTo_Registry;
var
  sCurUserName: string;
begin
  sCurUserName := UserID;
  if sCurUserName = '' then
      sCurUserName := XcDefault;
      
  if ReadRegLocalMachine(XcSubKeyIs, XcInstalledBy) = '' then
      WriteRegLocalMachine(XcSubKeyIs, XcInstalledBy, sCurUserName);
      
end;

function TfrmMain.UserID: String;
var 
  UID : PChar; 
  USize : DWord;
  bRetSuccess: boolean;
begin 
  GetMem(UID, 128);
  USize:=128;
  bRetSuccess := GetUserName(UID, USize);
  if bRetSuccess then
      Result:= String(UID)
  else
      Result := '';
      
  FreeMem(UID, 128);
end;

procedure TfrmMain.SetBool_CheckType(bChecked: boolean);
begin
  bCheckType := bChecked;
end;

procedure TfrmMain.SetBool_CheckCRC(bChecked: boolean);
begin
  bCheckCRC := bChecked;
end;

procedure TfrmMain.SetBool_CheckAttributes(bChecked: boolean);
begin
  bCheckAttributes := bChecked;
end;

function TfrmMain.GetAttributes(iFileAttri: Cardinal): string;
begin
  Result := '';
  if Bool(iFileAttri and FILE_ATTRIBUTE_READONLY) then
      Result := Result + 'R';

  if Bool(iFileAttri and FILE_ATTRIBUTE_HIDDEN) then
      Result := Result + 'H';

  if Bool(iFileAttri and FILE_ATTRIBUTE_SYSTEM) then
      Result := Result + 'S';

  if Bool(iFileAttri and FILE_ATTRIBUTE_ARCHIVE) then
      Result := Result + 'A';
      
end;

procedure TfrmMain.mnuSubMoveArchiveClick(Sender: TObject);
var
  sPathFilename, sFilename, sOutPath, sThisFile: string;
begin
  if ZipMaster1.ZipFileName = '' then begin
      beep;
      MessageDlg('There has not opened file.', mtError, [mbOK], 0);
      exit;
  end;
  
  if not SelectDirectory('Move file to...', '', sOutPath) then exit;
  sOutPath := IncludeTrailingPathDelimiter(sOutPath);
  
  sPathFilename := ZipFName.Caption; // preserve filename
  sFilename := ExtractFilename(sPathFilename);
  mnuCloseAchiveClick(nil); // close it firstly
  if MoveFile(sPathFilename, sOutPath) then
      sThisFile := sOutPath + sFilename
  else
      sThisFile := sPathFilename;

  if FileExists(sThisFile) then begin // open new target or previous file
      if btnShowVirtual.Down then begin
          SetZipFile(sThisFile, False);
          StartUnzip(True, sThisFile, '', nil, False);
      end
      else
          SetZipFile(sThisFile, True);

      ListView1.SetFocus;
      // Are some files extracted to Windows\Temp\xxx ? True to remove it.
      if bExtractedFiles and (edShadow.Text <> '') then
          Remove_TempFiles_UnderWinTemp;

      RecentUnzipFilesToReg(sThisFile);
      bExtractedFiles := False;
  end
  else begin
      beep;
      MessageDlg('Failing to open target zip file. File not found.',
      mtError, [mbOK], 0);
      exit;
  end;
end;

function TfrmMain.MoveFile(const sSource: string; const sTargetPath: string): boolean;
var
  sChosenPath, sFilename, sSourcePath, sOutPath: string;
  bSuccess, bCancel: boolean;
begin
  bSuccess := False;
  bCancel := False;
  if not FileExists(sSource) then begin
      beep;
      MessageDlg('Source file does not exist. Failing to move file.',
      mtError, [mbOK], 0);
      Result := False;
      exit;
  end;

  sSourcePath := AnsiLowercase( ExtractFileDir(sSource) );
  sSourcePath := IncludeTrailingPathDelimiter(sSourcePath);
  sFilename := ExtractFilename(sSource);

  //if SelectDirectory('Move file to...', '', sOutPath) then begin
      sOutPath := sTargetPath;
      sOutPath := IncludeTrailingPathDelimiter(sOutPath);

      if AnsiLowercase(sOutPath) = sSourcePath then begin
          beep;
          MessageDlg('Source and target folder are same. Moving file ' +
          'wil not continue.', mtError, [mbOK], 0);
          Result := False;
          exit;
      end;

      if DirectoryExists(sOutPath) then begin
          if FileExists(sOutPath + sFilename) then begin
              if MessageDlg('File already exists. Do you want to overwrite?',
              mtConfirmation, [mbYes, mbNo], 0) <> mrYes then begin
                  Result := False;
                  exit;
              end
              else begin// Yes
                  if not DeleteFile(sOutPath + sFilename) then begin
                      beep;
                      MessageDlg('Failing to overwrite existed file.' +
                      ' Possible reason is ' +
                      'read-only file.', mtError, [mbOK], 0);
                      Result := False;
                      exit;
                  end;
              end;
          end;
              bSuccess := RenameFile(sSource, sOutPath + sFilename);
      end
      else begin
          beep;
          MessageDlg('Target folder does not exist.', mtError, [mbOK], 0);
          Result := False;
          exit;
      end;
  //end
  //else // bCancelProcess = True
  //    bCancel := True;

  if (not bSuccess) and (not bCancel) then begin
      beep;
      MessageDlg('Failing to move file.', mtError, [mbOK], 0);
  end;
  Result := bSuccess;
end;

function TfrmMain.GetInputOfFilename(sFilenameIs: string): string;
var
  bInvalidChr, bPressOK: boolean;
  sGetInput: string;
begin
  Result := '';
  bPressOK := False;
  sGetInput := ChangeFileExt(sFilenameIs, '');
  repeat
    if InputQuery(XcProgramName, 'Rename file : ' + sFilenameIs + #13#10 +
    #13#10 + 'Please input a new filename without file extension :',
    sGetInput) = False then begin // Cancel
        //Result := 'QuitRenameFile';
        exit;
    end
    else begin
        bInvalidChr := False;
        if AnsiPos('\', sGetInput) > 0 then
            bInvalidChr := True
        else if AnsiPos('/', sGetInput) > 0 then
            bInvalidChr := True
        else if AnsiPos(':', sGetInput) > 0 then
            bInvalidChr := True
        else if AnsiPos('*', sGetInput) > 0 then
            bInvalidChr := True
        else if AnsiPos('?', sGetInput) > 0 then
            bInvalidChr := True
        else if AnsiPos('"', sGetInput) > 0 then
            bInvalidChr := True
        else if AnsiPos('<', sGetInput) > 0 then
            bInvalidChr := True
        else if AnsiPos('>', sGetInput) > 0 then
            bInvalidChr := True
        else if AnsiPos('|', sGetInput) > 0 then
            bInvalidChr := True;

        if bInvalidChr then begin
            beep;
            MessageDlg('Filename may not include below characters.' +
             #13#10 + '\ / : * ? " < > |', mtError, [mbOK], 0);
            //sGetInput := '';
        end
        else begin
            if sGetInput <> '' then begin
                bPressOK := True;
                Result := sGetInput;
            end
            else begin
                beep;
                MessageDlg('Empty string found.', mtError, [mbOK], 0);
            end;
        end;
    end;
  until bPressOK;
end;

procedure TfrmMain.mnuSubRenameArchiveClick(Sender: TObject);
var
  sPathFilename, sPath, sFilename, sFileExt, sNewFilename: string;
  f: file;
begin
  if ZipMaster1.ZipFileName = '' then begin
      beep;
      MessageDlg('There has not opened file.', mtError, [mbOK], 0);
      exit;
  end;
  
  sPathFilename := ZipFName.Caption;
  sPath := ExtractFileDir(sPathFilename);
  sPath := IncludeTrailingPathDelimiter(sPath);
      
  sFilename := ExtractFilename(sPathFilename);
  sFileExt := ExtractFileExt(sPathFilename);
  sNewFilename := GetInputOfFilename(sFilename);
  if sNewFilename <> '' then begin
      sNewFilename := sNewFilename + sFileExt;
      AssignFile(f, sPathFilename);
      try
        mnuCloseAchiveClick(nil); // close it firstly
        Rename(f, sPath + sNewFilename);

        if FileExists(sPath + sNewFilename) then begin // open new target or previous file
            if btnShowVirtual.Down then begin
                SetZipFile(sPath + sNewFilename, False);
                StartUnzip(True, sPath + sNewFilename, '', nil, False);
            end
            else
                SetZipFile(sPath + sNewFilename, True);

            ListView1.SetFocus;
            // Are some files extracted to Windows\Temp\xxx ? True to remove it.
            if bExtractedFiles and (edShadow.Text <> '') then
                Remove_TempFiles_UnderWinTemp;

            RecentUnzipFilesToReg(sPath + sNewFilename);
            bExtractedFiles := False;
        end
        else begin
            beep;
            MessageDlg('Failing to open new target zip file. File not found.',
            mtError, [mbOK], 0);
            exit;
        end;
      except on E: Exception do
        begin
          beep;
          MessageDlg('Failing to rename file.' +
          ' Please verifying it is not a duplicated filename' +
          Chr(13) + Chr(10) + 'Error # '  + E.Message, mtError, [mbOK], 0);
        end;
      end;
  end;
end;

procedure TfrmMain.mnuSubDeleteArchiveClick(Sender: TObject);
var
  SHFileOp: TSHFILEOPSTRUCT;
  sPathFilename: string;
begin
  if ZipMaster1.ZipFileName = '' then begin
      beep;
      MessageDlg('There has not opened file.', mtError, [mbOK], 0);
      exit;
  end;

  if MessageDlg('Do you really want to delete opened archive?',
  mtConfirmation, [mbYes, mbCancel], 0) <> mrYes then
      exit;

  sPathFilename := ZipFName.Caption;
  if not FileExists(sPathFilename) then begin
      beep;
      MessageDlg('Unexpected error. File not found.', mtError, [mbOK], 0);
      exit;
  end;

  mnuCloseAchiveClick(nil); // close it firstly
  with SHFileOp do begin
      Wnd    := Application.Handle;
      wFunc  := FO_DELETE;
      pFrom  := pChar( sPathFilename + #0 );
      pTo    := nil;
      fFlags := FOF_SILENT or FOF_NOCONFIRMATION;
      if bDeletedFiles_To_RecycleBin then
          fFlags := fFlags or FOF_ALLOWUNDO;
          
  end;
  SHFileOperation( SHFileOp );
end;

procedure TfrmMain.SetBool_CheckFileToRecycleBin(bChecked: boolean);
begin
  bDeletedFiles_To_RecycleBin := bChecked;
end;

procedure TfrmMain.SetBool_CheckFavoriteOpen(bChecked: boolean);
begin
  bFavorite_OpenFolder := bChecked;
end;

procedure TfrmMain.SetBool_CheckFavoriteAdd(bChecked: boolean);
begin
  bFavorite_AddFolder := bChecked;
end;

procedure TfrmMain.SetBool_CheckFavoriteExtract(bChecked: boolean);
begin
  bFavorite_ExtractFolder := bChecked;
end;

procedure TfrmMain.SetStr_FavoriteOpen(sPath: string);
begin
  sFavorite_OpenFolder := sPath;
end;

procedure TfrmMain.SetStr_FavoriteAdd(sPath: string);
begin
  sFavorite_AddFolder := sPath;
end;

procedure TfrmMain.SetStr_FavoriteExtract(sPath: string);
begin
  sFavorite_ExtractFolder := sPath;
end;

procedure TfrmMain.SetStr_VirusScannerProg(sPathFName: string);
begin
  sVirusScanner_PathFName := sPathFName;
end;

procedure TfrmMain.SetStr_VirusScannerPara(sStringIs: string);
begin
  sVirusScanner_Para := sStringIs;
end;

procedure TfrmMain.SetBool_CntOf_ColumnChanged_OnMainLst(bChanged: boolean);
begin
  bCol_AddedOrRemoved_OnMainLst := bChanged;
end;

procedure TfrmMain.ClearColumnsImage_OnlvMain;
//var
  //i: integer;
  //AColumn: TLVColumnEx;
  //AColumn: THdItem;
begin
  VT.Header.SortColumn := NoColumn;
 { AColumn.mask := 0; // important! reset
  AColumn.fmt := 0;  // important! reset
  AColumn.iImage := -1; // important! reset
  iPreImgIndex := 0; // important! reset

  for i := 0 to ListView1.Columns.Count -1 do begin
    with AColumn do begin
      Mask := HDI_FORMAT;
      Header_GetItem(GetDlgItem(ListView1.Handle, 0), i, AColumn);
      Mask := HDI_BITMAP or HDI_FORMAT;

      fmt := fmt And Not HDF_BITMAP;

      Header_SetItem(GetDlgItem(ListView1.Handle, 0), i, AColumn);
    end;
  end;

  for i := 0 to ListView1.Columns.Count -1 do begin
    AColumn.mask := LVCF_FMT;
    AColumn.fmt  := AColumn.fmt and not LVCFMT_IMAGE and not LVCFMT_BITMAP_ON_RIGHT;
    SendMessage(ListView1.Handle, LVM_SETCOLUMN, i, Longint(@AColumn));
  end; }
end;

procedure TfrmMain.Sort_SubMenuClick(Sender: TObject);
var
  NewMenuItem: TMenuItem;
  i: integer;
  sCap: string;
begin
  NewMenuItem := (Sender as TMenuItem);
  if NewMenuItem.Checked then begin
      sCap := NewMenuITem.Caption;
      sCap := AnsiReplaceStr(sCap, '&', '');
      WriteIniToReg(XcSubKeyIs, cMainMenuSort, sCap);
  end
  else // uncheck, use By Default
      WriteIniToReg(XcSubKeyIs, cMainMenuSort, '');
      
end;

procedure TfrmMain.mnuSubByNameClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByName);
end;

procedure TfrmMain.mnuSubByModifiedClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByModified);
end;

procedure TfrmMain.mnuSubBySizeClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubBySize);
end;

procedure TfrmMain.mnuSubByRatioClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByRatio);
end;

procedure TfrmMain.mnuSubByPackedClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByPacked);
end;

procedure TfrmMain.mnuSubByFileTypeClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByFileType);
end;

procedure TfrmMain.mnuSubByCRCClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByCRC);
end;

procedure TfrmMain.mnuSubByAttributesClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByAttributes);
end;

procedure TfrmMain.mnuSubByPathClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByPath);
end;

procedure TfrmMain.mnuSubByDefaultClick(Sender: TObject);
begin
  Sort_SubMenuClick(mnuSubByDefault);
end;

function TfrmMain.Get_ColIndex_ToSort_FromMenuSortBy: integer;
var
  i, j: integer;
  sCap: string;
begin
  Result := -1;
  for i := 0 to mnuSubSort.Count -1 do begin
    if mnuSubSort.Items[i].Checked then begin
        sCap := mnuSubSort.Items[i].Caption;
        AnsiReplaceStr(sCap, '&', '');
        for j := 0 to ListView1.Columns.Count -1 do begin
          if AnsiPos(ListView1.Column[j].Caption, sCap)
          > 0 then begin
              Result := j;
              break;
          end
        end;
        break;
    end;
  end;
end;

procedure TfrmMain.SetBool_CheckTxtWithNotePad(bChecked: boolean);
begin
  bOpen_TextFile_WithNotePad := bChecked;
end;


procedure TfrmMain.Update_lblCurDir(sDirIs: string);
begin
  lblCurDir.Caption := sDirIs;
  imgCurDir.Visible := (lblCurDir.Caption <> '');
end;

procedure TfrmMain.SetBool_CheckHiddenFiles(bChecked: boolean);
begin
  bDontShow_HiddenFiles := bChecked;
end;

procedure TfrmMain.SetBool_CheckSystemFiles(bChecked: boolean);
begin
  bDontShow_SystemFiles := bChecked;
end;

procedure TfrmMain.SetBool_RadioButtonParaFilename(bChecked: boolean);
begin
  bParaFilename := bChecked;
end;


procedure TfrmMain.mnuRenameClick(Sender: TObject);
var
  ZipRenameList: TList;
  RenRec: pZipRenameRec;
  sPathFilename, sFilename, sZipFolder, sFileExt, sNewFilename: string;
  sNewPathFilename: string;
  bIsVirtualFolder, bIsEncrypted: boolean;
  i, iRet: integer;
  Node: PVirtualNode;
  Data: PEntry;
begin
  if VT.SelectedCount = 0 then begin
      beep;
      MessageDlg('Nothing selected.', mtInformation, [mbOK], 0);
      exit;
  end;

  if VT.SelectedCount > 1 then exit;

  ZipRenameList := TList.Create( );
  try
    sPathFilename := GetPathFileName_FromSelected_Item(bIsVirtualFolder, bIsEncrypted);
    if (sPathFileName <> '') and (bIsVirtualFolder = False) then begin
        //bEncrypted := (AnsiPos('*', ListView1.Selected.Caption) > 0);
        sFilename := ExtractFilename(sPathFileName); // File name
        sZipFolder := ExtractFileDir(sPathFilename); // Folder name
        if sZipFolder <> '' then
            sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);

       { sZipFolder := AnsiReplaceStr(sZipFolder, '\', '/');
        if (sZipFolder <> '') and ( copy(sZipFolder, length(sZipFolder), 1)
        <> '/' ) then
            sZipFolder := sZipFolder + '/'; }

        sFileExt := ExtractFileExt(sPathFilename); // File Extension
        //sModifiedDate := ListView1.Selected.SubItems.Strings[0];
        sNewFilename := GetInputOfFilename(sFilename);
        if sNewFilename <> '' then begin
            sNewFilename := sNewFilename + sFileExt;
            sNewPathFilename := sZipFolder + sNewFilename;

            for i := 0 to ZipMaster1.ZipContents.Count -1 do begin
              with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
                if AnsiLowercase(sNewPathFilename) = AnsiLowercase(Filename) then begin
                    beep;
                    MessageDlg('Filename already exists.', mtError, [mbOK], 0);
                    exit;
                end;
              end;
            end;
            
            New( RenRec );
            RenRec^.Source := sPathFileName;
            RenRec^.Dest := sNewPathFilename;
            RenRec^.DateTime := 0; //DateTimeToFileDate( Today);//StrToDateTime(sModifiedDate) );
            ZipRenameList.Add( RenRec );
            iRet := ZipMaster1.Rename( ZipRenameList, 0 );
            if iRet = 0 then begin // success
                Node := VT.GetFirstSelected;
                if Node <> nil then begin
                    Data := VT.GetNodeData(Node);
                    if bIsEncrypted then
                        Data.Value[0] := sNewFilename + '*'
                    else 
                        Data.Value[0] := sNewFilename; // Name

                    Screen.Cursor := crHourGlass;
                    try
                      self.Update_FDataList;
                      VT.InvalidateNode(Node);
                    finally
                      Screen.Cursor := crDefault;
                    end;
                    //VT.Refresh;
                end
                else begin
                    beep;
                    MessageDlg('Filename was updated to zip file but ' +
                    'an error occured when updating the data on List View. ' +
                    'Please close this archive and open it again to get ' +
                    'updated contents of zip file on List View.' , mtError, [mbOK], 0);
                end;
            end
            else begin
                beep;
                MessageDlg('Failing to rename file. Error code: ' + inttostr(iRet), mtError, [mbOK], 0);
            end;

            Dispose( RenRec );
        end;
    end;
  finally
    ZipRenameList.Free( );
  end;
end;


procedure TfrmMain.mnuSubRenameClick(Sender: TObject);
begin
  mnuRenameClick(nil);
end;

procedure TfrmMain.mnuAddCommentClick(Sender: TObject);
begin
  pm1AddCommentClick(nil);
end;

procedure TfrmMain.smHomePageClick(Sender: TObject);
var
  RetValue: integer;
  sHomePage: string;
begin
  sHomePage := 'http://www.ccyjchk.com/';
  if sHomePage <> '' then begin
      RetValue := ShellExecute(handle, PChar('open'),
      PChar(sHomePage), '', '', SW_SHOW); // if success, return different value where depends on program

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

function TfrmMain.FreeDiskSpace(sDriveIs: string; var bNoErr: boolean): Int64;
var
  FreeBytesAvailableToCaller:  TLargeInteger;    // Int64
  TotalNumberOfBytes        :  TLargeInteger;
  TotalNumberOfFreeBytes    :  TLargeInteger;
begin
  if sDriveIs = '' then exit;

  sDriveIs := IncludeTrailingPathDelimiter(sDriveIs);
  bNoErr := GetDiskFreeSpaceEx(PAnsiChar(sDriveIs),
                               FreeBytesAvailableToCaller,
                               TotalNumberOfBytes,
                               @TotalNumberOfFreeBytes);

  if bNoErr then
      Result := TotalNumberOfFreeBytes;
      
end;

function TfrmMain.LevelOf_FreeDiskSpace_Extract(sDriveIs: string): integer;
var
  SectorsPerCluster, BytesPerSector, NumberOfFreeClusters: Cardinal;
  TotalNumberOfClusters, FreeBytes: Cardinal;
  sLetter: string;
begin
  // Return -1, 0 or 1. -1 = dangerous, 0 = normal, 1 = quite safe
  Result := 0;
  if sDriveIs = '' then exit;
  sDriveIs := IncludeTrailingPathDelimiter(sDriveIs);

  try
    GetDiskFreeSpace(PAnsiChar(sDriveIs), SectorsPerCluster, BytesPerSector,
    NumberOfFreeClusters, TotalNumberOfClusters);
    FreeBytes := NumberOfFreeClusters * SectorsPerCluster * BytesPerSector;

    sLetter := copy(sDriveIs, 1, 1);
    sLetter := AnsiLowercase(sLetter);

    if FreeBytes >= (ZipMaster1.ZipFileSize * 2) then
        Result := 1
    else if FreeBytes <= ZipMaster1.ZipFileSize then
        Result := -1
    else if (sLetter = 'a') and (ZipMaster1.ZipFileSize >= 300000) then // 300 KB
         Result := -1;
         
  except

  end;
end;

function TfrmMain.Get_FileSize(sPathFName: string): integer;
var
  s: TSearchRec;
begin
  Result := 0;
  if not FileExists(sPathFName) then exit;

  FindFirst(sPathFName, faAnyFile, s);
  Result := s.Size;
  FindClose(s);
end;

function TfrmMain.FilterSysOrHidFile_And_Get_FileSize(sPathFName: string;
 var bSysOrHid_File_BeRemoved: boolean): integer;
var
  s: TSearchRec;
begin
  Result := 0;
  bSysOrHid_File_BeRemoved := False;
  if not FileExists(sPathFName) then exit;

  FindFirst(sPathFName, faAnyFile, s);
  Result := s.Size;

  if self.RadioSystem.ItemIndex = 0 then begin // No system file
      if (s.Attr and faSysFile) <> 0 then // same attribute
          bSysOrHid_File_BeRemoved := True;

  end;

  if self.RadioHidden.ItemIndex = 0 then begin // No hidden file
      if (s.Attr and faHidden	) <> 0 then // same attribute
          bSysOrHid_File_BeRemoved := True;

  end;
  FindClose(s);
end;

function TfrmMain.Get_Predict_Compressed_FileSize(sPathFName: string): integer;
var
  s: TSearchRec;
  sFileExt: string;
begin
  Result := -1;
  if not FileExists(sPathFName) then exit;

  sFileExt := ExtractFileExt(sPathFName);
  sFileExt := AnsiLowercase(sFileExt);
  FindFirst(sPathFName, faAnyFile, s);
  // assuming file is about 100 kb or over because free disk space less than 100 kb will be intecepted
  if (sFileExt = '.exe') or (sFileExt = '.txt') or (sFileExt = '.doc') then
      Result := round(s.Size * 0.5) // compressed ratio 50%
  else if (sFileExt = '.htm') or (sFileExt = '.html') or (sFileExt = '.bmp') then
      Result := round(s.Size * 0.37) // compressed ratio 63%
  else if (sFileExt = '.jpg') or (sFileExt = '.jpeg') then begin
      if s.Size > 100000 then // greater than 100 kb
          Result := round(s.Size * 0.85); // compressed ratio 15%

  end
  else if sFileExt = '.hlp' then
      Result := round(s.Size * 0.7) // compressed ratio 30%
  else
      Result := s.Size;
      
  FindClose(s);
end;

procedure TfrmMain.mnuRepairClick(Sender: TObject);
begin
 { if ZipFname.Caption = '' then begin
      ShowMessage('There has not opened file.');
      exit;
  end; }

  with TfrmRepair.Create(Self) do
  try
    edRepairArchive.Text := ZipFname.Caption;
    showmodal;
  finally
    free;
  end;
end;

procedure TfrmMain.ZipMaster1ItemProgress(Sender: TObject; Item: String;
  TotalSize: Cardinal; PerCent: Integer);
begin
  // When handling a file, this will come in twice. First is initial and next is real
  //Gauge2.Progress := PerCent;
end;

procedure TfrmMain.mnuSubSearchClick(Sender: TObject);
begin
  Search.Show;
end;

function TfrmMain.ChangeFolder_OnVirtualTree_After_SearchFile(sPathIs: string):
  boolean;
var
  i, j, k, iRetPos, iCnt, iFrom: integer;
  sTemp, sDirIs: string;
  sArrayDir: array of string;
  UpTreeNode, NodeIs: TTreeNode;
  bFoundSameNode: boolean;
begin
  Result := False;
  sTemp := sPathIs;
  if sTemp <> '' then
      sTemp := IncludeTrailingPathDelimiter(sTemp); // important! below testing will add last dir

  if sTemp <> '' then begin
      k := 0;
      repeat
        iRetPos := AnsiPos('\', sTemp);
        if iRetPos > 0 then begin
            SetLength(sArrayDir, k+1);
            sDirIs := copy(sTemp, 1, iRetPos-1); // dir without "\"
            sArrayDir[k] := sDirIs;
            sTemp := copy(sTemp, iRetPos+1, length(sTemp)-iRetPos);
            Inc(k);
        end;
      until iRetPos = 0;

      NodeIs := nil;
      UpTreeNode := nil;
      iFrom := 1; // 0 is filename on top item
      for i := Low(sArrayDir) to High(sArrayDir) do begin
        bFoundSameNode := False;
        for j := iFrom to TreeView1.Items.Count -1 do begin
          if TreeView1.Items.Item[j].Level = i +1 then begin // starting level is 1, so i +1
              if AnsiLowercase(TreeView1.Items.Item[j].Text) = AnsiLowercase(sArrayDir[i]) then begin
                  bFoundSameNode := True;
                  iFrom := j +1; // Start from next node
                  NodeIs := TreeView1.Items.Item[j]; // Preserve found node
                  break;
              end;
          end;
        end;

        if bFoundSameNode = False then begin
            MessageDlg('Unexpected error. Failing to find correct location of ' +
            'the searched file.', mtError, [mbOK],0);
            exit;
        end;
      end;

      if bFoundSameNode and (NodeIs <> nil) then begin
          Result := True;
          TreeView1.Items.Item[NodeIs.AbsoluteIndex].Selected := True; // update selected tree
      end
      else begin
          MessageDlg('Unexpected error. Failing to find correct location of ' +
          'the searched file.', mtError, [mbOK],0);
          exit;
      end;
  end
  else begin // is / dir
      Result := True;
      TreeView1.Items.Item[0].Selected := True;
  end;
end;
procedure TfrmMain.btnVTupClick(Sender: TObject);
begin
  if TreeView1.Selected = nil then exit;
  if TreeView1.Selected.IsFirstNode then exit;
  ChangeFolder_OnVirtualMainList_Click('..');
end;

procedure TfrmMain.btnVThomeClick(Sender: TObject);
begin
  if TreeView1.Items.Count = 0 then exit;
  TreeView1.Items[0].Selected := True;
end;

procedure TfrmMain.btnVTcollapseClick(Sender: TObject);
begin
  if TreeView1.Items.Count = 0 then exit;
  TreeView1.FullCollapse;
  TreeView1_Change;
end;

procedure TfrmMain.mnuActionsClick(Sender: TObject);
begin
  SetStatusOfPopupUnzip(nil);
end;

procedure TfrmMain.mnuVirusScanClick(Sender: TObject);
var
  sPath_VSProg, sCommandPara: string;
  RetValue: Cardinal;
begin
  if sVirusScanner_PathFName = '' then begin
      beep;
      MessageDlg('Please configure Virus Scanner! Go to Settings -> ' +
      'Preferences -> Program Location -> Virus Scanner.', mtInformation, [mbOK], 0);
      exit;
  end;

  if not FileExists(sVirusScanner_PathFName) then begin
      beep;
      MessageDlg('Virus Scanner not found. Please make sure path and filename ' +
      'are still correct! Go to Settings -> Preferences -> Program Location ' +
      '-> Virus Scanner.', mtInformation, [mbOK], 0);
      exit;
  end;

  sPath_VSProg := sVirusScanner_PathFName;
  sPath_VSProg := ExtractShortPathName(sPath_VSProg); // run program
  sCommandPara := sVirusScanner_Para; // parameters
  sCommandPara := Trim(sCommandPara); // trim

  if bParaFilename then // assume AntiVir
      sCommandPara := '"' + sCommandPara + ' ' + ZipFName.Caption + '"'
  else begin // Filename + Para
      if sCommandPara = '' then
          sCommandPara := ZipFName.Caption
      else
          sCommandPara := '"' + ZipFName.Caption + ' ' + sCommandPara + '"';
          
  end;

  RetValue := ShellExecute(handle, 'open', PChar(sPath_VSProg),
  PChar(sCommandPara), '', SW_SHOW);

  if RetValue = 31 then begin // Has not the filename association?
       beep;
       MessageDlg('Failing to open specified program.', mtError, [mbOK], 0);
  end;
end;

procedure TfrmMain.Reading_Filenames_InFolder(sPathIs: string;
  sWildcards: string; var RetStrings: TStrings);
var
  sr: TSearchRec;
  iFileAttrs: Integer;
  //iRetValue: Integer;
begin
  if not DirectoryExists(sPathIs) then Exit;

  if sPathIs <> '' then
      sPathIs := IncludeTrailingPathDelimiter(sPathIs);

  iFileAttrs := faAnyFile;  //faArchive;   //archives only
  if FindFirst(sPathIs + sWildcards, iFileAttrs, sr) = 0 then begin
      repeat
        // if ((sr.Attr and iFileAttrs) = sr.Attr) and // (SearchRec.Attr and faHidden) <> 0 mean True, = 0 means False
        if (sr.Name <> '.') and (sr.Name <> '..') then begin
            try
              if FileExists(sPathIs + sr.Name) then
                  RetStrings.Append(sPathIs + sr.Name);

            except

            end;
        end;
      until FindNext(sr) <> 0;
      FindClose(sr);
  end
end;


procedure TfrmMain.Set_MainListView_ToUse_SystemIcons;
var
  FileInfo: TSHFileInfo;
  BM, BM2: TBitmap;
  //sFileType: string;
  //iIndex: integer;
begin
  // Get system icons
  ilSysIcons.Handle := SHGetFileInfo('', FILE_ATTRIBUTE_READONLY, FileInfo,
  SizeOf(TSHFileInfo), SHGFI_SYSICONINDEX or SHGFI_SMALLICON or SHGFI_TYPENAME);

  ilSysIcons.ShareImages := True; // If you found system icons were changed in Windows, setting ShareImages to False and run program again. When you found system icons disappeared, restart Windows and system icons will be restored.
  //APrevColor := ilSysIcons.BkColor; // preserve it! When form is destroyed, restore original. This is important to make no affect on system icons
  //ilSysIcons.BkColor := clBtnFace; // Sometimes, it affects background colour of icons
  //ilSysIcons.Assign(ilMainList); // XP does not accept this method

  BM := TBitmap.Create;  // by me
  BM2 := TBitmap.Create;  // by me
  try
    BM.Canvas.Brush.Color := clBtnFace;
    BM.Canvas.FillRect(Rect(0,0,16,16));

    if ilSystem.GetBitmap(0, BM) then begin
        BM.PixelFormat := pf24bit; // by me, change source 256 col to true col. This makes background col is completely transparent
        BM.Canvas.FloodFill(0, 0, BM.Canvas.Pixels[0, 0], fsSurFace);
        ArrowDown := BM.Handle;
        iIndex_Descend_SysIcons := ArrowDown;
        //iIndex_Descend_SysIcons := ilSysIcons.Add(BM, nil);
    end;

    BM2.Canvas.Brush.Color := clBtnFace;
    BM2.Canvas.FillRect(Rect(0,0,16,16));

    if ilSystem.GetBitmap(1, BM2) then begin
        BM2.PixelFormat := pf24bit; // by me, change source 256 col to true col. This makes background col is completely transparent
        BM2.Canvas.FloodFill(0, 0, BM2.Canvas.Pixels[0, 0], fsSurFace);
        ArrowUp := BM2.Handle;
        iIndex_Ascend_SysIcons := ArrowUp;
        //iIndex_Ascend_SysIcons := ilSysIcons.Add(BM, nil);
    end;
  finally
    BM.ReleaseHandle;
    BM.Free;
    BM2.ReleaseHandle;
    BM2.Free;
  end;
 { iIndex := Get_SysIconIndex_Of_Given_PathFileExt('.bmp', sFileType);
  ilMainList.GetIcon(iIndex, AIcon);
  ilSystem.InsertIcon(ilSystem.Count, AIcon);

  iIndex := Get_SysIconIndex_Of_Given_PathFileExt('.emf', sFileType);
  ilMainList.GetIcon(iIndex, AIcon);
  ilSystem.InsertIcon(ilSystem.Count, AIcon);

  iIndex := Get_SysIconIndex_Of_Given_PathFileExt('.gif', sFileType);
  ilMainList.GetIcon(iIndex, AIcon);
  ilSystem.InsertIcon(ilSystem.Count, AIcon);

  iIndex := Get_SysIconIndex_Of_Given_PathFileExt('.htm', sFileType);
  ilMainList.GetIcon(iIndex, AIcon);
  ilSystem.InsertIcon(ilSystem.Count, AIcon);

  iIndex := Get_SysIconIndex_Of_Given_PathFileExt('.jpg', sFileType);
  ilMainList.GetIcon(iIndex, AIcon);
  ilSystem.InsertIcon(ilSystem.Count, AIcon);

  iIndex := Get_SysIconIndex_Of_Given_PathFileExt('.pcx', sFileType);
  ilMainList.GetIcon(iIndex, AIcon);
  ilSystem.InsertIcon(ilSystem.Count, AIcon);

  iIndex := Get_SysIconIndex_Of_Given_PathFileExt('.png', sFileType);
  ilMainList.GetIcon(iIndex, AIcon);
  ilSystem.InsertIcon(ilSystem.Count, AIcon);

  iIndex := Get_SysIconIndex_Of_Given_PathFileExt('.wmf', sFileType);
  ilMainList.GetIcon(iIndex, AIcon);
  ilSystem.InsertIcon(ilSystem.Count, AIcon);

  iCntImageListSystem := ilSystem.Count; // know actually fixed items on ilSystem
  }
end;

procedure TfrmMain.Insert_HistoryOpenWithApp_ToMenu;
var
  iCnt, iIndex: integer;
  i: integer;
  sPathFilename: string;
  AFiles: TStrings;
begin
  if mnuOpenWith_ThisApp.Tag = 0 then begin // not exist
      mnuOpenWith_ThisApp.Tag := 1;
      iIndex := mnuOpenWith.MenuIndex; // get index from PopupMenu1.mnuOPenWith
      mnuOpenWith_ThisApp.ImageIndex := mnuOpenWith.ImageIndex;
      mnuOpenWith_ThisApp.Enabled := mnuOpenWith.Enabled;
      mnuOpenWith_ThisApp.AutoHotkeys := maManual; // not need shourtcut symbol
      PopupMenu1.Items.Insert(iIndex +1, mnuOpenWith_ThisApp); // dynamically create

      AFiles := TStringList.Create;
      try
        RecentOpenWithAppFromReg(AFiles);
        //iCnt := length(mSubOpenWith_ThisApp);
        if AFiles = nil then exit;

        iCnt := AFiles.Count;
        SetLength(mSubOpenWith_ThisApp, iCnt);
        for i := 0 to iCnt -1 do begin
          mSubOpenWith_ThisApp[i] := TMenuItem.Create(Self);
          mSubOpenWith_ThisApp[i].Caption := AFiles.Strings[i];
          mSubOpenWith_ThisApp[i].OnClick := mSubOpenWith_ThisApp_OnClick;
          mnuOpenWith_ThisApp.Insert(i, mSubOpenWith_ThisApp[i]);
        end;
      finally
        AFiles.Free;
      end;
  end
  else if mnuOpenWith_ThisApp.Tag = 1 then  // exist
      mnuOpenWith_ThisApp.Enabled := mnuOpenWith.Enabled;
  
end;

procedure TfrmMain.RecentOpenWithAppToReg(const sPathFilenameIs: string);
var
  i: integer;
  sPathFilenameInMenu, sCap, sPreCap: string;
  iC, iCnt: integer;
begin
  // Is count over limit?
  if mnuOpenWith_ThisApp.Count > cCountOfRecentOpenWithApp then begin
      for i := mnuOpenWith_ThisApp.Count -1 downto cCountOfRecentOpenWithApp
      do begin
        mnuOpenWith_ThisApp.Delete(i);
      end;
  end;

  sPreCap := '';
  // Is path already exist?
  for i := mnuOpenWith_ThisApp.Count -1 downto 0 do begin
    sCap := AnsiLowercase(mnuOpenWith_ThisApp.Items[i].Caption);
    sCap := AnsiReplaceStr(sCap, '&', ''); // remove shortcut symbol
    if sCap = sPreCap then
        mnuOpenWith_ThisApp.Delete(i)
    else if sCap = AnsiLowercase(sPathFilenameIs) then
        exit;

    sPreCap := sCap;
  end;
  // Full?
  if mnuOpenWith_ThisApp.Count >= cCountOfRecentOpenWithApp then
      mnuOpenWith_ThisApp.Delete(mnuOpenWith_ThisApp.Count -1);

  iC := 0;
  if FileExists(sPathFilenameIs) then begin
      iCnt := length(mSubOpenWith_ThisApp);
      SetLength(mSubOpenWith_ThisApp, iCnt +1);
      mSubOpenWith_ThisApp[iC] := TMenuItem.Create(Self);
      mSubOpenWith_ThisApp[iC].Caption := sPathFilenameIs;
      mSubOpenWith_ThisApp[iC].OnClick := mSubOpenWith_ThisApp_OnClick;
      mnuOpenWith_ThisApp.Insert(iC, mSubOpenWith_ThisApp[iC]);
  end;

  // update to registry
  for i := 0 to mnuOpenWith_ThisApp.Count -1 do begin
    sCap := mnuOpenWith_ThisApp.Items[i].Caption;
    sCap := AnsiReplaceStr(sCap, '&', ''); // remove shortcut symbol
    sPathFilenameInMenu := sCap;
    WriteIniToReg(sSubKeyRecentOpenWithApp, inttostr(i), sPathFilenameInMenu);
  end;
end;

procedure TfrmMain.RecentOpenWithAppFromReg(var AFiles: TStrings);
var
  i: integer;
  sPathFilename: string;
begin
  for i := 0 to cCountOfRecentOpenWithApp -1 do begin
    sPathFilename := ReadIniFromReg(sSubKeyRecentOpenWithApp, inttostr(i));
    if sPathFilename = '' then break;
    if FileExists(sPathFilename) then
        AFiles.Append(sPathFilename);
        
  end;
end;

procedure TfrmMain.mSubOpenWith_ThisApp_OnClick(Sender: TObject);
var
  sPathFileName, sWithApplication: string;
  bIsVirtualFolder: boolean;
  AParentMenuItem: TMenuItem;
  AIndex: integer;
  bIsEncrypted: boolean;
begin
  sWithApplication := (Sender as TMenuItem).Caption;
  sWithApplication := AnsiReplaceStr(sWithApplication, '&', ''); // remove shortcut symbol
  if not FileExists(sWithApplication) then begin
      AParentMenuItem := (Sender as TMenuItem).Parent;
      if AParentMenuItem <> nil then begin
          AIndex := AParentMenuItem.IndexOf( (Sender as TMenuItem) );
          if AIndex <> -1 then
              AParentMenuItem.Delete(AIndex);
          
      end;

      beep;
      MessageDlg('The application file not found.', mtError, [mbOK], 0);
      exit;
  end;

  if VT.SelectedCount = 0 then begin
      beep;
      MessageDlg('Nothing selected on list.', mtError, [mbOK], 0);
      exit;
  end;
  sPathFileName := GetPathFileName_FromSelected_Item(bIsVirtualFolder, bIsEncrypted);
  if (sPathFileName <> '') and (bIsVirtualFolder = False) then
      ExecuteFileClick(nil, sPathFileName, False, sWithApplication);
      
end;

procedure TfrmMain.ZipMaster1Progress(Sender: TObject;
  ProgrType: ProgressType; Filename: String; FileSize: Cardinal);
var
    Step: Integer;
begin
    case ProgrType of
   { TotalFiles2Process: begin // intercept extracting and deleting files, extracting comes once, deleting comes in when every new file
                        end;
     TotalSize2Process: begin // intercept extracting and deleting files
                        end; }
               NewFile: begin // add, extract files, come every time to have new file found
                          Gauge2.MaxValue := FileSize;
                          Gauge2.Progress := 0;
                          Gauge1.Progress := Gauge1.Progress +1;
                          if HourGlass.Showing then begin
                              HourGlass.Gauge2.MaxValue := FileSize;
                              HourGlass.Gauge2.Progress := 0;
                              HourGlass.Gauge1.Progress := HourGlass.Gauge1.Progress +1;
                          end;
                        end;
        ProgressUpdate: begin // extracting files will come in. When adding files, every new file come here first berfore ExtraUpdate
                          Gauge2.Progress := Gauge2.Progress + FileSize;
                          if HourGlass.Showing then
                              HourGlass.Gauge2.Progress := HourGlass.Gauge2.Progress + FileSize;

                        end;
          { ExtraUpdate: begin // intercept adding files to archive

                        end;
              NewExtra: begin // intercept adding files when a new file found

                        end; }
            EndOfBatch: begin // Reset the Gauge and filename.
                          Gauge2.Progress := 0;
                          Gauge2.Visible := False;
                        end;
    end;                                // EOF Case
end;

{procedure TfrmMain.ThreadDone(Sender: TObject);
begin
  Dec(ThreadsRunning);
  if ThreadsRunning = 0 then
      bFinished_Threads := True;

end; }


{function TfrmMain.ExecAndWait(const ExecuteFile, ParamString : string): boolean;
var
  SEInfo: TShellExecuteInfo;
  ExitCode: DWORD;
begin
  FillChar(SEInfo, SizeOf(SEInfo), 0);
  SEInfo.cbSize := SizeOf(TShellExecuteInfo);
  with SEInfo do begin
    fMask := SEE_MASK_NOCLOSEPROCESS;
    Wnd := Application.Handle;
    lpFile := PChar(ExecuteFile);
    lpParameters := PChar(ParamString);
    nShow := SW_HIDE;
  end;
  if ShellExecuteEx(@SEInfo) then
  begin
    repeat
      Application.ProcessMessages;
      GetExitCodeProcess(SEInfo.hProcess, ExitCode);
    until (ExitCode <> STILL_ACTIVE) or Application.Terminated;
    Result:=True;
  end
  else Result:=False;
end; }


procedure TfrmMain.VTGetNodeDataSize(Sender: TBaseVirtualTree;
  var NodeDataSize: Integer);
begin
  NodeDataSize := SizeOf(TEntry);
end;

procedure TfrmMain.VTGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
  Column: TColumnIndex; TextType: TVSTTextType; var CellText: WideString);
var
  Data: PEntry;
  sColCap: string;
begin
  Data := Sender.GetNodeData(Node);
  CellText := '';
 // if TextType = ttNormal then begin
      case Column of
        0: begin // Name
             CellText := MBToUnicode(Data.value[0]);
           end;
        1: begin // Modified
             CellText := Data.value[1];
           end;
        2: begin // Size
             CellText := Data.value[2];
           end;
        3: begin // Ratio
             CellText := Data.value[3];
           end;
        4: begin // Packed
             CellText := Data.value[4];
           end;
        5: begin // Below from File type ... Path
                 sColCap := VT.Header.Columns[Column].Text;
                 if sColCap = 'File Type' then
                     CellText := MBToUnicode(Data.value[5])
                 else if sColCap = 'CRC' then
                     CellText := Data.value[6]
                 else if sColCap = 'Attributes' then
                     CellText := Data.value[7]
                 else if sColCap = 'Path' then
                     CellText := MBToUnicode(Data.value[8]);

           end;
        6: begin
                 sColCap := VT.Header.Columns[Column].Text;
                 if sColCap = 'File Type' then
                     CellText := MBToUnicode(Data.value[5])
                 else if sColCap = 'CRC' then
                     CellText := Data.value[6]
                 else if sColCap = 'Attributes' then
                     CellText := Data.value[7]
                 else if sColCap = 'Path' then
                     CellText := MBToUnicode(Data.value[8]);

           end;
        7: begin
                 sColCap := VT.Header.Columns[Column].Text;
                 if sColCap = 'File Type' then
                     CellText := MBToUnicode(Data.value[5])
                 else if sColCap = 'CRC' then
                     CellText := Data.value[6]
                 else if sColCap = 'Attributes' then
                     CellText := Data.value[7]
                 else if sColCap = 'Path' then
                     CellText := MBToUnicode(Data.value[8]);

           end;
        8: begin
                 sColCap := VT.Header.Columns[Column].Text;
                 if sColCap = 'File Type' then
                     CellText := MBToUnicode(Data.value[5])
                 else if sColCap = 'CRC' then
                     CellText := Data.value[6]
                 else if sColCap = 'Attributes' then
                     CellText := Data.value[7]
                 else if sColCap = 'Path' then
                     CellText := MBToUnicode(Data.value[8]);

           end;
      end; // case end
 // end;
end;

procedure TfrmMain.VTGetImageIndex(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
  var Ghosted: Boolean; var ImageIndex: Integer);
var
  Data: PEntry;
  sCaption, sFileExt, sFile_Type: string;
  iIndex: integer;
begin
  if Column = 0 then begin
      Data := Sender.GetNodeData(Node);
      sCaption := Data.Value[0];
      sCaption := AnsiReplaceStr(sCaption, '*' , ''); // remove password char
      if Uppercase(sCaption) = 'SETUP.EXE' then
          sFileExt := Uppercase(sCaption)
      else
          sFileExt := Uppercase( ExtractFileExt(sCaption) );
          
      iIndex := Get_SysIconIndex_Of_Given_FileExt(sFileExt, sFile_Type);
      ImageIndex := iIndex;
  end;
end;

procedure TfrmMain.VTHeaderClick(Sender: TVTHeader; Column: TColumnIndex;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbLeft then
  begin
    with Sender do
    begin
      if SortColumn <> Column then
      begin
        SortColumn := Column;
        SortDirection := sdAscending;
      end
      else
      case SortDirection of
        sdAscending:
          SortDirection := sdDescending;
        sdDescending:
          //SortColumn := NoColumn;
          SortDirection := sdAscending;
      end;
      Treeview.SortTree(SortColumn, SortDirection, False);
      //VT.SortTree( Column, VT.Header.SortDirection );
    end;

  end;
end;

procedure TfrmMain.VTCompareNodes(Sender: TBaseVirtualTree; Node1,
  Node2: PVirtualNode; Column: TColumnIndex; var Result: Integer);
var
  Data1, Data2: PEntry;
  iV1, iV2: Cardinal;
  Modified1, Modified2: TDateTime;
  sColumnCap, sTem1, sTem2: string;
begin
  Data1 := Sender.GetNodeData(Node1);
  Data2 := Sender.GetNodeData(Node2);

  sColumnCap := VT.Header.Columns[Column].Text;

  if AnsiPos('Name', sColumnCap) > 0 then begin// Name
      Result := AnsiCompareStr(Data1.Value[0], Data2.Value[0])
  end
  else if AnsiPos('Modified', sColumnCap) > 0 then begin // Modified
      if (Data1.Value[1] <> '') and (Data2.Value[1] <> '') then begin
          Modified1 := StrToDateTime(Trim(Data1.Value[1]));
          Modified2 := StrToDateTime(Trim(Data2.Value[1]));
          Result := CompareDateTime(Modified1,Modified2);
      end
      else
          Result := 0;

  end
  else if AnsiPos('Size', sColumnCap) > 0 then begin //
      if (Data1.Value[2] <> '') and (Data2.Value[2] <> '') then begin
          iV1 := round(strtoint(Trim(Data1.Value[2])));
          iV2 := round(strtoint(Trim(Data2.Value[2])));
          Result := CompareValue(iV1, iV2);
      end
      else
          Result := 0;
  end
  else if AnsiPos('Ratio', sColumnCap) > 0 then begin //
      sTem1 := Trim( AnsiReplaceStr(Data1.Value[3], '%', '') );
      sTem2 := Trim( AnsiReplaceStr(Data2.Value[3], '%', '') );
      if (sTem1 <> '') and (sTem2 <> '') then begin
          iV1 := round(strtoint(sTem1));
          iV2 := round(strtoint(sTem2));
          Result := CompareValue(iV1, iV2);
      end
      else
          Result := 0;
          
  end
  else if AnsiPos('Packed', sColumnCap) > 0 then begin //
      if (Data1.Value[4] <> '') and (Data2.Value[4] <> '') then begin
          iV1 := round(strtoint(Trim(Data1.Value[4])));
          iV2 := round(strtoint(Trim(Data2.Value[4])));
          Result := CompareValue(iV1, iV2);
      end
      else
          Result := 0;
          
  end
  else if AnsiPos('File Type', sColumnCap) > 0 then begin //
      Result := AnsiCompareStr(Data1.Value[5], Data2.Value[5]);
  end
  else if AnsiPos('CRC', sColumnCap) > 0 then begin //
      Result := AnsiCompareStr(Data1.Value[6], Data2.Value[6]);
  end
  else if AnsiPos('Attributes', sColumnCap) > 0 then begin //
      Result := AnsiCompareStr(Data1.Value[7], Data2.Value[7]);
  end
  else if AnsiPos('Path', sColumnCap) > 0 then begin //
      Result := AnsiCompareStr(Data1.Value[8], Data2.Value[8]);
  end;
end;

procedure TfrmMain.AddNodesToVirtualTree(TotalCnt: integer);
var
  Node: PVirtualNode;
  Data: PEntry;
  i: integer;
begin
  VT.BeginUpdate;
  try
    for i := 0 to TotalCnt -1 do begin
      Node := VT.AddChild(nil); // add top level node
      Data := VT.GetNodeData(Node);

      Data.Value[0] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).Caption.DelimitedText, '"', '' );
      Data.Value[1] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).Modified.DelimitedText, '"', '' );
      Data.Value[2] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).Size.DelimitedText, '"' , '' );
      Data.Value[3] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).Ratio.DelimitedText, '"' , '' );
      Data.Value[4] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).ColPacked.DelimitedText, '"' , '' );
      Data.Value[5] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).FileType.DelimitedText, '"' , '' );
      Data.Value[6] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).CRC.DelimitedText, '"' , '' );
      Data.Value[7] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).Attributes.DelimitedText, '"' , '' );
      Data.Value[8] := AnsiReplaceStr( TVTPropEditData( FDataList[Node.Index] ).Path.DelimitedText, '"' , '' );
    end;
  finally
    VT.EndUpdate;
  end;
end;

procedure TfrmMain.VTDblClick(Sender: TObject);
var
  sPathFileName: string;
  bIsVirtualFolder, bIsEncrypted: boolean;
begin
  if VT.SelectedCount <> 1 then exit;
  sPathFileName := GetPathFileName_FromSelected_Item(bIsVirtualFolder, bIsEncrypted); // get path name and filename first then execute
 { if (sPathFileName <> '') and (not bIsVirtualFolder) then // Normal view on ListView and it is a file
      ExecuteFileClick(nil, sPathFileName, True, '')
  else if (sPathFileName <> '') and bIsVirtualFolder then // Virtual view on ListView and it is folder
      ChangeFolder_OnVirtualMainList_Click(sPathFileName); }

  btnEditSelXML.Click;

end;

procedure TfrmMain.Update_FDataList;
var
  i, iTotalFiles: integer;
  sPath, sFilename, sFileExt, sFile_Type: string;
begin
  FDataList.Clear; // Empty the FDataList
  iTotalFiles := ZipMaster1.ZipContents.Count;

  Screen.Cursor := crHourGlass;
  try
    for i := 0 to iTotalFiles -1 do begin
      with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
        sPath := ExtractFilePath(FileName);
        sFilename := ExtractFileName(FileName);
        sFileExt := Uppercase( ExtractFileExt(FileName) );

        if sFilename <> '' then begin
            ped := TVTPropEditData.Create;
            if Encrypted then
                ped.Caption.Add( sFilename + '*' ) // Name
            else
                ped.Caption.Add( sFilename ); // Name

            ped.Modified.Add(FormatDateTime('ddddd  t',
            FileDateToDateTime(DateTime))); // Modified
            ped.Size.Add(IntToStr(UncompressedSize)); // Size
            if UncompressedSize <> 0 Then
                ped.Ratio.Add(IntToStr(Round((1 - (CompressedSize /
                UncompressedSize)) * 100)) + '% ') //Ratio
            else
                ped.Ratio.Add('0% ');

            ped.ColPacked.Add(IntToStr(CompressedSize)); // Packed

            Get_SysIconIndex_Of_Given_FileExt(sFileExt, sFile_Type);
            if sFile_Type <> '' then
                ped.FileType.Add(sFile_Type) // File Type
            else begin
                sFile_Type := AnsiReplaceStr(sFileExt, '.', ' ');
                ped.FileType.Add(sFile_Type);
            end;

            //if bCheckCRC then  // CRC
            ped.CRC.Add(inttohex(CRC32, 2));

            //if bCheckAttributes then  // Attributes
            ped.Attributes.Add(GetAttributes(ExtFileAttrib));

            ped.Path.Add(sPath); //Path

            FDataList.Add( ped );
        end;
      end; // with end
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmMain.VTIncrementalSearch(Sender: TBaseVirtualTree;
  Node: PVirtualNode; const SearchText: WideString; var Result: Integer);
var
  sCompare1, sCompare2 : string;
  DisplayText : WideString;

begin
      VT.IncrementalSearchDirection := sdForward;   // note can be backward

     // Note: This code requires a proper Unicode/WideString comparation routine which I did not want to link here for
     // size and clarity reasons. For now strings are (implicitely) converted to ANSI to make the comparation work.
     // Search is not case sensitive.
      VTGetText( Sender, Node, 0 {Column}, ttNormal, DisplayText );
      sCompare1 := SearchText;
      sCompare2 := DisplayText;

     // By using StrLIComp we can specify a maximum length to compare. This allows us to find also nodes
     // which match only partially. Don't forget to specify the shorter string length as search length.
     Result := StrLIComp( pAnsiChar(sCompare1), pAnsiChar(sCompare2), Min(Length(sCompare1), Length(sCompare2)) )
end;

procedure TfrmMain.VTKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
 { i, iStart: integer;
  AChr: Char;
  sA: string[1];
  bFound: boolean; }
  Pos: TPoint;
begin
  if VT.TotalCount = 0 then exit;
  //if (Key = vk_Control) or (Key = vk_Shift) or ( Key = vk_Up) then exit;
  if (Key = vk_Left) then begin
      VT.ClearSelection;
      exit;
  end;

 { if Shift = [ssCtrl] then begin
      if (Shift = [ssCtrl]) and (Key = 191) then begin // char '/'
          Pos := VT.ClientOrigin;
          PopupMenu1.Popup(Pos.X + 75, Pos.Y + 20);
      end;
      exit;
  end; }

  if Shift = [ssCtrl] then begin
      if (Shift = [ssCtrl]) and (Key = 191) then begin // char '/'
          Pos := VT.ClientOrigin;
          pmWord.Popup(Pos.X + 75, Pos.Y + 20);
      end;
      exit;
  end;
  
  if ( not ( (Key >= 48) and (Key <= 57) ) ) and // 0..9
  ( not ( (Key >= 65) and (Key <= 122) ) ) and // 'A..Z' and 'a..z'
  ( not (Key = 46) ) then // Delete
      exit;

  if Key = 46 then begin// Delete key
      mnuDeleteClick(nil);
      exit;
  end;

 { AChr := Char(Key);

  if btnShowVirtual.Down then
      ListView1.ClearSelection;

  if chPreSearch = AChr then
      iStart := iStartItemSearch + 1
  else
      iStart := 0;

  chPreSearch := AChr;
  bFound := False;
  for i := iStart to ListView1.Items.Count -1 do begin
    sA := copy(ListView1.Items[i].Caption, 1, 1);
    sA := uppercase(sA);
    if sA = AChr then begin
        ListView1.Items[i].Selected := True;
        ListView1.Items[i].MakeVisible(False);
        iStartItemSearch := i;
        bFound := True;
        break;
    end;
  end;

  if not bFound then
      iStartItemSearch := -1; }
end;

function TfrmMain.MBToUnicode(sFrom: string): WideString;
var
  Buffer: PWideChar;
  iSize, iNewSize, iRet: integer;
begin
  Result := sFrom;
  iSize := Length(sFrom) + 1;
  iNewSize := iSize * 2;

  {Allocate memory}
  GetMem(Buffer, iNewSize);
  try
    iRet := MultiByteToWideChar(CP_ACP, 0, PChar(sFrom), iSize, Buffer, iNewSize);
    if iRet <> -1 then
        Result := Buffer;

  finally
    FreeMem(Buffer);
  end;
end;

procedure TfrmMain.Output_Txt_Contents_Of_XML_From_Docx;
const
  cXML_Document = 'word\document.xml';
var
 // Filename : string;
 // f        : TEXTFILE;
  sMyTxt, sXML_Doc, sOutputPFilename: string;
  sSpaces: string;
  NN: TNvpNode;
  bTabFound: boolean;
  vEnter: variant;

  iPBar: integer;

  //frmPBar: TfrmPBar;
 // frmViewText: TMessageOutput;
begin
  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      ShowMessage('Err: Failing to extract text from xml file. Target directory not exist.' + #13#10);
      exit;
  end;

  sXML_Doc := XsExtractToThisPath + cXML_Document;

  if not FileExists(sXML_Doc) then begin
      ShowMessage('Err: Failing to extract text from xml file. Xml file not found.' + #13#10);
      exit;
  end;



  try
    Screen.Cursor := crHourGlass;
    //frmPBar := TfrmPBar.Create(nil);
    pnPB.Visible := True;

    iPBar := 0;

    // know length of progress bar
    if FileExists(sXML_Doc) then begin

        XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
        XmlParser.StartScan;
        XmlParser.Normalize := FALSE;
        while XmlParser.Scan do begin

          if (XmlParser.CurPartType = ptStartTag)	then begin
              if XmlParser.CurName = 'w:p' then
                  Inc(iPBar);

          end;

        end;

    end;


    lbPB.Caption := 'Loading...';
    GaugePB.MinValue := 0;
    GaugePB.MaxValue := iPBar;
    GaugePB.Progress := 0;
    //frmPBar.Show;


    sSpaces := '';
    sMyTxt := '';
    bTabFound := False;
    vEnter := #$D#$A;

    memOutput.Clear;

 { frmViewText := TMessageOutput.Create(self);
  with frmViewText do
  try
    Caption := 'Text Content';
    AutoSize := False;
    Width := 800;
    Height := 600;
    btnSaveFile.Visible := True; 
    memOutput.Clear;
    memOutput.Align := alClient;
    memOutput.WordWrap := False;
    memOutput.ScrollBars := ssBoth; }

    XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
    XmlParser.StartScan;
    XmlParser.Normalize := FALSE;
    while XmlParser.Scan do begin
      if (XmlParser.CurPartType = ptStartTag) or (XmlParser.CurPartType = ptEmptyTag)	then begin
          if XmlParser.CurName = 'w:jc' then begin
              NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('w:val'));
              if NN.Value = 'center' then
                  sSpaces := '               ';

              NN.Free;
          end
          else if XmlParser.CurName = 'w:tab' then begin
              sSpaces := '        ';
              bTabFound := True;
          end
          else if XmlParser.CurName = 'w:p' then begin
              if sMyTxt <> '' then begin // text found before this paragraph, print it first
                  if bTabFound then begin
                      memOutput.Lines.Append(''); // Payer wants adding one more line
                      bTabFound := False;
                  end;

                  memOutput.Lines.Append(sMyTxt);
                  sMyTxt := '';
              end
              else
                  memOutput.Lines.Append('');

              GaugePB.Progress := GaugePB.Progress + 1;
              GaugePB.Repaint;
              Application.ProcessMessages;

          end;
      end;

      if (XmlParser.CurPartType = ptContent) or (XmlParser.CurPartType = ptCData) then begin
          if XmlParser.CurName = 'w:t' then begin
              if AnsiPos(vEnter, XmlParser.CurContent) = 0 then // not accept #$D#$A
                  sMyTxt := sSpaces + sMyTxt + XmlParser.CurContent;

          end;

         { if ( XmlParser.CurName = 'w:t' ) or ( XmlParser.CurName = 'w:instrText' ) then begin
              if XmlParser.CurName = 'w:instrText' then begin
                  if AnsiPos(XmlParser.CurContent, vEnter) = 0 then
                      sMyTxt := sSpaces + sMyTxt + XmlParser.CurContent;

              end        
              else
                  sMyTxt := sSpaces + sMyTxt + XmlParser.CurContent;
          end; }

          sSpaces := '';
      end;
    end;


    if sMyTxt <> '' then begin
        if bTabFound then
            memOutput.Lines.Append(''); // Payer wants adding one more line

        memOutput.Lines.Append(sMyTxt);
    end;

    memOutput.SetFocus;
    memOutput.SelStart := 0;
    memOutput.SelLength := 0;

  finally
    //frmPBar.Free;
    pnPB.Visible := False;
    Screen.Cursor := crDefault;
  end;

  {  showmodal;

  finally
    frmViewText.Free;
  end; }



 { sOutputPFilename := XsExtractToThisPath + 'document.txt';

  if FileExists(sOutputPFilename) then // delete old one first if found
      DeleteFile(sOutputPFilename);

  Filename := sOutputPFilename; // output to
  AssignFile (f, Filename);
  TRY
    Rewrite (f);
    TRY
      XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
      XmlParser.StartScan;
      XmlParser.Normalize := FALSE;
      WHILE XmlParser.Scan DO begin

        if ( XmlParser.CurPartType = ptStartTag ) or
        ( XmlParser.CurPartType = ptEmptyTag )	then begin


            if XmlParser.CurName = 'w:jc' then begin
                NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('w:val'));
                if NN.Value = 'center' then
                    sSpaces := '               ';

                NN.Free;
            end
            else if XmlParser.CurName = 'w:tab' then
                sSpaces := '        '
            else if XmlParser.CurName = 'w:p' then
                WriteLn(f, '');


        end;

        sMyTxt := '';
        if (XmlParser.CurPartType = ptContent) OR
           (XmlParser.CurPartType = ptCData) THEN begin
            if not ( XmlParser.CurContent = #$D#$A ) then begin
                sMyTxt := XmlParser.CurContent;
                Write(f, sSpaces + sMyTxt);
            end;

            sSpaces := '';
        end;
      end;
    FINALLY
      CloseFile (f);
      END;


  EXCEPT

  END; }

end;

procedure TfrmMain.Output_Txt_Contents_Of_XML_From_Pptx;
//const
//  cXML_Document = 'ppt\slides\slide1.xml';
var
 // Filename : string;
 // f        : TEXTFILE;
  sMyTxt, sXML_Doc, sOutputPFilename, sErrorMessage: string;
  sSpaces: string;
  NN: TNvpNode;
  bTabFound: boolean;
  //vEnter: variant;
  i: integer;
 // frmViewText: TMessageOutput;
  slShort, slLong: TStringList;
  sFilename: string;
begin
  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      ShowMessage('Err: Failing to extract text from xml file. Target directory not exist.' + #13#10);
      exit;
  end;

  if XslOfficeDoc.Count = 0 then exit;

  sErrorMessage := '';
  memOutput.Clear;

  slShort := TStringList.Create;
  slLong := TStringList.Create;

  try // sorting

    for i := 0 to XslOfficeDoc.Count -1 do begin
      sFilename := ExtractFilename(XslOfficeDoc.Strings[i]);

      if length(sFilename) = 10 then
          slShort.Add(XslOfficeDoc.Strings[i]) // e.g: slide1.xml
      else
          slLong.Add(XslOfficeDoc.Strings[i]); // e.g: slide10.xml

    end;

    slShort.Sorted := True;
    slLong.Sorted := True;

    XslOfficeDoc.Clear;

    XslOfficeDoc.AddStrings(slShort);
    XslOfficeDoc.AddStrings(slLong);

  finally
    slShort.Free;
    slLong.Free;
  end;


  for i := 0 to XslOfficeDoc.Count -1 do begin
    sXML_Doc := XsExtractToThisPath + XslOfficeDoc.Strings[i];

    if not FileExists(sXML_Doc) then
        sErrorMessage := sErrorMessage + 'Err: Failing to extract text. File not found. filename: ' + sXML_Doc + #13#10

    else begin // File found
        sSpaces := '';
        sMyTxt := '';
        bTabFound := False;
        //vEnter := #$D#$A;


        XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
        XmlParser.StartScan;
        XmlParser.Normalize := FALSE;
        while XmlParser.Scan do begin
          if (XmlParser.CurPartType = ptStartTag) or (XmlParser.CurPartType = ptEmptyTag)	then begin
            { if XmlParser.CurName = 'w:jc' then begin
                  NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('w:val'));
                  if NN.Value = 'center' then
                      sSpaces := '               ';

                  NN.Free;
              end
              else if XmlParser.CurName = 'w:tab' then begin
                  sSpaces := '        ';
                  bTabFound := True;
              end }
              if XmlParser.CurName = 'a:p' then begin
                  if sMyTxt <> '' then begin // text found before this paragraph, print it first
                      if bTabFound then begin
                          memOutput.Lines.Append(''); // Payer wants adding one more line
                          bTabFound := False;
                      end;

                      memOutput.Lines.Append(sMyTxt);
                      sMyTxt := '';
                  end
                  else
                      memOutput.Lines.Append('');

              end;
          end;

          if (XmlParser.CurPartType = ptContent) or (XmlParser.CurPartType = ptCData) then begin
              if XmlParser.CurName = 'a:t' then begin
                  //if AnsiPos(vEnter, XmlParser.CurContent) = 0 then // not accept #$D#$A
                      sMyTxt := sSpaces + sMyTxt + XmlParser.CurContent;

              end;

              sSpaces := '';
          end;


        end;
    end;


    if sMyTxt <> '' then begin
        if bTabFound then
            memOutput.Lines.Append(''); // Payer wants adding one more line

        memOutput.Lines.Append(sMyTxt);
    end;

  end;

  memOutput.SetFocus;
  memOutput.SelStart := 0;
  memOutput.SelLength := 0;

  if sErrorMessage <> '' then
      PopUp_ErrorMessage(sErrorMessage);


end;


function TfrmMain.Xlsx_To_CSV(const sSheetName: string; const sPFilenameIs: string): boolean;
var
  f: TextFile;
  sMyTxt, sXML_Doc, sOutputPFilename, sFilename, sErrorMessage, sFile: string;

  NN: TNvpNode;
  bSrcIndexFound, bFoundSheetFile: boolean;
  vEnter, vSpace: variant;


  sRow, sCell, sC, sTemp, sTmp, sPart: string;
  iRowNo, iPreRow, iColNo, iPreCol, iT, iGap, iNum, iMaxCol: integer;
  iIndex: Int64;
  iLen, iP, iPos, i, j, k, iG: integer;

  // for progress bar
  sTotalRow: string;
  iPBar, iTotalRow: integer;
  //frmPBar: TfrmPBar;
begin
  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      ShowMessage('Err: Failing to extract data from xml file(s). Target directory not exist.' + #13#10);
      exit;
  end;

  Get_SharedStrings_From_Xlsx; // Get sharedStrings first

  if XslSheetsFilenames.Count = 0 then exit;

  sErrorMessage := '';


  bFoundSheetFile := False;
  for i := 0 to XslSheetsFilenames.Count -1 do begin
    sXML_Doc := XsExtractToThisPath + XslSheetsFilenames.Strings[i];
    sFile := Ansilowercase(ExtractFilename(sXML_Doc));
    if sFile = sSheetName then begin
        bFoundSheetFile := True;
        break;
    end;
  end;
  

  if not bFoundSheetFile then begin
      ShowMessage('Err. The active worksheet not found. The sheetname is: ' + sSheetName);
      exit;
  end;


  try
    Screen.Cursor := crHourGlass;
    //frmPBar := TfrmPBar.Create(nil);
    pnPB.Visible := True;

    iPBar := 0;

    // know length of progress bar
    if FileExists(sXML_Doc) then begin

        XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
        XmlParser.StartScan;
        XmlParser.Normalize := FALSE;
        while XmlParser.Scan do begin

          if (XmlParser.CurPartType = ptStartTag)	then begin
              if XmlParser.CurName = 'row' then begin

                  // Get Row number
                  NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('r'));
                  sTotalRow := NN.Value;
                  NN.Free;

              end;
          end;

        end;

    end;

    if TryStrtoInt(sTotalRow, iTotalRow) then
        iPBar := iPBar + iTotalRow;  // No of rows from all sheets

    sTotalRow := '';

    lbPB.Caption := 'Saving file...';
    GaugePB.MinValue := 0;
    GaugePB.MaxValue := iPBar;
    GaugePB.Progress := 0;
    //frmPBar.Show;



    FileMode := 2;  // 2 = read and write, 0 = read only
    AssignFile(f, sPFilenameIs);

    try
       Rewrite(f);  // open an old file and access read or write, error occurred when file doesn't exist
    except on E: Exception do
       begin
         sErrorMessage := sErrorMessage + 'Error occurred when trying to open '+
         sPFilenameIs + Chr(13) + Chr(10) + 'Error # '  + E.Message + #13#10;

         PopUp_ErrorMessage(sErrorMessage);
         exit;
       end;
    end;



    if not FileExists(sXML_Doc) then
        sErrorMessage := sErrorMessage + 'Err: Failing to extract text. File not found. filename: ' + sXML_Doc + #13#10

    // ******** remember to modify **************
    else begin // File found
        sMyTxt := '';
        bSrcIndexFound := False;
        vEnter := #$D#$A;
        vSpace := #$A0;
        iMaxCol := 0;
        iPreCol := 0;
        iPreRow := 1; // not include first fixed row, so, is 1 


        XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
        XmlParser.StartScan;
        XmlParser.Normalize := FALSE;
        while XmlParser.Scan do begin

          if (XmlParser.CurPartType = ptStartTag) or (XmlParser.CurPartType = ptEmptyTag)	then begin
              if XmlParser.CurName = 'row' then begin

                  if sMyTxt <> '' then begin // text found before this paragraph, print it first
                      //AMemo.Lines.Append(sMyTxt);
                      Writeln(f, sMyTxt);
                      sMyTxt := '';
                  end;

                  // Get Row number
                  NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('r'));
                  sRow := NN.Value;
                  NN.Free;

                  if TryStrToInt(sRow, iRowNo) then begin
                      iG := iRowNo - iPreRow -1;

                      //if iG > 1 then begin
                      for j := 0 to iG -1 do begin
                        Writeln(f, ',');
                      end;
                      //end;

                      iPreRow := iRowNo;

                      GaugePB.Progress := iRowNo;
                      GaugePB.Repaint;
                      Application.ProcessMessages;
                  end;

                  iPreCol := 0;

              end;


              if (XmlParser.CurPartType = ptEmptyTag) then begin
                  if (XmlParser.CurName = 'col') then
                      iMaxCol := iMaxCol + 1;

              end;


              if (XmlParser.CurPartType = ptStartTag) then begin
                  if (XmlParser.CurName = 'c') then begin

                      NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('t'));

                      if NN.Value = 's' then
                          bSrcIndexFound := True;

                      NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('r'));

                      try
                        sCell := NN.Value;
                        sCell := AnsiReplaceStr(sCell, sRow, ''); // remove Row no

                        iColNo := 0;
                        sC := sCell;


                        repeat
                          iLen := length(sC);
                          if iLen > 0 then
                              sTemp := copy(sC, 1, 1)
                          else
                              break;

                          iP := XslAlphabet.IndexOf(sTemp); // iP can tell what block is

                          k := iLen;
                          iNum := 1;

                          while k > 1 do begin
                            iNum := iNum * 26;
                            k := k -1;
                          end;

                          iT := iP * iNum; 

                          iColNo := iColNo + iT;
                          sC := copy(sC, 2, iLen -1);
                        until length(sC) = 0;


                        iGap := iColNo - iPreCol;


                        if iColNo > 0 then // prevent iP = -1, skip this error
                            iPreCol := iColNo;        // preserve Col no

                        if iGap > 1 then begin
                            repeat
                              sMyTxt := sMyTxt + ',';
                              iGap := iGap -1;
                            until iGap = 1;
                        end;

                      finally
                        NN.Free;
                      end;

                  end;
              end;
          end;


          if (XmlParser.CurPartType = ptEndTag) then begin
              if (XmlParser.CurName = 'c') then begin
                  bSrcIndexFound := False;
                  sMyTxt := sMyTxt + ',';
              end;
             { else if (XmlParser.CurName = 'row') then
                  iCnt := 0; }

          end;


          if (XmlParser.CurPartType = ptContent) or (XmlParser.CurPartType = ptCData) then begin
              if XmlParser.CurName = 'v' then begin
                  if bSrcIndexFound then begin
                      if TryStrToInt64(XmlParser.CurContent, iIndex) then begin
                          if XslSharedStr.Count > iIndex then begin

                                  sTmp := XslSharedStr.Strings[iIndex];

                                  iPos := AnsiPos('"', sTmp);
                                  if iPos > 0 then begin
                                      sPart := AnsiReplaceStr(sTmp, '"', '""');
                                      sTmp := sPart;

                                  end;


                                  if (AnsiPos(vEnter, sTmp) > 0) or
                                  (AnsiPos(',', sTmp) > 0) then begin
                                      if AnsiPos(vSpace, sTmp) > 0 then
                                          sTmp := AnsiReplaceStr(sTmp, vSpace, '');

                                      sTmp := '"' + sTmp + '"';
                                  end;

                                  sMyTxt := sMyTxt + sTmp;

                          end;
                      end;
                  end
                  else begin // is inStr?

                      sTmp := XmlParser.CurContent;

                      if (AnsiPos(vEnter, sTmp) > 0) or
                      (AnsiPos(',', sTmp) > 0) then begin

                          if AnsiPos(vSpace, sTmp) > 0 then
                              sTmp := AnsiReplaceStr(sTmp, vSpace, '');

                          sTmp := '"' + sTmp + '"';
                      end;

                      sMyTxt := sMyTxt + sTmp;

                  end;
              end;
          end;
        end;


        if sMyTxt <> '' then begin
            //AMemo.Lines.Append(sMyTxt);
            Writeln(f, sMyTxt);
        end;

    end;

  finally
    //frmPBar.Free;
    pnPB.Visible := False;
    Screen.Cursor := crDefault;
    CloseFile(f);
  end;
  

  //if AMemo.Text = '' then  // prevent rubblish
  //    sErrorMessage := sErrorMessage + 'No data found on sheet.' + #13#10;


  if sErrorMessage <> '' then
      PopUp_ErrorMessage(sErrorMessage)
  else
      Result := True;

  

end;


function TfrmMain.Xlsx_To_CSV_Speed(iSheetIndex: integer; bSaveAll: boolean): boolean;
var
  i, j, k, iPos, iFromPage, iToPage: integer;
  sMyTxt, sTmp, sPart, sSheet, sPath, sPathFilename: string;
  vEnter, vSpace: variant;
  VTr: TVirtualStringTree;
  Node: PVirtualNode;
  Data: PMyEntry;

  f: TextFile;
  SaveDlg: TSaveDialog;
begin
  iFromPage := 0;
  iToPage := PgeCtr.PageCount -1;

  if not bSaveAll then begin
      if iSheetIndex < PgeCtr.PageCount then begin
          iFromPage := iSheetIndex;
          iToPage := iSheetIndex;
      end
      else begin
          ShowMessage('Error, sheet index is out of bound');
          exit;
      end;
  end;

  sPath := '';

  for i := iFromPage to iToPage do begin
    for j := 0 to PgeCtr.Pages[i].ComponentCount -1 do begin
      if PgeCtr.Pages[i].Components[j] is TVirtualStringTree then begin
                  sSheet := PgeCtr.Pages[i].Caption;

                  SaveDlg := TSaveDialog.Create(nil);
                  try
                    SaveDlg.FileName := sSheet;
                    SaveDlg.DefaultExt := 'csv';
                    SaveDlg.Options := [ofOverwritePrompt,ofHideReadOnly,ofEnableSizing];
                    if sPath = '' then
                        sPath := ExtractFileDir(frmMain.ZipFName.Caption);

                    if DirectoryExists(sPath) then
                        SaveDlg.InitialDir := sPath;

                    SaveDlg.Filter := 'CSV files (*.csv)|*.csv|All files (*.*)|*.*';

                    if SaveDlg.Execute then begin
                        if SaveDlg.FileName <> '' then begin
                            sPathFilename := SaveDlg.FileName;
                            sPath := ExtractFileDir(SaveDlg.FileName); 
                        end;
                    end
                    else
                        break;

                  finally
                    SaveDlg.Free;
                  end;


                  VTr := PgeCtr.Pages[i].Components[j] as TVirtualStringTree;

                  sErrorMessage := '';
                  sMyTxt := '';
                  vEnter := #$D#$A;
                  vSpace := #$A0;

                  try
                    VTr.Cursor := crHandPoint;

                    FileMode := 2;  // 2 = read and write, 0 = read only
                    AssignFile(f, sPathFilename);

                    try
                      Rewrite(f);  // open an old file and access read or write, error occurred when file doesn't exist
                    except on E: Exception do
                      begin
                        sErrorMessage := sErrorMessage + 'Error occurred when trying to open '+
                        sPathFilename + Chr(13) + Chr(10) + 'Error # '  + E.Message + #13#10;

                        PopUp_ErrorMessage(sErrorMessage);
                        exit;
                      end;
                    end;


                    Node := VTr.GetFirst;

                    while Node <> nil do begin
                      Data := VTr.GetNodeData(Node);

                      // not include col 0
                      for k := 1 to VTr.Header.Columns.Count - 1 do begin
                        sTmp := Data.Value[k];

                        iPos := Pos('"', sTmp);
                        if iPos > 0 then begin
                            sPart := AnsiReplaceStr(sTmp, '"', '""');
                            sTmp := sPart;
                        end;


                        if (AnsiPos(vEnter, sTmp) > 0) or
                        (AnsiPos(',', sTmp) > 0) then begin
                            if Pos(vSpace, sTmp) > 0 then
                                sTmp := AnsiReplaceStr(sTmp, vSpace, '');

                            sTmp := '"' + sTmp + '"';
                        end;

                        sMyTxt := sMyTxt + sTmp + ',';
                      end;


                      Writeln(f, sMyTxt);

                      sMyTxt := '';
                      Node := VTr.GetNext(Node);
                    end;


                  finally
                    CloseFile(f);
                    VTr.Cursor := crDefault;
                  end;
                  
      end;
    end;
  end;
end;


procedure TfrmMain.Output_Txt_Contents_Of_XML_From_Xlsx;
var
  sMyTxt, sXML_Doc, sOutputPFilename, sFilename, sErrorMessage: string;
  sSheetname, sS, sExt, sRid: string;
  iL: integer;

  NN, ZZ: TNvpNode;
  bSrcIndexFound: boolean;
  vEnter, vSpace: variant;

  slShort, slLong: TStringList;

  sRow, sCell, sC, sTemp, sTmp, sPart, sCol: string;
  sDefaultRowH, sDefaultColW, sColMin, sColW, sRowH: string;
  iColNo, iPreCol, iRowNo, iPreRow, iT, iGap, iNum, iMaxCol, iMCol, iCnt: integer;
  iIndex, iR: int64;
  iLen, iP, iPos, i, j, k, m, iG, iRetFS, iColMin: integer;
  iLenHOfRow, iLLL_HofRow: integer;

  iPreviousColmin, iPreviousColminW: integer;
  ia, ib, ic: integer;
  iDD: Double;
  bFoundDrawing: boolean;

  // For images
  ix, iCol_is, iRow_is, iImgID, iImgWidth, iImgHeight: integer;
  iColWidth_is, iRowHeight_is: integer;

  // For Progress Bar
  sTotalRow: string;
  iPBar, iGo, iTotalRow, isss: integer;
  //frmPBar: TfrmPBar;

  TS: TTabSheet;
  SG: TStringGrid;
  VTr: TVirtualStringTree;
  Node, NextNode: PVirtualNode;
  Data: PMyEntry;
  NewCol: TCollectionItem;
begin
  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      ShowMessage('Err: Failing to extract data from xml file(s). Target directory not exist.' + #13#10);
      exit;
  end;

  Get_SheetNames_From_Xlsx;    // Get sheetnames
  Get_SharedStrings_From_Xlsx; // Get sharedStrings first

  if XslSheetsFilenames.Count = 0 then exit;

  sErrorMessage := '';
  memOutput.Clear;

  slShort := TStringList.Create;
  slLong := TStringList.Create;

  try // sorting

    for i := 0 to XslSheetsFilenames.Count -1 do begin
      sFilename := ExtractFilename(XslSheetsFilenames.Strings[i]);

      if length(sFilename) = 10 then
          slShort.Add(XslSheetsFilenames.Strings[i]) // e.g: slide1.xml
      else
          slLong.Add(XslSheetsFilenames.Strings[i]); // e.g: slide10.xml

    end;

    slShort.Sorted := True;
    slLong.Sorted := True;

    XslSheetsFilenames.Clear;

    XslSheetsFilenames.AddStrings(slShort);
    XslSheetsFilenames.AddStrings(slLong);

  finally
    slShort.Free;
    slLong.Free;
  end;


  try
    Screen.Cursor := crHourGlass;
    //frmPBar := TfrmPBar.Create(nil);
    pnPB.Visible := True;

  iPBar := 0;
  iGo := 0;

  // know length of progress bar
  for isss := 0 to XslSheetsFilenames.Count -1 do begin

    sXML_Doc := XsExtractToThisPath + XslSheetsFilenames.Strings[isss];

    if FileExists(sXML_Doc) then begin

        XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
        XmlParser.StartScan;
        XmlParser.Normalize := FALSE;
        while XmlParser.Scan do begin

          if (XmlParser.CurPartType = ptStartTag)	then begin
              if XmlParser.CurName = 'row' then begin

                  // Get Row number
                  NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('r'));
                  sTotalRow := NN.Value;
                  NN.Free;

              end;
          end;

        end;

    end;

    if TryStrtoInt(sTotalRow, iTotalRow) then
        iPBar := iPBar + iTotalRow;  // No of rows from all sheets

    sTotalRow := '';

  end;


  lbPB.Caption := 'Loading...';
  GaugePB.MinValue := 0;
  GaugePB.MaxValue := iPBar;
  GaugePB.Progress := 0;

  SetLength(ImagesOnSheet, 1);    // Reset
  ImagesOnSheet[0].iSheetID := -2;
  SetLength(AHeightOfRow, 1);    // Reset
  AHeightOfRow[0].iSheetID := -1;
  AHeightOfRow[0].iIndex := -1;
  SetLength(iStart_HOfRow, 1);   // Reset
  iStart_HOfRow[0] := 0;
  //frmPBar.Show;


  for i := 0 to XslSheetsFilenames.Count -1 do begin
    sXML_Doc := XsExtractToThisPath + XslSheetsFilenames.Strings[i];

    if not FileExists(sXML_Doc) then
        sErrorMessage := sErrorMessage + 'Err: Failing to extract text. File not found. filename: ' + sXML_Doc + #13#10

    // ********  **************
    else begin // File found

        if i = PgeCtr.PageCount then begin

            TS := TTabSheet.Create(PgeCtr);
            //with TTabSheet.Create(PgeCtr) do begin
            TS.PageControl := PgeCtr;

            //TS.Name := 'sheet' + inttostr(i +1);
            sS := Ansilowercase(ExtractFilename(sXML_Doc));
            sExt := ExtractFileExt(sS);
            iL := length(sExt);

            if iL > 0 then
                sS := copy(sS, 1, length(sS) -iL);

            TS.Name := sS;

            if XslRID.Count > i then begin
                sSheetname := XslRID.Strings[i];
                TS.Caption := sSheetname;
            end
            else
                TS.Caption := TS.Name;


           { SG := TStringGrid.Create(TS);
            SG.Parent := TS;
            SG.Align := alClient;
            SG.FixedCols := 1;
            SG.FixedRows := 1;
            SG.DefaultColWidth := 64;
            SG.DefaultRowHeight := 24;
            SG.Options := SG.Options + [Grids.goDrawFocusSelected]; //, Grids.goRowSizing, Grids.goColSizing];
            SG.Name := 'StringGrid' + inttostr(i +1);
            SG.PopupMenu := pmSG;
            SG.OnDrawCell := SGDrawCell;
            SG.OnMouseMove := SGMouseMove;
            SG.OnMouseDown := SGMouseDown; }

            VTr := TVirtualStringTree.Create(TS);
            VTr.Parent := TS;
            VTr.Align := alClient;
            VTr.Name := 'VTree' + inttostr(i+1);
            VTr.Visible := True;
            

            VTr.ScrollBarOptions.AlwaysVisible := True;
            VTr.ScrollBarOptions.ScrollBars := ssBoth;

            VTr.TreeOptions.AutoOptions := [toAutoDropExpand, toAutoScrollOnExpand,
                toAutoTristateTracking, toAutoDeleteMovedNodes];

            VTr.TreeOptions.PaintOptions := [toShowButtons,toShowDropmark,
                toThemeAware,toUseBlendedImages,
                toShowHorzGridLines, toShowVertGridLines];

            VTr.TreeOptions.SelectionOptions := [toDisableDrawSelection,
                toExtendedFocus];

            VTr.NodeDataSize := SizeOf(TMyEntry);
            VTr.Header.Assign(VTHide.Header);

            VTr.PopupMenu := pmSG;

            VTr.OnGetText := VTrGetText;
            VTr.OnBeforeCellPaint := VTrBeforeCellPaint;
            VTr.OnAfterCellPaint := VTrAfterCellPaint;
            VTr.OnFocusChanging := VTrFocusChanging;
            VTr.OnMeasureItem := VTrMeasureItem;
            VTr.OnMouseMove := VTrMouseMove;
            VTR.OnMouseDown := VTrMouseDown;

            VTr.BeginUpdate;
        end;

        
        sMyTxt := '';
        bSrcIndexFound := False;
        vEnter := #$D#$A;
        vSpace := #$A0;
        iMaxCol := 0;
        iMCol := 0;
        iRowNo := 0;
        iColNo := 0;
        iPreCol := 0;
        iPreRow := 0; // not include first fixed row, so, is 1
        iPreviousColmin := 0;
        Setlength(AThisColW, 1);                    // initial


        if i > 0 then begin
            iLLL_HofRow := Length(iStart_HOfRow);
            SetLength(iStart_HOfRow, iLLL_HofRow + 1);
            iStart_HOfRow[iLLL_HofRow] := 0;           // initial
        end;

        bFoundDrawing := False;


        XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
        XmlParser.StartScan;
        XmlParser.Normalize := FALSE;
        while XmlParser.Scan do begin

          if (XmlParser.CurPartType = ptStartTag) then begin
           if XmlParser.CurName = 'sheetData' then begin
            while XmlParser.Scan do begin
             if (XmlParser.CurPartType = ptStartTag) then begin
              // c found
              if (XmlParser.CurName = 'c') then begin
                      NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('t'));

                      if NN.Value = 's' then
                          bSrcIndexFound := True;

                      NN.Free;
                      // Cell position
                      NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('r'));

                      try
                        sCell := NN.Value;
                        sCell := AnsiReplaceStr(sCell, sRow, ''); // remove Row no

                        iColNo := 0;                        // initial from 0
                        sC := sCell;

                        // Get Col no
                        repeat
                          iLen := length(sC);
                          if iLen > 0 then                  // Get first char (A...
                              sTemp := copy(sC, 1, 1)
                          else
                              break;

                          // iP can tell what line of block is (AA or BA...
                          iP := XslAlphabet.IndexOf(sTemp);

                          k := iLen;
                          iNum := 1;

                          while k > 1 do begin
                            iNum := iNum * 26;   // first: A..Z, 2nd: AA..ZZ, 3rd: AAA...ZZZ
                            k := k -1;
                          end;

                          iT := iP * iNum;       // know pos on a block


                          iColNo := iColNo + iT;       // subtotal col to know exact pos
                          sC := copy(sC, 2, iLen -1);  // remove first char
                        until length(sC) = 0;

                        iGap := iColNo - iPreCol;      // jumping col found!

                        if iColNo > 0 then // prevent iP = -1, skip this error
                            iPreCol := iColNo;         // preserve Col no


                        //SetLength(Data.Value, iColNo + 1);

                        if iColNo > iMCol then begin
                            //if SG.ColCount <= iColNo then
                            //    SG.ColCount := iColNo +1;

                            iMCol := iColNo;

                            for j := 0 to iGap -1 do begin
                              NewCol := VTr.Header.Columns.Insert(VTr.Header.Columns.Count); // following after


                              VTr.Header.Columns[NewCol.Index].Text := '';
                              VTr.Header.Columns[NewCol.Index].Alignment := taRightJustify;

                              VTr.Header.Columns[NewCol.Index].CaptionAlignment := taCenter;
                              VTr.Header.Columns[NewCol.Index].MinWidth := 64;
                              if ( Length(AThisColW) > 1 ) and ( NewCol.Index <= High(AThisColW) ) then
                                  VTr.Header.Columns[NewCol.Index].Width := AThisColW[NewCol.Index]
                              else
                                  VTr.Header.Columns[NewCol.Index].Width := 64;

                              // must set this on last to make CaptionAlignment works
                              {VTr.Header.Columns[NewCol.Index].Options := [coEnabled,
                                  coParentBidiMode, coParentColor,
                                  coVisible, coUseCaptionAlignment];  }
                              VTr.Header.Columns[NewCol.Index].Options := VTr.Header.Columns[NewCol.Index].Options +
                                  [coUseCaptionAlignment];

                            end;

                        end;

                      finally
                        NN.Free;
                      end;

              end
              // row found
              else if XmlParser.CurName = 'row' then begin
                  if sMyTxt <> '' then begin // text found before this paragraph, print it first
                      //SG.Cells[iColNo, iRowNo] := sMyTxt;
                      Data.Value[iColNo] := sMyTxt;
                      sMyTxt := '';
                  end;

                  // Get Row number
                  NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('r'));
                  sRow := NN.Value;
                  NN.Free;

                  if TryStrToInt(sRow, iRowNo) then begin
                      //SG.RowCount := iRowNo +1;

                      iG := iRowNo - iPreRow -1;

                      //if iG > 1 then begin
                      for j := 0 to iG -1 do begin
                        Node := VTr.AddChild(nil);    // add root level node
                        Data := VTr.GetNodeData(Node);
                        Data.Value[0] := inttostr( iPreRow + (j+1) );
                      end;
                      //end;

                      iPreRow := iRowNo;
                  end;


                  Node := VTr.AddChild(nil);    // add root level node
                  Data := VTr.GetNodeData(Node);
                  Data.Value[0] := sRow;        // draw 0 Col, Row no
                  //SetLength(Data.Value, 1);  // new start from beginning

                  iPreCol := 0;
                  GaugePB.Progress := iGo + iRowNo;
                  GaugePB.Repaint;
                  Application.ProcessMessages;

                  // Row height
                  ZZ := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('ht'));
                  try
                    sRowH := ZZ.Value;

                    if TryStrToFloat(sRowH, iDD) then begin

                        iDD := iDD/0.6;
                        iRetFS := Round(iDD);

                        if iRetFS > 24 then begin
                            //Node.NodeHeight := iRetFS;
                            iLenHOfRow := length(AHeightOfRow);

                            AHeightOfRow[iLenHOfRow -1].iSheetID := i;
                            AHeightOfRow[iLenHOfRow -1].iIndex := Node.Index;
                            AHeightOfRow[iLenHOfRow -1].iHeight := iRetFS;

                            SetLength(AHeightOfRow, iLenHOfRow + 1);
                            AHeightOfRow[iLenHOfRow].iSheetID := -1;
                            AHeightOfRow[iLenHOfRow].iIndex := -1;

                        end;

                        //if iRetFS > SG.DefaultRowHeight then
                        //    SG.RowHeights[iRowNo] := iRetFS;

                    end;
                  finally
                    ZZ.Free;
                  end;

              end;
             end    // if ptStartTag end
             // content found
             else if (XmlParser.CurPartType = ptContent) or (XmlParser.CurPartType = ptCData) then begin
                 if XmlParser.CurName = 'v' then begin
                  if bSrcIndexFound then begin
                      if TryStrToInt64(XmlParser.CurContent, iIndex) then begin
                          if XslSharedStr.Count > iIndex then begin

                                  sTmp := XslSharedStr.Strings[iIndex];

                                  iPos := AnsiPos('"', sTmp);
                                  if iPos > 0 then begin // double quotes can show
                                      sPart := AnsiReplaceStr(sTmp, '"', '""');
                                      sTmp := sPart;

                                  end;

                                  // Comma can show
                                  if (AnsiPos(vEnter, sTmp) > 0) or
                                  (AnsiPos(',', sTmp) > 0) then begin
                                      if AnsiPos(vSpace, sTmp) > 0 then
                                          sTmp := AnsiReplaceStr(sTmp, vSpace, '');

                                      sTmp := '"' + sTmp + '"';
                                  end;

                                  sMyTxt := sMyTxt + sTmp;

                          end;
                      end;
                  end
                  else begin // is inStr?

                      sTmp := XmlParser.CurContent;

                      if (AnsiPos(vEnter, sTmp) > 0) or
                      (AnsiPos(',', sTmp) > 0) then begin

                          if AnsiPos(vSpace, sTmp) > 0 then
                              sTmp := AnsiReplaceStr(sTmp, vSpace, '');

                          sTmp := '"' + sTmp + '"';
                      end;

                      sMyTxt := sMyTxt + sTmp;

                  end;
                 end;
             end
             else if (XmlParser.CurPartType = ptEndTag) then begin
                 if (XmlParser.CurName = 'c') then begin
                     bSrcIndexFound := False;
                     //sMyTxt := sMyTxt + ',';

                     //SG.Cells[iColNo, iRowNo] := sMyTxt;
                     Data.Value[iColNo] := sMyTxt;
                     sMyTxt := '';

                 end
                 else if (XmlParser.CurName = 'sheetData') then
                     break;       // get all cells data and leave

                 { else if (XmlParser.CurName = 'row') then
                     iCnt := 0; }

             end;   // if ptEndTag end
            end;    // while end
           end;     // if sheetData end
          end       // if ptStartTag end
          else if (XmlParser.CurPartType = ptEmptyTag)	then begin

                  if (XmlParser.CurName = 'col') then begin
                      iMaxCol := iMaxCol + 1;

                      NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('width'));

                      try
                        sColW := NN.Value;

                        if TryStrToFloat(sColW, iDD) then begin
                          iRetFS := Round(iDD *10);  // about equal to Excel width

                          if iRetFS < VTr.Header.Columns.DefaultWidth then
                              iRetFS := VTr.Header.Columns.DefaultWidth;
                          //if iRetFS < SG.DefaultColWidth then
                          //    iRetFS := SG.DefaultColWidth;

                          ZZ := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('min'));
                          try
                            sColMin := ZZ.Value;
                            TryStrToInt(sColMin, iColMin);
                          finally
                            ZZ.Free;
                          end;


                          if (iColMin > 0) then begin
                              //SG.ColCount := iColMin +1;
                              for m := iPreviousColmin +1 to iColmin do begin
                                Setlength(AThisColW, length(AThisColW) +1);

                                if m = iColmin then
                                    AThisColW[m] := iRetFS
                                //    SG.ColWidths[m] := iRetFS
                                else // previous min, use previous width
                                    AThisColW[m] := iPreviousColminW;
                                //    SG.ColWidths[m] := iPreviousColminW;

                              end;

                              iPreviousColmin := iColmin;
                              iPreviousColminW := iRetFS;
                          end;

                        end;

                      finally
                        NN.Free;
                      end;

                  end
                  else if (XmlParser.CurName = 'sheetFormatPr') then begin
                      // Default row height
                      NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('defaultRowHeight'));

                      try
                        sDefaultRowH := NN.Value;

                        if TryStrToFloat(sDefaultRowH, iDD) then begin
                          iRetFS := Round(iDD);
                          //if iRetFS > 16 then  // mininum height
                          //    SG.DefaultRowHeight := iRetFS;

                        end;

                      finally
                        NN.Free;
                      end;

                      // Default col width
                      NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('defaultColWidth'));

                      try
                        sDefaultColW := NN.Value;

                        if TryStrToFloat(sDefaultColW, iDD) then begin
                          iRetFS := Round(iDD);
                          //if iRetFS >= 54 then  // minimum width
                          //    SG.DefaultColWidth := iRetFS;

                        end;

                      finally
                        NN.Free;
                      end;

                  end  // Drawing
                  else if (XmlParser.CurName = 'drawing') then begin
                      NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('r:id'));
                      try
                        if NN.Value <> '' then begin
                            sRid := NN.Value;
                            bFoundDrawing := True;
                        end;
                      finally
                        NN.Free;
                      end;
                  end;

          end;


        end;  // while end


        if sMyTxt <> '' then
            //SG.Cells[iColNo, iRowNo] := sMyTxt;
            Data.Value[iColNo] := sMyTxt;

        VTr.RootNodeCount := iRowNo + 1;
        VTr.EndUpdate;


        if bFoundDrawing then begin// Drawing found, need to preserve information of images
            // Knowing drawing filename from sheet filename
            Get_Drawing_Filename( ExtractFilename(sXML_Doc), i, sRid );

            // Above got information of images, one more step to preserve
            // ToCol and ToRow
            for ix := 0 to High(ImagesOnSheet) do begin
              if ImagesOnSheet[ix].iSheetID = i then begin // this sheet no
                  iImgID := ImagesOnSheet[ix].iImageID;
                  if iImgID <> -1 then begin

                      iImgWidth := AImage[iImgID].Picture.Width;
                      iImgHeight := AImage[iImgID].Picture.Height;

                      iCol_is := ImagesOnSheet[ix].iCol;
                      ImagesOnSheet[ix].iToCol := ImagesOnSheet[ix].iCol; // initialize

                      if VTr.Header.Columns.Count > iCol_is then begin
                          iColWidth_is := VTr.Header.Columns[iCol_is].Width;
                      //if SG.ColCount > iCol_is then begin
                      //    iColWidth_is := SG.ColWidths[iCol_is];

                          while iImgWidth > iColWidth_is do begin
                            if iCol_is < (VTr.Header.Columns.Count -1) then begin
                            //if iCol_is < (SG.ColCount -1) then begin
                                ImagesOnSheet[ix].iToCol := iCol_is + 1;

                                iCol_is := iCol_is + 1; // next col
                                // Accumulation of col width
                                iColWidth_is := iColWidth_is + VTr.Header.Columns[iCol_is].Width +
                                                1;
                                //iColWidth_is := iColWidth_is + SG.ColWidths[iCol_is] +
                                //                SG.GridLineWidth;

                            end
                            else
                                break;

                          end;
                      end;


                      iRow_is := ImagesOnSheet[ix].iRow;
                      ImagesOnSheet[ix].iToRow := ImagesOnSheet[ix].iRow; // initialize

                      //iNodeCnt := 0;
                      if VTr.RootNodeCount > iRow_is then begin
                          NextNode := VTr.GetFirst;
                          if NextNode <> nil then begin
                              repeat
                                //Inc(iNodeCnt);
                                if NextNode.Index = iRow_is then begin
                                    iRowHeight_is := NextNode.NodeHeight;
                                    break;
                                end;
                                NextNode := VTr.GetNext(NextNode);
                              until NextNode = nil;
                          end;

                          //iRowHeight_is := SG.RowHeights[iRow_is];

                      //if SG.RowCount > iRow_is then begin
                      //    iRowHeight_is := SG.RowHeights[iRow_is];

                          while iImgHeight > iRowHeight_is do begin
                            if iRow_is < (VTr.RootNodeCount -1) then begin
                            //if iRow_is < (SG.RowCount -1) then begin
                                ImagesOnSheet[ix].iToRow := iRow_is + 1;

                                iRow_is := iRow_is + 1; // next row
                                // Accumulation of row height

                                //iNodeCnt := 0;
                                NextNode := VTr.GetFirst;
                                if NextNode <> nil then begin
                                    repeat
                                      //Inc(iNodeCnt);
                                      if NextNode.Index = iRow_is then begin
                                          iRowHeight_is := iRowHeight_is +
                                                           NextNode.NodeHeight +
                                                           1;
                                          break;
                                      end;
                                      NextNode := VTr.GetNext(NextNode);
                                    until NextNode = nil;
                                end;


                                //iRowHeight_is := iRowHeight_is + SG.RowHeights[iRow_is] +
                                //                SG.GridLineWidth;

                            end
                            else
                                break;

                          end;
                      end;

                  end;

              end;
            end;

        end;

    end;


    iGo := iRowNo;  // for progress bar


    iCnt := 1;

    while iCnt <= iMCol do begin    // draw Col header
      if iCnt <= 26 then begin
          sCol := XslAlphabet.Strings[iCnt];

          VTr.Header.Columns[iCnt].Text := sCol;
          //SG.Cells[iCnt, 0] := sCol;
          iCnt := iCnt +1;

          //if iCnt > iMCol then break;
      end
      else if iCnt <= 702 then begin  // 26*26+26  (AA)
          for ia := 1 to 26 do begin
            for ib := 1 to 26 do begin
              sCol := XslAlphabet.Strings[ia] + XslAlphabet.Strings[ib];

              VTr.Header.Columns[iCnt].Text := sCol;
              //SG.Cells[iCnt, 0] := sCol;
              iCnt := iCnt +1;

              if iCnt > iMCol then break;
            end;

            if iCnt > iMCol then break;
          end;

          if iCnt > iMCol then break;
      end
      else if iCnt <= 18278 then begin  // 26*26*26+702  (AAA)
          for ia := 1 to 26 do begin
            for ib := 1 to 26 do begin
              for ic := 1 to 26 do begin
                sCol := XslAlphabet.Strings[ia] + XslAlphabet.Strings[ib] +
                        XslAlphabet.Strings[ic];

                VTr.Header.Columns[iCnt].Text := sCol;
                //SG.Cells[iCnt, 0] := sCol;
                iCnt := iCnt +1;

                if iCnt > iMCol then break;
              end;

              if iCnt > iMCol then break;
            end;

            if iCnt > iMCol then break;
          end;

          if iCnt > iMCol then break;
      end;

    end;


   { iR := 1;
    while iR < SG.RowCount do begin // draw Row number
      SG.Cells[0, iR] := inttostr(iR);
      iR := iR +1;
    end;
    }

  end;

  finally
    //frmPBar.Free;
    pnPB.Visible := False;
    Screen.Cursor := crDefault;
  end;

  if sErrorMessage <> '' then
      PopUp_ErrorMessage(sErrorMessage);


end;


procedure TfrmMain.Output_Txt_Contents_Of_XML_From_Xlsx_Speed;
const
  cDrawing = 'drawing';
var
  sMyTxt, sXML_Doc, sOutputPFilename, sFilename, sErrorMessage: string;
  sSheetname, sS, sExt, sRid: string;

  //vEnter, vSpace: variant;

  slShort, slLong: TStringList;

  sRow, sC, sTemp, sCol: string;
  sColMin, sColW, sRowH: string;
  iColNo, iRowNo, iT, iNum, iCnt, iL: integer;
  iIndex: int64;
  iLen, iP, iRetFS, iColMin: integer;
  iLenHOfRow, iLLL_HofRow: integer;

  iDD: Double;

  // For images
  ix, iCol_is, iRow_is, iImgID, iImgWidth, iImgHeight: integer;
  iColWidth_is, iRowHeight_is: integer;

  // For Progress Bar
  iPBar, iGo, iTotalRow, isss: integer;

  TS: TTabSheet;
  SG: TStringGrid;
  VTr: TVirtualStringTree;
  Node, NextNode: PVirtualNode;
  NN: TNvpNode;
  Data: PMyEntry;
  NewCol: TCollectionItem;

  i, j, k, m: integer;
  ia, ib, ic: integer;
  iPreviousColmin, iPreviousColminW: integer;

  sDimension, sOneChr, sColCap, sR: string;
  sHeaderCol, sRowNo: array of string;
  iDPos, iFroPos, iColCount: integer;
  bAssignedDataToCell, bFoundDrawing, bSrcIndexFound: boolean;

  f: FILE;
  ReadIn, OldFileMode, iTagPos: integer;
  sCellLoc, sSourceIdx, sValue: string;
begin
  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      ShowMessage('Err: Failing to extract data from xml file(s). Target directory not exist.' + #13#10);
      exit;
  end;

  Get_SheetNames_From_Xlsx;    // Get sheetnames
  Get_SharedStrings_From_Xlsx; // Get sharedStrings first

  if XslSheetsFilenames.Count = 0 then exit;

  sErrorMessage := '';
  memOutput.Clear;

  slShort := TStringList.Create;
  slLong := TStringList.Create;

  try // sorting

    for i := 0 to XslSheetsFilenames.Count -1 do begin
      sFilename := ExtractFilename(XslSheetsFilenames.Strings[i]);

      if length(sFilename) = 10 then
          slShort.Add(XslSheetsFilenames.Strings[i]) // e.g: slide1.xml
      else
          slLong.Add(XslSheetsFilenames.Strings[i]); // e.g: slide10.xml

    end;

    slShort.Sorted := True;
    slLong.Sorted := True;

    XslSheetsFilenames.Clear;

    XslSheetsFilenames.AddStrings(slShort);
    XslSheetsFilenames.AddStrings(slLong);

  finally
    slShort.Free;
    slLong.Free;
  end;


  try
    Screen.Cursor := crHourGlass;
    pnPB.Visible := True;

  iPBar := 0;
  iGo := 0;

  // know length of progress bar
  for isss := 0 to XslSheetsFilenames.Count -1 do begin

    sXML_Doc := XsExtractToThisPath + XslSheetsFilenames.Strings[isss];

    if FileExists(sXML_Doc) then begin

        Setlength(sHeaderCol, length(sHeaderCol) + 1);
        Setlength(sRowNo, length(sRowNo) + 1);


        XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
        XmlParser.StartScan;
        XmlParser.Normalize := FALSE;
        while XmlParser.Scan do begin

          if (XmlParser.CurPartType = ptEmptyTag)	then begin
              if XmlParser.CurName = 'dimension' then begin

                  // Get dimension of sheet
                  NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('ref'));
                  sDimension := NN.Value;
                  NN.Free;

                  iDPos := Pos(':', sDimension);
                  if iDPos > 0 then                 // get last dim (ex: AB150 )
                      sDimension := copy(sDimension, iDPos + 1,
                                    Length(sDimension) - iDPos);

                  if sDimension <> '' then begin
                      // first must be a Letter
                      sHeaderCol[isss] := copy(sDimension, 1, 1);
                      sDimension := copy(sDimension, 2, length(sDimension) -1);
                      //id := 1;
                      repeat
                        sOneChr := copy(sDimension, 1, 1);
                        if sOneChr <> '' then begin
                            if (XslAlphabet.IndexOf(sOneChr) <> -1 ) then begin
                                sHeaderCol[isss] := sHeaderCol[isss] + sOneChr;
                                sDimension := copy(sDimension, 2, length(sDimension) -1);
                            end
                            else begin
                                sRowNo[isss] := sDimension;  // remain is a number
                                break;
                            end;
                        end;
                        //Inc(id);          // move to next char
                      until sOneChr = '';
                  end;

                  break;                  // not search for others
              end;      // if CurName
          end;          // if ptEmptyTag

        end;            // while wend
    end;


    if TryStrtoInt(sRowNo[isss], iTotalRow) then
        iPBar := iPBar + iTotalRow;  // No of rows from all sheets

  end;


  lbPB.Caption := 'Loading...';
  GaugePB.MinValue := 0;
  GaugePB.MaxValue := iPBar;
  GaugePB.Progress := 0;

  SetLength(ImagesOnSheet, 1);    // Reset
  ImagesOnSheet[0].iSheetID := -2;
  SetLength(AHeightOfRow, 1);    // Reset
  AHeightOfRow[0].iSheetID := -1;
  AHeightOfRow[0].iIndex := -1;
  SetLength(iStart_HOfRow, 1);   // Reset
  iStart_HOfRow[0] := 0;
  

  // load data from sheets
  for i := 0 to XslSheetsFilenames.Count -1 do begin
    sXML_Doc := XsExtractToThisPath + XslSheetsFilenames.Strings[i];

    if not FileExists(sXML_Doc) then
        sErrorMessage := sErrorMessage + 'Err: Failing to extract text. File not found. filename: ' + sXML_Doc + #13#10

    // ********  **************
    else begin // File found

        if i = PgeCtr.PageCount then begin

            TS := TTabSheet.Create(PgeCtr);
            //with TTabSheet.Create(PgeCtr) do begin
            TS.PageControl := PgeCtr;

            //TS.Name := 'sheet' + inttostr(i +1);
            sS := Ansilowercase(ExtractFilename(sXML_Doc));
            sExt := ExtractFileExt(sS);
            iL := length(sExt);

            if iL > 0 then
                sS := copy(sS, 1, length(sS) -iL);

            TS.Name := sS;

            if XslRID.Count > i then begin
                sSheetname := XslRID.Strings[i];
                TS.Caption := sSheetname;
            end
            else
                TS.Caption := TS.Name;


            VTr := TVirtualStringTree.Create(TS);
            VTr.Parent := TS;
            VTr.Align := alClient;
            VTr.Name := 'VTree' + inttostr(i+1);
            VTr.Visible := True;
            

            VTr.ScrollBarOptions.AlwaysVisible := True;
            VTr.ScrollBarOptions.ScrollBars := ssBoth;

            VTr.TreeOptions.AutoOptions := [toAutoDropExpand, toAutoScrollOnExpand,
                toAutoTristateTracking, toAutoDeleteMovedNodes];

            VTr.TreeOptions.MiscOptions := [toGridExtensions];

            VTr.TreeOptions.PaintOptions := [toShowButtons,toShowDropmark,
                toThemeAware,toUseBlendedImages,
                toShowHorzGridLines, toShowVertGridLines];

            VTr.TreeOptions.SelectionOptions := [toDisableDrawSelection,
                toExtendedFocus];

            VTr.NodeDataSize := SizeOf(TMyEntry);
            VTr.Header.Assign(VTHide.Header);

            VTr.PopupMenu := pmSG;

            VTr.OnGetText := VTrGetText;
            VTr.OnBeforeCellPaint := VTrBeforeCellPaint;
            VTr.OnAfterCellPaint := VTrAfterCellPaint;
            VTr.OnFocusChanging := VTrFocusChanging;
            VTr.OnMeasureItem := VTrMeasureItem;
            VTr.OnMouseMove := VTrMouseMove;
            VTR.OnMouseDown := VTrMouseDown;

            VTr.BeginUpdate;
        end;


        if sHeaderCol[i] <> '' then begin    // add col

            iColNo := 1;              // Col 0 was created, initial from 1
            sC := sHeaderCol[i];

            // Get Col no
            repeat
              iLen := length(sC);
              if iLen > 0 then                  // Get first char (A...
                  sTemp := copy(sC, 1, 1)
              else
                  break;

              // iP can tell what line of block is (AA or BA...
              iP := XslAlphabet.IndexOf(sTemp);

              k := iLen;
              iNum := 1;

              while k > 1 do begin
                iNum := iNum * 26;   // first: A..Z, 2nd: AA..ZZ, 3rd: AAA...ZZZ
                k := k -1;
              end;

              iT := iP * iNum;       // know pos on a block


              iColNo := iColNo + iT;       // subtotal col to know exact pos
              sC := copy(sC, 2, iLen -1);  // remove first char
            until length(sC) = 0;


            for j := 0 to iColNo -1 do begin
              NewCol := VTr.Header.Columns.Insert(VTr.Header.Columns.Count); // following after


              VTr.Header.Columns[NewCol.Index].Text := '';
              VTr.Header.Columns[NewCol.Index].Alignment := taRightJustify;

              VTr.Header.Columns[NewCol.Index].CaptionAlignment := taCenter;
              VTr.Header.Columns[NewCol.Index].MinWidth := 64;
              if ( Length(AThisColW) > 1 ) and ( NewCol.Index <= High(AThisColW) ) then
                  VTr.Header.Columns[NewCol.Index].Width := AThisColW[NewCol.Index]
              else
                  VTr.Header.Columns[NewCol.Index].Width := 64;

              // must set this on last to make CaptionAlignment works
              {VTr.Header.Columns[NewCol.Index].Options := [coEnabled,
                                  coParentBidiMode, coParentColor,
                                  coVisible, coUseCaptionAlignment];  }
              VTr.Header.Columns[NewCol.Index].Options := VTr.Header.Columns[NewCol.Index].Options +
              [coUseCaptionAlignment];

            end;

        end;


        iColCount := VTr.Header.Columns.Count;

        iCnt := 1;

        while iCnt <= iColNo do begin    // draw Col header
              if iCnt <= 26 then begin
                  sCol := XslAlphabet.Strings[iCnt];

                  VTr.Header.Columns[iCnt].Text := sCol;
                  iCnt := iCnt +1;

              end
              else if iCnt <= 702 then begin  // 26*26+26  (AA)
                  for ia := 1 to 26 do begin
                    for ib := 1 to 26 do begin
                      sCol := XslAlphabet.Strings[ia] + XslAlphabet.Strings[ib];

                      VTr.Header.Columns[iCnt].Text := sCol;
                      iCnt := iCnt +1;

                      if iCnt > iColNo then break;
                    end;

                    if iCnt > iColNo then break;
                  end;

                  if iCnt > iColNo then break;
              end
              else if iCnt <= 18278 then begin  // 26*26*26+702  (AAA)
                  for ia := 1 to 26 do begin
                    for ib := 1 to 26 do begin
                      for ic := 1 to 26 do begin
                        sCol := XslAlphabet.Strings[ia] + XslAlphabet.Strings[ib] +
                                XslAlphabet.Strings[ic];

                        VTr.Header.Columns[iCnt].Text := sCol;
                        iCnt := iCnt +1;

                        if iCnt > iColNo then break;
                      end;

                      if iCnt > iColNo then break;
                    end;

                    if iCnt > iColNo then break;
                  end;

                  if iCnt > iColNo then break;
              end;
        end;              // while end





        if sRowNo[i] <> '' then begin       // draw loc to each cell

            iRowNo := strtoint(sRowNo[i]);
            for j := 0 to iRowNo - 1 do begin
              Node := VTr.AddChild(nil);    // add root level node
              Data := VTr.GetNodeData(Node);
              Data.Value[0] := inttostr(j + 1);        // draw Row no to Col 0

              for m := 1 to iColNo -1 do begin
                sColCap := VTr.Header.Columns[m].Text;
                sR := inttostr(j + 1);
                Data.Cell_Loc[m] := sColCap + sR;
              end;
            end;

        end;



        iFroPos := 1;
        //sRCell := '';
        sMyTxt := '';

        bSrcIndexFound := False;
        //vEnter := #$D#$A;
        //vSpace := #$A0;
        //iMaxCol := 0;
        //iMCol := 0;
        iRowNo := 0;
        iColNo := 0;
        //iPreCol := 0;
        //iPreRow := 0; // not include first fixed row, so, is 1
        iPreviousColmin := 0;
        Setlength(AThisColW, 1);                    // initial
        bFoundDrawing := False;

        Node := VTr.GetFirst;

        if i > 0 then begin
            iLLL_HofRow := Length(iStart_HOfRow);
            SetLength(iStart_HOfRow, iLLL_HofRow + 1);
            iStart_HOfRow[iLLL_HofRow] := 0;           // initial
        end;




        // --- Open File
        OldFileMode := SYSTEM.FileMode;
        try
          SYSTEM.FileMode := fmOpenRead or fmShareDenyNone;
          try
            AssignFile (f, sXML_Doc);
            Reset (f, 1);
          except
            sErrorMessage := sErrorMessage + 'Err: Failing to open file: ' + sXML_Doc + #13#10;
            break;
          end;

          try
            // --- Allocate Memory
            try
              FBufferSize := Filesize (f) + 1;
              GetMem (FBuffer, FBufferSize);
            except
              Release_Memory_Of_Xlsx_File_Open;
              sErrorMessage := sErrorMessage + 'Err: Failing to allocate memory for file: ' + sXML_Doc + #13#10;
              break;
            end;

            // --- Read File
            try
              BlockRead (f, FBuffer^, FBufferSize, ReadIn);
              (FBuffer+ReadIn)^ := #0;  // NULL termination
            except
              Release_Memory_Of_Xlsx_File_Open;
              sErrorMessage := sErrorMessage + 'Err: Failing to read file into memory: ' + sXML_Doc + #13#10;
              break;
            end;

          finally
            CloseFile (f);
          end;

          //FSource := Filename;
          //Result  := TRUE;
          CurStart := FBuffer;              // pointer to CurStart

        finally
          SYSTEM.FileMode := OldFileMode;
        end;

        // Is xml file?
        if not ( StrLComp(CurStart, '<?xml', 5) = 0 ) then begin
            Release_Memory_Of_Xlsx_File_Open;
            sErrorMessage := sErrorMessage + 'Err: Not xml file found: ' + sXML_Doc + #13#10;
            break;
        end;


        repeat                                            // do col width

          if  CurStart^ = '<' then begin                  // find first <
               inc(CurStart);

               if CurStart^ = 'c' then begin              // find col!
                   inc(CurStart);

                   if CurStart^ = 'o' then begin
                       inc(CurStart);

                       if CurStart^ = 'l' then begin
                           inc(CurStart);

                           if CurStart^ = ' ' then begin
                               inc(CurStart);

                               repeat

                                 if CurStart^ = 'm' then begin
                                     inc(CurStart);

                                     if CurStart^ = 'i' then begin
                                         inc(CurStart);

                                         if CurStart^ = 'n' then begin
                                             inc(CurStart);

                                             if CurStart^ = '=' then begin
                                                 inc(CurStart);

                                                 if CurStart^ = '"' then begin
                                                     inc(CurStart);
                                                     sColMin := '';

                                                     repeat
                                                       if CurStart^ <> '"' then
                                                           // get width of eache col
                                                           sColMin := sColMin + CurStart^;

                                                       inc(CurStart);
                                                     until CurStart^ = '"';

                                                     if TryStrToInt(sColMin, iColMin) then begin
                                                        { if (iColMin > 0) then begin

                                                             for m := iPreviousColmin +1 to iColmin do begin
                                                               Setlength(AThisColW, length(AThisColW) +1);

                                                               if m = iColmin then
                                                                   AThisColW[m] := iRetFS
                                                               else // previous min, use previous width
                                                                   AThisColW[m] := iPreviousColminW;

                                                             end;

                                                             iPreviousColmin := iColmin;
                                                             iPreviousColminW := iRetFS;

                                                         end; }
                                                     end;

                                                 end;
                                             end;
                                         end;
                                     end;
                                 end
                                 else if CurStart^ = 'w' then begin
                                     inc(CurStart);

                                     if CurStart^ = 'i' then begin
                                         inc(CurStart);

                                         if CurStart^ = 'd' then begin
                                             inc(CurStart);

                                             if CurStart^ = 't' then begin
                                                 inc(CurStart);

                                                 if CurStart^ = 'h' then begin
                                                     inc(CurStart);

                                                     if CurStart^ = '=' then begin
                                                         inc(CurStart);

                                                         if CurStart^ = '"' then begin
                                                             inc(CurStart);
                                                             sColW := '';

                                                             repeat
                                                               if CurStart^ <> '"' then
                                                                   // get width of eache col
                                                                   sColW := sColW + CurStart^;

                                                               inc(CurStart);
                                                             until CurStart^ = '"';


                                                             if TryStrToFloat(sColW, iDD) then begin
                                                                 iRetFS := Round(iDD *10);  // about equal to Excel width

                                                                  if iRetFS < 20 then
                                                                      iRetFS := VTr.Header.Columns.DefaultWidth;

                                                                  if (iColMin > 0) then begin

                                                                      for m := iPreviousColmin +1 to iColmin do begin

                                                                        if m < VTr.Header.Columns.Count then begin
                                                                            Setlength(AThisColW, length(AThisColW) +1);

                                                                            if m = iColmin then
                                                                                //AThisColW[m] := iRetFS
                                                                                VTr.Header.Columns[m].Width := iRetFS
                                                                            else                     // previous min, use previous width
                                                                                //AThisColW[m] := iPreviousColminW;
                                                                                VTr.Header.Columns[m].Width := iPreviousColminW;

                                                                        end;

                                                                      end;

                                                                      iPreviousColmin := iColmin;
                                                                      iPreviousColminW := iRetFS;

                                                                  end;

                                                                  break;

                                                             end;
                                                         end;

                                                     end;
                                                 end;
                                             end;
                                         end;
                                     end;
                                 end;

                                 inc(CurStart);
                               until CurStart^ = '>';
                           end;

                       end;

                   end;

               end
               else if CurStart^ = '/' then begin              // is end tag col!
                   inc(CurStart);

                   if CurStart^ = 'c' then begin
                       inc(CurStart);

                       if CurStart^ = 'o' then begin
                           inc(CurStart);

                           if CurStart^ = 'l' then begin
                               inc(CurStart);

                               if CurStart^ = 's' then begin
                                   inc(CurStart);

                                   if CurStart^ = '>' then begin
                                       inc(CurStart);

                                       break;
                                   end;
                               end;
                           end;
                       end;
                   end;
               end;
          end;

          inc(CurStart);                        // not found, move to next
        until CurStart^ = #0;

        CurStart := FBuffer;              // start from beggining
        iTagPos := Pos('<sheetData>', CurStart);
        CurStart := CurStart + iTagPos;   // move to pos sheetData after <



        repeat                                            // do cell data
          
          if  CurStart^ = '<' then begin                  // find first <
            inc(CurStart);

            if CurStart^ = 'c' then begin                 // find col!
                inc(CurStart);

              if CurStart^ = ' ' then begin
                inc(CurStart);

                bSrcIndexFound := False;
                sCellLoc := '';
                sSourceIdx := '';

                repeat
                  if CurStart^ = 'r' then begin           // attr r (cell loc)
                      inc(CurStart);

                      if CurStart^ = '=' then begin
                          inc(CurStart);

                          if CurStart^ = '"' then begin
                              inc(CurStart);

                              repeat
                                if CurStart^ <> '"' then
                                    // get cell location such as: AC101
                                    sCellLoc := sCellLoc + CurStart^;

                                inc(CurStart);
                              until CurStart^ = '"';
                          end;
                      end;
                  end
                  else if CurStart^ = 't' then begin      // attr t (source id)
                      inc(CurStart);

                      if CurStart^ = '=' then begin
                          inc(CurStart);

                          if CurStart^ = '"' then begin
                              inc(CurStart);

                              repeat
                                if CurStart^ <> '"' then
                                    sSourceIdx := sSourceIdx + CurStart^;

                                inc(CurStart);
                              until CurStart^ = '"';

                              if sSourceIdx = 's' then
                                  bSrcIndexFound := True;

                          end;
                      end;

                  end;

                  inc(CurStart);

                until CurStart^ = '>';

                inc(CurStart);
              end;

            end
            else if CurStart^ = 'v' then begin            // find data!
                inc(CurStart);

                sValue := '';

                if CurStart^ = '>' then begin
                    inc(CurStart);

                    repeat
                      sValue := sValue + CurStart^;
                      inc(CurStart);
                    until CurStart^ = '<';

                    inc(CurStart);
                end;


                if bSrcIndexFound then begin              // is source index
                    if TryStrToInt64(sValue, iIndex) then begin
                        if XslSharedStr.Count > iIndex then begin

                            sMyTxt := XslSharedStr.Strings[iIndex];

                        end;
                    end;
                end
                else
                    sMyTxt := sValue;                     // get data directly


                bAssignedDataToCell := False;
                repeat
                  if Node.Index = iRowNo - 1 then begin   // look for row
                      Data := VTr.GetNodeData(Node);
                      for k := iFroPos to iColCount -1 do begin // look for col
                        if Data.Cell_Loc[k] = sCellLoc then begin
                            Data.Value[k] := sMyTxt;            // match!
                            iFroPos := k + 1;    // next time start from this pos
                            bAssignedDataToCell := True;
                            break;
                        end;
                      end;
                      break;                    // maybe, not found!
                  end;

                  if bAssignedDataToCell then break;

                  if Node.Index < iRowNo - 1 then   // prevent over exceed row
                      Node := VTr.GetNext(Node)
                  else
                      break;

                until Node = nil;

                sMyTxt := '';                       // initial

            end
            else if CurStart^ = 'r' then begin            // find row!
                inc(CurStart);

                if CurStart^ = 'o' then begin
                    inc(CurStart);

                    if CurStart^ = 'w' then begin
                        inc(CurStart);

                      if CurStart^ = ' ' then begin
                        inc(CurStart);

                        repeat

                          if CurStart^ = 'r' then begin   // get row no
                              inc(CurStart);

                              if CurStart^ = '=' then begin
                                  inc(CurStart);

                                  if CurStart^ = '"' then begin
                                      inc(CurStart);

                                      sRow := '';

                                      repeat
                                        if CurStart^ <> '"' then
                                            sRow := sRow + CurStart^;

                                        inc(CurStart);

                                      until CurStart^ = '"';

                                      iFroPos := 1;       // initial (For search col no)

                                      // preserve row no and refresh PBar
                                      if TryStrToInt(sRow, iRowNo) then begin
                                          GaugePB.Progress := iGo + iRowNo;
                                          GaugePB.Repaint;
                                          Application.ProcessMessages;
                                      end;

                                  end;
                              end;
                          end
                          else if CurStart^ = 'h' then begin   // get row height
                              inc(CurStart);

                              if CurStart^ = 't' then begin
                                  inc(CurStart);

                                  if CurStart^ = '=' then begin
                                      inc(CurStart);

                                      if CurStart^ = '"' then begin
                                          inc(CurStart);
                                          sRowH := '';

                                          repeat
                                            if CurStart^ <> '"' then
                                                sRowH := sRowH + CurStart^;

                                                inc(CurStart);

                                          until CurStart^ = '"';

                                          if TryStrToFloat(sRowH, iDD) then begin

                                              iDD := iDD/0.6;
                                              iRetFS := Round(iDD);

                                              if iRetFS > 24 then begin
                                                  //Node.NodeHeight := iRetFS;
                                                  iLenHOfRow := length(AHeightOfRow);

                                                  AHeightOfRow[iLenHOfRow -1].iSheetID := i;
                                                  AHeightOfRow[iLenHOfRow -1].iIndex := iRowNo -1;
                                                  AHeightOfRow[iLenHOfRow -1].iHeight := iRetFS;

                                                  SetLength(AHeightOfRow, iLenHOfRow + 1);
                                                  AHeightOfRow[iLenHOfRow].iSheetID := -1;
                                                  AHeightOfRow[iLenHOfRow].iIndex := -1;

                                              end;
                                          end;
                                      end;

                                  end;
                              end;
                          end;

                          inc(CurStart);
                        until CurStart^ = '>';

                      end;

                    end;
                end;

            end
            else if CurStart^ = 'd' then begin            // find drawing!
                inc(CurStart);
                bFoundDrawing := False;
                sRid := '';

                for k := 2 to Length(cDrawing) do begin   // check name: drawing
                  if CurStart^ = copy(cDrawing, k, 1) then begin
                      if CurStart^ = 'g' then begin       // checked: is drawing
                          inc(CurStart);

                          if CurStart^ = ' ' then begin
                              inc(CurStart);

                              repeat
                                if CurStart^ = 'r' then begin
                                    inc(CurStart);

                                    if CurStart^ = ':' then begin
                                        inc(CurStart);

                                        if CurStart^ = 'i' then begin
                                            inc(CurStart);

                                            if CurStart^ = 'd' then begin
                                                inc(CurStart);

                                                if CurStart^ = '=' then begin
                                                    inc(CurStart);

                                                    if CurStart^ = '"' then begin
                                                        inc(CurStart);

                                                        // get rid no
                                                        repeat
                                                          if CurStart^ <> '"' then
                                                              sRid := sRid + CurStart^;

                                                          inc(CurStart);
                                                        until CurStart^ = '"';

                                                        bFoundDrawing := True;
                                                        inc(CurStart);
                                                        break;
                                                    end;
                                                end;
                                            end;
                                        end;
                                    end;
                                end;

                                inc(CurStart);
                              until CurStart^ = '>';

                              break;
                          end;

                      end;

                      inc(CurStart);
                  end
                  else
                      break;
                      
                end;
            end
            else
                inc(CurStart);      // found no needed elements



          end
          else
              inc(CurStart);        // no < found, move to next

        until CurStart^ = #0;       // end of doc

        // --- No Document or End Of Document: Terminate Scan
        if (CurStart = nil) or (CurStart^ = #0) then
            CurStart := StrEnd (FBuffer);

        Release_Memory_Of_Xlsx_File_Open;       // release buffer

        VTr.RootNodeCount := iRowNo + 1;
        VTr.EndUpdate;


        if bFoundDrawing then begin// Drawing found, need to preserve information of images
            // Knowing drawing filename from sheet filename
            Get_Drawing_Filename( ExtractFilename(sXML_Doc), i, sRid );

            // Above got information of images, one more step to preserve
            // ToCol and ToRow
            for ix := 0 to High(ImagesOnSheet) do begin
              if ImagesOnSheet[ix].iSheetID = i then begin // this sheet no
                  iImgID := ImagesOnSheet[ix].iImageID;
                  if iImgID <> -1 then begin

                      iImgWidth := AImage[iImgID].Picture.Width;
                      iImgHeight := AImage[iImgID].Picture.Height;

                      iCol_is := ImagesOnSheet[ix].iCol;
                      ImagesOnSheet[ix].iToCol := ImagesOnSheet[ix].iCol; // initialize

                      if VTr.Header.Columns.Count > iCol_is then begin
                          iColWidth_is := VTr.Header.Columns[iCol_is].Width;
                      //if SG.ColCount > iCol_is then begin
                      //    iColWidth_is := SG.ColWidths[iCol_is];

                          while iImgWidth > iColWidth_is do begin
                            if iCol_is < (VTr.Header.Columns.Count -1) then begin
                            //if iCol_is < (SG.ColCount -1) then begin
                                ImagesOnSheet[ix].iToCol := iCol_is + 1;

                                iCol_is := iCol_is + 1; // next col
                                // Accumulation of col width
                                iColWidth_is := iColWidth_is + VTr.Header.Columns[iCol_is].Width +
                                                1;
                                //iColWidth_is := iColWidth_is + SG.ColWidths[iCol_is] +
                                //                SG.GridLineWidth;

                            end
                            else
                                break;

                          end;
                      end;


                      iRow_is := ImagesOnSheet[ix].iRow;
                      ImagesOnSheet[ix].iToRow := ImagesOnSheet[ix].iRow; // initialize

                      //iNodeCnt := 0;
                      if VTr.RootNodeCount > iRow_is then begin
                          NextNode := VTr.GetFirst;
                          if NextNode <> nil then begin
                              repeat
                                //Inc(iNodeCnt);
                                if NextNode.Index = iRow_is then begin
                                    iRowHeight_is := NextNode.NodeHeight;
                                    break;
                                end;
                                NextNode := VTr.GetNext(NextNode);
                              until NextNode = nil;
                          end;

                          //iRowHeight_is := SG.RowHeights[iRow_is];

                      //if SG.RowCount > iRow_is then begin
                      //    iRowHeight_is := SG.RowHeights[iRow_is];

                          while iImgHeight > iRowHeight_is do begin
                            if iRow_is < (VTr.RootNodeCount -1) then begin
                            //if iRow_is < (SG.RowCount -1) then begin
                                ImagesOnSheet[ix].iToRow := iRow_is + 1;

                                iRow_is := iRow_is + 1; // next row
                                // Accumulation of row height

                                //iNodeCnt := 0;
                                NextNode := VTr.GetFirst;
                                if NextNode <> nil then begin
                                    repeat
                                      //Inc(iNodeCnt);
                                      if NextNode.Index = iRow_is then begin
                                          iRowHeight_is := iRowHeight_is +
                                                           NextNode.NodeHeight +
                                                           1;
                                          break;
                                      end;
                                      NextNode := VTr.GetNext(NextNode);
                                    until NextNode = nil;
                                end;


                                //iRowHeight_is := iRowHeight_is + SG.RowHeights[iRow_is] +
                                //                SG.GridLineWidth;

                            end
                            else
                                break;

                          end;
                      end;

                  end;

              end;
            end;

        end;

    end;


    iGo := iGo + iRowNo;        // for progress bar, subtotal row no of sheets


  end;

  finally
    pnPB.Visible := False;
    Screen.Cursor := crDefault;
  end;

  if sErrorMessage <> '' then   // show error messages
      PopUp_ErrorMessage(sErrorMessage);


end;



procedure TfrmMain.Get_SharedStrings_From_Xlsx;
const
  cXML_Document = 'xl\sharedStrings.xml';
var
  sXML_Doc, sMyText: string;
  NN: TNvpNode;
  //vEnter: variant;
  //i, iIndex: int64;
begin
  XslSharedStr.Clear; // delete old data

  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      ShowMessage('Err: Failing to extract data from xml file. Target directory not exist.' + #13#10);
      exit;
  end;

  sXML_Doc := XsExtractToThisPath + cXML_Document;

  if not FileExists(sXML_Doc) then begin
      ShowMessage('Err: Failing to extract data from xml file. File not found: ' + sXML_Doc + #13#10);
      exit;
  end;

  //vEnter := #$D#$A;
  //i := -1;
  //iIndex := -1;
  sMyText := '';

    XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
    XmlParser.StartScan;
    XmlParser.Normalize := FALSE;
    while XmlParser.Scan do begin

     { if (XmlParser.CurPartType = ptStartTag) or (XmlParser.CurPartType = ptEmptyTag)	then begin
          //if XmlParser.CurName = 'si' then
          //    i := i + 1;

      end; }

      if (XmlParser.CurPartType = ptContent) or (XmlParser.CurPartType = ptCData) then begin
          if XmlParser.CurName = 't' then begin
              //if AnsiPos(vEnter, XmlParser.CurContent) = 0 then begin // not accept #$D#$A
                  sMyText := sMyText + XmlParser.CurContent;

                 { if i = iIndex then
                      sMyText := sMyText + XmlParser.CurContent
                  else begin // new source index found
                      if sMyText <> '' then
                          XslSharedStr.Insert(iIndex, sMyText);

                      XslSharedStr.Insert(i, XmlParser.CurContent);
                      iIndex := i;
                      sMyText := '';
                  end; }
             // end;
          end;
      end;

      if (XmlParser.CurPartType = ptEndTag) then begin
          if XmlParser.CurName = 'si' then begin
              XslSharedStr.Append(sMyText);
              sMyText := '';
          end;
      end;

    end;

    //if sMyText <> '' then
    //    XslSharedStr.Append(sMyText);
        

end;

procedure TfrmMain.Get_SheetNames_From_Xlsx;
const
  cXML_Document = 'xl\workbook.xml';
var
  sXML_Doc, sRID: string;
  NN: TNvpNode;
begin
  XslRID.Clear; // delete old data

  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      ShowMessage('Err: Failing to extract data from xml file. Target directory not exist.' + #13#10);
      exit;
  end;

  sXML_Doc := XsExtractToThisPath + cXML_Document;

  if not FileExists(sXML_Doc) then begin
      ShowMessage('Err: Failing to extract data from xml file. File not found: ' + sXML_Doc + #13#10);
      exit;
  end;


  XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
  XmlParser.StartScan;
  XmlParser.Normalize := FALSE;
  while XmlParser.Scan do begin

      if (XmlParser.CurPartType = ptEmptyTag) then begin
          if (XmlParser.CurName = 'sheet') then begin
              NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('name'));
              try
                sRID := NN.Value;
                //sRID := AnsiReplaceStr(sRID, 'rId', '');
                XslRID.Append(SRID);
              finally
                NN.Free;
              end;
          end;
      end;


     { if (XmlParser.CurPartType = ptContent) or (XmlParser.CurPartType = ptCData) then begin

      end;

      if (XmlParser.CurPartType = ptEndTag) then begin

      end; }

  end;

end;


procedure TfrmMain.btnViewTxtClick(Sender: TObject);
var
  sExt: string;
begin
  if ZipFName.Caption = '' then begin
      beep;
      MessageDlg('There has not opened file.', mtError, [mbOK], 0);
      exit;
  end;

  WhatCompressionFile_WillBe_DirectlyExtracted(ZipFName.Caption);
  
  if XbIsDoUnzip then begin
      sExt := Ansilowercase(ExtractFileExt(ZipFName.Caption));

      if sExt = '.docx' then
          Output_Txt_Contents_Of_XML_From_Docx
      else if sExt = '.pptx' then
          Output_Txt_Contents_Of_XML_From_Pptx
      else if sExt = '.xlsx' then begin
          Remove_StringGrids_And_Tabsheets;
          Output_Txt_Contents_Of_XML_From_Xlsx_Speed;
          btnSaveAllToCSV.Enabled := True;
      end;
      XbIsDoUnzip := False;
  end;

  {$IFDEF IS_DEMO}

  {$ELSE}
  btnSaveTxtToFile.Enabled := True;
  {$ENDIF}
end;

procedure TfrmMain.btnSaveTxtToFileClick(Sender: TObject);
var
  sExt, sPath, sPathFilename: string;
  SaveDlg: TSaveDialog;

  i, j: integer;
  TS: TTabSheet;
  SG: TStringGrid;
  sSheet: string;
begin
  {$IFDEF IS_DEMO}
  Exit;
  {$ELSE}

  sExt := ExtractFileExt(ZipFName.Caption);
  sExt := Ansilowercase(sExt);

  if (sExt = '.docx') or (sExt = '.pptx') then begin

      SaveDlg := TSaveDialog.Create(nil);
      try
        SaveDlg.DefaultExt := 'txt';
        SaveDlg.Options := [ofOverwritePrompt,ofHideReadOnly,ofEnableSizing];
        sPath := ExtractFileDir(ZipFName.Caption);
        if DirectoryExists(sPath) then
            SaveDlg.InitialDir := sPath;

        SaveDlg.Filter := 'Text files (*.txt)|*.txt|All files (*.*)|*.*';

        if SaveDlg.Execute then begin
            if SaveDlg.FileName <> '' then begin
                sPathFilename := SaveDlg.FileName;
              //  if Lowercase(ExtractFileExt(sPathFilename)) <> '.txt' then
              //     sPathFilename := sPathFilename + '.txt';

                memOutput.Lines.SaveToFile(sPathFilename);
            end;
        end;

      finally
        SaveDlg.Free;
      end;
      
  end
  else if sExt = '.xlsx' then begin

      if pnWorksheets.Visible then begin
          i := PgeCtr.ActivePageIndex;

          if i >=0  then begin
              Xlsx_To_CSV_Speed(PgeCtr.ActivePageIndex, False);

              {
              sSheet := PgeCtr.Pages[i].Name;
              sSheet := sSheet + '.xml';

                  SaveDlg := TSaveDialog.Create(nil);
                  try
                    SaveDlg.DefaultExt := 'csv';
                    SaveDlg.Options := [ofOverwritePrompt,ofHideReadOnly,ofEnableSizing];
                    sPath := ExtractFileDir(frmMain.ZipFName.Caption);
                    if DirectoryExists(sPath) then
                        SaveDlg.InitialDir := sPath;

                    SaveDlg.Filter := 'CSV files (*.csv)|*.csv|All files (*.*)|*.*';

                    if SaveDlg.Execute then begin
                        if SaveDlg.FileName <> '' then begin
                            sPathFilename := SaveDlg.FileName;
                          //  if Lowercase(ExtractFileExt(sPathFilename)) <> '.txt' then
                          //     sPathFilename := sPathFilename + '.txt';
                            //VMemo := TMemo.Create(nil);
                            //VMemo.Parent := Self;  // Must have parent
                            //try
                              Xlsx_To_CSV_Speed(PgeCtr.ActivePageIndex, sPathFilename, False);


                            //finally
                            //  VMemo.Free;
                            //end;
                        end;
                    end;

                  finally
                    SaveDlg.Free;
                  end;
              }
          end;

      end;

  end;

  {$ENDIF}
end;


procedure TfrmMain.btnAdvancedViewClick(Sender: TObject);
var
  sNumFiles: string;
begin
  if btnAdvancedView.Down then begin
      Splitter2.Visible := True;
      pnM.Visible := True;
      btnViewTxt.Enabled := True;
      if (ZipFName.Caption <> '') then begin
          if ZipMaster1.Count > 1 then
              sNumFiles := 'Objects = '
          else
              sNumFiles := 'Object = ';

          StatusBar1.Panels[1].Text := sNumFiles + inttostr(ZipMaster1.Count);
      end;
  end
  else begin
      Splitter2.Visible := False;
      pnM.Visible := False;
      //if (ZipFName.Caption = '') then
          btnViewTxt.Enabled := False;
          
      StatusBar1.Panels[1].Text := ''; // remove it, wanted by payer
  end;
end;

procedure TfrmMain.btnAboutClick(Sender: TObject);
begin
  with TAboutBox.Create(Self) do
  try
    ShowModal;
  finally
    free;
  end;
end;

procedure TfrmMain.btnEditSelXMLClick(Sender: TObject);
var
  sPathFileName: string;
  bIsVirtualFolder, bIsEncrypted: boolean;
  sStrings: TStrings;
begin
  if VT.SelectedCount <> 1 then exit;

  sPathFileName := GetPathFileName_FromSelected_Item(bIsVirtualFolder, bIsEncrypted); // get path name and filename first then execute
  if (sPathFileName <> '') and (not bIsVirtualFolder) then begin // Normal view on ListView and it is a file
      //ExecuteFileClick(nil, sPathFileName, True, '');
      ZipMaster1.ExtractFileToStream(sPathFileName);

      with TfrmEditXML.Create(self) do
      try
        Screen.Cursor := crHourGlass;

        EditXML.sVirtual_PathFName := sPathFileName;
        Caption := Caption + ' - ' + sVirtual_PathFName;

        if XpObjModel1.LoadStream(ZipMaster1.ZipStream) then
            SynEdit1.Text := XpObjModel1.XmlDocument

        else begin
            ShowMessage('Error. Failing to open file. The file will be ' +
            'extracted and try to open it from disk after OK button pressed.' +
            ' More messages below:' + #10#13 + #10#13 +
            XpObjModel1.Errors.Text);

            RadioExtractWithDir.ItemIndex := 1;
            RadioOverwrite.ItemIndex := 1;

           { sToPath := ExtractFileDir(XsAppPath + sPathFileName);
            // Is the directory not exit?
            if not DirectoryExists(sToPath) then
                ForceDirectories(sToPath);  // create directory

            }

            if FileExists(XsAppPath + sPathFileName) then // remove old existed file first
                DeleteFile(XsAppPath + sPathFileName);

            try
              sStrings := TStringList.Create;
              sStrings.Append(sPathFileName);
              StartUnzip(False, ZipFName.Caption, XsAppPath, sStrings, True);
            finally
              sStrings.Free;
            end;

            SynEdit1.Lines.LoadFromFile(XsAppPath + sPathFileName);

        end;

        ShowModal;

      finally
        Screen.Cursor := crDefault;
        free;
      end;
  end;
end;

procedure TfrmMain.btnTestClick(Sender: TObject);
begin
  mnuTestClick(nil);
end;

procedure TfrmMain.btnRepairClick(Sender: TObject);
begin
  mnuRepairClick(nil);
end;

procedure TfrmMain.mnuViewXMLContentsClick(
  Sender: TObject);
begin
  btnEditSelXML.Click;
end;

procedure TfrmMain.pmWordOnPopup(Sender: TObject);
begin
  if VT.SelectedCount = 1 then
      mnuViewXMLContents.Visible := True
  else
      mnuViewXMLContents.Visible := False;

end;


procedure TfrmMain.AfterOpeningArchive_And_Do;
var
 { sFileNameIs, sFileExt: string;
  Node: PVirtualNode;
  Data: PEntry; }

  sZipFolder, sFileNameIs, sTem, sExt, sMediaDir: string;
  i: integer;
begin
  if btnAdvancedView.Down then begin
      VT.SetFocus;
      btnViewTxt.Enabled := True;
  end
  else
      btnViewTxt.Click;

  btnTest.Enabled := True;
  //btnRepair.Enabled := True;
  btnRecoverIMAGES.Visible := False; // Assume False first
  lblOpenFileIs.Caption := XcContentsOfDocx + ExtractFilename(ZipFName.Caption);

 { Node := VT.GetFirst;
  while Assigned(Node) do begin
    Data := VT.GetNodeData(Node);
    sFileNameIs := Data.Value[0]; // Name
    sFileExt := ExtractFileExt(sFileNameIs);
    sFileExt := Ansilowercase(sFileExt);

    if (sFileExt = '.bmp') or (sFileExt = '.jpg') or (sFileExt = '.jpeg') or
    (sFileExt = '.gif') or (sFileExt = '.png') or (sFileExt = '.wmf') then begin
        btnRecoverIMAGES.Visible := True;
        break;
    end;

    Node := VT.GetNext(Node);

    if Node = nil then
        break;

  end; }

  sExt := Ansilowercase(ExtractFileExt(ZipFname.Caption));

  pnWorksheets.Visible := (sExt = '.xlsx');

  if sExt = '.docx' then
      sMediaDir := 'word\media\'
  else if sExt = '.pptx' then
      sMediaDir := 'ppt\media\'
  else if sExt = '.xlsx' then
      sMediaDir := 'xl\media\'
  else
      exit;


  for i := 0 to ZipMaster1.ZipContents.Count -1 do begin
    with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
      sZipFolder := ExtractFilePath(FileName);
      if (sZipFolder <> '') then
          sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);

      sFileNameIs := ExtractFileName(FileName);
      sFileNameIs := AnsiReplaceStr(sFileNameIs, '*', ''); // remove password char

      sTem := Ansilowercase(sZipFolder);

      if sTem = sMediaDir then begin
          if sFileNameIs <> '' then begin
              btnRecoverIMAGES.Visible := True;
              break;
          end;
      end;

    end;
  end;


end;

procedure TfrmMain.btnHLImgFilesClick(Sender: TObject);
begin
 { if ZipFname.Caption <> '' then
      VT.Repaint;  }

end;

procedure TfrmMain.btnRecoverIMAGESClick(Sender: TObject);
type
  TImgTypeCount = record

  sImageType: string[5];
  iCount: integer;
end;

var
  sFileNameIs, sZipFolder, sFileExt, sOutPath, sTem: string;
  sMediaDir, sExt, sTheFileExt: string;
  sStrings: TStrings;
  //Node: PVirtualNode;
  //Data: PEntry;
  i, j, iLen, iL: integer;
  sMsg, sCnt, sImages, sImgType: string;
  AImgTypeCount: array of TImgTypeCount;
  bFound: boolean;
begin
  if ZipFname.Caption = '' then exit;

  sExt := Ansilowercase(ExtractFileExt(ZipFname.Caption));

  if sExt = '.docx' then
      sMediaDir := 'word\media\'
  else if sExt = '.pptx' then
      sMediaDir := 'ppt\media\'
  else if sExt = '.xlsx' then
      sMediaDir := 'xl\media\'
  else begin
      ShowMessage('Error: No media files found.');
      exit;
  end;

  try
    sStrings := TStringList.Create;
    SetLength(AImgTypeCount, 1);       // reset
    AImgTypeCount[0].sImageType := '';
    AImgTypeCount[0].iCount := 0;

   { Node := VT.GetFirst;
    while Assigned(Node) do begin
      Data := VT.GetNodeData(Node);
      sFileNameIs := Data.Value[0]; // Name
      //sFileNameIs := AnsiReplaceStr(sFileNameIs, '*', ''); // remove password char
      sFileExt := ExtractFileExt(sFileNameIs);
      sFileExt := Ansilowercase(sFileExt);

      if (sFileExt = '.bmp') or (sFileExt = '.jpg') or (sFileExt = '.jpeg') or
      (sFileExt = '.gif') or (sFileExt = '.png') or (sFileExt = '.wmf') then begin
          sZipFolder := Data.Value[8]; // Path
          if (sZipFolder <> '') then
              sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);

          sStrings.Append(sZipFolder + sFileNameIs);
      end;

      Node := VT.GetNext(Node);
    end; }

    for i := 0 to ZipMaster1.ZipContents.Count -1 do begin
      with ZipDirEntry(ZipMaster1.ZipContents[i]^) do begin
        sZipFolder := ExtractFilePath(FileName);
        if (sZipFolder <> '') then
            sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);

        sFileNameIs := ExtractFileName(FileName);
        sFileNameIs := AnsiReplaceStr(sFileNameIs, '*', ''); // remove password char

        sTheFileExt := lowercase( ExtractFileExt(sFileNameIs) );


        sTem := Ansilowercase(sZipFolder);

        if sTem = sMediaDir then begin
            sStrings.Append(sZipFolder + sFileNameIs);

            if sTheFileExt <> '' then begin

                bFound := False;
                for j := Low(AImgTypeCount) to High(AImgTypeCount) do begin
                  if sTheFileExt = AImgTypeCount[j].sImageType then begin
                      bFound := True;
                      AImgTypeCount[j].iCount := AImgTypeCount[j].iCount + 1;
                      break;
                  end;
                end;

                if not bFound then begin
                    iLen := length(AImgTypeCount);
                    AImgTypeCount[iLen -1].sImageType := sTheFileExt;
                    AImgTypeCount[iLen -1].iCount := 1;
                    SetLength(AImgTypeCount, iLen +1);
                    AImgTypeCount[iLen].sImageType := '';     // initialize
                    AImgTypeCount[iLen].iCount := -1;
                end;

            end;

            {$IFDEF IS_DEMO}
            if sStrings.Count >= 2 then
                break;

            {$ENDIF}
            
        end;

      end;
    end;


    if sStrings.Count > 0 then begin
        JvBrowseFolder1.Title := 'Extract file(s) to...' + #13#10 +
        'Same filenames on target location will be overwritten.';

        if JvBrowseFolder1.Execute then begin
            sOutPath := JvBrowseFolder1.Directory;
            if sOutPath <> '' then
                sOutPath := IncludeTrailingPathDelimiter(sOutPath);
                
            //if Directoryexists(sOutPath) then begin
                RadioExtractWithDir.ItemIndex := 0; // 0 = no dir, 1 = with dir
                RadioOverwrite.ItemIndex := 1; // always overwrite
                StartUnzip(False, ZipFName.Caption, sOutPath, sStrings, True);
            //end;

            sMsg := '';

            {$IFDEF IS_DEMO}
            sMsg := sMsg + 'Demo version only extracts two image files at a time.' + #13#10;
            {$ENDIF}

            for j := Low(AImgTypeCount) to High(AImgTypeCount) do begin
              if (AImgTypeCount[j].sImageType <> '') and
              (AImgTypeCount[j].iCount <> -1) then begin

                  iL := Length(AImgTypeCount[j].sImageType);
                  sImgType := copy(AImgTypeCount[j].sImageType, 2, iL -1);   // remove first '.'

                  sImages := '';
                  if AImgTypeCount[j].iCount > 1 then
                      sImages := 'images'
                  else
                      sImages := 'image';

                  sCnt := inttostr(AImgTypeCount[j].iCount);
                  sMsg := sMsg + sCnt + ' ' + sImgType + ' ' + sImages + #13#10;
              end;
            end;

            //sMsg := sMsg + #13#10;
            sMsg := sMsg + 'saved to ' + sOutPath;

            ShowMessage(sMsg);

        end;
    end;

  finally
    sStrings.Free;
  end;
end;

{procedure TfrmMain.VTBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellRect: TRect);
var
  sFileNameIs, sFileExt: string;
  Data: PEntry;
begin
  if ( Column = 0 ) and ( btnHLImgFiles.Down ) then begin
      if ZipFName.Caption <> '' then begin
          Data := VT.GetNodeData(Node);
          sFileNameIs := Data.Value[0]; // Name
          sFileExt := ExtractFileExt(sFileNameIs);
          sFileExt := Ansilowercase(sFileExt);

          if (sFileExt = '.bmp') or (sFileExt = '.jpg') or (sFileExt = '.jpeg') or
          (sFileExt = '.gif') or (sFileExt = '.png') or (sFileExt = '.wmf') then begin
              TargetCanvas.Brush.Color := $E0E0E0; // paint grey
              TargetCanvas.FillRect(CellRect);
          end;
      end;
  end;
end; }

procedure TfrmMain.VTBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellPaintMode: TVTCellPaintMode; CellRect: TRect;
  var ContentRect: TRect);
var
  sFileNameIs, sFileExt: string;
  Data: PEntry;
begin
 { if ( Column = 0 ) and ( btnHLImgFiles.Down ) then begin
      if ZipFName.Caption <> '' then begin
          Data := VT.GetNodeData(Node);
          sFileNameIs := Data.Value[0]; // Name
          sFileExt := ExtractFileExt(sFileNameIs);
          sFileExt := Ansilowercase(sFileExt);

          if (sFileExt = '.bmp') or (sFileExt = '.jpg') or (sFileExt = '.jpeg') or
          (sFileExt = '.gif') or (sFileExt = '.png') or (sFileExt = '.wmf') then begin
              TargetCanvas.Brush.Color := $E0E0E0; // paint grey
              TargetCanvas.FillRect(CellRect);
          end;
      end;
  end; }
end;

procedure TfrmMain.btnHelpClick(Sender: TObject);
var
  RetValue: integer;
  sPathFilename: string;
begin
  sPathFilename := XsAppPath + 'crworde.chm';
  if FileExists(sPathFilename) then begin
      RetValue := ShellExecute(handle, 'open', PChar(sPathFilename),
      '', '', SW_SHOW); // if success, return different value where depends on program

      if RetValue = 31 then begin // Has not the filename association?
          beep;
          MessageDlg('No application is associated with this kind of file extension.' +
          ' Failing to open file.', mtError, [mbOK], 0);
      end;
  end
  else begin
    beep;
    MessageDlg('Help file not found.', mtInformation, [mbOK], 0);
  end;
end;


procedure TfrmMain.VETreeEnumFolder(
  Sender: TCustomVirtualExplorerTree; Namespace: TNamespace;
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

procedure TfrmMain.VETreeClick(Sender: TObject);
var
  Node: PVirtualNode;
  NS: TNamespace;
  sFileExt, sPathFilename: string;
begin
  Node := VETree.GetFirstSelected;
  if VETree.ValidateNamespace(Node, NS) then begin
      sPathFilename := NS.NameForParsing;

      sFileExt := AnsiLowercase(ExtractFileExt(sPathFilename));
      if not FileExists(sPathFilename) then
          exit
      { else if (sFileExt <> '.zip') and (sFileExt <> '.exe')
      and (sFileExt <> '.docx') then
          exit; }
      else if (sFileExt <> '.docx') and (sFileExt <> '.pptx') and
      (sFileExt <> '.xlsx') then
          exit;


      if btnShowVirtual.Down then begin
          SetZipFile(sPathFilename, False);
          StartUnzip(True, sPathFilename, '', nil, False);
      end
      else
          SetZipFile(sPathFilename, True);

      Remove_StringGrids_And_Tabsheets; // must do first before AfterOpeningArchive_And_Do
      AfterOpeningArchive_And_Do;

      // Are some files extracted to Windows\Temp\xxx ? True to remove it.
      if bExtractedFiles and (edShadow.Text <> '') then
          Remove_TempFiles_UnderWinTemp;

      bExtractedFiles := False;
  end;
end;


// ***

{procedure TfrmMain.ShellDir1AddFile(Sender: TObject;
  var SearchRec: TSearchRec; var AddFile: Boolean);
begin
  if ((SearchRec.Attr and faDirectory) = 0) then begin // greater than 0 is Dir, equal to 0 is not Dir
      if FilterComboBox1.ItemIndex = 0 then begin // Filter is docx format
          if AnsiLowercase( ExtractFileExt(SearchRec.Name) ) <> '.docx' then
              AddFile := False;

      end; }
     { else if FilterComboBox1.ItemIndex = 1 then begin // Filter is zip format
          if AnsiLowercase( ExtractFileExt(SearchRec.Name) ) <> '.zip' then
              AddFile := False;

      end
      else if FilterComboBox1.ItemIndex = 2 then begin
          if AnsiLowercase( ExtractFileExt(SearchRec.Name) ) <> '.exe' then
              AddFile := False;

      end; }
//  end;
//end;

{procedure TfrmMain.ShellDir1Changing(Sender: TObject; Item: TListItem;
  Change: TItemChange; var AllowChange: Boolean);
begin
  Update_lblCurDir(ShellDir1.Path);
end; }

{procedure TfrmMain.ShellDir1Click(Sender: TObject);
var
  sFileExt, sPathFilename: string;
begin
  if ShellDir1.Selected = nil then exit;

  sPathFilename := ShellDir1.GetFullFileName(ShellDir1.Selected);
  sFileExt := AnsiLowercase(ExtractFileExt(sPathFilename));
  if not FileExists(sPathFilename) then
      exit
  else if (sFileExt <> '.zip') and (sFileExt <> '.exe')
  and (sFileExt <> '.docx') then
      exit;
  else if (sFileExt <> '.docx') then
      exit;


  if btnShowVirtual.Down then begin
      SetZipFile(sPathFilename, False);
      StartUnzip(True, sPathFilename, '', nil, False);
  end
  else
      SetZipFile(sPathFilename, True);

  AfterOpeningArchive_And_Do;

  //VT.SetFocus;

  // Are some files extracted to Windows\Temp\xxx ? True to remove it.
  if bExtractedFiles and (edShadow.Text <> '') then
      Remove_TempFiles_UnderWinTemp;

  bExtractedFiles := False;
end; }

{procedure TfrmMain.ShellDir1DblClick(Sender: TObject);
var
  sFileExt, sPathFilename: string;
begin
  if ShellDir1.Selected = nil then exit;  // after pressing Enter will come in

  sPathFilename := ShellDir1.GetFullFileName(ShellDir1.Selected);
  sFileExt := AnsiLowercase(ExtractFileExt(sPathFilename));
  if not FileExists(sPathFilename) then
      exit
  else if (sFileExt <> '.zip') and (sFileExt <> '.exe')
  and (sFileExt <> '.docx') then
      exit;

  if btnShowVirtual.Down then begin
      SetZipFile(sPathFilename, False);
      StartUnzip(True, sPathFilename, '', nil, False);
  end
  else
      SetZipFile(sPathFilename, True);

  AfterOpeningArchive_And_Do;

  //VT.SetFocus;
  // Are some files extracted to Windows\Temp\xxx ? True to remove it.
  if bExtractedFiles and (edShadow.Text <> '') then
      Remove_TempFiles_UnderWinTemp;

  bExtractedFiles := False;
end; }

{procedure TfrmMain.ShellDir1DDFileOperationExecuted(Sender: TObject;
  dwEffect: Integer; SourcePath, TargetPath: String);
begin
  ShellDir1.ReLoad(True);
end; }

{procedure TfrmMain.ShellDir1ExecFile(Sender: TObject; Item: TListItem;
  var AllowExec: Boolean);
begin
  if FileExists(ShellDir1.GetFullFileName(Item)) then
      AllowExec := False;

end; }

{procedure TfrmMain.ShellDir1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Pos: TPoint;
  iWidth, iHeight: integer;
begin
  if (Shift = [ssCtrl]) and (Key = 85) then
      btnUp1.Click
  else if (Shift = [ssCtrl]) and (Key = vk_Left) then
      btnRefresh1.Click
  else if (Shift = [ssCtrl]) and (Key = vk_Home) then
      btnHome.Click
  else if (Shift = [ssCtrl]) and (Key = 83) then begin// s
      btnViewStyle.Down := not btnViewStyle.Down;
      btnViewStyle.Click;
  end
  else if (Shift = [ssCtrl]) and (Key = 70) then // f
      btnFavoriteDir.Click
  else if (Shift = [ssCtrl]) and (Key = vk_Right) then begin
      Pos := btnFavoriteDir.ClientOrigin;
      iWidth := btnFavoriteDir.Width;
      iHeight := btnFavoriteDir.Height;
      pmFavoriteDir.Popup(Pos.X + iWidth, Pos.Y + iHeight);
  end;
end; }

{procedure TfrmMain.cbShellDrive1Change(Sender: TObject);
var
  cA: char;
  sDrive: string;
begin
  cA := cbShellDrive1.Drive;
  if length(cA) = 1 then
      sDrive := cA + ':';

  if DirectoryExists(sDrive) then
      ShellDir1.Path := sDrive
  else begin
      if bWin9598 then
          MessageDlg('Drive not ready or inaccessible.', mtInformation, [mbOK], 0);
          
  end;
end; }

{procedure TfrmMain.cbShellDrive1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  cA: char;
  sDrive: string;
begin
  if Key = 13 then begin
      cA := cbShellDrive1.Drive;
      if length(cA) = 1 then
          sDrive := cA + ':';

      if DirectoryExists(sDrive) then begin
          ShellDir1.Path := sDrive;
          //ShellDir1.SetFocus;
      end;
  end;
end; }

{procedure TfrmMain.WMDeviceChange(Var Message : TMessage);
const
  DBT_DEVICEARRIVAL	   = $8000;	// system detected a new device
  DBT_DEVICEREMOVECOMPLETE = $8004;     // device is gone

Begin
  Inherited;
  IF (Message.wParam = DBT_DEVICEARRIVAL) OR
     (Message.wParam = DBT_DEVICEREMOVECOMPLETE) Then
  Begin
    DriveView1.RefreshRootNodes(True, dsAll And Not dvdsFloppy Or dvdsRereadAllways);
    cbShellDrive1.ResetItems;
  End;
End; } {WMDeviceChange}

procedure TfrmMain.VETreeFocusChanged(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex);
var
  NS: TNamespace;
  sDir: string;
begin
  if VETree.ValidateNamespace(Node, NS) then begin
      sDir := NS.NameForParsing;
      if not NS.Folder then
          sDir := ExtractFileDir(sDir);

      Update_lblCurDir(sDir);
  end;
end;

procedure TfrmMain.Remove_StringGrids_And_Tabsheets;
var
  i, j, k, m: integer;
  //ChildControl: TControl;
  //TS: TTabSheet;
  SG: TStringGrid;
  VTr: TVirtualStringTree;
  //KG: TKGrid;
begin
  if pnWorksheets.Visible then begin

      for i := PgeCtr.PageCount -1 downto 0 do begin
        for j := 0 to PgeCtr.Pages[i].ComponentCount -1 do begin
          if PgeCtr.Pages[i].Components[j] is TVirtualStringTree then begin
              VTr := PgeCtr.Pages[i].Components[j] as TVirtualStringTree;

              VTr.Clear;
              VTr.Free;
             { for k := 0 to SG.RowCount -1 do begin
                for m := 0 to SG.ColCount -1 do begin
                  SG.Cells[k, m] := '';
                  //SG.Objects[k, m].Free;

                end;
              end; }

              //FreeAndNil(SG);
              //SG := nil;

             { SG.Perform(WM_SETREDRAW, 0, 0);

              try
                if SG.RowCount < SG.ColCount then begin
                    for k := SG.FixedRows to SG.RowCount  -1 do begin
                      SG.Rows[k].Clear;
                    end;
                end
                else begin
                    for k := SG.FixedCols to SG.ColCount  -1 do begin
                      SG.Cols[k].Clear;
                    end;
                end;
              finally
                SG.Perform(WM_SETREDRAW, 1, 0);
                SG.Invalidate;
              end; }

          end;
        end;
      end;

      for i := PgeCtr.PageCount -1 downto 0 do begin
        PgeCtr.Pages[i].Free;
      end;

      FreeBitmap;
      Release_Memory_Of_Xlsx_File_Open;
      
  end;
end;

procedure TfrmMain.Get_Drawing_Filename(const sSheetFile: string;
  const iSheetIndexIs: integer; const sRid: string);
const
  cSheet_Rels_Dir = 'xl\worksheets\_rels\';
  //cDrawing_Dir = 'xl\drawings\';
  //cRelsDrawing_Dir = 'xl\drawings\_rels\';
var
  sXML_Doc, sDrawFilename: string;
  // sXML_Draw, sXML_Rels_Draw, sRelsDrawFilename, sDrawFilename: string;
  NN: TNvpNode;
  bFoundRid: boolean;
begin
  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      sErrorMessage := sErrorMessage + 'Err: Target directory not exist when trying to get drawing filename.' + #13#10;
      exit;
  end;

  // Example:   xl\worksheets\_rels\sheet1.xml.rels
  sXML_Doc := XsExtractToThisPath + cSheet_Rels_Dir + sSheetFile + '.rels';

  if not FileExists(sXML_Doc) then begin
      sErrorMessage := sErrorMessage + 'Err: The file not found. Filename: ' + sXML_Doc + #13#10;
      exit;
  end;


  bFoundRid := False;


    XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
    XmlParser.StartScan;
    XmlParser.Normalize := FALSE;
    while XmlParser.Scan do begin

      if (XmlParser.CurPartType = ptStartTag) then begin

      end;



      if (XmlParser.CurPartType = ptEndTag) then begin

      end;



      if (XmlParser.CurPartType = ptEmptyTag) then begin

          if XmlParser.CurName = 'Relationship' then begin
              // go to find drawing.xml
              NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('Id'));
              try
                if NN.Value = sRid then
                    bFoundRid := True;

              finally
                NN.Free;
              end;


              if bFoundRid then begin
                  NN := TNvpNode.Create(XmlParser.CurStart, XmlParser.CurAttr.Value('Target'));
                  try
                    // likes this: ../drawings/drawing.xml
                    sDrawFilename := AnsiReplaceStr(NN.Value, '/', '\');
                    sDrawFilename := ExtractFilename(sDrawFilename); // get it!
                  finally
                    NN.Free;
                  end;

                  bFoundRid := False;
              end;
          end

      end;


      if (XmlParser.CurPartType = ptContent) or (XmlParser.CurPartType = ptCData) then begin

      end;

    end;


    if sDrawFilename <> '' then  // The next step to get images pos
        Get_Images_And_Positions(sDrawFilename, iSheetIndexIs);


end;


procedure TfrmMain.Get_Images_And_Positions(const sDrawFileIs: string;
  const iSheetIndexIs: integer);
const
  cDrawing_Dir = 'xl\drawings\';
  cMedia_Dir = 'xl\media\';
var
  sXML_Doc, sRID, sCx, sCy, sImgFname, sImgFullPath: string;
  // sXML_Draw, sXML_Rels_Draw, sRelsDrawFilename, sDrawFilename: string;
  NN: TNvpNode;
  iL: integer;
  bXdrFrom, bXdrBlipFill: boolean;
  bXdrSpPr, bAxFrm: boolean;
  iCx, iCy: integer;
  //TemBitmap, DestBitmap: TBitmap;
begin

  if DirectoryExists(XsExtractToThisPath) then
      XsExtractToThisPath := IncludeTrailingPathDelimiter(XsExtractToThisPath)
  else begin
      sErrorMessage := sErrorMessage + 'Err: Target directory not exist.' + #13#10;
      exit;
  end;

  // Example:  xl\drawings\drawing1.xml
  sXML_Doc := XsExtractToThisPath + cDrawing_Dir + sDrawFileIs;

  if not FileExists(sXML_Doc) then begin
      sErrorMessage := sErrorMessage + 'Err: The file not found. Filename: ' + sXML_Doc + #13#10;
      exit;
  end;


  iL := High(ImagesOnSheet); // no record is: -1, one record is: 0, two is: 1
  if iL = 0 then begin
      if ImagesOnSheet[0].iSheetID = -2 then  // is first initail record
          iL := 0                             // don't want this record
      else                                    // preserved record of previous sheets
          iL := iL +1;
  end
  else                           // found preserved record of previous sheets
      iL := iL +1;

  bXdrFrom := False;
  bXdrBlipFill := False;
  bXdrSpPr := False;
  bAxFrm := False;
  iCx := -1;
  iCy := -1;


    XmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
    XmlParser.StartScan;
    XmlParser.Normalize := FALSE;
    while XmlParser.Scan do begin 

      if (XmlParser.CurPartType = ptStartTag) then begin
          if XmlParser.CurName = 'xdr:from' then begin
              bXdrFrom := True;
              iL := iL +1;  // iL = 1 is first image...

              SetLength(ImagesOnSheet, iL);
              SetLength(AImage, iL);

              ImagesOnSheet[iL -1].iSheetID := iSheetIndexIs; // sheet no
              ImagesOnSheet[iL -1].iImageID := -1; // initialize
          end
          else if XmlParser.CurName = 'xdr:blipFill' then
              bXdrBlipFill := True
          else if XmlParser.CurName = 'xdr:spPr' then
              bXdrSpPr := True
          else if XmlParser.CurName = 'a:xfrm' then
              bAxFrm := True;

      end;



      if (XmlParser.CurPartType = ptEndTag) then begin
          if XmlParser.CurName = 'xdr:from' then
              bXdrFrom := False
          else if XmlParser.CurName = 'xdr:blipFill' then
              bXdrBlipFill := False
          else if XmlParser.CurName = 'xdr:spPr' then
              bXdrSpPr := False
          else if XmlParser.CurName = 'a:xfrm' then begin
              bAxFrm := False;
              iCx := -1;
              iCy := -1;
          end
      end;



      if (XmlParser.CurPartType = ptEmptyTag) then begin
          if XmlParser.CurName = 'a:blip' then begin
              if bXdrBlipFill then begin
                  NN := TNvpNode.Create(XmlParser.CurStart, Xmlparser.CurAttr.Value('r:embed'));
                  sRID := NN.Value;

                  if sRID <> '' then begin
                      sImgFname := Get_Image_Filename_From_Rels(sDrawFileIs, sRID);
                      sImgFullPath := XsExtractToThisPath + cMedia_Dir + sImgFname;

                      // Example:  xl\meida\Image1.png
                      if FileExists(sImgFullPath) then begin

                          AImage[iL -1] := TImage.Create(pnBack);
                          AImage[iL -1].Align := alNone;
                          AImage[iL -1].Parent := pnBack;
                          AImage[iL -1].Width := 32;
                          AImage[iL -1].Height := 32;

                          AImage[iL -1].Picture.LoadFromFile(sImgFullPath);

                          if AImage[iL -1].Picture.Graphic <> nil then begin
                              AImage[iL -1].Tag := iL -1; // is image id for iImageID
                              ImagesOnSheet[iL -1].iImageID := AImage[iL -1].Tag;
                          end;

                      end
                      else
                          sErrorMessage := sErrorMessage + 'Err: File not found. Filename: ' + sImgFullPath + #13#10;

                  end
                  else
                      sErrorMessage := sErrorMessage + 'Err: Failing to get image filename.' + #13#10;

              end;
          end
          else if XmlParser.CurName = 'a:ext' then begin
              if bXdrSpPr and bAxFrm then begin
                  try
                    NN := TNvpNode.Create(XmlParser.CurStart, Xmlparser.CurAttr.Value('cx'));
                    sCx := NN.Value;

                   if Trystrtoint(sCx, iCx) then begin
                       iCx := iCx div 7620;
                       ImagesOnSheet[iL -1].iCx := iCx;
                   end;
                  finally
                    NN.Free;
                  end;

                  try
                    NN := TNvpNode.Create(XmlParser.CurStart, Xmlparser.CurAttr.Value('cy'));
                    sCy := NN.Value;

                    if Trystrtoint(sCy, iCy) then begin
                       iCy := iCy div 7620;
                       ImagesOnSheet[iL -1].iCy := iCy;
                    end;
                  finally
                    NN.Free;
                  end;


                 { if (iCx <> -1) and (iCy <> -1) then begin
                      if AImage[iL -1].Picture.Graphic <> nil then begin
                          try
                            TemBitmap := TBitmap.Create;
                            DestBitmap := TBitmap.Create;

                            TemBitmap.Width := AImage[iL -1].Picture.Graphic.Width;
                            TemBitmap.Height := AImage[iL -1].Picture.Graphic.Height;
                            TemBitmap.Assign(AImage[iL -1].Picture.Graphic);



                            DestBitmap.Width := iCx;
                            DestBitmap.Height := iCy;

                            // shrink with proportional
                            // important to show original color in Win 2000
                            SetStretchBltMode(DestBitmap.Canvas.Handle, HALFTONE);

                            StretchBlt(DestBitmap.Canvas.Handle, 0, 0,
                                        DestBitmap.Width, DestBitmap.Height,
                                        TemBitmap.Canvas.Handle, 0, 0,
                                        TemBitmap.Width,
                                        TemBitmap.Height, SrcCopy);


                            if AImage[iL -1].Picture.Graphic is TPNGObject then begin
                                DestBitmap.Transparent := TPNGObject(AImage[iL -1].Picture.Graphic).Transparent;
                                DestBitmap.TransparentColor := TPNGObject(AImage[iL -1].Picture.Graphic).TransparentColor;
                                //TemBitmap.TransparentMode := tmAuto;
                            end;


                            AImage[iL -1].Picture.Graphic := nil;
                            AImage[iL -1].Picture.Graphic := DestBitmap;

                          finally
                            DestBitmap.Free;
                            TemBitmap.Free;
                          end
                      end;
                  end; }

              end;
          end;
      end;


      if (XmlParser.CurPartType = ptContent) or (XmlParser.CurPartType = ptCData) then begin
          if (XmlParser.CurName = 'xdr:col') and bXdrFrom then begin
              ImagesOnSheet[iL -1].iCol := strtoint( XmlParser.CurContent ) +1; // Excel sheets 0,0 not include col header
          end
          else if (XmlParser.CurName = 'xdr:row') and bXdrFrom then begin
              ImagesOnSheet[iL -1].iRow := strtoint( XmlParser.CurContent );

          end;
      end;
    end;



end;


function TfrmMain.Get_Image_Filename_From_Rels(const sDrawFileIs: string;
  const sRidIs: string): string;
const
  cRelsDrawing_Dir = 'xl\drawings\_rels\';
var
  sXML_Doc, sImgFilename: string;
  NN: TNvpNode;
  bFoundRID: boolean;
begin

  if sDrawFileIs <> '' then
      sXML_Doc := XsExtractToThisPath + cRelsDrawing_Dir + sDrawFileIs + '.rels';


  if not FileExists(sXML_Doc) then begin
      ShowMessage('Err: Failing to draw images on sheet. The file not found. Filename: ' + sXML_Doc + #13#10);
      exit;
  end;


  bFoundRID := False;


  DrawXmlParser.LoadFromFile(sXML_Doc, fmOpenRead or fmShareDenyNone);
  DrawXmlParser.StartScan;
  DrawXmlParser.Normalize := FALSE;
  while DrawXmlParser.Scan do begin

    if (DrawXmlParser.CurPartType = ptStartTag) then begin
        //if XmlParser.CurName = 'Relationships' then begin



        //end;
    end;


    if (DrawXmlParser.CurPartType = ptEmptyTag) then begin

        if DrawXmlParser.CurName = 'Relationship' then begin
            NN := TNvpNode.Create(DrawXmlParser.CurStart, DrawXmlParser.CurAttr.Value('Id'));
            try
              if sRidIs = NN.Value then
                  bFoundRID := True;
            finally
                NN.Free;
            end;

            if bFoundRID then begin
                NN := TNvpNode.Create(DrawXmlParser.CurStart, DrawXmlParser.CurAttr.Value('Target'));
                try
                  // Example:  ../media/image1.png
                  sImgFilename := AnsiReplaceStr(NN.Value, '/', '\');
                  if sImgFilename <> '' then 
                      Result := ExtractFilename(sImgFilename);

                finally
                  NN.Free;
                end;

                bFoundRID := False;
                exit;

            end;

        end;

    end;

  end;

end;


procedure TfrmMain.FreeBitmap;
var
  i: integer;
begin
  if pnBack.ComponentCount = 0 then exit;
  if not ( High(AImage) < Low(AImage) ) then begin
      For i := 0 to High(AImage) do begin
        //AImage[i].Picture.Bitmap.FreeImage;
        if Assigned(Aimage[i]) then
            AImage[i].Picture.Graphic := nil;
            
      end;
  end;
end;


procedure TfrmMain.SGDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  i, k, iSheetNo, iCol, iRow, iToCol, iToRow, iImagIndex, iCx, iCy: integer;
  iCnt, iH, iW, iWidth, iHeight: integer;
  TS: TTabSheet;
  SG: TStringGrid;
  ARect: TRect;
begin

  if Sender is TStringGrid then begin
      SG := Sender as TStringGrid;
      if SG.Parent is TTabSheet then begin
          TS := SG.Parent as TTabSheet;

          For i := 0 to High(ImagesOnSheet) do begin
            if ImagesOnSheet[i].iSheetID = TS.PageIndex then begin
                iCol := ImagesOnSheet[i].iCol;
                iRow := ImagesOnSheet[i].iRow;
                iToCol := ImagesOnSheet[i].iToCol;
                iToRow := ImagesOnSheet[i].iToRow;
                iImagIndex := ImagesOnSheet[i].iImageID;
                iCx := ImagesOnSheet[i].iCx;
                iCy := ImagesOnSheet[i].iCy;

                if AImage[iImagIndex].Picture.Graphic = nil then
                    break;

               { if (ACol = iCol) and (ARow = iRow) then begin
                    SG.Canvas.Draw(Rect.Left, Rect.Top,
                                   AImage[iImagIndex].Picture.Graphic);


                end }
                if ((ACol >= iCol) and (ACol <= iToCol)) and
                ((ARow >= iRow) and (ARow <= iToRow)) then begin

                    iH := 0;
                    iCnt := iRow;
                    for k := iRow to ARow -1 do begin
                      iH := iH + SG.RowHeights[iCnt] + SG.GridLineWidth;
                      iCnt := iCnt + 1; // next row
                    end;

                    iW := 0;
                    iCnt := iCol;
                    for k := iCol to ACol -1 do begin
                      iW := iW + SG.ColWidths[iCnt] + SG.GridLineWidth;
                      iCnt := iCnt + 1; // next col
                    end;


                    ABitmap.FreeImage;
                    BBitmap.FreeImage;
                    TeBitmap.FreeImage;

                    ABitmap.Width := AImage[iImagIndex].Picture.Graphic.Width;
                    ABitmap.Height := AImage[iImagIndex].Picture.Graphic.Height;
                    ABitmap.PixelFormat := pf24bit;
                    ABitmap.Assign(AImage[iImagIndex].Picture.Graphic);
                    // shrink with proportional
                    // important to show original color in Win 2000
                    SetStretchBltMode(TeBitmap.Canvas.Handle, HALFTONE);

                    if (iCx > -1) and (iCy > -1) and ( (iCx <> ABitmap.Width)
                    or (iCy <> ABitmap.Height) ) then begin
                        TeBitmap.Width := iCx;  // must initilize first before StretchBlt
                        TeBitmap.Height := iCy; // must initilize first before StretchBlt
                        TeBitmap.PixelFormat := pf24bit;

                        StretchBlt(TeBitmap.Canvas.Handle, 0, 0,
                                   TeBitmap.Width, TeBitmap.Height,
                                   ABitmap.Canvas.Handle, 0, 0,
                                   ABitmap.Width, ABitmap.Height, SrcCopy);

                        ABitmap.FreeImage;
                        ABitmap.Width := TeBitmap.Width;
                        ABitmap.Height := TeBitmap.Height;
                        
                        ABitmap.Assign(TeBitmap);
                    end;


                    iWidth := ABitmap.Width - iW;
                    if iWidth <= 0 then begin
                        ImagesOnSheet[i].iToCol := ACol -1; // update last col
                        break;
                    end;

                    iHeight := ABitmap.Height - iH;
                    if iHeight <= 0 then begin
                        ImagesOnSheet[i].iToRow := ARow -1; // update last row
                        break;
                    end;

                    BBitmap.Width := iWidth; // less than 0, cause error
                    BBitmap.Height := iHeight; // less than 0, cause error
                    BBitmap.PixelFormat := pf24bit;

                    BitBlt(BBitmap.Canvas.Handle, 0, 0, BBitmap.Width,
                    BBitmap.Height, ABitmap.Canvas.Handle, iW, iH, SrcCopy);

                    if AImage[iImagIndex].Picture.Graphic is TPNGObject then begin
                        BBitmap.Transparent := TPNGObject(AImage[iImagIndex].Picture.Graphic).Transparent;
                        BBitmap.TransparentColor := TPNGObject(AImage[iImagIndex].Picture.Graphic).TransparentColor;
                        //BBitmap.TransparentMode := tmAuto;
                    end;

                    //ARect := SG.CellRect(iCol, iRow);

                    //if (ARect.Left > 0) and (ARect.Top > 0) then begin
                        SG.Canvas.Draw(Rect.Left, Rect.Top, BBitmap);
                    //end;
                end;
            end;
          end;
      end;
  end;

end;

procedure TfrmMain.SGMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  i, iSheetNo, iCol, iRow, iToCol, iToRow, iImagIndex, iCx, iCy: integer;
  ACol, ARow: integer;
  ACoord: TGridCoord;
  SG: TStringGrid;
  TS: TTabSheet;
begin

  if Sender is TStringGrid then begin
      SG := Sender as TStringGrid;
      if SG.Parent is TTabSheet then begin
          TS := SG.Parent as TTabSheet;

          For i := 0 to High(ImagesOnSheet) do begin
            if ImagesOnSheet[i].iSheetID = TS.PageIndex then begin
                iCol := ImagesOnSheet[i].iCol;
                iRow := ImagesOnSheet[i].iRow;
                iToCol := ImagesOnSheet[i].iToCol;
                iToRow := ImagesOnSheet[i].iToRow;
                iImagIndex := ImagesOnSheet[i].iImageID;
                iCx := ImagesOnSheet[i].iCx;
                iCy := ImagesOnSheet[i].iCy;

                ACoord := SG.MouseCoord(X, Y);

                SG.MouseToCell(X, Y, ACol, ARow);

                if ((ACol >= iCol) and (ACol <= iToCol)) and
                ((ARow >= iRow) and (ARow <= iToRow)) then begin
                    SG.Cursor := crHandPoint;
                    break;
                end
                else
                    SG.Cursor := crDefault;

            end;
          end;
      end;
  end;

end;

procedure TfrmMain.SGMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  i, iSheetNo, iCol, iRow, iToCol, iToRow, iImagIndex, iCx, iCy: integer;
  ACol, ARow: integer;
  ACoord: TGridCoord;
  SG: TStringGrid;
  TS: TTabSheet;
begin
  if Button <> mbRight then exit;

  if Sender is TStringGrid then begin
      SG := Sender as TStringGrid;
      if SG.Parent is TTabSheet then begin
          TS := SG.Parent as TTabSheet;

          For i := 0 to High(ImagesOnSheet) do begin
            if ImagesOnSheet[i].iSheetID = TS.PageIndex then begin
                iCol := ImagesOnSheet[i].iCol;
                iRow := ImagesOnSheet[i].iRow;
                iToCol := ImagesOnSheet[i].iToCol;
                iToRow := ImagesOnSheet[i].iToRow;
                iImagIndex := ImagesOnSheet[i].iImageID;
                iCx := ImagesOnSheet[i].iCx;
                iCy := ImagesOnSheet[i].iCy;

                ACoord := SG.MouseCoord(X, Y);

                SG.MouseToCell(X, Y, ACol, ARow);

                if ((ACol >= iCol) and (ACol <= iToCol)) and
                ((ARow >= iRow) and (ARow <= iToRow)) then begin
                    ClipBitmap.FreeImage;
                    ClipBitmap.Assign(AImage[i].Picture.Graphic);

                    if AImage[i].Picture.Graphic is TPNGObject then begin
                        ClipBitmap.Transparent := TPNGObject(AImage[i].Picture.Graphic).Transparent;
                        ClipBitmap.TransparentColor := TPNGObject(AImage[i].Picture.Graphic).TransparentColor;
                        //ClipBitmap.TransparentMode := tmAuto;
                    end;

                    pmSG.Popup(self.Left + pnB.Left + X + 20,
                    self.Top + pnB.Top + Y - 10);
                    break;
                end;
               { else
                    SG.Cursor := crDefault; }

            end;
          end;
      end;
  end;

end;

procedure TfrmMain.pmCopyImgObjClick(Sender: TObject);
begin
  {$IFDEF IS_DEMO}
  {$ELSE}
  Clipboard.Assign(ClipBitmap);
  {$ENDIF}
end;

procedure TfrmMain.pnWorksheetsResize(Sender: TObject);
begin
  pnPB.Top := (pnWorksheets.Height div 2) - (pnPB.Height div 2);
  pnPB.Left := (pnWorksheets.Width div 2) - (pnPB.Width div 2);
end;

procedure TfrmMain.VTrGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: WideString);
var
  Data: PMyEntry;
begin
  //if Column > 0 then begin
      Data := Sender.GetNodeData(Node);
      CellText := Data.Value[Column];
  //end
  //else
  //    CellText := '';
      
end;

procedure TfrmMain.VTrAfterCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellRect: TRect);
var
  i, k, iSheetNo, iCol, iRow, iToCol, iToRow, iImagIndex, iCx, iCy: integer;
  iCnt, iH, iW, iWidth, iHeight, ACol, ARow, iColWidth: integer;
  TS: TTabSheet;
  NextNode: PVirtualNode;
begin

      if Sender.Parent is TTabSheet then begin
          TS := Sender.Parent as TTabSheet;

          For i := 0 to High(ImagesOnSheet) do begin
            if ImagesOnSheet[i].iSheetID = TS.PageIndex then begin
                iCol := ImagesOnSheet[i].iCol;
                iRow := ImagesOnSheet[i].iRow;
                iToCol := ImagesOnSheet[i].iToCol;
                iToRow := ImagesOnSheet[i].iToRow;
                iImagIndex := ImagesOnSheet[i].iImageID;
                iCx := ImagesOnSheet[i].iCx;
                iCy := ImagesOnSheet[i].iCy;

                if AImage[iImagIndex].Picture.Graphic = nil then
                    break;

               { if (ACol = iCol) and (ARow = iRow) then begin
                    SG.Canvas.Draw(Rect.Left, Rect.Top,
                                   AImage[iImagIndex].Picture.Graphic);


                end }

                ACol := Column;
                ARow := Node.Index;
                
                if ((ACol >= iCol) and (ACol <= iToCol)) and
                ((ARow >= iRow) and (ARow <= iToRow)) then begin

                    iH := 0;
                    iCnt := iRow;
                    for k := iRow to ARow -1 do begin

                      NextNode := Sender.GetFirst;
                      if NextNode <> nil then begin
                          repeat
                            //Inc(iNodeCnt);
                            if NextNode.Index = iCnt then begin
                                iH := iH + NextNode.NodeHeight + 1;
                                break;
                            end;
                            NextNode := Sender.GetNext(NextNode);
                          until NextNode = nil;
                      end;

                      //iH := iH + SG.RowHeights[iCnt] + SG.GridLineWidth;
                      iCnt := iCnt + 1; // next row
                    end;

                    iW := 0;
                    iCnt := iCol;
                    for k := iCol to ACol -1 do begin
                      iColWidth := (Sender as TVirtualStringTree).Header.Columns[iCnt].Width;
                      iW := iW + iColWidth + 1;
                      //iW := iW + SG.ColWidths[iCnt] + SG.GridLineWidth;
                      iCnt := iCnt + 1; // next col
                    end;


                    ABitmap.FreeImage;
                    BBitmap.FreeImage;
                    TeBitmap.FreeImage;

                    ABitmap.Width := AImage[iImagIndex].Picture.Graphic.Width;
                    ABitmap.Height := AImage[iImagIndex].Picture.Graphic.Height;
                    ABitmap.PixelFormat := pf24bit;
                    ABitmap.Assign(AImage[iImagIndex].Picture.Graphic);
                    // shrink with proportional
                    // important to show original color in Win 2000
                    SetStretchBltMode(TeBitmap.Canvas.Handle, HALFTONE);

                    if (iCx > -1) and (iCy > -1) and ( (iCx <> ABitmap.Width)
                    or (iCy <> ABitmap.Height) ) then begin
                        TeBitmap.Width := iCx;  // must initilize first before StretchBlt
                        TeBitmap.Height := iCy; // must initilize first before StretchBlt
                        TeBitmap.PixelFormat := pf24bit;

                        StretchBlt(TeBitmap.Canvas.Handle, 0, 0,
                                   TeBitmap.Width, TeBitmap.Height,
                                   ABitmap.Canvas.Handle, 0, 0,
                                   ABitmap.Width, ABitmap.Height, SrcCopy);

                        ABitmap.FreeImage;
                        ABitmap.Width := TeBitmap.Width;
                        ABitmap.Height := TeBitmap.Height;
                        
                        ABitmap.Assign(TeBitmap);
                    end;


                    iWidth := ABitmap.Width - iW;
                    if iWidth <= 0 then begin
                        ImagesOnSheet[i].iToCol := ACol -1; // update last col
                        break;
                    end;

                    iHeight := ABitmap.Height - iH;
                    if iHeight <= 0 then begin
                        ImagesOnSheet[i].iToRow := ARow -1; // update last row
                        break;
                    end;

                    BBitmap.Width := iWidth; // less than 0, cause error
                    BBitmap.Height := iHeight; // less than 0, cause error
                    BBitmap.PixelFormat := pf24bit;

                    BitBlt(BBitmap.Canvas.Handle, 0, 0, BBitmap.Width,
                    BBitmap.Height, ABitmap.Canvas.Handle, iW, iH, SrcCopy);

                    if AImage[iImagIndex].Picture.Graphic is TPNGObject then begin
                        BBitmap.Transparent := TPNGObject(AImage[iImagIndex].Picture.Graphic).Transparent;
                        BBitmap.TransparentColor := TPNGObject(AImage[iImagIndex].Picture.Graphic).TransparentColor;
                        //BBitmap.TransparentMode := tmAuto;
                    end;

                    //ARect := SG.CellRect(iCol, iRow);

                    //if (ARect.Left > 0) and (ARect.Top > 0) then begin
                        TargetCanvas.Draw(CellRect.Left, CellRect.Top, BBitmap);
                        //SG.Canvas.Draw(Rect.Left, Rect.Top, BBitmap);
                    //end;
                end;
            end;
          end;
      end;


 { if Column = 0 then begin
      with TargetCanvas do begin
        Inc(CellRect.Right);
        Inc(CellRect.Bottom);
        DrawEdge(Handle, CellRect, BDR_RAISEDINNER, BF_RECT or BF_MIDDLE);
        //if Node = Sender.FocusedNode then

      end;
  end; }
end;

procedure TfrmMain.VTrBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellPaintMode: TVTCellPaintMode; CellRect: TRect;
  var ContentRect: TRect);
begin
  if Column = 0 then begin
      with TargetCanvas do begin
        Inc(CellRect.Right);
        Inc(CellRect.Bottom);
        DrawEdge(Handle, CellRect, BDR_RAISEDINNER, BF_RECT or BF_MIDDLE);
        //if Node = Sender.FocusedNode then

      end;
  end
  else begin
      TargetCanvas.Brush.Color := $E0E0E0;
      TargetCanvas.FillRect(CellRect);
  end;

  // paint whole focused cell
  if ((Column = Sender.FocusedColumn) and (Node = Sender.FocusedNode)) then begin
      TargetCanvas.Brush.Color := clHighlight;
      TargetCanvas.FillRect(CellRect);
  end;
end;

procedure TfrmMain.VTrFocusChanging(Sender: TBaseVirtualTree; OldNode,
  NewNode: PVirtualNode; OldColumn, NewColumn: TColumnIndex;
  var Allowed: Boolean);
begin
  Allowed := NewColumn > 0;
end;

procedure TfrmMain.VTrMeasureItem(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; var NodeHeight: Integer);
var
  i, k: integer;
  VTr: TVirtualStringTree;
  TS: TTabSheet;
begin
  if Sender.Parent is TTabSheet then begin

      TS := Sender.Parent as TTabSheet;
      k := TS.PageIndex;
      for i := iStart_HOfRow[k] to High(AHeightOfRow) do begin
        if AHeightOfRow[i].iSheetID = TS.PageIndex then begin
            if AHeightOfRow[i].iIndex = Node.Index then begin
                NodeHeight := AHeightOfRow[i].iHeight;
                iStart_HOfRow[k] := i + 1;
                break;
            end;
        end;
      end;
      
  end;
end;

procedure TfrmMain.VTrMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var
  i, iSheetNo, iCol, iRow, iToCol, iToRow, iImagIndex, iCx, iCy: integer;
  ACol, ARow: integer;
  VTr: TVirtualStringTree;
  TS: TTabSheet;
  Node: PVirtualNode;
  P: TPoint;
begin

  if Sender is TVirtualStringTree then begin
      VTr := Sender as TVirtualStringTree;

      Node := VTr.GetNodeAt(X, Y);
      if Node = nil then
          exit;



      if VTr.Parent is TTabSheet then begin
          TS := VTr.Parent as TTabSheet;

          For i := 0 to High(ImagesOnSheet) do begin
            if ImagesOnSheet[i].iSheetID = TS.PageIndex then begin
                iCol := ImagesOnSheet[i].iCol;
                iRow := ImagesOnSheet[i].iRow;
                iToCol := ImagesOnSheet[i].iToCol;
                iToRow := ImagesOnSheet[i].iToRow;
                iImagIndex := ImagesOnSheet[i].iImageID;
                iCx := ImagesOnSheet[i].iCx;
                iCy := ImagesOnSheet[i].iCy;

                ARow := Node.Index;           // Col header is not count!
                //ACoord := SG.MouseCoord(X, Y);

                P.X := X;
                P.Y := Y;
                ACol := VTr.Header.Columns.ColumnFromPosition(P);

                //SG.MouseToCell(X, Y, ACol, ARow);

                if ((ACol >= iCol) and (ACol <= iToCol)) and
                ((ARow >= iRow) and (ARow <= iToRow)) then begin
                    VTr.Cursor := crHandPoint;
                    //SG.Cursor := crHandPoint;
                    break;
                end
                else
                    VTr.Cursor := crDefault;
                    //SG.Cursor := crDefault;
            end;
          end;
      end;
  end;

end;


procedure TfrmMain.VTrMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  i, iSheetNo, iCol, iRow, iToCol, iToRow, iImagIndex, iCx, iCy: integer;
  ACol, ARow: integer;
  //ACoord: TGridCoord;
  VTr: TVirtualStringTree;
  TS: TTabSheet;
  Node: PVirtualNode;
  P: TPoint;
begin
  if Button <> mbRight then exit;

  if Sender is TVirtualStringTree then begin
      VTr := Sender as TVirtualStringTree;

      Node := VTr.GetNodeAt(X, Y);
      if Node = nil then
          exit;


      if VTr.Parent is TTabSheet then begin
          TS := VTr.Parent as TTabSheet;

          For i := 0 to High(ImagesOnSheet) do begin
            if ImagesOnSheet[i].iSheetID = TS.PageIndex then begin
                iCol := ImagesOnSheet[i].iCol;
                iRow := ImagesOnSheet[i].iRow;
                iToCol := ImagesOnSheet[i].iToCol;
                iToRow := ImagesOnSheet[i].iToRow;
                iImagIndex := ImagesOnSheet[i].iImageID;
                iCx := ImagesOnSheet[i].iCx;
                iCy := ImagesOnSheet[i].iCy;


                ARow := Node.Index;           // Col header is not count!
                //ACoord := SG.MouseCoord(X, Y);

                P.X := X;
                P.Y := Y;
                ACol := VTr.Header.Columns.ColumnFromPosition(P);

                //ACoord := SG.MouseCoord(X, Y);

                //SG.MouseToCell(X, Y, ACol, ARow);

                if ((ACol >= iCol) and (ACol <= iToCol)) and
                ((ARow >= iRow) and (ARow <= iToRow)) then begin
                    ClipBitmap.FreeImage;
                    ClipBitmap.Assign(AImage[i].Picture.Graphic);

                    if AImage[i].Picture.Graphic is TPNGObject then begin
                        ClipBitmap.Transparent := TPNGObject(AImage[i].Picture.Graphic).Transparent;
                        ClipBitmap.TransparentColor := TPNGObject(AImage[i].Picture.Graphic).TransparentColor;
                        //ClipBitmap.TransparentMode := tmAuto;
                    end;

                    pmSG.Popup(self.Left + pnB.Left + X + 25,
                    self.Top + pnB.Top + Y + 15);
                    break;
                end;
               { else
                    SG.Cursor := crDefault; }

            end;
          end;
      end;
  end;

end;


procedure TfrmMain.Release_Memory_Of_Xlsx_File_Open;
begin
  if (FBufferSize > 0) and (FBuffer <> nil) then begin
      FreeMem(FBuffer);
      FBuffer := NIL;
      FBufferSize := 0;
  end;
end;

procedure TfrmMain.btnSaveAllToCSVClick(Sender: TObject);
var
  sExt: string;
  i: integer;
begin
  {$IFDEF IS_DEMO}
  Exit;
  {$ELSE}

  sExt := ExtractFileExt(ZipFName.Caption);
  sExt := Ansilowercase(sExt);

  if sExt = '.xlsx' then begin

      if pnWorksheets.Visible then begin
          i := PgeCtr.ActivePageIndex;

          if i >=0  then 
              Xlsx_To_CSV_Speed(-1, True);

      end;

  end
  else
      ShowMessage('Error, Not xlsx file found. Filename: ' + ZipFName.Caption);

  {$ENDIF}
end;

procedure TfrmMain.memOutputMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  {$IFDEF IS_DEMO}
  memOutput.SelLength := 0;
  {$ENDIF}
end;

procedure TfrmMain.memOutputKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  {$IFDEF IS_DEMO}
  memOutput.SelLength := 0;
  {$ENDIF}
end;




end.



