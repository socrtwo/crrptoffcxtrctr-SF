unit addfiles;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, ToolWin, ShlObj, ImgList, Menus, Buttons, ExtCtrls,
  Registry;

type
  PShellItem = ^TShellItem;
  TShellItem = record
    FullID,
    ID: PItemIDList;
    Empty: Boolean;
    DisplayName,
    TypeName: string;
    ImageIndex,
    Size,
    Attributes: Integer;
    ModDate: string;
  end;

  TfrmAddFiles = class(TForm)
    ListView: TListView;
    CoolBar1: TCoolBar;
    ToolBar2: TToolBar;
    ToolbarImages: TImageList;
    btnBrowse: TToolButton;
    btnLargeIcons: TToolButton;
    btnSmallIcons: TToolButton;
    btnList: TToolButton;
    btnReport: TToolButton;
    cbPath: TComboBox;
    ToolButton3: TToolButton;
    PopupMenu1: TPopupMenu;
    btnBack: TToolButton;
    PageControl1: TPageControl;
    tabAdd: TTabSheet;
    tabUnzip: TTabSheet;
    Bevel1: TBevel;
    GroupBox1: TGroupBox;
    btnAll: TButton;
    btnShowZips: TButton;
    btnHideFolders: TSpeedButton;
    ToolButton1: TToolButton;
    btnUnzipThis: TBitBtn;
    btnCancelUnzipThis: TBitBtn;
    lblFileExt: TLabel;
    lblPath: TLabel;
    edPath: TEdit;
    lblFileExt2: TLabel;
    ToolButton2: TToolButton;
    pnTop: TPanel;
    lblCurPath: TLabel;
    cbUpper: TComboBox;
    Bevel2: TBevel;
    Bevel3: TBevel;
    gbAddFiles: TGroupBox;
    btnAdd: TBitBtn;
    btnCancelAdd: TBitBtn;
    cbCompressRatio: TComboBox;
    chKeepAboveSet: TCheckBox;
    gb1: TGroupBox;
    Label1: TLabel;
    chAddWithDir: TCheckBox;
    chOverwriteWithNew: TCheckBox;
    btnComment: TButton;
    lblCompressRatio: TLabel;
    lblThisFile: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure ListViewData(Sender: TObject; Item: TListItem);
    procedure btnBrowseClick(Sender: TObject);
    procedure cbPathKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cbPathClick(Sender: TObject);
    procedure btnLargeIconsClick(Sender: TObject);
    procedure ListViewDblClick(Sender: TObject);
    procedure ListViewDataHint(Sender: TObject; StartIndex,
      EndIndex: Integer);
    procedure ListViewKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ListViewDataFind(Sender: TObject; Find: TItemFind;
      const FindString: String; const FindPosition: TPoint;
      FindData: Pointer; StartIndex: Integer; Direction: TSearchDirection;
      Wrap: Boolean; var Index: Integer);
    procedure ListViewCustomDrawItem(Sender: TCustomListView;
      Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure ListViewCustomDrawSubItem(Sender: TCustomListView;
      Item: TListItem; SubItem: Integer; State: TCustomDrawState;
      var DefaultDraw: Boolean);
    procedure btnBackClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnCancelAddClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure chKeepAboveSetMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure chKeepAboveSetKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure chAddWithDirMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure chAddWithDirKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure chOverwriteWithNewMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure chOverwriteWithNewKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cbCompressRatioChange(Sender: TObject);
    procedure btnUnzipThisClick(Sender: TObject);
    procedure btnCancelUnzipThisClick(Sender: TObject);
    procedure ListViewClick(Sender: TObject);
    procedure ListViewColumnClick(Sender: TObject; Column: TListColumn);
    procedure SetFlagsForListView(const NewFlags: integer);
    procedure SetFilterForListView(const NewFilter: string);
    procedure btnAllClick(Sender: TObject);
    procedure RefreshListView;
    procedure btnShowZipsClick(Sender: TObject);
    procedure ListViewKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnHideFoldersClick(Sender: TObject);
    procedure cbPathDropDown(Sender: TObject);
    procedure cbUpperDropDown(Sender: TObject);
    procedure cbUpperClick(Sender: TObject);
    procedure btnCommentClick(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    FPIDL: PItemIDList;
    FIDList: TList;
    FIShellFolder,
    FIDesktopFolder: IShellFolder;
    FPath: string;
    procedure SetPath(const Value: string); overload;
    procedure SetPath(ID: PItemIDList); overload;
    procedure PopulateIDList(ShellFolder: IShellFolder);
    procedure ClearIDList;
    procedure CheckShellItems(StartIndex, EndIndex: Integer);
    function  ShellItem(Index: Integer): PShellItem;

    procedure SettingsBeforeAdd;
    procedure OnOffchkKeepSettings;
    procedure OnOffchkAddWithDir;
    procedure OnOffchkOverwriteWithNew;
    procedure RestoreSettingsOffrmAddFiles;
    procedure HistoryPathToReg(const sPath: string);
    procedure RestoreCaptionOfColumn;
    procedure FoundDataInListView;
    function Restore_WinNT_ClassID_To_Name(sClassID: string): string;
    function Read_NullValue_In_ClassID_FromReg(sSubKey: string; sValueName: string): string;
  end;

var
  frmAddFiles: TfrmAddFiles;

implementation

{$R *.dfm}

uses ShellAPI, ActiveX, ComObj, CommCtrl, FileCtrl, Main, InputMsgBox, ZipMstr;

const
  cCheckKeepAboveSet = 'CheckKeepAboveSet';
  cCheckAddWithDir = 'CheckAddWithDir';
  cCheckOverwriteWithNew = 'CheckOverwriteWithNew';
  cIndexOfCompressRatio = 'IndexOfCompressRatio';
  cSelfWidth = 'Width Of addfiles';
  cSelfHeight = 'Height Of addfiles';

  cCountOfHistoryPath = 10;
  cCountOfUpperPath = 10;
var
  sSubKeyHistoryDir: string = XcSubKeyIs + '\HistoryFolder';
  sCurPath: string;
  bCancelProcess: boolean = True;
  bDescending: boolean = False;
  iColBeSorted: byte = 0;

  CurFlags: integer = SHCONTF_FOLDERS or SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN;
  sCurFilter: string = '*.*';
  sComment: string = '';
//PIDL MANIPULATION

procedure DisposePIDL(ID: PItemIDList);
var
  Malloc: IMalloc;
begin
  if ID = nil then Exit;
  OLECheck(SHGetMalloc(Malloc));
  Malloc.Free(ID);
end;

function CopyITEMID(Malloc: IMalloc; ID: PItemIDList): PItemIDList;
begin
  Result := Malloc.Alloc(ID^.mkid.cb + SizeOf(ID^.mkid.cb));
  CopyMemory(Result, ID, ID^.mkid.cb + SizeOf(ID^.mkid.cb));
end;

function NextPIDL(IDList: PItemIDList): PItemIDList;
begin
  Result := IDList;
  Inc(PChar(Result), IDList^.mkid.cb);
end;

function GetPIDLSize(IDList: PItemIDList): Integer;
begin
  Result := 0;
  if Assigned(IDList) then
  begin
    Result := SizeOf(IDList^.mkid.cb);
    while IDList^.mkid.cb <> 0 do
    begin
      Result := Result + IDList^.mkid.cb;
      IDList := NextPIDL(IDList);
    end;
  end;
end;


procedure StripLastID(IDList: PItemIDList);
var
  MarkerID: PItemIDList;
begin
  MarkerID := IDList;
  if Assigned(IDList) then
  begin
     while IDList.mkid.cb <> 0 do
    begin
      MarkerID := IDList;
      IDList := NextPIDL(IDList);
    end;
    MarkerID.mkid.cb := 0;
  end;
end;

function CreatePIDL(Size: Integer): PItemIDList;
var
  Malloc: IMalloc;
  HR: HResult;
begin
  Result := nil;

  HR := SHGetMalloc(Malloc);
  if Failed(HR) then
    Exit;

  try
    Result := Malloc.Alloc(Size);
    if Assigned(Result) then
      FillChar(Result^, Size, 0);
  finally
  end;
end;

function CopyPIDL(IDList: PItemIDList): PItemIDList;
var
  Size: Integer;
begin
  Size := GetPIDLSize(IDList);
  Result := CreatePIDL(Size);
  if Assigned(Result) then
    CopyMemory(Result, IDList, Size);
end;

function ConcatPIDLs(IDList1, IDList2: PItemIDList): PItemIDList;
var
  cb1, cb2: Integer;
begin
  if Assigned(IDList1) then
    cb1 := GetPIDLSize(IDList1) - SizeOf(IDList1^.mkid.cb)
  else
    cb1 := 0;

  cb2 := GetPIDLSize(IDList2);

  Result := CreatePIDL(cb1 + cb2);
  if Assigned(Result) then
  begin
    if Assigned(IDList1) then
      CopyMemory(Result, IDList1, cb1);
    CopyMemory(PChar(Result) + cb1, IDList2, cb2);
  end;
end;

//SHELL FOLDER ITEM INFO

function GetDisplayName(ShellFolder: IShellFolder; PIDL: PItemIDList;
                        ForParsing: Boolean): string;
var
  StrRet: TStrRet;
  P: PChar;
  Flags: Integer;
begin
  Result := '';
  if ForParsing then
    Flags := SHGDN_FORPARSING
  else
    Flags := SHGDN_NORMAL;

  ShellFolder.GetDisplayNameOf(PIDL, Flags, StrRet);
  case StrRet.uType of
    STRRET_CSTR:
      SetString(Result, StrRet.cStr, lStrLen(StrRet.cStr));
    STRRET_OFFSET:
      begin
        P := @PIDL.mkid.abID[StrRet.uOffset - SizeOf(PIDL.mkid.cb)];
        SetString(Result, P, PIDL.mkid.cb - StrRet.uOffset);
      end;
    STRRET_WSTR:
      Result := StrRet.pOleStr;
  end;
end;

function GetShellImage(PIDL: PItemIDList; Large, Open: Boolean): Integer;
var
  FileInfo: TSHFileInfo;
  Flags: Integer;
begin
  FillChar(FileInfo, SizeOf(FileInfo), #0);
  Flags := SHGFI_PIDL or SHGFI_SYSICONINDEX or SHGFI_ICON;
  if Open then Flags := Flags or SHGFI_OPENICON;
  if Large then Flags := Flags or SHGFI_LARGEICON
  else Flags := Flags or SHGFI_SMALLICON;
  SHGetFileInfo(PChar(PIDL),
                0,
                FileInfo,
                SizeOf(FileInfo),
                Flags);
  Result := FileInfo.iIcon;
end;

function IsFolder(ShellFolder: IShellFolder; ID: PItemIDList): Boolean;
var
  Flags: UINT;
begin
  Flags := SFGAO_FOLDER;
  ShellFolder.GetAttributesOf(1, ID, Flags);
  Result := SFGAO_FOLDER and Flags <> 0;
end;


function ListSortFunc(Item1, Item2: Pointer): Integer;
begin
  if not bDescending then
      Result := SmallInt( frmAddFiles.FIShellFolder.CompareIDs(iColBeSorted,
                PShellItem(Item1).ID, PShellItem(Item2).ID ) )
  else
      Result := SmallInt( frmAddFiles.FIShellFolder.CompareIDs(iColBeSorted,
                PShellItem(Item2).ID, PShellItem(Item1).ID ) );
end;

{TForm1}

//GENERAL FORM METHODS

procedure TfrmAddFiles.FormCreate(Sender: TObject);
var
  FileInfo: TSHFileInfo;
  ImageListHandle: THandle;
  NewPIDL: PItemIDList;
begin
  Self.Icon := Application.Icon;
  bCancelProcess := True;
  edPath.Text := '';
  lblCurPath.Caption := '';
  cbCompressRatio.Items.Insert(0, 'Maximum');
  cbCompressRatio.Items.Insert(1, 'Normal');
  cbCompressRatio.Items.Insert(2, 'Low');
  cbCompressRatio.Items.Insert(3, 'None');
  cbCompressRatio.ItemIndex := 1;
  
  OLECheck(SHGetDesktopFolder(FIDesktopFolder));
  FIShellFolder := FIDesktopFolder;
  FIDList := TList.Create;
  ImageListHandle := SHGetFileInfo('C:\',
                           0,
                           FileInfo,
                           SizeOf(FileInfo),
                           SHGFI_SYSICONINDEX or SHGFI_SMALLICON);
  SendMessage(ListView.Handle, LVM_SETIMAGELIST, LVSIL_SMALL, ImageListHandle);

  ImageListHandle := SHGetFileInfo('C:\',
                           0,
                           FileInfo,
                           SizeOf(FileInfo),
                           SHGFI_SYSICONINDEX or SHGFI_LARGEICON);

  SendMessage(ListView.Handle, LVM_SETIMAGELIST, LVSIL_NORMAL, ImageListHandle);
  OLECheck(
    SHGetSpecialFolderLocation(
      Application.Handle,
      CSIDL_DRIVES,
      NewPIDL)
  );
  SetPath(NewPIDL);
  ActiveControl := cbPath;
  cbPath.SelStart := 0;
  cbPath.SelLength := Length(cbPath.Text);
  Self.SetFilterForListView('.*');

  RestoreSettingsOffrmAddFiles;
end;

procedure TfrmAddFiles.btnBrowseClick(Sender: TObject);
var
  S: string;
begin
  S := '';
  if SelectDirectory('Select Directory', '', S) then
    SetPath(S);
end;

procedure TfrmAddFiles.cbPathKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    if cbPath.Text[Length(cbPath.Text)] = ':' then
      cbPath.Text := cbPath.Text + '\';
    SetPath(cbPath.Text);
    Key := 0;
  end;
end;

procedure TfrmAddFiles.cbPathClick(Sender: TObject);
var
  I: Integer;
begin
  if cbPath.Text = '' then exit;
  I := cbPath.Items.IndexOf(cbPath.Text);
  if cbPath.Text[Length(cbPath.Text)] = ':' then // added by me
      cbPath.Text := cbPath.Text + '\';

  if I >= 0 then begin
    if cbPath.Items.Objects[I] <> nil then // added by me
        SetPath(PItemIDList(cbPath.Items.Objects[I]))
    else
        SetPath(cbPath.Text); // added by me

  end
  else
    SetPath(cbPath.Text);
end;

procedure TfrmAddFiles.btnLargeIconsClick(Sender: TObject);
begin
  ListView.ViewStyle := TViewStyle((Sender as TComponent).Tag);
end;

procedure TfrmAddFiles.ListViewDblClick(Sender: TObject);
var
  RootPIDL,
  ID: PItemIDList;
  sFileType: string;
begin
  if ListView.Selected <> nil then
  begin
    ID := ShellItem(ListView.Selected.Index).ID;
    if not IsFolder(FIShellFolder, ID) then begin
        if tabUnzip.TabVisible and (not tabAdd.TabVisible) then begin
            sFileType := ShellItem(ListView.Selected.Index).TypeName; // added by me
            sFileType := lowercase(sFileType);
            if (sFileType = 'zip') and ( FileExists(edPath.Text) ) then begin
                btnUnzipThis.Click;
            end;
        end;
        Exit;
    end;
    RootPIDL := ConcatPIDLs(FPIDL, ID);
    SetPath(RootPIDL);
    if ListView.ViewStyle <> vsReport then
        SetPath(RootPIDL); // I don't know why! This can get correct icons

    ListView.ClearSelection; // added by me
  end;
end;

function TfrmAddFiles.ShellItem(Index: Integer): PShellItem;
begin
  Result := PShellItem(FIDList[Index]);
end;

procedure TfrmAddFiles.ListViewKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_RETURN:
      ListViewDblClick(Sender);
    VK_BACK:
      btnBackClick(Sender);
  end;
end;

//SHELL-RELATED ROUTINES.

procedure TfrmAddFiles.ClearIDList;
var
  I: Integer;
begin
  for I := 0 to FIDList.Count-1 do
  begin
    DisposePIDL(ShellItem(I).ID);
    Dispose(ShellItem(I));
  end;
  FIDList.Clear;
end;

procedure TfrmAddFiles.PopulateIDList(ShellFolder: IShellFolder);
//const
//  Flags = SHCONTF_FOLDERS or SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN;
var
  ID: PItemIDList;
  EnumList: IEnumIDList;
  NumIDs: LongWord;
  SaveCursor: TCursor;
  ShellItem: PShellItem;

  FileInfo: TSHFileInfo;
begin
  SaveCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourglass;
    OleCheck(
      ShellFolder.EnumObjects(
        Application.Handle,
        CurFlags, // altered by me, origin is Flags
        EnumList)
    );

    FIShellFolder := ShellFolder;
    ClearIDList;
    while EnumList.Next(1, ID, NumIDs) = S_OK do
    begin
      ShellItem := New(PShellItem);
      ShellItem.ID := ID;
      ShellItem.DisplayName := GetDisplayName(FIShellFolder, ID, False);
      ShellItem.Empty := True;
      lblFileExt.Caption := ExtractFileExt(ShellItem.DisplayName);
      lblFileExt.Caption := lowercase(lblFileExt.Caption);

      ShellItem.FullID := ConcatPIDLs(FPIDL, ShellItem.ID);
      //ImageIndex := GetShellImage(FullID, ListView.ViewStyle = vsIcon, False);

      //File Type
      SHGetFileInfo(
        PChar(ShellItem.FullID),
        0,
        FileInfo,
        SizeOf(FileInfo),
        SHGFI_TYPENAME or SHGFI_PIDL
      );

      if ( IsFolder(FIShellFolder, ShellItem.ID) ) or (lblFileExt.Caption =
      sCurFilter) or ( lowercase(FileInfo.szTypeName) = 'zip' ) or (sCurFilter
      = '.*') then
          FIDList.Add(ShellItem);
    end;

    iColBeSorted := 0; // added by me
    FIDList.Sort(ListSortFunc);

    //We need to tell the ListView how many items it has.
    ListView.Items.Count := FIDList.Count;

    ListView.Repaint;
  finally
    Screen.Cursor := SaveCursor;
  end;
end;

procedure TfrmAddFiles.SetPath(const Value: string);
var
  P: PWideChar;
  NewPIDL: PItemIDList;
  Flags,
  NumChars: LongWord;
begin
  if not DirectoryExists(Value) then exit; // added by me
  lblCurPath.Caption := Value;
  
  NumChars := Length(Value);
  Flags := 0;
  P := StringToOleStr(Value);

  OLECheck(
    FIDesktopFolder.ParseDisplayName(
      Application.Handle,
      nil,
      P,
      NumChars,
      NewPIDL,
      Flags)
   );
  SetPath(NewPIDL);
end;

procedure TfrmAddFiles.SetPath(ID: PItemIDList);
var
  Index: Integer;
  NewShellFolder: IShellFolder;
  i: integer;
  sNTClassID: string;
  bOldObjFound: boolean;
begin
   OLECheck(
     FIDesktopFolder.BindToObject(
            ID,
            nil,
            IID_IShellFolder,
            Pointer(NewShellFolder))
   );

  ListView.Items.BeginUpdate;
  try
  try
    PopulateIDList(NewShellFolder);
    FPIDL := ID;
    FPath := GetDisplayName(FIDesktopFolder, FPIDL, True);

    //if (Win32Platform = VER_PLATFORM_WIN32_NT) then begin
        if copy(FPath, 1, 3) = '::{' then begin
            sNTClassID := copy( FPath, 3, length(FPath)-2 );
            FPath := Restore_WinNT_ClassID_To_Name(sNTClassID);
        end;
    //end;

    sCurPath := IncludeTrailingPathDelimiter(FPath); // added by me
    Index := cbPath.Items.IndexOf(FPath);
    lblCurPath.Caption := FPath;

    bOldObjFound := False;
    if (Index < 0) then begin // disabled by me
        //cbPath.Items.InsertObject(0, FPath, Pointer(FPIDL));
        //cbPath.Text := cbPath.Items[0];
        for i := 0 to cbUpper.Items.Count -1 do begin
          if cbUpper.Items.Strings[i] = FPath then begin
              bOldObjFound := True;
              break;
          end;
        end;
        if not bOldObjFound then
            cbUpper.Items.InsertObject(0, FPath, Pointer(FPIDL));

        cbUpper.Text := cbUpper.Items[0];
        for i := cbUpper.Items.Count -1 downto cCountOfUpperPath do begin
          cbUpper.Items.Delete(i);
        end;
    end
    else begin
      //cbPath.ItemIndex := Index;
      //cbPath.Text := cbPath.Items[cbPath.ItemIndex];
      for i := 0 to cbUpper.Items.Count -1 do begin
        if cbUpper.Items.Strings[i] = FPath then begin
              bOldObjFound := True;
              break;
        end;
      end;
      if not bOldObjFound then
          cbUpper.Items.Insert(0, FPath);

    end;

    if ListView.Items.Count > 0 then
    begin
      ListView.Selected := ListView.Items[0];
      ListView.Selected.Focused := True;
      ListView.Selected.MakeVisible(False);
    end;

  except// on E: Exception do

  end;
  finally
    ListView.Items.EndUpdate;
  end;
  ListView.ClearSelection; // added by me
  RestoreCaptionOfColumn; // added by me
end;

//ROUTINES FOR MANAGING VIRTUAL DATA

procedure TfrmAddFiles.CheckShellItems(StartIndex, EndIndex: Integer);

 function ValidFileTime(FileTime: TFileTime): Boolean;
 begin
   Result := (FileTime.dwLowDateTime <> 0) or (FileTime.dwHighDateTime <> 0);
 end;

var
  FileData: TWin32FindData;
  FileInfo: TSHFileInfo;
  SysTime: TSystemTime;
  I: Integer;
  LocalFileTime: TFILETIME;
begin
  //Here all the data that wasn't initialized in PopulateIDList is
  //filled in.
  for I := StartIndex to EndIndex do
  begin
    if ShellItem(I)^.Empty then
    with ShellItem(I)^ do
    begin
      FullID := ConcatPIDLs(FPIDL, ID);
      ImageIndex := GetShellImage(FullID, ListView.ViewStyle = vsIcon, False);

      //File Type
      SHGetFileInfo(
        PChar(FullID),
        0,
        FileInfo,
        SizeOf(FileInfo),
        SHGFI_TYPENAME or SHGFI_PIDL
      );
      TypeName := FileInfo.szTypeName;

      //Get File info from Windows
      FillChar(FileData, SizeOf(FileData), #0);
      SHGetDataFromIDList(
        FIShellFolder,
        ID,
        SHGDFIL_FINDDATA,
        @FileData,
        SizeOf(FileData)
      );

      //File Size, in KB
      Size := (FileData.nFileSizeLow + 1023 ) div 1024;
      if Size = 0 then Size := 1;

      //Modified Date
      FillChar(LocalFileTime, SizeOf(TFileTime), #0);
      with FileData do
        if ValidFileTime(ftLastWriteTime)
        and FileTimeToLocalFileTime(ftLastWriteTime, LocalFileTime)
        and FileTimeToSystemTime(LocalFileTime, SysTime) then
        try
          ModDate := DateTimeToStr(SystemTimeToDateTime(SysTime))
        except
          on EConvertError do ModDate := '';
        end
        else
          ModDate := '';

      //Attributes
      Attributes := FileData.dwFileAttributes;

      //Flag this record as complete.
      Empty := False;
    end;
  end;
end;

procedure TfrmAddFiles.ListViewDataHint(Sender: TObject; StartIndex,
  EndIndex: Integer);
begin
  //OnDataHint is called before OnData. This gives you a chance to
  //initialize only the data structures that need to be drawn.
  //You should keep track of which items have been initialized so no
  //extra work is done.
  if (StartIndex > FIDList.Count) or (EndIndex > FIDList.Count) then Exit;
  CheckShellItems(StartIndex, EndIndex);
end;

procedure TfrmAddFiles.ListViewData(Sender: TObject; Item: TListItem);
var
  Attrs: string;
begin
  //OnData gets called once for each item for which the ListView needs
  //data. If the ListView is in Report View, be sure to add the subitems.
  //Item is a "dummy" item whose only valid data is it's index which
  //is used to index into the underlying data.
  if (Item.Index > FIDList.Count) then Exit;
  with ShellItem(Item.Index)^ do
  begin
    Item.Caption := DisplayName;
    Item.ImageIndex := ImageIndex;

    if ListView.ViewStyle <> vsReport then Exit;

    if not IsFolder(FIShellFolder, ID) then
      Item.SubItems.Add(Format('%dKB', [Size]))
    else
      Item.SubItems.Add('');
    Item.SubItems.Add(TypeName);
    try
      Item.SubItems.Add(ModDate);
    except
    end;

    if Bool(Attributes and FILE_ATTRIBUTE_READONLY) then Attrs := Attrs + 'R';
    if Bool(Attributes and FILE_ATTRIBUTE_HIDDEN) then Attrs := Attrs + 'H';
    if Bool(Attributes and FILE_ATTRIBUTE_SYSTEM) then Attrs := Attrs + 'S';
    if Bool(Attributes and FILE_ATTRIBUTE_ARCHIVE) then Attrs := Attrs + 'A';
  end;
  Item.SubItems.Add(Attrs);
end;

procedure TfrmAddFiles.ListViewDataFind(Sender: TObject; Find: TItemFind;
  const FindString: String; const FindPosition: TPoint; FindData: Pointer;
  StartIndex: Integer; Direction: TSearchDirection; Wrap: Boolean;
  var Index: Integer);
//OnDataFind gets called in response to calls to FindCaption, FindData,
//GetNearestItem, etc. It also gets called for each keystroke sent to the
//ListView (for incremental searching)
var
  I: Integer;
  Found: Boolean;
begin
  I := StartIndex;
  if (Find = ifExactString) or (Find = ifPartialString) then
  begin
    repeat
      if (I = FIDList.Count-1) then
        if Wrap then I := 0 else Exit;
      Found := Pos(UpperCase(FindString), UpperCase(ShellItem(I)^.DisplayName)) = 1;
      Inc(I);
    until Found or (I = StartIndex);
    if Found then Index := I-1;
  end;
end;

procedure TfrmAddFiles.ListViewCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
var
  Attrs: Integer;
begin
  if Item = nil then Exit;
  Attrs := ShellItem(Item.Index).Attributes;
  if Bool(Attrs and FILE_ATTRIBUTE_READONLY) then
    ListView.Canvas.Font.Color := clGrayText;
  if Bool(Attrs and FILE_ATTRIBUTE_HIDDEN) then
    ListView.Canvas.Font.Style :=
       ListView.Canvas.Font.Style + [fsStrikeOut];
  if Bool(Attrs and FILE_ATTRIBUTE_SYSTEM) then
    Listview.Canvas.Font.Color := clHighlight;
end;

procedure TfrmAddFiles.ListViewCustomDrawSubItem(Sender: TCustomListView;
  Item: TListItem; SubItem: Integer; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin
  if SubItem = 0 then Exit;
  ListView.Canvas.Font.Color := GetSysColor(COLOR_WINDOWTEXT);
  //workaround for Win98 bug.
end;

procedure TfrmAddFiles.btnBackClick(Sender: TObject);
var
  Temp: PItemIDList;
begin
  Temp := CopyPIDL(FPIDL);
  if Assigned(Temp) then
    StripLastID(Temp);

  if Temp.mkid.cb <> 0 then begin
    SetPath(Temp);
    if ListView.ViewStyle <> vsReport then
        SetPath(Temp); // I don't know why! This can get correct icons
  end
  else
    Beep;

  ListView.ClearSelection; // added by me
end;

procedure TfrmAddFiles.btnAddClick(Sender: TObject);
var
  sPath: string;
  sPathFilename: string;
  i: integer;
begin
  if ListView.SelCount = 0 then begin
      beep;
      MessageDlg('Selected nothing. Please try again.', mtError, [mbOK], 0);
      exit;
  end;

  if sComment <> '' then begin
      frmMain.ZipMaster1.ZipComment := sComment;
      frmMain.ZipMaster1.List;
  end;
  
  SettingsBeforeAdd; // initial all settings

  for i := 0 to ListView.Items.Count -1 do begin
    sPathFilename := '';
    sPath := '';
    lblThisFile.Caption := '';

    if ListView.Items.Item[i].Selected then begin
        lblThisFile.Caption := ListView.Items[i].Caption;
        sPathFilename := sCurPath + lblThisFile.Caption;

        if DirectoryExists(sPathFilename) then begin
            sPath := IncludeTrailingPathDelimiter(sPathFilename);
            frmMain.ReadFilesOrFolders(nil, sPath);
        end
        else if FileExists(sPathFilename) then
            frmMain.lboFilesToZip.Items.Add(sPathFileName);

    end;
  end;
  Close;
end;

procedure TfrmAddFiles.SettingsBeforeAdd;
var
  zmOpts: AddOpts;
begin
  zmOpts := [];
  if chAddWithDir.Checked then begin
      //zmOpts := zmOpts + [AddRecurseDirs]; // Recursively find a specified filename 
      zmOpts := zmOpts + [AddDirNames];
  end;

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

procedure TfrmAddFiles.btnCancelAddClick(Sender: TObject);
begin
  bCancelProcess := True;
  Close;
end;

procedure TfrmAddFiles.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
var
  NewFlags: integer;
begin
  Main.bCancelOpenZip_Onaddfiles := bCancelProcess;
  if bCancelProcess then
      Self.Tag := 1
  else
      Self.Tag := 0;

  NewFlags := SHCONTF_FOLDERS or SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN;
  SetFlagsForListView(NewFlags);
  btnHideFolders.Down := False;

  if Self.WindowState = wsNormal then begin
      frmMain.WriteIniToReg(XcSubKeyIs, cSelfWidth, inttostr(Self.Width));
      frmMain.WriteIniToReg(XcSubKeyIs, cSelfHeight, inttostr(Self.Height));
  end;
end;

procedure TfrmAddFiles.FormDestroy(Sender: TObject);
begin
  ClearIDList;
  FIDList.Free;
end;

procedure TfrmAddFiles.chKeepAboveSetMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  OnOffchkKeepSettings;
end;

procedure TfrmAddFiles.OnOffchkKeepSettings;
var
  sChecked: string;
begin
  if chKeepAboveSet.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckKeepAboveSet, sChecked);
end;

procedure TfrmAddFiles.chKeepAboveSetKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 32 then begin
      chKeepAboveSet.Checked := not chKeepAboveSet.Checked;
      OnOffchkKeepSettings;
      chKeepAboveSet.Checked := not chKeepAboveSet.Checked;
  end;
end;

procedure TfrmAddFiles.chAddWithDirMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  OnOffchkAddWithDir;
end;

procedure TfrmAddFiles.OnOffchkAddWithDir;
var
  sChecked: string;
begin
  if chAddWithDir.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckAddWithDir, sChecked);
end;

procedure TfrmAddFiles.chAddWithDirKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 32 then begin
      chAddWithDir.Checked := not chAddWithDir.Checked;
      OnOffchkAddWithDir;
      chAddWithDir.Checked := not chAddWithDir.Checked;
  end;
end;

procedure TfrmAddFiles.OnOffchkOverwriteWithNew;
var
  sChecked: string;
begin
  if chOverwriteWithNew.Checked then
      sChecked := '1'
  else
      sChecked := '0';

  frmMain.WriteIniToReg(XcSubKeyIs, cCheckOverwriteWithNew, sChecked);
end;

procedure TfrmAddFiles.chOverwriteWithNewMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  OnOffchkOverwriteWithNew;
end;

procedure TfrmAddFiles.chOverwriteWithNewKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if Key = 32 then begin
      chOverwriteWithNew.Checked := not chOverwriteWithNew.Checked;
      OnOffchkOverwriteWithNew;
      chOverwriteWithNew.Checked := not chOverwriteWithNew.Checked;
  end;
end;

procedure TfrmAddFiles.cbCompressRatioChange(Sender: TObject);
var
  i: Shortint;
begin
  i := cbCompressRatio.ItemIndex;
  frmMain.WriteIniToReg(XcSubKeyIs, cIndexOfCompressRatio, inttostr(i));
end;

procedure TfrmAddFiles.RestoreSettingsOffrmAddFiles;
var
  sIndexOfCompress: string;
  iIndexOfCompress: smallint;
  sStr: string;
  i: integer;
  sOldPath: string;
  sSelfWidth, sSelfHeight: string;
begin
  for i := 0 to 9 do begin // assume nine history path written by Directory form
    sOldPath := frmMain.ReadIniFromReg(sSubKeyHistoryDir, inttostr(i));
    if sOldPath <> '' then  begin
        sOldPath := ExcludeTrailingPathDelimiter(sOldPath);
        cbPath.Items.Append(sOldPath);
    end;
  end;

  sIndexOfCompress := frmMain.ReadIniFromReg(XcSubKeyIs, cIndexOfCompressRatio);
  if sIndexOfCompress <> '' then begin
      iIndexOfCompress := strtoint(sIndexOfCompress);
      cbCompressRatio.ItemIndex := iIndexOfCompress;
  end;
  sStr := '';
  sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckKeepAboveSet);
  if sStr = '1' then
      chKeepAboveSet.Checked := True;

  sStr := '';
  sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckAddWithDir);
  if sStr = '0' then
      chAddWithDir.Checked := False;

  sStr := '';
  sStr := frmMain.ReadIniFromReg(XcSubKeyIs, cCheckOverwriteWithNew);
  if sStr = '0' then
      chOverwriteWithNew.Checked := False;

  sSelfWidth := frmMain.ReadIniFromReg(XcSubKeyIs, cSelfWidth);
  if sSelfWidth <> '' then
      Self.Width := strtoint(sSelfWidth);

  sSelfHeight := frmMain.ReadIniFromReg(XcSubKeyIs, cSelfHeight);
  if sSelfHeight <> '' then
      Self.Height := strtoint(sSelfHeight);
      
end;

procedure TfrmAddFiles.btnUnzipThisClick(Sender: TObject);
var
  sPathFilename: string;
  sPath: string;
begin
  sPathFilename := edPath.Text;
  if (sPathFilename <> '') and FileExists(sPathFilename) then begin
      sPath := ExtractFilePath(sPathFilename);
      sPath := ExcludeTrailingPathDelimiter(sPath);
      //sPath := IncludeTrailingPathDelimiter(sPath);
      //frmMain.WriteIniToReg(XcSubKeyIs, cPreviousDir, sPath);
      HistoryPathToReg(sPath);
      bCancelProcess := False;
  end
  else begin
      beep;
      bCancelProcess := True;
      Raise Exception.Create('File does not exist.');
  end;
  close;
end;

procedure TfrmAddFiles.btnCancelUnzipThisClick(Sender: TObject);
begin
  bCancelProcess := True;
  Close;
end;

procedure TfrmAddFiles.ListViewClick(Sender: TObject);
var
  sThisFile: string;
  sFileType: string;
  sFileExt: string;
begin
  if (Sender as TListView).Selected = nil then exit;

  lblFileExt2.Caption := (Sender as TListView).Selected.Caption;
  sThisFile := lblFileExt2.Caption;
  //if lowercase( ExtractFileExt(sThisFile) ) <> '.zip' then exit;

  sFileExt := copy(sThisFile, length(sThisFile) -4+1, 4);
  //sFileExt := Lowercase( ExtractFileExt(sThisFile) );
  sFileExt := lowercase(sFileExt);
  sFileType := ShellItem(ListView.Selected.Index).TypeName;
  sFileType := lowercase(sFileType);
  if (sFileExt <> '.zip') and (sFileType = 'zip') then
      sThisFile := sThisFile + '.zip';

  if DirectoryExists(sCurPath) then
      edPath.Text := sCurPath + sThisFile;

end;

procedure TfrmAddFiles.ListViewColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  if (Column.Index <> 0) and (Column.Index <> 3) then exit; // sort Name & Modified only

 { if Column.ImageIndex = 6 then begin
      Column.ImageIndex := 7;
      bDescending := False;
  end
  else if Column.ImageIndex = 7 then begin
      Column.ImageIndex := 6;
      bDescending := True;
  end
  else begin
      Column.ImageIndex := 6;
      bDescending := True;
  end; }
  bDescending := not bDescending;
  RestoreCaptionOfColumn;
  if not bDescending then begin
      if Column.Index = 0 then
          Column.Caption := 'Name  a - z'
      else if Column.Index = 3 then
          Column.Caption := 'Modified  r - o';
  end
  else begin
      if Column.Index = 0 then
          Column.Caption := 'Name  z - a'
      else if Column.Index = 3 then
          Column.Caption := 'Modified  o - r';
  end;
  iColBeSorted := Column.Index; // added by me
  FIDList.Sort(ListSortFunc);
  //We need to tell the ListView how many items it has.
  ListView.Items.Count := FIDList.Count;
  ListView.Repaint;
end;

procedure TfrmAddFiles.HistoryPathToReg(const sPath: string);
var
  i: integer;
  sPathNameInList: string;
begin
  // Is count over limit?
  repeat
    if cbPath.Items.Count > cCountOfHistoryPath then 
      //for i := cbPath.Items.Count -1 downto 0 do begin
        cbPath.Items.Delete(0);
        //if cbPath.Items.Count <= cCountOfHistoryPath then // be my count
        //    break;

      //end;
    //end;
  until cbPath.Items.Count <= cCountOfHistoryPath;
  // Is path already exist?
  for i := 0 to cbPath.Items.Count -1 do begin
    if lowercase(cbPath.Items.Strings[i]) = lowercase(sPath) then
        exit;

  end;
  // not full?
  if cbPath.Items.Count < cCountOfHistoryPath then
      cbPath.Items.Append(sPath)
  else begin // full
      cbPath.Items.Delete(0); // remove first item
      cbPath.Items.Append(sPath);
  end;
  // update to registry
  for i := cbPath.Items.Count - 1 downto 0 do begin
    sPathNameInList := cbPath.Items.Strings[i];
    frmMain.WriteIniToReg(sSubKeyHistoryDir, inttostr(i), sPathNameInList);
  end;
end;

procedure TfrmAddFiles.SetFlagsForListView(const NewFlags: integer);
begin
  CurFlags := NewFlags;
end;

procedure TfrmAddFiles.RestoreCaptionOfColumn;
begin
  ListView.Column[0].Caption := 'Name';
  ListView.Column[3].Caption := 'Modified';
end;

procedure TfrmAddFiles.SetFilterForListView(const NewFilter: string);
begin
  sCurFilter := NewFilter;
end;

procedure TfrmAddFiles.btnAllClick(Sender: TObject);
begin
  Self.SetFilterForListView('.*');
  Self.SetPath(sCurPath);
end;

procedure TfrmAddFiles.RefreshListView;
begin
  Self.SetPath(sCurPath);
end;

procedure TfrmAddFiles.btnShowZipsClick(Sender: TObject);
begin
  Self.SetFilterForListView('.zip');
  Self.SetPath(sCurPath);
end;

procedure TfrmAddFiles.FoundDataInListView;
begin
  if Self.tabUnzip.Visible then
      ListViewClick(ListView);
end;

procedure TfrmAddFiles.ListViewKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  FoundDataInListView;
end;

procedure TfrmAddFiles.btnHideFoldersClick(Sender: TObject);
var
  NewFlags: integer;
begin
  if btnHideFolders.Down then
      NewFlags := SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN
  else
      NewFlags := SHCONTF_FOLDERS or SHCONTF_NONFOLDERS or SHCONTF_INCLUDEHIDDEN;

  SetFlagsForListView(NewFlags);
  Self.RefreshListView;
end;

procedure TfrmAddFiles.cbPathDropDown(Sender: TObject);
begin
  cbPath.Text := '';
end;

procedure TfrmAddFiles.cbUpperDropDown(Sender: TObject);
begin
  cbUpper.Text := '';
end;

procedure TfrmAddFiles.cbUpperClick(Sender: TObject);
var
  I: Integer;
begin
  if cbUpper.Text = '' then exit;
  I := cbUpper.Items.IndexOf(cbUpper.Text);
  if cbUpper.Text[Length(cbUpper.Text)] = ':' then // added by me
      cbUpper.Text := cbUpper.Text + '\';

  if I >= 0 then begin
      if cbUpper.Items.Objects[I] <> nil then // added by me
          SetPath(PItemIDList(cbUpper.Items.Objects[I]))
      else
          SetPath(cbUpper.Text); // added by me

  end
  else
      SetPath(cbUpper.Text);
      
end;

procedure TfrmAddFiles.btnCommentClick(Sender: TObject);
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

function TfrmAddFiles.Restore_WinNT_ClassID_To_Name(sClassID: string): string;
var
  sResultName: string;
  iRet: integer;
  sTemp: string;
  sSendID: string;
  sRetName: string;
  sPrevRetName: string;
begin
  sTemp := sClassID; // e.g, ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} Win 2000 
  repeat
    iRet := AnsiPos('\', sTemp); // divide class id if more than one
    if iRet > 0 then begin
        sSendID := copy(sTemp, 1, iRet-1);
        sTemp := copy( sTemp, iRet+3, length(sTemp) - (iRet+2) );
    end;
    sRetName := Read_NullValue_In_ClassID_FromReg('CLSID\' + sTemp, '');
    if sRetName <> sPrevRetName then begin
        sResultName := sResultName + sRetName + '\';
        sPrevRetName := sRetName;
    end;
  until iRet = 0;
  Result := ExcludeTrailingPathDelimiter(sResultName);
end;

function TfrmAddFiles.Read_NullValue_In_ClassID_FromReg(sSubKey: string; sValueName: string):
 string;
begin
  Result := '';

  with TRegistry.Create do
  try
    RootKey := HKEY_CLASSES_ROOT;
    if OpenKey(sSubKey, False) then begin
        Result := ReadString(sValueName);

        CloseKey;
    end;
  finally
    Free;
  end;
end;

procedure TfrmAddFiles.ToolButton2Click(Sender: TObject);
var
  NewPIDL: PItemIDList;
begin
  cbUpper.Clear;

  OLECheck(
    SHGetSpecialFolderLocation(
      Application.Handle,
      CSIDL_DRIVES,
      NewPIDL)
  );
  SetPath(NewPIDL);
  if ListView.ViewStyle <> vsReport then // I don't know why! This can get correct icons
      SetPath(NewPIDL);
      
end;



procedure TfrmAddFiles.FormActivate(Sender: TObject);
begin
  bCancelProcess := True;
  bCancelOpenZip_Onaddfiles := True;
end;

end.
