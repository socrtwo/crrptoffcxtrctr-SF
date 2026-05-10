unit AddItemsThr;

interface

uses
  Classes, Windows, ZipMstr, ComCtrls, SysUtils, StrUtils, ShellAPI;

type
  TAdd_Items_Thread = class(TThread)
  private
    { Private declarations }
    bCheckType: boolean;
    bCheckCRC: boolean;
    bCheckAttributes: boolean;
    FZM: TZipMaster;
    FLV: TListView;
    F_Start: integer;
    F_End: integer;
  protected
    procedure Execute; override;
    procedure AddItems(iStart: integer; iEnd: integer); virtual; abstract;
    function Get_SysIconIndex_Of_Given_FileExt(sWhatFileExt: string;
     var sFileTypeIs: string): integer;
    function GetAttributes(iFileAttri: Cardinal): string;
  public
    constructor Create(ZM: TZipMaster; LV: TListView; iSS: integer;
     iEE: integer; bFileType, bCRC, bAttr: boolean);
  end;

  { TAddItems_One }

  TAddItems_One = class(TAdd_Items_Thread)
  protected
    procedure AddItems(iStart: integer; iEnd: integer); override;
  end;

  { TAddItems_Two }

  TAddItems_Two = class(TAdd_Items_Thread)
  protected
    procedure AddItems(iStart: integer; iEnd: integer); override;
  end;

  { TAddItems_Three }

  TAddItems_Three = class(TAdd_Items_Thread)
  protected
    procedure AddItems(iStart: integer; iEnd: integer); override;
  end;

implementation


var
  sPreExt: string;
  iPreImgI: integer;
{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TAdd_Items_Thread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TAdd_Items_Thread }

constructor TAdd_Items_Thread.Create(ZM: TZipMaster; LV: TListView; iSS: integer;
iEE: integer; bFileType, bCRC, bAttr: boolean);
begin
  FZM := ZM;
  FLV := LV;
  F_Start := iSS;
  F_End := iEE;
  bCheckType := bFileType;
  bCheckCRC := bCRC;
  bCheckAttributes := bAttr;
  FreeOnTerminate := True;
  inherited Create(False);
end;

procedure TAdd_Items_Thread.Execute;
begin
  { Place thread code here }
  AddItems(F_Start, F_End);
end;

{ TAddItems_One }
procedure TAddItems_One.AddItems(iStart: integer; iEnd: integer);
var
  i, iIndex: integer;
  ListItem: TListItem;
  sFileExt, sPath, sFile_Type: string;
begin
  for i := iStart to iEnd do begin
      with ZipDirEntry(FZM.ZipContents[i]^) do begin
        ListItem := FLV.Items.Add;
        if Encrypted then
            ListItem.Caption := ExtractFileName(FileName) + '*' // Name
        else
            ListItem.Caption := ExtractFileName(FileName); // Name

        //ListItem.ImageIndex := -1;
       // bSameFileType := False;
        sPath := ExtractFilePath(FileName);
        //sFilename := lowercase(FileName);
        sFileExt := Uppercase( ExtractFileExt(FileName) );

       { if sFilename = 'setup.exe' then
            ListItem.ImageIndex := 3
        else if (sFilename = 'readme.txt') or (sFilename = 'readme1st.txt') then
            ListItem.ImageIndex := 4
        else } if sPreExt = sFileExt then begin
            ListItem.ImageIndex := iPreImgI;
           // bSameFileType := True;
        end
        else begin
            iIndex := Get_SysIconIndex_Of_Given_FileExt(sFileExt, sFile_Type);
            ListItem.ImageIndex := iIndex;
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


        ListItem.SubItems.Add(IntToStr(CompressedSize)); // Packed

        if bCheckType then begin // Type
            if sFile_Type <> '' then
                ListItem.SubItems.Add(sFile_Type)
            else begin
                sFile_Type := AnsiReplaceStr(sFileExt, '.', ' ');
                ListItem.SubItems.Add(sFile_Type);
            end;
        end;

        if bCheckCRC then  // CRC
            ListITem.SubItems.Add(inttohex(CRC32, 2));

        if bCheckAttributes then  // Attributes
            ListITem.SubItems.Add(GetAttributes(ExtFileAttrib));

        ListITem.SubItems.Add(sPath); //Path

        //Gauge1.Progress := i;
      end;
  end;
end;

{ TAddItems_Two }
procedure TAddItems_Two.AddItems(iStart: integer; iEnd: integer);
var
  i, iIndex: integer;
  ListItem: TListItem;
  sFileExt, sPath, sFile_Type: string;
begin
  for i := iStart to iEnd do begin
      with ZipDirEntry(FZM.ZipContents[i]^) do begin
        ListItem := FLV.Items.Add;
        if Encrypted then
            ListItem.Caption := ExtractFileName(FileName) + '*' // Name
        else
            ListItem.Caption := ExtractFileName(FileName); // Name

        //ListItem.ImageIndex := -1;
       // bSameFileType := False;
        sPath := ExtractFilePath(FileName);
        //sFilename := lowercase(FileName);
        sFileExt := Uppercase( ExtractFileExt(FileName) );

       { if sFilename = 'setup.exe' then
            ListItem.ImageIndex := 3
        else if (sFilename = 'readme.txt') or (sFilename = 'readme1st.txt') then
            ListItem.ImageIndex := 4
        else } if sPreExt = sFileExt then begin
            ListItem.ImageIndex := iPreImgI;
           // bSameFileType := True;
        end
        else begin
            iIndex := Get_SysIconIndex_Of_Given_FileExt(sFileExt, sFile_Type);
            ListItem.ImageIndex := iIndex;
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


        ListItem.SubItems.Add(IntToStr(CompressedSize)); // Packed

        if bCheckType then begin // Type
            if sFile_Type <> '' then
                ListItem.SubItems.Add(sFile_Type)
            else begin
                sFile_Type := AnsiReplaceStr(sFileExt, '.', ' ');
                ListItem.SubItems.Add(sFile_Type);
            end;
        end;

        if bCheckCRC then  // CRC
            ListITem.SubItems.Add(inttohex(CRC32, 2));

        if bCheckAttributes then  // Attributes
            ListITem.SubItems.Add(GetAttributes(ExtFileAttrib));

        ListITem.SubItems.Add(sPath); //Path

        //Gauge1.Progress := i;
      end;
  end;
end;

{ TAddItems_Three }
procedure TAddItems_Three.AddItems(iStart: integer; iEnd: integer);
var
  i, iIndex: integer;
  ListItem: TListItem;
  sFileExt, sPath, sFile_Type: string;
begin
  for i := iStart to iEnd do begin
      with ZipDirEntry(FZM.ZipContents[i]^) do begin
        ListItem := FLV.Items.Add;
        if Encrypted then
            ListItem.Caption := ExtractFileName(FileName) + '*' // Name
        else
            ListItem.Caption := ExtractFileName(FileName); // Name

        //ListItem.ImageIndex := -1;
       // bSameFileType := False;
        sPath := ExtractFilePath(FileName);
        //sFilename := lowercase(FileName);
        sFileExt := Uppercase( ExtractFileExt(FileName) );

       { if sFilename = 'setup.exe' then
            ListItem.ImageIndex := 3
        else if (sFilename = 'readme.txt') or (sFilename = 'readme1st.txt') then
            ListItem.ImageIndex := 4
        else } if sPreExt = sFileExt then begin
            ListItem.ImageIndex := iPreImgI;
           // bSameFileType := True;
        end
        else begin
            iIndex := Get_SysIconIndex_Of_Given_FileExt(sFileExt, sFile_Type);
            ListItem.ImageIndex := iIndex;
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


        ListItem.SubItems.Add(IntToStr(CompressedSize)); // Packed

        if bCheckType then begin // Type
            if sFile_Type <> '' then
                ListItem.SubItems.Add(sFile_Type)
            else begin
                sFile_Type := AnsiReplaceStr(sFileExt, '.', ' ');
                ListItem.SubItems.Add(sFile_Type);
            end;
        end;

        if bCheckCRC then  // CRC
            ListITem.SubItems.Add(inttohex(CRC32, 2));

        if bCheckAttributes then  // Attributes
            ListITem.SubItems.Add(GetAttributes(ExtFileAttrib));

        ListITem.SubItems.Add(sPath); //Path

        //Gauge1.Progress := i;
      end;
  end;
end;

function TAdd_Items_Thread.Get_SysIconIndex_Of_Given_FileExt(sWhatFileExt: string;
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

function TAdd_Items_Thread.GetAttributes(iFileAttri: Cardinal): string;
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
end.
