unit SearchBox;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, VirtualTrees;

type
  TSearch = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    edSearch: TEdit;
    lblMessage: TLabel;
    btnFind: TButton;
    btnCancel: TButton;
    GroupBox1: TGroupBox;
    chkCaseSensitive: TCheckBox;
    chkWholeWords: TCheckBox;
    btnFindNext: TButton;
    procedure btnFindClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure edSearchKeyPress(Sender: TObject; var Key: Char);
    procedure edSearchClick(Sender: TObject);
    procedure btnFindNextClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure edSearchChange(Sender: TObject);
  private
    { Private declarations }
    procedure StartSearch(iFromPos: integer);
  public
    { Public declarations }
  end;

var
  Search: TSearch;

 { procedure SetWindowPos(Wnd:HWnd;
     hWndInsertAfter: Longint; X: Longint;
      Y: Longint; cx: Longint; cy: Longint;
       DWORD:Word) stdcall; }

implementation

uses
  Main, ZipMstr, strUtils;

var
  iLastSelectedPos: integer = 0;
  LastNode: PVirtualNode;
{$R *.dfm}

procedure SetWindowPos; external 'USER32';

procedure TSearch.StartSearch(iFromPos: integer);
var
  sSearchThis, sFileName: string;
  bFound: boolean;
  Node: PVirtualNode;
  Data: PEntry;
begin
  lblMessage.Caption := '';
  sSearchThis := edSearch.Text;
  if not chkCaseSensitive.Checked then // not case sensitive
      sSearchThis := AnsiLowercase(sSearchThis);

  if sSearchThis = '' then exit;
  if frmMain.ZipFName.Caption = '' then exit;

 { if iFromPos > (frmMain.ZipMaster1.ZipContents.Count -1) then begin // eof
      lblMessage.Caption := 'End of file.';
      btnFindNext.Enabled := False;
      exit;
  end; }

  if frmMain.btnShowVirtual.Down then begin

  end
  else begin
      if iFromPos = 0 then
          Node := frmMain.VT.GetFirst
      else begin
          if Assigned(LastNode) then
              Node := LastNode
          else begin
              beep;
              lblMessage.Caption := 'End of file.';
              iLastSelectedPos := 0;
              btnFindNext.Enabled := False;
              exit;
          end;
      end;

      bFound := False;
      while Assigned(Node) do begin
        Data := frmMain.VT.GetNodeData(Node);
        sFileName := Data.Value[0]; // Name
        sFileName := AnsiReplaceStr(sFileName, '*', ''); // remove password char
        if not chkCaseSensitive.Checked then // not case sensitive
            sFileName := AnsiLowercase(sFileName);

        if chkWholeWords.Checked then begin
            if sSearchThis = sFileName then
                bFound := True;

        end
        else begin
           if AnsiPos(sSearchThis, sFileName) > 0 then
               bFound := True;

        end;

        if bFound then begin
            frmMain.VT.ClearSelection;
            frmMain.VT.Selected[Node] := True;
            frmMain.VT.InvalidateNode(Node);
            frmMain.VT.ScrollIntoView(Node, True, False);
            LastNode := frmMain.VT.GetNext(Node);
            btnFindNext.Enabled := True;
            exit;
        end
        else
            Node := frmMain.VT.GetNext(Node);

        if Node = nil then begin
            beep;
            lblMessage.Caption := 'End of file.';
            iLastSelectedPos := 0;
            btnFindNext.Enabled := False;
            exit;
        end;
      end;
  end;

 { bFound := False;
  for i := iFromPos to frmMain.ZipMaster1.ZipContents.Count -1 do begin
    with ZipDirEntry(frmMain.ZipMaster1.ZipContents[i]^) do begin
        sFilenameIs := ExtractFileName(Filename);
        if not chkCaseSensitive.Checked then // not case sensitive
            sFilenameIs := AnsiLowercase(sFilenameIs);

        if chkWholeWords.Checked then begin // whole words only
            if sSearchThis = sFilenameIs then
                bFound := True;

        end
        else begin // not whole words
            if AnsiPos(sSearchThis, sFilenameIs) > 0 then
                bFound := True;

        end;

        if bFound then begin // found the file!
            //frmMain.ListView1.ClearSelection; // clear all selected files first
            frmMain.VT.ClearSelection;
            if frmMain.btnShowVirtual.Down then begin // virtual list view
                sPath := ExtractFileDir(Filename);
                // change folder on TreeView
                if frmMain.ChangeFolder_OnVirtualTree_After_SearchFile(sPath) then begin // success
                    iCntCol := frmMain.ListView1.Columns.Count;
                    sZipFolder := '';
                    sPathFName := Filename;
                    // find the filename on ListView
                    for j := 0 to frmMain.ListView1.Items.Count -1 do begin
                      sCap := frmMain.ListView1.Items[j].Caption; // File name
                      sCap := AnsiReplaceStr(sCap, '*', ''); // remove password char
                      sZipFolder := frmMain.ListView1.Items[j].SubItems.Strings[iCntCol -2]; // Folder name
                      if (sZipFolder <> '') then
                          sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);
                      // Is the selected item not a folder
                      if (frmMain.ListView1.Items.Item[j].SubItems.Strings[0] <> '') and
                      (frmMain.ListView1.Items.Item[j].SubItems.Strings[1] <> '') then begin
                          if sPathFName = (sZipFolder + sCap) then begin // found !
                              iLastSelectedPos := i; // last pos from ZipMaster1
                              btnFindNext.Enabled := True;
                              frmMain.ListView1.Items[j].Selected := True;
                              break;
                          end;
                      end;
                    end;
                end;
            end
            else begin // normal listview
                // the file on list is same index with content of ZipMaster1
                if AnsiLowercase(sFilenameIs) = AnsiLowercase(frmMain.ListView1.Items[i].Caption) then begin
                    iLastSelectedPos := i;
                    btnFindNext.Enabled := True;
                    frmMain.ListView1.Items[i].Selected := True
                end
                else begin // try to find the file from item 0
                    iCntCol := frmMain.ListView1.Columns.Count;
                    sZipFolder := '';
                    for j := 0 to frmMain.ListView1.Items.Count -1 do begin
                      sCap := frmMain.ListView1.Items[j].Caption; // File name
                      sCap := AnsiReplaceStr(sCap, '*', ''); // remove password char
                      sZipFolder := frmMain.ListView1.Items[j].SubItems.Strings[iCntCol -2]; // Folder name
                      if (sZipFolder <> '') then
                          sZipFolder := IncludeTrailingPathDelimiter(sZipFolder);

                      if Filename = (sZipFolder + sCap) then begin // found !
                          iLastSelectedPos := j;
                          btnFindNext.Enabled := True;
                          frmMain.ListView1.Items[j].Selected := True;
                          break;
                      end;
                    end;
                end;
            end;

            if frmMain.ListView1.Selected <> nil then // got it!
                frmMain.ListView1.Selected.MakeVisible(False);;
                
            frmMain.ListView1.SetFocus;
            break;
        end;
    end;
  end;

  if not bFound then begin
      lblMessage.Caption := 'File not found.';
      iLastSelectedPos := 0;
      btnFindNext.Enabled := False;
  end; }
end;

procedure TSearch.btnFindClick(Sender: TObject);
begin
  frmMain.VT.ClearSelection;
  iLastSelectedPos := 0;
  StartSearch(0);
end;

procedure TSearch.btnCancelClick(Sender: TObject);
begin
  self.Close;
end;

procedure TSearch.FormCreate(Sender: TObject);
begin
  self.Icon := Application.Icon;
  self.Caption := XcProgramName + '...Search';
  edSearch.Text := '';
  lblMessage.Caption := '';
  self.Top := frmMain.Top + 10;
  self.Left := frmMain.Left + 10;
  //SetWindowPos(self.handle, -1, 0, 0, 0, 0, 3);
end;

procedure TSearch.edSearchKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      btnFind.Click;
      Key := #0;
  end;
end;

procedure TSearch.edSearchClick(Sender: TObject);
begin
  edSearch.SelectAll;
end;

procedure TSearch.btnFindNextClick(Sender: TObject);
begin
  StartSearch(iLastSelectedPos +1);
end;

procedure TSearch.FormShow(Sender: TObject);
begin
  iLastSelectedPos := 0;
  lblMessage.Caption := '';
  btnFindNext.Enabled := False;
  LastNode := nil;
  edSearch.SetFocus;
end;

procedure TSearch.edSearchChange(Sender: TObject);
begin
  iLastSelectedPos := 0;
  lblMessage.Caption := '';
  btnFindNext.Enabled := False;
end;

end.
