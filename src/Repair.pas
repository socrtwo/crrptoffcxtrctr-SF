unit Repair;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, StdCtrls, Buttons, ImgList, JvBrowseFolder;

type
  TfrmRepair = class(TForm)
    lvRepair: TListView;
    pnBottom: TPanel;
    edRepairArchive: TEdit;
    btnOpenArchive: TBitBtn;
    lblRepairArchive: TLabel;
    lblFixedArchive: TLabel;
    edFixedArchive: TEdit;
    btnOpenFixed: TBitBtn;
    btnStart: TButton;
    btnClose: TBitBtn;
    ilRepair: TImageList;
    pnTop: TPanel;
    lblDisclaimer: TLabel;
    OpenDlg: TOpenDialog;
    procedure FormCreate(Sender: TObject);
    procedure btnOpenArchiveClick(Sender: TObject);
    procedure edRepairArchiveChange(Sender: TObject);
    procedure btnOpenFixedClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
  private
    { Private declarations }
    OKCnt, TotalCnt: integer;
    procedure OnZipFileFound(Sender: TObject; const Filename: string;
     FileInfoIsOK: boolean);
  public
    { Public declarations }
  end;

var
  frmRepair: TfrmRepair;

implementation

uses
  Main, ZipFix, filectrl;

{$R *.dfm}

procedure TfrmRepair.FormCreate(Sender: TObject);
begin
  self.Caption := XcProgramName + '...Repairing';
end;

procedure TfrmRepair.btnOpenArchiveClick(Sender: TObject);
//var
//  sFileExt: string;
begin
  if OpenDlg.Execute then begin
     { sFileExt := ExtractFileExt(OpenDlg.FileName);
      sFileExt := AnsiLowercase(sFileExt);
      if sFileExt <> '.zip' then begin
          beep;
          MessageDlg('This is not a zip file.', mtError, [mbOK], 0);
          exit;
      end; }
      edRepairArchive.Text := OpenDlg.FileName;
  end;
end;

procedure TfrmRepair.edRepairArchiveChange(Sender: TObject);
var
  sExt: string;
begin
  sExt := ExtractFileExt(edRepairArchive.Text);
  edFixedArchive.Text := ChangeFileExt(edRepairArchive.Text, '_fixed' + sExt);
end;

procedure TfrmRepair.btnOpenFixedClick(Sender: TObject);
var
  sExt, sFilename, sOutPath: string;
  JvBrowser: TJvBrowseForFolderDialog;
begin
  JvBrowser := TJvBrowseForFolderDialog.Create(nil);
  try
    JvBrowser.Title := 'Select Folder...';

    if JvBrowser.Execute then begin
        sOutPath := JvBrowser.Directory;
        if sOutPath <> '' then
            sOutPath := IncludeTrailingPathDelimiter(sOutPath);

        sFilename := ExtractFilename(edRepairArchive.Text);
        sExt := ExtractFileExt(edRepairArchive.Text);
        if sFilename <> '' then
            sFilename := ChangeFileExt(sFilename, '_fixed' + sExt);

        edFixedArchive.Text := sOutPath + sFilename;

    end;
  finally
    JvBrowser.Free;
  end;
end;

procedure TfrmRepair.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmRepair.btnStartClick(Sender: TObject);
var
  sFilename, sWhatDrive, sFrom, sTo: string;
  ListItem: TListItem;
  bNoRetErr: boolean;
  iFreeSpace: Cardinal;
  iFileSize: integer;
begin
  sFrom := edRepairArchive.Text;
  sTo := edFixedArchive.Text;
  edRepairArchive.Text := Trim(sFrom); // remove invalid leading or trailing spaces
  edFixedArchive.Text := Trim(sTo);

  if not FileExists(edRepairArchive.Text) then begin
      beep;
      MessageDlg('Repairing archive not found.', mtError, [mbOK], 0);
      exit;
  end;

  if not DirectoryExists( ExtractFileDir(edFixedArchive.Text) ) then begin
      beep;
      MessageDlg('Target directory to store fixed archive not exist.',
      mtError, [mbOK], 0);
      exit;
  end;

  sFilename := ExtractFilename(edFixedArchive.Text);
  if sFilename = '' then begin
      beep;
      MessageDlg('The output filename of fixed archive is empty.',
      mtError, [mbOK], 0);
      exit;
  end;

  if FileExists(edFixedArchive.Text) then begin
      if MessageDlg('The output file of fixed archive already exists. Do you ' +
      'want to overwrite?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
          exit;
          
  end;

  if edRepairArchive.Text = edFixedArchive.Text then begin
      beep;
      MessageDlg('The output filename of Fixed archive is same with Repairing ' +
      'archive. Please use a new filename on Fixed archive text box.',
      mtError, [mbOK], 0);
      exit;
  end;

  sWhatDrive := ExtractFileDrive(edFixedArchive.Text);
  if sWhatDrive <> '' then
      sWhatDrive := IncludeTrailingPathDelimiter(sWhatDrive);

  iFreeSpace := frmMain.FreeDiskSpace(sWhatDrive, bNoRetErr);
  iFileSize := frmMain.Get_FileSize(edRepairArchive.Text);

  if bNoRetErr then begin
      if iFreeSpace < iFileSize then begin
          beep;
          MessageDlg('Insufficient disk space on target location.',
          mtError, [mbOK], 0);
          exit;
      end;
  end;

  OKCnt := 0;
  TotalCnt := 0;
  lvREpair.Clear;
  with TZipFix.create(self) do begin
    try
      OnFileFound := self.OnZipFileFound;
      Execute(edRepairArchive.Text, edFixedArchive.Text);
      //display totals...
      ListItem := lvRepair.Items.Add;
      ListItem.ImageIndex := -1;
      ListItem.Caption := '';

      ListItem := lvRepair.Items.Add;
      ListItem.ImageIndex := 2;
      ListItem.Caption := 'Total Files: ' + inttostr(TotalCnt) ;

      ListItem := lvRepair.Items.Add;
      ListItem.ImageIndex := 2;
      //ListItem.Caption := Format('%d out of %d files are OK',[OKCnt, TotalCnt]);
      ListItem.Caption := inttostr(OKCnt) + ' files are OK';
      ListItem.MakeVisible(False);
    finally
      free;
    end;
  end;
end;

procedure TfrmRepair.OnZipFileFound(Sender: TObject; const Filename: string;
 FileInfoIsOK: boolean);
var
  s: string;
  ListItem: TListItem;
begin
  inc(TotalCnt);
  if Filename = '' then
      s := '??filename??'
  else
      s := Filename;

  ListItem := lvRepair.Items.Add;
  if FileInfoIsOK then begin
      inc(OKCnt);
      ListItem.ImageIndex := 0;
  end
  else
      ListItem.ImageIndex := 1;

  ListItem.Caption := s;
end;
end.
