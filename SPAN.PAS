unit span;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons;

type
  TDiskSpan = class(TForm)
    GroupBox5: TGroupBox;
    Label1: TLabel;
    edLocation: TEdit;
    btnSource: TButton;
    GroupBox1: TGroupBox;
    MaxVolSizeEdit: TEdit;
    Label2: TLabel;
    MinFreeVolEdit: TEdit;
    Label3: TLabel;
    btnStart: TBitBtn;
    BitBtn1: TBitBtn;
    Label4: TLabel;
    Label5: TLabel;
    lblCurZip: TLabel;
    lblFileSize: TLabel;
    Label7: TLabel;
    edFreeSpace: TEdit;
    Label6: TLabel;
    procedure btnSourceClick(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure MaxVolSizeEditKeyPress(Sender: TObject; var Key: Char);
    procedure MaxVolSizeEditExit(Sender: TObject);
    procedure MinFreeVolEditKeyPress(Sender: TObject; var Key: Char);
    procedure MinFreeVolEditExit(Sender: TObject);
    procedure edLocationChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure edFreeSpaceKeyPress(Sender: TObject; var Key: Char);
    procedure edFreeSpaceExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DiskSpan: TDiskSpan;

implementation

uses
  Main, FileCtrl;

{$R *.dfm}

procedure TDiskSpan.btnSourceClick(Sender: TObject);
var
  sPath: string;
begin
  sPath := '';
  if SelectDirectory('Select Directory', '', sPath) then
      edLocation.Text := sPath;
      
end;

procedure TDiskSpan.btnStartClick(Sender: TObject);
var
  sInFile, sOutFile: String;
  //iRetSuc: integer;
  iMaxVal, iMinVal, iFreeVal: integer;
begin
   sInFile := frmMain.ZipMaster1.ZipFilename;
   if sInFile = '' then begin
       beep;
       MessageDlg('There has not opened zip file.', mtError, [mbOK], 0);
       Exit;
   end;

   if not DirectoryExists(edLocation.Text) then begin
       beep;
       MessageDlg('Target location to save files not exist.', mtError, [mbOK], 0);
       exit; 
   end;

   sOutFile := edLocation.Text; // set path first
   if sOutFile <> '' then
       sOutFile := IncludeTrailingPathDelimiter(sOutFile);
       
   iMaxVal := strtoint(MaxVolSizeEdit.Text);
   iMinVal := strtoint(MinFreeVolEdit.Text);
   iFreeVal := strtoint(edFreeSpace.Text);

   frmMain.ZipMaster1.MaxVolumeSize := iMaxVal * 1000;
   frmMain.ZipMaster1.MinFreeVolumeSize := iMinVal * 1000;
   frmMain.ZipMaster1.KeepFreeOnDisk1 := iFreeVal * 1000;

   sOutFile := sOutFile + ExtractFileName(sInFile);
   frmMain.ZipMaster1.WriteSpan(sInFile, sOutFile);
  { if iRetSuc <> 0 then begin
       beep;
       MessageDlg('Failing to create span. Error code: ' + inttostr(iRetSuc),
       mtError, [mbOK], 0);
       exit;
   end; }

   Close;
end;

procedure TDiskSpan.MaxVolSizeEditKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then exit;
  if not ((Key >= '0') and (Key <= '9') or (Key = chr(vk_Back))) then begin
      Key := #0;
      beep;
  end;
end;

procedure TDiskSpan.MaxVolSizeEditExit(Sender: TObject);
begin
  if MaxVolSizeEdit.Text = '' then
      MaxVolSizeEdit.Text := '65'
  else if strtoint(MaxVolSizeEdit.Text) < 65 then
      MaxVolSizeEdit.Text := '65';

end;

procedure TDiskSpan.MinFreeVolEditKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then exit;
  if not ((Key >= '0') and (Key <= '9') or (Key = chr(vk_Back))) then begin
      Key := #0;
      beep;
  end;
end;

procedure TDiskSpan.MinFreeVolEditExit(Sender: TObject);
begin
  if MinFreeVolEdit.Text = '' then
      MinFreeVolEdit.Text := '0';
end;

procedure TDiskSpan.edLocationChange(Sender: TObject);
begin
  if DirectoryExists(edLocation.Text) then
      btnStart.Enabled := True
  else
      btnStart.Enabled := False;
      
end;

procedure TDiskSpan.FormCreate(Sender: TObject);
begin
  lblCurZip.Caption := 'Zip file is ' + frmMain.ZipMaster1.ZipFileName;
  lblFileSize.Caption := 'Size = ' + inttostr(frmMain.ZipMaster1.ZipFileSize
  div 1024) + ' KB. ' + 'Files = ' + inttostr(frmMain.ZipMaster1.Count);
  Self.Caption := XcProgramName + '...Disk Span';
  edLocation.Text := GetCurrentDir;
end;

procedure TDiskSpan.edFreeSpaceKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then exit;
  if not ((Key >= '0') and (Key <= '9') or (Key = chr(vk_Back))) then begin
      Key := #0;
      beep;
  end;
end;

procedure TDiskSpan.edFreeSpaceExit(Sender: TObject);
begin
  if edFreeSpace.Text = '' then
      edFreeSpace.Text := '0';

  if strtoint(edFreeSpace.Text) > 1024 then
      edFreeSpace.Text := '1024';

end;

end.
