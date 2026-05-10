unit OptAddFl;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons;

type
  TOptionsOfPopupAdd = class(TForm)
    gb1: TGroupBox;
    chAddWithDir: TCheckBox;
    chOverwriteWithNew: TCheckBox;
    Label1: TLabel;
    cbCompressRatio: TComboBox;
    lblCompressRatio: TLabel;
    btnOK: TBitBtn;
    btnCancel: TBitBtn;
    btnComment: TButton;
    chPassword: TCheckBox;
    chIncludeHidden: TCheckBox;
    lblWildcards: TLabel;
    edWildcards: TEdit;
    btnAddWithWildcards: TBitBtn;
    chkExcludeWildcards: TCheckBox;
    chIncludeSystem: TCheckBox;
    procedure btnOKClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCommentClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure edWildcardsChange(Sender: TObject);
    procedure btnAddWithWildcardsClick(Sender: TObject);
    procedure chkExcludeWildcardsClick(Sender: TObject);
  private
    { Private declarations }
    //procedure ChangeLanguage;
    procedure RestoreSettingsOfThisForm;
    procedure BtnAddFiles_ToArchive_Click(bWithWildcards: boolean);
  public
    { Public declarations }
  end;

var
  OptionsOfPopupAdd: TOptionsOfPopupAdd;

implementation

{$R *.dfm}
uses
  Main, InputMsgBox, ZipMstr, StrUtils;

var
  sComment: string = '';

procedure TOptionsOfPopupAdd.FormCreate(Sender: TObject);
begin
  cbCompressRatio.Items.Insert(0, 'Maximum');
  cbCompressRatio.Items.Insert(1, 'Normal');
  cbCompressRatio.Items.Insert(2, 'Low');
  cbCompressRatio.Items.Insert(3, 'None');

  cbCompressRatio.ItemIndex := 0;
  RestoreSettingsOfThisForm;
  //if XbNewLanguage then
  //    ChangeLanguage;

end;

procedure TOptionsOfPopupAdd.btnOKClick(Sender: TObject);
begin
  BtnAddFiles_ToArchive_Click(False);
end;

procedure TOptionsOfPopupAdd.btnCancelClick(Sender: TObject);
begin
  close;
end;

procedure TOptionsOfPopupAdd.RestoreSettingsOfThisForm;
const
  cCheckAddWithDir = 'CheckAddWithDir';  // refer to addfiles
  cCheckOverwriteWithNew = 'CheckOverwriteWithNew'; // refer to addfiles
  cIndexOfCompressRatio = 'IndexOfCompressRatio'; // refer to addfiles
  cCheckPassword = 'CheckPassword'; // refer to addfiles
  cCheckIncludeSystem = 'CheckIncludeSystem'; // refer to addfiles
  cCheckIncludeHidden = 'CheckIncludeHidden'; // refer to addfiles
var
  sIndex: string;
  iIndex: smallint;
  sStr: string;
begin
  // Get original settings from addfiles form
  sIndex := frmMain.ReadIniFromReg(XcSubKeyIs, cIndexOfCompressRatio);
  if sIndex <> '' then begin
      iIndex := strtoint(sIndex);
      cbCompressRatio.ItemIndex := iIndex;
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
  // ^^^

end;
{procedure TOptionsOfPopupAdd.ChangeLanguage;
var
  iComboIndex: shortint;
begin
  Self.Caption := sOptAdd_Caption;
  chAddWithDir.Caption := sDirTabAddchk_AddFilesWithDir;
  chOverwriteWithNew.Caption := sDirTabAddchk_OverwriteIfNewer;
  Label1.Caption := sDirTabAdd_nocheckIsalwaysoverwrite;
  lblCompressRatio.Caption := sDirTabAdd_CompressionLevel;
  chKeepAboveSet.Caption := sDirTabAddchk_KeepSettings;
  btnOK.Caption := sBtnOK;
  btnCancel.Caption := sBtnCancel;

  iComboIndex := cbCompressRatio.ItemIndex;
  cbCompressRatio.Clear;
  cbCompressRatio.Items.Append(sCombo_HighlyCompress);
  cbCompressRatio.Items.Append(sCombo_Normal);
  cbCompressRatio.Items.Append(sCombo_SlightlyCompress);
  cbCompressRatio.Items.Append(sCombo_None);
  cbCompressRatio.ItemIndex := iComboIndex;
end; }


procedure TOptionsOfPopupAdd.btnCommentClick(Sender: TObject);
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

procedure TOptionsOfPopupAdd.FormActivate(Sender: TObject);
begin
  bCancelAddFiles := True;
end;

procedure TOptionsOfPopupAdd.edWildcardsChange(Sender: TObject);
begin
  btnAddWithWildcards.Enabled := (edWildcards.Text <> '');
  btnOK.Enabled := not btnAddWithWildcards.Enabled;
end;

procedure TOptionsOfPopupAdd.BtnAddFiles_ToArchive_Click(bWithWildcards: boolean);
var
  zmOpts: AddOpts;
  j, k, iA, iRetPos: integer;
  sTemp, sWildcard, sFileExt: string;
  sArrayCatFileExt: array of string;
  bFound: boolean;
begin
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

  Main.bCancelAddFiles := False;
  close;
end;

procedure TOptionsOfPopupAdd.btnAddWithWildcardsClick(Sender: TObject);
begin
  BtnAddFiles_ToArchive_Click(True);
end;

procedure TOptionsOfPopupAdd.chkExcludeWildcardsClick(Sender: TObject);
begin
  if chkExcludeWildcards.Checked then
      lblWildcards.Caption := 'Exclude:'
  else
      lblWildcards.Caption := 'Add with wildcards:';
      
end;

end.
