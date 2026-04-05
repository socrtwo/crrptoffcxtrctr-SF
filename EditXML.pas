unit EditXML;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, SynEdit, ExtCtrls, XpBase, XpDOM, Menus, SynEditHighlighter,
  SynHighlighterXML, XpParser;

type
  TfrmEditXML = class(TForm)
    pnBack: TPanel;
    SynEdit1: TSynEdit;
    XpObjModel1: TXpObjModel;
    MainMenu1: TMainMenu;
    Exit1: TMenuItem;
    Save1: TMenuItem;
    SynXMLSyn1: TSynXMLSyn;
    pmnuEditor: TPopupMenu;
    Cut1: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    Delete1: TMenuItem;
    N1: TMenuItem;
    SelectAll1: TMenuItem;
    N2: TMenuItem;
    Repaint1: TMenuItem;
    Redo1: TMenuItem;
    Undo1: TMenuItem;
    N3: TMenuItem;
    ClearAll1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Save1Click(Sender: TObject);
    procedure Cut1Click(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Delete1Click(Sender: TObject);
    procedure SelectAll1Click(Sender: TObject);
    procedure Repaint1Click(Sender: TObject);
    procedure Undo1Click(Sender: TObject);
    procedure Redo1Click(Sender: TObject);
    procedure ClearAll1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  {$DEFINE IS_DEMO} // Main, EditXML Forms


var
  frmEditXML: TfrmEditXML;
  sVirtual_PathFName: string = '';
implementation


uses
  Main, ZipMstr;


const
cDefaultActiveLineColor = $00C1FFFF;
cEditXMLWidth = 'EditXMLWidth';
cEditXMLHeight = 'EditXMLHeight';

{$R *.dfm}

procedure TfrmEditXML.FormCreate(Sender: TObject);
var
  iEditXMLWidth, iEditXMLHeight: integer;
  sEditXMLWidth, sEditXMLHeight: string;
begin
  {$IFDEF IS_DEMO}
  Save1.Enabled := False;   // MainMenu1
  {$ENDIF}

 // self.Width := 800;
 // self.Height := 600;
  SynEdit1.Clear;
  sVirtual_PathFName := '';
  SynEdit1.ActiveLineColor := cDefaultActiveLineColor;

  sEditXMLWidth := frmMain.ReadIniFromReg(XcSubKeyIs, cEditXMLWidth);
  if sEditXMLWidth <> '' then begin
      iEditXMLWidth := strtoint(sEditXMLWidth);
      if (iEditXMLWidth >= 800) and (iEditXMLWidth <= Screen.Width) then
          self.Width := strtoint(sEditXMLWidth);

  end;

  sEditXMLHeight := frmMain.ReadIniFromReg(XcSubKeyIs, cEditXMLHeight);
  if sEditXMLHeight <> '' then begin
      iEditXMLHeight := strtoint(sEditXMLHeight);
      if (iEditXMLHeight >= 600) and (iEditXMLHeight <= Screen.Height) then
          self.Height := strtoint(sEditXMLHeight);

  end;

end;

procedure TfrmEditXML.Exit1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmEditXML.Save1Click(Sender: TObject);
var
  zmOpts: AddOpts;
begin
  {$IFDEF IS_DEMO}

  {$ELSE}
  Main.sErrorMessage := ''; // empty error message first

  zmOpts := [];
  zmOpts := zmOpts + [AddDirNames];
  zmOpts := zmOpts + [AddUpdate]; // must before ExtrUpdate
  zmOpts := zmOpts + [AddFreshen];
  zmOpts := zmOpts + [AddResetArchive];
  frmMain.ZipMaster1.AddOptions := zmOpts;
  frmMain.ZipMaster1.AddCompLevel := 8;

  frmMain.ZipMaster1.ZipStream.Clear; // Clear old memory data first
  SynEdit1.Lines.SaveToStream(frmMain.ZipMaster1.ZipStream);
 // frmMain.ZipMaster1.FSpecArgs.Clear;
 // frmMain.ZipMaster1.FSpecArgs.Append('word\document.xml' );
  frmMain.ZipMaster1.AddStreamToFile(sVirtual_PathFName, 0, FILE_ATTRIBUTE_ARCHIVE);

  if Main.sErrorMessage <> '' then // catched error message
      frmMain.PopUp_ErrorMessage('Error occurred when saving file to archive')
  else
      ShowMessage('File has been saved to archive.');

  {$ENDIF}    
end;

procedure TfrmEditXML.Cut1Click(Sender: TObject);
begin
  if SynEdit1.SelText <> '' then
      SynEdit1.CutToClipboard;
      
end;

procedure TfrmEditXML.Copy1Click(Sender: TObject);
begin
  if SynEdit1.SelText <> '' then
      SynEdit1.CopyToClipboard;
      
end;

procedure TfrmEditXML.Paste1Click(Sender: TObject);
begin
  SynEdit1.PasteFromClipboard;
end;

procedure TfrmEditXML.Delete1Click(Sender: TObject);
begin
  SynEdit1.SelText := '';
end;

procedure TfrmEditXML.SelectAll1Click(Sender: TObject);
begin
  SynEdit1.SelectAll;
end;

procedure TfrmEditXML.Repaint1Click(Sender: TObject);
begin
  SynEdit1.Repaint;
end;

procedure TfrmEditXML.Undo1Click(Sender: TObject);
begin
  SynEdit1.Undo;
end;

procedure TfrmEditXML.Redo1Click(Sender: TObject);
begin
  SynEdit1.Redo;
end;

procedure TfrmEditXML.ClearAll1Click(Sender: TObject);
begin
  SynEdit1.ClearAll;
end;

procedure TfrmEditXML.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  frmMain.WriteIniToReg(XcSubKeyIs, cEditXMLWidth, inttostr(Self.Width));
  frmMain.WriteIniToReg(XcSubKeyIs, cEditXMLHeight, inttostr(Self.Height));
end;

end.
