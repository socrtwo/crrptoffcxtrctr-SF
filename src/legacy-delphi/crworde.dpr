program crworde;
uses
  Forms,
  SysUtils,
  Windows,
  Registry,
  StrUtils,
  ShellCtrls,
  Dialogs,
  Main in 'Main.pas' {frmMain},
  OptAddFl in 'OptAddFl.pas' {OptionsOfPopupAdd},
  MsgOut in 'MsgOut.pas' {MessageOutput},
  settings in 'settings.pas' {frmSettings},
  reabout in 'reabout.pas' {AboutBox},
  hrglass in 'hrglass.pas' {HourGlass},
  InputMsgBox in 'InputMsgBox.pas' {InputMsg},
  span in 'span.pas' {DiskSpan},
  Open_Add in 'Open_Add.pas' {frmDir},
  Extract_To in 'Extract_To.pas' {frmToFolder},
  Repair in 'Repair.pas' {frmRepair},
  SearchBox in 'SearchBox.pas' {Search},
  EditXML in 'EditXML.pas' {frmEditXML},
  OpenDir in 'OpenDir.pas' {frmOpenDir};

{$R *.res}
var
  i: integer;
  sPath: string;
  sDirOrFilename: string;
  sOutputPathZip: string = '';
  sFromDir: string = '';
  sTailPara: string;
  sHeader: string = '';
  sExt: string;
  //FFileName: array[0..MAX_PATH] of Char;
begin
  Application.Initialize;
  Application.Title := 'Corrupt Office 2007 Extractor';
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TMessageOutput, MessageOutput);
  Application.CreateForm(THourGlass, HourGlass);
  Application.CreateForm(TDiskSpan, DiskSpan);
  Application.CreateForm(TSearch, Search);
  Application.CreateForm(TfrmEditXML, frmEditXML);
  Application.CreateForm(TfrmOpenDir, frmOpenDir);
  if (ParamCount > 0) then begin    //and FileExists(ParamStr(1)) then begin
      Main.slZipStrings.Clear;
      for i := 1 to ParamCount do begin
            //sDirOrFilename := ExpandFilename(ParamStr(i));
            //sDirOrFilename := AnsiReplaceStr(ParamStr(i), '#13#10', ' ');
            //StrCopy( FFileName, PChar( ParamStr(i) ) );
            frmMain.lblExtractParaStr.Caption := ParamStr(i); // don't expand filename here, cause error
            sDirOrFilename := frmMain.lblExtractParaStr.Caption;
            sExt := AnsiLowercase(ExtractFileExt(sDirOrFilename));

           { if (ParamCount=1) and (AnsiLowercase(ExtractFileExt(sDirOrFilename))
            ='.docx') then begin
                frmMain.WhatCompressionFile_WillBe_DirectlyExtracted(sDirOrFilename);
                Application.ShowMainForm := False;
                Application.Terminate;
                break;
            end }
          {  if (ParamCount=1) and ( (sDirOrFilename = '/?') or
            (AnsiLowercase(sDirOrFilename) = '-h') ) then begin
                Write('          Extract xml files from docx Word document' + #13#10);
                Write('Copyright (C) 2009 Ccy JC_HK       www.ccyjchk.com' + #13#10 );
                Write('hhzip v1.0       This program is designed for www.s2services.com' + #13#10 + #13#10);
                Write('Usage: hhzip [-thx] ''''file''''' + #13#10 + #13#10);
                Write('Commands:' + #13#10);
                Write('  -h    this help' + #13#10);
                Write('  -t    output text' + #13#10);
                Write('  -x    extract xml files' + #13#10);
                Application.ShowMainForm := False;
                Application.Terminate;
                break;
            end;

            if (ParamCount=2) and (sDirOrFilename = '-t') then begin
                Main.XbOutputTxtContentsFile := True;
            end
            else if (ParamCount=2) and (sDirOrFilename = '-x') then begin
                Main.XbOutputXmlFile := True;
            end
            else if (ParamCount=2) and
            (AnsiLowercase(ExtractFileExt(sDirOrFilename)) ='.docx') then begin
                // is correct command input
                if Main.XbOutputTxtContentsFile or Main.XbOutputXmlFile then begin
                    frmMain.WhatCompressionFile_WillBe_DirectlyExtracted(sDirOrFilename);
                    if Main.XbOutputTxtContentsFile and Main.XbIsDoUnzip then begin
                        frmMain.Output_Txt_Contents_Of_XML_From_Docx;
                        Main.XbIsDoUnzip := False;
                    end;
                end
                else
                    Write('Err: invalid input of Command.');

                Application.ShowMainForm := False;
                Application.Terminate;
                break;
            end
            else begin
                if Main.XbOutputTxtContentsFile or Main.XbOutputXmlFile then
                    Write('Err: Not docx file found.')
                else
                    Write('Err: invalid input of Command.');

                Application.ShowMainForm := False;
                Application.Terminate;
                break;
            end; }

            // Only one file and it is docx file, try unzip
            if (ParamCount=1) and ( (sExt = '.docx') or (sExt = '.pptx') or
            (sExt = '.xlsx') ) then begin
                frmMain.Show;
                frmMain.Repaint;
                frmMain.SetZipFile(sDirOrFilename, True);

                frmMain.AfterOpeningArchive_And_Do;

                //frmMain.StatusBar1.Panels[0].Text := sDirOrFilename;
                //frmMain.StartUnzip(True, sDirOrFilename, '', nil, False);
                break;
            end
            else begin
                ShowMessage('Not Office 2007 file found.');
                Application.ShowMainForm := False;
                Application.Terminate;
                break;
            end;






            
      end;



      
  end
  else begin
     { Write('          Extract xml files from docx Word document' + #13#10);
      Write('Copyright (C) 2009 Ccy JC_HK       www.ccyjchk.com' + #13#10 );
      Write('hhzip v1.0       This program is designed for www.s2services.com' + #13#10 + #13#10);
      Write('This program only run in command-line mode.' + #13#10 + #13#10);
      Application.ShowMainForm := False;
      Application.Terminate; }
  end;

  // Check no termination before calling. prevent dead loop!
  if not Application.Terminated then
      Application.Run;
      
end.
