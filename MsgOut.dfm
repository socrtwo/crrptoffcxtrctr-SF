object MessageOutput: TMessageOutput
  Left = 241
  Top = 247
  AutoSize = True
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  BorderWidth = 4
  Caption = 'MessageOutput'
  ClientHeight = 343
  ClientWidth = 520
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object memOutput: TMemo
    Left = 0
    Top = 0
    Width = 520
    Height = 302
    Align = alClient
    Lines.Strings = (
      'memOutput')
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 0
  end
  object pnBot: TPanel
    Left = 0
    Top = 302
    Width = 520
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    DesignSize = (
      520
      41)
    object btnOK: TBitBtn
      Left = 424
      Top = 6
      Width = 75
      Height = 27
      Anchors = [akTop, akBottom]
      Caption = '&OK'
      TabOrder = 0
      OnClick = btnOKClick
      Kind = bkOK
    end
    object btnSaveFile: TButton
      Left = 16
      Top = 6
      Width = 75
      Height = 27
      Caption = 'Save'
      TabOrder = 1
      Visible = False
      OnClick = btnSaveFileClick
    end
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = 'txt'
    Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 136
    Top = 184
  end
end
