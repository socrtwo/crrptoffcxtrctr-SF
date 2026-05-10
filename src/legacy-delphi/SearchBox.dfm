object Search: TSearch
  Left = 215
  Top = 177
  AutoSize = True
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Search'
  ClientHeight = 107
  ClientWidth = 526
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 120
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 526
    Height = 107
    BevelOuter = bvNone
    TabOrder = 0
    object Label1: TLabel
      Left = 184
      Top = 8
      Width = 66
      Height = 16
      Caption = 'Search file:'
    end
    object lblMessage: TLabel
      Left = 184
      Top = 72
      Width = 71
      Height = 16
      Caption = 'lblMessage'
    end
    object edSearch: TEdit
      Left = 184
      Top = 32
      Width = 209
      Height = 24
      TabOrder = 0
      Text = 'edSearch'
      OnChange = edSearchChange
      OnClick = edSearchClick
      OnKeyPress = edSearchKeyPress
    end
    object btnFind: TButton
      Left = 408
      Top = 8
      Width = 105
      Height = 25
      Caption = '&Find'
      Default = True
      TabOrder = 2
      OnClick = btnFindClick
    end
    object btnCancel: TButton
      Left = 408
      Top = 72
      Width = 105
      Height = 25
      Caption = '&Cancel'
      TabOrder = 4
      OnClick = btnCancelClick
    end
    object GroupBox1: TGroupBox
      Left = 16
      Top = 6
      Width = 145
      Height = 92
      TabOrder = 1
      object chkCaseSensitive: TCheckBox
        Left = 14
        Top = 24
        Width = 121
        Height = 17
        Caption = 'case sensitive'
        Checked = True
        State = cbChecked
        TabOrder = 0
      end
      object chkWholeWords: TCheckBox
        Left = 14
        Top = 48
        Width = 129
        Height = 17
        Caption = 'whole words only'
        Checked = True
        State = cbChecked
        TabOrder = 1
      end
    end
    object btnFindNext: TButton
      Left = 408
      Top = 40
      Width = 105
      Height = 25
      Caption = 'Find &Next'
      Default = True
      Enabled = False
      TabOrder = 3
      OnClick = btnFindNextClick
    end
  end
end
