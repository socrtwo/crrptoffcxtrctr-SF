object OptionsOfPopupAdd: TOptionsOfPopupAdd
  Left = 251
  Top = 186
  AutoSize = True
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  BorderWidth = 7
  Caption = 'options before adding files'
  ClientHeight = 270
  ClientWidth = 304
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object lblCompressRatio: TLabel
    Left = 8
    Top = 169
    Width = 115
    Height = 16
    Caption = 'Compression level:'
  end
  object lblWildcards: TLabel
    Left = 8
    Top = 226
    Width = 113
    Height = 16
    Caption = 'Add with wildcards:'
  end
  object gb1: TGroupBox
    Left = 0
    Top = 0
    Width = 304
    Height = 153
    TabOrder = 6
    object Label1: TLabel
      Left = 41
      Top = 88
      Width = 176
      Height = 14
      Caption = '( no check = always overwrite )'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object chAddWithDir: TCheckBox
      Left = 16
      Top = 47
      Width = 137
      Height = 16
      Caption = 'with directories'
      Checked = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      State = cbChecked
      TabOrder = 1
    end
    object chOverwriteWithNew: TCheckBox
      Left = 16
      Top = 69
      Width = 153
      Height = 16
      Caption = 'overwrite if newer'
      Checked = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      State = cbChecked
      TabOrder = 2
    end
    object btnComment: TButton
      Left = 200
      Top = 24
      Width = 83
      Height = 25
      Hint = 'Add comment to archive'
      Caption = 'Comment'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = btnCommentClick
    end
    object chPassword: TCheckBox
      Left = 16
      Top = 24
      Width = 97
      Height = 17
      Caption = 'password'
      TabOrder = 0
    end
    object chIncludeHidden: TCheckBox
      Left = 16
      Top = 130
      Width = 153
      Height = 17
      Hint = 
        'To include hidden files or not although hidden files are selecte' +
        'd'
      Caption = 'Include hidden files'
      Checked = True
      ParentShowHint = False
      ShowHint = True
      State = cbChecked
      TabOrder = 3
    end
    object chIncludeSystem: TCheckBox
      Left = 16
      Top = 108
      Width = 153
      Height = 17
      Hint = 
        'To include system files or not although system files are selecte' +
        'd'
      Caption = 'Include system files'
      Checked = True
      ParentShowHint = False
      ShowHint = True
      State = cbChecked
      TabOrder = 5
    end
  end
  object cbCompressRatio: TComboBox
    Left = 8
    Top = 192
    Width = 161
    Height = 24
    ItemHeight = 16
    TabOrder = 3
  end
  object btnOK: TBitBtn
    Left = 208
    Top = 168
    Width = 81
    Height = 28
    Caption = '&OK'
    TabOrder = 0
    OnClick = btnOKClick
    Kind = bkOK
  end
  object btnCancel: TBitBtn
    Left = 208
    Top = 240
    Width = 81
    Height = 28
    Caption = '&Cancel'
    TabOrder = 2
    OnClick = btnCancelClick
    Kind = bkCancel
  end
  object edWildcards: TEdit
    Left = 8
    Top = 246
    Width = 161
    Height = 24
    Hint = 'For example:'#13#10'*.doc,*.gif'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    OnChange = edWildcardsChange
  end
  object btnAddWithWildcards: TBitBtn
    Left = 224
    Top = 202
    Width = 65
    Height = 24
    Hint = 
      'Adding files:'#13#10'1.) based on specified wildcards or'#13#10'2.) exclude ' +
      'specified wildcards'
    Caption = 'A&dd  *.?'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    OnClick = btnAddWithWildcardsClick
  end
  object chkExcludeWildcards: TCheckBox
    Left = 136
    Top = 226
    Width = 32
    Height = 17
    Hint = 'Exclude'
    Caption = 'E'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    OnClick = chkExcludeWildcardsClick
  end
end
