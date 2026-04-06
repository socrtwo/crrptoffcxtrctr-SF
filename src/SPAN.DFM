object DiskSpan: TDiskSpan
  Left = 238
  Top = 90
  AutoSize = True
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  BorderWidth = 7
  Caption = 'DiskSpan'
  ClientHeight = 393
  ClientWidth = 369
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object lblCurZip: TLabel
    Left = 8
    Top = 0
    Width = 345
    Height = 33
    AutoSize = False
    Caption = 'lblCurZip'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    WordWrap = True
  end
  object lblFileSize: TLabel
    Left = 8
    Top = 40
    Width = 74
    Height = 16
    Caption = 'lblFileSize'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label7: TLabel
    Left = 100
    Top = 328
    Width = 97
    Height = 16
    Caption = 'KB  (Max = 1024)'
  end
  object Label6: TLabel
    Left = 8
    Top = 296
    Width = 238
    Height = 16
    Caption = 'Free space on disk 1 ( Only for drive A: ):'
  end
  object GroupBox5: TGroupBox
    Left = 0
    Top = 72
    Width = 369
    Height = 89
    Caption = 'Save split zip files into location '
    TabOrder = 0
    object Label1: TLabel
      Left = 12
      Top = 24
      Width = 60
      Height = 16
      Caption = 'Directory :'
      FocusControl = edLocation
    end
    object edLocation: TEdit
      Left = 8
      Top = 42
      Width = 313
      Height = 24
      TabOrder = 0
      OnChange = edLocationChange
    end
    object btnSource: TButton
      Left = 328
      Top = 44
      Width = 25
      Height = 21
      Caption = '...'
      TabOrder = 1
      OnClick = btnSourceClick
    end
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 176
    Width = 369
    Height = 105
    Caption = ' Size of split file ( 1 MB = 1024 KB ) '
    TabOrder = 1
    object Label2: TLabel
      Left = 16
      Top = 32
      Width = 63
      Height = 16
      Caption = 'Maximum :'
    end
    object Label3: TLabel
      Left = 184
      Top = 32
      Width = 59
      Height = 16
      Caption = 'Minimum :'
    end
    object Label4: TLabel
      Left = 108
      Top = 64
      Width = 17
      Height = 16
      Caption = 'KB'
    end
    object Label5: TLabel
      Left = 276
      Top = 64
      Width = 17
      Height = 16
      Caption = 'KB'
    end
    object MaxVolSizeEdit: TEdit
      Left = 16
      Top = 56
      Width = 85
      Height = 24
      Hint = 'Maximum size of any archive part in bytes'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      Text = '65'
      OnExit = MaxVolSizeEditExit
      OnKeyPress = MaxVolSizeEditKeyPress
    end
    object MinFreeVolEdit: TEdit
      Left = 184
      Top = 56
      Width = 85
      Height = 24
      Hint = 'Min. number of bytes available before disk will be used'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Text = '65'
      OnExit = MinFreeVolEditExit
      OnKeyPress = MinFreeVolEditKeyPress
    end
  end
  object btnStart: TBitBtn
    Left = 160
    Top = 360
    Width = 89
    Height = 33
    Caption = '&Start'
    Enabled = False
    TabOrder = 3
    OnClick = btnStartClick
  end
  object BitBtn1: TBitBtn
    Left = 264
    Top = 360
    Width = 89
    Height = 33
    Caption = '&Cancel'
    TabOrder = 4
    Kind = bkCancel
  end
  object edFreeSpace: TEdit
    Left = 10
    Top = 320
    Width = 85
    Height = 24
    Hint = 'Bytes on the first disk which will not be used for the zipfile'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
    Text = '0'
    OnExit = edFreeSpaceExit
    OnKeyPress = edFreeSpaceKeyPress
  end
end
