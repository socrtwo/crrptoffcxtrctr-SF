object frmSettings: TfrmSettings
  Left = 218
  Top = 124
  AutoSize = True
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  BorderWidth = 7
  Caption = 'Settings'
  ClientHeight = 397
  ClientWidth = 482
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
  object btnOK: TButton
    Left = 407
    Top = 16
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 0
    OnClick = btnOKClick
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 482
    Height = 353
    ActivePage = tabMisc
    TabOrder = 1
    object tabMisc: TTabSheet
      Caption = 'Misc'
      object Label2: TLabel
        Left = 27
        Top = 242
        Width = 14
        Height = 16
        Alignment = taCenter
        AutoSize = False
        Caption = '*'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object Label3: TLabel
        Left = 48
        Top = 240
        Width = 144
        Height = 16
        Caption = 'Take effect after clicking'
      end
      object GroupBox1: TGroupBox
        Left = 24
        Top = 16
        Width = 417
        Height = 129
        Caption = ' System '
        TabOrder = 0
        object lblAsterisk: TLabel
          Left = 3
          Top = 26
          Width = 14
          Height = 16
          Alignment = taCenter
          AutoSize = False
          Caption = '*'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object Label1: TLabel
          Left = 3
          Top = 50
          Width = 14
          Height = 16
          Alignment = taCenter
          AutoSize = False
          Caption = '*'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object chkContextMenu: TCheckBox
          Left = 24
          Top = 25
          Width = 137
          Height = 16
          Caption = 'Context menu'
          TabOrder = 0
          OnKeyUp = chkContextMenuKeyUp
          OnMouseUp = chkContextMenuMouseUp
        end
        object chkAssociate: TCheckBox
          Left = 24
          Top = 49
          Width = 249
          Height = 16
          Caption = 'Associate HaHaZip with archives'
          TabOrder = 1
          OnKeyUp = chkAssociateKeyUp
          OnMouseUp = chkAssociateMouseUp
        end
        object chkFileToRecycleBin: TCheckBox
          Left = 24
          Top = 72
          Width = 297
          Height = 17
          Caption = 'Use Recycle Bin to store deleted archives'
          Checked = True
          State = cbChecked
          TabOrder = 2
        end
        object chkTxtWithNotePad: TCheckBox
          Left = 24
          Top = 96
          Width = 241
          Height = 17
          Caption = 'Always open txt files with Notepad'
          Checked = True
          State = cbChecked
          TabOrder = 3
        end
      end
      object cbLanguageFiles: TComboBox
        Left = 24
        Top = 273
        Width = 105
        Height = 24
        ItemHeight = 16
        TabOrder = 1
        TabStop = False
        Text = 'cbLanguageFiles'
        Visible = False
      end
      object GroupBox6: TGroupBox
        Left = 24
        Top = 160
        Width = 417
        Height = 65
        Caption = ' Temporary folder to handle compressing files '
        TabOrder = 2
        object rbWinTemp: TRadioButton
          Left = 16
          Top = 32
          Width = 129
          Height = 17
          Caption = 'Windows Temp'
          Checked = True
          TabOrder = 0
          TabStop = True
        end
        object rbCurrent: TRadioButton
          Left = 152
          Top = 32
          Width = 81
          Height = 17
          Caption = 'Current'
          TabOrder = 1
        end
        object rbTarget: TRadioButton
          Left = 248
          Top = 32
          Width = 81
          Height = 17
          Caption = 'Target'
          TabOrder = 2
        end
      end
    end
    object tabMainList: TTabSheet
      Caption = 'Main'
      ImageIndex = 1
      object GroupBox5: TGroupBox
        Left = 264
        Top = 16
        Width = 185
        Height = 41
        TabOrder = 2
        object chkBrowseList: TCheckBox
          Left = 24
          Top = 16
          Width = 145
          Height = 17
          Caption = 'Show Browse List'
          Checked = True
          State = cbChecked
          TabOrder = 0
        end
      end
      object GroupBox2: TGroupBox
        Left = 264
        Top = 76
        Width = 185
        Height = 65
        Caption = ' Position of main screen '
        TabOrder = 3
        object rbLast: TRadioButton
          Left = 16
          Top = 24
          Width = 113
          Height = 17
          Caption = 'Last position'
          Checked = True
          TabOrder = 0
          TabStop = True
        end
        object rbNormal: TRadioButton
          Tag = 1
          Left = 16
          Top = 40
          Width = 81
          Height = 17
          Caption = 'Normal'
          TabOrder = 1
        end
      end
      object GroupBox4: TGroupBox
        Left = 16
        Top = 16
        Width = 233
        Height = 51
        Caption = ' Columns width '
        TabOrder = 0
        object chkAutoSize: TCheckBox
          Left = 24
          Top = 24
          Width = 161
          Height = 17
          Caption = 'Automatically resize'
          TabOrder = 0
          OnClick = chkAutoSizeClick
        end
      end
      object GroupBox3: TGroupBox
        Left = 16
        Top = 88
        Width = 233
        Height = 97
        Caption = ' Font '
        TabOrder = 1
        object edFontName: TEdit
          Left = 16
          Top = 24
          Width = 121
          Height = 24
          TabStop = False
          ReadOnly = True
          TabOrder = 1
        end
        object btnFonts: TBitBtn
          Left = 152
          Top = 24
          Width = 57
          Height = 25
          Caption = 'Fonts'
          TabOrder = 0
          OnClick = btnFontsClick
        end
        object edFontSize: TEdit
          Left = 16
          Top = 56
          Width = 121
          Height = 24
          TabStop = False
          ReadOnly = True
          TabOrder = 2
        end
      end
    end
    object tabView: TTabSheet
      Caption = 'View'
      ImageIndex = 2
      object GroupBox7: TGroupBox
        Left = 16
        Top = 16
        Width = 185
        Height = 265
        Caption = ' Columns '
        TabOrder = 0
        object chkName: TCheckBox
          Left = 16
          Top = 24
          Width = 81
          Height = 17
          Caption = 'Name'
          Checked = True
          Enabled = False
          State = cbChecked
          TabOrder = 0
        end
        object chkModified: TCheckBox
          Left = 16
          Top = 48
          Width = 97
          Height = 17
          Caption = 'Modified'
          Checked = True
          Enabled = False
          State = cbChecked
          TabOrder = 1
        end
        object chkSize: TCheckBox
          Left = 16
          Top = 72
          Width = 81
          Height = 17
          Caption = 'Size'
          Checked = True
          Enabled = False
          State = cbChecked
          TabOrder = 2
        end
        object chkRatio: TCheckBox
          Left = 16
          Top = 96
          Width = 73
          Height = 17
          Caption = 'Ratio'
          Checked = True
          Enabled = False
          State = cbChecked
          TabOrder = 3
        end
        object chkPacked: TCheckBox
          Left = 16
          Top = 120
          Width = 89
          Height = 17
          Caption = 'Packed'
          Checked = True
          Enabled = False
          State = cbChecked
          TabOrder = 4
        end
        object chkPath: TCheckBox
          Left = 16
          Top = 216
          Width = 65
          Height = 17
          Caption = 'Path'
          Checked = True
          Enabled = False
          State = cbChecked
          TabOrder = 8
        end
        object chkType: TCheckBox
          Left = 16
          Top = 144
          Width = 73
          Height = 17
          Caption = 'Type'
          TabOrder = 5
          OnClick = chkTypeClick
        end
        object chkCRC: TCheckBox
          Left = 16
          Top = 168
          Width = 73
          Height = 17
          Caption = 'CRC'
          TabOrder = 6
          OnClick = chkCRCClick
        end
        object chkAttributes: TCheckBox
          Left = 16
          Top = 192
          Width = 97
          Height = 17
          Caption = 'Attributes'
          TabOrder = 7
          OnClick = chkAttributesClick
        end
      end
      object GroupBox8: TGroupBox
        Left = 216
        Top = 16
        Width = 233
        Height = 105
        Caption = ' General '
        TabOrder = 1
        object chkGridLines: TCheckBox
          Left = 16
          Top = 24
          Width = 97
          Height = 17
          Caption = 'Grid lines'
          TabOrder = 0
        end
        object chkRowSelect: TCheckBox
          Left = 16
          Top = 48
          Width = 97
          Height = 17
          Caption = 'Row select'
          TabOrder = 1
        end
        object chkHotTrack: TCheckBox
          Left = 16
          Top = 72
          Width = 97
          Height = 17
          Caption = 'HotTrack'
          TabOrder = 2
        end
      end
    end
    object tabFolders: TTabSheet
      Caption = 'Folders'
      ImageIndex = 3
      object GroupBox9: TGroupBox
        Left = 24
        Top = 16
        Width = 425
        Height = 78
        Caption = ' Open folder '
        TabOrder = 0
        object rbLastOpen: TRadioButton
          Left = 24
          Top = 24
          Width = 89
          Height = 17
          Caption = 'Last open'
          Checked = True
          TabOrder = 0
          TabStop = True
        end
        object rbFavoriteOpen: TRadioButton
          Left = 24
          Top = 43
          Width = 81
          Height = 17
          Caption = 'Favorite:'
          TabOrder = 1
        end
        object edFavoriteOpen: TEdit
          Left = 112
          Top = 40
          Width = 265
          Height = 24
          TabStop = False
          ReadOnly = True
          TabOrder = 3
        end
        object btnFavoriteOpen: TBitBtn
          Left = 384
          Top = 40
          Width = 27
          Height = 25
          TabOrder = 2
          OnClick = btnFavoriteOpenClick
          Glyph.Data = {
            76020000424D760200000000000036000000280000000F0000000C0000000100
            1800000000004002000074120000741200000000000000000000FF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FF000000FF00FF0E575737AABC008FB40086AD00719B006D9800
            6D982A6D7F2B6D7F1E6D861E6D8641375BFF00FFFF00FF000000FF00FF137777
            4AE6FF00CEFF00BBEF00B8ED0094CE0094CE0D94C63194B13294B13294B1174A
            59FF00FFFF00FF000000FF00FF51B8B84AB9BD6CE9FF10D2FF00CBFC00C3F600
            B9ED00A0D80A94C82A94B52A94B53188A1FF00FFFF00FF000000FF00FF5DC5C5
            4AB1B184EBFF6BE7FF00CEFF00CBFC00B6EA00B0E600A1D91A94BF2194BB3194
            B141375BFF00FF000000FF00FF67D3D72CA2B175E4ED8AE9FF63E3FF49DFFF00
            CEFF00B5EA00B5EA00B2E708AFE02294BA2A6A80FF00FF000000FF00FF68D4D8
            2FA4B56EDFE78BEAFF77E6FF69E6FF49E0FF08D1FF08D1FF08B7EA08B5E8089A
            D106739CFF00FF000000FF00FF6FD5D87FE7FF238A8D56BDBD56BDBD56BDBD56
            BDBD56BDBD56BDBD56BDBD56BDBD56BDBD408E8EFF00FF000000FF00FF5AC4C6
            78E1E664CACA5DC6C65DC6C65DC6C65DC6C65DC6C65DC6C65DC6C65DC6C62F6D
            6DFF00FFFF00FF000000FF00FF3260824399A831909B31909B1D889B13849700
            5656FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
            8092B770CFD167CFD163CDD15EA1AEFF00FFFC00FCFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000}
        end
      end
      object GroupBox10: TGroupBox
        Left = 24
        Top = 104
        Width = 425
        Height = 121
        Caption = ' Add folder '
        TabOrder = 1
        object rbLastAdd: TRadioButton
          Left = 24
          Top = 24
          Width = 81
          Height = 17
          Caption = 'Last add'
          Checked = True
          TabOrder = 0
          TabStop = True
        end
        object rbFavoriteAdd: TRadioButton
          Left = 24
          Top = 43
          Width = 81
          Height = 17
          Caption = 'Favorite:'
          TabOrder = 1
        end
        object edFavoriteAdd: TEdit
          Left = 112
          Top = 40
          Width = 265
          Height = 24
          TabStop = False
          ReadOnly = True
          TabOrder = 5
        end
        object btnFavoriteAdd: TBitBtn
          Left = 384
          Top = 40
          Width = 27
          Height = 25
          TabOrder = 2
          OnClick = btnFavoriteAddClick
          Glyph.Data = {
            76020000424D760200000000000036000000280000000F0000000C0000000100
            1800000000004002000074120000741200000000000000000000FF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FF000000FF00FF0E575737AABC008FB40086AD00719B006D9800
            6D982A6D7F2B6D7F1E6D861E6D8641375BFF00FFFF00FF000000FF00FF137777
            4AE6FF00CEFF00BBEF00B8ED0094CE0094CE0D94C63194B13294B13294B1174A
            59FF00FFFF00FF000000FF00FF51B8B84AB9BD6CE9FF10D2FF00CBFC00C3F600
            B9ED00A0D80A94C82A94B52A94B53188A1FF00FFFF00FF000000FF00FF5DC5C5
            4AB1B184EBFF6BE7FF00CEFF00CBFC00B6EA00B0E600A1D91A94BF2194BB3194
            B141375BFF00FF000000FF00FF67D3D72CA2B175E4ED8AE9FF63E3FF49DFFF00
            CEFF00B5EA00B5EA00B2E708AFE02294BA2A6A80FF00FF000000FF00FF68D4D8
            2FA4B56EDFE78BEAFF77E6FF69E6FF49E0FF08D1FF08D1FF08B7EA08B5E8089A
            D106739CFF00FF000000FF00FF6FD5D87FE7FF238A8D56BDBD56BDBD56BDBD56
            BDBD56BDBD56BDBD56BDBD56BDBD56BDBD408E8EFF00FF000000FF00FF5AC4C6
            78E1E664CACA5DC6C65DC6C65DC6C65DC6C65DC6C65DC6C65DC6C65DC6C62F6D
            6DFF00FFFF00FF000000FF00FF3260824399A831909B31909B1D889B13849700
            5656FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
            8092B770CFD167CFD163CDD15EA1AEFF00FFFC00FCFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000}
        end
        object chkHiddenFiles: TCheckBox
          Left = 24
          Top = 72
          Width = 169
          Height = 17
          Caption = 'Don'#39't show hidden files'
          TabOrder = 3
        end
        object chkSystemFiles: TCheckBox
          Left = 24
          Top = 94
          Width = 169
          Height = 17
          Caption = 'Don'#39't show system files'
          TabOrder = 4
        end
      end
      object GroupBox11: TGroupBox
        Left = 24
        Top = 232
        Width = 425
        Height = 78
        Caption = ' Extract folder '
        TabOrder = 2
        object rbLastExtract: TRadioButton
          Left = 24
          Top = 24
          Width = 97
          Height = 17
          Caption = 'Last extract'
          Checked = True
          TabOrder = 0
          TabStop = True
        end
        object rbFavoriteExtract: TRadioButton
          Left = 24
          Top = 43
          Width = 81
          Height = 17
          Caption = 'Favorite:'
          TabOrder = 1
        end
        object edFavoriteExtract: TEdit
          Left = 112
          Top = 40
          Width = 265
          Height = 24
          TabStop = False
          ReadOnly = True
          TabOrder = 3
        end
        object btnFavoriteExtract: TBitBtn
          Left = 384
          Top = 40
          Width = 27
          Height = 25
          TabOrder = 2
          OnClick = btnFavoriteExtractClick
          Glyph.Data = {
            76020000424D760200000000000036000000280000000F0000000C0000000100
            1800000000004002000074120000741200000000000000000000FF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FF000000FF00FF0E575737AABC008FB40086AD00719B006D9800
            6D982A6D7F2B6D7F1E6D861E6D8641375BFF00FFFF00FF000000FF00FF137777
            4AE6FF00CEFF00BBEF00B8ED0094CE0094CE0D94C63194B13294B13294B1174A
            59FF00FFFF00FF000000FF00FF51B8B84AB9BD6CE9FF10D2FF00CBFC00C3F600
            B9ED00A0D80A94C82A94B52A94B53188A1FF00FFFF00FF000000FF00FF5DC5C5
            4AB1B184EBFF6BE7FF00CEFF00CBFC00B6EA00B0E600A1D91A94BF2194BB3194
            B141375BFF00FF000000FF00FF67D3D72CA2B175E4ED8AE9FF63E3FF49DFFF00
            CEFF00B5EA00B5EA00B2E708AFE02294BA2A6A80FF00FF000000FF00FF68D4D8
            2FA4B56EDFE78BEAFF77E6FF69E6FF49E0FF08D1FF08D1FF08B7EA08B5E8089A
            D106739CFF00FF000000FF00FF6FD5D87FE7FF238A8D56BDBD56BDBD56BDBD56
            BDBD56BDBD56BDBD56BDBD56BDBD56BDBD408E8EFF00FF000000FF00FF5AC4C6
            78E1E664CACA5DC6C65DC6C65DC6C65DC6C65DC6C65DC6C65DC6C65DC6C62F6D
            6DFF00FFFF00FF000000FF00FF3260824399A831909B31909B1D889B13849700
            5656FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
            8092B770CFD167CFD163CDD15EA1AEFF00FFFC00FCFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000}
        end
      end
    end
    object tabProgLoc: TTabSheet
      Caption = 'Program Location'
      ImageIndex = 4
      object gbVirusScanner: TGroupBox
        Left = 24
        Top = 16
        Width = 433
        Height = 185
        Caption = ' Virus Scanner '
        TabOrder = 0
        object lblScanner: TLabel
          Left = 18
          Top = 24
          Width = 55
          Height = 16
          Caption = 'Program:'
        end
        object lblPara: TLabel
          Left = 16
          Top = 120
          Width = 73
          Height = 16
          Caption = 'Parameters:'
        end
        object edVirusScanner: TEdit
          Left = 16
          Top = 46
          Width = 369
          Height = 24
          TabOrder = 0
        end
        object btnVirusScanner: TBitBtn
          Left = 392
          Top = 46
          Width = 27
          Height = 25
          TabOrder = 1
          OnClick = btnVirusScannerClick
          Glyph.Data = {
            76020000424D760200000000000036000000280000000F0000000C0000000100
            1800000000004002000074120000741200000000000000000000FF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FF000000FF00FF0E575737AABC008FB40086AD00719B006D9800
            6D982A6D7F2B6D7F1E6D861E6D8641375BFF00FFFF00FF000000FF00FF137777
            4AE6FF00CEFF00BBEF00B8ED0094CE0094CE0D94C63194B13294B13294B1174A
            59FF00FFFF00FF000000FF00FF51B8B84AB9BD6CE9FF10D2FF00CBFC00C3F600
            B9ED00A0D80A94C82A94B52A94B53188A1FF00FFFF00FF000000FF00FF5DC5C5
            4AB1B184EBFF6BE7FF00CEFF00CBFC00B6EA00B0E600A1D91A94BF2194BB3194
            B141375BFF00FF000000FF00FF67D3D72CA2B175E4ED8AE9FF63E3FF49DFFF00
            CEFF00B5EA00B5EA00B2E708AFE02294BA2A6A80FF00FF000000FF00FF68D4D8
            2FA4B56EDFE78BEAFF77E6FF69E6FF49E0FF08D1FF08D1FF08B7EA08B5E8089A
            D106739CFF00FF000000FF00FF6FD5D87FE7FF238A8D56BDBD56BDBD56BDBD56
            BDBD56BDBD56BDBD56BDBD56BDBD56BDBD408E8EFF00FF000000FF00FF5AC4C6
            78E1E664CACA5DC6C65DC6C65DC6C65DC6C65DC6C65DC6C65DC6C65DC6C62F6D
            6DFF00FFFF00FF000000FF00FF3260824399A831909B31909B1D889B13849700
            5656FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
            8092B770CFD167CFD163CDD15EA1AEFF00FFFC00FCFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000}
        end
        object edPara: TEdit
          Left = 16
          Top = 142
          Width = 369
          Height = 24
          TabOrder = 4
        end
        object rbFilenamePara: TRadioButton
          Left = 208
          Top = 88
          Width = 177
          Height = 17
          Caption = 'Filename + Parameters'
          Checked = True
          TabOrder = 3
          TabStop = True
        end
        object rbParaFilename: TRadioButton
          Left = 16
          Top = 88
          Width = 177
          Height = 17
          Caption = 'Parameters + Filename'
          TabOrder = 2
        end
      end
    end
  end
  object pnBottom: TPanel
    Left = 0
    Top = 359
    Width = 482
    Height = 38
    BevelOuter = bvNone
    Ctl3D = False
    ParentCtl3D = False
    TabOrder = 2
    object btnSettingsApply: TBitBtn
      Left = 216
      Top = 6
      Width = 75
      Height = 30
      Hint = 'Apply changes on current page'
      Caption = '&Apply'
      Default = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = btnSettingsApplyClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333333333333333333333330000333333333333333333333333F33333333333
        00003333344333333333333333388F3333333333000033334224333333333333
        338338F3333333330000333422224333333333333833338F3333333300003342
        222224333333333383333338F3333333000034222A22224333333338F338F333
        8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
        33333338F83338F338F33333000033A33333A222433333338333338F338F3333
        0000333333333A222433333333333338F338F33300003333333333A222433333
        333333338F338F33000033333333333A222433333333333338F338F300003333
        33333333A222433333333333338F338F00003333333333333A22433333333333
        3338F38F000033333333333333A223333333333333338F830000333333333333
        333A333333333333333338330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
    object btnSettingsCancel: TBitBtn
      Left = 384
      Top = 6
      Width = 81
      Height = 30
      Cancel = True
      Caption = '&Cancel'
      TabOrder = 2
      OnClick = btnSettingsCancelClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333333333000033338833333333333333333F333333333333
        0000333911833333983333333388F333333F3333000033391118333911833333
        38F38F333F88F33300003339111183911118333338F338F3F8338F3300003333
        911118111118333338F3338F833338F3000033333911111111833333338F3338
        3333F8330000333333911111183333333338F333333F83330000333333311111
        8333333333338F3333383333000033333339111183333333333338F333833333
        00003333339111118333333333333833338F3333000033333911181118333333
        33338333338F333300003333911183911183333333383338F338F33300003333
        9118333911183333338F33838F338F33000033333913333391113333338FF833
        38F338F300003333333333333919333333388333338FFF830000333333333333
        3333333333333333333888330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
    object btnSettingsOK: TBitBtn
      Left = 296
      Top = 6
      Width = 75
      Height = 30
      Hint = 'Apply changes on current page'
      Caption = '&OK'
      Default = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = btnSettingsOKClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333333333333333333333330000333333333333333333333333F33333333333
        00003333344333333333333333388F3333333333000033334224333333333333
        338338F3333333330000333422224333333333333833338F3333333300003342
        222224333333333383333338F3333333000034222A22224333333338F338F333
        8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
        33333338F83338F338F33333000033A33333A222433333338333338F338F3333
        0000333333333A222433333333333338F338F33300003333333333A222433333
        333333338F338F33000033333333333A222433333333333338F338F300003333
        33333333A222433333333333338F338F00003333333333333A22433333333333
        3338F38F000033333333333333A223333333333333338F830000333333333333
        333A333333333333333338330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
  end
  object FontDialog1: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Left = 384
    Top = 264
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Executable(*.exe)|*.exe'
    Options = [ofHideReadOnly, ofFileMustExist, ofEnableSizing]
    Left = 348
    Top = 267
  end
end
