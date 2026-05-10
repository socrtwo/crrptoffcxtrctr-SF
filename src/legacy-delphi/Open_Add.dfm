object frmDir: TfrmDir
  Left = 247
  Top = 196
  Width = 729
  Height = 610
  BorderIcons = [biSystemMenu]
  BorderWidth = 7
  Caption = 'OpenAdd'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object pnBottom: TPanel
    Left = 0
    Top = 350
    Width = 697
    Height = 203
    Align = alBottom
    TabOrder = 2
    object PageControl1: TPageControl
      Left = 1
      Top = 1
      Width = 695
      Height = 201
      ActivePage = tabUnzip
      Align = alClient
      Style = tsFlatButtons
      TabOrder = 0
      object tabAdd: TTabSheet
        Caption = 'Add files'
        TabVisible = False
        OnShow = tabAddShow
        object gbAddFiles: TGroupBox
          Left = 0
          Top = 0
          Width = 687
          Height = 191
          Align = alClient
          Caption = ' Only accept files or directories from above right pane '
          TabOrder = 0
          object lblCompressRatio: TLabel
            Left = 16
            Top = 28
            Width = 118
            Height = 16
            Caption = 'Compression level :'
          end
          object lblWildcards: TLabel
            Left = 16
            Top = 86
            Width = 113
            Height = 16
            Caption = 'Add with wildcards:'
          end
          object btnAdd: TBitBtn
            Left = 552
            Top = 32
            Width = 81
            Height = 30
            Caption = '&Add'
            TabOrder = 5
            OnClick = btnAddClick
            Glyph.Data = {
              06030000424D060300000000000036000000280000000F0000000F0000000100
              180000000000D002000074120000741200000000000000000000FF00FFFF00FF
              FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFB813C362478325305CFF00
              FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFA0
              0FAB4B4C93A9A2C293BBD327444D57416AFF00FFFF00FF000000FF00FFFF00FF
              FF00FFFF00FFFF00FFA509B36A46957D96B4C68498E183979BC0D9455D631F34
              3F6C597FFF00FF000000FF00FFFF00FFFF00FF800D947650907AA8C8D37C8EF0
              1B1EFF0000ED333A96B4C8598388162222476262FF00FF000000FF00FFFF00FF
              8B0D9388B4CECAC6E5A9D7F4D3616CD86E7DFF0000FB181CA596A444B5BF1A2C
              349E45AEFF00FF000000FF00FFFF00FF854288A8D8F4D2CDEBA9D8F4BAA5BAF5
              1214FF0000C6B1C97B97A03DDBE12B3B40FF00FFFF00FF000000FF00FFFF00FF
              6A90AFDF7C8E98CEE5B2D8F2D7505AF81114E898B0A2DAF55FAAB33EECEE3A45
              5F9E08A2FF00FF000000FF00FF762075A2CCDFFF0000EE2428F42B30FF0000EB
              434CC2D1EE99D3EFBBB7CBD0F9FA576773531D5FFF00FF000000FF00FF5B275C
              D2D5E3FF0000FF0000E64E57CB8C96CAC4DCC6CCEA9ABAD0C1CFD7E6FEFC687B
              84485169FF00FF000000990099537E85E0DAEDB6E7F595EBF4B0DBF1A3D2E4A0
              B1C3A4B6C2BCDBDDE1F9F8E6FEFC7992923B5E77FF00FF0000009E069F8C8B92
              E2EEF59FE5EEA4D3E085C5CAD8C3D7D4EBECE0F7F6E6FDFCDFF6F5CCE8E7708F
              9C567997FF00FF000000B31CB5A9CAC978A5A492B0B1BADBDBDEF4F2E6FEFCE6
              FEFCE6FEFCB7CCCF91B2C46697B9628CB58183C2FF00FF0000003A1A3F832F88
              FF00FF93A2A1E6FEFCE4FCFADDECEFA9C4CD6889A53F68A46290BD86A0D5688A
              AF894C93FF00FF000000FF00FFFF00FFFF00FF849392C1BFCA8A8EA66A8AA061
              689479349CA70AB69F10AD931D9FA510AEFF00FFFF00FF000000FF00FFFF00FF
              FF00FF47234CFF00FF4F0D59542266C508C6FF00FFFF00FFFF00FFFF00FFFF00
              FFFF00FFFF00FF000000}
          end
          object btnCancelAdd: TBitBtn
            Left = 552
            Top = 104
            Width = 81
            Height = 30
            Caption = '&Cancel'
            TabOrder = 7
            OnClick = btnCancelAddClick
            Kind = bkCancel
          end
          object cbCompressRatio: TComboBox
            Left = 16
            Top = 48
            Width = 185
            Height = 24
            ItemHeight = 0
            TabOrder = 0
          end
          object chKeepAboveSet: TCheckBox
            Left = 16
            Top = 138
            Width = 129
            Height = 17
            Caption = 'Keep settings'
            TabOrder = 3
          end
          object gb1: TGroupBox
            Left = 216
            Top = 22
            Width = 297
            Height = 140
            TabOrder = 4
            object Label1: TLabel
              Left = 49
              Top = 75
              Width = 191
              Height = 14
              Caption = '( always overwrite when disabled )'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -12
              Font.Name = 'Tahoma'
              Font.Style = []
              ParentFont = False
            end
            object chAddWithDir: TCheckBox
              Left = 16
              Top = 38
              Width = 129
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
              Top = 58
              Width = 145
              Height = 17
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
              Left = 192
              Top = 24
              Width = 75
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
              Top = 16
              Width = 97
              Height = 17
              Caption = 'password'
              TabOrder = 0
            end
            object chIncludeHidden: TCheckBox
              Left = 16
              Top = 112
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
              Top = 92
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
          object edWildcards: TEdit
            Left = 16
            Top = 106
            Width = 185
            Height = 24
            Hint = 'For example:'#13#10'*.doc,*.gif'
            ParentShowHint = False
            ShowHint = True
            TabOrder = 2
            OnChange = edWildcardsChange
          end
          object btnAddWithWildcards: TBitBtn
            Left = 568
            Top = 66
            Width = 65
            Height = 24
            Hint = 
              'Adding files:'#13#10'1.) based on specified wildcards or'#13#10'2.) exclude ' +
              'specified wildcards'
            Caption = 'A&dd  *.?'
            Enabled = False
            ParentShowHint = False
            ShowHint = True
            TabOrder = 6
            OnClick = btnAddWithWildcardsClick
          end
          object chkExcludeWildcards: TCheckBox
            Left = 168
            Top = 86
            Width = 32
            Height = 17
            Hint = 'Exclude'
            Caption = 'E'
            ParentShowHint = False
            ShowHint = True
            TabOrder = 1
            OnClick = chkExcludeWildcardsClick
          end
        end
      end
      object tabUnzip: TTabSheet
        Caption = 'Open'
        ImageIndex = 1
        TabVisible = False
        OnShow = tabUnzipShow
        object GroupBox1: TGroupBox
          Left = 0
          Top = 0
          Width = 687
          Height = 191
          Align = alClient
          Caption = 
            ' Double-click on a docx file to open it directly. Or, input file' +
            'name on text box and press Open button. '
          TabOrder = 0
          object btnHideFolders: TSpeedButton
            Left = 208
            Top = 32
            Width = 105
            Height = 25
            AllowAllUp = True
            GroupIndex = 1
            Caption = '&Hide folders'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            OnClick = btnHideFoldersClick
            OnMouseDown = btnHideFoldersMouseDown
          end
          object lblPath: TLabel
            Left = 21
            Top = 80
            Width = 28
            Height = 16
            Alignment = taRightJustify
            Caption = 'File :'
          end
          object btnUnzipThis: TBitBtn
            Left = 430
            Top = 29
            Width = 107
            Height = 30
            Caption = 'Open docx file'
            TabOrder = 1
            OnClick = btnUnzipThisClick
          end
          object btnCancelUnzipThis: TBitBtn
            Left = 552
            Top = 29
            Width = 81
            Height = 30
            Caption = '&Cancel'
            TabOrder = 2
            Kind = bkCancel
          end
          object edPath: TEdit
            Left = 56
            Top = 72
            Width = 577
            Height = 22
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 0
            Text = 'edPath'
            OnKeyPress = edPathKeyPress
          end
        end
      end
    end
  end
  object pnMiddle: TPanel
    Left = 0
    Top = 33
    Width = 697
    Height = 317
    Align = alClient
    BevelOuter = bvNone
    Caption = 'pnMiddle'
    TabOrder = 1
    object Splitter1: TSplitter
      Left = 218
      Top = 0
      Height = 317
    end
    object pnTree: TPanel
      Left = 0
      Top = 0
      Width = 218
      Height = 317
      Align = alLeft
      Caption = 'pnTree'
      TabOrder = 0
      object DriveView1: TDriveView
        Left = 1
        Top = 82
        Width = 216
        Height = 234
        FullDriveScan = False
        DimmHiddenDirs = False
        WatchDirectory = False
        DirView = DirView1
        ShowDirSize = False
        ShowVolLabel = True
        ShowAnimation = False
        FileNameDisplay = fndStored
        Align = alClient
        DragMode = dmAutomatic
        Indent = 23
        ParentColor = False
        TabOrder = 0
        TabStop = True
        OnChange = DriveView1Change
        OnKeyUp = DriveView1KeyUp
      end
      object CoolBar2: TCoolBar
        Left = 1
        Top = 1
        Width = 216
        Height = 81
        AutoSize = True
        Bands = <
          item
            Control = IEDriveComboBox1
            ImageIndex = -1
            MinHeight = 24
            Width = 212
          end
          item
            Control = ToolBar1
            ImageIndex = -1
            Width = 212
          end
          item
            Control = FilterComboBox1
            ImageIndex = -1
            MinHeight = 24
            Width = 212
          end>
        object IEDriveComboBox1: TIEDriveComboBox
          Left = 9
          Top = 0
          Width = 199
          Height = 24
          DropDownFixedWidth = 0
          DriveView = DriveView1
          TabOrder = 0
          OnChange = IEDriveComboBox1Change
          OnKeyUp = IEDriveComboBox1KeyUp
        end
        object ToolBar1: TToolBar
          Left = 9
          Top = 26
          Width = 199
          Height = 25
          Caption = 'ToolBar1'
          TabOrder = 1
          object pnOnCoolBar: TPanel
            Left = 0
            Top = 2
            Width = 201
            Height = 22
            Align = alClient
            BevelOuter = bvNone
            TabOrder = 0
            object btnUp1: TSpeedButton
              Left = 0
              Top = 0
              Width = 25
              Height = 22
              Flat = True
              Glyph.Data = {
                A6020000424DA6020000000000003600000028000000100000000D0000000100
                1800000000007002000074120000741200000000000000000000FF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
                FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
                00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF868686
                FFFFFFCCCCCCCCCCCCCCCCCCB2B2B2B2B2B2868686B2B2B28686868686868686
                86868686FF00FFFF00FFFF00FF868686FFFFFFDDDDDDCCCCCCCCCCCCCCCCCCB2
                B2B2B2B2B2868686B2B2B2868686868686868686FF00FFFF00FFFF00FF868686
                FFFFFFDDDDDDCCCCCC4D4D4D4D4D4D4D4D4D4D4D4D4D4D4DB2B2B2B2B2B28686
                86868686FF00FFFF00FFFF00FF868686FFFFFFFFFFFFDDDDDD4D4D4DDDDDDDCC
                CCCCCCCCCCB2B2B2B2B2B2B2B2B2B2B2B2868686FF00FFFF00FFFF00FF868686
                FFFFFFFFFFFFDDDDDD4D4D4DDDDDDDCCCCCCCCCCCCB2B2B2B2B2B28686868686
                86868686FF00FFFF00FFFF00FF868686FFFFFFFFFFFF4D4D4D4D4D4D4D4D4DDD
                DDDDCCCCCCCCCCCCCCCCCCB2B2B2868686B2B2B2FF00FFFF00FFFF00FF868686
                FFFFFFFFFFFFFFFFFF4D4D4DFFFFFFDDDDDDDDDDDDCCCCCCCCCCCCB2B2B2B2B2
                B2868686FF00FFFF00FFFF00FF868686FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFF00FFFF00FF868686
                8686868686868686868686868686868686868686868686868686868686868686
                86868686FF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFFFFFFFFFFFFFF04
                0404FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
                FFFF00FFFF00FFFF00FF}
              OnClick = btnUp1Click
            end
            object btnRefresh1: TSpeedButton
              Left = 32
              Top = 0
              Width = 25
              Height = 22
              Flat = True
              Glyph.Data = {
                CA020000424DCA0200000000000036000000280000000E0000000F0000000100
                1800000000009402000074120000741200000000000000000000FF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
                FFFF00FF0000FF00FF7E7E7EA3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3
                A3A3A3A3A3A3A3A3A3A3747474FF00FF0000FF00FFA7A7A7E3E3E3DEDEDEDDDD
                DDDDDDDDD7DAD7C7D4C9DDDDDDDDDDDDDDDDDDDDDDDD919191FF00FF0000FF00
                FFACACACEAEAEAE3E3E3DFDFDFDDDDDD4CA0586EAE78DDDDDDDDDDDDDDDDDDDD
                DDDD919191FF00FF0000FF00FFAEAEAEEEEEEEE9E9E9E3E3E330953F00801300
                8013008013DDDDDDDDDDDDDDDDDD919191FF00FF0000FF00FFAEAEAEEEEEEEE9
                E9E9E5E5E5E3E3E34EA25A6EAE78C7C7C72F933EDDDDDDDDDDDD919191FF00FF
                0000FF00FFAEAEAEF6F6F6EAEAEAE9E9E9E9E9E9DCE0DDCCD8CEDDDDDD2F933E
                DDDDDDDDDDDD919191FF00FF0000FF00FFAEAEAEF8F8F8EDEDED85BC8DB7D3BB
                E3E7E4CFDBD1E3E3E3DEDEDEDDDDDDDDDDDD919191FF00FF0000FF00FFAEAEAE
                F8F8F8F8F8F88AC1929DB9A1C8DACB2D943DCFDAD1E1E1E1DDDDDDDDDDDD9191
                91FF00FF0000FF00FFAEAEAEF8F8F8F8F8F8F8F8F83599440080130080130080
                13E3E3E3DFDFDFDDDDDD919191FF00FF0000FF00FFAEAEAEF8F8F8F8F8F8F8F8
                F8F8F8F8D4E6D730963FDCE7DEDADADA8080807A7A7A575757FF00FF0000FF00
                FFAEAEAEF8F8F8F8F8F8F8F8F8F8F8F8F1F4F1DBE8DDF1F1F1D4D4D49F9F9FC6
                C6C63A003AFF00FF0000FF00FFAEAEAEF8F8F8F8F8F8F8F8F8F8F8F8F8F8F8F8
                F8F8F3F3F3DADADAABABAB3D0E3DFF00FFFF00FF0000FF00FF969696B4B4B4B4
                B4B4B4B4B4B4B4B4B4B4B4B4B4B4B4B4B4A8A8A8913991FF00FFFF00FFFF00FF
                0000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FF0000}
              OnClick = btnRefresh1Click
            end
            object btnHome: TSpeedButton
              Left = 64
              Top = 0
              Width = 25
              Height = 22
              Flat = True
              Glyph.Data = {
                36030000424D3603000000000000360000002800000010000000100000000100
                1800000000000003000074120000741200000000000000000000FF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
                FFFF00FFFF00FFFF00FFFF00FFFF00FF00000000000000000000000000000000
                0000000000000000000000000000000000000000FF00FFFF00FFFF00FFFF00FF
                8E7D8ED3C7D3D3C7D3D3C7D38E7D8E8E7D8E6D5B6D000000D3C7D3D3C7D3D3C7
                D3000000FF00FFFF00FFFF00FFFF00FF8E7D8EFFF7FFE4DAE4E4DAE48E7D8EB9
                ABB98E7D8E000000E4DAE4E4DAE4D3C7D3000000FF00FFFF00FFFF00FFFF00FF
                8E7D8EFFF7FFF1E7F1E4DAE48E7D8ED3C7D3B9ABB9000000E4DAE4E4DAE4D3C7
                D3000000FF00FFFF00FFFF00FFFF00FF8E7D8EFFF7FFFFF7FFF1E7F18E7D8EFF
                F7FFD3C7D3000000E4DAE4E4DAE4D3C7D3000000FF00FFFF00FFFF00FFFF00FF
                8E7D8EFFF7FFFFF7FFF1E7F18E7D8E8E7D8E8E7D8E000000E4DAE4E4DAE4D3C7
                D3000000FF00FFFF00FFFF00FF8E7D8E000000B9ABB9FFF7FFFFF7FFF1E7F1F1
                E7F1F1E7F1E4DAE4E4DAE4E4DAE4B9ABB9000000000000FF00FFFF00FF8E7D8E
                F1E7F1000000B9ABB9FFF7FFFFF7FFF1E7F1F1E7F1F1E7F1E4DAE4B9ABB90000
                00C7BBC7000000FF00FFFF00FFFF00FF8E7D8EF1E7F1000000B9ABB9FFF7FFFF
                F7FFF1E7F1F1E7F1B9ABB9000000C7BBC7000000FF00FFFF00FFFF00FFFF00FF
                FF00FF8E7D8EF1E7F1000000B9ABB9FFF7FFFFF7FFB9ABB9000000C7BBC70000
                00000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8E7D8EF1E7F1000000B9
                ABB9B9ABB9000000C7BBC7000000E4DAE4000000FF00FFFF00FFFF00FFFF00FF
                FF00FFFF00FFFF00FF8E7D8EF1E7F1000000000000C7BBC70000008E7D8EE4DA
                E4000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8E7D8EF1
                E7F1C7BBC7000000FF00FF8E7D8EE4DAE4000000FF00FFFF00FFFF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FFFF00FF8E7D8E000000FF00FFFF00FF8E7D8E8E7D
                8E000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
                00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
              OnClick = btnHomeClick
            end
            object btnCollapse1: TSpeedButton
              Left = 96
              Top = 0
              Width = 25
              Height = 22
              Flat = True
              Glyph.Data = {
                66020000424D660200000000000036000000280000000D0000000E0000000100
                1800000000003002000074120000741200000000000000000000FF00FFFF00FF
                FF00FFFF00FF404000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
                FF00FF00FFFF00FFFF00FFFF00FF404000FF00FFFF00FFFF00FFFF00FFFF00FF
                FF00FFFF00FFFF00FF00FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
                00FFFF00FFFF00FFFF00FFFF00FFFF00FF004040004040004040004040004040
                00404000404000404000404000FF00FFFF00FFFF00FFFF00FF00404000FFFFFF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF404000FF00FFFF00FFFF00FFFF00
                FF00404000FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFF404000FF00FF
                FF00FFFF00FFFF00FF00404000FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFF
                FFFF404000FF00FFFF00FFFF00FFFF00FF00404000FFFFFF0000000000000000
                00000000000000FFFFFF404000FF00FF40400040400040400000404000FFFFFF
                FFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFF404000FF00FFFF00FFFF00FFFF00
                FF00404000FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFF404000FF00FF
                FF00FFFF00FFFF00FF00404000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFF404000FF00FFFF00FFFF00FFFF00FF004040004040004040004040004040
                00404000404000404000404000FF00FFFF00FFFF00FFFF00FF00FF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
                FF00FF00FFFF00FFFF00FFFF00FF404000FF00FFFF00FFFF00FFFF00FFFF00FF
                FF00FFFF00FFFF00FF00}
              OnClick = btnCollapse1Click
            end
            object btnHiddenFiles: TSpeedButton
              Left = 128
              Top = 0
              Width = 25
              Height = 22
              Hint = 'Do not show hidden files'
              AllowAllUp = True
              GroupIndex = 1
              Flat = True
              Glyph.Data = {
                0A020000424D0A0200000000000036000000280000000C0000000D0000000100
                180000000000D401000074120000741200000000000000000000402000402000
                402000402000402000402000402000402000402000402000402000402000FFFF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000FFFFFFFF000040ABD9FFFFFFFF00
                00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF40ABD9FF0000FF000040
                ABD9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000FF000040ABD9
                40ABD9FF0000FF0000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF40AB
                D9FF0000FF000040ABD9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                0000FFFFFF40ABD9FF0000FFFFFFFF0000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFFFFC080FFC080FFA060808040808040806020404020404020000000000000
                402000000000FFC080FFC080FFA0608080408080404060604040204040200000
                00000000402000402000FFC080FFC080FFA06080804080804040606040402040
                4020000000000000402000402000}
              ParentShowHint = False
              ShowHint = True
              OnClick = btnHiddenFilesClick
            end
            object btnSystemFiles: TSpeedButton
              Left = 160
              Top = 0
              Width = 25
              Height = 22
              Hint = 'Do not show system files'
              AllowAllUp = True
              GroupIndex = 2
              Flat = True
              Glyph.Data = {
                06030000424D060300000000000036000000280000000F0000000F0000000100
                180000000000D002000074120000741200000000000000000000FF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
                FFFF00FFFF00FF000000FF00FF00000000000000000000000000000000000000
                0000000000000000000000000000000000000000FF00FF000000FF00FF7F7F7F
                BFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBF
                BF000000FF00FF000000FF00FF7F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFBFBFBF000000FF00FF000000FF00FF7F7F7F
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000000000FFFFFFFFFFFFFFFFFFBFBF
                BF000000FF00FF000000FF00FF7F7F7FFFFFFFFFFFFFFFFFFFFFFFFF000000BF
                BFBFBFBFBF000000FFFFFFFFFFFFBFBFBF000000FF00FF000000FF00FF7F7F7F
                FFFFFFFFFFFFFFFFFF7F7F7FFFFFFF0000007F7F7F7F7F7F000000FFFFFFBFBF
                BF000000FF00FF000000FF00FF7F7F7FFFFFFFFFFFFF000000BFBFBF000000FF
                FFFF7F7F7F000000FFFFFFFFFFFFBFBFBF000000FF00FF000000FF00FF7F7F7F
                FFFFFFFFFFFF00000000FFFF007F7F000000FFFFFFFFFFFFFFFFFFFFFFFFBFBF
                BF000000FF00FF000000FF00FF7F7F7FFFFFFFFFFFFF007F7F007F7F000000FF
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFBFBFBF000000FF00FF000000FF00FF7F7F7F
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFBFBF
                BF000000FF00FF000000FF00FF7F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                FFFFFFFFFFFFFFFF000000000000000000000000FF00FF000000FF00FF7F7F7F
                FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFBFBFBFFFFFFF7F7F
                7FFF00FFFF00FF000000FF00FF7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F
                7F7F7F7F7F7F7F7F7F7F7FFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
                FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
                FFFF00FFFF00FF000000}
              ParentShowHint = False
              ShowHint = True
              OnClick = btnSystemFilesClick
            end
          end
        end
        object FilterComboBox1: TFilterComboBox
          Left = 9
          Top = 53
          Width = 199
          Height = 24
          Filter = 'Word files (*.docx)|*.docx|All files (*.*)|*.*'
          TabOrder = 2
          OnClick = FilterComboBox1Click
        end
      end
    end
    object DirView1: TDirView
      Left = 221
      Top = 0
      Width = 476
      Height = 317
      SortColumn = 0
      DateTimeDisplay = dtdDateTimeSec
      Align = alClient
      TabOrder = 1
      ViewStyle = vsReport
      OnClick = DirView1Click
      OnDblClick = DirView1DblClick
      OnKeyUp = DirView1KeyUp
      ColWidthName = 240
      ColWidthSize = 70
      ColWidthType = 60
      ColWidthDate = 150
      ColWidthAttr = 35
      WatchForChanges = False
      UseIconUpdateThread = False
      DimmHiddenFiles = False
      ShowSubDirSize = False
      DirsOnTop = False
      UseIconCache = False
      FileNameDisplay = fndStored
      SingleClickToExec = False
      SelFileSizeFrom = 0
      SelFileTimeFrom = 0
      ShowMaskInHeader = False
      ShowAnimation = False
      SortByExtension = False
      OnAddFile = DirView1AddFile
      OnExecFile = DirView1ExecFile
      Mask = '*.*'
    end
  end
  object pnTop: TPanel
    Left = 0
    Top = 0
    Width = 697
    Height = 33
    Align = alTop
    BevelOuter = bvNone
    Caption = 'pnTop'
    TabOrder = 0
    object CoolBar1: TCoolBar
      Left = 0
      Top = 0
      Width = 697
      Height = 32
      Align = alNone
      Bands = <
        item
          Control = cbHistoryPath
          ImageIndex = -1
          MinHeight = 24
          Width = 697
        end>
      EdgeBorders = [ebTop, ebBottom]
      object cbHistoryPath: TComboBox
        Left = 9
        Top = 0
        Width = 684
        Height = 24
        AutoComplete = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ItemHeight = 16
        ParentFont = False
        TabOrder = 0
        OnClick = cbHistoryPathClick
        OnKeyPress = cbHistoryPathKeyPress
      end
    end
  end
end
