object frmToFolder: TfrmToFolder
  Left = 256
  Top = 213
  Width = 581
  Height = 603
  BorderIcons = [biSystemMenu]
  BorderWidth = 7
  Caption = 'Extract'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object pnBottom: TPanel
    Left = 0
    Top = 391
    Width = 549
    Height = 155
    Align = alBottom
    BevelOuter = bvNone
    BorderWidth = 4
    TabOrder = 0
    object gbExtract: TGroupBox
      Left = 4
      Top = 4
      Width = 541
      Height = 147
      Align = alClient
      Caption = 
        ' Extract to below directory. If bold text found, new directory w' +
        'ill be created. '
      TabOrder = 0
      object edToPath: TEdit
        Left = 16
        Top = 24
        Width = 505
        Height = 24
        TabOrder = 0
        Text = 'c:\Temp'
        OnChange = edToPathChange
      end
      object GroupBox2: TGroupBox
        Left = 17
        Top = 64
        Width = 136
        Height = 73
        Caption = ' Files to extract '
        TabOrder = 1
        object rbtnAllFiles: TRadioButton
          Left = 16
          Top = 25
          Width = 81
          Height = 16
          Caption = 'All files'
          TabOrder = 0
        end
        object rbSelectedFiles: TRadioButton
          Left = 16
          Top = 41
          Width = 113
          Height = 16
          Caption = 'Selected files'
          TabOrder = 1
        end
      end
      object GroupBox3: TGroupBox
        Left = 185
        Top = 52
        Width = 208
        Height = 85
        Caption = ' Options '
        TabOrder = 2
        object Label2: TLabel
          Left = 48
          Top = 62
          Width = 153
          Height = 16
          AutoSize = False
          Caption = '( skip if older when disabled )'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -10
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object chkWithDir: TCheckBox
          Left = 16
          Top = 25
          Width = 137
          Height = 16
          Caption = 'with directories'
          Checked = True
          State = cbChecked
          TabOrder = 0
        end
        object chkAlwaysOverwrite: TCheckBox
          Left = 16
          Top = 44
          Width = 153
          Height = 17
          Caption = 'always overwrite'
          Checked = True
          State = cbChecked
          TabOrder = 1
        end
      end
      object btnOK: TBitBtn
        Left = 432
        Top = 72
        Width = 83
        Height = 28
        Caption = '&Extract'
        TabOrder = 3
        OnClick = btnOKClick
        Glyph.Data = {
          06030000424D060300000000000036000000280000000F0000000F0000000100
          180000000000D002000074120000741200000000000000000000FF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF62478325305CFF00
          FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFA0
          0FAB4B4C93A9A2C293BBD327444D57416AFF00FFFF00FF000000FF00FFFF00FF
          FF00FFFF00FFFF00FFA509B36A46957D96B499D5F2D6CCEA9BC0D9455D631F34
          3F6C597FFF00FF000000FF00FFFF00FFFF00FF800D947650907AA8C8B8D1EEBD
          D2EF9FDBF6D0CDEB90C5DC598388162222476262FF00FF000000FF00FFFF00FF
          8B0D9388B4CECAC6E5A9D7F4BBD3EFBDD2EF9FDBF6D0CDEB96C0D444B5BF1A2C
          349E45AEFF00FF000000FF00FFFF00FF854288A8D8F4D2CDEBA9D8F4B2D6F2B9
          D4F0B8D4F0BAD3F075A8B33DDBE12B3B40FF00FFFF00FF000000FF00FFFF00FF
          6A90AFCACFEC95DFF8B2D8F2B7D4F1B5D5F1D3CCEAA2DAF55FAAB33EECEE3A45
          5F9E08A2FF00FF000000FF00FF762075A2CCDFC1D9F1ACD9F3C2D5EEBCD3EFB9
          D3F0C2D1EE99D3EFBBB7CBD0F9FA576773531D5FFF00FF000000FF00FF5B275C
          D2D5E3A6E2F4B0DAF1B8D5EDB5DFF5C0D6F0C6CCEA9ABAD0C1CFD7E6FEFC687B
          84485169FF00FF000000990099537E85E0DAEDB6E7F595EBF4B0DBF1A3D2E4A0
          B1C3A4B6C2BCDBDDE1F9F8E6FEFC7992923B5E77FF00FF0000009E069F8C8B92
          E2EEF59FE5EEA4D3E085C5CAD8C3D7CBE1E2B2C4C3B8CBCAB8CBCAB5C9C76576
          7D44616FFF00FF000000B31CB5A9CAC978A5A492B0B1BADBDBDEF4F2E6FEFCB8
          CBCA5C5C5C9999999999999999999999999999993D3D3D0000003A1A3F832F88
          FF00FF93A2A1E6FEFCE4FCFADDECEF91A2A9999999FFFFFFFFFFFFFFFFFFFFFF
          FFF4EBF44F3D4F000000FF00FFFF00FFFF00FF849392C1BFCA8A8EA66A8AA04D
          5576999999FFFFFFFFFFFFFFFFFFFFFFFFB57AB55B525B000000FF00FFFF00FF
          FF00FF47234CFF00FF4F0D595422669208933D3D3D6666666666666666666666
          66442944292929000000}
      end
      object btnCancel: TBitBtn
        Left = 432
        Top = 104
        Width = 83
        Height = 28
        Caption = '&Cancel'
        TabOrder = 4
        OnClick = btnCancelClick
        Kind = bkCancel
      end
    end
  end
  object pnMiddle: TPanel
    Left = 0
    Top = 33
    Width = 549
    Height = 358
    Align = alClient
    BevelOuter = bvNone
    Caption = 'pnMiddle'
    TabOrder = 2
    object Splitter1: TSplitter
      Left = 218
      Top = 0
      Height = 358
    end
    object DirView1: TDirView
      Left = 221
      Top = 0
      Width = 328
      Height = 358
      SortColumn = 0
      DateTimeDisplay = dtdDateTimeSec
      Align = alClient
      TabOrder = 1
      ViewStyle = vsReport
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
      Mask = '*.*'
    end
    object pnTree: TPanel
      Left = 0
      Top = 0
      Width = 218
      Height = 358
      Align = alLeft
      Caption = 'pnTree'
      TabOrder = 0
      object DriveView1: TDriveView
        Left = 1
        Top = 56
        Width = 216
        Height = 301
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
        Height = 55
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
            object btnRefresh1: TSpeedButton
              Left = 30
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
          end
        end
      end
    end
  end
  object pnTop: TPanel
    Left = 0
    Top = 0
    Width = 549
    Height = 33
    Align = alTop
    BevelOuter = bvNone
    Caption = 'pnTop'
    TabOrder = 1
    object CoolBar1: TCoolBar
      Left = 0
      Top = 0
      Width = 540
      Height = 35
      Align = alNone
      Bands = <
        item
          Control = cbHistoryPath
          ImageIndex = -1
          MinHeight = 24
          Width = 540
        end>
      EdgeBorders = [ebTop, ebBottom]
      object cbHistoryPath: TComboBox
        Left = 9
        Top = 0
        Width = 527
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
