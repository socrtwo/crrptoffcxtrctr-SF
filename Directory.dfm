object GetFolderName: TGetFolderName
  Tag = 1
  Left = 193
  Top = 29
  Width = 572
  Height = 559
  BorderIcons = [biSystemMenu, biMaximize]
  BorderWidth = 8
  Caption = 'Directory'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 120
  TextHeight = 16
  object PageControl1: TPageControl
    Left = 0
    Top = 298
    Width = 548
    Height = 213
    ActivePage = tabExtract
    Align = alBottom
    Style = tsButtons
    TabOrder = 0
    object tabExtract: TTabSheet
      Caption = 'Extract to'
      ImageIndex = 1
      object gbBottom: TGroupBox
        Left = 0
        Top = 0
        Width = 540
        Height = 179
        Align = alRight
        Caption = 
          '   Extract to below folder. If bold text found, new path will be' +
          ' created '
        TabOrder = 0
        object edToPath: TEdit
          Left = 25
          Top = 32
          Width = 456
          Height = 24
          TabOrder = 0
          Text = 'c:\Temp'
          OnChange = edToPathChange
        end
        object btnQuit: TBitBtn
          Left = 392
          Top = 112
          Width = 81
          Height = 28
          Caption = '&Cancel'
          TabOrder = 4
          OnClick = btnQuitClick
          Kind = bkCancel
        end
        object btnExtract: TButton
          Left = 392
          Top = 72
          Width = 81
          Height = 28
          Caption = '&Extract'
          TabOrder = 3
          OnClick = btnExtractClick
        end
        object GroupBox2: TGroupBox
          Left = 169
          Top = 64
          Width = 200
          Height = 89
          Caption = ' Options '
          TabOrder = 2
          object Label2: TLabel
            Left = 24
            Top = 68
            Width = 161
            Height = 16
            AutoSize = False
            Caption = '(no check = skip if older)'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'MS Sans Serif'
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
        object GroupBox1: TGroupBox
          Left = 25
          Top = 64
          Width = 136
          Height = 89
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
      end
    end
  end
  object pnShellList: TPanel
    Left = 0
    Top = 62
    Width = 548
    Height = 236
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 1
    TabOrder = 1
    object Splitter1: TSplitter
      Left = 186
      Top = 1
      Height = 234
    end
    object ShellListView1: TShellListView
      Left = 189
      Top = 1
      Width = 358
      Height = 234
      ObjectTypes = [otFolders, otNonFolders]
      Root = 'C:\'
      ShellTreeView = stvDirectory
      Sorted = True
      Align = alClient
      ReadOnly = False
      HideSelection = False
      OnChange = ShellListView1Change
      TabOrder = 0
      ViewStyle = vsReport
    end
    object btnOK: TButton
      Left = 412
      Top = 8
      Width = 75
      Height = 25
      Caption = 'OK'
      TabOrder = 1
      Visible = False
      OnClick = btnOKClick
    end
    object btnCancel: TButton
      Left = 412
      Top = 40
      Width = 75
      Height = 25
      Caption = 'Cancel'
      TabOrder = 2
      Visible = False
    end
    object pnTree: TPanel
      Left = 1
      Top = 1
      Width = 185
      Height = 234
      Align = alLeft
      BevelOuter = bvNone
      TabOrder = 3
      object stvDirectory: TShellTreeView
        Left = 0
        Top = 30
        Width = 185
        Height = 204
        ObjectTypes = [otFolders]
        Root = 'C:\'
        ShellListView = ShellListView1
        UseShellImages = True
        Align = alClient
        AutoRefresh = False
        Indent = 19
        ParentColor = False
        RightClickSelect = True
        ShowRoot = False
        TabOrder = 0
        OnMouseDown = stvDirectoryMouseDown
        OnKeyDown = stvDirectoryKeyDown
      end
      object CoolBar2: TCoolBar
        Left = 0
        Top = 0
        Width = 185
        Height = 30
        Bands = <>
        object btnUp1: TSpeedButton
          Left = 5
          Top = 0
          Width = 25
          Height = 25
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
          Left = 69
          Top = 0
          Width = 25
          Height = 25
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
        object btnCollapse1: TSpeedButton
          Left = 37
          Top = 0
          Width = 25
          Height = 25
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
      end
    end
  end
  object CoolBar1: TCoolBar
    Left = 0
    Top = 0
    Width = 548
    Height = 62
    AutoSize = True
    Bands = <
      item
        Control = cbHistoryPath
        ImageIndex = -1
        MinHeight = 24
        Width = 544
      end
      item
        Control = pnCB
        ImageIndex = -1
        MinHeight = 32
        Width = 544
      end>
    object pnCB: TPanel
      Left = 9
      Top = 26
      Width = 531
      Height = 32
      Align = alClient
      BevelOuter = bvNone
      BevelWidth = 3
      TabOrder = 1
      object lblPath: TLabel
        Left = 240
        Top = 8
        Width = 33
        Height = 16
        AutoSize = False
        Caption = 'lblPath'
        Visible = False
        WordWrap = True
      end
      object lblPreDrive: TLabel
        Left = 288
        Top = 8
        Width = 33
        Height = 16
        AutoSize = False
        Caption = 'lblPreDrive'
        Visible = False
        WordWrap = True
      end
      object lblGetDirName: TLabel
        Left = 336
        Top = 8
        Width = 33
        Height = 16
        AutoSize = False
        Caption = 'lblGetDirName'
        Visible = False
        WordWrap = True
      end
      object lblPathFilename: TLabel
        Left = 384
        Top = 8
        Width = 41
        Height = 16
        AutoSize = False
        Caption = 'lblPathFilename'
        Visible = False
        WordWrap = True
      end
      object dcDrive: TDriveComboBox
        Left = 0
        Top = 6
        Width = 169
        Height = 22
        TabOrder = 0
        OnChange = dcDriveChange
      end
    end
    object cbHistoryPath: TComboBox
      Left = 9
      Top = 0
      Width = 531
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
      OnDropDown = cbHistoryPathDropDown
      OnKeyPress = cbHistoryPathKeyPress
    end
  end
end
