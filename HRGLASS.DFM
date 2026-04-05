object HourGlass: THourGlass
  Left = 298
  Top = 236
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'HourGlass'
  ClientHeight = 134
  ClientWidth = 222
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object lblOnProcess: TLabel
    Left = 16
    Top = 8
    Width = 193
    Height = 49
    AutoSize = False
    Caption = 'lblOnProcess'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    WordWrap = True
  end
  object Gauge1: TGauge
    Left = 0
    Top = 112
    Width = 193
    Height = 22
    Progress = 0
    Visible = False
  end
  object Gauge2: TGauge
    Left = 198
    Top = 114
    Width = 18
    Height = 18
    BackColor = cl3DLight
    BorderStyle = bsNone
    ForeColor = clGray
    Kind = gkPie
    Progress = 0
    ShowText = False
  end
  object Animate1: TAnimate
    Left = 72
    Top = 64
    Width = 41
    Height = 41
  end
  object btnCancel: TButton
    Left = 136
    Top = 80
    Width = 75
    Height = 25
    Caption = 'Cancel'
    TabOrder = 1
    OnClick = btnCancelClick
  end
end
