object ManagePresets: TManagePresets
  Left = 577
  Top = 188
  Width = 200
  Height = 250
  Caption = 'Manage export presets'
  Color = clBtnFace
  Constraints.MinHeight = 250
  Constraints.MinWidth = 200
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 120
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 118
    Height = 215
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object ListBox1: TListBox
      Left = 0
      Top = 25
      Width = 118
      Height = 190
      Align = alClient
      ItemHeight = 16
      TabOrder = 0
      OnClick = ListBox1Click
      OnKeyPress = ListBox1KeyPress
    end
    object Panel2: TPanel
      Left = 0
      Top = 0
      Width = 118
      Height = 25
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 1
      OnResize = Panel2Resize
      object Edit1: TEdit
        Left = 0
        Top = 1
        Width = 190
        Height = 24
        TabOrder = 0
        Text = 'Untitled'
      end
    end
  end
  object Panel3: TPanel
    Left = 118
    Top = 0
    Width = 74
    Height = 215
    Align = alRight
    BevelOuter = bvNone
    TabOrder = 1
    object Button1: TButton
      Left = 4
      Top = 8
      Width = 66
      Height = 25
      Caption = 'Save'
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 4
      Top = 36
      Width = 66
      Height = 25
      Caption = 'Delete...'
      TabOrder = 1
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 4
      Top = 79
      Width = 66
      Height = 25
      Cancel = True
      Caption = 'Close'
      ModalResult = 1
      TabOrder = 2
    end
  end
end
