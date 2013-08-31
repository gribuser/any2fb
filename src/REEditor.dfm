object ReEditorForm: TReEditorForm
  Left = 449
  Top = 204
  Width = 300
  Height = 250
  ActiveControl = Edit1
  Caption = 'RE Editor'
  Color = clBtnFace
  Constraints.MinHeight = 150
  Constraints.MinWidth = 300
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnResize = FormResize
  PixelsPerInch = 120
  TextHeight = 16
  object Memo2: TMemo
    Left = 0
    Top = 58
    Width = 292
    Height = 157
    Align = alClient
    TabOrder = 1
  end
  object Panel3: TPanel
    Left = 0
    Top = 0
    Width = 292
    Height = 58
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object Label1: TLabel
      Left = 2
      Top = 0
      Width = 68
      Height = 16
      Caption = 'Search RE:'
    end
    object Label2: TLabel
      Left = 2
      Top = 41
      Width = 77
      Height = 16
      Caption = 'Replace RE:'
    end
    object Panel2: TPanel
      Left = 207
      Top = 0
      Width = 85
      Height = 58
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
      object Button1: TButton
        Left = 7
        Top = 2
        Width = 75
        Height = 25
        Caption = 'Ok'
        Default = True
        ModalResult = 1
        TabOrder = 0
      end
      object Button2: TButton
        Left = 7
        Top = 30
        Width = 75
        Height = 25
        Cancel = True
        Caption = 'Cancel'
        ModalResult = 2
        TabOrder = 1
      end
    end
    object Edit1: TEdit
      Left = 0
      Top = 16
      Width = 163
      Height = 24
      TabOrder = 0
      Text = 'Edit1'
    end
  end
end
