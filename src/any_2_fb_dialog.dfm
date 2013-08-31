object UpenFileModalDialog: TUpenFileModalDialog
  Left = 419
  Top = 256
  Width = 330
  Height = 440
  BorderIcons = [biSystemMenu]
  Caption = 'ANY to FB2 import by GribUser'
  Color = clBtnFace
  Constraints.MinHeight = 440
  Constraints.MinWidth = 330
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Icon.Data = {
    000001000200101010000000000028010000260000002020100000000000E802
    00004E0100002800000010000000200000000100040000000000C00000000000
    0000000000000000000000000000000000000000800000800000008080008000
    0000800080008080000080808000C0C0C0000000FF0000FF000000FFFF00FF00
    0000FF00FF00FFFF0000FFFFFF00000000000000000000000000000000000000
    00000000000000000000000000000AAAAAAA000000000AAAAAAA000000000AAA
    AAAA000000000AAA0000000000000AAA0000000000000AAA0000000000000AAA
    0000000000000AAA0000000000000AAAAAAA000000000AAAAAAA000000000AAA
    AAAA000000000000000000000000FF0F0000FF0F0000FF0F0000000F0000000F
    0000000F0000000F00000000000000000000000000000000000000000000007F
    0000007F0000007F0000007F0000280000002000000040000000010004000000
    0000800200000000000000000000000000000000000000000000000080000080
    00000080800080000000800080008080000080808000C0C0C0000000FF0000FF
    000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000AAAAAAAAAAAAAAA00000000000000000AAAAAAAAAAAAAAA0000
    0000000000000AAAAAAAAAAAAAAA00000000000000000AAAAAAAAAAAAAAA0000
    0000000000000AAAAAAAAAAAAAAA00000000000000000AAAAAAAAAAAAAAA0000
    0000000000000AAAAAAAAAAAAAAA00000000000000000AAAAAAAAAAAAAAA0000
    0000000000000AAAAAAA0000000000000000000000000AAAAAAA000000000000
    0000000000000AAAAAAA0000000000000000000000000AAAAAAA000000000000
    0000000000000AAAAAAA0000000000000000000000000AAAAAAA000000000000
    0000000000000AAAAAAA0000000000000000000000000AAAAAAA000000000000
    0000000000000AAAAAAA0000000000000000000000000AAAAAAAAAAAAAAA0000
    0000000000000AAAAAAAAAAAAAAA00000000000000000AAAAAAAAAAAAAAA0000
    0000000000000AAAAAAAAAAAAAAA00000000000000000AAAAAAAAAAAAAAA0000
    0000000000000AAAAAAAAAAAAAAA00000000000000000AAAAAAAAAAAAAAA0000
    00000000000000000000000000000000000000000000FFFF00FFFFFF00FFFFFF
    00FFFFFF00FFFFFF00FFFFFF00FF000000FF000000FF000000FF000000FF0000
    00FF000000FF000000FF000000FF000000FF0000000000000000000000000000
    0000000000000000000000000000000000000000000000007FFF00007FFF0000
    7FFF00007FFF00007FFF00007FFF00007FFF00007FFF}
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnDeactivate = FormDeactivate
  PixelsPerInch = 120
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 322
    Height = 59
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    OnResize = Panel1Resize
    object Label7: TLabel
      Left = 8
      Top = 40
      Width = 48
      Height = 16
      Caption = 'Settings'
    end
    object Edit1: TEdit
      Left = 8
      Top = 9
      Width = 281
      Height = 24
      TabOrder = 0
      Text = 'http://'
    end
    object Panel5: TPanel
      Left = 286
      Top = 0
      Width = 36
      Height = 59
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
      object Button1: TButton
        Left = 2
        Top = 8
        Width = 26
        Height = 25
        Caption = '...'
        TabOrder = 0
        OnClick = Button1Click
      end
    end
    object ComboBox1: TComboBox
      Left = 64
      Top = 35
      Width = 249
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 2
      OnChange = CheckBox1Click
      Items.Strings = (
        '<Defaults>'
        '<Last used>'
        'Guttenberg'
        'HTML Documentation'
        'Lib.ru'
        'Mad HTML')
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 352
    Width = 322
    Height = 53
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 2
    object Panel3: TPanel
      Left = 7
      Top = 0
      Width = 315
      Height = 53
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object Button2: TButton
        Left = 2
        Top = 0
        Width = 79
        Height = 29
        Caption = '&Import'
        Default = True
        TabOrder = 0
        OnClick = Button2Click
      end
      object Button3: TButton
        Left = 204
        Top = 0
        Width = 69
        Height = 29
        Cancel = True
        Caption = '&Cancel'
        ModalResult = 2
        TabOrder = 2
        OnClick = Button3Click
      end
      object Button4: TButton
        Left = 278
        Top = 0
        Width = 29
        Height = 29
        Caption = '?'
        TabOrder = 3
        OnClick = Button4Click
      end
      object Button5: TButton
        Left = 86
        Top = 0
        Width = 113
        Height = 29
        Caption = 'Save settings...'
        TabOrder = 1
        OnClick = Button5Click
      end
    end
    object CheckBox24: TCheckBox
      Left = 9
      Top = 33
      Width = 305
      Height = 17
      Caption = 'Automaticaly close this window when finished'
      TabOrder = 1
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 59
    Width = 322
    Height = 293
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 7
    TabOrder = 1
    object PageControl1: TPageControl
      Left = 7
      Top = 7
      Width = 308
      Height = 279
      ActivePage = TabSheet1
      Align = alClient
      TabIndex = 0
      TabOrder = 0
      OnResize = Panel5Resize
      object TabSheet1: TTabSheet
        Caption = 'General'
        object CheckBox6: TCheckBox
          Left = 132
          Top = 62
          Width = 131
          Height = 17
          Caption = 'No <emptyline/>'
          TabOrder = 4
        end
        object CheckBox7: TCheckBox
          Left = 6
          Top = 97
          Width = 290
          Height = 17
          Caption = 'Do not convert "quotes" to '#171'quotes'#187
          TabOrder = 7
        end
        object CheckBox15: TCheckBox
          Left = 6
          Top = 115
          Width = 290
          Height = 17
          Caption = 'Do not convert [text] and {text} into footnotes'
          TabOrder = 8
        end
        object CheckBox16: TCheckBox
          Left = 6
          Top = 133
          Width = 290
          Height = 17
          Caption = 'Do not detect _italic_ text'
          TabOrder = 9
        end
        object CheckBox17: TCheckBox
          Left = 6
          Top = 62
          Width = 115
          Height = 17
          Caption = 'No epigraphs'
          TabOrder = 3
        end
        object CheckBox18: TCheckBox
          Left = 6
          Top = 150
          Width = 290
          Height = 17
          Caption = 'Do not restore broken paragraphs'
          TabOrder = 10
        end
        object CheckBox21: TCheckBox
          Left = 6
          Top = 167
          Width = 290
          Height = 17
          Caption = 'Do not detect poems'
          TabOrder = 11
        end
        object CheckBox12: TCheckBox
          Left = 6
          Top = 44
          Width = 123
          Height = 17
          Caption = 'Preserve <form>'
          TabOrder = 1
        end
        object CheckBox8: TCheckBox
          Left = 6
          Top = 220
          Width = 290
          Height = 17
          Caption = 'Deny convert '#39'-'#39' to '#39#8212#39' (long dash in dialogs fe)'
          TabOrder = 14
        end
        object CheckBox14: TCheckBox
          Left = 6
          Top = 202
          Width = 290
          Height = 17
          Caption = 'Ignore line-indent (spaces at the line start)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 13
        end
        object CheckBox5: TCheckBox
          Left = 6
          Top = 184
          Width = 290
          Height = 17
          Caption = 'Only use marked with <h#>|^T^U headers'
          TabOrder = 12
        end
        object CheckBox9: TCheckBox
          Left = 132
          Top = 44
          Width = 165
          Height = 17
          Caption = 'Do not convert charset'
          TabOrder = 2
        end
        object RadioGroup1: TRadioGroup
          Left = 6
          Top = -1
          Width = 288
          Height = 39
          Caption = 'Text structure'
          Columns = 3
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemIndex = 0
          Items.Strings = (
            'Auto'
            'Indented'
            'Empty lines')
          ParentFont = False
          TabOrder = 0
          OnClick = RadioGroup1Click
        end
        object CheckBox20: TCheckBox
          Left = 132
          Top = 79
          Width = 155
          Height = 17
          Caption = 'Allow 1000 error fixes'
          TabOrder = 6
        end
        object CheckBox19: TCheckBox
          Left = 6
          Top = 80
          Width = 115
          Height = 17
          Caption = 'No description'
          TabOrder = 5
        end
      end
      object TabSheet2: TTabSheet
        Caption = 'Links'
        ImageIndex = 1
        object Panel7: TPanel
          Left = 0
          Top = 0
          Width = 300
          Height = 248
          Align = alClient
          BevelOuter = bvNone
          BorderWidth = 4
          TabOrder = 0
          object GroupBox1: TGroupBox
            Left = 4
            Top = 4
            Width = 292
            Height = 77
            Align = alTop
            Caption = 'Images'
            TabOrder = 0
            object CheckBox11: TCheckBox
              Left = 8
              Top = 18
              Width = 274
              Height = 17
              Caption = 'Remove ALL images from the document'
              TabOrder = 0
              OnClick = CheckBox11Click
              OnKeyPress = CheckBox11KeyPress
            end
            object CheckBox3: TCheckBox
              Left = 8
              Top = 36
              Width = 250
              Height = 17
              Caption = 'Remove off-site images'
              TabOrder = 1
            end
            object CheckBox4: TCheckBox
              Left = 8
              Top = 54
              Width = 250
              Height = 17
              Caption = 'Preserve dynamic images'
              TabOrder = 2
            end
          end
          object GroupBox2: TGroupBox
            Left = 4
            Top = 81
            Width = 292
            Height = 158
            Align = alTop
            Caption = 'Linked documents'
            TabOrder = 1
            object Label3: TLabel
              Left = 214
              Top = 44
              Width = 71
              Height = 16
              Caption = 'levels deep'
            end
            object Label6: TLabel
              Left = 9
              Top = 111
              Width = 223
              Height = 16
              Caption = 'Never follow matching this expression'
            end
            object Label5: TLabel
              Left = 8
              Top = 70
              Width = 213
              Height = 16
              Caption = 'Only follow matching this expression'
            end
            object CheckBox2: TCheckBox
              Left = 8
              Top = 18
              Width = 182
              Height = 17
              Caption = 'Remove external links'
              TabOrder = 0
              OnClick = CheckBox2Click
              OnKeyPress = CheckBox2KeyPress
            end
            object CheckBox22: TCheckBox
              Left = 8
              Top = 36
              Width = 97
              Height = 17
              Caption = 'Follow links'
              TabOrder = 1
              OnClick = CheckBox22Click
              OnKeyPress = CheckBox22KeyPress
            end
            object UpDown1: TUpDown
              Left = 193
              Top = 41
              Width = 18
              Height = 24
              Associate = Edit2
              Min = 1
              Max = 50
              Position = 1
              TabOrder = 2
              Wrap = False
            end
            object Edit2: TEdit
              Left = 161
              Top = 41
              Width = 32
              Height = 24
              TabOrder = 3
              Text = '1'
            end
            object CheckBox23: TCheckBox
              Left = 8
              Top = 54
              Width = 143
              Height = 17
              Caption = 'Follow off-site links'
              TabOrder = 4
            end
            object Edit5: TEdit
              Left = 8
              Top = 127
              Width = 275
              Height = 24
              TabOrder = 6
            end
            object Edit4: TEdit
              Left = 8
              Top = 86
              Width = 276
              Height = 24
              TabOrder = 5
            end
          end
        end
      end
      object TabSheet4: TTabSheet
        Caption = 'Regexp'
        ImageIndex = 3
        object Label1: TLabel
          Left = 0
          Top = 41
          Width = 300
          Height = 16
          Align = alTop
          Caption = 'Regular expressions to run on-load'
        end
        object Label2: TLabel
          Left = 0
          Top = 138
          Width = 300
          Height = 16
          Align = alTop
          Caption = 'Regular expressions to run on result document'
        end
        object Panel6: TPanel
          Left = 0
          Top = 0
          Width = 300
          Height = 41
          Align = alTop
          BevelOuter = bvNone
          TabOrder = 0
          OnResize = Panel6Resize
          object Label4: TLabel
            Left = 0
            Top = 0
            Width = 300
            Height = 16
            Align = alTop
            Caption = 'Headers detection regular expression'
          end
          object Edit3: TEdit
            Left = 0
            Top = 16
            Width = 298
            Height = 24
            TabOrder = 0
          end
        end
        object StringGrid1: TStringGrid
          Left = 0
          Top = 154
          Width = 300
          Height = 94
          Align = alClient
          ColCount = 2
          DefaultColWidth = 130
          DefaultRowHeight = 18
          FixedCols = 0
          RowCount = 1
          FixedRows = 0
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goAlwaysShowEditor]
          ScrollBars = ssVertical
          TabOrder = 2
          OnEnter = StringGrid2Enter
          OnGetEditText = StringGrid1GetEditText
          OnSetEditText = StringGrid1SetEditText
        end
        object StringGrid2: TStringGrid
          Left = 0
          Top = 57
          Width = 300
          Height = 81
          Align = alTop
          ColCount = 2
          DefaultColWidth = 130
          DefaultRowHeight = 18
          FixedCols = 0
          RowCount = 1
          FixedRows = 0
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goAlwaysShowEditor]
          ScrollBars = ssVertical
          TabOrder = 1
          OnEnter = StringGrid2Enter
          OnGetEditText = StringGrid1GetEditText
          OnSetEditText = StringGrid1SetEditText
          OnTopLeftChanged = StringGrid2TopLeftChanged
        end
        object Button6: TButton
          Left = 264
          Top = 60
          Width = 18
          Height = 18
          Caption = '...'
          TabOrder = 3
          TabStop = False
          Visible = False
          OnClick = Button6Click
        end
      end
      object TabSheet3: TTabSheet
        Caption = 'Log'
        ImageIndex = 2
        object ProgressBar1: TProgressBar
          Left = 0
          Top = 231
          Width = 300
          Height = 17
          Align = alBottom
          Min = 0
          Max = 100
          TabOrder = 0
        end
        object RichEdit1: TRichEdit
          Left = 0
          Top = 0
          Width = 300
          Height = 231
          Align = alClient
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = True
          HideSelection = False
          Lines.Strings = (
            'Ready...')
          ParentCtl3D = False
          PopupMenu = PopupMenu1
          ReadOnly = True
          ScrollBars = ssBoth
          TabOrder = 1
          WantReturns = False
          WordWrap = False
        end
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 
      'All supported documents (*.html; *.htm; *.txt; *.prt; *.rtf; *.d' +
      'oc; *.dot; *.wri; *.wk1; *.wk3; *.wk4; *.mcw)|*.html;*.htm;*.txt' +
      ';*.prt;*.rtf;*.doc;*.dot;*.wri;*.wk1;*.wk3;*.wk4;*.mcw;|Native s' +
      'uported txt/html (*.html; *.htm; *.txt; *.prt)|*.html;*.htm;*.tx' +
      't;*.prt;|MSWord documents (*.rtf; *.doc; *.dot; *.wri; *.wk1; *.' +
      'wk3; *.wk4; *.mcw)|*.rtf;*.doc;*.dot;*.wri;*.wk1;*.wk3;*.wk4;*.m' +
      'cw;|All files (*.*)|*.*'
    Left = 38
    Top = 12
  end
  object PopupMenu1: TPopupMenu
    Left = 232
    Top = 416
    object Copy1: TMenuItem
      Caption = '&Copy the whole log'
      OnClick = Copy1Click
    end
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = 'fb2'
    Filter = 'FictionBook 2.0 files (*.fb2)|*.fb2|All files (*.*)|*.*'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 64
    Top = 10
  end
  object Timer1: TTimer
    Enabled = False
    Interval = 200
    OnTimer = Timer1Timer
    Left = 11
    Top = 13
  end
end
