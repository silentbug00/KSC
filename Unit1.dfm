object Form1: TForm1
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Form1'
  ClientHeight = 395
  ClientWidth = 536
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 536
    Height = 395
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 532
    ExplicitHeight = 383
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object Memo1: TMemo
        Left = 0
        Top = 177
        Width = 528
        Height = 190
        Align = alClient
        Lines.Strings = (
          'Memo1')
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 2
        ExplicitWidth = 524
        ExplicitHeight = 178
      end
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 528
        Height = 133
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 0
        ExplicitWidth = 524
        object Label1: TLabel
          Left = 34
          Top = 24
          Width = 45
          Height = 13
          Alignment = taRightJustify
          Caption = 'Base URL'
        end
        object Label4: TLabel
          Left = 37
          Top = 105
          Width = 42
          Height = 13
          Alignment = taRightJustify
          Caption = 'Filename'
        end
        object Label3: TLabel
          Left = 33
          Top = 78
          Width = 46
          Height = 13
          Alignment = taRightJustify
          Caption = 'Password'
        end
        object Label2: TLabel
          Left = 57
          Top = 51
          Width = 22
          Height = 13
          Alignment = taRightJustify
          Caption = 'User'
        end
        object Edit1: TEdit
          Left = 85
          Top = 21
          Width = 416
          Height = 21
          TabOrder = 0
          Text = 'https://127.0.0.1:57512/config'
        end
        object Edit2: TEdit
          Left = 85
          Top = 48
          Width = 212
          Height = 21
          TabOrder = 1
          Text = 'Administrator'
        end
        object Edit3: TEdit
          Left = 85
          Top = 75
          Width = 212
          Height = 21
          PasswordChar = #9679
          TabOrder = 2
        end
        object Edit4: TEdit
          Left = 85
          Top = 102
          Width = 389
          Height = 21
          TabOrder = 3
          Text = 'D:\OPC\RTUList.xlsx'
        end
        object Button1: TButton
          Left = 480
          Top = 100
          Width = 21
          Height = 25
          Caption = '...'
          TabOrder = 4
          OnClick = Button1Click
        end
      end
      object Panel2: TPanel
        Left = 0
        Top = 133
        Width = 528
        Height = 44
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 1
        ExplicitWidth = 524
        object Button2: TButton
          Left = 426
          Top = 6
          Width = 75
          Height = 25
          Caption = 'Process'
          TabOrder = 0
          OnClick = Button2Click
        end
        object Button3: TButton
          Left = 302
          Top = 6
          Width = 75
          Height = 25
          Caption = 'Cancel'
          TabOrder = 1
          Visible = False
          OnClick = Button3Click
        end
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 476
    Top = 88
  end
  object LbBlowfish1: TLbBlowfish
    CipherMode = cmECB
    Left = 300
    Top = 24
  end
end
