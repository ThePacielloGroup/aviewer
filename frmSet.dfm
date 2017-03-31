object WndSet: TWndSet
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Settings'
  ClientHeight = 366
  ClientWidth = 534
  Color = clBtnFace
  DoubleBuffered = True
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Icon.Data = {
    0000010001001010000000000000680300001600000028000000100000002000
    0000010018000000000000000000480000004800000000000000000000002DDD
    FC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2D
    DDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFCFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF2DDDFC2DDD
    FCFFFFFFFCA5B8D25F71D56476D36173E47C8FFFE5EEFFE5EEE47C8FD36173D5
    6476D25F71FCA5B8FFFFFF2DDDFC2DDDFCFFFFFFD56476D56476D56476D56476
    D25A6DECBAC2ECBAC2D25A6DD56476D56476D56476D56476FFFFFF2DDDFC2DDD
    FCFFFFFFD56476D56476D56476D56476D25A6DECBAC2ECBAC2D25A6DD56476D5
    6476D56476D56476FFFFFF2DDDFC2DDDFCFFFFFFD56476D56476D56476D56476
    D25A6DECBAC2ECBAC2D25A6DD56476D56476D56476D56476FFFFFF2DDDFC2DDD
    FCFFFFFFE67A8CD46274D56476D46274DE6B7EFBD0D8FBD0D8DE6B7ED46274D5
    6476D46274E67A8CFFFFFF2DDDFC2DDDFCFFFFFFFFFEFEFFD8E6FFDBE8FFD9E7
    FFE8F0FFFFFFFFFFFFFFE8F0FFD9E7FFDBE8FFD8E6FFFEFEFFFFFF2DDDFC2DDD
    FCFFFFFFFFFEFEFFD8E6FFDBE8FFD9E7FFE8F0FFFFFFFFFEFFE5F5FFD5F2FFD7
    F2FFD4F1FFFEFBFFFFFFFF2DDDFC2DDDFCFFFFFFE67A8CD46274D56476D46274
    DE6B7EFFC9CFBAF9FF1CD5FF00D1FF00D2FF00D2FF43D8FFFFFFFF2DDDFC2DDD
    FCFFFFFFD56476D56476D56476D56476D25A6DFFB0B67AF4FF00CFFF00D2FF00
    D2FF00D2FF00D2FFFFFFFF2DDDFC2DDDFCFFFFFFD56476D56476D56476D56476
    D25A6DFFB0B67AF4FF00CFFF00D2FF00D2FF00D2FF00D2FFFFFFFF2DDDFC2DDD
    FCFFFFFFD56476D56476D56476D56476D25A6DFFB0B67AF4FF00CFFF00D2FF00
    D2FF00D2FF00D2FFFFFFFF2DDDFC2DDDFCFFFFFFFCA5B8D25F71D56476D36173
    E47C8FFFE1E9E2FCFF38D9FF00D1FF00D2FF00D1FF95E5FFFFFFFF2DDDFC2DDD
    FCFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
    FFFFFFFFFFFFFFFFFFFFFF2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC
    2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC2DDDFC0000
    0000000000000000000000000000000000000000000000000000000000000000
    000000000000000000000000000000000000000000000000000000000000}
  OldCreateOrder = False
  Position = poMainFormCenter
  ShowHint = True
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object gbSelProp: TAccGroupBox
    Left = 8
    Top = 8
    Width = 297
    Height = 353
    Caption = 'Use of properties'
    TabOrder = 0
    TabStop = True
    CtrlFirstChild = chkList
    object chkList: TAccCheckList
      Left = 16
      Top = 16
      Width = 265
      Height = 313
      HeaderColor = clWindow
      HeaderBackgroundColor = clWindowText
      ItemHeight = 13
      TabOrder = 0
      CtrlRight = btnFont
    end
  end
  object gbFont: TAccGroupBox
    Left = 311
    Top = 8
    Width = 214
    Height = 153
    Caption = 'Font'
    TabOrder = 1
    TabStop = True
    CtrlFirstChild = btnFont
    object btnFont: TAccButton
      Left = 16
      Top = 16
      Width = 89
      Height = 25
      Caption = '&Font...'
      TabOrder = 0
      OnClick = btnFontClick
      CtrlNext = Memo1
      CtrlPrev = chkList
      CtrlLeft = chkList
      CtrlDown = Memo1
    end
    object Memo1: TAccMemo
      Left = 16
      Top = 55
      Width = 185
      Height = 82
      Alignment = taCenter
      ReadOnly = True
      TabOrder = 1
      CtrlNext = btnReg
      CtrlPrev = btnFont
      CtrlUp = btnFont
    end
  end
  object gbRegDll: TAccGroupBox
    Left = 311
    Top = 167
    Width = 122
    Height = 98
    Caption = 'IAccessible2Proxy.dll'
    TabOrder = 2
    TabStop = True
    object btnReg: TAccButton
      Left = 16
      Top = 24
      Width = 89
      Height = 25
      Caption = '&Register'
      ElevationRequired = True
      TabOrder = 0
      OnClick = btnRegClick
      CtrlNext = btnUnreg
      CtrlPrev = Memo1
      CtrlLeft = chkList
      CtrlUp = Memo1
    end
    object btnUnreg: TAccButton
      Left = 15
      Top = 55
      Width = 89
      Height = 25
      Caption = '&Unregister'
      ElevationRequired = True
      TabOrder = 1
      OnClick = btnUnregClick
      CtrlNext = btnOK
      CtrlPrev = gbRegDll
      CtrlLeft = chkList
      CtrlUp = btnReg
      CtrlDown = btnOK
    end
  end
  object btnOK: TAccButton
    Left = 450
    Top = 302
    Width = 75
    Height = 25
    Caption = '&OK'
    ModalResult = 1
    TabOrder = 3
    CtrlNext = btnCancel
    CtrlDown = btnCancel
  end
  object btnCancel: TAccButton
    Left = 451
    Top = 333
    Width = 75
    Height = 25
    Caption = '&Cancel'
    Default = True
    ModalResult = 2
    TabOrder = 4
    CtrlPrev = btnOK
    CtrlUp = btnOK
  end
  object cbExTip: TTransCheckBox
    Left = 311
    Top = 271
    Width = 210
    Height = 17
    Caption = 'Expansion tooltip'
    TabOrder = 5
    Visible = False
  end
  object FontDialog1: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = []
    Left = 360
    Top = 24
  end
end
