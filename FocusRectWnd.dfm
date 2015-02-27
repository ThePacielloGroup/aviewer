object WndFocusRect: TWndFocusRect
  Left = 0
  Top = 0
  BorderIcons = []
  BorderStyle = bsNone
  ClientHeight = 48
  ClientWidth = 221
  Color = clWindow
  Ctl3D = False
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  OnClose = FormClose
  OnDeactivate = FormDeactivate
  DesignSize = (
    221
    48)
  PixelsPerInch = 96
  TextHeight = 13
  object Shape1: TShape
    Left = 0
    Top = 0
    Width = 221
    Height = 48
    Align = alClient
    Brush.Color = clRed
    ExplicitWidth = 345
    ExplicitHeight = 65
  end
  object Shape2: TShape
    Left = 5
    Top = 5
    Width = 211
    Height = 38
    Anchors = [akLeft, akTop, akRight, akBottom]
  end
end
