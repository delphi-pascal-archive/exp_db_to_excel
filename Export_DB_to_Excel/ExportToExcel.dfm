object Form1: TForm1
  Left = 224
  Top = 133
  Width = 641
  Height = 489
  Caption = #1069#1082#1089#1087#1086#1088#1090' '#1073#1072#1079' '#1076#1072#1085#1085#1099#1093' '#1074' Excel'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Icon.Data = {
    0000010001002020100000000000E80200001600000028000000200000004000
    0000010004000000000080020000000000000000000000000000000000000000
    0000000080000080000000808000800000008000800080800000C0C0C0008080
    80000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00CCC0
    000CCCC0000000000CCCC7777CCCCCCC0000CCCC00000000CCCC7777CCCCCCCC
    C0000CCCCCCCCCCCCCC7777CCCCC0CCCCC0000CCCCCCCCCCCC7777CCCCC700CC
    C00CCCC0000000000CCCC77CCC77000C0000CCCC00000000CCCC7777C7770000
    00000CCCC000000CCCC777777777C000C00000CCCC0000CCCC77777C777CCC00
    CC00000CCCCCCCCCC77777CC77CCCCC0CCC000CCCCC00CCCCC777CCC7CCCCCCC
    CCCC0CCCCCCCCCCCCCC7CCCCCCCCCCCC0CCCCCCCCCCCCCCCCCCCCCC7CCC70CCC
    00CCCCCCCC0CC0CCCCCCCC77CC7700CC000CCCCCC000000CCCCCC777CC7700CC
    0000CCCC00000000CCCC7777CC7700CC0000C0CCC000000CCC7C7777CC7700CC
    0000C0CCC000000CCC7C7777CC7700CC0000CCCC00000000CCCC7777CC7700CC
    000CCCCCC000000CCCCCC777CC7700CC00CCCCCCCC0CC0CCCCCCCC77CC770CCC
    0CCCCCCCCCCCCCCCCCCCCCC7CCC7CCCCCCCC0CCCCCCCCCCCCCC7CCCCCCCCCCC0
    CCC000CCCCC00CCCCC777CCC7CCCCC00CC00000CCCCCCCCCC77777CC77CCC000
    C00000CCCC0000CCCC77777C777C000000000CCCC000000CCCC777777777000C
    0000CCCC00000000CCCC7777C77700CCC00CCCC0000000000CCCC77CCC770CCC
    CC0000CCCCCCCCCCCC7777CCCCC7CCCCC0000CCCCCCCCCCCCCC7777CCCCCCCCC
    0000CCCC00000000CCCC7777CCCCCCC0000CCCC0000000000CCCC7777CCC0000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    000000000000000000000000000000000000000000000000000000000000}
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 416
    Width = 633
    Height = 46
    Align = alBottom
    TabOrder = 0
    object BtnExport: TBitBtn
      Left = 232
      Top = 8
      Width = 241
      Height = 25
      Caption = #1069#1082#1089#1087#1086#1088#1090' '#1073#1072#1079#1099' '#1076#1072#1085#1085#1099#1093' '#1074' Excel'
      TabOrder = 0
      OnClick = BtnExportClick
    end
    object BtnDB: TBitBtn
      Left = 16
      Top = 8
      Width = 209
      Height = 25
      Caption = #1042#1099#1073#1088#1072#1090#1100' '#1073#1072#1079#1091' '#1076#1072#1085#1085#1099#1093' ...'
      TabOrder = 1
      OnClick = BtnDBClick
    end
  end
  object TabbedNotebook: TTabbedNotebook
    Left = 0
    Top = 0
    Width = 633
    Height = 416
    Align = alClient
    TabFont.Charset = DEFAULT_CHARSET
    TabFont.Color = clBtnText
    TabFont.Height = -11
    TabFont.Name = 'MS Sans Serif'
    TabFont.Style = []
    TabOrder = 1
    object TTabPage
      Left = 4
      Top = 27
      Caption = #1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093
      object DBGrid: TDBGrid
        Left = 0
        Top = 0
        Width = 625
        Height = 385
        Align = alClient
        DataSource = DataSource
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -14
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
      end
    end
  end
  object TableDB: TTable
    TableName = 'Name.DBF'
    Left = 144
    Top = 168
  end
  object DataSource: TDataSource
    DataSet = TableDB
    Left = 176
    Top = 168
  end
  object OpenDialog: TOpenDialog
    Filter = #1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093' (dBase)|*.dbf|'#1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093' (Paradox)|*.db'
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 176
    Top = 136
  end
end
