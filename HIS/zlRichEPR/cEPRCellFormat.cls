VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRCellFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarFontBold As Boolean
Private mvarFontItalic As Boolean
Private mvarFontStrikeout As Boolean
Private mvarFontUnderline As Boolean
Private mvarMergeCells As Boolean
Private mvarProtectionHidden As Boolean
Private mvarProtectionLocked As Boolean
Private mvarWordWrap As Boolean
Private mvarPatternStyle As Integer
Private mvarFontSize As Integer
Private mvarAlignHorizontal As Long
Private mvarAlignVertical As Long
Private mvarFontCharSet As Long
Private mvarFontColor As OLE_COLOR
Private mvarPatternBG As OLE_COLOR
Private mvarPatternFG As OLE_COLOR
Private mvarBorderColor(0 To 5) As OLE_COLOR
Private mvarBorderStyle(0 To 5) As Integer
Private mvarValidationRule As String
Private mvarValidationRuleLocal As String
Private mvarValidationText As String
Private mvarNumberFormat As String
Private mvarNumberFormatLocal As String
Private mvarFontName As String

Public Property Let FontBold(ByVal vData As Boolean)
    mvarFontBold = vData
End Property

Public Property Get FontBold() As Boolean
    FontBold = mvarFontBold
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
    mvarFontItalic = vData
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mvarFontItalic
End Property

Public Property Let FontStrikeout(ByVal vData As Boolean)
    mvarFontStrikeout = vData
End Property

Public Property Get FontStrikeout() As Boolean
    FontStrikeout = mvarFontStrikeout
End Property

Public Property Let FontUnderline(ByVal vData As Boolean)
    mvarFontUnderline = vData
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = mvarFontUnderline
End Property

Public Property Let MergeCells(ByVal vData As Boolean)
    mvarMergeCells = vData
End Property

Public Property Get MergeCells() As Boolean
    MergeCells = mvarMergeCells
End Property

Public Property Let ProtectionHidden(ByVal vData As Boolean)
    mvarProtectionHidden = vData
End Property

Public Property Get ProtectionHidden() As Boolean
    ProtectionHidden = mvarProtectionHidden
End Property

Public Property Let ProtectionLocked(ByVal vData As Boolean)
    mvarProtectionLocked = vData
End Property

Public Property Get ProtectionLocked() As Boolean
    ProtectionLocked = mvarProtectionLocked
End Property

Public Property Let WordWrap(ByVal vData As Boolean)
    mvarWordWrap = vData
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = mvarWordWrap
End Property

Public Property Let PatternStyle(ByVal vData As Integer)
    mvarPatternStyle = vData
End Property

Public Property Get PatternStyle() As Integer
    PatternStyle = mvarPatternStyle
End Property

Public Property Let FontSize(ByVal vData As Integer)
    mvarFontSize = vData
End Property

Public Property Get FontSize() As Integer
    FontSize = mvarFontSize
End Property

Public Property Let AlignHorizontal(ByVal vData As Long)
    mvarAlignHorizontal = vData
End Property

Public Property Get AlignHorizontal() As Long
    AlignHorizontal = mvarAlignHorizontal
End Property

Public Property Let AlignVertical(ByVal vData As Long)
    mvarAlignVertical = vData
End Property

Public Property Get AlignVertical() As Long
    AlignVertical = mvarAlignVertical
End Property

Public Property Let FontCharSet(ByVal vData As Long)
    mvarFontCharSet = vData
End Property

Public Property Get FontCharSet() As Long
    FontCharSet = mvarFontCharSet
End Property

Public Property Let FontColor(ByVal vData As OLE_COLOR)
    mvarFontColor = vData
End Property

Public Property Get FontColor() As OLE_COLOR
    FontColor = mvarFontColor
End Property

Public Property Let PatternBG(ByVal vData As OLE_COLOR)
    mvarPatternBG = vData
End Property

Public Property Get PatternBG() As OLE_COLOR
    PatternBG = mvarPatternBG
End Property

Public Property Let PatternFG(ByVal vData As OLE_COLOR)
    mvarPatternFG = vData
End Property

Public Property Get PatternFG() As OLE_COLOR
    PatternFG = mvarPatternFG
End Property

Public Property Let BorderColor(lngBorder As F1BorderConstants, vData As OLE_COLOR)
    Select Case lngBorder
    Case F1HInsideBorder
        mvarBorderColor(0) = vData
    Case F1VInsideBorder
        mvarBorderColor(1) = vData
    Case F1TopBorder
        mvarBorderColor(2) = vData
    Case F1LeftBorder
        mvarBorderColor(3) = vData
    Case F1BottomBorder
        mvarBorderColor(4) = vData
    Case F1RightBorder
        mvarBorderColor(5) = vData
    End Select
End Property

Public Property Get BorderColor(lngBorder As F1BorderConstants) As OLE_COLOR
    Select Case lngBorder
    Case F1HInsideBorder
        BorderColor = mvarBorderColor(0)
    Case F1VInsideBorder
        BorderColor = mvarBorderColor(1)
    Case F1TopBorder
        BorderColor = mvarBorderColor(2)
    Case F1LeftBorder
        BorderColor = mvarBorderColor(3)
    Case F1BottomBorder
        BorderColor = mvarBorderColor(4)
    Case F1RightBorder
        BorderColor = mvarBorderColor(5)
    End Select
End Property

Public Property Let BorderStyle(lngBorder As F1BorderConstants, vData As Integer)
    Select Case lngBorder
    Case F1HInsideBorder
        mvarBorderStyle(0) = vData
    Case F1VInsideBorder
        mvarBorderStyle(1) = vData
    Case F1TopBorder
        mvarBorderStyle(2) = vData
    Case F1LeftBorder
        mvarBorderStyle(3) = vData
    Case F1BottomBorder
        mvarBorderStyle(4) = vData
    Case F1RightBorder
        mvarBorderStyle(5) = vData
    End Select
End Property

Public Property Get BorderStyle(lngBorder As F1BorderConstants) As Integer
    Select Case lngBorder
    Case F1HInsideBorder
        BorderStyle = mvarBorderStyle(0)
    Case F1VInsideBorder
        BorderStyle = mvarBorderStyle(1)
    Case F1TopBorder
        BorderStyle = mvarBorderStyle(2)
    Case F1LeftBorder
        BorderStyle = mvarBorderStyle(3)
    Case F1BottomBorder
        BorderStyle = mvarBorderStyle(4)
    Case F1RightBorder
        BorderStyle = mvarBorderStyle(5)
    End Select
End Property

Public Property Let ValidationRule(ByVal vData As String)
    mvarValidationRule = vData
End Property

Public Property Get ValidationRule() As String
    ValidationRule = mvarValidationRule
End Property

Public Property Let ValidationRuleLocal(ByVal vData As String)
    mvarValidationRuleLocal = vData
End Property

Public Property Get ValidationRuleLocal() As String
    ValidationRuleLocal = mvarValidationRuleLocal
End Property

Public Property Let ValidationText(ByVal vData As String)
    mvarValidationText = vData
End Property

Public Property Get ValidationText() As String
    ValidationText = mvarValidationText
End Property

Public Property Let NumberFormat(ByVal vData As String)
    mvarNumberFormat = vData
End Property

Public Property Get NumberFormat() As String
    NumberFormat = mvarNumberFormat
End Property

Public Property Let NumberFormatLocal(ByVal vData As String)
    mvarNumberFormatLocal = vData
End Property

Public Property Get NumberFormatLocal() As String
    NumberFormatLocal = mvarNumberFormatLocal
End Property

Public Property Let FontName(ByVal vData As String)
    mvarFontName = Trim(vData)
End Property

Public Property Get FontName() As String
    FontName = Trim(mvarFontName)
End Property
