VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParameter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4350
   Icon            =   "frmParameter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4350
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPaiType 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2715
      _cx             =   4789
      _cy             =   5953
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   11
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmParameter.frx":6852
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3240
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":6955
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":6EEF
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":7489
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":DCEB
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":1454D
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":1ADAF
            Key             =   "add"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":21611
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameter.frx":22023
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "医生交接班病人类型顺序"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim strTemp As String
    
    With vsfPaiType
        For i = 1 To .Rows - 1
            strTemp = IIf(strTemp = "", "", strTemp & ";") & .TextMatrix(i, .ColIndex("类型号")) & "," & .TextMatrix(i, .ColIndex("简称"))
        Next
        Call zlDatabase.SetPara(1, strTemp, glngSys, 1242)
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strTemp As String
    Dim varTemp As Variant, varData As Variant
    Dim i As Long
    
    strTemp = zlDatabase.GetPara(1, glngSys, 1242)
    If strTemp = "" Then strTemp = "4,术后;2,抢救;6,死亡;7,输血;10,危/重;8,危;3,一级护理;5,术前;1,新入;12,留观"
    varTemp = Split(strTemp, ";")
    With vsfPaiType
        For i = 0 To UBound(varTemp)
            varData = Split(varTemp(i), ",")
            .TextMatrix(i + 1, .ColIndex("类型号")) = Val(varData(0))
            .TextMatrix(i + 1, .ColIndex("简称")) = varData(1)
            strTemp = Decode(Val(varData(0)), 10, "病情变化/ 病危/重患者", 8, "危急值患者", 5, "拟行手术患者", 1, "新入院患者", varData(1) & "患者")
            .TextMatrix(i + 1, .ColIndex("病人类型")) = strTemp
        Next
    End With
    Call vsfPaiType_AfterRowColChange(2, 1, 1, 1)
End Sub

Private Sub vsfPaiType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfPaiType
        If NewRow = 1 Then
            .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = ""
            .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = imgList.ListImages("Down").Picture
        Else
            If NewRow = .Rows - 1 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = imgList.ListImages("Up").Picture
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = imgList.ListImages("Up").Picture
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = imgList.ListImages("Down").Picture
            End If
        End If
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("上移")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("下移")) = ""
        End If
    End With
End Sub

Private Sub vsfPaiType_Click()
    Dim lngNum As Long, lngRow As Long
    Dim strPati As String, strName As String
    Dim blnAdjust As Boolean
    
    With vsfPaiType
        If .Row < 1 Then Exit Sub
        If .Col = .ColIndex("上移") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("上移")) Is Nothing Then
                lngRow = .Row - 1
                blnAdjust = True
            End If
        ElseIf .Col = .ColIndex("下移") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("下移")) Is Nothing Then
                lngRow = .Row + 1
                blnAdjust = True
            End If
        End If
        If Not blnAdjust Then Exit Sub
        strPati = .TextMatrix(.Row, .ColIndex("病人类型"))
        lngNum = Val(.TextMatrix(.Row, .ColIndex("类型号")))
        strName = .TextMatrix(.Row, .ColIndex("简称"))
        .TextMatrix(.Row, .ColIndex("简称")) = .TextMatrix(lngRow, .ColIndex("简称"))
        .TextMatrix(.Row, .ColIndex("病人类型")) = .TextMatrix(lngRow, .ColIndex("病人类型"))
        .TextMatrix(.Row, .ColIndex("类型号")) = .TextMatrix(lngRow, .ColIndex("类型号"))
        .TextMatrix(lngRow, .ColIndex("病人类型")) = strPati
        .TextMatrix(lngRow, .ColIndex("类型号")) = lngNum
        .TextMatrix(lngRow, .ColIndex("简称")) = strName
        .Row = lngRow
        .ShowCell lngRow, 1
    End With
End Sub
