VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBloodCheckItemOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "血液核查项顺序设置"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4440
   Icon            =   "frmBloodCheckItemOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdUp 
      Caption         =   "上移(&U)"
      Height          =   350
      Left            =   3075
      TabIndex        =   5
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "下移(&D)"
      Height          =   350
      Left            =   3075
      TabIndex        =   4
      Top             =   1215
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1905
      TabIndex        =   3
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3075
      TabIndex        =   2
      Top             =   4560
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   120
      Left            =   0
      TabIndex        =   1
      Top             =   4290
      Width           =   5625
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   3450
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2775
      _cx             =   4895
      _cy             =   6085
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBloodCheckItemOrder.frx":08CA
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
   Begin VB.Label lbl 
      Caption         =   "请通过右侧按钮设置血液核查项目的顺序，点击确定保存"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3420
   End
End
Attribute VB_Name = "frmBloodCheckItemOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean

Private marrCheckItem
Private marrCheckItemKey
Private marrKey(13) As String

Public Function ShowColumn(ByVal frmMain As Object) As Boolean
    
    mblnOk = False

    If ExecuteCommand("初始控件") = False Then Exit Function
    
    Me.Show 1, frmMain
    
    If mblnOk Then
        ShowColumn = mblnOk
    End If
End Function


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSql As String
    Dim arrName
    
    On Error GoTo ErrHand
    
    Select Case strCommand
    Case "初始控件"
        '初始化核查项对应值
        arrName = Array("血袋无破损渗漏", "血液质量颜色", "姓名", "性别", "年龄", "病室", "床号", "住院号", "交叉配血试验结果", _
                        "血袋号", "血型", "血液成分", "血液剂量", "血液效期")
        strTmp = zlDatabase.GetPara(2, 2200, 9005)
        'ReDim arrName(13) As String
        With vsf
            .ScrollBars = flexScrollBarNone
            .ColHidden(0) = True
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, 2) = arrName(Split(strTmp, ",")(intRow - 1))
                .TextMatrix(intRow, 0) = Split(strTmp, ",")(intRow - 1)
                .TextMatrix(intRow, 1) = intRow
            Next
        
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "保存"
        strTmp = ""
        For intRow = 1 To vsf.Rows - 1
            strTmp = strTmp & vsf.TextMatrix(intRow, 0) & ","
        Next
        Call zlDatabase.SetPara(2, strTmp, 2200, 9005)
    End Select


    ExecuteCommand = True

    Exit Function
ErrHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Dim intNum As Integer
    Dim strName As String
    With vsf
        If .Row < 14 And .Row > 0 Then
            intNum = .TextMatrix(.Row, 0)
            .TextMatrix(.Row, 0) = .TextMatrix(.Row + 1, 0)
            .TextMatrix(.Row + 1, 0) = intNum
            strName = .TextMatrix(.Row, 2)
            .TextMatrix(.Row, 2) = .TextMatrix(.Row + 1, 2)
            .TextMatrix(.Row + 1, 2) = strName
            .Row = .Row + 1
            Call .Select(.Row, .Col)
        End If
    End With
End Sub

Private Sub cmdOK_Click()
    If ExecuteCommand("保存") = True Then
        Unload Me
    End If
End Sub

Private Sub cmdUp_Click()
    Dim intNum As Integer
    Dim strName As String
    
    With vsf
        If .Row < 15 And .Row > 1 Then
            intNum = .TextMatrix(.Row, 0)
            .TextMatrix(.Row, 0) = .TextMatrix(.Row - 1, 0)
            .TextMatrix(.Row - 1, 0) = intNum
            strName = .TextMatrix(.Row, 2)
            .TextMatrix(.Row, 2) = .TextMatrix(.Row - 1, 2)
            .TextMatrix(.Row - 1, 2) = strName
            .Row = .Row - 1
            Call .Select(.Row, .Col)
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set marrCheckItem = Nothing
    Set marrCheckItemKey = Nothing
End Sub
