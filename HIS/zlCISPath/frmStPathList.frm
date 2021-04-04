VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmStPathList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "标准路径选择"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   Icon            =   "frmStPathList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8580
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   8580
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3555
      Width           =   8580
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6960
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5520
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   240
         X2              =   10240
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsStPathList 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   8580
      _cx             =   15134
      _cy             =   6324
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   360
      RowHeightMax    =   360
      ColWidthMin     =   200
      ColWidthMax     =   5000
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmStPathList.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmStPathList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent     As Object '父窗体
Private mlStPathID     As Long '选择的标准路径ID
Private mblnOK         As Boolean
Private mrsStPath      As ADODB.Recordset '根据诊断获取的病人可用的标准路径列表

Private Enum Cols
    COL科室 = 0
    COL编码 = 1
    COL名称 = 2
    COL版本说明 = 3
End Enum

Public Function ShowMe(frmParent As Object, ByVal str疾病编码 As String, Optional ByVal intMode As Integer) As Boolean
'功能：根据疾病编码获取相关标准路径
'参数：frmParent  父窗体
'       str疾病编码 :格式：疾病编码1,疾病编码2,疾病编码3...
'       说明：本函数提供二种参看标准路径的模式
'       1、以疾病编码来参看相关标准路径 str疾病编码<>""
'       2、参看所有标准路径 str疾病编码=""
'       intMode 0-住院；1-门诊
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, str条件 As String, strSub条件 As String
    Dim i As Long
    Dim arrtmp As Variant, strSub编码 As String '对有分类的疾病编码进行截取
    Dim strTables As String
    
    On Error GoTo errH:
    Set mfrmParent = frmParent
    If intMode = 1 Then         '门诊
        strTables = "标准门诊路径目录 A, 标准门诊路径病种 B"
    Else                        '住院
        strTables = "标准路径目录 A, 标准路径病种 B"
    End If
    
    If Trim(str疾病编码) = "" Then
        Call frmStandardPathRef.ShowMe(mfrmParent, 0, , intMode)
        Exit Function
    Else
        arrtmp = Split(str疾病编码, ",")
        For i = LBound(arrtmp) To UBound(arrtmp)
            If i <> LBound(arrtmp) Then
                If InStr(Trim(arrtmp(i)), ".") > 0 Then
                    strSub条件 = strSub条件 & " Or InStr(b.疾病编码,'" & Left(Trim(arrtmp(i)), InStr(Trim(arrtmp(i)), ".") - 1) & "') > 0"
                Else
                    strSub条件 = strSub条件 & " Or InStr(b.疾病编码,'" & Trim(arrtmp(i)) & "') > 0"
                End If
                str条件 = str条件 & " Or InStr(b.疾病编码,'" & Trim(arrtmp(i)) & "') > 0"
            Else
                If InStr(Trim(arrtmp(i)), ".") > 0 Then
                    strSub条件 = " ( InStr(b.疾病编码,'" & Left(Trim(arrtmp(i)), InStr(Trim(arrtmp(i)), ".") - 1) & "') > 0"
                Else
                    strSub条件 = " ( InStr(b.疾病编码,'" & Trim(arrtmp(i)) & "') > 0"
                End If
                str条件 = " ( InStr(b.疾病编码,'" & Trim(arrtmp(i)) & "') > 0"
            End If
        Next
        str条件 = str条件 & " )"
        strSub条件 = strSub条件 & " )"
        strSql = " Select a.Id, a.科室名称, a.编码, a.路径名称, a.版本说明, b.疾病编码" & vbNewLine & _
                 " From " & strTables & vbNewLine & _
                 " Where  a.Id = b.标准路径id  And " & str条件
        Set mrsStPath = zlDatabase.OpenSQLRecord(strSql, gstrSysName)
    End If
    
    If mrsStPath.RecordCount = 0 Then '没有符合首要诊断的标准路径,就截取疾病编码大类，不要子分类进行匹配查找
        strSql = " Select a.Id, a.科室名称, a.编码, a.路径名称, a.版本说明, b.疾病编码" & vbNewLine & _
                 " From " & strTables & vbNewLine & _
                 " Where  a.Id = b.标准路径id  And " & strSub条件
        Set mrsStPath = zlDatabase.OpenSQLRecord(strSql, gstrSysName)
        If mrsStPath.RecordCount = 0 Then '没有符合首要诊断的标准路径
            Call frmStandardPathRef.ShowMe(mfrmParent, 0, , intMode)
            Exit Function
        ElseIf mrsStPath.RecordCount = 1 Then '仅有一条标准路径符合首要诊断
            Call frmStandardPathRef.ShowMe(mfrmParent, mrsStPath!ID, , intMode)
            Exit Function
        Else '仅有多条标准路径符合首要诊断
            Me.Show 1, frmParent
            Call frmStandardPathRef.ShowMe(mfrmParent, mlStPathID, , intMode)
            ShowMe = mblnOK
            Exit Function
        End If
    ElseIf mrsStPath.RecordCount = 1 Then '仅有一条标准路径符合首要诊断
        Call frmStandardPathRef.ShowMe(mfrmParent, mrsStPath!ID, , intMode)
        Exit Function
    Else '仅有多条标准路径符合首要诊断
        Me.Show 1, frmParent
        Call frmStandardPathRef.ShowMe(mfrmParent, mlStPathID, , intMode)
        ShowMe = mblnOK
        Exit Function
    End If
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
'功能：取消时推出窗体
    mblnOK = False
    mlStPathID = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
'功能：确定时获取选中标准路径
    If mlStPathID = 0 Then
        MsgBox "你还未选择标准路径，请选择标准路径", vbOKOnly, gstrSysName
    Else
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
'功能：标准路径数据加载到控件中
    Dim i As Long
    With vsStPathList
        .Rows = .FixedRows
        mrsStPath.MoveFirst
        For i = 1 To mrsStPath.RecordCount
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL科室) = mrsStPath!科室名称
            .TextMatrix(.Rows - 1, COL编码) = mrsStPath!编码 & ""
            .TextMatrix(.Rows - 1, COL名称) = mrsStPath!路径名称 & ""
            .TextMatrix(.Rows - 1, COL版本说明) = mrsStPath!版本说明 & ""
            .RowData(.Rows - 1) = mrsStPath!ID & ""
            mrsStPath.MoveNext
        Next
    End With
End Sub

Private Sub Form_Resize()
'功能：调整窗体中控件位置
    vsStPathList.Width = Me.ScaleWidth - vsStPathList.Left
    picBottom.Top = Me.ScaleHeight - picBottom.Height
    vsStPathList.Height = picBottom.Top - vsStPathList.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
'功能：确定退出时，并且标准路径已经选择时不阻止退出，否则阻止退出
    If mblnOK And mlStPathID = 0 Then
        Cancel = True
    Else
        Set mrsStPath = Nothing
        Set mfrmParent = Nothing
    End If
End Sub

Private Sub vsStPathList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能：选择标准路径
    mlStPathID = Val(vsStPathList.RowData(NewRow))
End Sub

Private Sub vsStPathList_DblClick()
'功能：双击某条标准路径时，确认选择
    mlStPathID = Val(vsStPathList.RowData(vsStPathList.Row))
    Call cmdOK_Click
End Sub


