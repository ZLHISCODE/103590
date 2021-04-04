VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProcRelating 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关联过程"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   Icon            =   "frmProcRelating.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraTip 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   7830
      Begin VB.Label lblTipPath 
         AutoSize        =   -1  'True
         Caption         =   "调用查看路径："
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label lblTipPath 
         Caption         =   "C:\AppSoft\Log\过程管理\RelProc.ini"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   600
         Width           =   6450
      End
      Begin VB.Label lblTip 
         Caption         =   "您可以直接删除当前存储过程，正在使用它的其他存储过程将会以文本文件的方式存储在本地，请及时对其他存储过程进行处理。"
         Height          =   390
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.Frame fraSplit 
      BackColor       =   &H80000012&
      Height          =   30
      Index           =   1
      Left            =   -240
      TabIndex        =   8
      Top             =   4440
      Width           =   8625
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8310
      TabIndex        =   6
      Top             =   0
      Width           =   8310
      Begin VB.Image ImgTop 
         Height          =   720
         Left            =   240
         Picture         =   "frmProcRelating.frx":6852
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblTop 
         Caption         =   $"frmProcRelating.frx":771C
         Height          =   600
         Left            =   1095
         TabIndex        =   7
         Top             =   135
         Width           =   7050
      End
   End
   Begin VB.Frame fraSplit 
      BackColor       =   &H80000012&
      Height          =   30
      Index           =   0
      Left            =   -360
      TabIndex        =   5
      Top             =   960
      Width           =   8625
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   700
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   8310
      TabIndex        =   2
      Top             =   4470
      Width           =   8310
      Begin VB.CommandButton cmdIgnor 
         Caption         =   "忽略(&O)"
         Height          =   350
         Left            =   6960
         TabIndex        =   4
         Top             =   175
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "强制删除(&O)"
         Height          =   350
         Left            =   5280
         TabIndex        =   3
         Top             =   175
         Width           =   1335
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleMode       =   0  'User
      ScaleWidth      =   8290.589
      TabIndex        =   0
      Top             =   1080
      Width           =   7830
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   2145
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7815
         _cx             =   13785
         _cy             =   3784
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmProcRelating.frx":77B8
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
   End
End
Attribute VB_Name = "frmProcRelating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==模块变量
'==============================================================
Private mobjMain As Object
Private mstrAllIDs As String
Private mstrValiIDs As String
Private mrsCheckData As ADODB.Recordset
Private mblnOk As Boolean
Private Enum RelationCol
    RC_当前过程 = 0
    RC_相关过程 = 1
    RC_相关过程本次删除 = 2
End Enum
'==============================================================
'==公共接口
'==============================================================
Public Function CheckRelation(ByVal objMain As Object, strIDs As String) As Boolean
    Set mobjMain = objMain
    Set mrsCheckData = GetCheckRelation(strIDs)
    If mrsCheckData Is Nothing Then
        CheckRelation = False
        Exit Function
    Else
        mrsCheckData.Filter = "状态=1" '查看是否存在引用对象本次不删除
        If mrsCheckData.RecordCount = 0 Then
            CheckRelation = True
            strIDs = mstrAllIDs
            Exit Function
        Else
            mrsCheckData.Filter = "状态<>0"
            mrsCheckData.Sort = "ID,Referenced_Name"
        End If
    End If
    Me.Show 1, mobjMain
    CheckRelation = True
    If mblnOk Then
        strIDs = mstrAllIDs
    Else
        strIDs = mstrValiIDs
    End If
End Function

'==============================================================
'==控件事件
'==============================================================
Private Sub cmdDel_Click()
    Dim objFSO As TextStream
    Dim i As Long
    '生成脚本
    If gobjFile.FileExists(lblTipPath(1).Caption) = False Then
        Set objFSO = gobjFile.CreateTextFile(lblTipPath(1).Caption, True)
    Else
        Set objFSO = gobjFile.OpenTextFile(lblTipPath(1).Caption, ForWriting)
    End If
    With vsfMain
        objFSO.WriteLine RPAD("当前过程", 30) & RPAD("相关过程", 30)
        For i = .FixedRows To .Rows - 1
             objFSO.WriteLine RPAD(.TextMatrix(i, RC_当前过程), 30) & RPAD(.TextMatrix(i, RC_相关过程), 30)
        Next
        objFSO.Close
    End With
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdIgnor_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strFolder As String
    lblTipPath(1).Caption = GetLogPath(LT_自定义, , , "自定义过程", "ProcedureRelation")
    Call LoadData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsfMain.Move 15, 15, picMain.ScaleWidth - 30, picMain.ScaleHeight - 30
    picTop.Height = ImgTop.Top + ImgTop.Height + 30
    picTop.Width = Me.ScaleWidth
    fraSplit(0).Top = picTop.Height + 15
    picMain.Top = fraSplit(0).Top + fraSplit(0).Height + 15
    fraSplit(1).Top = picBottom.Top - fraSplit(1).Height - 15
    fraTip.Top = fraSplit(1).Top - 15 - fraTip.Height
    picMain.Height = fraTip.Top - picMain.Top
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    vsfMain.Move 0, 0, picMain.ScaleWidth, picMain.ScaleHeight
End Sub

'==============================================================
'==私有方法
'==============================================================
Private Function GetCheckRelation(ByVal strIDs As String) As ADODB.Recordset
'功能：获取相关联数据
'参数：strIDs=ID条件，*类型 表示该类型的所有对象。-类型,ID1...:表示该类型去掉特定的ID,ID1,...：表示只获取这些ID
'           依赖关系记录集
    Dim strSQL As String, rsRelation As ADODB.Recordset
    Dim strProcs As String, intType As Integer, strTmp As String
    Dim lngPos As String, i As Integer
    Dim strPreID As String, blnValid As Boolean
    On Error GoTo errH
    mstrAllIDs = "": mstrValiIDs = ""
    '获取本次检查的存储过程
    If strIDs Like "[*]*" Then
        strProcs = "Select Id, Upper(名称) 名称, Upper(所有者) 所有者 From Zlprocedure Where 类型 = [1]"
        intType = Val(Mid(strIDs, 2))
    ElseIf strIDs Like "-*" Then
        lngPos = InStr(strIDs, ",")
        strTmp = Mid(strIDs, 1, lngPos - 1)
        strIDs = Mid(strIDs, lngPos + 1)
        intType = Val(Mid(strTmp, 2))
        strProcs = "Select Id, Upper(名称) 名称, Upper(所有者) 所有者" & vbNewLine & _
                    "From Zlprocedure a, Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) b" & vbNewLine & _
                    "Where 类型 = [1] And a.Id = b.Column_Value(+) And b.Column_Value Is Null"
    Else
        strProcs = "Select Id, Upper(名称) 名称, Upper(所有者) 所有者" & vbNewLine & _
                        "From Zlprocedure" & vbNewLine & _
                        "Where Id In (Select Column_Value From Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)))"
    End If
    strSQL = "Select b.Id, b.名称, Name, Referenced_Name, Decode(Name, Null, 0, Decode(c.名称, Null, 1, 2)) 状态" & vbNewLine & _
                    "From (Select Name, Referenced_Name" & vbNewLine & _
                    "       From User_Dependencies" & vbNewLine & _
                    "       Where Type In ('PROCEDURE', 'FUNCTION') And Name Not Like 'BIN$%' And Referenced_Type In ('PROCEDURE', 'FUNCTION') And" & vbNewLine & _
                    "             Referenced_Name Not Like 'BIN$%') a, (" & strProcs & ") b," & vbNewLine & _
                    "     (" & strProcs & ") c" & vbNewLine & _
                    "Where a.Referenced_Name(+) = b.名称 And c.名称(+) = a.Name " & vbNewLine & _
                    "order by b.Id,a.Name,c.名称"
    Set rsRelation = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "检查自定义过程关联对象", intType, strIDs)
    strPreID = "": blnValid = True
    For i = 1 To rsRelation.RecordCount
        If rsRelation!Id <> strPreID Then
            If strPreID <> "" Then
                If blnValid Then
                    mstrValiIDs = mstrValiIDs & "," & strPreID
                End If
                mstrAllIDs = mstrAllIDs & "," & strPreID
            End If
            strPreID = rsRelation!Id
            blnValid = rsRelation!状态 <> 1
        Else
            blnValid = blnValid And rsRelation!状态 <> 1
        End If
        rsRelation.MoveNext
    Next
    If strPreID <> "" Then
        If blnValid Then
            mstrValiIDs = mstrValiIDs & "," & strPreID
        End If
        mstrAllIDs = mstrAllIDs & "," & strPreID
    End If
    If mstrAllIDs <> "" Then mstrAllIDs = Mid(mstrAllIDs, 2)
    If mstrValiIDs <> "" Then mstrValiIDs = Mid(mstrValiIDs, 2)
    Set GetCheckRelation = rsRelation
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "对象依赖检查失败：" & err.Description, vbInformation, "关联过程"
End Function

Private Sub LoadData()
'功能：数据加载
    Dim lngRow As Long, strPreID As String, lngPreRow As Long, blnValid As Boolean
    With vsfMain
        .Redraw = flexRDNone
        .Rows = .FixedRows
        strPreID = "": blnValid = True
        Do While Not mrsCheckData.EOF
            .Rows = .Rows + 1: lngRow = .Rows - 1
            If mrsCheckData!Id <> strPreID Then
                If strPreID <> "" Then
                    .Cell(flexcpData, lngPreRow, RC_当前过程, lngRow - 1, RC_当前过程) = IIf(blnValid, 1, 0)
                End If
                strPreID = mrsCheckData!Id
                blnValid = mrsCheckData!状态 <> 1
                lngPreRow = lngRow
            Else
                blnValid = blnValid And mrsCheckData!状态 <> 1
            End If
            .TextMatrix(lngRow, RC_当前过程) = mrsCheckData!名称 & ""
            .TextMatrix(lngRow, RC_相关过程) = mrsCheckData!name & ""
            .TextMatrix(lngRow, RC_相关过程本次删除) = IIf(mrsCheckData!状态 = 1, "", "√")
            mrsCheckData.MoveNext
        Loop
        If strPreID <> "" Then
            .Cell(flexcpData, lngPreRow, RC_当前过程, .Rows - 1, RC_当前过程) = IIf(blnValid, 1, 0)
        End If
        .Redraw = flexRDDirect
    End With
End Sub

