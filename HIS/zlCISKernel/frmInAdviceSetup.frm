VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInAdviceSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "住院医嘱选项"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "frmInAdviceSetUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5985
      Left            =   90
      ScaleHeight     =   5985
      ScaleWidth      =   4215
      TabIndex        =   3
      Top             =   -240
      Width           =   4215
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   4965
         Left            =   0
         TabIndex        =   4
         Top             =   975
         Width           =   4095
         _cx             =   7223
         _cy             =   8758
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
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInAdviceSetUp.frx":000C
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
      Begin VB.Label lblKYYF 
         Caption         =   $"frmInAdviceSetUp.frx":0095
         Height          =   615
         Left            =   15
         TabIndex        =   5
         Top             =   375
         Width           =   4215
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   4380
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   4380
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   1890
         TabIndex        =   0
         Top             =   60
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3060
         TabIndex        =   1
         Top             =   60
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmInAdviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const VsPubBackColor = &HFAEADA

Private Sub cmdOK_Click()
    Dim i As Long
    Dim bytType As Long
    Dim arr可用药房(4) As String, arr缺省药房(4) As String, arrTmp() As String
    Dim blnSetup As Boolean
    
    blnSetup = InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱选项设置;") > 0
    '药房
    With vsfDrugStore
        For i = .FixedRows To .Rows - 1
            Select Case .TextMatrix(i, .ColIndex("类别"))
            Case "西药房"
                bytType = 0
            Case "成药房"
                bytType = 1
            Case "中药房"
                bytType = 2
            Case "发料部门"
                bytType = 3
            End Select
            If .TextMatrix(i, .ColIndex("可用")) <> 0 Then arr可用药房(bytType) = arr可用药房(bytType) & "," & .RowData(i)
            If .TextMatrix(i, .ColIndex("缺省")) = "√" Then arr缺省药房(bytType) = .RowData(i)
        Next
    End With
    arrTmp = Split("西药房,成药房,中药房,发料部门", ",")
    For bytType = 0 To UBound(arrTmp)
        Call zlDatabase.SetPara("住院可用" & arrTmp(bytType), Mid(arr可用药房(bytType), 2), glngSys, p住院医嘱下达, blnSetup)
        Call zlDatabase.SetPara("住院缺省" & arrTmp(bytType), arr缺省药房(bytType), glngSys, p住院医嘱下达, blnSetup)
    Next
     
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '表格的checkbox按回车，不转移焦点
        If Not Me.ActiveControl Is vsfDrugStore Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPar As String, i As Long
    Dim bln下达设置 As Boolean
    Dim arrTmp() As String
    Dim strDSIDs As String, strDefault As String, lngBackColor As Long, bytLockEdit As Byte
    Dim intType1 As Integer, intType2 As Integer, lngRow As Long
 
    
    On Error GoTo errH
             
    '药房与发料部门
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " AND B.部门ID=A.ID And B.服务对象 IN(2,3) and B.工作性质 in('中药房','西药房','成药房','发料部门')" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by 工作性质,编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With vsfDrugStore
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        .MergeCol(.ColIndex("类别")) = True
        .MergeCells = flexMergeFixedOnly
        
        
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = .FixedRows
            arrTmp = Split("西药房,成药房,中药房,发料部门", ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "工作性质='" & arrTmp(i) & "'"
                strDefault = zlDatabase.GetPara("住院缺省" & arrTmp(i), glngSys, p住院医嘱下达, , , , intType1)
                strDSIDs = "," & zlDatabase.GetPara("住院可用" & arrTmp(i), glngSys, p住院医嘱下达, , , , intType2) & ","
                Do While Not rsTmp.EOF
                    .TextMatrix(lngRow, .ColIndex("类别")) = arrTmp(i)
                    .TextMatrix(lngRow, .ColIndex("药房")) = rsTmp!名称
                    .RowData(lngRow) = Val(rsTmp!ID)
                    
                    If Val(rsTmp!ID) = Val(strDefault) Then
                        .TextMatrix(lngRow, .ColIndex("缺省")) = "√"
                        .TextMatrix(lngRow, .ColIndex("可用")) = -1   'true
                    Else
                        .TextMatrix(lngRow, .ColIndex("缺省")) = ""
                        .TextMatrix(lngRow, .ColIndex("可用")) = IIF(InStr(strDSIDs, "," & rsTmp!ID & ",") > 0, -1, 0)
                    End If
                    
                    '缺省单元格
                    'intType-'返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType1 & ",") > 0 Then
                        lngBackColor = IIF(bln下达设置, VsPubBackColor, &H8000000F)      '授权限控制
                        bytLockEdit = IIF(bln下达设置, 0, 1)
                    ElseIf intType1 = 5 Then
                        lngBackColor = VsPubBackColor       '公共模块,但不授权限控制
                    Else
                        lngBackColor = &H80000005     '正常编辑
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("缺省")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("缺省")) = bytLockEdit
                     
                    '可用单元格
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType2 & ",") > 0 Then
                        lngBackColor = IIF(bln下达设置, VsPubBackColor, &H8000000F)      '授权限控制
                        bytLockEdit = IIF(bln下达设置, 0, 1)
                    ElseIf intType2 = 5 Then
                        lngBackColor = VsPubBackColor       '公共模块,但不授权限控制
                    Else
                        lngBackColor = &H80000005     '正常编辑
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("可用")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("可用")) = bytLockEdit
                    
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Loop
                If lngRow < .Rows - 1 Then  '划分隔线
                    .Select lngRow, .FixedCols, lngRow, .Cols - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
            Next
        End If
    End With
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    cmdCancel.Left = Me.ScaleLeft + Me.ScaleWidth - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
End Sub

Private Sub vsfDrugStore_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDrugStore.ColIndex("可用") Then
        Call Set可用药房(Row, True)
    ElseIf Col = vsfDrugStore.ColIndex("可用") Then
        Call Set缺省药房
    End If
    Cancel = True
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("可用")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("缺省")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    With vsfDrugStore
        If .MouseCol = .ColIndex("缺省") Then
            Call Set缺省药房
        ElseIf .MouseCol = .ColIndex("药房") Then
            Call Set可用药房(.Row, True)
        ElseIf .MouseCol = .ColIndex("可用") And .MouseRow = .FixedRows - 1 Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                Call Set可用药房(i)
            Next
        End If
    End With
End Sub
Private Sub vsfDrugStore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If vsfDrugStore.Col = vsfDrugStore.ColIndex("缺省") Then
            Call Set缺省药房
        End If
    End If
End Sub

Private Sub Set缺省药房()
'功能：设置当前行的缺省药房，同时处理相同类型的其他行的缺省药房
    Dim i As Long
    
    With vsfDrugStore
        If Val("" & .Cell(flexcpData, .Row, .ColIndex("缺省"))) = 0 Then  '该参数允许修改的情况下
            If .TextMatrix(.Row, .ColIndex("缺省")) = "√" Then
                .TextMatrix(.Row, .ColIndex("缺省")) = ""
            Else
                '当没有有权限修改可用时且可用为0（false)时不允许设置缺省
                If Not (Val(.TextMatrix(.Row, .ColIndex("可用"))) = 0 And Val("" & .Cell(flexcpData, .Row, .ColIndex("可用"))) = 1) Then
                    '同类别的其他行取消缺省
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(.Row, .ColIndex("类别")) = .TextMatrix(i, .ColIndex("类别")) Then
                            If .TextMatrix(i, .ColIndex("缺省")) = "√" Then .TextMatrix(i, .ColIndex("缺省")) = ""
                        End If
                    Next
                    .TextMatrix(.Row, .ColIndex("可用")) = -1    '自动设置为可用
                    .TextMatrix(.Row, .ColIndex("缺省")) = "√"
                Else
                    MsgBox "设置当前药房为缺省时，会同时将当前药房设置为可用，" & vbNewLine & "你没有修改可用药房的权限。", vbInformation, gstrSysName
                End If
            End If
        Else
            MsgBox "你没有修改缺省药房的权限。", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Set可用药房(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = False)
'功能：设置当前行的可用药房，同时处理当前行的缺省药房
    With vsfDrugStore
        If Val("" & .Cell(flexcpData, lngRow, .ColIndex("可用"))) = 0 Then   '该参数允许修改的情况下
            If Val(.TextMatrix(lngRow, .ColIndex("可用"))) = -1 Then
                '当前科室勾选可用
                If Not (Val("" & .Cell(flexcpData, lngRow, .ColIndex("缺省"))) = 1 And .TextMatrix(lngRow, .ColIndex("缺省")) = "√") Then
                    .TextMatrix(lngRow, .ColIndex("可用")) = 0
                    .TextMatrix(lngRow, .ColIndex("缺省")) = ""
                Else
                    If blnAsk Then
                        MsgBox "取消当前药房可用时，会同时取消当前药房缺省，" & vbNewLine & "你没有修改缺省药房的权限。", vbInformation, gstrSysName
                    End If
                End If
            Else
                .TextMatrix(lngRow, .ColIndex("可用")) = -1    '自动设置为可用
            End If
        Else
            If blnAsk Then
                MsgBox "你没有修改可用药房的权限。", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
