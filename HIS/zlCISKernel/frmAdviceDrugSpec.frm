VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceDrugSpec 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品规格"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6435
   Icon            =   "frmAdviceDrugSpec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4560
      TabIndex        =   2
      Top             =   3135
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3345
      TabIndex        =   1
      Top             =   3135
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   2145
      Left            =   180
      TabIndex        =   0
      Top             =   855
      Width           =   6045
      _cx             =   10663
      _cy             =   3784
      Appearance      =   2
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceDrugSpec.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      Editable        =   2
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "切换为临时药品医嘱时，需要明确药品的规格；请为下列医嘱中按品种下达的药品医嘱选择一个药品规格。"
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   795
      TabIndex        =   3
      Top             =   255
      Width           =   5430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmAdviceDrugSpec.frx":0660
      Top             =   195
      Width           =   480
   End
End
Attribute VB_Name = "frmAdviceDrugSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSpec As String
Private mblnOK As Boolean

Private Enum COL_NUM
    COL_医嘱内容 = 0
    COL_规格 = 1
    COL_单量 = 2
    COL_医嘱ID = 3
    COL_药品ID = 4
    COL_药名ID = 5
    COL_药房ID = 6
End Enum

Public Function ShowMe(frmParent As Object, strSpec As String) As Boolean
'参数：strSpec="医嘱内容<Split2>单量<Split2>医嘱ID<Split2>药品ID<Split2>药名ID<Split2>药房ID<Split1>..."
'返回：strSpec="医嘱ID,药品ID;..."，只包含要选择规格的医嘱
    On Error Resume Next
    
    mstrSpec = strSpec
    Me.Show 1, frmParent
    ShowMe = mblnOK
    If mblnOK Then
        strSpec = mstrSpec
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSpec As String, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                strSpec = strSpec & ";" & Val(.TextMatrix(i, COL_医嘱ID)) & "," & Val(.TextMatrix(i, COL_药品ID))
                If Val(.TextMatrix(i, COL_药品ID)) = 0 Then
                    MsgBox "请为药品医嘱""" & .TextMatrix(i, COL_医嘱内容) & """确定一个药品规格。", vbInformation, gstrSysName
                    .Row = i: .Col = COL_规格
                    .ShowCell .Row, .Col: Exit Sub
                End If
            End If
        Next
        strSpec = Mid(strSpec, 2)
    End With
    
    mstrSpec = strSpec
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim arrAdvice As Variant, arrSub As Variant
    Dim rsDrug As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vBookMark As Variant
    Dim lngRow As Long
    
    mblnOK = False
    
    On Error GoTo errH
    
    arrAdvice = Split(mstrSpec, "<Split1>")
    With vsAdvice
        .Rows = UBound(arrAdvice) + 2
        For i = 1 To .Rows - 1
            arrSub = Split(arrAdvice(i - 1), "<Split2>")
            .TextMatrix(i, COL_医嘱内容) = arrSub(0)
            .TextMatrix(i, COL_规格) = ""
            .TextMatrix(i, COL_单量) = arrSub(1)
            .TextMatrix(i, COL_医嘱ID) = arrSub(2)
            .TextMatrix(i, COL_药品ID) = arrSub(3)
            .TextMatrix(i, COL_药名ID) = arrSub(4)
            .TextMatrix(i, COL_药房ID) = arrSub(5)
            
            '读取药品及规格信息
            Set rsDrug = New ADODB.Recordset
            Set rsDrug = GetDrugInfo(Val(.TextMatrix(i, COL_药名ID)), Val(.TextMatrix(i, COL_药品ID)), Val(.TextMatrix(i, COL_药房ID)), 2, False)
            If rsDrug.EOF Then
                MsgBox "药品医嘱""" & .TextMatrix(i, COL_医嘱内容) & """当前没有有效的规格。", vbInformation, gstrSysName
                Unload Me: Exit Sub
            ElseIf Val(.TextMatrix(i, COL_药品ID)) = 0 Then
                '标记为要选择规格
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80FFFF
                .RowData(i) = 1
                If lngRow = 0 Then lngRow = i
                
                '寻找合适的规格
                Call SeekMatchDrug(rsDrug, Val(.TextMatrix(i, COL_单量)), vBookMark, strSQL)
                If vBookMark <> 0 Then
                    rsDrug.Bookmark = vBookMark
                Else
                    rsDrug.MoveFirst
                End If
                .Cell(flexcpData, i, COL_规格) = strSQL '可选择的规格
                If strSQL = "" Then '全部规格都停用的药品
                    MsgBox "药品医嘱""" & .TextMatrix(i, COL_医嘱内容) & """当前没有有效的规格。", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
                
                .TextMatrix(i, COL_药品ID) = rsDrug!药品ID
            End If
            '显示指定或缺省规格信息
            .TextMatrix(i, COL_规格) = rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "")
        Next
        
        .Row = lngRow: .Col = COL_规格
        Call .AutoSize(COL_医嘱内容)
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SeekMatchDrug(rsDrug As ADODB.Recordset, ByVal sng单量 As Single, vBookMark As Variant, strList As String)
'功能：根据药品的多个规格定位缺省合适的规格,并设置相关药品信息到表格中
'参数：rsDrug=药品信息
'      vBookMark=返回用于定位规格位置的书签
'      strList=返回有效可供选择的规格,用于设置下拉框数据
    Dim vPreBookMark As Variant
    Dim lng倍数 As Long
        
    vPreBookMark = 0
    If Not rsDrug.EOF And Not rsDrug.BOF Then
        vPreBookMark = rsDrug.Bookmark
    End If
    
    rsDrug.MoveFirst
    vBookMark = 0: strList = ""
    Do While Not rsDrug.EOF
        '排开停用的药品
        If Nvl(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", Nvl(rsDrug!服务对象, 0)) > 0 Then
            If CInt(Nvl(sng单量, 0)) <> 0 And Nvl(rsDrug!库存, 0) > 0 Then
                '寻找剂量单位为单量的最小倍数的规格
                If rsDrug!剂量系数 / sng单量 = Int(rsDrug!剂量系数 / sng单量) Then
                    If rsDrug!剂量系数 / sng单量 < lng倍数 Or lng倍数 = 0 Then
                        vBookMark = rsDrug.Bookmark
                        lng倍数 = rsDrug!剂量系数 / sng单量
                    End If
                End If
            End If
            strList = strList & "|#" & rsDrug!药品ID & ";" & rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "") & _
                vbTab & IIF(InStr(GetInsidePrivs(p住院医嘱下达), "显示药品库存") = 0, _
                    IIF(Nvl(rsDrug!库存, 0) > 0, "有库存", "无库存"), "库存:" & Nvl(rsDrug!库存, 0) & rsDrug!住院单位)
        End If
        rsDrug.MoveNext
    Loop
    If vBookMark = 0 Then
        rsDrug.MoveFirst
        Do While Not rsDrug.EOF
            If Nvl(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", Nvl(rsDrug!服务对象, 0)) > 0 Then
                If Nvl(rsDrug!库存, 0) > 0 Then
                    vBookMark = rsDrug.Bookmark: Exit Do
                End If
            End If
            rsDrug.MoveNext
        Loop
    End If
    strList = Mid(strList, 2)
    
    If vBookMark = 0 And vPreBookMark <> 0 Then '没找到时恢复原有位置
        rsDrug.Bookmark = vPreBookMark
    End If
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = COL_规格 Then
        With vsAdvice
            If Val(.TextMatrix(Row, COL_药品ID)) = .ComboData Then Exit Sub
            .TextMatrix(Row, COL_药品ID) = .ComboData
        End With
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        '根据可否编辑设置编辑特性及光标特性
        If NewCol = COL_规格 And .RowData(NewRow) = 1 Then
            .ComboList = .Cell(flexcpData, NewRow, NewCol)
        Else
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(0, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.8 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.8
        End If
        If Col = COL_医嘱内容 Then Call vsAdvice.AutoSize(COL_医嘱内容)
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call EnterNextCell
    End If
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAdvice
        If Col = COL_规格 And .RowData(Row) = 1 Then
            If .ComboList = "" Then Cancel = True
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub EnterNextCell()
    With vsAdvice
        If .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1
            .Col = COL_规格
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub
