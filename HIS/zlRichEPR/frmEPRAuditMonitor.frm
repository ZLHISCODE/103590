VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRAuditMonitor 
   BorderStyle     =   0  'None
   Caption         =   "病历内容监测"
   ClientHeight    =   3840
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   6975
   Icon            =   "frmEPRAuditMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   0
      Left            =   465
      ScaleHeight     =   2535
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   660
      Width           =   5145
      Begin VSFlex8Ctl.VSFlexGrid vfgContent 
         Height          =   1890
         Left            =   480
         TabIndex        =   1
         Top             =   375
         Width           =   4185
         _cx             =   7382
         _cy             =   3334
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
   End
End
Attribute VB_Name = "frmEPRAuditMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    标志 = 0: 病历提纲: 提示等级: 提示内容
End Enum

'变量
'------------------------------------------------------------------------------------------------------------------
Private mlngRecordId As Long
Private mclsContent As clsVsf

Public Event GotFocus()

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
'    Set mfrmMain = frmMain
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsContent = New clsVsf
    With mclsContent
        
        Call .Initialize(Me.Controls, vfgContent, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("病历提纲", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("提示内容", 990, flexAlignLeftCenter, flexDTString, "", , True)
        vfgContent.FixedCols = 1
    End With
    
    zlInitData = True
    
End Function

Public Function zlClearData() As Boolean
    
    '
    zlClearData = mclsContent.ClearGrid
    
End Function

Public Sub zlPrintData(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '       strSubhead，打印的副标题
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = vfgContent
    objPrint.Title.Text = "病历内容监测记录"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
End Sub

Public Function zlRefreshData(ByVal lngRecordId As Long) As Boolean
'******************************************************************************************************************
'功能：上级程序调用本窗体的，传递参数，并显示窗体
'参数： lngRecordId-病人记录id
'******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

    mlngRecordId = lngRecordId
    
    Err = 0
    On Error GoTo errHand
    
    mclsContent.ClearGrid

    If lngRecordId = 0 Then Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Zl_病历内容监测_Neaten(" & lngRecordId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "病历内容监测")

    
    '提取内容监测数据
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select Lpad(' ', (提纲层次 - 1) * 3, ' ') || 提纲文本 As 病历提纲, 提示级别, 提示内容" & _
            " From 病历内容监测" & _
            " Where 病历记录ID = [1]" & _
            " Order By 提纲序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病历内容监测", lngRecordId)

    With Me.vfgContent
        .Clear
        Set .DataSource = rsTemp
        .TextMatrix(0, mCol.提示等级) = ""
        .TextMatrix(0, mCol.标志) = ""
        .ColWidth(mCol.提示等级) = 250
        
        .FixedCols = 1
        .FixedAlignment(mCol.标志) = flexAlignRightCenter
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftTop
        Next

        For lngCount = .FixedRows To .Rows - 1
            Select Case .TextMatrix(lngCount, mCol.提示等级)
            Case 0
                .TextMatrix(lngCount, mCol.提示等级) = ""
                .TextMatrix(lngCount, mCol.提示内容) = " "
            Case 1
                .TextMatrix(lngCount, mCol.提示等级) = "！"
                .Cell(flexcpForeColor, lngCount, mCol.提示等级, lngCount, mCol.提示内容) = RGB(0, 0, 255)
            Case 2
                .TextMatrix(lngCount, mCol.提示等级) = "！"
                .Cell(flexcpForeColor, lngCount, mCol.提示等级, lngCount, mCol.提示内容) = RGB(255, 0, 0)
            End Select
        Next
    End With
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

'######################################################################################################################

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsContent Is Nothing) Then Set mclsContent = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    vfgContent.Move 0, 0, picPane(Index).Width, picPane(Index).Height
End Sub

Private Sub vfgContent_GotFocus()
    RaiseEvent GotFocus
End Sub
