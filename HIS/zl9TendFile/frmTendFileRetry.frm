VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileRetry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "重算数据行"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8310
   Icon            =   "frmTendFileRetry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8310
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraMain 
      Height          =   5190
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   8295
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         Height          =   330
         Left            =   7185
         TabIndex        =   3
         Top             =   4770
         Width           =   900
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定"
         Height          =   330
         Left            =   6165
         TabIndex        =   2
         Top             =   4770
         Width           =   900
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfFile 
         Height          =   4575
         Left            =   15
         TabIndex        =   1
         Top             =   105
         Width           =   8160
         _cx             =   14393
         _cy             =   8070
         Appearance      =   2
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
         ForeColorSel    =   -2147483640
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
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         Begin VB.CheckBox chkChoose 
            Height          =   165
            Left            =   660
            TabIndex        =   4
            Top             =   60
            Width           =   180
         End
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTendFileRetry.frx":06EA
               Key             =   "体温单"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTendFileRetry.frx":0DFC
               Key             =   "记录单"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTendFileRetry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum mCol
    f标志 = 0: fID: 选择: f文件名称: f开始时间: f结束时间: 续打文件: 续打文件id
End Enum
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mintBaby As Integer
Private mblnSave As Boolean           '是否重算

Public Function ShowMe(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intBabby As Integer, ByVal strPrivs As String) As Boolean
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mintBaby = intBabby
    mblnSave = False
    Call zlRefData
    Me.Show 1
    ShowMe = mblnSave
End Function

Private Function zlRefData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim intRow As Integer
    Dim lngID As Long
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    Err = 0
    On Error GoTo ErrHand
    '------------------------------------------------------------------------------------------------------------------
    '护理文件刷新
    
    With vsfFile
        .Clear
        .Rows = 1
        .Cols = 8
        .FixedCols = 1
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .TextMatrix(0, mCol.f标志) = ""
        .TextMatrix(0, mCol.fID) = "ID"
        .TextMatrix(0, mCol.f文件名称) = "文件名称"
        .TextMatrix(0, mCol.f开始时间) = "开始时间"
        .TextMatrix(0, mCol.f结束时间) = "结束时间"
        .TextMatrix(0, mCol.续打文件) = "续打记录单"
        .TextMatrix(0, mCol.续打文件id) = "续打文件id"
        
        
        .ColDataType(mCol.选择) = flexDTBoolean
        .ColWidth(mCol.f标志) = 270: .ColWidth(mCol.选择) = 500
        .ColWidth(mCol.fID) = 0: .ColWidth(mCol.f文件名称) = 2200: .ColWidth(mCol.f开始时间) = 1800
        .ColWidth(mCol.f结束时间) = 1800: .ColWidth(mCol.续打文件) = 2200: .ColWidth(mCol.续打文件id) = 0
    End With
    
    intRow = vsfFile.FixedRows
    '--------------------------------------------------------------------------------------------------------------
    strSQL = " Select A.ID,A.文件名称, B.名称 AS 文件来源,B.保留,A.开始时间,A.结束时间,A.创建人,A.创建时间,A.归档人,C.文件名称 AS 续打文件,C.ID AS 续打文件ID,B.保留 " & _
              " From 病人护理文件 A,病历文件列表 B,病人护理文件 C" & _
              " Where A.格式ID=B.ID And A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3] And A.续打ID=C.ID(+) and B.保留=0" & _
              " Order by B.保留,A.开始时间"
    Call SQLDIY(strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "显示指定病人的护理文件列表", mlng病人ID, mlng主页ID, mintBaby)
    rsTemp.Filter = 0
    rsTemp.Sort = "开始时间"
    With Me.vsfFile
        Do While Not rsTemp.EOF
            vsfFile.Rows = vsfFile.Rows + 1
            i = vsfFile.Rows
            If Val(.TextMatrix(i - 1, mCol.fID)) > 0 Then .AddItem ""
            
            lngID = Val(NVL(rsTemp!ID, 0))
            .TextMatrix(i - 1, mCol.fID) = lngID
            Set .Cell(flexcpPicture, i - 1, mCol.f标志) = imgList.ListImages("记录单").Picture
            .TextMatrix(i - 1, mCol.f文件名称) = NVL(rsTemp!文件名称)
            .TextMatrix(i - 1, mCol.f开始时间) = Format(NVL(rsTemp!开始时间), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(i - 1, mCol.f结束时间) = Format(NVL(rsTemp!结束时间), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(i - 1, mCol.续打文件) = NVL(rsTemp!续打文件)
            .TextMatrix(i - 1, mCol.续打文件id) = NVL(rsTemp!续打文件id)
            
            rsTemp.MoveNext
        Loop
    End With
    With chkChoose
        .Value = 0
        .Top = vsfFile.Top - (vsfFile.RowHeight(0) - .Height) / 2
        .Left = vsfFile.ColWidth(mCol.f标志) + (vsfFile.ColWidth(mCol.选择) - .Width) / 2 - vsfFile.Left
    End With
    
    '选择行
    Call vsfFile.Select(intRow, mCol.fID)
    
    zlRefData = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub Allcheck()
    Dim i As Integer
    If chkChoose.Value = Checked Then
        For i = 1 To vsfFile.Rows - 1
            vsfFile.Cell(flexcpChecked, i, mCol.选择, i, mCol.选择) = flexChecked
        Next
    Else
        For i = 1 To vsfFile.Rows - 1
            vsfFile.Cell(flexcpChecked, i, mCol.选择, i, mCol.选择) = flexUnchecked
        Next
    End If
End Sub

Private Sub chkChoose_Click()
    Call Allcheck
End Sub

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngFileID As Long
    Dim lng续打ID As Long
    Dim intCountRetry As Integer
    Dim lngLoop As Long, j As Long
    Dim strFileID As String
    Dim blnCheck As Boolean
    
    lngFileID = 0
    lng续打ID = 0
    intCountRetry = 0
    strFileID = ""
    
    On Error GoTo ErrHand
    
    blnCheck = False
    For lngLoop = 0 To vsfFile.Rows - 1
        If vsfFile.Cell(flexcpChecked, lngLoop, mCol.选择) = flexChecked Then
            blnCheck = True
            Exit For
        End If
    Next
    If Not blnCheck Then Exit Sub
    
    If MsgBox("重算将会对当前选中的记录单文件以及相关记录单文件进行重算数据行,并且会根据参数:护理文件页码规则,对当前选中记录单文件及之后的记录" & _
            "单文件进行页码重整,此操作将会清除记录单文件的打印信息。" & vbCrLf & "请问您是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Screen.MousePointer = 11
        For lngLoop = 0 To vsfFile.Rows - 1
            If vsfFile.Cell(flexcpChecked, lngLoop, mCol.选择) = flexChecked Then
                lngFileID = Val(vsfFile.TextMatrix(lngLoop, mCol.fID))
                lng续打ID = Val(vsfFile.TextMatrix(lngLoop, mCol.续打文件id))
                If lngFileID > 0 And Not InStr(1, strFileID & ",", "," & lngFileID & ",") > 0 Then
                    If frmTendFilePreview.AnaliseData(Me, lngFileID, mlng病人ID, mlng主页ID) Then intCountRetry = intCountRetry + 1
                    strFileID = strFileID & "," & lngFileID
                    If lng续打ID <> 0 And Not chkChoose.Value = Checked And Not InStr(1, strFileID & ",", "," & lng续打ID & ",") > 0 Then
                        For j = 0 To vsfFile.Rows - 1
                            If frmTendFilePreview.AnaliseData(Me, lng续打ID, mlng病人ID, mlng主页ID) Then intCountRetry = intCountRetry + 1
                            strFileID = strFileID & "," & lng续打ID
                            lng续打ID = CheckContinue(lng续打ID)
                            If lng续打ID = 0 Then Exit For
                        Next j
                    End If
                End If
            End If
        Next lngLoop
        For lngLoop = 0 To vsfFile.Rows - 1
            If vsfFile.Cell(flexcpChecked, lngLoop, mCol.选择) = flexChecked Then
                Exit For
            End If
        Next lngLoop
        For j = lngLoop To vsfFile.Rows - 1
            lngFileID = Val(vsfFile.TextMatrix(j, mCol.fID))
            If lngFileID > 0 Then
                gstrSQL = "Zl_病人护理打印_Batchretrypage(" & lngFileID & ",'1;1',0)"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "页码重整")
            End If
        Next
        
        Screen.MousePointer = 0
        mblnSave = True
        MsgBox "已重算" & intCountRetry & "份记录单文件！", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Form_Resize()
    fraMain.Move 0, -90, Me.ScaleWidth, Me.ScaleHeight + 90
    vsfFile.Move 15, 105, vsfFile.Width, fraMain.Height - 105 - cmdOK.Height - 120 * 2
    cmdOK.Move 6015, vsfFile.Height + vsfFile.Top + 120
    cmdCancel.Move cmdOK.Left + cmdOK.Width + 165, cmdOK.Top
End Sub

Private Function CheckContinue(ByVal FileID As Long) As Long
    '功能 : 根据续打id查找续打id,没有返回0
    Dim i As Integer
    Dim lng续打ID As Long
    
    lng续打ID = 0
    For i = 0 To vsfFile.Rows - 1
        If Val(vsfFile.TextMatrix(i, mCol.fID)) = FileID Then
            lng续打ID = Val(vsfFile.TextMatrix(i, mCol.续打文件id))
            Exit For
        End If
    Next i
    CheckContinue = lng续打ID
    
End Function

Private Sub vsfFile_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If NewLeftCol <= mCol.选择 Then
        chkChoose.Visible = True
    Else
        chkChoose.Visible = False
    End If

End Sub

Private Sub vsfFile_AfterUserResize(ByVal ROW As Long, ByVal COL As Long)
    If COL <= mCol.选择 Then
        With chkChoose
            .Value = 0
            .Top = vsfFile.Top - (vsfFile.RowHeight(0) - .Height) / 2
            .Left = vsfFile.ColWidth(mCol.f标志) + (vsfFile.ColWidth(mCol.选择) - .Width) / 2 - vsfFile.Left
        End With
    End If
End Sub

Private Sub vsfFile_Click()
    Call vsfFile_DblClick
End Sub

Private Sub vsfFile_DblClick()
    Dim i As Integer
    Dim blnAllChoose As Boolean
    If vsfFile.COL <> mCol.选择 Then Exit Sub
    If vsfFile.Cell(flexcpChecked, vsfFile.ROW, mCol.选择) = flexUnchecked Then
        vsfFile.Cell(flexcpChecked, vsfFile.ROW, mCol.选择, vsfFile.ROW, mCol.选择) = flexChecked
        blnAllChoose = True
        For i = 1 To vsfFile.Rows - 1
            If vsfFile.Cell(flexcpChecked, i, mCol.选择, i, mCol.选择) <> flexChecked Then
                blnAllChoose = False
                Exit For
            End If
        Next
        If blnAllChoose = True Then chkChoose.Value = Checked
    Else
        If chkChoose.Value = Checked Then
            chkChoose.Value = Unchecked
            For i = 1 To vsfFile.Rows - 1
                If i <> vsfFile.ROW Then
                    vsfFile.Cell(flexcpChecked, i, mCol.选择, i, mCol.选择) = flexChecked
                End If
            Next
        Else
           vsfFile.Cell(flexcpChecked, vsfFile.ROW, mCol.选择, vsfFile.ROW, mCol.选择) = flexUnchecked
        End If
        
    End If
End Sub



