VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmAutoPrintOther 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自动打印其他单据"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9750
   Icon            =   "frmAutoPrintOther.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   " 参数"
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   4560
      TabIndex        =   9
      Top             =   1560
      Width           =   5055
      Begin VSFlex8Ctl.VSFlexGrid vsfRptPara 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4800
         _cx             =   8467
         _cy             =   7011
         Appearance      =   0
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
         BackColorSel    =   8421376
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAutoPrintOther.frx":030A
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   " 报表"
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4335
      Begin VSFlex8Ctl.VSFlexGrid vsfRpt 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4080
         _cx             =   7197
         _cy             =   7011
         Appearance      =   0
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
         BackColorSel    =   8421376
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAutoPrintOther.frx":04AB
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
   Begin VB.CommandButton CmdOK 
      Caption         =   "添加(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7200
      TabIndex        =   6
      Top             =   6120
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8400
      TabIndex        =   5
      Top             =   6120
      Width           =   1100
   End
   Begin VB.TextBox txt查找 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   600
      TabIndex        =   3
      Top             =   1140
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10660
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "提示："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   6190
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "*输入报表编码或名称进行查找"
      Height          =   180
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   2430
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "查找"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAutoPrintOther.frx":051C
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmAutoPrintOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrRpt As String   '返回主界面的报表信息：编码,名称

Public Sub ShowForm(ByVal frmParent As Object, ByRef strRpt As String)
        
    frmAutoPrintOther.Show vbModal, frmParent
    
    strRpt = mstrRpt
End Sub

Private Sub ShowRpt(Optional ByVal strCodeOrName As String)
    If grsRpt Is Nothing Then Exit Sub
    
    With vsfRptPara
        .rows = 1
    End With
    
    CmdOK.Enabled = False
    mstrRpt = ""
    
    With vsfRpt
        .Redraw = flexRDNone
        .rows = 1
        
        grsRpt.Filter = IIf(strCodeOrName = "", "", "编号 Like '*" & strCodeOrName & "*' Or 名称 Like '*" & strCodeOrName & "*'")
          
        Do While Not grsRpt.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("编号")) = grsRpt!编号
            .TextMatrix(.rows - 1, .ColIndex("名称")) = grsRpt!名称
            
            grsRpt.MoveNext
        Loop
        
        .Redraw = flexRDDirect
        
        '如果只有一条数据，则定位到该行
        If .rows = 2 Then
            If .TextMatrix(1, .ColIndex("编号")) <> "" Then
                .Row = 1
            End If
        End If
    End With
End Sub

Private Sub ShowRptPara(ByVal strCode As String)
    Dim n As Integer
    Dim blnCheck药房 As Boolean, blnCheck单据 As Boolean, blnCheckNO As Boolean, blnCheckOther As Boolean
     
    lblComment.Caption = ""
    
    With vsfRptPara
        .rows = 1
        
        If grsRpt Is Nothing Then Exit Sub
        If grsRptPara Is Nothing Then Exit Sub
        
        grsRptPara.Filter = "编号= '" & strCode & "'"
        
        .Redraw = flexRDNone
        
        Do While Not grsRptPara.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("编号")) = grsRptPara!编号
            .TextMatrix(.rows - 1, .ColIndex("源名称")) = grsRptPara!数据源名称
            .TextMatrix(.rows - 1, .ColIndex("序号")) = grsRptPara!参数序号
            .TextMatrix(.rows - 1, .ColIndex("参数名称")) = grsRptPara!参数名称
            
            grsRptPara.MoveNext
        Loop
        
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(.ColIndex("源名称")) = True
        
        .Redraw = flexRDDirect
        
        '检查参数是否符合该功能条件，必要有且只有“药房”，“单据”，“NO”这三个参数
        For n = 1 To .rows - 1
            If Trim(.TextMatrix(n, .ColIndex("参数名称"))) = "药房" Then
                blnCheck药房 = True
            End If
            
            If Trim(.TextMatrix(n, .ColIndex("参数名称"))) = "单据" Then
                blnCheck单据 = True
            End If
            
            If Trim(.TextMatrix(n, .ColIndex("参数名称"))) = "NO" Then
                blnCheckNO = True
            End If
            
            '检查有无其他参数
            If blnCheckOther = False Then
                If InStr(1, ",药房,单据,NO,", "," & .TextMatrix(n, .ColIndex("参数名称")) & ",") = 0 Then
                    blnCheckOther = True
                End If
            End If
        Next
        
        If blnCheck药房 = False Or blnCheck单据 = False Or blnCheckNO = False Or blnCheckOther = True Then
            lblComment.Caption = "提示："
            
            If blnCheck药房 = False Or blnCheck单据 = False Or blnCheckNO = False Then
                lblComment.Caption = lblComment.Caption & "缺少"
                
                If blnCheck药房 = False Then
                    lblComment.Caption = lblComment.Caption & "“药房”"
                End If
                
                If blnCheck单据 = False Then
                    lblComment.Caption = lblComment.Caption & IIf(blnCheck药房 = False, "、", "") & "“单据”"
                End If
                
                If blnCheckNO = False Then
                    lblComment.Caption = lblComment.Caption & IIf(blnCheck药房 = False Or blnCheck单据 = False, "、", "") & "“NO”"
                End If
            
                lblComment.Caption = lblComment.Caption & "参数"
            End If
            
            If blnCheckOther = True Then
                If blnCheck药房 = False Or blnCheck单据 = False Or blnCheckNO = False Then
                    lblComment.Caption = lblComment.Caption & "；并且不能有其他参数"
                Else
                    lblComment.Caption = lblComment.Caption & "不能有“药房”、“单据”、“NO”外的其他参数"
                End If
            End If
            
            CmdOK.Enabled = False
        Else
            CmdOK.Enabled = True
        End If
    End With
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If vsfRpt.Row = 0 Then Exit Sub
    If vsfRpt.TextMatrix(vsfRpt.Row, vsfRpt.ColIndex("编号")) = "" Then Exit Sub
    
    mstrRpt = vsfRpt.TextMatrix(vsfRpt.Row, vsfRpt.ColIndex("编号")) & "," & vsfRpt.TextMatrix(vsfRpt.Row, vsfRpt.ColIndex("名称"))
    
    Unload Me
End Sub


Private Sub Form_Load()
    If grsRpt Is Nothing Then
        Set grsRpt = New ADODB.Recordset
        With grsRpt
            If .State = 1 Then .Close
            .Fields.Append "编号", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "名称", adLongVarChar, 50, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        gstrSQL = "Select ID, 编号, 名称 From zlReports Where 系统 = 100 Order By 编号 "
        Set grsRpt = zlDatabase.OpenSQLRecord(gstrSQL, "取所有报表")
    End If
    
    If grsRptPara Is Nothing Then
        Set grsRptPara = New ADODB.Recordset
        With grsRptPara
            If .State = 1 Then .Close
            .Fields.Append "编号", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "数据源名称", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "序号", adDouble, 2, adFldIsNullable
            .Fields.Append "参数名称", adLongVarChar, 60, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        gstrSQL = "Select c.编号, b.名称 As 数据源名称, a.序号 As 参数序号, a.名称 As 参数名称 " & _
            " From zlRPTPars A, zlRPTDatas B, zlReports C " & _
            " Where c.Id = b.报表id And b.Id = a.源id " & _
            " Order By c.编号, b.名称, a.序号 "
        Set grsRptPara = zlDatabase.OpenSQLRecord(gstrSQL, "取所有报表参数")
    End If
    
    Call ShowRpt
End Sub

Private Sub txt查找_Change()
    Call ShowRpt(Trim(txt查找.Text))
End Sub

Private Sub vsfRpt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    
    With vsfRpt
        If .TextMatrix(NewRow, 0) = "" Then Exit Sub
        
        Call ShowRptPara(.TextMatrix(NewRow, .ColIndex("编号")))
    End With
End Sub

