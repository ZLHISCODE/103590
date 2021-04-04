VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmNoticeList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdEdit 
      Caption         =   "编辑(&E)"
      Height          =   345
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   990
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "停用(&S)"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Top             =   2400
      Width           =   990
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfNoticeSet 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      _cx             =   8916
      _cy             =   3625
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   End
End
Attribute VB_Name = "frmNoticeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEdit_Click()
    With vsfNoticeSet
        .TextMatrix(.Row, .ColIndex("通知间隔(S)")) = frmListSet.ShowEdit(.TextMatrix(.Row, .ColIndex("编号")), .TextMatrix(.Row, .ColIndex("名称")), .TextMatrix(.Row, .ColIndex("说明")), Val(.TextMatrix(.Row, .ColIndex("通知间隔(S)"))))
    End With
End Sub

Private Sub cmdStop_Click()
    Dim strSql As String
    
    On Error GoTo errH
    With vsfNoticeSet
        If .TextMatrix(.Row, .ColIndex("状态")) = "在用" Then
            strSql = "Update zltools.zlnoticelists Set Status = 0 Where NoticeCode = " & .TextMatrix(.Row, .ColIndex("编号"))
            .TextMatrix(.Row, .ColIndex("状态")) = "停用"
            .Cell(flexcpForeColor, .Row, .ColIndex("状态")) = vbRed
            cmdStop.Caption = "启用(&S)"
        Else
            strSql = "Update zltools.zlnoticelists Set Status = 1 Where NoticeCode = " & .TextMatrix(.Row, .ColIndex("编号"))
            .TextMatrix(.Row, .ColIndex("状态")) = "在用"
            .Cell(flexcpForeColor, .Row, .ColIndex("状态")) = vbBlack
            cmdStop.Caption = "停用(&S)"
        End If
    End With
    
    gcnOracle.Execute strSql
    MsgBox "修改成功，重启程序后生效。", , "提示"
    
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub Form_Load()
    Call InitVsf
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With vsfNoticeSet
        .Width = Me.ScaleWidth - .Left - 60
        .Height = Me.ScaleHeight - .Top - cmdStop.Height - 120
    End With
    
    With cmdStop
        .Top = vsfNoticeSet.Top + vsfNoticeSet.Height + 60
        .Left = Me.ScaleWidth - .Width - 60
    End With
    
    With cmdEdit
        .Top = cmdStop.Top
        .Left = cmdStop.Left - .Width - 60
    End With
End Sub

Public Sub SetDataSource(ByVal rsData As ADODB.Recordset)
    Dim i As Integer
    
    rsData.Filter = 0
    
    With vsfNoticeSet
        .Redraw = flexRDNone
        
        .Rows = 1: i = 1
        .Rows = rsData.RecordCount + 1
        
        Do While Not rsData.EOF
            .TextMatrix(i, .ColIndex("编号")) = rsData!Noticecode & ""
            .TextMatrix(i, .ColIndex("名称")) = rsData!Noticename & ""
            .TextMatrix(i, .ColIndex("所有者")) = rsData!Tableowner & ""
            .TextMatrix(i, .ColIndex("表名")) = rsData!Tablename & ""
            .TextMatrix(i, .ColIndex("说明")) = rsData!Comments & ""
            .TextMatrix(i, .ColIndex("通知间隔(S)")) = rsData!Interval & ""
            
            If rsData!Status & "" = "1" Then
                .TextMatrix(i, .ColIndex("状态")) = "在用"
                .Cell(flexcpForeColor, i, .ColIndex("状态")) = vbBlack
            Else
                .TextMatrix(i, .ColIndex("状态")) = "停用"
                .Cell(flexcpForeColor, i, .ColIndex("状态")) = vbRed
            End If
            
            i = i + 1
            rsData.MoveNext
        Loop
        
        .Redraw = flexRDDirect
        .AutoResize = True
        .AutoSize 0, .Cols - 1
        
        If .Rows > 1 Then .Select 1, 0
    End With

End Sub

Private Sub InitVsf()
    Dim strCols As String

    strCols = "编号,0,1;名称,500,1;所有者,300,1;表名,300,1;通知间隔(S),200,1;说明,2000,1;状态,300,1"
    InitTable vsfNoticeSet, strCols
End Sub

Private Sub vsfNoticeSet_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If OldRow = NewRow Then Exit Sub
    
    If vsfNoticeSet.TextMatrix(NewRow, vsfNoticeSet.ColIndex("状态")) = "在用" Then
        cmdStop.Caption = "停用(&S)"
    Else
        cmdStop.Caption = "启用(&S)"
    End If
End Sub

