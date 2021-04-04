VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmMediSpecExp 
   Caption         =   "规格扩展信息定义"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5040
   Icon            =   "frmMediSpecExp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   5040
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   2555
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   2
      Top             =   4560
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11850
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfItem 
      Height          =   3500
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   4870
      _cx             =   8590
      _cy             =   6174
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMediSpecExp.frx":6852
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
      Editable        =   2
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
   Begin VB.Label lblComment 
      Caption         =   "保存成功！"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4638
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   在最后一行的项目名称中按回车来增加新项目；在选中行按“Del”键删除该行项目"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   750
      TabIndex        =   0
      Top             =   173
      Width           =   4125
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmMediSpecExp.frx":68EB
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMediSpecExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintItemNameLength As Integer









Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strItems As String
    Dim n As Integer
    Dim i As Integer
    
    With vsfItem
        For n = 1 To .Rows - 1
            '检查名称是否为空
            If Trim(.TextMatrix(n, .ColIndex("项目名称"))) = "" Then
                MsgBox "项目名称不能为空，请录入名称！", vbInformation, gstrSysName
                
                .Row = n
                .TopRow = n
                .Col = .ColIndex("项目名称")
                Exit Sub
            End If
            
            '检查名称是否超长
            If LenB(StrConv(Trim(.TextMatrix(n, .ColIndex("项目名称"))), vbFromUnicode)) > mintItemNameLength Then
                MsgBox "项目名称超长，最大" & mintItemNameLength & "个字符或" & Int(mintItemNameLength / 2) & "个汉字 ！", vbInformation, gstrSysName
                
                .Row = n
                .TopRow = n
                .Col = .ColIndex("项目名称")
                Exit Sub
            End If
            
            '检查名称是否重复
            For i = 1 To .Rows - 1
                If i <> n And Trim(.TextMatrix(i, .ColIndex("项目名称"))) = Trim(.TextMatrix(n, .ColIndex("项目名称"))) Then
                    MsgBox "项目名称已存在，请重新录入名称！", vbInformation, gstrSysName
                    
                    .Row = n
                    .TopRow = n
                    .Col = .ColIndex("项目名称")
                    Exit Sub
                End If
            Next
            
            '拼凑项目串
            If .TextMatrix(n, .ColIndex("项目名称")) <> "" And .TextMatrix(n, .ColIndex("项目名称")) <> .TextMatrix(n, .ColIndex("原项目名称")) Then
                strItems = IIf(strItems = "", "", strItems & "|") & .TextMatrix(n, .ColIndex("编码")) & "," & .TextMatrix(n, .ColIndex("项目名称"))
            End If
        Next
        
        If strItems <> "" Then
            gstrSql = "Zl_药品规格扩展项目_Update("
            '项目串
            gstrSql = gstrSql & "'" & strItems & "'"
            gstrSql = gstrSql & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSql, "保存药品规格扩展项目")
            
            lblComment.Caption = "保存成功！"
            lblComment.Visible = True
        End If
        
        For n = 1 To .Rows - 1
            .TextMatrix(n, .ColIndex("原项目名称")) = .TextMatrix(n, .ColIndex("项目名称"))
        Next
        
        .Cell(flexcpForeColor, 1, .ColIndex("项目名称"), .Rows - 1, .ColIndex("项目名称")) = vbBlack
    End With
End Sub

Private Sub Form_Load()
    Dim rsData As ADODB.Recordset
    
    gstrSql = "Select 编码, 名称 From 药品规格扩展项目 Order By 编码"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "查询药品规格扩展项目")
    
    mintItemNameLength = rsData.Fields("名称").DefinedSize
    
    With vsfItem
        .Rows = 1
        
        If rsData.RecordCount = 0 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("编码")) = "1"
            Exit Sub
        End If
        
        Do While Not rsData.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsData!编码
            .TextMatrix(.Rows - 1, .ColIndex("项目名称")) = rsData!名称
            .TextMatrix(.Rows - 1, .ColIndex("原项目名称")) = rsData!名称
            
            rsData.MoveNext
        Loop
        
    End With
        
End Sub


Private Sub vsfItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 0 Then Exit Sub
    With vsfItem
        If Col <> .ColIndex("项目名称") Then Exit Sub
        If .TextMatrix(Row, .ColIndex("项目名称")) <> .TextMatrix(Row, .ColIndex("原项目名称")) Then
            .Cell(flexcpForeColor, Row, .ColIndex("项目名称")) = vbRed
            lblComment.Visible = False
        Else
            .Cell(flexcpForeColor, Row, .ColIndex("项目名称")) = vbBlack
        End If
    End With
End Sub

Private Sub vsfItem_EnterCell()
    With vsfItem
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        If .Col = .ColIndex("项目名称") Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub vsfItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsData As ADODB.Recordset
    
    With vsfItem
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("项目名称") Then
                .Col = .Col + 1
            ElseIf .Row = .Rows - 1 And .Col = .ColIndex("项目名称") And .TextMatrix(.Row, .ColIndex("项目名称")) <> "" Then
                .Rows = .Rows + 1

                .TextMatrix(.Rows - 1, .ColIndex("编码")) = Val(.TextMatrix(.Rows - 2, .ColIndex("编码"))) + 1
                
                .Row = .Rows - 1
                .Col = .ColIndex("项目名称")
                lblComment.Visible = False
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, .ColIndex("原项目名称")) <> "" Then
                gstrSql = "Select 1 From 药品规格扩展信息 Where 项目 = [1] And Rownum < 2 "
                Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "查询药品规格扩展信息", .TextMatrix(.Row, .ColIndex("项目名称")))
                
                If rsData.RecordCount > 0 Then
                    If MsgBox("已有药品设置了扩展项目“" & .TextMatrix(.Row, .ColIndex("项目名称")) & "”，是否删除？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Else
                    If MsgBox("是否删除扩展项目“" & .TextMatrix(.Row, .ColIndex("项目名称")) & "”？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                gstrSql = "Zl_药品规格扩展项目_Del("
                '编码
                gstrSql = gstrSql & Val(.TextMatrix(.Row, .ColIndex("编码")))
                gstrSql = gstrSql & ")"
        
                Call zlDatabase.ExecuteProcedure(gstrSql, "删除扩展项目")
            End If
            
            If .Row = .Rows - 1 Then
                .TextMatrix(.Row, .ColIndex("项目名称")) = ""
            Else
                .RemoveItem .Row
            End If
            
            lblComment.Visible = False
        End If
    End With
End Sub


Private Sub vsfItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then Exit Sub
    If Col = vsfItem.ColIndex("项目名称") Then
        If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub


