VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFgaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加审计规则"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   Icon            =   "frmFgaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkOpt 
      Caption         =   "Delete"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   15
      Top             =   3240
      Width           =   855
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "Update"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   14
      Top             =   3240
      Width           =   855
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "Insert"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   13
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "Select"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   12
      Top             =   2880
      Width           =   855
   End
   Begin VB.PictureBox pctBottom 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   4320
      ScaleHeight     =   4185
      ScaleWidth      =   1320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1320
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   1095
      End
   End
   Begin VB.TextBox txtPolicy 
      Height          =   300
      Left            =   960
      TabIndex        =   8
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox txtSchema 
      Height          =   300
      Left            =   960
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtObject 
      Height          =   300
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfCols 
      Height          =   1935
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   3375
      _cx             =   5953
      _cy             =   3413
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
      BackColorBkg    =   16777215
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
      ExplorerBar     =   1
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
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   1680
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
            Picture         =   "frmFgaEdit.frx":6852
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFgaEdit.frx":6DEC
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPolicy 
      AutoSize        =   -1  'True
      Caption         =   "规则名"
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   3780
      Width           =   540
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "操作类型"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label lblSchame 
      AutoSize        =   -1  'True
      Caption         =   "所有者"
      Height          =   180
      Left            =   300
      TabIndex        =   4
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lblObject 
      AutoSize        =   -1  'True
      Caption         =   "表名"
      Height          =   180
      Left            =   2280
      TabIndex        =   3
      Top             =   300
      Width           =   360
   End
   Begin VB.Label lblCols 
      AutoSize        =   -1  'True
      Caption         =   "指定列"
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   720
      Width           =   540
   End
End
Attribute VB_Name = "frmFgaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowMe()
    Me.Show 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSchema As String, strObject As String
    Dim strCols As String, strPolicy As String, strType As String
    Dim blnIsAllSelected As Boolean, i As Integer
    Dim strSql As String
    
    On Error GoTo errh
    strSchema = txtSchema.Text
    strObject = txtObject.Text
    strPolicy = txtPolicy.Text
    
    strType = chkOpt(1) & chkOpt(2) & chkOpt(3) & chkOpt(4)
    
    If strSchema = "" Or strSchema = "" Or strSchema = "" Or strType = "" Then
        MsgBox "所有者、表名、操作类型、规则名称必须填写。"
        Exit Sub
    End If
    
    blnIsAllSelected = True
    With vsfCols
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = "-1" Then
                    strCols = IIf(strCols = "", .TextMatrix(i, 1), strCols & "," & .TextMatrix(i, 1))
                Else
                    blnIsAllSelected = False
                End If
            Next
        End If
    End With
    
    For i = 1 To 4
        If chkOpt(i).value = 1 Then
            strPolicy = txtPolicy.Text & "_" & chkOpt(i).Caption
            strType = chkOpt(i).Caption
            
            strSql = "declare" & vbNewLine & _
                            "begin" & vbNewLine & _
                            "    dbms_fga.add_policy(object_schema => '" & strSchema & "'," & vbNewLine & _
                            "                      object_name => '" & strObject & "'," & vbNewLine & _
                            "                      policy_name => '" & strPolicy & "'," & vbNewLine & _
                            "                      audit_column => " & IIf(strCols = "" Or blnIsAllSelected, "Null", "'" & strCols & "'") & "," & vbNewLine & _
                            "                      statement_types => '" & strType & "');" & vbNewLine & _
                            "end;"
        
            gcnOracle.Execute strSql
        End If
    Next
    Unload Me
    Exit Sub
errh:
    MsgBox err.Description
End Sub


Private Sub Form_Activate()
    txtObject.SetFocus
End Sub

Private Sub Form_Load()
    Dim strCol As String
    
    '初始化表格
    With vsfCols
        .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        .Cell(flexcpText, 0, 0) = ""
        .Cell(flexcpPictureAlignment, 0, 0) = flexPicAlignCenterCenter
        .ColWidth(0) = 300
        .ColDataType(0) = flexDTBoolean
        
        .Cell(flexcpText, 0, 1) = "列名"
        .Cell(flexcpPictureAlignment, 0, 0) = flexPicAlignCenterCenter
    End With
    
    If gstrSTOwner = "" Then
        gstrSTOwner = GetOwnerName(100, gcnOracle)
    End If
    txtSchema.Text = gstrSTOwner
End Sub


Private Sub InitCols()
'功能: 根据界面上的所有者和表名获取表对应的列,初始化表格
     Dim strSchema As String, strObject As String
     Dim strSql As String, rsTmp As ADODB.Recordset, i As Integer
    
    On Error GoTo errh
    
    strSchema = txtSchema.Text: strObject = txtObject.Text
    
    If strSchema = "" Or strObject = "" Then Exit Sub
    strSql = "Select Column_Name From Dba_Tab_Cols Where Owner = [1] And Table_Name = [2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "InitCols", UCase(strSchema), UCase(strObject))

    With vsfCols
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rsTmp.RecordCount + .FixedRows
        i = .FixedRows
        Do While Not rsTmp.EOF
            .TextMatrix(i, 0) = "-1"
            .TextMatrix(i, 1) = rsTmp!Column_Name
            i = i + 1
            rsTmp.MoveNext
        Loop
        
        If rsTmp.RecordCount > 0 Then .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
        
        .ColDataType(0) = flexDTBoolean
        .Redraw = flexRDDirect
        If .Rows > 1 Then .Select 1, 0
    End With
    
    Exit Sub
errh:
    MsgBox err.Description
End Sub

Private Sub cboType_Click()
    txtPolicy.Text = "ZLP_" & txtSchema.Text & "_" & txtObject.Text
End Sub
Private Sub txtObject_Change()
    txtPolicy.Text = "ZLP_" & txtSchema.Text & "_" & txtObject.Text
End Sub

Private Sub txtObject_GotFocus()
    txtObject.SelStart = Len(txtObject.Text)
End Sub

Private Sub txtPolicy_GotFocus()
    txtPolicy.SelStart = Len(txtPolicy.Text)
End Sub

Private Sub txtPolicy_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtSchema_Change()
    txtPolicy.Text = "ZLP_" & txtSchema.Text & "_" & txtObject.Text
End Sub
Private Sub txtSchema_GotFocus()
    txtSchema.SelStart = Len(txtSchema.Text)
End Sub

Private Sub txtSchema_KeyPress(KeyAscii As Integer)
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If KeyAscii = 13 Then   '按下回车
        Call InitCols
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtObject_KeyPress(KeyAscii As Integer)
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then   '按下回车
        Call InitCols
        SendKeys "{tab}"
    End If
End Sub


Private Sub vsfCols_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer
    
    With vsfCols
        If Col = 0 Then
            If .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture Then
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "-1"
                Next
            Else
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "0"
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfCols_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, blnAllSelectd As Boolean
    
    blnAllSelectd = True
    With vsfCols
        If .Redraw = flexRDNone Then Exit Sub
        
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 0) = "0" Then
                blnAllSelectd = False
                Exit For
            End If
        Next

        
        If blnAllSelectd Then
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
        Else
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        End If
    End With
End Sub

Private Sub vsfCols_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsfCols_DblClick()
    Call vsfCols_KeyPress(32)
End Sub

Private Sub vsfCols_KeyPress(KeyAscii As Integer)
    Dim i As Integer, blnAllSelectd As Boolean
    
    blnAllSelectd = True
    With vsfCols
        If KeyAscii = 32 And .Col <> 0 Then   '按下空格,进行勾选
             For i = .FixedRows To .Rows - .FixedRows
                If .IsSelected(i) Then
                    .TextMatrix(i, 0) = IIf(.TextMatrix(i, 0) = "-1", 0, -1)
                End If
                If .TextMatrix(i, 0) = "0" Then
                    blnAllSelectd = False
                End If
            Next
        End If
        
        If blnAllSelectd Then
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
        Else
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        End If
    End With
End Sub

