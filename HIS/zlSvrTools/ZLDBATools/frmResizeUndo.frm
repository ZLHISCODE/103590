VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmResizeUndo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "收缩Undo表空间"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSize 
      Height          =   300
      Left            =   6120
      TabIndex        =   8
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtTableSpaceName 
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   3735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfUndo 
      Height          =   1935
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   6375
      _cx             =   1998990317
      _cy             =   1998982485
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
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
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "取消"
      Height          =   350
      Left            =   5760
      TabIndex        =   3
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "确认"
      Height          =   350
      Left            =   4560
      TabIndex        =   2
      Top             =   4440
      Width           =   1100
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3840
      Width           =   5775
   End
   Begin VB.Label lblPrompt 
      Caption         =   $"frmResizeUndo.frx":0000
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label lblNew 
      AutoSize        =   -1  'True
      Caption         =   "新建UNDO表空间"
      Height          =   180
      Left            =   480
      TabIndex        =   12
      Top             =   3000
      Width           =   1260
   End
   Begin VB.Label lblTbs 
      AutoSize        =   -1  'True
      Caption         =   "没有缺省的UNDO表空间"
      Height          =   180
      Left            =   5040
      TabIndex        =   11
      Top             =   720
      Width           =   1800
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "初始大小(M)"
      Height          =   180
      Left            =   5040
      TabIndex        =   7
      Top             =   3420
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "UNDO表空间"
      Height          =   180
      Left            =   480
      TabIndex        =   10
      Top             =   720
      Width           =   900
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   90
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      Caption         =   "名称"
      Height          =   180
      Left            =   480
      TabIndex        =   5
      Top             =   3420
      Width           =   360
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "位置"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   3900
      Width           =   360
   End
End
Attribute VB_Name = "frmResizeUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const conColor As Long = &HEFF0E0            '缺省UNDO表空间颜色

Public Sub ShowMe(ByRef frmParent As Form)
    Me.Show 1, frmParent
End Sub

Private Sub cmdNo_Click()
    Unload Me
End Sub

Private Sub cmdYes_Click()
    Dim strDbf As String, strPath As String, rsCount As ADODB.Recordset
    Dim dblSize As Double
    
    If Trim(txtTableSpaceName.Text) = "" Then
        MsgBox "请输入新的UNDO表空间名称!"
        Exit Sub
    End If
   
    If Trim(txtSize.Text) <> "" Then
        If Not IsNumeric(txtSize.Text) Then
            MsgBox "大小必须是数字!"
            Exit Sub
        End If
        dblSize = Val(txtSize.Text)
    Else
        '缺省100M大小
        dblSize = 100
    End If
    
    On Error GoTo errH
    gstrSQL = "Select 1 From dba_tablespaces where tablespace_name=[1]"
    Set rsCount = OpenSQLRecord(gstrSQL, Me.Caption, UCase(Trim(txtTableSpaceName.Text)))
    
    If rsCount.RecordCount > 0 Then
        MsgBox "指定的表空间已存在，请重新输入!"
    Else
        strDbf = Trim(txtPath.Text) & IIf(InStr(txtPath.Text, "\") > 0, "\", "/") & Trim(txtTableSpaceName.Text) & ".dbf"
        Call ResizeUndo(Trim(txtTableSpaceName.Text), strDbf, dblSize)
    End If
    
    Exit Sub
errH:
    Call ErrCenter(gstrSQL)
End Sub

Private Sub Form_Load()
    Dim strCol As String
    
    strCol = "表空间,1020,0;文件,4575,1;大小(M),705,1"
    vsfUndo.Editable = flexEDNone
    Call InitTable(vsfUndo, strCol)
    
    Call LoadUndo
End Sub

Private Sub LoadUndo()
'加载Undo表空间
    Dim rsTmp As ADODB.Recordset, i As Integer, lngStart As Integer
    Dim blnMult As Boolean, strPreTbs As String
    
    gstrSQL = "select  value  from v$parameter where name = 'undo_tablespace'"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        lblTbs.Caption = "缺省UNDO表空间" & rsTmp!Value
        lblTbs.Tag = rsTmp!Value
    End If
    
    With vsfUndo
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .Editable = flexEDNone
        .AllowUserResizing = flexResizeColumns
                
        gstrSQL = "Select a.Tablespace_Name, Trunc(b.Bytes / 1024 / 1024) Siz, b.File_Name," & vbNewLine & _
                    "       Substr(b.File_Name, 1," & vbNewLine & _
                    "               Decode(Sign(Instr(b.File_Name, '\', -1)), 1, Instr(b.File_Name, '\', -1), Instr(b.File_Name, '/', -1))) File_Path" & vbNewLine & _
                    "From Dba_Tablespaces A, Dba_Data_Files B" & vbNewLine & _
                    "Where a.Contents = 'UNDO' And a.Tablespace_Name = b.Tablespace_Name" & vbNewLine & _
                    "Order By b.File_Name"
        Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption)
        
        lngStart = .FixedRows
        .Redraw = flexRDNone
        .Rows = lngStart
        .Rows = lngStart + rsTmp.RecordCount
        
        For i = 1 To rsTmp.RecordCount
            If i = 1 Then strPreTbs = rsTmp!Tablespace_Name
            
            If strPreTbs <> rsTmp!Tablespace_Name Then
                blnMult = True
                Exit For
            End If
            rsTmp.MoveNext
        Next
        
        i = 1
        rsTmp.MoveFirst
        While Not rsTmp.EOF
            .TextMatrix(i, .ColIndex("表空间")) = rsTmp!Tablespace_Name
            .TextMatrix(i, .ColIndex("文件")) = rsTmp!File_Name
            .TextMatrix(i, .ColIndex("大小(M)")) = rsTmp!Siz
            If lblTbs.Tag = rsTmp!Tablespace_Name Then
                If txtPath.Text = "" Then
                    txtPath.Text = rsTmp!file_path
                    txtTableSpaceName.Text = rsTmp!Tablespace_Name
                    txtSize.Text = "500"
                    .Row = .Rows - 1
                End If
                
                If blnMult Then .Cell(flexcpBackColor, i, .ColIndex("文件"), i, .ColIndex("大小(M)")) = conColor
            End If
            i = i + 1
            rsTmp.MoveNext
        Wend
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub ResizeUndo(strTableSpaceName As String, strPath As String, Optional ByVal intSize As Integer = 50)

    Call SetCommandEnable(0)
    '创建表空间
    lblProgress.Caption = "正在创建新的UNDO表空间。"

    On Error GoTo errH
    gstrSQL = " create undo tablespace " & strTableSpaceName & " datafile '" & strPath & "' size " & intSize & "M AUTOEXTEND ON"
    gcnOracle.Execute gstrSQL
    
    '修改缺省表空间
    lblProgress.Caption = "正在修改缺省UNDO表空间。"
    gstrSQL = "alter system set undo_tablespace=" & strTableSpaceName & " scope=spfile "
    gcnOracle.Execute gstrSQL
    
    '提示信息
    lblProgress.Caption = "修改UNDO表空间为" & strTableSpaceName & "成功(重启后生效)！"

    MsgBox "修改成功,请重启数据库后执行以下语句：" & vbCrLf & _
        "drop tablespace " & lblTbs.Tag & " including contents and datafiles;" & vbCrLf & _
        "注：按快捷键CTRL+C可复制以上脚本。", vbInformation, "操作成功"
    
    Unload Me
    
    Exit Sub
errH:
    Call SetCommandEnable(1)
    Call ErrCenter(gstrSQL)
End Sub


Private Sub SetCommandEnable(bytEnable As Byte)
'功能：设置命令按钮的可用性
    cmdNo.Enabled = bytEnable = 1
    cmdYes.Enabled = cmdNo.Enabled
End Sub


Private Sub Form_Resize()
    
    lblTbs.Left = vsfUndo.Left + (vsfUndo.Width - lblTbs.Width)
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTableSpaceName_KeyPress(KeyAscii As Integer)
    '只允许输大小写字母和数字，下划线
    If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii > 64 And KeyAscii < 91 Or KeyAscii > 96 And KeyAscii < 123 _
        Or KeyAscii = 95 Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

