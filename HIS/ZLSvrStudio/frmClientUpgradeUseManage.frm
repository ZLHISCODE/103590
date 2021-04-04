VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClientUpgradeUseManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "客户端用途管理"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12450
   Icon            =   "frmClientUpgradeUseManage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   12450
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboUse 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5400
      Width           =   4005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   360
      Left            =   10080
      TabIndex        =   15
      Top             =   5910
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   11280
      TabIndex        =   16
      Top             =   5910
      Width           =   990
   End
   Begin VB.Frame fraFind 
      Caption         =   "定位"
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   870
         MaxLength       =   3
         TabIndex        =   20
         Tag             =   "IP地址"
         Top             =   390
         Width           =   300
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   21
         Left            =   1305
         MaxLength       =   3
         TabIndex        =   21
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   22
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   22
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   23
         Left            =   2205
         MaxLength       =   3
         TabIndex        =   23
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "定位(&F)"
         Height          =   330
         Left            =   6000
         TabIndex        =   26
         Top             =   330
         Width           =   990
      End
      Begin VB.TextBox txtClient 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3960
         TabIndex        =   25
         Top             =   360
         Width           =   1725
      End
      Begin VB.OptionButton optFind 
         Caption         =   "客户端"
         Height          =   180
         Index           =   1
         Left            =   2970
         TabIndex        =   24
         Top             =   405
         Width           =   850
      End
      Begin VB.OptionButton optFind 
         Caption         =   "IP"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   405
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   840
         MaxLength       =   20
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "IP"
         Text            =   "   ．   ．   ．"
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.Frame fraBatch 
      Caption         =   "客户端"
      Height          =   5055
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   4755
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   3615
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   3525
         _cx             =   6218
         _cy             =   6376
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
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
         Rows            =   1
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
      Begin VB.OptionButton optAdjustType 
         Caption         =   "部门"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optAdjustType 
         Caption         =   "IP范围"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   2910
         MaxLength       =   3
         TabIndex        =   7
         Tag             =   "IP地址"
         Top             =   390
         Width           =   300
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   3345
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   3795
         MaxLength       =   3
         TabIndex        =   9
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   13
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   17
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1110
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "IP地址"
         Top             =   390
         Width           =   300
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1545
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1995
         MaxLength       =   3
         TabIndex        =   5
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtSubIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2445
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "IP地址"
         Top             =   390
         Width           =   315
      End
      Begin VB.CommandButton cmdGetData 
         Caption         =   "刷新(&G)"
         Height          =   360
         Left            =   3600
         TabIndex        =   12
         Top             =   4560
         Width           =   990
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "IP"
         Text            =   "   ．   ．   ．"
         Top             =   360
         Width           =   1725
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "IP"
         Text            =   "   ．   ．   ．"
         Top             =   360
         Width           =   1725
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfUse 
      Height          =   5190
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
      _cx             =   12938
      _cy             =   9155
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
   Begin MSComctlLib.ImageList img16 
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
            Picture         =   "frmClientUpgradeUseManage.frx":038A
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeUseManage.frx":0852
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblUse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用途(&U)"
      Height          =   180
      Left            =   7560
      TabIndex        =   13
      Top             =   5400
      Width           =   630
   End
   Begin VB.Menu mnuPop 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCheck 
         Caption         =   "选定(&A)"
      End
      Begin VB.Menu mnuPopUncheck 
         Caption         =   "取消(&C)"
      End
   End
End
Attribute VB_Name = "frmClientUpgradeUseManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_VSFCOLS As String = "选定,,3,300,B|序号,,3,480|IP,,3,1500|客户端,,3,1500|用途,,3,1800|部门,,3,1500|" & _
    "院区,,3,1000|操作系统,,3,1000|说明,,3,3000"
Private Const MSTR_VSFCOLS_READONLY As String = "序号|IP|客户端|用途|部门|院区|操作系统|说明"

Private WithEvents mobjUse As clsVSFlexGridEx
Attribute mobjUse.VB_VarHelpID = -1
Private WithEvents mobjDept As clsVSFlexGridEx
Attribute mobjDept.VB_VarHelpID = -1
Private mblnClick As Boolean

Private Sub cboUse_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub
 
Private Sub cmdFind_Click()
    Dim i As Long, lngIdx As Long
    Dim strIP As String
    Dim blnFound As Boolean
    
    With vsfUse
        If optFind(1).value Then
            lngIdx = .ColIndex("客户端")
            For i = .FixedRows To .Rows - 1
                If UCase(Trim(.TextMatrix(i, lngIdx))) Like "*" & UCase(Trim(txtClient.Text)) & "*" Then
                    .Row = i
                    blnFound = True
                    Exit For
                End If
            Next
        Else
            lngIdx = .ColIndex("IP")
            For i = .FixedRows To .Rows - 1
                strIP = txtSubIP(20).Text & "." _
                      & txtSubIP(21).Text & "." _
                      & txtSubIP(22).Text & "." _
                      & txtSubIP(23).Text
                If GetIPVal(.TextMatrix(i, lngIdx)) = GetIPVal(strIP) Then
                    .Row = i
                    blnFound = True
                    Exit For
                End If
            Next
        End If
    End With
    
    If blnFound = False Then
        MsgBox "未找到对应的项目！", vbInformation, gstrSysName
        If optFind(1).value Then
            If txtClient.Visible And txtClient.Enabled Then txtClient.SetFocus
        Else
            If txtSubIP(20).Visible And txtSubIP(20).Enabled Then txtSubIP(20).SetFocus
        End If
    Else
        If vsfUse.Enabled And vsfUse.Visible Then vsfUse.SetFocus
    End If
End Sub

Private Sub cmdGetData_Click()
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strDepts As String
    Dim dblStartIP As Double, dblEndIP As Double
    Dim i As Integer, intCheck As Integer, intDept As Integer
    
    On Error GoTo hErr
    
    Screen.MousePointer = vbHourglass
    
    If optAdjustType(0).value Then
        'IP
        dblStartIP = GetIPVal(txtSubIP(0).Text & "." _
                   & txtSubIP(1).Text & "." _
                   & txtSubIP(2).Text & "." _
                   & txtSubIP(3).Text)
        dblEndIP = GetIPVal(txtSubIP(10).Text & "." _
                   & txtSubIP(11).Text & "." _
                   & txtSubIP(12).Text & "." _
                   & txtSubIP(13).Text)
        strSQL = _
            "Select a.工作站 客户端, a.IP, a.用途, a.部门, a.操作系统, a.说明, b.站点 院区 " & vbCr & _
            "From zlClients A, 部门表 B " & vbCr & _
            "Where a.部门 = b.名称(+) " & vbCr & _
            "  And Regexp_Substr(IP, '(\d{1,3})', 1, 1) * 16777216 " & _
            "    + Regexp_Substr(IP, '(\d{1,3})', 1, 2) * 65536 " & _
            "    + Regexp_Substr(IP, '(\d{1,3})', 1, 3) * 256 " & _
            "    + Regexp_Substr(IP, '(\d{1,3})', 1, 4) Between [1] And [2] " & _
            "  And IP Like '%.%.%.%' "
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "指定IP范围获取客户端信息", dblStartIP, dblEndIP)
    Else
        '部门
        strDepts = ""
        With vsfDept
            intDept = .ColIndex("部门")
            intCheck = .ColIndex("选定")
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, intCheck) = flexChecked Then
                    strDepts = strDepts & "," & .TextMatrix(i, intDept)
                End If
            Next
        End With
        If Left(strDepts, 1) = "," Then strDepts = Mid(strDepts, 2)
        
        strSQL = _
            "Select /*+ cardinality(B, 10)*/ a.工作站 客户端, a.Ip, a.用途, a.部门, a.操作系统, a.说明, c.站点 院区 " & vbCr & _
            "From zlClients A, Table(f_str2list([1])) B, 部门表 C " & vbCr & _
            "Where a.部门 = b.Column_Value And a.部门 = c.名称(+)"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "指定部门获取客户端信息", strDepts)
    End If
    
    mobjUse.Recordset = rsTemp
    mobjUse.Repaint RT_Rows
    rsTemp.Close
     
    With vsfUse
        .Col = .ColIndex("IP")
        .Sort = flexSortGenericAscending
        If .Rows > 1 Then
            .TopRow = 1
            .Row = 1
        End If
        mblnClick = True
        Call vsfUse_Click
        mblnClick = False
        Call RefreshSerialNumber
        If .Visible And .Enabled Then .SetFocus
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
hErr:
    Screen.MousePointer = vbDefault
    MsgBox err.Number & "：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String, strUse As String
    Dim colSQL As Collection
    Dim i As Long, lngIdx As Long
    
    If Trim(cboUse.Text) = "" Then
        MsgBox "请填写“用途”！", vbInformation, gstrSysName
        If cboUse.Enabled And cboUse.Visible Then cboUse.SetFocus
        Exit Sub
    ElseIf LenB(StrConv(Trim(cboUse.Text), vbFromUnicode)) > 50 Then
        MsgBox "“用途”不能超过25个汉字或50个字符！", vbInformation, gstrSysName
        If cboUse.Enabled And cboUse.Visible Then cboUse.SetFocus
        Exit Sub
    End If
    If GetChoosedCount(vsfUse, vsfUse.ColIndex("选定"), True) <= 0 Then
        MsgBox "请选定要调整的客户端项目！", vbInformation, gstrSysName
        If vsfUse.Enabled And vsfUse.Visible Then vsfUse.SetFocus
        Exit Sub
    End If
    
    With vsfUse
        '更新工作站的用途
        Set colSQL = New Collection
        lngIdx = .ColIndex("选定")
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, lngIdx) = flexChecked Then
                strSQL = "zl_zlClients_Control(17 " & _
                         ", '" & .TextMatrix(i, .ColIndex("客户端")) & "'" & _
                         ", Null, Null, Null, Null, Null, Null, Null, Null" & _
                         ", '" & Trim(cboUse.Text) & "'" & _
                         ")"
                colSQL.Add strSQL
            End If
        Next
        
        '更新用途字典表
        lngIdx = -1
        strUse = Trim(cboUse.Text)
        For i = 0 To cboUse.ListCount - 1
            If UCase(Trim(cboUse.List(i))) = UCase(strUse) Then
                lngIdx = i
                Exit For
            End If
        Next
        
        If lngIdx < 0 Then
            strSQL = "zl_zlClientsUse_Update(1, Null" & _
                     ", '" & Trim(cboUse.Text) & "'" & _
                     ", 0)"
            colSQL.Add strSQL
        End If
    End With

    If colSQL.Count > 0 Then
        On Error GoTo hErr
        Screen.MousePointer = vbHourglass
        gcnOracle.BeginTrans
        For i = 1 To colSQL.Count
            Call ExecuteProcedure(colSQL(i), Me.name & ".cmdSave", gcnOracle)
        Next
        gcnOracle.CommitTrans
        Set colSQL = Nothing
        Call SaveAuditLog(2, "更新用途", "更新工作站的用途")     '记录日志
        Screen.MousePointer = vbDefault
    End If
    
    Me.Tag = "1"
    Me.Hide
    Exit Sub
    
hErr:
    gcnOracle.RollbackTrans
    Screen.MousePointer = vbDefault
    MsgBox err.Number & "：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    cmdSave.Enabled = False
    Call InitVSF
    Call FillData(cboUse, 0)
    Call FillData(vsfDept, 1)
    Call optFind_Click(0)
    Call optAdjustType_Click(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjUse = Nothing
End Sub

Private Sub mnuPopCheck_Click()
    Call SetChoosed(vsfUse, vsfUse.ColIndex("选定"), True)
End Sub

Private Sub mnuPopUncheck_Click()
    Call SetChoosed(vsfUse, vsfUse.ColIndex("选定"), False)
End Sub

Private Sub optAdjustType_Click(Index As Integer)
    Dim blnIP As Boolean
    
    blnIP = Index = 0
    
    txtIP(0).Enabled = blnIP: txtIP(0).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(0).Enabled = blnIP: txtSubIP(0).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(1).Enabled = blnIP: txtSubIP(1).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(2).Enabled = blnIP: txtSubIP(2).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(3).Enabled = blnIP: txtSubIP(3).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    
    txtIP(1).Enabled = blnIP: txtIP(1).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(10).Enabled = blnIP: txtSubIP(10).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(11).Enabled = blnIP: txtSubIP(11).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(12).Enabled = blnIP: txtSubIP(12).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(13).Enabled = blnIP: txtSubIP(13).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    
    vsfDept.Enabled = Not blnIP
    If Not blnIP Then
        vsfDept.Cell(flexcpBackColor, vsfDept.FixedRows, 0, vsfDept.Rows - 1, vsfDept.Cols - 1) = &H80000005
    Else
        vsfDept.Cell(flexcpBackColor, vsfDept.FixedRows, 0, vsfDept.Rows - 1, vsfDept.Cols - 1) = &H8000000A
        vsfDept.Select 0, 0
    End If
End Sub

Private Sub optFind_Click(Index As Integer)
    Dim blnIP As Boolean
    
    blnIP = Index = 0
    
    txtIP(2).Enabled = blnIP: txtIP(2).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(20).Enabled = blnIP: txtSubIP(20).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(21).Enabled = blnIP: txtSubIP(21).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(22).Enabled = blnIP: txtSubIP(22).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    txtSubIP(23).Enabled = blnIP: txtSubIP(23).BackColor = IIf(blnIP, &H80000005, &H8000000A)
    
    txtClient.Enabled = Not blnIP: txtClient.BackColor = IIf(Not blnIP, &H80000005, &H8000000A)
End Sub

Private Sub txtClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn And Trim(txtClient.Text) <> "" Then
        Call cmdFind_Click
    End If
End Sub

Private Sub txtSubIP_GotFocus(Index As Integer)
    txtSubIP(Index).SelStart = 0
    txtSubIP(Index).SelLength = 3
End Sub

Private Sub txtSubIP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lngLineNo As Long, lngColNo As Long
    
    On Error Resume Next
    
    Call GetCursorPosEx(txtSubIP(Index).hwnd, lngLineNo, lngColNo)
    
    Select Case KeyCode
    Case vbKeyLeft
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtSubIP(Index - 1).Enabled Then
                txtSubIP(Index - 1).SelStart = Len(txtSubIP(Index - 1))
                txtSubIP(Index - 1).SetFocus
            End If
        End If
    Case vbKeyRight
        If Index < 3 Then
            If lngColNo <= Len(txtSubIP(Index)) Then Exit Sub
            If txtSubIP(Index + 1).Enabled Then
                txtSubIP(Index + 1).SelStart = 0
                txtSubIP(Index + 1).SetFocus
            End If
        End If
    Case vbKeyBack
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtSubIP(Index - 1).Enabled Then
                txtSubIP(Index - 1).SelStart = Len(txtSubIP(Index - 1))
                txtSubIP(Index - 1).SetFocus
            End If
        End If
    End Select
End Sub

Private Sub txtSubIP_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim strVal As String
    
    On Error Resume Next
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
        If Len(txtSubIP(Index)) = 2 And txtSubIP(Index).SelLength = 0 Then
            strVal = txtSubIP(Index) & Chr(KeyAscii)
            If val(strVal) > 255 Or val(strVal) < 0 Then
                KeyAscii = 0
                MsgBox "录入IP不正确！", vbInformation, gstrSysName
                txtSubIP(Index).SelStart = 1
                txtSubIP(Index).SelLength = 3
            Else
                txtSubIP(Index + 1).SetFocus
            End If
        End If
    Else
        If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then
                If KeyAscii = Asc(".") Then
                    Call txtSubIP_Validate(Index, blnCancel)
                    If Index + 1 Mod 4 <> 0 And Trim(txtSubIP(Index)) <> "" Then
                        If txtSubIP(Index + 1).Enabled Then
                            If Not blnCancel Then txtSubIP(Index + 1).SetFocus
                        End If
                    End If
                End If
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub txtSubIP_Validate(Index As Integer, Cancel As Boolean)
    If val(txtSubIP(Index).Text) > 255 Or val(txtSubIP(Index).Text) < 0 Then
        MsgBox "录入IP不正确！", vbInformation, gstrSysName
        txtSubIP(Index).SelStart = 0
        txtSubIP(Index).SelLength = 3
        Cancel = True
    ElseIf Trim(txtSubIP(Index)) = "" Then
        txtSubIP(Index) = "0"
    End If
End Sub

Private Sub vsfDept_KeyPress(KeyAscii As Integer)
    Dim blnVal As Boolean
    Dim i As Long, lngCheck As Long
    
    If KeyAscii = vbKeySpace Then
        With vsfDept
            If .SelectedRows <= 0 Then Exit Sub
            
            lngCheck = .ColIndex("选定")
            blnVal = .Cell(flexcpChecked, .SelectedRow(0), lngCheck) = flexChecked
            For i = 0 To .SelectedRows - 1
                .Cell(flexcpChecked, .SelectedRow(i), lngCheck) = IIf(blnVal, flexNoCheckbox, flexChecked)
            Next
        End With
    End If
End Sub

Private Sub vsfUse_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    cmdSave.Enabled = GetChoosedCount(vsfUse, vsfUse.ColIndex("选定"), True) > 0
End Sub

Private Sub vsfUse_AfterSort(ByVal Col As Long, Order As Integer)
    If Col = vsfUse.ColIndex("序号") Then
        Order = flexSortNone
    ElseIf Col = vsfUse.ColIndex("选定") Then
        Order = flexSortNone
        Call vsfUse_Click
    Else
        Call RefreshSerialNumber
    End If
End Sub

Private Sub vsfUse_BeforeSort(ByVal Col As Long, Order As Integer)
    If Col = vsfUse.ColIndex("序号") Then
        Order = flexSortNone
    ElseIf Col = vsfUse.ColIndex("选定") Then
        Order = flexSortCustom
    End If
End Sub

Private Sub vsfUse_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
        Case vsfUse.ColIndex("选定"), vsfUse.ColIndex("IP")
            Cancel = True
    End Select
End Sub

Private Sub InitVSF()
    Set mobjUse = New clsVSFlexGridEx
    With mobjUse
        .AppTemplate EM_Display, vsfUse, MSTR_VSFCOLS, MSTR_VSFCOLS_READONLY, True
        .Init True
        .Binding.Editable = flexEDKbdMouse
        .Binding.ScrollTrack = True
        .Binding.ExplorerBar = flexExSortShow
        .Binding.Cell(flexcpPicture, 0, .Binding.ColIndex("选定")) = img16.ListImages("UnCheck").Picture
    End With
    
    Set mobjDept = New clsVSFlexGridEx
    With mobjDept
        .AppTemplate EM_Display, vsfDept, "选定,,3,300,B|部门,,3,100", "部门", True
        .Init False
        .Binding.Editable = flexEDKbdMouse
        .Binding.ScrollTrack = True
        .Binding.ExplorerBar = flexExNone
        .Binding.ExtendLastCol = True
    End With
End Sub

Private Sub FillData(ByVal ctlVal As Control, ByVal bytType As Byte)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    If bytType = val("0-用途") Then
        strSQL = "Select Distinct 名称 as 用途 From zlClientsUse Where 名称 Is Not Null Order By 用途"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取用途信息")
        With ctlVal
            .Clear
            .addItem ""
            Do While rsTemp.EOF = False
                If Trim("" & rsTemp!用途) <> "" Then .addItem rsTemp!用途
                rsTemp.MoveNext
            Loop
            rsTemp.Close
        End With
    ElseIf bytType = val("1-部件") Then
        strSQL = "Select Distinct 部门 From zlClients Where 部门 Is Not Null Order By 部门"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取客户端部门信息")
        With mobjDept
            .Recordset = rsTemp
            .Repaint RT_Rows
            .Binding.RowHidden(0) = True
        End With
        rsTemp.Close
    End If
        
    Exit Sub
    
hErr:
    MsgBox err.Number & "：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub GetCursorPosEx(ByVal hwnd5 As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long, lParam As Long, wParam As Long, K As Long
    
    err = 0
    On Error Resume Next
    i = SendMessage(hwnd5, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16 '取得目前光标所在位置前有多少个Byte
    LineNo = SendMessage(hwnd5, EM_LINEFROMCHAR, j, 0) '取得光标前面有多少行
    LineNo = LineNo + 1
    K = SendMessage(hwnd5, EM_LINEINDEX, -1, 0)
    '取得目前光标所在行前面有多少个Byte
    ColNo = j - K + 1
End Sub

Private Sub vsfUse_Click()
    Dim intCheck As Integer
    Dim blnChecked As Boolean
    
    With vsfUse
        intCheck = .ColIndex("选定")
        If .MouseRow = 0 And .MouseCol = intCheck Or mblnClick Then
            If .Rows < .FixedRows Then Exit Sub
            
            If mblnClick Then
                .Cell(flexcpPicture, 0, intCheck) = img16.ListImages("AllCheck").Picture
                blnChecked = True
            Else
                If .Cell(flexcpPicture, 0, intCheck) = img16.ListImages("AllCheck").Picture Then
                    .Cell(flexcpPicture, 0, intCheck) = img16.ListImages("UnCheck").Picture
                    blnChecked = False
                Else
                    .Cell(flexcpPicture, 0, intCheck) = img16.ListImages("AllCheck").Picture
                    blnChecked = True
                End If
            End If
            If .Rows > 1 Then .Cell(flexcpChecked, .FixedRows, intCheck, .Rows - 1, intCheck) = blnChecked
            Call vsfUse_AfterEdit(0, vsfUse.ColIndex("选定"))
        End If
    End With
End Sub

Private Sub vsfUse_KeyPress(KeyAscii As Integer)
    Dim blnVal As Boolean
    Dim i As Long, lngCheck As Long
    
    If KeyAscii = vbKeySpace Then
        With vsfUse
            If .SelectedRows <= 0 Then Exit Sub
            
            lngCheck = .ColIndex("选定")
            blnVal = .Cell(flexcpChecked, .SelectedRow(0), lngCheck) = flexChecked
            For i = 0 To .SelectedRows - 1
                .Cell(flexcpChecked, .SelectedRow(i), lngCheck) = IIf(blnVal, flexNoCheckbox, flexChecked)
            Next
        End With
    End If
End Sub

Private Sub vsfUse_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim blnEnabled As Boolean
    
    If Button = vbRightButton And Shift = 0 Then
        blnEnabled = vsfUse.SelectedRows > 0
        mnuPopCheck.Enabled = blnEnabled
        mnuPopUncheck.Enabled = blnEnabled
        Call PopupMenu(mnuPop, , x, y)
    End If
End Sub

Private Function GetChoosedCount(ByVal vsfVal As VSFlexGrid, ByVal intIdx As Integer _
    , Optional ByVal blnExistsExit As Boolean = False) As Long
    
    Dim lngCount As Long, i As Long
    
    With vsfVal
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, intIdx) = flexChecked Then
                lngCount = lngCount + 1
                If blnExistsExit Then Exit For
            End If
        Next
    End With
    GetChoosedCount = lngCount
End Function

Private Sub SetChoosed(ByVal vsfVal As VSFlexGrid, ByVal intIdx As Integer, ByVal blnVal As Boolean)
    Dim i As Long, lngSelRow As Long
    
    With vsfVal
        For i = .FixedRows To .Rows - 1
            If .SelectedRow(lngSelRow) = i Then
                .Cell(flexcpChecked, i, intIdx) = IIf(blnVal, flexChecked, 0)
                lngSelRow = lngSelRow + 1
            End If
        Next
    End With
End Sub

Private Function GetIPVal(ByVal strIP As String) As Double
    Dim arrIP() As String
    
    arrIP = Split(strIP, ".")
    GetIPVal = val(arrIP(0)) * 2 ^ 24 _
             + val(arrIP(1)) * 2 ^ 16 _
             + val(arrIP(2)) * 2 ^ 8 _
             + val(arrIP(3))
End Function

Private Sub RefreshSerialNumber()
    Dim i As Long, lngIdx As Long
    
    With vsfUse
        .Redraw = flexRDNone
        lngIdx = .ColIndex("序号")
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, lngIdx) = i
        Next
        If .Rows > 1 Then
            .TopRow = 1
            .Row = 1
        End If
        .Redraw = flexRDDirect
    End With
End Sub
