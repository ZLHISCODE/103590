VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersonLoanSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "条件过滤"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2550
      Left            =   1095
      TabIndex        =   25
      Top             =   4395
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4498
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6480
      TabIndex        =   20
      Top             =   255
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6480
      TabIndex        =   21
      Top             =   750
      Width           =   1100
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   3975
      Left            =   135
      TabIndex        =   22
      Top             =   150
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "常规(&R)"
      TabPicture(0)   =   "frmPersonLoanSearch.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmPersonLoanSearch.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl病人ID"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbl病人"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Lbl科室(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtEDIT(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtEDIT(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtEDIT(4)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtEDIT(5)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmd科室"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmd科室 
         Caption         =   "…"
         Height          =   270
         Left            =   -69840
         TabIndex        =   19
         Top             =   2445
         Width           =   255
      End
      Begin VB.Frame fra 
         Caption         =   "范围"
         Height          =   1305
         Index           =   0
         Left            =   105
         TabIndex        =   24
         Top             =   585
         Width           =   5790
         Begin VB.TextBox txtEDIT 
            Height          =   300
            Index           =   1
            Left            =   3240
            MaxLength       =   8
            TabIndex        =   7
            Top             =   795
            Width           =   2085
         End
         Begin VB.TextBox txtEDIT 
            Height          =   300
            Index           =   0
            Left            =   705
            MaxLength       =   8
            TabIndex        =   5
            Top             =   765
            Width           =   2085
         End
         Begin MSComCtl2.DTPicker Dtp开始Date 
            Height          =   300
            Left            =   705
            TabIndex        =   1
            Top             =   375
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   67829763
            CurrentDate     =   37007
         End
         Begin MSComCtl2.DTPicker Dtp结束Date 
            Height          =   300
            Left            =   3240
            TabIndex        =   3
            Top             =   345
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   67829763
            CurrentDate     =   37007
         End
         Begin VB.Label lblCon 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "～"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   2925
            TabIndex        =   2
            Top             =   450
            Width           =   180
         End
         Begin VB.Label lblCon 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "日期"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   270
            TabIndex        =   0
            Top             =   420
            Width           =   360
         End
         Begin VB.Label lblCon 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "开始NO"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   4
            Top             =   810
            Width           =   540
         End
         Begin VB.Label lblCon 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "～"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   2925
            TabIndex        =   6
            Top             =   870
            Width           =   180
         End
      End
      Begin VB.TextBox txtEDIT 
         Height          =   300
         Index           =   5
         Left            =   -73515
         TabIndex        =   18
         Top             =   2445
         Width           =   3705
      End
      Begin VB.TextBox txtEDIT 
         Height          =   300
         Index           =   4
         Left            =   -73515
         MaxLength       =   18
         TabIndex        =   16
         Top             =   1995
         Width           =   2445
      End
      Begin VB.TextBox txtEDIT 
         Height          =   300
         Index           =   3
         Left            =   -73515
         MaxLength       =   18
         TabIndex        =   14
         Top             =   1560
         Width           =   2445
      End
      Begin VB.TextBox txtEDIT 
         Height          =   300
         Index           =   2
         Left            =   -73500
         MaxLength       =   18
         TabIndex        =   12
         Top             =   1110
         Width           =   2445
      End
      Begin VB.Frame fra 
         Caption         =   "请求选项"
         Height          =   705
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   2340
         Width           =   5730
         Begin VB.OptionButton opt范围 
            Caption         =   "所有单据(&0)"
            Height          =   225
            Index           =   0
            Left            =   465
            TabIndex        =   8
            Top             =   300
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton opt范围 
            Caption         =   "门诊收费及记帐(&1)"
            Height          =   225
            Index           =   1
            Left            =   1980
            TabIndex        =   9
            Top             =   300
            Width           =   1920
         End
         Begin VB.OptionButton opt范围 
            Caption         =   "住院记帐(&1)"
            Height          =   225
            Index           =   2
            Left            =   4095
            TabIndex        =   10
            Top             =   300
            Width           =   1485
         End
      End
      Begin VB.Label Lbl科室 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "对方科室(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   -74535
         TabIndex        =   17
         Top             =   2505
         Width           =   990
      End
      Begin VB.Label lbl病人 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人姓名(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74520
         TabIndex        =   15
         Top             =   2070
         Width           =   990
      End
      Begin VB.Label lbl病人ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74340
         TabIndex        =   13
         Top             =   1635
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院号(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74340
         TabIndex        =   11
         Top             =   1170
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmPersonLoanSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrStartDate As String, mstrEndDate As String
Private mstrStartNo As String, mstrEndNo As String
Private mstr单据 As String, mint业务请求 As Integer
Private mstr住院号 As String, mlng病人id As Long, mlng科室id As Long
Private mstr姓名 As String
Dim mblnOk As Boolean
Public Function ShowEdit(ByVal FrmMain As Form, strStartDate As String, strEndDate As String, strStartNo As String, strEndNo As String, _
    str单据 As String, int业务请求 As Integer, str住院号 As String, lng病人id As Long, str姓名 As String, _
    lng科室id As Long) As Boolean
    
    Me.Show 1, FrmMain
    strStartDate = mstrStartDate
    strEndDate = mstrEndDate
    strStartNo = mstrStartNo
    strEndNo = mstrEndNo
    str单据 = mstr单据
    
    int业务请求 = mint业务请求
    str住院号 = mstr住院号
    lng病人id = mlng病人id
    str姓名 = mstr姓名
    lng科室id = mlng科室id
    ShowEdit = mblnOk
    
    
    
End Function

Private Sub chk业务_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 2 Then
            sstFilter.Tab = 1
            If txtEdit(2).Enabled Then txtEdit(2).SetFocus
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub



Private Sub cmd科室_Click()
    Dim rsTemp As New ADODB.Recordset
    
     With rsTemp
        gstrSQL = " Select id,编码,编码||名称 as 名称,简码 " & _
                "   From 部门表 " & _
                "   Where (撤档时间 is null or to_char(撤档时间,'yyyy-mm-dd')='3000-01-01') " & _
                "   order by 编码"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "取部门信息"
        
        If .EOF Then
            MsgBox "部门体系建立不全,请在部门管理中建立！", vbInformation, gstrSysName
            txtEdit(5).SelStart = 0
            txtEdit(5).SelLength = Len(txtEdit(5).Text)
            Exit Sub
        End If
        
        If .RecordCount > 1 Then
            Set mshSelect.Recordset = rsTemp
            With mshSelect
                .Top = sstFilter.Top + txtEdit(5).Top - .Height
                .Left = sstFilter.Left + txtEdit(5).Left
                .Visible = True
                .SetFocus
                .ColWidth(0) = 800
                .ColWidth(1) = 800
                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                Exit Sub
                
            End With
        Else
            txtEdit(5) = IIf(IsNull(!名称), "", !名称)
            txtEdit(5).Tag = NVL(!Id)
            zlCommFun.PressKey (vbKeyTab)
        End If
    End With
End Sub

Private Sub Cmd取消_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmd确定_Click()
    
    If Val(txtEdit(5).Tag) = 0 And txtEdit(5).Text <> "" Then
        MsgBox "科室选择错误,请重新选择!"
        Exit Sub
    End If
    If txtEdit(2).Text <> "" And Not IsNumeric(txtEdit(2).Text) Then
        MsgBox "住院号必需为数字构成,请检查!"
        Exit Sub
    End If
    mstrStartDate = Format(Dtp开始Date.Value, "yyyy-mm-dd HH:MM:SS")
    mstrEndDate = Format(Dtp结束Date.Value, "yyyy-mm-dd HH:MM:SS")
    
    mstrStartNo = Trim(txtEdit(0).Text)
    mstrEndNo = Trim(txtEdit(1).Text)
'
'    mstr单据 = IIf(chk业务(0).Value = 1, ",24", "")
'    mstr单据 = mstr单据 & IIf(chk业务(1).Value = 1, ",25", "")
'    mstr单据 = mstr单据 & IIf(chk业务(2).Value = 1, ",26", "")
'
'    If mstr单据 <> "" Then
'        mstr单据 = Mid(mstr单据, 2)
'    End If
    
    
    mint业务请求 = IIf(opt范围(0).Value, 0, IIf(opt范围(1).Value, 1, 2))
    mstr住院号 = Trim(txtEdit(2).Text)
    mlng病人id = Val(txtEdit(3).Text)
    mstr姓名 = Trim(txtEdit(4).Text)
    
    mlng科室id = Val(txtEdit(5).Tag)
    mblnOk = True
    Unload Me
End Sub

Private Sub Dtp结束Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If

End Sub

Private Sub Dtp开始Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
        Dtp结束Date.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59"
        Dtp开始Date.Value = Format(DateAdd("d", -1, Dtp结束Date.Value), "yyyy-mm-dd") & " 00:00:00"
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
             txtEdit(5).Text = .TextMatrix(.Row, 2)
             txtEdit(5).Tag = Val(.TextMatrix(.Row, 0))
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub opt范围_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEDIT_Change(Index As Integer)
    txtEdit(Index).Tag = ""
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Index <= 1 Then
        Dim intYear As Integer, strYear As String
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(txtEdit(Index)) = "" Then Exit Sub
        '--如果不满八位,则按规则产生--
        Me.txtEdit(Index) = UCase(LTrim(Me.txtEdit(Index)))
        If Len(txtEdit(Index)) < 8 Then
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            Me.txtEdit(Index) = strYear & String(7 - Len(txtEdit(Index)), "0") & Me.txtEdit(Index)
        End If
        zlCommFun.PressKey (vbKeyTab)
        Exit Sub
    End If
    If Index <> 5 Then
          zlCommFun.PressKey (vbKeyTab)
          Exit Sub
    End If
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtEdit(5).Text) = "" Then
            zlCommFun.PressKey (vbKeyTab)
            Exit Sub
        End If
        With rsTemp
            gstrSQL = "" & _
                "   Select id,编码,编码||名称 as 名称,简码 " & _
                "   From 部门表 " & _
                "   Where (撤档时间 is null or to_char(撤档时间,'yyyy-mm-dd')='3000-01-01') and ( upper(编码) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.txtEdit(5) & "%' or Upper(名称) like '" & IIf(gstrMatchMethod = "0", "%", "") & txtEdit(5) & "%' or Upper(简码) like '" & UCase(txtEdit(5)) & "%')" & _
                "   order by 编码"
            
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "取部门信息"
            
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                txtEdit(5).SelStart = 0
                txtEdit(5).SelLength = Len(txtEdit(5).Text)
                
                Exit Sub
            End If
            
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + txtEdit(5).Top - .Height
                    .Left = sstFilter.Left + txtEdit(5).Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txtEdit(5) = IIf(IsNull(!名称), "", !名称)
                txtEdit(5).Tag = NVL(!Id)
                zlCommFun.PressKey (vbKeyTab)
            End If
        End With
    End If
End Sub

Private Sub txtEDIT_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Or Index = 2 Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m数字式
    Else
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    End If
End Sub

