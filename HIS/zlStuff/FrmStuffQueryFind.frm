VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmStuffQueryFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "查找材料"
   ClientHeight    =   3165
   ClientLeft      =   3135
   ClientTop       =   4320
   ClientWidth     =   5985
   Icon            =   "FrmStuffQueryFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Pic背景 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   -30
      ScaleHeight     =   3135
      ScaleWidth      =   6135
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   6135
      Begin VB.Frame fra 
         Height          =   75
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   2565
         Width           =   6075
      End
      Begin VB.Frame fra 
         Height          =   45
         Index           =   0
         Left            =   75
         TabIndex        =   18
         Top             =   645
         Width           =   5925
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
         Height          =   1575
         Left            =   585
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2778
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
      Begin VB.CommandButton CmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   135
         Picture         =   "FrmStuffQueryFind.frx":020A
         TabIndex        =   17
         Top             =   2745
         Width           =   1100
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "…"
         Height          =   300
         Left            =   5475
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2130
         Width           =   255
      End
      Begin VB.TextBox TxtSelect产地 
         Height          =   300
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2115
         Width           =   4520
      End
      Begin VB.CommandButton Cmd保存 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3390
         Picture         =   "FrmStuffQueryFind.frx":0354
         TabIndex        =   13
         Top             =   2745
         Width           =   1100
      End
      Begin VB.CommandButton Cmd放弃 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4800
         Picture         =   "FrmStuffQueryFind.frx":049E
         TabIndex        =   14
         Top             =   2745
         Width           =   1100
      End
      Begin VB.TextBox Txt名称 
         Height          =   300
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1290
         Width           =   1875
      End
      Begin VB.TextBox Txt材料编码 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox Txt简码 
         Height          =   300
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1290
         Width           =   1875
      End
      Begin VB.TextBox txt规格 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1710
         Width           =   1875
      End
      Begin VB.TextBox Txt产地 
         Height          =   300
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1740
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lbl 
         Caption         =   "按以述条件查找指定材料的库存,如果同时设置多项,则他们之间是且的关系."
         Height          =   345
         Left            =   915
         TabIndex        =   20
         Top             =   255
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   225
         Picture         =   "FrmStuffQueryFind.frx":05E8
         Top             =   60
         Width           =   480
      End
      Begin VB.Label lbl指定产地 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "指定产地"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   2175
         Width           =   720
      End
      Begin VB.Label Lbl产地 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "产地"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3390
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Lbl规格 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "规格"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   6
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label Lbl助记码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "简码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3390
         TabIndex        =   4
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label Lbl材料编码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "编码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   0
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Lbl通用名称 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   2
         Top             =   1350
         Width           =   360
      End
   End
End
Attribute VB_Name = "FrmStuffQueryFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrTemp As String
Public mstrBit As Byte '该程序查找的匹配方式
Dim mrsTemp As ADODB.Recordset
Public mstrOthers As Variant   '0-编码,1-名称,2-简码,3-规格,4-产地,5-指定产地

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdSelect_Click()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 编码,名称,简码 From 材料生产商  where (站点=[1] or 站点 is null) Order By 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-材料生产商", gstrNodeNo)
        
    With rsTemp
        If .EOF Then
            MsgBox "请初始化卫材生产商（字典管理）！", vbInformation, gstrSysName
             Me.TxtSelect产地.SetFocus: Exit Sub
        End If
                
        If .RecordCount > 1 Then
            Set mshSelect.Recordset = rsTemp
            With mshSelect
                .Top = TxtSelect产地.Top - .Height
                .Left = TxtSelect产地.Left
                .Visible = True
                .SetFocus
                .ColWidth(0) = 800
                .ColWidth(1) = 800
                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                .ZOrder
                Exit Sub
                
            End With
        Else
            TxtSelect产地 = IIf(IsNull(!名称), "", !名称)
            TxtSelect产地.Tag = 1
            Cmd保存.SetFocus
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd保存_Click()
    
    '0-编码,1-名称,2-简码,3-规格,4-产地,5-指定产地
    
    mstrOthers(4) = Trim(Txt产地.Text)
    
    '参数:[1]库房,[2]-编码 ,[3]-名称,[4]-简码,[5]-规格,[6]-产地,[7]-指定产地
    If LTrim(txt名称) = "" And LTrim(Txt材料编码) = "" And LTrim(Txt简码) = "" And LTrim(txt规格) = "" And LTrim(TxtSelect产地) = "" Then MsgBox "请输入至少一项信息！", vbInformation, gstrSysName
    
    mstrTemp = ""
    If LTrim(txt名称) <> "" Then
        mstrTemp = "Q.名称 like [3]"
        mstrOthers(1) = IIf(mstrBit = "0", "%", "") & LTrim(txt名称) & "%"
    End If
    
    If LTrim(Txt材料编码) <> "" Then
        If LTrim(mstrTemp) = "" Then
            mstrTemp = "Q.编码 like [2] "
            mstrOthers(0) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(Txt材料编码)) & "%"
        Else
            mstrTemp = mstrTemp & " And Q.编码 like [2] "
            mstrOthers(0) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(Txt材料编码)) & "%"
        End If
    End If
    
    If LTrim(Txt简码) <> "" Then
        If LTrim(mstrTemp) = "" Then
            mstrTemp = " M.材料id in (Select 收费细目ID from 收费项目别名  where 简码 like [4] )"
            mstrOthers(2) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(Txt简码)) & "%"
               
        Else
            mstrTemp = mstrTemp & " And  M.材料id in (Select 收费细目ID from 收费项目别名  where 简码 like [4] )"
            mstrOthers(2) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(Txt简码)) & "%"
        End If
    End If
    
    If LTrim(txt规格) <> "" Then
        mstrOthers(3) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(txt规格)) & "%"
        If LTrim(mstrTemp) = "" Then
            mstrTemp = " upper(Q.规格) like [5] "
        Else
            mstrTemp = mstrTemp & " And upper(Q.规格) like [5] "
        End If
    End If
    
    If LTrim(TxtSelect产地) <> "" Then
        mstrOthers(5) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(TxtSelect产地)) & "%"
            
        If LTrim(mstrTemp) = "" Then
        
            mstrTemp = "Upper(Q.产地) like [7] "
        Else
            mstrTemp = mstrTemp & " And upper(Q.产地) like [7] "
        End If
    End If
    Me.Hide
End Sub

Private Sub Cmd放弃_Click()
    mstrTemp = ""
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strOthers(0 To 6) As String
    For i = 0 To 6
        strOthers(i) = ""
    Next
    mstrOthers = strOthers
    mstrBit = gstrMatchMethod
End Sub


Private Sub Pic背景_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub TxtSelect产地_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        
        If Trim(TxtSelect产地) = "" Then Exit Sub
        TxtSelect产地 = UCase(TxtSelect产地)
    
        Dim rsTemp As New ADODB.Recordset
        
        On Error GoTo ErrHandle
        gstrSQL = "" & _
            "   Select 编码,名称,简码 " & _
            "   From 材料生产商 " & _
            "   Where (名称 like [1] or 编码 like upper([1]) or  简码 like upper([1])) And (站点=[1] or 站点 is null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "卫材生产商", IIf(gstrMatchMethod = "0", "%", "") & TxtSelect产地 & "%", gstrNodeNo)
        
        With rsTemp
            
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = TxtSelect产地.Top - .Height
                    .Left = TxtSelect产地.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 1000
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    .ZOrder
                    Exit Sub
                    
                End With
            Else
                TxtSelect产地 = IIf(IsNull(!名称), "", !名称)
                TxtSelect产地.Tag = 1
                Cmd保存.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TxtSelect产地_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt产地_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub txt规格_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub Txt简码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)

End Sub

Private Sub Txt名称_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub Txt材料编码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub


Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            TxtSelect产地.Text = .TextMatrix(.Row, 1)
            TxtSelect产地.Tag = 1
            Cmd保存.SetFocus
            .Visible = False
            Exit Sub
        End If
    End With
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

