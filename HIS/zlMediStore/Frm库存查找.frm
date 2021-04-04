VERSION 5.00
Begin VB.Form Frm库存查找 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "查找药品"
   ClientHeight    =   2580
   ClientLeft      =   3135
   ClientTop       =   4320
   ClientWidth     =   6525
   Icon            =   "Frm库存查找.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Pic背景 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2925
      ScaleWidth      =   6495
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton CmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   600
         Picture         =   "Frm库存查找.frx":020A
         TabIndex        =   18
         Top             =   2160
         Width           =   1100
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "…"
         Height          =   240
         Left            =   6050
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1695
         Width           =   255
      End
      Begin VB.TextBox TxtSelect产地 
         Height          =   300
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1665
         Width           =   5120
      End
      Begin VB.CommandButton Cmd保存 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3780
         Picture         =   "Frm库存查找.frx":0354
         TabIndex        =   15
         Top             =   2160
         Width           =   1100
      End
      Begin VB.CommandButton Cmd放弃 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5190
         Picture         =   "Frm库存查找.frx":049E
         TabIndex        =   17
         Top             =   2160
         Width           =   1100
      End
      Begin VB.TextBox Txt别名 
         Height          =   300
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   2
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox Txt通用名称 
         Height          =   300
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   1
         Top             =   390
         Width           =   1875
      End
      Begin VB.TextBox Txt药品编码 
         Height          =   300
         Left            =   1200
         TabIndex        =   0
         Top             =   390
         Width           =   1875
      End
      Begin VB.TextBox Txt简码 
         Height          =   300
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   3
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox txt规格 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1260
         Width           =   1875
      End
      Begin VB.TextBox Txt产地 
         Height          =   300
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1290
         Visible         =   0   'False
         Width           =   1875
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
         TabIndex        =   16
         Top             =   1725
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
         Left            =   3990
         TabIndex        =   14
         Top             =   1350
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
         TabIndex        =   13
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Lbl别名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "商品名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   12
         Top             =   900
         Width           =   720
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
         Left            =   3990
         TabIndex        =   11
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Lbl药品编码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "编码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   10
         Top             =   450
         Width           =   360
      End
      Begin VB.Label Lbl通用名称 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "通用名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3600
         TabIndex        =   9
         Top             =   450
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm库存查找"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strTmp As String
Public StrBit As Byte '该程序查找的匹配方式
Dim rsTmp As ADODB.Recordset
Private mfrmMain As Form    '父窗体

Private Type Type_SQLCondition
    str通用名 As String
    str编码 As String
    str简码 As String
    str别名 As String
    str规格 As String
    str产地 As String
End Type

Private SQLCondition As Type_SQLCondition

Public Function GetSearch(ByVal FrmMain As Form, _
    ByRef str通用名 As String, _
    ByRef str编码 As String, _
    ByRef str简码 As String, _
    ByRef str别名 As String, _
    ByRef str规格 As String, _
    ByRef str产地 As String) As String
    strTmp = ""
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, mfrmMain
    GetSearch = strTmp
    
    str通用名 = SQLCondition.str通用名
    str编码 = SQLCondition.str编码
    str简码 = SQLCondition.str简码
    str别名 = SQLCondition.str别名
    str规格 = SQLCondition.str规格
    str产地 = SQLCondition.str产地
    
End Function
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub CmdSelect_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(TxtSelect产地.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select 编码 as id,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码"
    'Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-药品生产商", gstrNodeNo)
    
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "药品生产商", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider.State = 0 Then
        TxtSelect产地.SetFocus
        Exit Sub
    End If
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    TxtSelect产地.Tag = 1
    TxtSelect产地.Text = rsProvider!名称
    Cmd保存.SetFocus
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd保存_Click()
    If LTrim(Txt通用名称) = "" And LTrim(Txt药品编码) = "" And LTrim(Txt别名) = "" & _
        LTrim(Txt简码) = "" And LTrim(txt规格) = "" And LTrim(TxtSelect产地) = "" Then MsgBox "请输入至少一项信息！", vbInformation, gstrSysName
    strTmp = ""
    If LTrim(Txt通用名称) <> "" Then strTmp = strTmp & " And A.名称 like [1] "
    If LTrim(Txt药品编码) <> "" Then strTmp = strTmp & " And A.编码 like [2] "
    If LTrim(Txt简码) <> "" Then strTmp = strTmp & " And B.简码 like [3] "
    If LTrim(Txt别名) <> "" Then strTmp = strTmp & " And B.名称 like [4] "
    If LTrim(txt规格) <> "" Then strTmp = strTmp & " And upper(A.规格) like [5] "
    If LTrim(TxtSelect产地) <> "" Then strTmp = strTmp & " And upper(A.产地) like [6] "
    
    SQLCondition.str通用名 = IIf(StrBit = "0", "%", "") & LTrim(Txt通用名称) & "%"
    SQLCondition.str编码 = IIf(StrBit = "0", "%", "") & UCase(LTrim(Txt药品编码)) & "%"
    SQLCondition.str简码 = IIf(StrBit = "0", "%", "") & UCase(LTrim(Txt简码)) & "%"
    SQLCondition.str别名 = IIf(StrBit = "0", "%", "") & UCase(LTrim(Txt别名)) & "%"
    SQLCondition.str规格 = IIf(StrBit = "0", "%", "") & UCase(LTrim(txt规格)) & "%"
    SQLCondition.str产地 = IIf(StrBit = "0", "%", "") & UCase(LTrim(TxtSelect产地)) & "%"
    
    Unload Me
End Sub

Private Sub Cmd放弃_Click()
    strTmp = ""
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    StrBit = GetSetting(appName:="ZLSOFT", Section:="公共模块\操作", Key:="输入匹配", Default:="0")
End Sub

Private Sub Pic背景_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub TxtSelect产地_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(TxtSelect产地.hWnd)
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(TxtSelect产地) = "" Then Exit Sub
        TxtSelect产地 = UCase(TxtSelect产地)
        
        gstrSQL = "Select 编码 as id ,简码,名称 From 药品生产商 Where (站点 = [3] Or 站点 is Null) And (upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [2]) Order By 编码"
'        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[药品生产商]", _
'                        IIf(gstrMatchMethod = "0", "%", "") & TxtSelect产地 & "%", _
'                        TxtSelect产地 & "%", gstrNodeNo)
        
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "药品生产商", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & TxtSelect产地 & "%", TxtSelect产地 & "%", gstrNodeNo)
        
        If blnCancel Then TxtSelect产地.SetFocus: Exit Sub
        
        With rsTemp
            If rsTemp.State = 0 Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                TxtSelect产地.SelStart = 0
                TxtSelect产地.SelLength = Len(TxtSelect产地)
                KeyCode = 0
                Exit Sub
            End If
            
            TxtSelect产地 = IIf(IsNull(!名称), "", !名称)
            TxtSelect产地.Tag = 1
            Cmd保存.SetFocus
            
        End With
        
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TxtSelect产地_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt别名_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
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

Private Sub Txt通用名称_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub Txt药品编码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub
