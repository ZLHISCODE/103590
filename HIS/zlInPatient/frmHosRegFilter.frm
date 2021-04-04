VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHosRegFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   1935
      TabIndex        =   9
      Top             =   2325
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   195
      TabIndex        =   10
      Top             =   60
      Width           =   5760
      Begin VB.TextBox txt门诊号 
         Height          =   300
         Left            =   990
         TabIndex        =   6
         Top             =   1620
         Width           =   2085
      End
      Begin VB.TextBox txt住院号E 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3510
         TabIndex        =   5
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox txt住院号B 
         Height          =   300
         Left            =   990
         TabIndex        =   4
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   990
         TabIndex        =   2
         Top             =   780
         Width           =   2085
      End
      Begin VB.ComboBox cbo登记员 
         Height          =   300
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker dtp入院E 
         Height          =   300
         Left            =   3495
         TabIndex        =   1
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   103743491
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp入院B 
         Height          =   300
         Left            =   990
         TabIndex        =   0
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   103743491
         CurrentDate     =   40544
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   390
         TabIndex        =   17
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   390
         TabIndex        =   16
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3210
         TabIndex        =   15
         Top             =   1260
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3210
         TabIndex        =   13
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "姓名"
         Height          =   180
         Left            =   390
         TabIndex        =   12
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记员"
         Height          =   180
         Left            =   3150
         TabIndex        =   11
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3105
      TabIndex        =   7
      Top             =   2325
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4290
      TabIndex        =   8
      Top             =   2325
      Width           =   1100
   End
End
Attribute VB_Name = "frmHosRegFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mbytType As Byte '入:病人清单类型
Public mstrFilter As String '出:条件
Public mcllFilter As Collection

Private Sub cbo登记员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo登记员.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo登记员.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo登记员.ListIndex = lngIdx
    If cbo登记员.ListIndex = -1 And cbo登记员.ListCount <> 0 Then cbo登记员.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub

Private Sub cmdOK_Click()
    If IsNumeric(txt住院号E.Text) And IsNumeric(txt住院号B.Text) Then
        If CLng(txt住院号E.Text) <= CLng(txt住院号B.Text) Then
            MsgBox "开始住院号应该小于结束住院号！", vbInformation, gstrSysName
            txt住院号B.SetFocus: Exit Sub
        End If
    End If
    Call MakeFilter
    gblnOK = True
    Hide
End Sub

Private Sub dtp入院E_Change()
    dtp入院B.MaxDate = dtp入院E.Value
End Sub

Private Sub Form_Activate()
    dtp入院B.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim Curdate As Date, i As Integer
    
    txt住院号B.Text = ""
    txt住院号E.Text = ""
    txt门诊号.Text = ""
    
    '设置初始条件(一内月入院)
    Curdate = zlDatabase.Currentdate
    dtp入院B.Value = Format(DateAdd("d", -7, Curdate), "yyyy-MM-dd 00:00:00")
    dtp入院E.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")

    cbo登记员.Clear
    cbo登记员.AddItem "所有登记员"
    cbo登记员.ListIndex = 0
    
    Set rsTmp = GetPersonnel("入院登记员", True)
    For i = 1 To rsTmp.RecordCount
        cbo登记员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        If rsTmp!ID = UserInfo.ID Then cbo登记员.ListIndex = cbo登记员.NewIndex
        rsTmp.MoveNext
    Next
End Sub

Private Sub MakeFilter()
    
    'by lesfeng 2010-1-11 性能优化
    Set mcllFilter = New Collection
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "入院日期"
    mcllFilter.Add Array("", ""), "住院号"
    mcllFilter.Add "", "病人姓名"
    mcllFilter.Add "", "登记人"
    mcllFilter.Add "", "门诊号"
    
'    mstrFilter = ""
'    mstrFilter = mstrFilter & " And B.入院日期 Between To_Date('" & Format(dtp入院B.Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp入院E.Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    
    mstrFilter = ""
    mstrFilter = mstrFilter & " And (B.入院日期  Between [1] And [2]) "
    mcllFilter.Remove "入院日期"
    mcllFilter.Add Array(Format(dtp入院B, "yyyy-MM-dd hh:mm:ss"), Format(dtp入院E, "yyyy-MM-dd hh:mm:ss")), "入院日期"
          
    If IsNumeric(txt住院号B.Text) And IsNumeric(txt住院号E.Text) Then
        mstrFilter = mstrFilter & " And A.病人ID In (Select Distinct 病人ID From 病案主页 Where 住院号 Between [3] And [4]) "
    ElseIf IsNumeric(txt住院号B.Text) Then
        mstrFilter = mstrFilter & " And A.病人ID = (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[3]) "
    End If
    
    mcllFilter.Remove "住院号"
    mcllFilter.Add Array(Trim(txt住院号B.Text), Trim(txt住院号E.Text)), "住院号"
    

    If cbo登记员.ListIndex <> 0 Then
        mstrFilter = mstrFilter & " And B.登记人=[5]"
    End If
    
    '问题17122 by lesfeng 2010-02-02
    If Trim(txt姓名.Text) <> "" Then
        mstrFilter = mstrFilter & " And NVL(B.姓名,A.姓名) like [7]"
    End If
    
    mcllFilter.Remove "病人姓名"
    mcllFilter.Add Trim(txt姓名.Text), "病人姓名"
    
    mcllFilter.Remove "登记人"
    mcllFilter.Add zlCommFun.GetNeedName(cbo登记员.Text), "登记人"
    
    If IsNumeric(txt门诊号.Text) Then mstrFilter = mstrFilter & " And A.门诊号 = [8] "
    mcllFilter.Remove "门诊号"
    mcllFilter.Add Trim(txt门诊号.Text), "门诊号"
    
End Sub

Private Sub txt门诊号_GotFocus()
    zlControl.TxtSelAll txt门诊号
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

'问题17122 by lesfeng 2010-02-02
Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub
'问题17122 by lesfeng 2010-02-02
Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt住院号B_Change()
    txt住院号E.Enabled = (Trim(txt住院号B.Text) <> "")
    If Not txt住院号E.Enabled Then txt住院号E.Text = ""
End Sub

Private Sub txt住院号B_GotFocus()
    zlControl.TxtSelAll txt住院号B
End Sub

Private Sub txt住院号B_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt住院号E_GotFocus()
    zlControl.TxtSelAll txt住院号E
End Sub

Private Sub txt住院号E_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
