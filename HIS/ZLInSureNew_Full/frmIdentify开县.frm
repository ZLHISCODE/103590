VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentify开县 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify开县.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      TabIndex        =   14
      Top             =   2985
      Width           =   6030
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   3765
      TabIndex        =   13
      Top             =   3135
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   2310
      TabIndex        =   12
      Top             =   3135
      Width           =   1305
   End
   Begin VB.TextBox txt身份证号 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2100
      MaxLength       =   18
      TabIndex        =   9
      Top             =   1890
      Width           =   2715
   End
   Begin VB.ComboBox cbo性别 
      Height          =   360
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   870
      Width           =   840
   End
   Begin VB.TextBox txt姓名 
      Height          =   360
      Left            =   2100
      TabIndex        =   3
      Top             =   870
      Width           =   1335
   End
   Begin VB.TextBox txtAccount 
      Height          =   360
      Left            =   2100
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtBanlance 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2100
      MaxLength       =   18
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   2715
   End
   Begin MSComCtl2.DTPicker dtpBirthday 
      Height          =   360
      Left            =   2100
      TabIndex        =   7
      Top             =   1380
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   635
      _Version        =   393216
      CustomFormat    =   "yyyy-mm-dd"
      Format          =   87031808
      CurrentDate     =   37243
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   555
      Picture         =   "frmIdentify开县.frx":030A
      Top             =   390
      Width           =   480
   End
   Begin VB.Label lbl身份证号 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号"
      Height          =   240
      Left            =   1065
      TabIndex        =   8
      Top             =   1950
      Width           =   960
   End
   Begin VB.Label lbl出生日期 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期"
      Height          =   240
      Left            =   1065
      TabIndex        =   6
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label lbl性别 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      Height          =   240
      Left            =   3465
      TabIndex        =   4
      Top             =   930
      Width           =   600
   End
   Begin VB.Label lbl姓名 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   240
      Left            =   1515
      TabIndex        =   2
      Top             =   930
      Width           =   510
   End
   Begin VB.Label lblCard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "帐号"
      Height          =   240
      Left            =   1500
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblbanlance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "个人余额"
      Height          =   240
      Left            =   1020
      TabIndex        =   10
      Top             =   2460
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmIdentify开县"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strPatiInfo As String

Private rsTmp As New ADODB.Recordset
Private strSQL As String
Private mintHIS收费 As Integer

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    strPatiInfo = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim(txt姓名.Text) = "" Then
        MsgBox "未正确地输入姓名,请重输！", vbInformation, gstrSysName
        txt姓名.SetFocus
        Exit Sub
    End If
    
    If mintHIS收费 = 1 Then
        If Trim(txtAccount.Text) = "" Then
            MsgBox "必须输入医保卡号,请重输！", vbInformation, gstrSysName
            txtAccount.SetFocus
    
            Exit Sub
        End If
        
        If Trim(txtBanlance.Text) <> "" Then
            If Not IsNumeric(txtBanlance.Text) Then
                MsgBox "个人余额必须为数字型!", vbOKOnly, gstrSysName
                txtBanlance.SelStart = 0
                txtBanlance.SelLength = Len(txtBanlance.Text)
                txtBanlance.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "个人余额不能为空!", vbOKOnly + vbExclamation, gstrSysName
            txtBanlance.SelStart = 0
            txtBanlance.SelLength = Len(txtBanlance.Text)
            txtBanlance.SetFocus
            Exit Sub
        End If
    
    End If
    
    If Trim(txt身份证号.Text) <> "" Then
        If Not IsNumeric(txt身份证号.Text) Then
            MsgBox "身份证号必须为数字如1,2,3等", vbOKOnly, gstrSysName
            txt身份证号.SelStart = 0
            txt身份证号.SelLength = Len(txt身份证号.Text)
            txt身份证号.SetFocus
            Exit Sub
        End If
    End If
    
    Call SaveInfo
    Me.Hide
End Sub

Private Sub SaveInfo()
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
        '9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
    
    Dim strKH As String
    Dim strSelfNo As String
    Dim strSelfPwd As String
    Dim STRNAME As String
    Dim strSex As String
    Dim strBirth As String
    Dim strSFZ As String
    Dim strDWMC As String
    Dim strdwbm As String
    Dim lng病人ID As Long
    
    If mintHIS收费 = 1 Then
        strKH = Trim(txtAccount.Text)
        strSelfNo = Trim(txtAccount.Text)
        gcur帐户余额 = Trim(txtBanlance.Text)
    Else
        strKH = Format(Now, "yyyymmddHHMMSS")
        strSelfNo = Format(Now, "yyyymmddHHMMSS")
        gcur帐户余额 = 0
    End If
        
    strSelfPwd = ""
    STRNAME = Trim(txt姓名.Text)
    strSex = Mid(cbo性别.List(cbo性别.ListIndex), InStr(1, cbo性别.List(cbo性别.ListIndex), "-") + 1)
    strBirth = Format(dtpBirthday.Value, "yyyy-mm-dd")
    strSFZ = Trim(txt身份证号.Text)
    strDWMC = ""
    strdwbm = ""
    
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    strPatiInfo = strKH & ";" & strSelfNo & ";" & strSelfPwd & ";" & _
                    STRNAME & ";" & strSex & ";" & _
                    strBirth & ";" & strSFZ & ";" & _
                    strDWMC & "(" & strdwbm & ")"
    
End Sub

Private Sub dtpBirthday_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New Recordset
    Dim i As Long
    
    mintHIS收费 = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("HIS收费"), 0)
    If mintHIS收费 = 1 Then
        lblCard.Visible = True
        txtAccount.Visible = True
        lblbanlance.Visible = True
        txtBanlance.Visible = True
    End If
    
    strSQL = "Select 编码,名称,简码,缺省标志 From 性别 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo性别.Clear
    If rsTmp.RecordCount <> 0 Then
        For i = 1 To rsTmp.RecordCount
            cbo性别.AddItem rsTmp!名称
            If rsTmp!缺省标志 Then
                cbo性别.ListIndex = i - 1
                cbo性别.ItemData(i - 1) = -1 '作标志
            End If
            rsTmp.MoveNext
        Next
        If cbo性别.ListIndex = -1 Then cbo性别.ListIndex = 0
    End If
    dtpBirthday.Value = Now()
    dtpBirthday.MaxDate = Now()
    gcur帐户余额 = 0
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strPatiInfo = ""
End Sub



Private Sub SeekCob(ByVal ConObj As ComboBox, ByVal strSeek As String)
    Dim intSeek As Integer
    
    ConObj.ListIndex = 0
    If strSeek = "" Then Exit Sub
    
    For intSeek = 0 To ConObj.ListCount
        If ConObj.List(intSeek) = strSeek Then
            ConObj.ListIndex = intSeek
            Exit For
        End If
    Next
End Sub

Private Function GetPatiRec(ByVal strAccount As String) As Recordset
    gstrSQL = "select a.卡号,a.医保号,a.密码,b.姓名,b.性别,b.出生日期,b.身份证号,b.工作单位 " _
        & " from 保险帐户 a,病人信息 b " _
        & " where a.病人id=b.病人id " _
        & " and a.卡号=[1] and a.险类=[2]"
        
        
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strAccount, TYPE_开县)
    Set GetPatiRec = rsTmp
End Function

Private Sub txtAccount_GotFocus()
    With txtAccount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
    Dim rsPati As New Recordset
    
    If KeyAscii = 13 And Trim(txtAccount.Text) <> "" Then
        Set rsPati = GetPatiRec(txtAccount.Text)
        If Not rsPati.EOF Then
            txt姓名.Text = IIf(IsNull(rsPati!姓名), "", rsPati!姓名)
            Call SeekCob(cbo性别, rsPati!性别)
            dtpBirthday.Value = Format(IIf(IsNull(rsPati!出生日期), zlDatabase.Currentdate, rsPati!出生日期), "yyyy-mm-dd")
            txt身份证号.Text = IIf(IsNull(rsPati!身份证号), "", rsPati!身份证号)
       
            txtBanlance.SetFocus
        Else
            txt姓名.SetFocus
            txt姓名.SelStart = 0
            txt姓名.SelLength = Len(txt姓名.Text)
        End If
    End If
End Sub

Private Sub txtBanlance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    End If
End Sub

Private Sub txt身份证号_GotFocus()
    With txt身份证号
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    End If
End Sub

Private Sub txt姓名_GotFocus()
    With txt姓名
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub


