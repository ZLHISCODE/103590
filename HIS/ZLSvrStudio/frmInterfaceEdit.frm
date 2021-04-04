VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInterfaceEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "三方接口授权"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7935
   Icon            =   "frmInterfaceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInterfaceEdit.frx":151A
   ScaleHeight     =   7110
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7935
      TabIndex        =   17
      Top             =   6375
      Width           =   7935
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F5F5F5&
         Caption         =   "取消(&Q)"
         Height          =   350
         Left            =   6600
         TabIndex        =   14
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00F5F5F5&
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5400
         TabIndex        =   13
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5505
      ScaleWidth      =   7905
      TabIndex        =   16
      Top             =   840
      Width           =   7935
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   2940
         Left            =   1320
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   6300
      End
      Begin VB.ComboBox cboExpiryDate 
         Height          =   300
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1764
         Width           =   975
      End
      Begin VB.TextBox txtExpiryDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1764
         Width           =   1140
      End
      Begin VB.CommandButton cmdCopy 
         BackColor       =   &H00F5F5F5&
         Caption         =   "复制(&C)"
         Default         =   -1  'True
         Height          =   350
         Left            =   3480
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   668
         Width           =   1100
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00F5F5F5&
         Cancel          =   -1  'True
         Caption         =   "重新生成(&N)"
         Height          =   350
         Left            =   4680
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   668
         Width           =   1335
      End
      Begin VB.TextBox txtAppKey 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   693
         Width           =   2145
      End
      Begin VB.TextBox txtAppName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   180
         Width           =   3225
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   345
         Left            =   1320
         TabIndex        =   7
         Top             =   1206
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy/MM/dd HH:mm:ss"
         Format          =   205455363
         CurrentDate     =   43077.4366782407
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         Caption         =   "说明"
         Height          =   180
         Left            =   840
         TabIndex        =   11
         Top             =   2280
         Width           =   360
      End
      Begin VB.Label lblExpiryDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         Caption         =   "效期"
         Height          =   180
         Left            =   840
         TabIndex        =   8
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         Caption         =   "生效时间"
         Height          =   180
         Left            =   480
         TabIndex        =   6
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lblAppKey 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         Caption         =   "授权码"
         Height          =   180
         Left            =   660
         TabIndex        =   2
         Top             =   753
         Width           =   540
      End
      Begin VB.Label lblAppName 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         Caption         =   "接口名称"
         Height          =   180
         Left            =   480
         TabIndex        =   0
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7935
      TabIndex        =   15
      Top             =   0
      Width           =   7935
      Begin VB.Image imgMain 
         Height          =   480
         Left            =   120
         Picture         =   "frmInterfaceEdit.frx":2A34
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblComment 
         BackColor       =   &H00DCDCDC&
         Caption         =   $"frmInterfaceEdit.frx":32FE
         Height          =   615
         Left            =   720
         TabIndex        =   18
         Top             =   120
         Width           =   7095
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmInterfaceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private mlngAppNo   As Long     '授权编号
Private mblnOK      As Boolean
Private mblnChange  As Boolean
Private mstrOldInfo As String
'===========================================================================
'==公共接口
'===========================================================================
Public Function ShowMe(Optional ByVal lngAPPNo As Long) As Boolean
    mlngAppNo = lngAPPNo
    mblnOK = False
    mblnChange = False
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOK
End Function

'===========================================================================
'==事件
'===========================================================================
Private Sub cboExpiryDate_Change()
    mblnChange = True
End Sub

Private Sub cboExpiryDate_Click()
    txtExpiryDate.Enabled = cboExpiryDate.ListIndex > 0
    If Not txtExpiryDate.Enabled Then
        txtExpiryDate.Text = ""
        txtExpiryDate.BackColor = &H8000000F
    Else
        txtExpiryDate.BackColor = &H80000005
    End If
End Sub

Private Sub cmdCancel_Click()
    If mblnChange = True Then
        If MsgBox("界面内容已经编辑，退出将会丢失编辑的内容，确认要退出吗？", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Clipboard.SetText txtAppKey.Text
End Sub

Private Sub cmdNew_Click()
    txtAppKey.Text = GetGUID
    mblnChange = True
End Sub

Private Sub cmdOK_Click()
    Dim strStart        As String
    Dim strEnd          As String
    Dim strInterval     As String
    Dim strNew          As String
    Dim strRemarks      As String
    On Error GoTo errH
    If Trim(txtAppName.Text) = "" Then
        MsgBox "请输入接口名称！", vbInformation, gstrSysName
        txtAppName.SetFocus
        Exit Sub
    End If
    If ActualLen(txtAppName.Text) > txtAppName.MaxLength Then
        MsgBox "接口名称最多允许输入" & txtAppName.MaxLength & "个英文字符或" & txtAppName.MaxLength / 2 & " 个汉字！", vbInformation, gstrSysName
        txtAppName.SetFocus
        Exit Sub
    End If
    If ActualLen(txtNote.Text) > txtNote.MaxLength Then
        MsgBox "接口说明最多允许输入" & txtNote.MaxLength & "个英文字符或" & txtNote.MaxLength / 2 & " 个汉字！", vbInformation, gstrSysName
        txtNote.SetFocus
        Exit Sub
    End If
    
    If txtExpiryDate.Enabled And Val(txtExpiryDate.Text) = 0 Then
        MsgBox "请输入效期！", vbInformation, gstrSysName
        txtExpiryDate.SetFocus
        Exit Sub
    End If
    strStart = "To_Date('" & Format(dtpStartTime.value, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If txtExpiryDate.Enabled = False Then
        strEnd = "To_Date('3000-01-01','YYYY-MM-DD HH24:MI:SS')"
    Else
        strInterval = Decode(cboExpiryDate.ListIndex, 1, "YYYY", 2, "M", 3, "ww", 4, "D", 5, "H")
        strEnd = "To_Date('" & Format(DateAdd(strInterval, Val(txtExpiryDate.Text), dtpStartTime.value), "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    If Not CheckAuditStatus("0316", "接口授权管理", strRemarks) Then Exit Sub
    Call ExecuteProcedure("Zl_Zlinterface_Edit(0," & mlngAppNo & "," & SQLAdjust(txtAppName.Text) & "," & SQLAdjust(Sm4EncryptEcb(txtAppKey.Text, GetGeneralAccountKey(G_APP_KEY))) & "," & SQLAdjust(txtNote.Text) & "," & strStart & "," & strEnd & ")", Me.Caption, gcnOracle)
    If mlngAppNo <> 0 Then
        strNew = strNew & "->"
    End If
    strNew = strNew & MidB("""" & txtAppName.Text & """(授权码:" & txtAppKey.Text & ",生效:" & Format(dtpStartTime.value, "YYYY-MM-DD hh:mm:ss") & ",效期" & IIf(txtExpiryDate.Enabled, txtExpiryDate.Text, "") & cboExpiryDate.Text & ",说明：" & txtNote.Text & "):" & strRemarks, 1, 240)
    Call SaveAuditLog(IIf(mlngAppNo = 0, 1, 2), "接口授权管理", IIf(mlngAppNo = 0, "新增", "修改") & "接口授权信息" & strNew)
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    MsgBox "保存失败！信息：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub Form_Activate()
    txtAppName.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0: PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select Max(Appname) Appname, Max(Key) Key, Max(Note) Note, To_Char(Max(Starttime), 'YYYY-MM-DD hh24:mi:ss') Starttime," & vbNewLine & _
        "       To_Char(Max(Stoptime), 'YYYY-MM-DD hh24:mi:ss') Stoptime,Max(Stoptime)-Max(Starttime) Days" & vbNewLine & _
        "From Zlinterface" & vbNewLine & _
        "Where Appno = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mlngAppNo)
    
    Me.Caption = IIf(mlngAppNo = 0, "新增三方接口授权", "修改三方接口授权")
    cmdNew.Visible = mlngAppNo <> 0
    txtAppName.Text = rsTmp!appName & ""
    txtAppKey.Text = Sm4DecryptEcb(rsTmp!key & "", GetGeneralAccountKey(G_APP_KEY))
    If txtAppKey.Text = "" Then txtAppKey.Text = GetGUID
    If IsNull(rsTmp!Starttime) Then
        dtpStartTime.value = Now
    Else
        dtpStartTime.value = CDate(rsTmp!Starttime & "")
    End If
    Call LoadExpiryDate(rsTmp!Starttime & "", rsTmp!Stoptime & "", rsTmp!Days & "")
    txtNote.Text = rsTmp!Note & ""
    If mlngAppNo <> 0 Then
        mstrOldInfo = MidB("""" & txtAppName.Text & """(授权码:" & txtAppKey.Text & ",生效:" & Format(dtpStartTime.value, "YYYY-MM-DD hh:mm:ss") & ",效期" & IIf(txtExpiryDate.Enabled, txtExpiryDate.Text, "") & cboExpiryDate.Text & ",说明：" & txtNote.Text & ")", 1, 100)
    End If
    mblnChange = False
    Exit Sub
errH:
    MsgBox "加载数据失败，信息：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If mblnChange = True Then
            If MsgBox("界面内容已经编辑，退出将会丢失编辑的内容，确认要退出吗？", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1
            End If
        End If
    End If
End Sub

Private Sub txtAppName_Change()
    mblnChange = True
End Sub

Private Sub txtExpiryDate_Change()
    mblnChange = True
End Sub

Private Sub txtExpiryDate_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 0 And KeyAscii < 32) Then
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNote_Change()
    mblnChange = True
End Sub

Private Sub LoadExpiryDate(ByVal datStart As String, ByVal strStop As String, ByVal strDays As String)

    cboExpiryDate.Clear
    cboExpiryDate.addItem "不限", 0
    cboExpiryDate.addItem "年", 1
    cboExpiryDate.addItem "月", 2
    cboExpiryDate.addItem "周", 3
    cboExpiryDate.addItem "天", 4
    cboExpiryDate.addItem "小时", 5

    If strDays = "" Or strStop Like "3000-01-01*" Then
        cboExpiryDate.ListIndex = 0
        txtExpiryDate.Enabled = False
    Else
        If strDays Like "*.*" Then '存在小数位，则认为是小时
            cboExpiryDate.ListIndex = 5
            txtExpiryDate.Enabled = True
            txtExpiryDate.Text = DateDiff("H", CDate(datStart), CDate(strStop))
        Else
            '先判断年单位,若差的年数包含的天数刚好等于对应天数，则以年为单位
            If DateDiff("D", CDate(datStart), DateAdd("YYYY", DateDiff("YYYY", CDate(datStart), CDate(strStop)), CDate(datStart))) = Val(strDays) Then
                cboExpiryDate.ListIndex = 1
                txtExpiryDate.Enabled = True
                txtExpiryDate.Text = DateDiff("YYYY", CDate(datStart), CDate(strStop))
            '判断月单位,若差的月数包含的天数刚好等于对应天数，则以月为单位
            ElseIf DateDiff("D", CDate(datStart), DateAdd("M", DateDiff("M", CDate(datStart), CDate(strStop)), CDate(datStart))) = Val(strDays) Then
                cboExpiryDate.ListIndex = 2
                txtExpiryDate.Enabled = True
                txtExpiryDate.Text = DateDiff("M", CDate(datStart), CDate(strStop))
            '判断周单位,若差的周数包含的天数刚好等于对应天数，则以月为单位
            ElseIf DateDiff("D", CDate(datStart), DateAdd("ww", DateDiff("ww", CDate(datStart), CDate(strStop)), CDate(datStart))) = Val(strDays) Then
                cboExpiryDate.ListIndex = 3
                txtExpiryDate.Enabled = True
                txtExpiryDate.Text = DateDiff("ww", CDate(datStart), CDate(strStop))
            '都不能取整，则以天为单位
            Else
                cboExpiryDate.ListIndex = 4
                txtExpiryDate.Enabled = True
                txtExpiryDate.Text = DateDiff("D", CDate(datStart), CDate(strStop))
            End If
        End If
    End If
End Sub

Public Function GetGUID() As String
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
                String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
                String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
                IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
                IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
                IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
                IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
                IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
                IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
                IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
                IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
        GetGUID = Mid(GetGUID, 1, 14)
    End If
End Function
