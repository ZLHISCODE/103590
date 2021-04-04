VERSION 5.00
Begin VB.Form frmParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frmParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   -120
      TabIndex        =   6
      Top             =   -120
      Width           =   3975
      Begin VB.ComboBox cboCapType 
         Height          =   300
         ItemData        =   "frmParaSet.frx":000C
         Left            =   1635
         List            =   "frmParaSet.frx":0016
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtComInterval 
         Height          =   300
         Left            =   1635
         TabIndex        =   11
         Text            =   "1"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkSaveUI 
         Caption         =   "保存用户界面"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox cboDrivers 
         Height          =   300
         ItemData        =   "frmParaSet.frx":002E
         Left            =   1635
         List            =   "frmParaSet.frx":0030
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cboPort 
         Height          =   300
         ItemData        =   "frmParaSet.frx":0032
         Left            =   1635
         List            =   "frmParaSet.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboDevice 
         Height          =   300
         ItemData        =   "frmParaSet.frx":005E
         Left            =   1635
         List            =   "frmParaSet.frx":006B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "脚踏采集方式"
         Height          =   180
         Left            =   360
         TabIndex        =   12
         Top             =   1520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "脚踏时间间隔                       秒"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   1880
         Width           =   3375
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "输入设备(&I)"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "脚踏端口(&P)"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   420
         Width           =   990
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "存储设备(&F)"
         Height          =   180
         Index           =   8
         Left            =   360
         TabIndex        =   2
         Top             =   780
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2310
      TabIndex        =   5
      Top             =   2790
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   2790
      Width           =   1100
   End
End
Attribute VB_Name = "frmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ifOK As Boolean

Private aDevices() As Variant

Public Function ShowMe(objParent As Object) As Boolean
    Me.Show vbModal, objParent
    ShowMe = ifOK
End Function

Private Sub cboDevice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboPort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "设备号", aDevices(0, cboDevice.ListIndex))
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏端口", cboPort.ListIndex + 1)
    
    If Me.cboDrivers.ListCount > 0 Then
        Call SaveSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Drivers", Me.cboDrivers.ListIndex)
        mConnCapDevice frmImgCapture.hwnd, Me.cboDrivers.ListIndex
    End If
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏擦集方式", cboCapType.ListIndex)
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "保存用户界面", chkSaveUI.Value)
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏时间间隔", IIf(Val(txtComInterval.Text) = 0, 1, Val(txtComInterval.Text)))
    
    ifOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strExeRoom As String
    Dim strDeviceNO As String, iPortNumber As Integer
    Dim i  As Integer
    Dim iCapType As Integer
    Dim strtmp() As String
    
    ifOK = False
    
    On Error GoTo DBError
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1"
    OpenRecordset rsTmp, Me.Caption
    If rsTmp.EOF Then
        MsgBox "未定义影像存储设备，请到影像设备目录中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    aDevices = rsTmp.GetRows: rsTmp.MoveFirst: strDeviceNO = rsTmp(0)
    Me.cboDevice.Clear
    Do While Not rsTmp.EOF
        cboDevice.AddItem Nvl(rsTmp(1))
        rsTmp.MoveNext
    Loop
    
    strDeviceNO = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "设备号", strDeviceNO)
    cboDevice.ListIndex = GetComboxIndex(aDevices, strDeviceNO)
    
    iPortNumber = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏端口", 1))
    If iPortNumber = 0 Then iPortNumber = 1
    cboPort.ListIndex = iPortNumber - 1
    
    strtmp = Split(mGetCapSureDevice(), ";")
    For i = 0 To UBound(strtmp)
        Me.cboDrivers.AddItem strtmp(i)
    Next
    
    i = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Drivers", 0)
    
    If i > 0 And i < Me.cboDrivers.ListCount Then
        Me.cboDevice.ListIndex = i
    Else
        If Me.cboDrivers.ListCount > 0 Then
            Me.cboDrivers.ListIndex = 0
        End If
    End If
    
    iCapType = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏擦集方式", 1))
    
    If iCapType = 0 Then
        cboCapType.ListIndex = 0
    Else
        cboCapType.ListIndex = 1
    End If
    
    chkSaveUI.Value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "保存用户界面", 0)
    
    txtComInterval.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏时间间隔", 1)
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function

