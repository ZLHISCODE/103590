VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSGetDeviceImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "提取设备图像"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   Icon            =   "frmPACSGetDeviceImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "影像接收主机"
      Height          =   630
      Left            =   30
      TabIndex        =   23
      Top             =   5505
      Width           =   10665
      Begin VB.TextBox TxtRemoteAE 
         Height          =   300
         Left            =   9240
         TabIndex        =   17
         Text            =   "XX_SUP"
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox TxtLocalAE 
         Height          =   300
         Left            =   6780
         TabIndex        =   15
         Text            =   "ZLSoftPACS"
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox TxtPort 
         Height          =   300
         Left            =   4680
         TabIndex        =   13
         Text            =   "104"
         Top             =   210
         Width           =   1095
      End
      Begin VB.TextBox TxtIP 
         Height          =   300
         Left            =   1290
         TabIndex        =   11
         Text            =   "LocalHost"
         Top             =   210
         Width           =   2295
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "远程AE(&R)"
         Height          =   180
         Left            =   8340
         TabIndex        =   16
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "本地AE(&L)"
         Height          =   180
         Left            =   5910
         TabIndex        =   14
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "端口号(&P)"
         Height          =   180
         Left            =   3810
         TabIndex        =   12
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "接收主机IP(&I)"
         Height          =   180
         Left            =   75
         TabIndex        =   10
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9585
      TabIndex        =   19
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdDownImage 
      Caption         =   "提取(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8460
      TabIndex        =   18
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdGetImageInfo 
      Caption         =   "检索(&G)"
      Height          =   350
      Left            =   7335
      TabIndex        =   20
      Top             =   6360
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "影像检索条件"
      Height          =   975
      Left            =   30
      TabIndex        =   22
      Top             =   4425
      Width           =   10635
      Begin VB.TextBox TxtStudyUID 
         Height          =   300
         Left            =   1290
         TabIndex        =   9
         Top             =   585
         Width           =   9195
      End
      Begin VB.ComboBox CboSex 
         Height          =   300
         ItemData        =   "frmPACSGetDeviceImage.frx":000C
         Left            =   9600
         List            =   "frmPACSGetDeviceImage.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   915
      End
      Begin VB.TextBox TxtName 
         Height          =   300
         Left            =   6510
         TabIndex        =   5
         Top             =   210
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   210
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   78577665
         CurrentDate     =   38617
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   315
         Left            =   3315
         TabIndex        =   3
         Top             =   210
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   78577665
         CurrentDate     =   38617
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "检查UID(&U)"
         Height          =   180
         Left            =   330
         TabIndex        =   8
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "病人性别(&S)"
         Height          =   180
         Left            =   8490
         TabIndex        =   6
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "病人姓名(&N)"
         Height          =   180
         Left            =   5430
         TabIndex        =   4
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Left            =   3180
         TabIndex        =   2
         Top             =   270
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "检查日期(&D)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   270
         Width           =   990
      End
   End
   Begin MSComctlLib.ListView LvwImageList 
      Height          =   3945
      Left            =   30
      TabIndex        =   21
      Top             =   360
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   6959
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "英文名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "病人ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "性别"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "检查日期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "检查UID"
         Object.Width           =   8114
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "设备影像记录："
      Height          =   180
      Left            =   90
      TabIndex        =   24
      Top             =   105
      Width           =   1260
   End
End
Attribute VB_Name = "frmPACSGetDeviceImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mObjDicomQuery As New DicomQuery
Dim mLngAdvice As Long                      '医嘱ID
Dim mstrDeviceName As String                '设备名

Private Sub cmdOK_Click()

End Sub

Private Sub cboSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetImage_Click()

End Sub

Public Sub ShowMe(objFrom As Object, strIp As String, IntPort As Integer, strDeviceName As String, strLocalAE As String, _
                  strRemoteAE As String, LngAdvice As Long)
    '------------------------------------------------
    '功能：供上级模块调用，并传入所需要的参数
    '参数：strIp;IntPort;strDeviceName;strLocalAE;strRemoteAE,LngAdvice
    '返回：无
    '上级函数或过程：frmPACStation.mnuExecFunc_Click
    '下级函数或过程：无
    '引用的外部参数：mObjDicomQuery
    '编制人：曾超 2005-9-22
    '------------------------------------------------
    With mObjDicomQuery
        .Node = strIp
        .Port = IntPort
        .CalledAE = strRemoteAE
        .CallingAE = strLocalAE
        .Root = "STUDY"
        .Level = "STUDY"
    End With
    mLngAdvice = LngAdvice
    mstrDeviceName = strDeviceName
    
    Me.Caption = "提取" & strDeviceName & "设备的图像"
    Me.Show vbModal, objFrom
    
End Sub


Private Sub cmdDownImage_Click()
    Dim dicGetImages As New DicomImages
    Dim dicGetImage As New DicomImage
            
    If Me.LvwImageList.ListItems.Count < 1 Then Exit Sub
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "接收主机IP", Me.TxtIP
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "端口号", Me.TxtPort
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "本地AE", Me.TxtLocalAE
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "远程AE", Me.TxtRemoteAE
    
    If Len(Trim(Me.TxtIP)) < 1 Then
        MsgBox "必需输入IP地址后才以提取图像!", vbInformation, gstrSysName
        Me.TxtIP.SetFocus
        Exit Sub
    End If
            
    If Len(Trim(Me.TxtPort)) < 1 Then
        MsgBox "必需输入端口号后才以提取图像!", vbInformation, gstrSysName
        Me.TxtPort.SetFocus
        Exit Sub
    End If
    
    With mObjDicomQuery
        .PatientID = Me.LvwImageList.SelectedItem.SubItems(1)
        .StudyUID = Me.LvwImageList.SelectedItem.SubItems(4)
        .SeriesUID = ""
        .InstanceUID = ""
        .Level = "STUDY"
    End With
    
    On Error GoTo GetImageError
    
    zl9comlib.zlCommFun.ShowFlash "请稍等正在读取图像....", Me
    
    '提取图像
    Set dicGetImages = mObjDicomQuery.GetImages
    
    '发送到网关
    For Each dicGetImage In dicGetImages
        dicGetImage.PatientID = mLngAdvice
        dicGetImage.Send Me.TxtIP, Me.TxtPort, TxtLocalAE, TxtRemoteAE
    Next
    
    zl9comlib.zlCommFun.StopFlash
    Unload Me
    
    Exit Sub
GetImageError:
    zl9comlib.zlCommFun.StopFlash

    If MsgBox("获取" & mstrDeviceName & "设备上图像不成功！是否重试？" & vbCrLf & Err.Description, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub cmdGetImageInfo_Click()
    Dim dicGetDates As New DicomDataSets
    Dim dicGetDate As Object
    Dim objItem As ListItem
    Dim strTmp As String
    
    On Error GoTo GetImageInfoErr
            
    With mObjDicomQuery
        .PatientID = ""
        .StudyDate = Format(Me.DTPBegin, "yyyyMMdd") & "-" & Format(Me.DTPEnd, "yyyyMMdd")
        .Name = Trim(Me.TxtName) & "*"
    End With
    
    strTmp = Me.CboSex
    strTmp = Replace(strTmp, "男", "M")
    strTmp = Replace(strTmp, "女", "F")
    mObjDicomQuery.Sex = strTmp
    mObjDicomQuery.StudyUID = Trim(Me.TxtStudyUID) & "*"
                
    Set dicGetDates = mObjDicomQuery.DoQuery
    Me.LvwImageList.ListItems.Clear
    
    For Each dicGetDate In dicGetDates
        Set objItem = Me.LvwImageList.FindItem(dicGetDate.StudyUID)
        If objItem Is Nothing Then
            Set objItem = Me.LvwImageList.ListItems.Add(, "_" & dicGetDate.StudyUID, dicGetDate.Name)
            objItem.SubItems(1) = Nvl(dicGetDate.PatientID)
            strTmp = Nvl(dicGetDate.Sex)
            strTmp = Replace(strTmp, "M", "男")
            strTmp = Replace(strTmp, "F", "女")
            objItem.SubItems(2) = strTmp
            objItem.SubItems(3) = Nvl(dicGetDate(&H8, &H20))
            objItem.SubItems(4) = Nvl(dicGetDate.StudyUID)
        End If
    Next
            
    If Me.LvwImageList.ListItems.Count > 0 Then
        Me.cmdDownImage.Enabled = True
    Else
        Me.cmdDownImage.Enabled = False
    End If
    Exit Sub
GetImageInfoErr:
    If Err.Number = 1011 Then
        If MsgBox("连接" & mstrDeviceName & "设备不成功！是否重试？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Resume
        End If
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub


Private Sub DTPBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub DTPEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    
    Me.TxtIP = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "接收主机IP", "localHost")
    Me.TxtPort = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "端口号", "104")
    Me.TxtLocalAE = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "本地AE", "ZLSoftPACS")
    Me.TxtRemoteAE = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "远程AE", "XX_SUP")
    
    '和查询病人的条件保持一致
    curDate = zlDatabase.Currentdate
    
    If frmPACSFilter.mBeforeDays <= 0 Then frmPACSFilter.mBeforeDays = 3
    
    DTPEnd.MaxDate = curDate: DTPBegin.MaxDate = curDate
    DTPBegin.Value = Format(curDate - frmPACSFilter.mBeforeDays, "yyyy-MM-dd")
    DTPEnd.Value = Format(curDate, "yyyy-MM-dd")
    
End Sub

Private Sub LvwImageList_DblClick()
    cmdDownImage_Click
End Sub

Private Sub TxtIP_GotFocus()
    TxtIP.SelStart = 0
    TxtIP.SelLength = Len(TxtIP.Text)
End Sub

Private Sub TxtIP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub TxtLocalAE_GotFocus()
    TxtLocalAE.SelStart = 0
    TxtLocalAE.SelLength = Len(TxtLocalAE.Text)
End Sub

Private Sub TxtLocalAE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtName_GotFocus()
    TxtName.SelStart = 0
    TxtName.SelLength = Len(TxtName.Text)
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub TxtPort_GotFocus()
    TxtPort.SelStart = 0
    TxtPort.SelLength = Len(TxtPort.Text)
End Sub

Private Sub TxtPort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub TxtRemoteAE_GotFocus()
    TxtRemoteAE.SelStart = 0
    TxtRemoteAE.SelLength = Len(TxtRemoteAE.Text)
End Sub

Private Sub TxtRemoteAE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub TxtStudyUID_GotFocus()
    TxtStudyUID.SelStart = 0
    TxtStudyUID.SelLength = Len(TxtStudyUID.Text)
End Sub

Private Sub TxtStudyUID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub
