VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSGetDeviceImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ȡ�豸ͼ��"
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
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      Caption         =   "Ӱ���������"
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
         Caption         =   "Զ��AE(&R)"
         Height          =   180
         Left            =   8340
         TabIndex        =   16
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "����AE(&L)"
         Height          =   180
         Left            =   5910
         TabIndex        =   14
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�˿ں�(&P)"
         Height          =   180
         Left            =   3810
         TabIndex        =   12
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��������IP(&I)"
         Height          =   180
         Left            =   75
         TabIndex        =   10
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9585
      TabIndex        =   19
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdDownImage 
      Caption         =   "��ȡ(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8460
      TabIndex        =   18
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdGetImageInfo 
      Caption         =   "����(&G)"
      Height          =   350
      Left            =   7335
      TabIndex        =   20
      Top             =   6360
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ӱ���������"
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
         Caption         =   "���UID(&U)"
         Height          =   180
         Left            =   330
         TabIndex        =   8
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�����Ա�(&S)"
         Height          =   180
         Left            =   8490
         TabIndex        =   6
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��������(&N)"
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
         Caption         =   "�������(&D)"
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
         Text            =   "Ӣ����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�Ա�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "���UID"
         Object.Width           =   8114
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "�豸Ӱ���¼��"
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
Dim mLngAdvice As Long                      'ҽ��ID
Dim mstrDeviceName As String                '�豸��

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
    '���ܣ����ϼ�ģ����ã�����������Ҫ�Ĳ���
    '������strIp;IntPort;strDeviceName;strLocalAE;strRemoteAE,LngAdvice
    '���أ���
    '�ϼ���������̣�frmPACStation.mnuExecFunc_Click
    '�¼���������̣���
    '���õ��ⲿ������mObjDicomQuery
    '�����ˣ����� 2005-9-22
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
    
    Me.Caption = "��ȡ" & strDeviceName & "�豸��ͼ��"
    Me.Show vbModal, objFrom
    
End Sub


Private Sub cmdDownImage_Click()
    Dim dicGetImages As New DicomImages
    Dim dicGetImage As New DicomImage
            
    If Me.LvwImageList.ListItems.Count < 1 Then Exit Sub
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��������IP", Me.TxtIP
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�˿ں�", Me.TxtPort
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "����AE", Me.TxtLocalAE
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Զ��AE", Me.TxtRemoteAE
    
    If Len(Trim(Me.TxtIP)) < 1 Then
        MsgBox "��������IP��ַ�������ȡͼ��!", vbInformation, gstrSysName
        Me.TxtIP.SetFocus
        Exit Sub
    End If
            
    If Len(Trim(Me.TxtPort)) < 1 Then
        MsgBox "��������˿ںź������ȡͼ��!", vbInformation, gstrSysName
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
    
    zl9comlib.zlCommFun.ShowFlash "���Ե����ڶ�ȡͼ��....", Me
    
    '��ȡͼ��
    Set dicGetImages = mObjDicomQuery.GetImages
    
    '���͵�����
    For Each dicGetImage In dicGetImages
        dicGetImage.PatientID = mLngAdvice
        dicGetImage.Send Me.TxtIP, Me.TxtPort, TxtLocalAE, TxtRemoteAE
    Next
    
    zl9comlib.zlCommFun.StopFlash
    Unload Me
    
    Exit Sub
GetImageError:
    zl9comlib.zlCommFun.StopFlash

    If MsgBox("��ȡ" & mstrDeviceName & "�豸��ͼ�񲻳ɹ����Ƿ����ԣ�" & vbCrLf & Err.Description, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    strTmp = Replace(strTmp, "��", "M")
    strTmp = Replace(strTmp, "Ů", "F")
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
            strTmp = Replace(strTmp, "M", "��")
            strTmp = Replace(strTmp, "F", "Ů")
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
        If MsgBox("����" & mstrDeviceName & "�豸���ɹ����Ƿ����ԣ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    
    Me.TxtIP = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��������IP", "localHost")
    Me.TxtPort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�˿ں�", "104")
    Me.TxtLocalAE = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "����AE", "ZLSoftPACS")
    Me.TxtRemoteAE = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Զ��AE", "XX_SUP")
    
    '�Ͳ�ѯ���˵���������һ��
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
