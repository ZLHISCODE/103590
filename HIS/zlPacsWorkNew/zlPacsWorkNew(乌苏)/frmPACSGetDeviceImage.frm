VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSGetDeviceImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ȡ�豸ͼ��"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   Icon            =   "frmPACSGetDeviceImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin DicomObjects.DicomViewer Viewer 
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   8040
      Width           =   495
      _Version        =   262147
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   35
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ѯ����"
      Height          =   3735
      Left            =   9120
      TabIndex        =   15
      Top             =   3360
      Width           =   1695
      Begin VB.CheckBox chkLog 
         Caption         =   "��¼ͨѶ��־"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "ȡͼ��ʽ"
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
         Begin VB.OptionButton optRetrieveType 
            Caption         =   "�첽Move"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optRetrieveType 
            Caption         =   "ͬ��Move"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame frmQueryRoot 
         Caption         =   "��ѯRoot"
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton optQueryRoot 
            Caption         =   "���"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optQueryRoot 
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ӱ���������"
      Height          =   630
      Left            =   120
      TabIndex        =   12
      Top             =   7200
      Width           =   10665
      Begin VB.TextBox TxtRemoteAE 
         Height          =   300
         Left            =   9240
         TabIndex        =   7
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox TxtLocalAE 
         Height          =   300
         Left            =   6780
         TabIndex        =   5
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox TxtPort 
         Height          =   300
         Left            =   4680
         TabIndex        =   3
         Top             =   210
         Width           =   1095
      End
      Begin VB.TextBox TxtIP 
         Height          =   300
         Left            =   1290
         TabIndex        =   1
         Top             =   210
         Width           =   2295
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Զ��AE(&R)"
         Height          =   180
         Left            =   8340
         TabIndex        =   6
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "����AE(&L)"
         Height          =   180
         Left            =   5910
         TabIndex        =   4
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�˿ں�(&P)"
         Height          =   180
         Left            =   3810
         TabIndex        =   2
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��������IP(&I)"
         Height          =   180
         Left            =   75
         TabIndex        =   0
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9585
      TabIndex        =   9
      Top             =   8040
      Width           =   1100
   End
   Begin VB.CommandButton cmdDownImage 
      Caption         =   "��ȡ(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8460
      TabIndex        =   8
      Top             =   8040
      Width           =   1100
   End
   Begin VB.CommandButton cmdGetImageInfo 
      Caption         =   "����(&G)"
      Height          =   350
      Left            =   7320
      TabIndex        =   10
      Top             =   8040
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ӱ���������"
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   3345
      Width           =   8895
      Begin VB.CheckBox chkLevelSeries 
         Caption         =   "���м���"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkLevelStudy 
         Caption         =   "��鼶��"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Height          =   1455
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   8655
         Begin VB.TextBox txtSPStepID 
            Height          =   300
            Left            =   1680
            TabIndex        =   45
            Top             =   600
            Width           =   2300
         End
         Begin VB.TextBox txtRProcedureID 
            Height          =   300
            Left            =   6240
            TabIndex        =   43
            Top             =   240
            Width           =   2300
         End
         Begin VB.TextBox txtSeriesNumber 
            Height          =   300
            Left            =   6240
            TabIndex        =   41
            Top             =   600
            Width           =   2300
         End
         Begin VB.TextBox txtModality 
            Height          =   300
            Left            =   1680
            TabIndex        =   39
            Top             =   240
            Width           =   2300
         End
         Begin MSComCtl2.DTPicker dtpPPSStart 
            Height          =   315
            Left            =   1680
            TabIndex        =   48
            Top             =   960
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            DateIsNull      =   -1  'True
            Format          =   62324739
            CurrentDate     =   38617
            MinDate         =   -109174
         End
         Begin MSComCtl2.DTPicker dtpPPSEnd 
            Height          =   315
            Left            =   6240
            TabIndex        =   49
            Top             =   960
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            DateIsNull      =   -1  'True
            Format          =   62324739
            CurrentDate     =   38617
            MinDate         =   -109174
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   180
            Left            =   5040
            TabIndex        =   52
            Top             =   1027
            Width           =   210
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "PPS Start Date"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   1027
            Width           =   1260
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "SP Step ID"
            Height          =   180
            Left            =   120
            TabIndex        =   46
            Top             =   660
            Width           =   900
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Request Procedure ID"
            Height          =   180
            Left            =   4320
            TabIndex        =   44
            Top             =   300
            Width           =   1800
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "���к�"
            Height          =   180
            Left            =   4920
            TabIndex        =   42
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Ӱ�����"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   8655
         Begin VB.ComboBox CboSex 
            Height          =   300
            ItemData        =   "frmPACSGetDeviceImage.frx":000C
            Left            =   6195
            List            =   "frmPACSGetDeviceImage.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   600
            Width           =   2300
         End
         Begin VB.TextBox txtReferringDoctor 
            Height          =   300
            Left            =   6195
            TabIndex        =   36
            Top             =   1320
            Width           =   2300
         End
         Begin VB.TextBox txtModalitiesInStudy 
            Height          =   300
            Left            =   1680
            TabIndex        =   34
            Top             =   1320
            Width           =   2300
         End
         Begin VB.TextBox txtStudyID 
            Height          =   300
            Left            =   6195
            TabIndex        =   32
            Top             =   960
            Width           =   2300
         End
         Begin VB.TextBox TxtName 
            Height          =   300
            Left            =   1680
            TabIndex        =   30
            Top             =   600
            Width           =   2300
         End
         Begin VB.TextBox txtAccessionNumber 
            Height          =   300
            Left            =   1680
            TabIndex        =   28
            Top             =   960
            Width           =   2300
         End
         Begin MSComCtl2.DTPicker DTPBegin 
            Height          =   315
            Left            =   1680
            TabIndex        =   24
            Top             =   240
            Width           =   2300
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   62324739
            CurrentDate     =   38617
            MinDate         =   -109174
         End
         Begin MSComCtl2.DTPicker DTPEnd 
            Height          =   315
            Left            =   6195
            TabIndex        =   25
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   62324739
            CurrentDate     =   38617
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "�����Ա�"
            Height          =   180
            Left            =   4920
            TabIndex        =   51
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "����ҽ��"
            Height          =   180
            Left            =   4920
            TabIndex        =   37
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "����Ӱ�����"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   1380
            Width           =   1260
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   4920
            TabIndex        =   33
            Top             =   1020
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Accession Number"
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   1020
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   307
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   180
            Left            =   4980
            TabIndex        =   26
            Top             =   307
            Width           =   210
         End
      End
   End
   Begin MSComctlLib.TreeView trvList 
      Height          =   2895
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   5106
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "�豸Ӱ���¼��"
      Height          =   180
      Left            =   90
      TabIndex        =   13
      Top             =   105
      Width           =   1260
   End
End
Attribute VB_Name = "frmPACSGetDeviceImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mObjDicomQuery As New DicomQuery
Dim mLngAdvice As Long                      'ҽ��ID
Dim mstrDeviceName As String                '�豸��

Private Sub cboSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub chkLevelSeries_Click()
    If chkLevelSeries.value = 1 Then
        chkLevelStudy.value = 0
    Else
        chkLevelStudy.value = 1
    End If
End Sub

Private Sub chkLevelStudy_Click()
    If chkLevelStudy.value = 0 Then
        chkLevelSeries.value = 1
    Else
        chkLevelSeries.value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub ShowMe(objFrom As Object, strIP As String, IntPort As Integer, strDeviceName As String, strLocalAE As String, _
                  strRemoteAE As String, LngAdvice As Long)
    '------------------------------------------------
    '���ܣ����ϼ�ģ����ã�����������Ҫ�Ĳ���
    '������strIp;IntPort;strDeviceName;strLocalAE;strRemoteAE,LngAdvice
    '���أ���
    '------------------------------------------------
    With mObjDicomQuery
        .Node = strIP
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
    Dim oneNode As Node
    Dim qry As DicomQuery
    Dim OK As Boolean
            
    If Me.trvList.SelectedItem Is Nothing Then Exit Sub
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������IP", Me.TxtIP
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�˿ں�", Me.TxtPort
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����AE", Me.TxtLocalAE
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Զ��AE", Me.TxtRemoteAE
    
    If Len(Trim(Me.TxtIP)) < 1 Then
        MsgBoxD Me, "��������IP��ַ�������ȡͼ��!", vbInformation, gstrSysName
        Me.TxtIP.SetFocus
        Exit Sub
    End If
            
    If Len(Trim(Me.TxtPort)) < 1 Then
        MsgBoxD Me, "��������˿ںź������ȡͼ��!", vbInformation, gstrSysName
        Me.TxtPort.SetFocus
        Exit Sub
    End If
    
    Set oneNode = trvList.SelectedItem
    Set qry = New DicomQuery
    
    Select Case oneNode.Tag(0)
        Case 1  '���˼���
            qry.PatientID = oneNode.Tag(1)
            qry.Level = "PATIENT"
        Case 2  '��鼶��
            qry.PatientID = oneNode.Tag(1)
            qry.StudyUID = oneNode.Tag(2)
            qry.Level = "STUDY"
        Case 3  '���м���
            qry.PatientID = oneNode.Tag(1)
            qry.StudyUID = oneNode.Tag(2)
            qry.SeriesUID = oneNode.Tag(3)
            qry.Level = "SERIES"
        Case 4  'ͼ�񼶱�
            qry.PatientID = oneNode.Tag(1)
            qry.StudyUID = oneNode.Tag(2)
            qry.SeriesUID = oneNode.Tag(3)
            qry.InstanceUID = oneNode.Tag(4)
            qry.Level = "IMAGE"
    End Select
    
'    If Me.LvwImageList.ListItems.Count < 1 Then Exit Sub
    
    '��ȡͼ��
    
'    With mObjDicomQuery
'        .PatientID = Me.LvwImageList.SelectedItem.SubItems(1)
'        .StudyUID = Me.LvwImageList.SelectedItem.SubItems(4)
'        .SeriesUID = ""
'        .InstanceUID = ""
'        .Level = "STUDY"
'    End With
    
    On Error GoTo GetImageError
    
    zl9comlib.zlCommFun.ShowFlash "���Ե����ڶ�ȡͼ��....", Me
    
    '���Qry�Ĳ���
    qry.Node = mObjDicomQuery.Node
    qry.Port = mObjDicomQuery.Port
    qry.CalledAE = mObjDicomQuery.CalledAE
    qry.CallingAE = mObjDicomQuery.CallingAE
    qry.Root = mObjDicomQuery.Root
    
    
    If optRetrieveType(1).value = True Then     'ͬ��MOVE
        qry.Destination = TxtLocalAE.Text
        qry.ReceivePort = TxtPort.Text
        
        Viewer.Unlisten TxtPort.Text
        Set dicGetImages = qry.GetUsingMove
        Do
            OK = Viewer.Listen(TxtPort.Text)
        Loop While Not OK
    Else
        qry.Destination = TxtLocalAE.Text
        qry.MoveImages
    End If
    
'    '��ȡͼ��
'    Set dicGetImages = mObjDicomQuery.GetImages
    
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

    If MsgBoxD(Me, "��ȡ" & mstrDeviceName & "�豸��ͼ�񲻳ɹ����Ƿ����ԣ�" & vbCrLf & err.Description, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdGetImageInfo_Click()
    Dim dicGetDates As New DicomDataSets
    Dim dicGetDate As DicomDataSet
    Dim objItem As ListItem
    Dim strTmp As String
'    Dim g As New DicomGlobal
    
    
'    g.RegWord("Timeout") = 120
    
    On Error GoTo GetImageInfoErr
            
    Set dicGetDates = funQueryData
    
    '���TreeView�б�
    Call subFillTrvList(dicGetDates)
    
    If trvList.Nodes.Count = 0 Then
        cmdDownImage.Enabled = False
    Else
        cmdDownImage.Enabled = True
    End If
            
    Exit Sub
GetImageInfoErr:
   
End Sub


Private Sub DTPBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub DTPEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    
    Me.TxtIP = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������IP", "localHost")
    Me.TxtPort = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�˿ں�", "104")
    Me.TxtLocalAE = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����AE", "ZLSoftPACS")
    Me.TxtRemoteAE = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Զ��AE", "XX_SUP")
    
    '�Ͳ�ѯ���˵���������һ��
    curDate = zlDatabase.Currentdate
    
    If frmPACSFilter.mBeforeDays <= 0 Then frmPACSFilter.mBeforeDays = 3
    
    dtpEnd.MaxDate = curDate: dtpBegin.MaxDate = curDate
    dtpBegin.value = Format(curDate - frmPACSFilter.mBeforeDays, "yyyy-MM-dd")
    dtpEnd.value = Format(curDate, "yyyy-MM-dd")
    
End Sub

Private Sub LvwImageList_DblClick()
    cmdDownImage_Click
End Sub

Private Sub trvList_Expand(ByVal Node As MSComctlLib.Node)
    Dim oneNode As Node
    Dim dsData As DicomDataSet
    Dim IDs(4) As String
    Dim i As Integer
    Dim dssReturn As DicomDataSets
    Dim dsReturn As DicomDataSet
    Dim strDesp As String
    
    On Error GoTo err
    
    '��ȡTAG������
    For i = 1 To Node.Tag(0)
        IDs(i) = Node.Tag(i)
    Next
    
    IDs(0) = Node.Tag(0) + 1
    Node.Tag(0) = -Node.Tag(0)
    
    Select Case Node.Tag(0)
    Case 1  'ԭLevelΪ����
        mObjDicomQuery.Level = "STUDY"
        mObjDicomQuery.PatientID = Node.Tag(1)
        mObjDicomQuery.SeriesUID = ""
        mObjDicomQuery.InstanceUID = ""
        
        Set dssReturn = mObjDicomQuery.DoQuery
        
        trvList.Nodes.Remove Node.Child.Index
        
        For Each dsReturn In dssReturn
    
            strDesp = "  ���������" & getDescription(dsReturn) & "  AccNum��" & dsReturn.AccessionNumber
            
            Set oneNode = trvList.Nodes.Add(Node.Index, tvwChild, , strDesp)
            IDs(2) = dsReturn.StudyUID
            oneNode.Tag = IDs
            trvList.Nodes.Add oneNode.Index, tvwChild, , "���Ժ����ڲ�ѯ�����б�..."
            oneNode.Expanded = False
            
            '��¼��־
            Call subLogDataset(dsReturn, "trvList_Expand", "��ѯ�����LevelΪ���")
        Next
        
    Case 2  'ԭLevelΪ���
        mObjDicomQuery.Level = "SERIES"
        mObjDicomQuery.PatientID = Node.Tag(1)
        mObjDicomQuery.StudyUID = Node.Tag(2)
        mObjDicomQuery.SeriesUID = ""
        mObjDicomQuery.InstanceUID = ""
        
        Set dssReturn = mObjDicomQuery.DoQuery
        
        trvList.Nodes.Remove Node.Child.Index
        
        For Each dsReturn In dssReturn
            If dsReturn.Attributes(&H8, &H60).Exists Then
                If Not IsNull(dsReturn.Attributes(&H8, &H60).value) Then
                    strDesp = "Ӱ�����" & dsReturn.Attributes(&H8, &H60).value
                End If
            End If
            
            If dsReturn.Attributes(&H20, &H11).Exists Then
                If Not IsNull(dsReturn.Attributes(&H20, &H11).value) Then
                    strDesp = strDesp & "  ���кţ�" & dsReturn.Attributes(&H20, &H11).value
                End If
            End If
            
            strDesp = strDesp & "  ����������" & dsReturn.SeriesDescription
            
            If strDesp = "" Then strDesp = "**����**"
            Set oneNode = trvList.Nodes.Add(Node.Index, tvwChild, , strDesp)
            IDs(3) = dsReturn.SeriesUID
            oneNode.Tag = IDs
            trvList.Nodes.Add oneNode.Index, tvwChild, , "���Ժ����ڲ�ѯͼ���б�..."
            oneNode.Expanded = False
            
            '��¼��־
            Call subLogDataset(dsReturn, "trvList_Expand", "��ѯ�����LevelΪ����")
        Next
    
    Case 3  'ԭLevelΪ����
        mObjDicomQuery.Level = "IMAGE"
        mObjDicomQuery.PatientID = Node.Tag(1)
        mObjDicomQuery.StudyUID = Node.Tag(2)
        mObjDicomQuery.SeriesUID = Node.Tag(3)
        mObjDicomQuery.InstanceUID = ""
        
        Set dssReturn = mObjDicomQuery.DoQuery
        
        trvList.Nodes.Remove Node.Child.Index
        
        For Each dsReturn In dssReturn
            strDesp = dsReturn.InstanceUID
            Set oneNode = trvList.Nodes.Add(Node.Index, tvwChild, , strDesp)
            IDs(4) = dsReturn.InstanceUID
            oneNode.Tag = IDs
            '��¼��־
            Call subLogDataset(dsReturn, "trvList_Expand", "��ѯ�����LevelΪͼ��")
        Next
    Case Else
        MsgBox "����ʶ��Ĳ�ѯ����"
    End Select
    
    Exit Sub
err:
    '�ݲ�����
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
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
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

Private Sub TxtStudyUID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub subFillTrvList(dssRtruen As DicomDataSets)
'------------------------------------------------
'���ܣ����ݲ�ѯ��������TreeView
'������ dssRtruen  --  ���ص����ݼ�
'���أ���
'------------------------------------------------
    Dim dsReturn As DicomDataSet
    Dim IDs(4) As String
    Dim oneNode As Node
    
    trvList.Nodes.Clear
    
    For Each dsReturn In dssRtruen
        '��¼����ID
        IDs(1) = dsReturn.PatientID
        
        '���ݲ�ѯ��Rootȷ�����ݵ�װ��
        If optQueryRoot(1).value = True Then
            '������Բ���Ϊ������ѯ
            Set oneNode = trvList.Nodes.Add(, , , "������" & dsReturn.Name & "  ����ID��" & dsReturn.PatientID)
            IDs(0) = 1
            trvList.Nodes.Add oneNode.Index, tvwChild, , "��ȴ������ڲ�ѯ����б�..."
        Else
            '������Լ��Ϊ������ѯ
            Set oneNode = trvList.Nodes.Add(, , , "������" & dsReturn.Name & "  ���������" & getDescription(dsReturn) & "  AccNum��" & dsReturn.AccessionNumber)
            IDs(0) = 2
            IDs(2) = dsReturn.StudyUID
            trvList.Nodes.Add oneNode.Index, tvwChild, , "��ȴ������ڲ�ѯ�����б�..."
        End If
        oneNode.Tag = IDs
        oneNode.Expanded = False
    Next
    
    
'    For Each dicGetDate In dicGetDates
'        Set objItem = Me.LvwImageList.FindItem(dicGetDate.StudyUID)
'        If objItem Is Nothing Then
'            Set objItem = Me.LvwImageList.ListItems.Add(, "_" & dicGetDate.StudyUID, dicGetDate.Name)
'            objItem.SubItems(1) = Nvl(dicGetDate.PatientID)
'            strTmp = Nvl(dicGetDate.Sex)
'            strTmp = Replace(strTmp, "M", "��")
'            strTmp = Replace(strTmp, "F", "Ů")
'            objItem.SubItems(2) = strTmp
'            objItem.SubItems(3) = IIf(dicGetDate.Attributes(&H8, &H20).Exists, Nvl(dicGetDate.Attributes(&H8, &H20)), "")
'            objItem.SubItems(4) = Nvl(dicGetDate.StudyUID)
'        End If
'    Next
'
'
'
'    Dim nd As Node
'    For Each R In res
'        IDs(1) = R.PatientID
'        If OptionsBox.PatientRoot Then
'            Set nd = Tree.Nodes.Add(, , , R.Name)
'            IDs(0) = 1
'            Tree.Nodes.Add nd.Index, tvwChild, , "Please wait, while the study list is retrieved"
'        Else
'            Set nd = Tree.Nodes.Add(, , , R.Name & " / " & Description(R))
'            IDs(0) = 2
'            IDs(2) = R.StudyUID
'            Tree.Nodes.Add nd.Index, tvwChild, , "Please wait, while the series list is retrieved"
'        End If
'        nd.Tag = IDs
'        nd.Expanded = False
'
'    Next
End Sub


Private Function getDescription(study As DicomDataSet) As String
'------------------------------------------------
'���ܣ���ȡ�������
'������ study  --  ������ݼ�
'���أ��������
'------------------------------------------------
    Dim strDesp As String
    Dim attr As DicomAttribute
    
    On Error Resume Next
    
    strDesp = study.StudyDescription
    If strDesp = "" Then strDesp = "��飺 "
    Set attr = study.Attributes(8, &H20)
    If attr.Exists Then strDesp = strDesp & attr.value
    getDescription = strDesp
End Function

Private Function funQueryData() As DicomDataSets
'------------------------------------------------
'���ܣ���ѯ������C-FIND��Ϣ
'������ study  --  ������ݼ�
'���أ���ѯ���ص����ݼ�
'------------------------------------------------
    Dim strLevel As String
    Dim strTmp As String
    Dim dsReturn As DicomDataSet
    Dim dsQuery As DicomDataSet
    Dim dssSub As DicomDataSets
    Dim dsSub As DicomDataSet
    Dim strTime As String
    
    On Error GoTo err
    
    If chkLevelStudy.value = 1 Then
        strLevel = "STUDY"
    Else
        strLevel = "SERIES"
    End If
    
    If optQueryRoot(1).value = True Then
        mObjDicomQuery.Root = "PATIENT"
    Else
        mObjDicomQuery.Root = "STUDY"
    End If
    
    If strLevel = "STUDY" Then
    
        Set dsQuery = funCreateDSS(strLevel)
        
        '������ں�ʱ��
        If dtpBegin.value <> "" Or dtpEnd.value <> "" Then
            If dtpBegin.value <> "" Then
                strTmp = Format(dtpBegin, "yyyymmdd")
                strTime = Format(dtpBegin, "HHMMSS")
            End If
            strTmp = strTmp & "-"
            strTime = strTime & "-"
            
            If dtpBegin.value <> "" Then
                strTmp = strTmp & Format(dtpEnd, "yyyymmdd")
                strTime = strTime & Format(dtpEnd, "HHMMSS")
            End If
            
            dsQuery.Attributes.Add &H8, &H20, strTmp
            
            'ʱ��
            If strTime <> "000000-000000" And strTime <> "-000000" And strTime <> "000000-" Then
                dsQuery.Attributes.Add &H8, &H30, strTime
            End If
        End If
        
        '��������
        If txtName.Text <> "" Then
            dsQuery.Attributes.Add &H10, &H10, Trim(txtName.Text) & "*"
        End If
        
        '�����Ա�
        strTmp = Me.cboSex
        strTmp = Replace(strTmp, "��", "M")
        strTmp = Replace(strTmp, "Ů", "F")
        dsQuery.Sex = strTmp
        
        'AccessionNumber
        If txtAccessionNumber.Text <> "" Then
            dsQuery.Attributes.Add &H8, &H50, txtAccessionNumber.Text
        End If
        
        '����
        If txtStudyID.Text <> "" Then
            dsQuery.Attributes.Add &H20, &H10, txtStudyID.Text
        End If '
        
        '����е�Ӱ�����
        If txtModalitiesInStudy.Text <> "" Then
            dsQuery.Attributes.Add &H8, &H61, txtModalitiesInStudy.Text
        End If
        
        '����ҽ��
        If txtReferringDoctor.Text <> "" Then
            dsQuery.Attributes.Add &H8, &H90, txtReferringDoctor.Text
        End If
                    
    ElseIf strLevel = "SERIES" Then
        Set dsQuery = funCreateDSS(strLevel)
        
        'Ӱ�����
        If txtModality.Text <> "" Then
            dsQuery.Attributes.Add &H8, &H60, txtModality.Text
        End If
        
        'Request Procedure ID�� Scheduled procedure step ID
        If txtRProcedureID.Text <> "" Or txtSPStepID.Text <> "" Then
            Set dssSub = New DicomDataSets
            Set dsSub = New DicomDataSet
            dsSub.Attributes.Add &H40, &H1001, txtRProcedureID.Text
            dsSub.Attributes.Add &H40, &H9, txtSPStepID.Text
            Call dssSub.Add(dsSub)
            dsQuery.Attributes.Add &H40, &H275, dssSub
        End If
        
        '���к�
        If txtSeriesNumber.Text <> "" Then
            dsQuery.Attributes.Add &H20, &H11, txtSeriesNumber.Text
        End If
        
        'PPS��ʼʱ��
        If dtpPPSStart.value <> "" Or dtpPPSEnd.value <> "" Then
            If dtpPPSStart.value <> "" Then
                strTmp = Format(dtpPPSStart, "yyyymmdd")
                strTime = Format(dtpPPSStart, "HHMMSS")
            End If
            strTmp = strTmp & "-"
            strTime = strTime & "-"
            
            If dtpPPSEnd.value <> "" Then
                strTmp = strTmp & Format(dtpPPSEnd, "yyyymmdd")
                strTime = strTime & Format(dtpPPSEnd, "HHMMSS")
            End If
            dsQuery.Attributes.Add &H40, &H244, strTmp
            
            'ʱ��
            If strTime <> "000000-000000" And strTime <> "-000000" And strTime <> "000000-" Then
                dsQuery.Attributes.Add &H40, &H245, strTime
            End If
        End If
        
    ElseIf strLevel = "IMAGE" Then
    
    
    End If
    
    
    '��¼��־
    Call subLogDataset(dsQuery, "funQueryData", "��ѯ�����ݼ�")
    
    Set funQueryData = mObjDicomQuery.DoRawQuery(dsQuery)
    
    '��¼��־
    For Each dsReturn In funQueryData
        Call subLogDataset(dsReturn, "funQueryData", "��ѯ���ؽ��")
    Next
    
    Exit Function
    
err:
     If err.Number = 1011 Then
        If MsgBoxD(Me, "����" & mstrDeviceName & "�豸���ɹ����Ƿ����ԣ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Resume
        End If
    Else
        If err.Number = 1049 Then
            err.Description = "���ӱ��ܾ������鱾��AE��Զ��AE�Ƿ���ȷ��" & vbCrLf & err.Description
        End If
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    
End Function


Private Sub subLogDataset(ds As DicomDataSet, logSubName As String, logTitle As String)
'------------------------------------------------
'���ܣ���¼���ݼ���־
'������ ds  --  ���ݼ�
'       logSubName
'       logTitle
'���أ���
'------------------------------------------------
    Dim strLog As String
    
    If chkLog.value = 1 Then
        AppendAttributes strLog, "", ds.Attributes
        WriteCommLog logSubName, logTitle, Replace(strLog, "'", "��")
    End If
    
End Sub

Private Sub AppendAttributes(ByRef list As String, prefix As String, ByRef ob As Object)
    Dim at As DicomAttribute
    Dim s As DicomDataSets
    Dim i As Integer
    Dim v As Variant
    For Each at In ob
        list = list & prefix & "(" & hex4(at.Group) & "," & hex4(at.Element) & ") : "
        list = list & Left(at.Description & Space(30), 30) & ": "
        If (at.Group = &H7FE0) Then ' pixel data
            list = list & "Pixel data" & vbCrLf
        ElseIf (VarType(at.value) = 9) Then ' i.e. a sequence
            Set s = at.value
            list = list & "Sequence of " & s.Count & " items:" & vbCrLf
            For i = 1 To s.Count
                list = list & prefix & ">---------------" & vbCrLf
                AppendAttributes list, prefix & ">", s(i).Attributes
            Next
            list = list & prefix & ">---------------" & vbCrLf
        Else
            v = at.value ' could be variant or array
            If (VarType(v) > 8192) Then ' i.e. an array
                list = list & "Multiple values :" & vbCrLf & "              "
                If UBound(v, 1) > 32 Then
                    list = list & "Array of " & UBound(v, 1) & " elements"
                Else
                    For i = LBound(v, 1) To UBound(v, 1)
                        list = list & v(i)
                        If i <> UBound(v, 1) Then list = list & " : "
                    Next
                End If
                list = list & vbCrLf
            Else
                list = list & v & vbCrLf
            End If
        End If
    Next
End Sub


Private Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String)
'------------------------------------------------
'���ܣ���¼ͨѶ��־
'������ logSubName  --  ������־�ĺ�����
'       logTitle   -- ��־����
'       logDesc   --  ��־����
'���أ���
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    Dim intHour As Integer
    
    On Error GoTo err
    
    intHour = Hour(time)
    intHour = intHour / 2
    intHour = intHour * 2
    
    strFileName = App.Path & "\" & date & "-" & intHour & ".log"
    
    If chkLog.value = 1 Then
        strLog = Now() & " ���⣺ " & logTitle & vbCrLf & "      ������ " & logSubName & vbCrLf & "     ��־���ݣ�" & logDesc & vbCrLf
        
        Open strFileName For Append As #1
        Print #1, strLog
        Close #1
    
    End If
    Exit Sub
err:
    Close #1
End Sub

Private Function hex4(ByVal v As Integer) As String
    hex4 = Right("000" & Hex(v), 4)
End Function


Private Function funCreateDSS(strLevel As String) As DicomDataSet
'------------------------------------------------
'���ܣ�����һ���յĲ�ѯ���ݼ�
'������ strLevel  -- ��ѯ�ļ���
'���أ��յĲ�ѯ���ݼ�
'------------------------------------------------
    Dim dsReturn As DicomDataSet
    Dim dssSub As DicomDataSets
    Dim dsSub As DicomDataSet
    
    On Error GoTo err
    
    Set dsReturn = New DicomDataSet
    
    If strLevel = "STUDY" Then
        dsReturn.Attributes.Add &H8, &H52, "STUDY"  'Level
        dsReturn.Attributes.Add &H8, &H20, ""   'study date
        dsReturn.Attributes.Add &H8, &H30, ""   'study time
        dsReturn.Attributes.Add &H8, &H50, ""   'accession number
        dsReturn.Attributes.Add &H10, &H10, ""  'patient name
        dsReturn.Attributes.Add &H10, &H20, ""  'patient id
        dsReturn.Attributes.Add &H20, &H10, ""  'study id
        dsReturn.Attributes.Add &H20, &HD, ""   'study instance uid
        dsReturn.Attributes.Add &H8, &H61, ""   'modalities in study
        dsReturn.Attributes.Add &H8, &H90, ""   'referring physican's name
        
        dsReturn.Attributes.Add &H10, &H30, ""  'birthday
        dsReturn.Attributes.Add &H10, &H40, ""  'sex
        dsReturn.Attributes.Add &H20, &H1206, ""    'number of  study related series
        dsReturn.Attributes.Add &H20, &H1208, ""    'number of stusy related instances
        
    ElseIf strLevel = "SERIES" Then
        dsReturn.Attributes.Add &H8, &H52, "SERIES"  'Level
        dsReturn.Attributes.Add &H8, &H60, ""   'modality
        dsReturn.Attributes.Add &H20, &H11, ""  'series number
        dsReturn.Attributes.Add &H20, &HE, ""   'series instance uid
        dsReturn.Attributes.Add &H20, &H1209, ""    'number of series related instances
        dsReturn.Attributes.Add &H8, &H103E, "" 'series description
        
        
        Set dssSub = New DicomDataSets
        Set dsSub = New DicomDataSet
        dsSub.Attributes.Add &H40, &H1001, ""   'requested procedure id
        dsSub.Attributes.Add &H40, &H9, ""      'scheduled procedure step id
        Call dssSub.Add(dsSub)
        dsReturn.Attributes.Add &H40, &H275, dssSub 'request attributes sequence
        
        dsReturn.Attributes.Add &H40, &H244, "" 'performed procedure step start date
        dsReturn.Attributes.Add &H40, &H245, "" 'performed procesure step start time
        
    Else
        'ͼ�񼶱�
        dsReturn.Attributes.Add &H8, &H52, "IMAGE"  'Level
        dsReturn.Attributes.Add &H20, &H13, ""  'instance number
        dsReturn.Attributes.Add &H8, &H18, ""   'SOP instance uid
        dsReturn.Attributes.Add &H8, &H16, ""   'SOP class uid
        
        'ͼ�񼶱����ӵ�
        dsReturn.Attributes.Add &H28, &H10, ""  'Rows
        dsReturn.Attributes.Add &H28, &H11, ""  'Columns
        dsReturn.Attributes.Add &H28, &H100, "" 'Bits allocated
        dsReturn.Attributes.Add &H28, &H8, ""   'number of frames
    End If
    
    
    Set funCreateDSS = dsReturn
    Exit Function
err:
    funCreateDSS = Nothing
    
End Function

