VERSION 5.00
Object = "{82809FC2-3B17-4941-8A37-713AA0519BB1}#1.0#0"; "DVDProX2.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreateCD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����CD"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   Icon            =   "frmCreateCD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame4 
      Caption         =   "��¼��Ϣ"
      Height          =   885
      Left            =   120
      TabIndex        =   21
      Top             =   6150
      Width           =   9255
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   60
         TabIndex        =   22
         Top             =   510
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label LabInformation 
         AutoSize        =   -1  'True
         Caption         =   "��Ϣ:"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5490
      TabIndex        =   6
      Top             =   7230
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "д��(&W)"
      Height          =   350
      Left            =   2370
      TabIndex        =   5
      Top             =   7230
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "CD ���ѡ��:"
      Height          =   1815
      Left            =   4920
      TabIndex        =   14
      Top             =   4290
      Width           =   4455
      Begin VB.OptionButton optPacking 
         Caption         =   "DICOMDIR�Ͷ���CD��Ƭվ"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optPacking 
         Caption         =   "ֻ��DICOMDIR"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��¼ѡ��:"
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   4290
      Width           =   4575
      Begin VB.CommandButton CmdWriterCDOption 
         Caption         =   "�߼�ѡ��"
         Height          =   345
         Left            =   3360
         TabIndex        =   20
         Top             =   1380
         Width           =   975
      End
      Begin VB.ComboBox CboWriterSpeeds 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   630
         Width           =   3195
      End
      Begin VB.ComboBox CboDrivers 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   270
         Width           =   3195
      End
      Begin VB.TextBox txtCDName 
         Height          =   300
         Left            =   1170
         TabIndex        =   2
         Top             =   990
         Width           =   3165
      End
      Begin DVDPROX2LibCtl.DVDWriterPro2 DVDWriterPro 
         Left            =   240
         OleObjectBlob   =   "frmCreateCD.frx":000C
         Top             =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�ٶ�:"
         Height          =   180
         Left            =   630
         TabIndex        =   18
         Top             =   690
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CD����:"
         Height          =   180
         Left            =   450
         TabIndex        =   16
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��¼����:"
         Height          =   180
         Left            =   270
         TabIndex        =   15
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DICOM CD ���ݣ�"
      Height          =   4125
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtSpacing 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   7920
         TabIndex        =   11
         Text            =   "600.00 MB"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtSpacing 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   7920
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "ȫ��ɾ��"
         Height          =   350
         Left            =   7920
         TabIndex        =   8
         Top             =   1080
         Width           =   1100
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "ɾ��"
         Height          =   350
         Left            =   7920
         TabIndex        =   7
         Top             =   480
         Width           =   1100
      End
      Begin MSComctlLib.TreeView trvCDContents 
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "���������"
         Height          =   255
         Left            =   7920
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "��Ҫ�Ŀռ䣺"
         Height          =   255
         Left            =   7920
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCreateCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As frmViewer
Dim dimgsInCD As DicomImages
Private Const DBL_PACSLITE_SIZE = 20
Private Const INT_MAX_CAPACITY = 600
Private Const STR_ATTACHED_FILE_PATH = "PACSLIST"



Private Sub CboDrivers_Click()
    Dim lngDriveIndex As Long

    lngDriveIndex = CboDrivers.ItemData(CboDrivers.ListIndex)

    If DVDWriterPro.OpenDrive(lngDriveIndex) = False Then
        MsgBox "���ܴ�ѡ��Ŀ�¼����!", vbInformation, gstrSysName
        Exit Sub
    End If

    LoadWriteSpeedCombo

End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim img As DicomImage
    Dim strMiddlePath  As String
    Dim strDicomPath As String
    Dim strRootPath As String
    Dim transfersyntax As String
    Dim dsetDicomDir As New DicomDataSet
    Dim strAppPath As String
    Dim strFileName As String
    Dim fs As Object
    Dim strMsg As String
    
    Dim intUseJoliet As Integer
    Dim intCDRWMode  As Integer
    Dim blHighCompatibilityMode  As Boolean
    Dim blCheckImage As Boolean
    Dim blCloseDisk As Boolean
    Dim blTestWriter As Boolean
    Dim blBufferProof As Boolean
    Dim blAutoVerify As Boolean
    
    intUseJoliet = GetSetting("ZLSOFT", "˽��ģ��\����\" & App.ProductName & "\��¼����", "ʹ��Joliet", 1)
    intCDRWMode = GetSetting("ZLSOFT", "˽��ģ��\����\" & App.ProductName & "\��¼����", "ʹ��CDRWģʽ", 1)
    blHighCompatibilityMode = GetSetting("ZLSOFT", "˽��ģ��\����\" & App.ProductName & "\��¼����", "�߼���DVDģʽ", 0)
    blCheckImage = GetSetting("ZLSOFT", "˽��ģ��\����\" & App.ProductName & "\��¼����", "��ʹ�ø��ٻ���", 0)
    blCloseDisk = GetSetting("ZLSOFT", "˽��ģ��\����\" & App.ProductName & "\��¼����", "�رչ���", 1)
    blTestWriter = GetSetting("ZLSOFT", "˽��ģ��\����\" & App.ProductName & "\��¼����", "����д��", 1)
    blBufferProof = GetSetting("ZLSOFT", "˽��ģ��\����\" & App.ProductName & "\��¼����", "����У��", 1)
    blAutoVerify = GetSetting("ZLSOFT", "˽��ģ��\����\" & App.ProductName & "\��¼����", "�Զ�����У��", 1)
    
    On Error GoTo errh
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    '��鴴��CD�������Ƿ�����
    If dimgsInCD.Count <= 0 Then
        MsgBox "û��ͼ���޷�����CD,��ӹ�Ƭվ����ѡ��ͼ��", vbInformation, gstrSysName
        Exit Sub
    End If
'    If Me.TxtPath = "" Then
'        MsgBox "������CD��·����", , "����CD"
'        Me.TxtPath.SetFocus
'        Exit Sub
'    End If
    If Me.txtCDName = "" Then
        MsgBox "������CD���ơ�", vbInformation, gstrSysName
        Me.txtCDName.SetFocus
        Exit Sub
    End If
'    If left(Me.txtSpacing(0).Text, InStr(Me.txtSpacing(0).Text, "MB") - 1) - INT_MAX_CAPACITY > 0 Then
'        MsgBox "ѡ���ͼ��ռ�ÿռ䳬��һ�Ź��̵�����������ɾ������ͼ��", , "����CD"
'        Exit Sub
'    End If
    '����ͼ�񣬲�����DICOMDIR
'    If Dir(Me.TxtPath, vbDirectory) = "" Then
'        MsgBox "������Ч��·��������ѡ��!", vbInformation, "��ʾ"
'        Me.TxtPath.SetFocus
'        Exit Sub
'    End If
    strRootPath = IIf(Len(App.Path) > 3, App.Path & "\CreateCDTmp", App.Path & "CreateCDTmp")
    strDicomPath = strRootPath & "\DICOM"
    If Dir(strRootPath, vbDirectory) = "" Then MkDir (strRootPath)
    If Dir(strDicomPath, vbDirectory) = "" Then MkDir (strDicomPath)
    transfersyntax = "1.2.840.10008.1.2.1"
    For Each img In dimgsInCD
        subSaveLabelToImg img
        '���Ŀ¼�����ڣ��򴴽�Ŀ¼
        strMiddlePath = "IMAGES"
        If Dir(strDicomPath & "\" & strMiddlePath, vbDirectory) = "" Then
            MkDir (strDicomPath & "\" & strMiddlePath)
        End If
        strMiddlePath = strMiddlePath & "\" & ChkDir(img.Name)
        If Dir(strDicomPath & "\" & strMiddlePath, vbDirectory) = "" Then
            MkDir (strDicomPath & "\" & strMiddlePath)
        End If
        strMiddlePath = strMiddlePath & "\" & img.StudyUID
        If Dir(strDicomPath & "\" & strMiddlePath, vbDirectory) = "" Then
            MkDir (strDicomPath & "\" & strMiddlePath)
        End If
        img.WriteFile strDicomPath & "\" & strMiddlePath & "\" & img.InstanceUID & ".DCM", True, transfersyntax
        dsetDicomDir.AddToDirectory img, strMiddlePath & "\" & img.InstanceUID & ".DCM", transfersyntax, 0
    Next img
    dsetDicomDir.Name = "ZLPACS"
    dsetDicomDir.WriteDirectory strDicomPath & "\DICOMDIR"
    
    '���桰����CD��Ƭվ��
    If Me.optPacking(1).Value = True Then
        '���ض�Ŀ¼���ļ����Ƶ�strPath\PACSLiteĿ¼��
        strAppPath = App.Path & IIf(Len(App.Path) > 3, "\", "") & STR_ATTACHED_FILE_PATH
        If Dir(strAppPath, vbDirectory) <> "" Then
            If Dir(strAppPath, vbDirectory) = "" Then
                MsgBox "û���ҵ������ļ�·����", vbInformation, gstrSysName
                Exit Sub
            End If
            fs.CopyFile strAppPath & "\*.*", strRootPath
        End If
    End If
    '��¼
    If DVDWriterPro.GetMediaType() = mtNotLoaded Then
        MsgBox "������д��Ĺ���!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call DVDWriterPro.CloneDirectoryToISO("\", strRootPath & "\*.*")
    
    
    If (DVDWriterPro.GetDiscFreeSpaceBlocks() < (DVDWriterPro.GetISOVolumeSizeBlocks())) Then
        strMsg = "���ɿռ�: " & ConvertBytesToMBString(DVDWriterPro.ConvertBlocksToBytes(DVDWriterPro.GetDiscFreeSpaceBlocks(), wtpDataMode1)) & " MB!" & _
                vbCrLf & "С����ʹ�ÿռ�: " & ConvertBytesToMBString(DVDWriterPro.ConvertBlocksToBytes(DVDWriterPro.GetISOVolumeSizeBlocks(), wtpDataMode1)) & " MB ."
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    
    With DVDWriterPro

        .AutoVerify = blAutoVerify
        .CloseDisc = blCloseDisk
        .CloseSession = True

        .VolumeType = intUseJoliet
        .VolumeIdentifier = txtCDName.Text 'Were setting only the Volume Identifier..You could set all the volume descriptors however

        .CacheImage = blCheckImage

        .SetBufferProtection blBufferProof

        If (.GetMediaType() = mtCD) Or (.GetMediaType() = mtCDRW) Then
            .DVDHighCompatibilityMode = blHighCompatibilityMode
            .WriteType = intCDRWMode
            .TestWrite = blTestWriter
        Else
            .DVDHighCompatibilityMode = blHighCompatibilityMode
            .WriteType = intCDRWMode
            .TestWrite = False
        End If
    End With

    If DVDWriterPro.WriteDisc() = False Then
        MsgBox "����д�����!", vbCritical, gstrSysName
        Exit Sub
    End If
    
    fs.DeleteFolder strRootPath, True
    '���������Ϣ��ע���

    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ʽ", IIf(optPacking(0).Value, -1, 0))

    Exit Sub
errh:
    MsgBox "��������:" & err.Description & vbCrLf & "�����:" & err.Number, vbExclamation, gstrSysName
End Sub

Private Sub cmdRemove_Click()
    Dim nodeTemp As Node
    Dim nodeUp As Node
    Dim img As DicomImage
    Dim strInstanceUID As String
    
    Set nodeTemp = Me.trvCDContents.SelectedItem
    If nodeTemp Is Nothing Then Exit Sub
    If Not nodeTemp.Child Is Nothing Then Exit Sub
    strInstanceUID = nodeTemp.Tag
    For Each img In dimgsInCD
        If img.InstanceUID = strInstanceUID Then
            dimgsInCD.Remove dimgsInCD.IndexOf(img)
            Set nodeUp = nodeTemp.Parent
            'ɾ��ͼ����Ϣ
            Me.trvCDContents.Nodes.Remove nodeTemp.Index
            '������μ����û��ͼ���򽫼����Ϣɾ��
            If nodeUp.Child Is Nothing Then
                Set nodeTemp = nodeUp.Parent
                Me.trvCDContents.Nodes.Remove nodeUp.Index
                '�����ǰ����û��ͼ���򽫲�����Ϣɾ��
                If nodeTemp.Child Is Nothing Then
                    Me.trvCDContents.Nodes.Remove nodeTemp.Index
                End If
            End If
            Exit For
        End If
    Next img
    subGetImageCapacity
End Sub

Private Sub cmdRemoveAll_Click()
    dimgsInCD.Clear
    Me.trvCDContents.Nodes.Clear
    subGetImageCapacity
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdWriterCDOption_Click()
    frmCreateAdvicedSetup.Show vbModal, Me
End Sub

Private Sub DVDWriterPro_CachingStatus(ByVal nPercentComplete As Integer)
    LabInformation.Caption = "���Cach" & Format(nPercentComplete, "0#") & " %"
End Sub

Private Sub DVDWriterPro_ClosingDisc()
    LabInformation.Caption = "�رչ���......."
End Sub

Private Sub DVDWriterPro_ClosingSession()
    LabInformation.Caption = "�ر���Ϣ...."
End Sub

Private Sub DVDWriterPro_ClosingTrack(ByVal lTrackNumber As Long)
    LabInformation.Caption = "�رչ켣...."
End Sub

Private Sub DVDWriterPro_CreatingDirectories()
    LabInformation.Caption = "����Ŀ¼...."
End Sub

Private Sub DVDWriterPro_FileVerifyComplete(ByVal lFilesCompared As Long, ByVal lFilesMatched As Long)
    LabInformation.Caption = "�Զ�����...(" & lFilesMatched & " �� " & lFilesCompared & " ƥ�����)"
    Call DVDWriterPro.EjectLoad(False)
    MsgBox "д���У�����!", vbInformation, gstrSysName

End Sub

Private Sub DVDWriterPro_FileVerifyStart(ByVal lFilesToCompare As Long)
    LabInformation.Caption = "�Զ�У��... " & lFilesToCompare & " �ļ�."
End Sub

Private Sub DVDWriterPro_FileVerifyStatus(ByVal sItemDestPath As String, ByVal sSourceFilePath As String, ByVal lFileBytesCompared As Long, ByVal lFileSizeTotal As Long, ByVal fvStatus As DVDPROX2LibCtl.eVerifyStatus, ByVal lCurrentFile As Long, ByVal lFilesToCompare As Long, bCancel As Boolean)
    Dim intPercentVerified As Integer

    Select Case fvStatus
        Case fvsComparing
        Case fvsMatched
        Case fvsIncorrectByteCount
            MsgBox "У����� - Դ�ļ���С��Ŀ�겻ƥ��: " & vbCrLf & _
                    sItemDestPath & " (Ŀ��)" & vbCrLf & sSourceFilePath & " (Դ)", vbExclamation, gstrSysName
    
        Case fvsNoMatch
            MsgBox "У����� - Դ�ļ���С��Ŀ�겻ƥ��: " & vbCrLf & _
                    sItemDestPath & " (Ŀ��)" & vbCrLf & sSourceFilePath & " (Դ)", vbExclamation, gstrSysName
    
        Case fvsReadingDiscError
            MsgBox "У����� - Դ�ļ���С��Ŀ�겻ƥ��: " & vbCrLf & _
                    sItemDestPath & " (Ŀ��)" & vbCrLf & sSourceFilePath & " (Դ)", vbExclamation, gstrSysName
    
        Case fvsReadingSourceError
            MsgBox "У����� - Դ�ļ���С��Ŀ�겻ƥ��: " & vbCrLf & _
                    sItemDestPath & " (Ŀ��)" & vbCrLf & sSourceFilePath & " (Դ)", vbExclamation, gstrSysName
    End Select

    LabInformation.Caption = "У��·��: " & sItemDestPath

    intPercentVerified = ((lCurrentFile / lFilesToCompare) * 100)

    ProgressBar1.Value = intPercentVerified

    DoEvents
End Sub

Private Sub DVDWriterPro_PreparingToWrite()

    LabInformation.Caption = "Ԥд��...."

    Me.cmdOK.Enabled = False
End Sub

Private Sub DVDWriterPro_ReadingTrackFile(ByVal sFileName As String, ByVal lFileIndex As Long, ByVal lTrackNumber As Long)
    LabInformation.Caption = "�켣: " & Format(lTrackNumber, "0#") & " - ��ȡ..." & CStr(lFileIndex) & " - " & sFileName
End Sub

Private Sub DVDWriterPro_ReadingTrackFileError(ByVal TrackFileError As DVDPROX2LibCtl.eTrackFileError, ByVal sFileName As String, ByVal lTrackNumber As Long)
    LabInformation.Caption = "���ļ�:" & sFileName & "ʱ��������!"
End Sub

Private Sub DVDWriterPro_ReplaceImportedISOFile(ByVal sDestPath As String, ByVal sNewSourcePath As String, ByVal sFileName As String, bReplaceFile As Boolean)
    Dim lngResult As Long

    lngResult = MsgBox("д���ļ�ʱ�����ļ�����ͬ�Ƿ��滻?", vbOKCancel + vbQuestion, gstrSysName)
    
    If lngResult = vbOK Then
        bReplaceFile = True
    Else
        bReplaceFile = False
    End If
End Sub

Private Sub DVDWriterPro_TrackWriteStatus(ByVal lTrackNumber As Long, ByVal lBlocksWritten As Long, ByVal lBlocksToWrite As Long)
    Dim intPercentTrackWritten As Integer
    On Error Resume Next

    intPercentTrackWritten = ((lBlocksWritten / lBlocksToWrite) * 100)

    ProgressBar1.Value = intPercentTrackWritten
End Sub

Private Sub DVDWriterPro_WriteCancelled()
    LabInformation.Caption = "ȡ��д��......"
    Me.cmdOK.Enabled = True
    MsgBox "д�뱻ȡ��!", vbInformation, gstrSysName
End Sub

Private Sub DVDWriterPro_WriteComplete()
    LabInformation.Caption = "д�����!"

    Me.cmdOK.Enabled = True

    If DVDWriterPro.AutoVerify = False Then
        MsgBox "д�����!", vbInformation, gstrSysName
    End If

    If (DVDWriterPro.TestWrite = False) And (DVDWriterPro.AutoVerify = False) Then
        Call DVDWriterPro.EjectLoad(False)
    End If
End Sub

Private Sub DVDWriterPro_WriteError(ByVal WriteError As DVDPROX2LibCtl.eWriteErrorType, ByVal DriveError As DVDPROX2LibCtl.eCDError, ByVal sErrorInfo As String, ByVal sSenseInfo As String)
    Dim strError As String

    strError = "д��ʱ��������: (" & CStr(WriteError) & ")   " & vbCrLf

    If WriteError = errDriveError Then
        strError = strError & GetDriveErrorMessage(DriveError) & vbCrLf & " ���ʹ�������: " & sSenseInfo
    End If

    MsgBox strError, vbCritical + vbOKOnly, gstrSysName

    Me.cmdOK.Enabled = True

End Sub

Private Sub Form_Load()
    If f Is Nothing Then Exit Sub
    Dim i As Integer
    Dim j As Integer
    Dim v As DicomViewer
    Dim img As DicomImage
    Dim blnAdd As Boolean
    Dim node1 As Node
    Dim node2 As Node
    Dim node3 As Node
    Dim blnInserted As Boolean
    
    '��ȡ��ѡ�е�ͼ��
    Set dimgsInCD = New DicomImages
    For Each v In f.Viewer
        For Each img In v.Images
            If img.Tag <> "" Then
                blnAdd = True
                For j = 1 To dimgsInCD.Count
                    If dimgsInCD(j).InstanceUID = img.InstanceUID Then blnAdd = False
                Next j
                If blnAdd = True Then
                    dimgsInCD.Add img
                    subLabelCopyRebuild img, dimgsInCD(dimgsInCD.Count)
                End If
            End If
        Next img
    Next v
    
    '��ͼ�����Ϣ��ӵ�treeview�У��ֳ����㣻
    '��һ�� ���ˣ�text������������ & ������tag��Patient ID
    '�ڶ��� ��飺text������飺�� & ���������tag��Study UID
    '������ ͼ��text����ͼ�񣺡� & ͼ��UID�� tag��Instance UID
    For Each img In dimgsInCD
        blnInserted = False
        '���treCDContents�ĵ�һ��
        If Me.trvCDContents.Nodes.Count > 1 Then
            '���PatientID�Ƿ����ظ���
            Set node1 = Me.trvCDContents.Nodes(1)
            While (Not node1 Is Nothing) And blnInserted = False
                '�ڲ��˲�β���
                If node1.Tag <> img.PatientID Then  '���Ҳ��˲�ε���һ���ڵ�
                    Set node1 = node1.Next
                Else    '���Ҽ����
                    Set node2 = node1.Child
                    While (Not node2 Is Nothing) And blnInserted = False
                        '�ڼ���β���
                        If node2.Tag <> img.StudyUID Then   '���Ҽ���ε���һ���ڵ�
                            Set node2 = node2.Next
                        Else    '����ͼ����
                            Set node3 = node2.Child
                            While (Not node3 Is Nothing) And blnInserted = False
                                If node3.Tag <> img.InstanceUID Then
                                    Set node3 = node3.Next
                                Else
                                    blnInserted = True
                                End If
                            Wend
                            If blnInserted = False Then
                                subAddNodeToContents "IMAGE", node2, img
                                blnInserted = True
                            End If
                        End If
                    Wend
                    If blnInserted = False Then
                        subAddNodeToContents "STUDY", node1, img
                        blnInserted = True
                    End If
                End If
            Wend
            If blnInserted = False Then
                subAddNodeToContents "PATIENT", Nothing, img
                blnInserted = True
            End If
        Else
            '��trvCDContents���������Ϣ
            subAddNodeToContents "PATIENT", Nothing, img
            blnInserted = True
        End If
    Next img
    subGetImageCapacity
    Me.txtSpacing(1).Text = INT_MAX_CAPACITY & "MB"
    'ע��ؼ�
    Me.DVDWriterPro.LicenseCode = "10LBTZY9V42HTZKTKL27S"
    LoadDriveCombo
    
    Me.optPacking(0).Value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ʽ", 0)
    Me.optPacking(1).Value = Not Me.optPacking(0).Value
    
End Sub

Private Sub subAddNodeToContents(strLevel As String, nodeCurrent As Node, img As DicomImage)
    Dim node1 As Node
    Dim node2 As Node
    Dim node3 As Node
    
    If UCase(strLevel) = "PATIENT" Then
        Set node1 = Me.trvCDContents.Nodes.Add(, , , "������" & img.Name)
        node1.Tag = img.PatientID
        Set node2 = Me.trvCDContents.Nodes.Add(node1, tvwChild, , "��飺" & img.StudyDescription)
        node2.Tag = img.StudyUID
        Set node3 = Me.trvCDContents.Nodes.Add(node2, tvwChild, , "ͼ��" & img.InstanceUID)
        node3.Tag = img.InstanceUID
    ElseIf UCase(strLevel) = "STUDY" Then
        Set node2 = Me.trvCDContents.Nodes.Add(nodeCurrent, tvwChild, , "��飺" & img.StudyDescription)
        node2.Tag = img.StudyUID
        Set node3 = Me.trvCDContents.Nodes.Add(node2, tvwChild, , "ͼ��" & img.InstanceUID)
        node3.Tag = img.InstanceUID
    ElseIf UCase(strLevel) = "IMAGE" Then
        Set node3 = Me.trvCDContents.Nodes.Add(nodeCurrent, tvwChild, , "ͼ��" & img.InstanceUID)
        node3.Tag = img.InstanceUID
    End If
End Sub
 
Private Sub subGetImageCapacity()
    '����ֵ��MBΪ��λ
    Dim img As DicomImage
    Dim dblCapacity As Double
    Dim lngRows As Long
    Dim lngCols As Long
    Dim lngBitAllocate As Long
    
    For Each img In dimgsInCD
        lngRows = img.sizey
        lngCols = img.sizex
        lngBitAllocate = img.Attributes(&H28, &H100).Value
        dblCapacity = dblCapacity + lngRows * lngCols * lngBitAllocate * img.FrameCount / 8 / 1024 / 1024
    Next img
    If Me.optPacking(1).Value = True Then
        dblCapacity = dblCapacity + DBL_PACSLITE_SIZE
    End If
    Me.txtSpacing(0).Text = Format(dblCapacity, "0.00") & "MB"
End Sub

Private Sub optPacking_Click(Index As Integer)
    subGetImageCapacity
End Sub

Private Sub optPacking_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub trvCDContents_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCDName_GotFocus()
    Me.txtCDName.SelStart = 0
    Me.txtCDName.SelLength = Len(Me.txtCDName.Text)
End Sub

Private Sub txtCDName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPath_GotFocus()
'    Me.TxtPath.SelStart = 0
'    Me.TxtPath.SelLength = Len(Me.TxtPath.Text)
End Sub

Private Sub txtPath_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub LoadDriveCombo()
    '�õ��ɶ�д��CD
    Dim intDrives As Integer

    If DVDWriterPro.InitDrives(False) = False Then
        MsgBox "���������ܳ�ʹ��!", vbInformation, gstrSysName
    End If
    
    CboDrivers.Clear

    '���ҿɶ�д��CD
    For intDrives = 0 To DVDWriterPro.GetDriveCount() - 1
        If DVDWriterPro.IsDriveWriter(intDrives) = True Then
            CboDrivers.AddItem DVDWriterPro.GetDriveLetter(intDrives) & ": " & DVDWriterPro.GetDriveVendor(intDrives) & " " & DVDWriterPro.GetDriveModel(intDrives)
            CboDrivers.ItemData(CboDrivers.NewIndex) = intDrives
        End If
    Next
    
    '�豸��һ��CD
    If CboDrivers.ListCount > 0 Then
        CboDrivers.ListIndex = 0
    Else
        MsgBox "û���ҵ����Զ�д�Ŀ�¼����!", vbExclamation, gstrSysName
    End If
End Sub

Private Sub LoadWriteSpeedCombo()
    Dim lngMaxWriteSpeedKBS As Long
    Dim lngSpeedKBS As Long
    Dim dblDisplaySpeed  As Double
    Dim bUseDVDspeeds As Boolean
    Dim DiscType As eMediaType
    
    DiscType = DVDWriterPro.GetMediaType()
    If (DiscType = mtCD) Or (DiscType = mtCDRW) Or (DiscType = mtNotLoaded) Then
        bUseDVDspeeds = False
    Else
        bUseDVDspeeds = True
    End If
    
    CboWriterSpeeds.Clear
    
    lngMaxWriteSpeedKBS = DVDWriterPro.GetMaxWriteSpeed()
    
    If lngMaxWriteSpeedKBS > 0 Then
        
        lngSpeedKBS = lngMaxWriteSpeedKBS
        
        If bUseDVDspeeds = True Then
            Do
                dblDisplaySpeed = CDbl(lngSpeedKBS) / 1380
                
                CboWriterSpeeds.AddItem Format(dblDisplaySpeed, "#.0") & "x"

                CboWriterSpeeds.ItemData(CboWriterSpeeds.NewIndex) = lngSpeedKBS
                
                lngSpeedKBS = lngSpeedKBS - 690
            Loop While (lngSpeedKBS >= 1380)
        Else
            Do
                dblDisplaySpeed = CDbl(lngSpeedKBS) / 176
                
                If (dblDisplaySpeed > 0) And (dblDisplaySpeed < 2) Then
                    dblDisplaySpeed = 1
                End If
                
                CboWriterSpeeds.AddItem Format(dblDisplaySpeed, "#") & "x"

                CboWriterSpeeds.ItemData(CboWriterSpeeds.NewIndex) = lngSpeedKBS
                
                If lngSpeedKBS >= 1200 Then
                    lngSpeedKBS = lngSpeedKBS - 704
                Else
                    lngSpeedKBS = lngSpeedKBS - 352
                End If
            Loop While (lngSpeedKBS > 0)
        End If
    Else
        CboWriterSpeeds.AddItem "Ĭ��"
    End If
    If CboWriterSpeeds.ListCount > 0 Then
        CboWriterSpeeds.ListIndex = 0
    End If
End Sub
Private Function ConvertBytesToMBString(ByVal dblBytes As Double) As String
    ConvertBytesToMBString = Format((dblBytes / 1048576), "#########0.#0 MB")
End Function
Private Function GetDriveErrorMessage(ByVal DriveError As DVDPROX2LibCtl.eCDError)
    
    Dim strMsg As String
    
    Select Case DriveError
    Case cdNoAdditionalErrorData '1000
        strMsg = "No additional error data was reported"
    Case cdIOTerminated '1001
        strMsg = "�쳣I/O��ֹ"
    Case cdLogicalUnitNotReady '1002
        strMsg = "���Ϸ��û��׼����"
    Case cdLogicalUnitCommFailed '1003
        strMsg = "����һ��������Ϣ���"
    Case cdDeviceTrackingError '1004
        strMsg = "�������������ɹ켣"
    Case cdWriteGenericError '1005
        strMsg = "������д����"
    Case cdWriteRecoveryNeeded '1006
        strMsg = "Writing occurred, but recovery is needed"
    Case cdWriteRecoveryFailed '1007
        strMsg = "��ͼ�ָ�ʧ��"
    Case cdWriteLossOfStreaming '1008
        strMsg = "A buffer under-run has occurred"
    Case cdReadUnrecovered '1009
        strMsg = "����������ǲ��ܶ�"
    Case cdReadRetriesExhausted '1010
        strMsg = "���������ͼ�ָ���ʧ��"
    Case cdReadErrorTooLong '1011
        strMsg = "��ȡ��ʱ"
    Case cdReadLECUncorrectable '1012
        strMsg = "While reading, the LEC was not recovered"
    Case cdReadCIRCUnrecovered '1013
        strMsg = "The CIRC could not be validated"
    Case cdReadUPCEANFailed '1014
        strMsg = "Reading of the UPC failed"
    Case cdReadISRCFailed '1015
        strMsg = "Reading of the ISRC failed"
    Case cdReadLossOfStreaming '1016
        strMsg = "��ȡ����ʱ���ж�"
    Case cdPositioningError '1017
        strMsg = "��������д��ý��"
    Case cdParameterListLengthError '1018
        strMsg = "һ�������ݵĳ��������͵�������"
    Case cdSynchronousTransferError '1019
        strMsg = "����������Ϸ�һ��Ǩ�ƴ���"
    Case cdInvalidCommandCode '1020
        strMsg = "һ��ʧЧ������͵������������"
    Case cdLBAOutOfRange '1021
        strMsg = "Error trying to write past the end of the media"
    Case cdInvalidCDBField '1022
        strMsg = "ʧЧ����ʧ��"
    Case cdInvalidParamterListField '1023
        strMsg = "һ�������ݵĲ������͵�������"
    Case cdParameterNotSupported '1024
        strMsg = "��֧��һ���������"
    Case cdParamterValueInvalid '1025
        strMsg = "һ���������ʧУ��ֵ"
    Case cdBusOrDeviceReset '1026
        strMsg = "The SCSI/ATAPI bus was reset and caused a write failure"
    Case cdParametersChanged '1027
        strMsg = "A command parameter changed while in progress"
    Case cdIncompatibleMedium '1028
        strMsg = "������̲��ܼ����������"
    Case cdReadUnknownMediumFormat '1029
        strMsg = "����������ܼ������ָ�ʽ�Ĺ���"
    Case cdReadIncompatibleMediumFormat '1030
        strMsg = "������̲��ܼ��ݵ�ǰ����"
    Case cdWriteUnknownMediumFormat '1031
        strMsg = "���̸�ʽδ֪"
    Case cdIncompatibleWriteFormat '1032
        strMsg = "�����������д������Ϊ��ʽì��"
    Case cdMediaNotPresent '1033
        strMsg = "������̲�������"
    Case cdLogicalUnitFailure '1034
        strMsg = "The drive had an unknown failure"
    Case cdLogicalUnitTimedOut '1035
        strMsg = "The drive has timed out while completing a command"
    Case cdEraseFailed '1036
        strMsg = "The disc could not be erased"
    Case cdUnableToRecoverTOC '1037
        strMsg = "The Table of Contents is unrecoverable"
    Case cdEndOfUserAreaOnTrack '1038
        strMsg = "Error trying to write past the user area of the media"
    Case cdPacketDoesNotFit '1039
        strMsg = "Packet recording is not configured correctly"
    Case cdIllegalTrackMode '1040
        strMsg = "The current track mode is incompatible with the disc format"
    Case cdInvalidPacketSize '1041
        strMsg = "Packet recording has incorrect size"
    Case cdSessionFixationError '1042
        strMsg = "A generic session closing error occurred"
    Case cdSessionFixationErrorLeadIn '1043
        strMsg = "Error closing Lead-in area"
    Case cdSessionFixationErrorLeadOut '1044
        strMsg = "Error closing Lead-out area"
    Case cdSessionFixationIncompleteTrack '1045
        strMsg = "While closing, the track was never completed"
    Case cdEmptyPartialReservedTrack '1046
        strMsg = "Error attempting to write to a reserved track"
    Case cdPowerCalibrationFull '1047
        strMsg = "Power calibration area is full"
    Case cdPowerCalibrationAreaError '1048
        strMsg = "A flaw exists in the Power calibration area"
    Case cdPMAUpdateFailure '1049
        strMsg = "The disc's PMA could not be updated"
    Case cdPMAFull '1050
        strMsg = "The disc's PMA is full"
    Case cdUnknownError '1051
        strMsg = "Unknown error - use extended data for more information"
    Case cdNoError '1052 - You will never see this most likely
        strMsg = "No Error Reported"
    Case cdNoSeekComplete '1053
        strMsg = "A seek command was interrupted by another command"
    Case cdNTIOError '1054
        strMsg = "A NT disc I/O operation failed"
    Case cdFormatInProgress '1055
        strMsg = "A format is in progress causing operation failure"
    End Select

    'return the error string
    strMsg = strMsg & " (" & CStr(DriveError) & ")."

    GetDriveErrorMessage = strMsg

End Function

Private Function ChkDir(StrDirectory As String) As String
    '���Ŀ¼�Ƿ��в��������ַ���������
    ChkDir = Replace(StrDirectory, "/", "")
    ChkDir = Replace(StrDirectory, "\", "")
    ChkDir = Replace(StrDirectory, ":", "")
    ChkDir = Replace(StrDirectory, "*", "")
    ChkDir = Replace(StrDirectory, "?", "")
    ChkDir = Replace(StrDirectory, """", "")
    ChkDir = Replace(StrDirectory, "<", "")
    ChkDir = Replace(StrDirectory, ">", "")
    ChkDir = Replace(StrDirectory, "|", "")
End Function
