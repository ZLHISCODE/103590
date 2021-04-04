VERSION 5.00
Object = "{82809FC2-3B17-4941-8A37-713AA0519BB1}#1.0#0"; "DVDProX2.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreateCD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "创建CD"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame4 
      Caption         =   "刻录信息"
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
         Caption         =   "信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5490
      TabIndex        =   6
      Top             =   7230
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "写入(&W)"
      Height          =   350
      Left            =   2370
      TabIndex        =   5
      Top             =   7230
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "CD 打包选项:"
      Height          =   1815
      Left            =   4920
      TabIndex        =   14
      Top             =   4290
      Width           =   4455
      Begin VB.OptionButton optPacking 
         Caption         =   "DICOMDIR和独立CD观片站"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optPacking 
         Caption         =   "只有DICOMDIR"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "刻录选项:"
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   4290
      Width           =   4575
      Begin VB.CommandButton CmdWriterCDOption 
         Caption         =   "高级选项"
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
         Caption         =   "速度:"
         Height          =   180
         Left            =   630
         TabIndex        =   18
         Top             =   690
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CD名称:"
         Height          =   180
         Left            =   450
         TabIndex        =   16
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "刻录光驱:"
         Height          =   180
         Left            =   270
         TabIndex        =   15
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DICOM CD 内容："
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
         Caption         =   "全部删除"
         Height          =   350
         Left            =   7920
         TabIndex        =   8
         Top             =   1080
         Width           =   1100
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "删除"
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
         Caption         =   "最大容量："
         Height          =   255
         Left            =   7920
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "需要的空间："
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
        MsgBox "不能打开选择的刻录光驱!", vbInformation, gstrSysName
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
    
    intUseJoliet = GetSetting("ZLSOFT", "私有模块\本地\" & App.ProductName & "\刻录设置", "使用Joliet", 1)
    intCDRWMode = GetSetting("ZLSOFT", "私有模块\本地\" & App.ProductName & "\刻录设置", "使用CDRW模式", 1)
    blHighCompatibilityMode = GetSetting("ZLSOFT", "私有模块\本地\" & App.ProductName & "\刻录设置", "高兼容DVD模式", 0)
    blCheckImage = GetSetting("ZLSOFT", "私有模块\本地\" & App.ProductName & "\刻录设置", "不使用高速缓存", 0)
    blCloseDisk = GetSetting("ZLSOFT", "私有模块\本地\" & App.ProductName & "\刻录设置", "关闭光盘", 1)
    blTestWriter = GetSetting("ZLSOFT", "私有模块\本地\" & App.ProductName & "\刻录设置", "测试写入", 1)
    blBufferProof = GetSetting("ZLSOFT", "私有模块\本地\" & App.ProductName & "\刻录设置", "缓存校验", 1)
    blAutoVerify = GetSetting("ZLSOFT", "私有模块\本地\" & App.ProductName & "\刻录设置", "自动数据校验", 1)
    
    On Error GoTo errh
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    '检查创建CD的条件是否满足
    If dimgsInCD.Count <= 0 Then
        MsgBox "没有图像无法创建CD,请从观片站重新选择图像。", vbInformation, gstrSysName
        Exit Sub
    End If
'    If Me.TxtPath = "" Then
'        MsgBox "请输入CD的路径。", , "创建CD"
'        Me.TxtPath.SetFocus
'        Exit Sub
'    End If
    If Me.txtCDName = "" Then
        MsgBox "请输入CD名称。", vbInformation, gstrSysName
        Me.txtCDName.SetFocus
        Exit Sub
    End If
'    If left(Me.txtSpacing(0).Text, InStr(Me.txtSpacing(0).Text, "MB") - 1) - INT_MAX_CAPACITY > 0 Then
'        MsgBox "选择的图像占用空间超过一张光盘的容量，请先删除部分图像。", , "创建CD"
'        Exit Sub
'    End If
    '保存图像，并创建DICOMDIR
'    If Dir(Me.TxtPath, vbDirectory) = "" Then
'        MsgBox "不是有效的路径请重新选择!", vbInformation, "提示"
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
        '如果目录不存在，则创建目录
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
    
    '保存“独立CD观片站”
    If Me.optPacking(1).Value = True Then
        '将特定目录的文件复制到strPath\PACSLite目录中
        strAppPath = App.Path & IIf(Len(App.Path) > 3, "\", "") & STR_ATTACHED_FILE_PATH
        If Dir(strAppPath, vbDirectory) <> "" Then
            If Dir(strAppPath, vbDirectory) = "" Then
                MsgBox "没有找到复制文件路径！", vbInformation, gstrSysName
                Exit Sub
            End If
            fs.CopyFile strAppPath & "\*.*", strRootPath
        End If
    End If
    '刻录
    If DVDWriterPro.GetMediaType() = mtNotLoaded Then
        MsgBox "请插入可写入的光盘!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call DVDWriterPro.CloneDirectoryToISO("\", strRootPath & "\*.*")
    
    
    If (DVDWriterPro.GetDiscFreeSpaceBlocks() < (DVDWriterPro.GetISOVolumeSizeBlocks())) Then
        strMsg = "自由空间: " & ConvertBytesToMBString(DVDWriterPro.ConvertBlocksToBytes(DVDWriterPro.GetDiscFreeSpaceBlocks(), wtpDataMode1)) & " MB!" & _
                vbCrLf & "小于需使用空间: " & ConvertBytesToMBString(DVDWriterPro.ConvertBlocksToBytes(DVDWriterPro.GetISOVolumeSizeBlocks(), wtpDataMode1)) & " MB ."
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
        MsgBox "不能写入光盘!", vbCritical, gstrSysName
        Exit Sub
    End If
    
    fs.DeleteFolder strRootPath, True
    '保存界面信息到注册表

    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "打包方式", IIf(optPacking(0).Value, -1, 0))

    Exit Sub
errh:
    MsgBox "发生错误:" & err.Description & vbCrLf & "错误号:" & err.Number, vbExclamation, gstrSysName
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
            '删除图像信息
            Me.trvCDContents.Nodes.Remove nodeTemp.Index
            '如果本次检查中没有图像，则将检查信息删除
            If nodeUp.Child Is Nothing Then
                Set nodeTemp = nodeUp.Parent
                Me.trvCDContents.Nodes.Remove nodeUp.Index
                '如果当前病人没有图像，则将病人信息删除
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
    LabInformation.Caption = "完成Cach" & Format(nPercentComplete, "0#") & " %"
End Sub

Private Sub DVDWriterPro_ClosingDisc()
    LabInformation.Caption = "关闭光盘......."
End Sub

Private Sub DVDWriterPro_ClosingSession()
    LabInformation.Caption = "关闭信息...."
End Sub

Private Sub DVDWriterPro_ClosingTrack(ByVal lTrackNumber As Long)
    LabInformation.Caption = "关闭轨迹...."
End Sub

Private Sub DVDWriterPro_CreatingDirectories()
    LabInformation.Caption = "创建目录...."
End Sub

Private Sub DVDWriterPro_FileVerifyComplete(ByVal lFilesCompared As Long, ByVal lFilesMatched As Long)
    LabInformation.Caption = "自动检验...(" & lFilesMatched & " 和 " & lFilesCompared & " 匹配完成)"
    Call DVDWriterPro.EjectLoad(False)
    MsgBox "写入和校验完成!", vbInformation, gstrSysName

End Sub

Private Sub DVDWriterPro_FileVerifyStart(ByVal lFilesToCompare As Long)
    LabInformation.Caption = "自动校验... " & lFilesToCompare & " 文件."
End Sub

Private Sub DVDWriterPro_FileVerifyStatus(ByVal sItemDestPath As String, ByVal sSourceFilePath As String, ByVal lFileBytesCompared As Long, ByVal lFileSizeTotal As Long, ByVal fvStatus As DVDPROX2LibCtl.eVerifyStatus, ByVal lCurrentFile As Long, ByVal lFilesToCompare As Long, bCancel As Boolean)
    Dim intPercentVerified As Integer

    Select Case fvStatus
        Case fvsComparing
        Case fvsMatched
        Case fvsIncorrectByteCount
            MsgBox "校验错误 - 源文件大小和目标不匹配: " & vbCrLf & _
                    sItemDestPath & " (目标)" & vbCrLf & sSourceFilePath & " (源)", vbExclamation, gstrSysName
    
        Case fvsNoMatch
            MsgBox "校验错误 - 源文件大小和目标不匹配: " & vbCrLf & _
                    sItemDestPath & " (目标)" & vbCrLf & sSourceFilePath & " (源)", vbExclamation, gstrSysName
    
        Case fvsReadingDiscError
            MsgBox "校验错误 - 源文件大小和目标不匹配: " & vbCrLf & _
                    sItemDestPath & " (目标)" & vbCrLf & sSourceFilePath & " (源)", vbExclamation, gstrSysName
    
        Case fvsReadingSourceError
            MsgBox "校验错误 - 源文件大小和目标不匹配: " & vbCrLf & _
                    sItemDestPath & " (目标)" & vbCrLf & sSourceFilePath & " (源)", vbExclamation, gstrSysName
    End Select

    LabInformation.Caption = "校验路径: " & sItemDestPath

    intPercentVerified = ((lCurrentFile / lFilesToCompare) * 100)

    ProgressBar1.Value = intPercentVerified

    DoEvents
End Sub

Private Sub DVDWriterPro_PreparingToWrite()

    LabInformation.Caption = "预写入...."

    Me.cmdOK.Enabled = False
End Sub

Private Sub DVDWriterPro_ReadingTrackFile(ByVal sFileName As String, ByVal lFileIndex As Long, ByVal lTrackNumber As Long)
    LabInformation.Caption = "轨迹: " & Format(lTrackNumber, "0#") & " - 读取..." & CStr(lFileIndex) & " - " & sFileName
End Sub

Private Sub DVDWriterPro_ReadingTrackFileError(ByVal TrackFileError As DVDPROX2LibCtl.eTrackFileError, ByVal sFileName As String, ByVal lTrackNumber As Long)
    LabInformation.Caption = "读文件:" & sFileName & "时发生错误!"
End Sub

Private Sub DVDWriterPro_ReplaceImportedISOFile(ByVal sDestPath As String, ByVal sNewSourcePath As String, ByVal sFileName As String, bReplaceFile As Boolean)
    Dim lngResult As Long

    lngResult = MsgBox("写入文件时发现文件名相同是否替换?", vbOKCancel + vbQuestion, gstrSysName)
    
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
    LabInformation.Caption = "取消写入......"
    Me.cmdOK.Enabled = True
    MsgBox "写入被取消!", vbInformation, gstrSysName
End Sub

Private Sub DVDWriterPro_WriteComplete()
    LabInformation.Caption = "写入完成!"

    Me.cmdOK.Enabled = True

    If DVDWriterPro.AutoVerify = False Then
        MsgBox "写入完成!", vbInformation, gstrSysName
    End If

    If (DVDWriterPro.TestWrite = False) And (DVDWriterPro.AutoVerify = False) Then
        Call DVDWriterPro.EjectLoad(False)
    End If
End Sub

Private Sub DVDWriterPro_WriteError(ByVal WriteError As DVDPROX2LibCtl.eWriteErrorType, ByVal DriveError As DVDPROX2LibCtl.eCDError, ByVal sErrorInfo As String, ByVal sSenseInfo As String)
    Dim strError As String

    strError = "写入时发生错误: (" & CStr(WriteError) & ")   " & vbCrLf

    If WriteError = errDriveError Then
        strError = strError & GetDriveErrorMessage(DriveError) & vbCrLf & " 发送错误数据: " & sSenseInfo
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
    
    '获取被选中的图像
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
    
    '将图像的信息添加到treeview中，分成三层；
    '第一层 病人：text＝“姓名：“ & 姓名；tag＝Patient ID
    '第二层 检查：text＝”检查：“ & 检查描述；tag＝Study UID
    '第三层 图像：text＝“图像：” & 图像UID； tag＝Instance UID
    For Each img In dimgsInCD
        blnInserted = False
        '检查treCDContents的第一层
        If Me.trvCDContents.Nodes.Count > 1 Then
            '检查PatientID是否有重复的
            Set node1 = Me.trvCDContents.Nodes(1)
            While (Not node1 Is Nothing) And blnInserted = False
                '在病人层次查找
                If node1.Tag <> img.PatientID Then  '查找病人层次的下一个节点
                    Set node1 = node1.Next
                Else    '查找检查层次
                    Set node2 = node1.Child
                    While (Not node2 Is Nothing) And blnInserted = False
                        '在检查层次查找
                        If node2.Tag <> img.StudyUID Then   '查找检查层次的下一个节点
                            Set node2 = node2.Next
                        Else    '查找图像层次
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
            '往trvCDContents里面添加信息
            subAddNodeToContents "PATIENT", Nothing, img
            blnInserted = True
        End If
    Next img
    subGetImageCapacity
    Me.txtSpacing(1).Text = INT_MAX_CAPACITY & "MB"
    '注册控件
    Me.DVDWriterPro.LicenseCode = "10LBTZY9V42HTZKTKL27S"
    LoadDriveCombo
    
    Me.optPacking(0).Value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "打包方式", 0)
    Me.optPacking(1).Value = Not Me.optPacking(0).Value
    
End Sub

Private Sub subAddNodeToContents(strLevel As String, nodeCurrent As Node, img As DicomImage)
    Dim node1 As Node
    Dim node2 As Node
    Dim node3 As Node
    
    If UCase(strLevel) = "PATIENT" Then
        Set node1 = Me.trvCDContents.Nodes.Add(, , , "姓名：" & img.Name)
        node1.Tag = img.PatientID
        Set node2 = Me.trvCDContents.Nodes.Add(node1, tvwChild, , "检查：" & img.StudyDescription)
        node2.Tag = img.StudyUID
        Set node3 = Me.trvCDContents.Nodes.Add(node2, tvwChild, , "图像：" & img.InstanceUID)
        node3.Tag = img.InstanceUID
    ElseIf UCase(strLevel) = "STUDY" Then
        Set node2 = Me.trvCDContents.Nodes.Add(nodeCurrent, tvwChild, , "检查：" & img.StudyDescription)
        node2.Tag = img.StudyUID
        Set node3 = Me.trvCDContents.Nodes.Add(node2, tvwChild, , "图像：" & img.InstanceUID)
        node3.Tag = img.InstanceUID
    ElseIf UCase(strLevel) = "IMAGE" Then
        Set node3 = Me.trvCDContents.Nodes.Add(nodeCurrent, tvwChild, , "图像：" & img.InstanceUID)
        node3.Tag = img.InstanceUID
    End If
End Sub
 
Private Sub subGetImageCapacity()
    '返回值以MB为单位
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
    '得到可读写的CD
    Dim intDrives As Integer

    If DVDWriterPro.InitDrives(False) = False Then
        MsgBox "驱动程序不能初使化!", vbInformation, gstrSysName
    End If
    
    CboDrivers.Clear

    '查找可读写的CD
    For intDrives = 0 To DVDWriterPro.GetDriveCount() - 1
        If DVDWriterPro.IsDriveWriter(intDrives) = True Then
            CboDrivers.AddItem DVDWriterPro.GetDriveLetter(intDrives) & ": " & DVDWriterPro.GetDriveVendor(intDrives) & " " & DVDWriterPro.GetDriveModel(intDrives)
            CboDrivers.ItemData(CboDrivers.NewIndex) = intDrives
        End If
    Next
    
    '设备第一个CD
    If CboDrivers.ListCount > 0 Then
        CboDrivers.ListIndex = 0
    Else
        MsgBox "没有找到可以读写的刻录光驱!", vbExclamation, gstrSysName
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
        CboWriterSpeeds.AddItem "默认"
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
        strMsg = "异常I/O中止"
    Case cdLogicalUnitNotReady '1002
        strMsg = "这个戏动没有准备好"
    Case cdLogicalUnitCommFailed '1003
        strMsg = "发送一个驱动信息挫败"
    Case cdDeviceTrackingError '1004
        strMsg = "这个驱动不能完成轨迹"
    Case cdWriteGenericError '1005
        strMsg = "不明的写错误"
    Case cdWriteRecoveryNeeded '1006
        strMsg = "Writing occurred, but recovery is needed"
    Case cdWriteRecoveryFailed '1007
        strMsg = "企图恢复失败"
    Case cdWriteLossOfStreaming '1008
        strMsg = "A buffer under-run has occurred"
    Case cdReadUnrecovered '1009
        strMsg = "这个光盘总是不能读"
    Case cdReadRetriesExhausted '1010
        strMsg = "这个驱动企图恢复读失败"
    Case cdReadErrorTooLong '1011
        strMsg = "读取超时"
    Case cdReadLECUncorrectable '1012
        strMsg = "While reading, the LEC was not recovered"
    Case cdReadCIRCUnrecovered '1013
        strMsg = "The CIRC could not be validated"
    Case cdReadUPCEANFailed '1014
        strMsg = "Reading of the UPC failed"
    Case cdReadISRCFailed '1015
        strMsg = "Reading of the ISRC failed"
    Case cdReadLossOfStreaming '1016
        strMsg = "读取数据时被中断"
    Case cdPositioningError '1017
        strMsg = "驱动不能写入媒体"
    Case cdParameterListLengthError '1018
        strMsg = "一个不兼容的长参数发送到了驱动"
    Case cdSynchronousTransferError '1019
        strMsg = "在这个驱动上发一个迁移错误"
    Case cdInvalidCommandCode '1020
        strMsg = "一个失效的命令发送到了这个驱动上"
    Case cdLBAOutOfRange '1021
        strMsg = "Error trying to write past the end of the media"
    Case cdInvalidCDBField '1022
        strMsg = "失效命令失败"
    Case cdInvalidParamterListField '1023
        strMsg = "一个不兼容的参数发送到了驱动"
    Case cdParameterNotSupported '1024
        strMsg = "不支持一个命令参数"
    Case cdParamterValueInvalid '1025
        strMsg = "一个命令参数失校的值"
    Case cdBusOrDeviceReset '1026
        strMsg = "The SCSI/ATAPI bus was reset and caused a write failure"
    Case cdParametersChanged '1027
        strMsg = "A command parameter changed while in progress"
    Case cdIncompatibleMedium '1028
        strMsg = "这个光盘不能兼容这个驱动"
    Case cdReadUnknownMediumFormat '1029
        strMsg = "这个驱动不能兼容这种格式的光盘"
    Case cdReadIncompatibleMediumFormat '1030
        strMsg = "这个光盘不能兼容当前驱动"
    Case cdWriteUnknownMediumFormat '1031
        strMsg = "光盘格式未知"
    Case cdIncompatibleWriteFormat '1032
        strMsg = "这个驱动不能写入是因为格式矛盾"
    Case cdMediaNotPresent '1033
        strMsg = "这个光盘不能引导"
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
    '检查目录是否有不正常的字符，并修正
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
