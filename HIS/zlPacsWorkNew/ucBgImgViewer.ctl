VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucBgImgViewer 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ScaleHeight     =   5220
   ScaleWidth      =   7620
   Begin DicomObjects.DicomViewer dcmViewer 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _Version        =   262147
      _ExtentX        =   12938
      _ExtentY        =   7011
      _StockProps     =   35
      BackColor       =   0
      CellSpacing     =   1
      UseMouseWheel   =   -1  'True
   End
   Begin VB.Timer timerState 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   4200
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6000
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   400
      ImageHeight     =   300
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucBgImgViewer.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucBgImgViewer.ctx":57E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucBgImgViewer.ctx":AFD24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucBgImgViewer.ctx":107BB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picScroll 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   310
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   4800
      Width           =   7455
      Begin VB.TextBox txtRecordCount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "8"
         ToolTipText     =   "ÿҳ��¼����"
         Top             =   0
         Width           =   375
      End
      Begin VB.ComboBox cbxPage 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label labState 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " --"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   10
         Width           =   495
      End
   End
End
Attribute VB_Name = "ucBgImgViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const VALIDTIME = 30

Private Enum TResultState
    rsClear = 0
    rsOk = 1
    rsFailed = 2
End Enum
 

Private mlngViewRows As Integer
Private mlngViewCols As Integer

Private mlngSelectIndex As Integer

Private mlngPageIndex As Long       '��ǰҳ����
Private mlngPageCount As Long       '��ǰҳ����
Private mlngPageRecord As Long      'ÿҳ��ʾ��

Private mlngServerTime As Long

Private mblnIsBGReadProcessing As Boolean   '�Ƿ���к�̨ͼ���ȡ����
Private mblnIsTimerWorking As Boolean
Private mblnIsRefreshing As Boolean
Private mblnIsPageConfig As Boolean

Private mlngTimeOut As Long

Private mdtStartTime As Date

Private maryImgInfo() As clsBgImgInfo
Private maryImgBuf() As clsBgImgInfo

Private mblnIsDrawOrder As Boolean
Private mblnIsDrawHint As Boolean
Private mblnIsShowCheck As Boolean
Private mblnIsShowState As Boolean
Private mlngSelColorStyle As ColorConstants

Private mblnBGServerStarted As Boolean

Private mstrUploadCmdNames As String

Private mblnIsClickEvent As Boolean


'�¼�����
'Public Event OnCmdEvent(ByVal strCmd As String)
Public Event OnClick(ByVal lngImgIndex As Long)
Public Event OnDbClick(ByVal lngImgIndex As Long)
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)


Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property


Property Get PageRecordCount() As Long
    PageRecordCount = mlngPageRecord
End Property

Property Let PageRecordCount(ByVal value As Long)
    mlngPageRecord = value
    
    If mlngPageRecord <= 0 Then mlngPageRecord = 8
    
    txtRecordCount.tag = "1"
    
    txtRecordCount.Text = mlngPageRecord
    
    txtRecordCount.tag = ""
End Property


'ѡ����ɫ��ʽ
Property Get SelColorStyle() As ColorConstants
    SelColorStyle = mlngSelColorStyle
End Property

Property Let SelColorStyle(ByVal value As ColorConstants)
    mlngSelColorStyle = value
End Property

'�Ƿ�������
Property Get IsDrawOrder() As Boolean
    IsDrawOrder = mblnIsDrawOrder
End Property

Property Let IsDrawOrder(ByVal value As Boolean)
    mblnIsDrawOrder = value
End Property

'�Ƿ������ʾ
Property Get IsDrawHint() As Boolean
    IsDrawHint = mblnIsDrawHint
End Property

Property Let IsDrawHint(ByVal value As Boolean)
    mblnIsDrawHint = value
End Property

'�Ƿ���ʾ״̬
Property Get IsShowState() As Boolean
    IsShowState = mblnIsShowState
End Property

Property Let IsShowState(ByVal value As Boolean)
    mblnIsShowState = value
    picScroll.Visible = value
End Property

'�Ƿ���ʾ��ѡ��
Property Get IsShowCheck() As Boolean
    IsShowCheck = mblnIsShowCheck
End Property

Property Let IsShowCheck(ByVal value As Boolean)
    mblnIsShowCheck = value
End Property


'��ʱ
Property Get TimeOut() As Long
    TimeOut = mlngTimeOut
End Property

Property Let TimeOut(ByVal value As Long)
    mlngTimeOut = value
End Property

'ͼ������
Property Get ImgCount() As Long
On Error GoTo errhandle
    ImgCount = UBound(maryImgInfo) + 1
Exit Property
errhandle:
    ImgCount = 0
End Property
 
'����Ԥ������
Property Get Viewer() As Object
    Set Viewer = dcmViewer
End Property

'ѡ���ͼ������
Property Get SelImgIndex() As Long
    SelImgIndex = mlngSelectIndex
End Property


Private Function IsValid(aryImgInfos() As clsBgImgInfo) As Boolean
On Error GoTo errhandle
    IsValid = IIf(UBound(aryImgInfos) >= 0, True, False)
Exit Function
errhandle:
    IsValid = False
End Function

Public Sub ConstructionImgData(objBgImgInfo As clsBgImgInfo)
'����ͼ������
    Dim lngBound As Long
    
    lngBound = ImgBufCount
    ReDim Preserve maryImgBuf(lngBound)
    
    Set maryImgBuf(lngBound) = objBgImgInfo
End Sub

Private Function ImgBufCount()
On Error GoTo errhandle
    ImgBufCount = UBound(maryImgBuf) + 1
Exit Function
errhandle:
    ImgBufCount = 0
End Function
 
Public Sub EraseImgData()
'����ͼ������
    EraseAry maryImgInfo
    EraseAry maryImgBuf
End Sub

Private Sub EraseAry(ByRef ary)
On Error Resume Next
    Erase ary
    If err.Number <> 0 Then
        Debug.Print "ucBgImgViewer.EraseAry Err:" & err.Description
    End If
End Sub

Property Get ImageInfo(ByVal lngIndex As Long) As clsBgImgInfo
    Set ImageInfo = maryImgInfo(lngIndex)
End Property


'Public Function GetImageInfos() As clsBgImgInfo()
'    Dim i As Long
'    Dim aryImgInfos() As clsBgImgInfo
'
'    If ImgCount <= 0 Then Exit Function
'
'    ReDim aryImgInfos(UBound(maryImgInfo))
'
'    For i = 0 To UBound(maryImgInfo)
'        Set aryImgInfos(i) = Nothing
'
'        If maryImgInfo(i) Is Nothing Then
'            Set aryImgInfos(i) = maryImgInfo(i).CopyNew
'        End If
'    Next
'
'    GetImageInfos = aryImgInfos
'End Function

Private Sub ReadImgInfoBuf(Optional ByVal blnIsReset As Boolean = True)
    Dim i As Long
    Dim lngBound As Long
    
    If blnIsReset Then
            EraseAry maryImgInfo
    End If
    
    For i = 0 To ImgBufCount - 1
        lngBound = ImgCount
        ReDim Preserve maryImgInfo(lngBound)
        
        Set maryImgInfo(lngBound) = maryImgBuf(i)
        
        If maryImgInfo(lngBound).ImageOrder <= 0 Then
            maryImgInfo(lngBound).ImageOrder = lngBound + 1
        End If
    Next
     
    EraseAry maryImgBuf
End Sub

Private Function GetRootHwnd() As Long
    GetRootHwnd = GetAncestor(hwnd, GA_ROOT)
End Function

Public Sub Init()
    mblnBGServerStarted = StartBackGroundServer(IIf(CheckExeIsRun("ZL9PACSIMGTRANS.EXE"), False, True))
End Sub

Public Sub Refresh(Optional ByVal blnIsReset As Boolean = True)
'����ͼ��
On Error GoTo errhandle
    mdtStartTime = 0
    mlngPageIndex = 1
    mblnIsBGReadProcessing = False
    
    If mblnIsRefreshing Then

        MsgboxH GetRootHwnd, "ͼ�������δ��ɣ����Ժ�����!", vbOKOnly, "��ʾ"
        
        Call ReadImgInfoBuf(blnIsReset)
        
        EraseAry maryImgBuf
        
        Exit Sub
    End If
    
    mblnIsRefreshing = True
    
    timerState.Enabled = False
    
    '��ȡ�����⣬������ʵ�ʶ�ȡ������
    Call ReadImgInfoBuf(blnIsReset)
    
    Call WaitUnlock
    
    If IsValid(maryImgInfo) = False Then
        mblnIsRefreshing = False
        Exit Sub
    End If

    '����ͼ������
    Call ProcessImgCmds
    
    '�����ǰͼ����ʾ
    Call ClearImgView

    Call DrawResultState(rsClear)
    
    
'    ���÷�ҳ��Ϣ
    Call ConfigPage
    
    If mblnIsBGReadProcessing Then
        Call ReDrawImages(mlngPageIndex, True)
        
        timerState.Enabled = True
        Call timerState_Timer 'ִ��timer���̣�������̨����
        
        If mblnBGServerStarted = False And mblnIsBGReadProcessing Then
            MsgboxH GetRootHwnd, "��̨���ݴ����������ʧ�ܣ����ݴ���ģʽ������Ϊʵʱģʽ��", vbOKOnly, "��ʾ"
'            Call ShowBall("��̨��������ʧ�ܣ�������Ϊʵʱ���䡣")
            
            mblnIsBGReadProcessing = False
            timerState.Enabled = False
            
            '��̨����ִ��ʧ��ʱ���Ž��д˴���
            Call ProcessImgCmds(True)
            
            Call DrawImgStates(True)
        End If
    End If
    
    mblnIsRefreshing = False
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
    mblnIsRefreshing = False
End Sub

  


Public Function DetectionImgProcess(ByRef lngErrorCount As Long) As Long
'���ͼ����
    Dim i As Long
    Dim objImgInfo As clsBgImgInfo
    Dim lngProcesCount As Long
    
    lngProcesCount = 0
    lngErrorCount = 0
    
    For i = 0 To ImgCount - 1
        If timerState.Enabled = False Then Exit Function
        
        Set objImgInfo = maryImgInfo(i)
        If objImgInfo.LoadState = lsRedo _
            Or objImgInfo.LoadState = lsSent Then
            
            Call UpdateProcessState(objImgInfo)
        End If
        
        '���㴦���е�����
        If objImgInfo.LoadState = lsRedo _
            Or objImgInfo.LoadState = lsSent Then
            
            lngProcesCount = lngProcesCount + 1
        End If
        
        '�����������
        If objImgInfo.LoadState = lsError Then
            lngErrorCount = lngErrorCount + 1
        End If
    Next
    
    DetectionImgProcess = lngProcesCount
End Function

Private Sub UpdateProcessState(objImgInfo As clsBgImgInfo)
    Dim strCmdFile As String
    Dim dtEndTime As Date
    Dim dtStartTime As Date
    Dim objIni As New clsIniFile
    Dim lngReDo As Long
    Dim strProcessErr As String
    Dim strFile As String
    Dim blnIsEnterFailedDir As Boolean
    Dim strCheckFilePath As String
    
    If objImgInfo.LoadState <> lsRedo And objImgInfo.LoadState <> lsSent Then Exit Sub
    
    If objImgInfo.ImgCommand <> icUpLoad Then
        '�жϱ����Ƿ�����ļ�
        strFile = FormatFilePath(objImgInfo.FilePath & "\" & objImgInfo.Filename)
        If IsValidFile(strFile) Then
            '������ش����ļ���˵������ɹ�
            objImgInfo.IsReDrawed = False
            objImgInfo.LoadState = lsLocal
            objImgInfo.ErrorInfo = ""
            
            Exit Sub
        End If
    End If
    
    
    strCmdFile = GetImgCmdFile(objImgInfo)
    
    blnIsEnterFailedDir = False
    If Trim(Dir(strCmdFile, vbHidden)) = "" Then
        objImgInfo.StartTime = CDate(0)
        strCmdFile = GetImgCmdFailed(objImgInfo)
        
        '�ж�ʧ��Ŀ¼�Ƿ��������ļ�
        If Trim(Dir(strCmdFile, vbHidden)) = "" Then
            '�п��ܺ�̨�����Ѿ�������ɲ�ɾ���������ļ�
            '���Ը���mstrUploadCmdNames��������Ƿ�ɹ����ɹ������ļ�
            'Ҳ���ж�ftp�ļ��Ƿ���ڣ�������ftp����ʱ�����ӣ�...
            '��Ҫ�����ļ��ϴ�ʱ�������ɾ�������
            If InStr(mstrUploadCmdNames, ";" & objImgInfo.Key & ";") > 0 Then
                objImgInfo.IsReDrawed = False
                objImgInfo.LoadState = lsLocal
                objImgInfo.ErrorInfo = ""
                
                mstrUploadCmdNames = Replace(mstrUploadCmdNames, ";" & objImgInfo.Key & ";", "")
            Else
                objImgInfo.LoadState = lsRedo
                objImgInfo.ErrorInfo = "δ��⵽ִ������,�������,������."
            End If
            
            Exit Sub
        Else
            blnIsEnterFailedDir = True
        End If
    End If
    
    '�����ϴ������ļ��Ĵ���ʽ
    Call objIni.SetIniFile(strCmdFile)
    
    dtStartTime = objImgInfo.StartTime
    
    dtEndTime = CDate(objIni.ReadValue("OTHERINFO", "ENDTIME", "0"))
    lngReDo = Val(objIni.ReadValue("OTHERINFO", "REDO", "0"))
    strProcessErr = objIni.ReadValue("OTHERINFO", "ERRORINFO", "")
'    Debug.Print strProcessErr
    objImgInfo.Redo = lngReDo
    
    If dtEndTime > CDate(0) And dtEndTime > dtStartTime Then
        If lngReDo = -1 Or blnIsEnterFailedDir Then
            'Ϊ-1��ʾ������ɣ�����ʧ�ܣ���Ϊ�ɹ���Ӧ��ɾ�������ļ���������ɾ������ʱ
            If Len(strProcessErr) > 0 Then
                objImgInfo.IsReDrawed = False
                objImgInfo.LoadState = lsError
                objImgInfo.ErrorInfo = strProcessErr
            Else
                objImgInfo.IsReDrawed = False
                
                '���¼�Ȿ���ļ��Ƿ����
                strFile = FormatFilePath(objImgInfo.FilePath & "\" & objImgInfo.Filename)
                If IsValidFile(strFile) Then
                    '������ش����ļ���˵������ɹ�
                    objImgInfo.LoadState = lsLocal
                    objImgInfo.ErrorInfo = ""
                Else
                    objImgInfo.LoadState = lsError
                    objImgInfo.ErrorInfo = "δ��⵽���غ��ͼ���ļ�."
                End If
            End If
        Else
            '����0��ʾ���Դ���Ĵ���
            objImgInfo.LoadState = lsRedo
            objImgInfo.ErrorInfo = strProcessErr
        End If
    End If
    
    Set objIni = Nothing
End Sub

Private Sub ResetImgDrawState()
    Dim i As Long
    
    For i = 0 To ImgCount - 1
        maryImgInfo(i).IsReDrawed = False
    Next
End Sub

Private Sub ReDrawImages(ByVal lngPageIndex As Long, Optional ByVal blnIsPageChange As Boolean = False)
    Dim lngStartIndex As Long
    Dim lngCount As Long
    Dim i As Long
    Dim blnIsAbort As Boolean
    
    If blnIsPageChange Then
        '�ı䵱ǰ��ʾҳ
        Call ClearImgView
        Call ResetImgDrawState
    End If
    
    lngStartIndex = (lngPageIndex - 1) * mlngPageRecord
    lngCount = ImgCount
    
    blnIsAbort = False
    For i = lngStartIndex To (lngStartIndex + mlngPageRecord) - 1
        
        If i >= lngCount Then Exit Sub  '����ͼ��Χ
        
        If maryImgInfo(i).IsReDrawed = False Then
            '���������ļ�
            If maryImgInfo(i).IsCreateCmdFile = False Then
                If blnIsAbort Then
                    maryImgInfo(i).LoadState = lsError
                    maryImgInfo(i).ErrorInfo = "����ֹ����"
                    maryImgInfo(i).IsReDrawed = False
                Else
                    If ProcessImgCmd(maryImgInfo(i), , Not mblnBGServerStarted) = frAbort Then blnIsAbort = True
                End If
                maryImgInfo(i).IsCreateCmdFile = True
            End If
            
            Call ReDrawImage(maryImgInfo(i))
        End If
    Next
End Sub


Private Sub ProcessImgCmds(Optional ByVal blnIsForceOnline As Boolean = False)
'����ͼ������
    Dim i As Long
    Dim lngImgCount As Long
    Dim blnIsAbort As Boolean
    
    lngImgCount = mlngPageRecord
    If lngImgCount > ImgCount Then
        lngImgCount = ImgCount
    End If
    
    blnIsAbort = False
    For i = 0 To lngImgCount - 1  'ImgCount - 1
        If maryImgInfo(i).IsCreateCmdFile = False Or blnIsForceOnline Then
            If blnIsAbort Then
                maryImgInfo(i).LoadState = lsError
                maryImgInfo(i).ErrorInfo = "����ֹ����"
                maryImgInfo(i).IsReDrawed = False
            Else
                If ProcessImgCmd(maryImgInfo(i), , blnIsForceOnline) = frAbort Then blnIsAbort = True
            End If
            maryImgInfo(i).IsCreateCmdFile = True
        End If
    Next
     
End Sub


Private Function ProcessImgCmd(objImgInfo As clsBgImgInfo, _
    Optional ByVal blnIsRedo As Boolean = False, Optional ByVal blnIsForceOnline As Boolean = False) As ftpResult
    Dim strError As String
    Dim strFailedFile As String
    Dim objIni As clsIniFile
    Dim strTransErr As String
    
    ProcessImgCmd = frNormal
        
    If objImgInfo.IsBackGround And blnIsForceOnline = False Then
        '��̨����
        If objImgInfo.LoadState = lsNone Or blnIsRedo Then
            
            If objImgInfo.ImgCommand = icDownload Or objImgInfo.ImgCommand = icReadly Then
                If FileExists(objImgInfo.FilePath & objImgInfo.Filename) Then
                
                    '�ж��Ƿ�Ϊ�ϴ��ϴδ���ʧ�ܵ��ļ�
                    strError = GetTransFailedState(objImgInfo.Key, icUpLoad)
                    If Len(strError) > 0 Then
                        If InStr(strError, "��...") > 0 Then
                            objImgInfo.ImgCommand = icUpLoad
                            objImgInfo.LoadState = lsRedo
                            
                            mdtStartTime = Now
                        Else
                            objImgInfo.ImgCommand = icUpLoad
                            objImgInfo.LoadState = lsError
                            If Len(objImgInfo.ErrorInfo) <= 0 Then objImgInfo.ErrorInfo = "ͼ���ϴ�ʧ��>>" & Replace(strError, "ͼ���ϴ�ʧ��", "")
                        End If
                    Else
                        objImgInfo.LoadState = lsLocal
                        objImgInfo.ErrorInfo = ""
                    End If
                    
                    mblnIsBGReadProcessing = True
                    Exit Function
                End If
            End If
            
            If objImgInfo.ImgCommand = icUpLoad And FileExists(objImgInfo.FilePath & objImgInfo.Filename) = False Then
                objImgInfo.LoadState = lsError
                objImgInfo.ErrorInfo = "δ�ҵ����ϴ��������ļ�"
                
                mblnIsBGReadProcessing = True
                Exit Function
            End If
            
            
            If TransCmd(objImgInfo, GetImgCmdFile(objImgInfo), strError) = False Then
                objImgInfo.LoadState = lsError
                objImgInfo.ErrorInfo = strError
            Else
                objImgInfo.LoadState = lsSent
                objImgInfo.ErrorInfo = ""
                
                If objImgInfo.ImgCommand = icUpLoad Then
                    mstrUploadCmdNames = mstrUploadCmdNames & ";" & objImgInfo.Key & ";"
                End If
            End If
            
            mdtStartTime = Now
            mlngServerTime = VALIDTIME
            
            mblnIsBGReadProcessing = True
'        ElseIf objImgInfo.LoadState = lsLocal Then
'
'            mdtStartTime = Now
'            mblnIsBGReadProcessing = True
            
        End If
    Else
        '�жϱ����ļ��Ƿ����
        If objImgInfo.ImgCommand = icDownload Then
            If FileExists(objImgInfo.FilePath & objImgInfo.Filename) Then
                '�ж��Ƿ��ϴ�����ʧ���ļ����ϴ�ʧ���ļ�����ʧ��Ŀ¼��������ļ�
                strError = GetTransFailedState(objImgInfo.Key, icUpLoad)
                If Len(strError) > 0 Then
                    If InStr(strError, "��...") > 0 Then
                        objImgInfo.ImgCommand = icUpLoad
                        objImgInfo.LoadState = lsRedo
                    Else
                        objImgInfo.ImgCommand = icUpLoad
                        objImgInfo.LoadState = lsError
                        If Len(objImgInfo.ErrorInfo) <= 0 Then objImgInfo.ErrorInfo = "ͼ���ϴ�ʧ��>>" & Replace(strError, "ͼ���ϴ�ʧ��", "")
                    End If
                Else
                    objImgInfo.LoadState = lsLocal
                    objImgInfo.IsReDrawed = False
                    objImgInfo.ErrorInfo = ""
                    
                    Exit Function
                End If
            End If
        End If
        
        ProcessImgCmd = DirectProcessFtpFile(objImgInfo)
        If objImgInfo Is Nothing Then Exit Function
        
        If objImgInfo.ImgCommand = icUpLoad Then
            strFailedFile = GetImgCmdFailed(objImgInfo)
            
            If ProcessImgCmd = frNormal Then
                'ɾ��ʧ��Ŀ¼�ı���ļ�
                RemoveFile strFailedFile
            Else
                strTransErr = objImgInfo.ErrorInfo
                '��ʧ��Ŀ¼����һ��ͬ������ļ�
                Call TransCmd(objImgInfo, strFailedFile, strError)
                
                objImgInfo.LoadState = lsError
                objImgInfo.ErrorInfo = strTransErr
            End If
        End If
    End If
End Function


Private Function DirectProcessFtpFile(objImgInfo As clsBgImgInfo) As ftpResult
'����ftp�ļ�
    Dim ftpConTag As TFtpConTag
    Dim strFtpFile As String
    Dim strLocalFile As String
    
    If objImgInfo Is Nothing Then
        MsgboxH GetRootHwnd, "��ǰ���������Ч�����Զ���ֹ��", vbOKOnly
        DirectProcessFtpFile = frAbort
        Exit Function
    End If
    
    ftpConTag = FtpTagInstance(objImgInfo.FtpIp, objImgInfo.FtpUser, objImgInfo.FtpPwd, objImgInfo.FtpVirtualPath)
    
    strLocalFile = objImgInfo.FilePath & objImgInfo.Filename
    strFtpFile = objImgInfo.FtpFile
    
    If objImgInfo.ImgCommand = icUpLoad Then
        DirectProcessFtpFile = FtpUpload(ftpConTag, strFtpFile, strLocalFile)
        
        objImgInfo.IsReDrawed = False
        
        If DirectProcessFtpFile = frNormal Then
            objImgInfo.LoadState = lsLocal
        Else
            objImgInfo.LoadState = lsError
            objImgInfo.ErrorInfo = "ͼ���ϴ�ʧ�ܡ�"
        End If
    Else
        If DirExists(objImgInfo.FilePath) = False Then Call MkLocalDir(objImgInfo.FilePath)
        
        DirectProcessFtpFile = FtpDownload(ftpConTag, strFtpFile, strLocalFile)
        
        objImgInfo.IsReDrawed = False
        
        If DirectProcessFtpFile = frNormal Then
            objImgInfo.LoadState = lsLocal
        Else
            objImgInfo.LoadState = lsError
            objImgInfo.ErrorInfo = "ͼ������ʧ�ܡ�"
        End If
    End If
End Function


Private Sub AddImgToViewer(objDcmImg As DicomImage, objImgInfo As clsBgImgInfo)
    Dim strCmdFile As String
    Dim objIni As New clsIniFile
    Dim strErr As String
    Dim strCmd As String
    
    Call DrawBorder(objDcmImg, mlngSelColorStyle)
    
    If mblnIsShowCheck Then Call DrawCheckBox(objDcmImg, mlngSelColorStyle)
    If mblnIsDrawOrder Then Call DrawImgOrder(objDcmImg)
    If mblnIsDrawHint Then Call DrawHints(objDcmImg)
    
    '�ж��Ƿ���ڶ�Ӧ����ʧ�ܵ�����
    strCmdFile = GetImgCmdPath(True) & objDcmImg.InstanceUID
    If FileExists(strCmdFile) = False Then
        strCmdFile = GetImgCmdPath() & objDcmImg.InstanceUID
    End If
    
    If FileExists(strCmdFile) Then
        '����ʧ�ܵĴ���������ȡʧ����Ϣ����ʾ
        Call objIni.SetIniFile(strCmdFile)
        
        strErr = objIni.ReadValue("OTHERINFO", "ERRORINFO", "")
        strCmd = objIni.ReadValue("OTHERINFO", "IMGCOMMAND", "")
        
        If Val(strCmd) = 2 Then
            strCmd = "�ϴ�:"
            
        ElseIf Val(strCmd) = 1 Then
            If objImgInfo.LoadState = lsLocal And objImgInfo.ImgCommand = icDownload _
                And FileExists(objImgInfo.FilePath & objImgInfo.Filename) Then
            '����Ѿ��������ص����ص�ͼ����ֱ��ɾ��֮ǰ����ʧ�ܵ�����
                RemoveFile strCmdFile
                
                strErr = ""
            End If
            
            strCmd = "����:"
        Else
            strCmd = ""
        End If
        
        If Len(strErr) > 0 Then
            Call DrawErrorText(objDcmImg, strCmd & strErr)
        End If
    End If
    
    Call dcmViewer.Images.Add(objDcmImg)
End Sub




Private Sub ReDrawImage(imgInfo As clsBgImgInfo)
'ִ��ͼ������.�������أ����أ��ϴ�
On Error GoTo errhandle
    Dim strFile As String
    Dim objNewImg As DicomImage
    Dim strError As String
    Dim lngImgIndex As Long

    
    If imgInfo.IsReDrawed Then Exit Sub
    
'    '����ý���ļ�
'    If imgInfo.Format = ifAvi Or imgInfo.Format = ifWav Then
'        '������Ƶ����Ƶ���滻ͼ��ֻ���ھ��岥�ŵ�ʱ�򣬲����ض�Ӧ����Ƶ�ļ�
'        If imgInfo.Format = ifAvi Then
'            Set objNewImg = ReadMediaFile(sitAvi, strError)
'        Else
'            Set objNewImg = ReadMediaFile(sitWav, strError)
'        End If
'
''        imgInfo.LoadState = lsMedia
'        imgInfo.IsReDrawed = True
'
'        Set objNewImg.tag = imgInfo
'
'        objNewImg.InstanceUID = imgInfo.Key
'        Call AddImgToViewer(objNewImg)
'
'        Exit Sub
'    End If
'    Debug.Print GetTickCount & ":ReDrawImage --"

    imgInfo.IsReDrawed = True
    
    Select Case imgInfo.LoadState
        Case lsRedo, lsSent, lsError '������
            If imgInfo.ImgCommand <> icUpLoad Then
                Set objNewImg = ReadMediaFile(sitDown, strError)
            Else
                '������Ƶ����Ƶ���滻ͼ��ֻ���ھ��岥�ŵ�ʱ�򣬲����ض�Ӧ����Ƶ�ļ�
                If imgInfo.Format = ifAvi Then
                    Set objNewImg = ReadMediaFile(sitAvi, strError)
                ElseIf imgInfo.Format = ifWav Then
                    Set objNewImg = ReadMediaFile(sitWav, strError)
                Else
                    strFile = FormatFilePath(imgInfo.FilePath & "\" & imgInfo.Filename)
                    Set objNewImg = ReadDicomFile(strFile, strError, IIf(imgInfo.Format = ifDcm, True, False))
                End If
              
                
                If objNewImg Is Nothing Then
                    '�����ļ���ȡ����
                    If imgInfo.ErrorInfo <> "" Then imgInfo.ErrorInfo = strError
                    Set objNewImg = ReadMediaFile(sitErr, strError)
                End If
            End If
            
            objNewImg.InstanceUID = imgInfo.Key
            
            Set objNewImg.tag = imgInfo
            
            If imgInfo.LoadState = lsRedo Or imgInfo.LoadState = lsError Then
                Call DrawErrorInfo(objNewImg, imgInfo)
            End If
            
            Call AddImgToViewer(objNewImg, imgInfo)
            

            
        Case lsLocal
            '����ý���ļ�
            '������Ƶ����Ƶ���滻ͼ��ֻ���ھ��岥�ŵ�ʱ�򣬲����ض�Ӧ����Ƶ�ļ�
            If imgInfo.Format = ifAvi Then
                Set objNewImg = ReadMediaFile(sitAvi, strError)
            ElseIf imgInfo.Format = ifWav Then
                Set objNewImg = ReadMediaFile(sitWav, strError)
            Else
                strFile = FormatFilePath(imgInfo.FilePath & "\" & imgInfo.Filename)
                Set objNewImg = ReadDicomFile(strFile, strError, IIf(imgInfo.Format = ifDcm, True, False))
            End If

                
            If objNewImg Is Nothing Then
                '�����ļ���ȡ����
                imgInfo.LoadState = lsError
                
                '�����ļ���ȡʧ��
                imgInfo.ErrorInfo = strError
                Set objNewImg = ReadMediaFile(sitErr, strError)
                
                
                Set objNewImg.tag = imgInfo
                
                objNewImg.InstanceUID = imgInfo.Key
                
                Call DrawErrorInfo(objNewImg, imgInfo)
                Call AddImgToViewer(objNewImg, imgInfo)
            Else
                '��ȡ�ɹ�
                Set objNewImg.tag = imgInfo
                Call AddImgToViewer(objNewImg, imgInfo)
            End If
    End Select
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
'    Resume
End Sub

Private Function StartBackGroundServer(Optional ByVal blnForceStart As Boolean = False) As Boolean
'������̨����
    Dim strTimeoutFile As String
    Dim objIni As New clsIniFile
    Dim blnIsStart As Boolean
    Dim dtLastTime As Date
    Dim lngTimeoutCount As Long
    Dim strBgFile As String
    
On Error GoTo errhandle
    StartBackGroundServer = False
    
    strTimeoutFile = GetImgCmdPath
    
    If DirExists(strTimeoutFile) = False Then
        Call MkLocalDir(strTimeoutFile)
    End If
    
    strTimeoutFile = strTimeoutFile & "TimeOut.dat"
    blnIsStart = False
    
    If Trim(Dir(strTimeoutFile, vbHidden)) = "" Then
        '���û���ļ�����ֱ��������̨���̣���̨��������ʱ�������ȴ���timeout�ļ�
        blnIsStart = True '������̨����
    Else
'        '������ļ����жϺ�̨�����Ƿ���ڣ���������ڣ���ֱ������,������ڣ���ֱ���ж��Ƿ�ʱ
'        If HasProcess(strBgExe) = False Then
'            blnIsStart = True
'        Else
            '���ڽ������ȡtimeout�ļ�,�ж�����ʱ���Ƿ�ʱ
            objIni.SetIniFile strTimeoutFile
            dtLastTime = CDate(objIni.ReadValue("TIMEOUT", "value", 0))
            
            lngTimeoutCount = DateDiff("s", dtLastTime, Now)
            If lngTimeoutCount > 30 Then
                blnIsStart = True
            End If
            
'        End If
    End If
    
    If blnIsStart Or blnForceStart Then
        blnIsStart = False
        '������̨����
        strBgFile = IsExistsBGServer()
        
        If strBgFile <> "" Then
            Shell strBgFile & " " & GetImgCmdPath
            Call UpdateTimeout
            blnIsStart = True
        Else
            StartBackGroundServer = False
            Set objIni = Nothing
            
            Exit Function
        End If
    End If
    
    Set objIni = Nothing
    
    StartBackGroundServer = True
Exit Function
errhandle:
    Set objIni = Nothing
    StartBackGroundServer = False
End Function


Private Sub UpdateTimeout()
'���³�ʱʱ��
    Dim objIni As New clsIniFile
On Error GoTo errhandle

    Call objIni.SetIniFile(FormatFilePath(GetImgCmdPath & "\TimeOut.dat"))
    Call objIni.WriteValue("TIMEOUT", "value", Now)
    
    Set objIni = Nothing
Exit Sub
errhandle:
    Set objIni = Nothing
End Sub



Public Function ReadMediaFile(ByVal stateImgType As TStateImageType, _
    Optional ByRef strError As String) As DicomImage
    
    Dim strFile As String
    Dim strCurError As String
    
On Error GoTo errhandle
    strFile = ""
    strError = ""
    
    Set ReadMediaFile = Nothing
    
    If stateImgType = sitNul Then
        Set ReadMediaFile = ReadDicomFile("NULL", strCurError)
        Exit Function
    End If
    
    If stateImgType = sitAvi Then
        strFile = FormatFilePath(SysRootPath & "\AVI.BMP")
        If Trim(Dir(strFile, vbHidden)) = "" Then Call SavePicture(imgList.ListImages(2).Picture, strFile)
    End If
    
    If stateImgType = sitWav Then
        strFile = FormatFilePath(SysRootPath & "\WAV.BMP")
        If Trim(Dir(strFile, vbHidden)) = "" Then Call SavePicture(imgList.ListImages(3).Picture, strFile)
    End If
    
    If stateImgType = sitDown Then
        strFile = FormatFilePath(SysRootPath & "\DOWN.BMP")
        If Trim(Dir(strFile, vbHidden)) = "" Then Call SavePicture(imgList.ListImages(1).Picture, strFile)
    End If
    
    If stateImgType = sitErr Then
        strFile = FormatFilePath(SysRootPath & "\ERROR.BMP")
        If Trim(Dir(strFile, vbHidden)) = "" Then Call SavePicture(imgList.ListImages(4).Picture, strFile)
    End If
    
    If Len(strFile) <= 0 Then
        strError = "��Ч��״̬�ļ�."
        Exit Function
    End If
    
    Set ReadMediaFile = ReadDicomFile(strFile, strCurError)
    
    If ReadMediaFile Is Nothing Then
        Set ReadMediaFile = ReadDicomFile("NULL", strCurError)
    End If
Exit Function
errhandle:
    strError = err.Description
    Set ReadMediaFile = ReadDicomFile("NULL", strCurError)
    
    If strCurError <> "" Then strError = "1:" & strError & vbCrLf & "2:" & strCurError
End Function





Private Sub ConfigPage(Optional ByVal blnIsRefresh = True)
'���÷�ҳ
    Dim lngTotalCount As Long
    Dim i As Long
    
On Error GoTo errhandle
    mblnIsPageConfig = True

    lngTotalCount = ImgCount 'UBound(maryImgInfo) + 1
    cbxPage.Clear
    
    If lngTotalCount <= 0 Then
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Call cbxPage.AddItem("0/0")
        
        Exit Sub
    End If
    
    If (lngTotalCount Mod mlngPageRecord) <> 0 Then
        mlngPageCount = Int(lngTotalCount / mlngPageRecord) + 1
    Else
        mlngPageCount = lngTotalCount / mlngPageRecord
    End If
    
    For i = 1 To mlngPageCount
        Call cbxPage.AddItem("�� " & i & "/" & mlngPageCount & " ҳ")
    Next
    
    If blnIsRefresh Then
        mlngPageIndex = 1
        cbxPage.ListIndex = 0
    Else
        If mlngPageIndex > cbxPage.ListCount Then
            cbxPage.ListIndex = cbxPage.ListCount - 1
        Else
            If cbxPage.ListCount > 0 Then
                cbxPage.ListIndex = mlngPageIndex - 1
            End If
            
        End If
        
        mlngPageIndex = cbxPage.ListIndex + 1
    End If
    
    
    '����dcmViewer����ʾ�к���
    Call ResizeRegion(mlngPageRecord, dcmViewer.Width, dcmViewer.Height, mlngViewRows, mlngViewCols)
    
    dcmViewer.MultiColumns = mlngViewCols
    dcmViewer.MultiRows = mlngViewRows
    
    mblnIsPageConfig = False
Exit Sub
errhandle:
    mblnIsPageConfig = False
End Sub

Public Sub ProxyTransfer(objImgInfo As clsBgImgInfo)
    Call ProcessImgCmd(objImgInfo, , Not mblnBGServerStarted)
End Sub

Public Sub AddImg(objImgInfo As clsBgImgInfo, Optional ByVal blnInsertFirst As Boolean = False)
'���ͼ��
    Dim blnIsSend As Boolean
    Dim lngBound As Long
    Dim strError As String
    Dim lngOrder As Long
    Dim i As Long
    Dim objProxyImgInfo As clsBgImgInfo
    
    blnIsSend = False
    
    lngBound = ImgCount
    ReDim Preserve maryImgInfo(lngBound)
    
    If blnInsertFirst = False Then
        Set maryImgInfo(lngBound) = objImgInfo
        
        If maryImgInfo(lngBound).ImageOrder <= 0 Then
            maryImgInfo(lngBound).ImageOrder = lngBound + 1
        End If
    Else
        '��������һ��λ�ã�˵���ǵ���
        For i = lngBound To 1 Step -1
            Set maryImgInfo(i) = maryImgInfo(i - 1).CopyNew()
            maryImgInfo(i).ImageOrder = i + 1
        Next
        
        Set maryImgInfo(0) = objImgInfo
        lngBound = 0
        
        If maryImgInfo(lngBound).ImageOrder <= 0 Then
            maryImgInfo(lngBound).ImageOrder = 1
        End If
    End If
     
'    If maryImgInfo(lngBound).DrawOrder <= 0 Then
'        maryImgInfo(lngBound).DrawOrder = lngBound + 1
'    End If
            
    Set objProxyImgInfo = maryImgInfo(lngBound).CopyNew
    mdtStartTime = Now
    
    '������ͼ��
    If maryImgInfo(lngBound).LoadState = lsLocal Then  '���ָ�����Ǽ��ر���ͼ�񣬱���ͬ�����ͼ����ʾ������Ҫ���������
        mblnIsBGReadProcessing = True
    Else
        Call ProcessImgCmd(maryImgInfo(lngBound), , Not mblnBGServerStarted)
    End If
     
     If maryImgInfo(lngBound) Is Nothing Then
        '����������л����ʱ�������maryImgInfo�������Ϊnothing,
        Call ProxyTransfer(objProxyImgInfo)
        Exit Sub
    Else
        If objProxyImgInfo.AdviceId <> maryImgInfo(lngBound).AdviceId Then
            Call ProxyTransfer(objProxyImgInfo)
            Exit Sub
        End If
     End If
     
'    ���÷�ҳ��Ϣ
    Call ConfigPage(False)
       
    If mblnIsBGReadProcessing Or Not mblnBGServerStarted Then
        LockUpdateEx dcmViewer.hwnd
        
        If dcmViewer.Images.Count < dcmViewer.MultiRows * dcmViewer.MultiColumns Then
'            lngOrder = objImgInfo.DrawOrder
'
'            objImgInfo.DrawOrder = 0
            Call ReDrawImage(objImgInfo)
            
'            objImgInfo.DrawOrder = lngOrder
        Else
            If dcmViewer.Images.Count > 0 Then Call dcmViewer.Images.Remove(1)
            
            lngOrder = objImgInfo.ImageOrder
            
            objImgInfo.ImageOrder = 0
            Call ReDrawImage(objImgInfo)
            
            objImgInfo.ImageOrder = lngOrder
            
            Call dcmViewer.Images.Move(dcmViewer.Images.Count, 1)
        End If
        
        LockUpdateEx 0
        
        If mblnBGServerStarted Then timerState.Enabled = True
    End If
End Sub

Public Sub ClearChecked()
'���ѡ��
    Dim i As Long
    
    If mblnIsShowCheck = False Then Exit Sub
    
    For i = 1 To dcmViewer.Images.Count
        Call DrawCheckBox(dcmViewer.Images(i), False)
    Next i
End Sub



Public Sub SelectedAll()
'ȫѡ
    Dim i As Long
    
    If mblnIsShowCheck = False Then Exit Sub
    
    For i = 1 To dcmViewer.Images.Count
        Call DrawCheckBox(dcmViewer.Images(i), ColorConstants.vbRed, True)
    Next i
End Sub

Public Sub Selected(ByVal lngIndex As Long)
'ѡ��ͼ��
    If dcmViewer.Images.Count <= 0 Then
        mlngSelectIndex = 0
        Exit Sub
    End If
    
    If lngIndex > dcmViewer.Images.Count Then
        lngIndex = lngIndex - 1
        If lngIndex > dcmViewer.Images.Count Then Exit Sub
    End If
    
    Call DrawBorder(dcmViewer.Images(lngIndex), mlngSelColorStyle, True)
    
    RaiseEvent OnClick(lngIndex)
End Sub



Public Sub DelImgView(Optional ByVal lngSelIndex As Long = 0)
'ɾ��ͼ��
'���lngSelIndexΪ-1�����ʾɾ����ǰѡ���ͼ��
    Dim i As Long
    Dim strKey As String
    Dim arySelIndex() As Long
    Dim objImg As DicomImage
    
    If lngSelIndex > dcmViewer.Images.Count Then Exit Sub
    
    If lngSelIndex <= 0 Then
        'ɾ��ѡ���ͼ��
        arySelIndex = GetSelects()
        
        For i = UBound(arySelIndex) To 1 Step -1
            Set objImg = dcmViewer.Images(arySelIndex(i))
            
            If Not objImg Is Nothing Then
                strKey = objImg.tag.Key
                
                Call RemoveImgInfo(strKey)
                Call RemoveImgCmdFile(strKey)
                
                Set dcmViewer.Images(arySelIndex(i)).tag = Nothing
                Call dcmViewer.Images.Remove(arySelIndex(i))
            End If
        Next
        
    Else
        'ɾ��ָ����ͼ��
        strKey = dcmViewer.Images(lngSelIndex).tag.Key
        
        '�Ƴ�����Ԫ��
        Call RemoveImgInfo(strKey)
        Call RemoveImgCmdFile(strKey)
        
        '�Ƴ�����Ԥ��
        Set dcmViewer.Images(lngSelIndex).tag = Nothing
        Call dcmViewer.Images.Remove(lngSelIndex)
    End If
        
    dcmViewer.Refresh
    
    Call ConfigPage(False)
    
    If dcmViewer.Images.Count <= 0 Then
        Call cbxPage_Click
    End If
End Sub

Public Function GetImage(ByVal lngIndex As Long, Optional ByRef objImgInfo As clsBgImgInfo = Nothing) As DicomImage
'��ȡͼ��
    Dim objSelImg As DicomImage
    Dim strError As String
    
    Set GetImage = Nothing
    
    If lngIndex <= 0 Or lngIndex > dcmViewer.Images.Count Then Exit Function
    
    Set objSelImg = dcmViewer.Images(lngIndex)
    Set objImgInfo = objSelImg.tag.CopyNew
    
    '�ж��Ƿ��ϴ�������ʧ�ܵ��ļ�
    strError = GetTransFailedState(objSelImg.InstanceUID)
    If Len(strError) > 0 Then
        objImgInfo.LoadState = lsError
        
        objImgInfo.ErrorInfo = strError
    End If
    
    
    Set GetImage = objSelImg.SubImage(0, 0, objSelImg.SizeX, objSelImg.SizeY, 1, 1)
    
    GetImage.InstanceUID = objSelImg.InstanceUID
End Function

Public Sub Redo(Optional ByVal lngSelIndex As Long = 0)
'lngIndex���С��0�����ѡ�������ͼ���������
    Dim i As Long
    Dim strKey As String
    Dim arySelIndex() As Long
    Dim objImgInfo As clsBgImgInfo
    Dim blnIsAbort As Boolean
    
    If lngSelIndex > dcmViewer.Images.Count Then Exit Sub
    
    If lngSelIndex <= 0 Then
        arySelIndex = GetSelects()
        
        blnIsAbort = False
        For i = UBound(arySelIndex) To 1 Step -1
            Set objImgInfo = dcmViewer.Images(arySelIndex(i)).tag
            objImgInfo.IsReDrawed = False
            
            If blnIsAbort Then
                objImgInfo.LoadState = lsError
                objImgInfo.ErrorInfo = "����ֹ����"
            Else
                Call ProcessImgCmd(objImgInfo, True, Not mblnBGServerStarted)
            End If
            
            If Not mblnBGServerStarted Then Call DrawImgState(arySelIndex(i), True)
        Next
        
    Else
        Set objImgInfo = dcmViewer.Images(lngSelIndex).tag
        objImgInfo.IsReDrawed = False
        
        Call ProcessImgCmd(objImgInfo, True, Not mblnBGServerStarted)
        If Not mblnBGServerStarted Then Call DrawImgState(lngSelIndex, True)
    End If
    
    If mblnIsBGReadProcessing Then
        Call DrawResultState(rsClear)
        timerState.Enabled = True
    End If
End Sub




Public Sub ReDown(Optional ByVal lngSelIndex As Long = 0)
'��������
'ע���ϴ�ʧ�ܵ�ͼ�������������
    Dim i As Long
    Dim strKey As String
    Dim arySelIndex() As Long
    Dim objImgInfo As clsBgImgInfo
    Dim lngDownCount As Long
    Dim strFailedFile As String
    Dim blnIsAbort As Boolean
    
    If lngSelIndex > dcmViewer.Images.Count Then Exit Sub
    
    lngDownCount = 0
    
    If lngSelIndex <= 0 Then
        arySelIndex = GetSelects()
        If UBound(arySelIndex) = 1 Then
            lngSelIndex = arySelIndex(1)
        End If
    End If
    
    If lngSelIndex <= 0 Then
        
        If UBound(arySelIndex) <= 0 Then
            MsgboxH GetRootHwnd, "��ѡ����Ҫ�������ص�ͼ��", vbOKOnly, "��ʾ"
            Exit Sub
        End If
        
        '�ж��Ƿ��ϴ�ʧ�ܵ�ͼ��
        
        If MsgboxH(GetRootHwnd, "�������ؽ���ɾ������ͼ���Ƿ������", vbYesNo, "��ʾ") = vbNo Then Exit Sub
        
        blnIsAbort = False
        For i = UBound(arySelIndex) To 1 Step -1
            Set objImgInfo = dcmViewer.Images(arySelIndex(i)).tag
            
            '��Ϊ�ϴ�ʧ�ܵ�ͼ�񣬲��ܽ�����������
            If objImgInfo.ImgCommand <> icUpLoad And GetTransFailedState(objImgInfo.Key, icUpLoad) = "" Then
                lngDownCount = lngDownCount + 1
                
                objImgInfo.ImgCommand = icDownload
                objImgInfo.IsReDrawed = False
                
                If blnIsAbort Then
                    objImgInfo.LoadState = lsError
                    objImgInfo.ErrorInfo = "����ֹ����"
                Else
                    '��ɾ���Ѿ�����ʧ�ܵ�����
                    strFailedFile = GetImgCmdFailed(objImgInfo)
                    If FileExists(strFailedFile) Then
                        RemoveFile strFailedFile
                    End If
                    
                    RemoveFile objImgInfo.FilePath & objImgInfo.Filename
                    
                    If ProcessImgCmd(objImgInfo, True, Not mblnBGServerStarted) = frAbort Then blnIsAbort = True
                End If
                If Not mblnBGServerStarted Then Call DrawImgState(arySelIndex(i), True)
                
            End If
        Next
        
    Else
        Set objImgInfo = dcmViewer.Images(lngSelIndex).tag
        
        If GetTransFailedState(objImgInfo.Key, icUpLoad) <> "" Then
             Call MsgboxH(GetRootHwnd, "��ǰͼ��δ�ϴ��ɹ��������������ء�", vbOKOnly, "��ʾ")
             Exit Sub
        End If
        
        If MsgboxH(GetRootHwnd, "�������ؽ���ɾ������ͼ���Ƿ������", vbYesNo, "��ʾ") = vbNo Then Exit Sub
        
        
        If objImgInfo.ImgCommand <> icUpLoad Then
            lngDownCount = lngDownCount + 1
            
            objImgInfo.ImgCommand = icDownload
            objImgInfo.IsReDrawed = False
            
            '��ɾ���Ѿ�����ʧ�ܵ�����
            strFailedFile = GetImgCmdFailed(objImgInfo)
            If FileExists(strFailedFile) Then
                RemoveFile strFailedFile
            End If
                
            RemoveFile objImgInfo.FilePath & objImgInfo.Filename
            
            Call ProcessImgCmd(objImgInfo, True, Not mblnBGServerStarted)
            If Not mblnBGServerStarted Then Call DrawImgState(lngSelIndex, True)
            
        End If
    End If
    
    If lngDownCount <= 0 Then
        MsgboxH GetRootHwnd, "δ���ֿ������������ص�ͼ��", vbOKOnly, "��ʾ"
    End If
    
    If mblnIsBGReadProcessing Then
        Call DrawResultState(rsClear)
        timerState.Enabled = True
    End If
End Sub


Private Function GetTransFailedState(ByVal strInstanceUID As String, Optional ByVal imgType As TImageCommand = icReadly) As String
'�ж��Ƿ���ʧ�ܵ�ͼ��
    Dim strFailedFile As String
    Dim objIni As New clsIniFile
    Dim strError As String
    Dim strCmd As String
    Dim strEndTime As String
    
    GetTransFailedState = ""
    
    strFailedFile = GetImgCmdPath(True) & strInstanceUID
    If FileExists(strFailedFile) = False Then
        strFailedFile = GetImgCmdPath() & strInstanceUID
    End If
    
    If FileExists(strFailedFile) Then
        objIni.SetIniFile strFailedFile
        strCmd = objIni.ReadValue("OTHERINFO", "IMGCOMMAND", "")
        
        If imgType <> icReadly Then
            If Val(strCmd) <> imgType Then Exit Function
        End If
        
        '�ж��Ƿ��������ط������ϴ���û�м�ʱ���ļ�������������ظ���ȡ�������Ҫ�жϴ�����
        strEndTime = objIni.ReadValue("OTHERINFO", "STARTTIME", Now)
        If Now - (3 / 24 / 60 / 60) < CDate(strEndTime) Then
            GetTransFailedState = "ͼ������..."
            Exit Function
        End If
        
        GetTransFailedState = objIni.ReadValue("OTHERINFO", "ERRORINFO", "")
        
        If Val(strCmd) = 2 Then
                If Len(GetTransFailedState) <= 0 Then GetTransFailedState = "ͼ���ϴ���..."
        Else
                If Len(GetTransFailedState) <= 0 Then GetTransFailedState = "ͼ��������..."
        End If
    End If
End Function


Public Sub ReUp(Optional ByVal lngSelIndex As Long = 0)
'�����ϴ�
    Dim i As Long
    Dim strKey As String
    Dim arySelIndex() As Long
    Dim objImgInfo As clsBgImgInfo
    Dim lngUpCount As Long
    Dim strFailedFile As String
    Dim blnIsAbort As Boolean
    
    If lngSelIndex > dcmViewer.Images.Count Then Exit Sub
    
    lngUpCount = 0
    
    If lngSelIndex <= 0 Then
        arySelIndex = GetSelects()
        
        If UBound(arySelIndex) = 1 Then
            lngSelIndex = arySelIndex(1)
        End If
    End If
    
    If lngSelIndex <= 0 Then
        
        If UBound(arySelIndex) <= 0 Then
            MsgboxH GetRootHwnd, "��ѡ����Ҫ�����ϴ���ͼ��", vbOKOnly, "��ʾ"
            Exit Sub
        End If
        
        blnIsAbort = False
        For i = UBound(arySelIndex) To 1 Step -1
            Set objImgInfo = dcmViewer.Images(arySelIndex(i)).tag
            
            If (objImgInfo.LoadState = lsLocal Or (objImgInfo.LoadState = lsError And FileExists(objImgInfo.FilePath & objImgInfo.Filename))) And GetTransFailedState(objImgInfo.Key, icDownload) = "" Then
                lngUpCount = lngUpCount + 1
                
                objImgInfo.ImgCommand = icUpLoad
                objImgInfo.IsReDrawed = False
                
                If blnIsAbort Then
                    objImgInfo.LoadState = lsError
                    objImgInfo.ErrorInfo = "����ֹ����"
                Else
                    '��ɾ���Ѿ�����ʧ�ܵ�����
                    strFailedFile = GetImgCmdFailed(objImgInfo)
                    If FileExists(strFailedFile) Then
                        RemoveFile strFailedFile
                    End If
                    
                    If ProcessImgCmd(objImgInfo, True, Not mblnBGServerStarted) = frAbort Then blnIsAbort = True
                End If
                
                If Not mblnBGServerStarted Then Call DrawImgState(arySelIndex(i), True)
            End If
        Next
        
    Else
        Set objImgInfo = dcmViewer.Images(lngSelIndex).tag
        
        If GetTransFailedState(objImgInfo.Key, icDownload) <> "" Then
            Call MsgboxH(GetRootHwnd, "��ǰͼ��δ���سɹ������������ϴ���", vbOKOnly, "��ʾ")
            Exit Sub
        End If
        
        If objImgInfo.LoadState = lsLocal Or (objImgInfo.LoadState = lsError And FileExists(objImgInfo.FilePath & objImgInfo.Filename)) Then
            lngUpCount = lngUpCount + 1
            
            objImgInfo.ImgCommand = icUpLoad
            objImgInfo.IsReDrawed = False
            
            '��ɾ���Ѿ�����ʧ�ܵ�����
            strFailedFile = GetImgCmdFailed(objImgInfo)
            If FileExists(strFailedFile) Then
                RemoveFile strFailedFile
            End If
                
            Call ProcessImgCmd(objImgInfo, True, Not mblnBGServerStarted)
            
            If Not mblnBGServerStarted Then Call DrawImgState(lngSelIndex, True)
        End If
    End If
    
    If lngUpCount <= 0 Then
        MsgboxH GetRootHwnd, "δ���ֿ����������ϴ���ͼ��", vbOKOnly, "��ʾ"
    End If
    
    If mblnIsBGReadProcessing Then
        Call DrawResultState(rsClear)
        timerState.Enabled = True
    End If
End Sub

Private Sub RemoveImgInfo(ByVal strKey As String)
'�Ƴ�ͼ��������Ϣ
    Dim i As Long
    Dim objImgInfo As clsBgImgInfo
    Dim lngIndex As Long
    Dim lngBound As Long
    
    If ImgCount <= 0 Then Exit Sub
    
    lngIndex = -1
    lngBound = UBound(maryImgInfo)
    For i = 0 To lngBound
        Set objImgInfo = maryImgInfo(i)
        
        If objImgInfo.Key = strKey Then
            lngIndex = i
            Exit For
        End If
    Next
    
    If lngIndex >= 0 Then
        For i = lngIndex + 1 To UBound(maryImgInfo)
            Set maryImgInfo(i - 1) = maryImgInfo(i).CopyNew()
        Next
        
        If lngBound > 0 Then
            ReDim Preserve maryImgInfo(lngBound - 1)
        Else
            EraseAry maryImgInfo
        End If
    End If
    
End Sub

Private Sub RemoveImgCmdFile(ByVal strKey As String)
'�Ƴ������ļ�
    Dim strCmdFile As String
On Error GoTo errhandle:
    strCmdFile = FormatFilePath(GetImgCmdPath & strKey)
    Call RemoveFile(strCmdFile)
    
    strCmdFile = FormatFilePath(GetImgCmdPath(True) & strKey)
    Call RemoveFile(strCmdFile)
    
Exit Sub
errhandle:
    
End Sub


Public Function GetSelects() As Long()
'��ȡѡ�е�ͼ������
'������1��ʼ
    Dim i As Long
    Dim lngBound As Long
    Dim arySelIndex() As Long
    
    ReDim arySelIndex(0)
    
    For i = 1 To dcmViewer.Images.Count
        
        If IsSelected(i) Then
            '����Ƿ�͸����ɫ,˵���Ǳ�ѡ�е�ͼ��
            lngBound = UBound(arySelIndex) + 1
            ReDim Preserve arySelIndex(lngBound)
            
            arySelIndex(lngBound) = i
        End If
    Next
    
    GetSelects = arySelIndex
End Function

Public Function IsSelected(ByVal lngIndex As Long) As Boolean
'�ж��Ƿ�ѡ��
    Dim objImg As DicomImage
    Dim objLab As DicomLabel
    Dim i As Long
    
    IsSelected = False
    
    If lngIndex <= 0 Or lngIndex > dcmViewer.Images.Count Then Exit Function
    
    Set objImg = dcmViewer.Images(lngIndex)
    
    For i = 1 To objImg.Labels.Count
        Set objLab = objImg.Labels(i)
        
        If objLab.tag = IMG_LAB_CHECKBOX_TAG Then
            If objLab.Transparent = False Then
                IsSelected = True
                Exit Function
            End If
        End If
    Next
    
    If objImg.BorderColour <> IMG_BACK_BORDER_COLOR Then
        IsSelected = True
    End If
    
End Function


Public Function IsChecked(ByVal lngIndex As Long) As Boolean
'�ж��Ƿ�ѡ��
    Dim objImg As DicomImage
    Dim objLab As DicomLabel
    Dim i As Long
    
    IsChecked = False
    
    If lngIndex <= 0 Or lngIndex > dcmViewer.Images.Count Then Exit Function
    
    Set objImg = dcmViewer.Images(lngIndex)
    
    For i = 1 To objImg.Labels.Count
        Set objLab = objImg.Labels(i)
        
        If objLab.tag = IMG_LAB_CHECKBOX_TAG Then
            If objLab.Transparent = False Then
                IsChecked = True
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsErrorImg(ByVal lngIndex As Long) As Boolean
'�ж��Ƿ����ͼ��
    Dim objImgInfo As clsBgImgInfo
    
    IsErrorImg = False
    
    If lngIndex <= 0 Or lngIndex > dcmViewer.Images.Count Then Exit Function
    
    Set objImgInfo = dcmViewer.Images(lngIndex).tag
    
    If objImgInfo Is Nothing Then Exit Function
    
    If objImgInfo.LoadState = lsError Then IsErrorImg = True
End Function


Public Sub ClearImgView()
'���ͼ��
    dcmViewer.Images.Clear
End Sub


Public Sub ClearDrawHint(ByVal strSourState As String)
'���������ʾ
    Dim i As Long
    
    For i = i To ImgCount - 1
        maryImgInfo(i).DrawHint = Replace(maryImgInfo(i).DrawHint, strSourState, "")
    Next
    
    Call ReDrawImages(mlngPageIndex, True)
    
End Sub

Public Sub ImgDrawHint(ByVal strImgKey As String, ByVal strState As String, Optional ByVal strClear As String = "")
    Dim i As Long
    Dim j As Long
    
    For i = 0 To ImgCount - 1
        If strImgKey = maryImgInfo(i).Key Then
             
            For j = 1 To Len(strClear)
                maryImgInfo(i).DrawHint = Replace(maryImgInfo(i).DrawHint, Mid(strClear, j, 1), "")
            Next
            
            For j = 1 To Len(strClear)
                maryImgInfo(i).DrawHint = Replace(maryImgInfo(i).DrawHint, Mid(strState, j, 1), "")
            Next
            
            maryImgInfo(i).DrawHint = strState & maryImgInfo(i).DrawHint
            
            Exit For
        End If
    Next
    
    Call ReDrawImages(mlngPageIndex, True)

End Sub

Public Sub SyncDrawHint(ByVal strImgKeys As String, ByVal strState As String, Optional ByVal strClear As String = "")
'ͬ��������ʾ
    Dim i As Long
    Dim j As Long
    
    For i = i To ImgCount - 1
        
        For j = 1 To Len(strClear)
            maryImgInfo(i).DrawHint = Replace(maryImgInfo(i).DrawHint, Mid(strClear, j, 1), "")
        Next
        
        If InStr(strImgKeys, maryImgInfo(i).Key) >= 1 Then
            For j = 1 To Len(strState)
                maryImgInfo(i).DrawHint = Replace(maryImgInfo(i).DrawHint, Mid(strState, j, 1), "") '��Ҫ��֮ǰ��״̬��գ�������ʾ�ظ���״̬�ı�
            Next
            
            maryImgInfo(i).DrawHint = strState & maryImgInfo(i).DrawHint
        End If
    Next
    
    Call ReDrawImages(mlngPageIndex, True)
End Sub


Public Sub ClearAll()
    Dim i As Long
    
    dcmViewer.Images.Clear
    
    
    For i = 0 To ImgCount - 1
        Set maryImgInfo(i) = Nothing
    Next
     
    EraseAry maryImgInfo
    
    timerState.Enabled = False
    
    Call ConfigPage(True)
    
    DrawRunState True
End Sub


 
Private Sub WaitUnlock()
On Error GoTo errhandle
    Dim i As Long
    
    i = 0
    While True
        If mblnIsTimerWorking = False Then Exit Sub
        
        i = i + 1
        If i > 300 Then
            Exit Sub
        End If
        
        Sleep 10
        
        DoEvents
    Wend
Exit Sub
errhandle:
    
End Sub

Private Sub cbxPage_Click()
On Error GoTo errhandle
    If mblnIsRefreshing Then Exit Sub
    If mblnIsPageConfig Then Exit Sub
    
    If cbxPage.ListCount <= 0 Then Exit Sub
    
    mlngPageIndex = cbxPage.ListIndex + 1
    
    Call ReDrawImages(mlngPageIndex, True)
    
'    If mblnIsProcessing = False Then
'        Call DrawImgStates(Not mblnBGServerStarted)
'    End If
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
End Sub


'Private Sub cmdRefresh_Click()
'On Error GoTo errhandle
'    RaiseEvent OnCmdEvent("REFRESH")
'Exit Sub
'errhandle:
'    MsgBoxH hWnd, err.Description, vbOKOnly, "��ʾ"
'End Sub

Private Sub dcmViewer_Click()
On Error GoTo errhandle
    If mblnIsClickEvent = False Then Exit Sub
    
    RaiseEvent OnClick(mlngSelectIndex)
Exit Sub
errhandle:
End Sub

Private Sub dcmViewer_DblClick()
On Error GoTo errhandle
    If mblnIsClickEvent = False Then Exit Sub
    
    RaiseEvent OnDbClick(mlngSelectIndex)
Exit Sub
errhandle:
End Sub

Public Sub FullView()
On Error GoTo errhandle
    If dcmViewer.MultiColumns = 1 And dcmViewer.MultiRows = 1 Then
        dcmViewer.MultiColumns = mlngViewCols
        dcmViewer.MultiRows = mlngViewRows
        
        dcmViewer.CurrentIndex = 1
    Else
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        dcmViewer.CurrentIndex = mlngSelectIndex
    End If
    
    
Exit Sub
errhandle:

End Sub
 

Private Sub dcmViewer_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle
    Dim lngSelectIndex As Long
    Dim i As Long
    
    If dcmViewer.Images.Count <= 0 Then Exit Sub
    
    Select Case KeyCode
        Case 36     'home
            If cbxPage.ListCount > 0 Then cbxPage.ListIndex = 0
            Exit Sub
        Case 35  'end
            If cbxPage.ListCount > 0 Then cbxPage.ListIndex = cbxPage.ListCount - 1
            Exit Sub
        Case 33  'pageup
            If cbxPage.ListIndex - 1 >= 0 Then cbxPage.ListIndex = cbxPage.ListIndex - 1
            Exit Sub
        Case 34   'pagedown
            If cbxPage.ListIndex + 1 < cbxPage.ListCount Then
                cbxPage.ListIndex = cbxPage.ListIndex + 1
            End If
            Exit Sub
        Case 37         '�������
            lngSelectIndex = mlngSelectIndex - 1
            If lngSelectIndex <= 0 Then Exit Sub
        Case 38    '�Ϲ���
            lngSelectIndex = mlngSelectIndex - dcmViewer.MultiColumns
            If lngSelectIndex <= 0 Then Exit Sub
        Case 39      '�ҹ���
            lngSelectIndex = mlngSelectIndex + 1
            If lngSelectIndex > dcmViewer.Images.Count Then Exit Sub
        Case 40      '�¹���
            lngSelectIndex = mlngSelectIndex + dcmViewer.MultiColumns
            If lngSelectIndex > dcmViewer.Images.Count Then Exit Sub
        Case 32
            '�ո���
            If mlngSelectIndex > 0 And mblnIsShowCheck Then
                If IsChecked(mlngSelectIndex) Then
                    Call DrawCheckBox(dcmViewer.Images(mlngSelectIndex), mlngSelColorStyle, False)
                Else
                    Call DrawCheckBox(dcmViewer.Images(mlngSelectIndex), mlngSelColorStyle, True)
                End If
            End If
            
            Exit Sub
        Case 65
            If Shift = 2 Then
                Call SelectedAll
                Exit Sub
            End If
            
        Case Else
            Exit Sub
    End Select
    
    For i = 1 To dcmViewer.Images.Count
        Call DrawBorder(dcmViewer.Images(i), mlngSelColorStyle)
    Next
        
    If lngSelectIndex > 0 Then
        Call DrawBorder(dcmViewer.Images(lngSelectIndex), mlngSelColorStyle, True)
    End If
    
    mlngSelectIndex = lngSelectIndex
    
    RaiseEvent OnClick(mlngSelectIndex)
Exit Sub
errhandle:
    Call MsgboxH(GetRootHwnd, err.Description, vbOKOnly, "��ʾ")
End Sub

Private Sub dcmViewer_LostFocus()
On Error GoTo errhandle
    mblnIsClickEvent = True
Exit Sub
errhandle:

End Sub

Private Sub dcmViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle
    Dim i As Long
    Dim objLabs As DicomLabels
    
    mblnIsClickEvent = True
    
    If Button = 2 Then mblnIsClickEvent = False '����Ҽ�������click�¼�
    
    
    Set objLabs = dcmViewer.LabelHits(X, Y, False, True, True)
    For i = 1 To objLabs.Count
        If objLabs(i).tag = "CHECKBOX" And objLabs(i).Visible Then
            
            mblnIsClickEvent = False
            Exit For
        End If
    Next
    
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
Exit Sub
errhandle:

End Sub

Private Sub dcmViewer_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle

    If dcmViewer.Images.Count <= 0 Then Exit Sub
    
    If mblnIsClickEvent = False Then
'        dcmViewer.SetFocus
        Exit Sub
    End If
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
Exit Sub
errhandle:

End Sub

Private Sub dcmViewer_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Long
    Dim arySelIndex() As Long
    Dim blnIsSelect As Boolean
    
On Error GoTo errhandle
  
    blnIsSelect = False
    If Button = 2 Then
        arySelIndex = GetSelects()
        
        If UBound(arySelIndex) > 1 Then blnIsSelect = True
    End If
    
    If Not blnIsSelect Then
        mlngSelectIndex = dcmViewer.ImageIndex(X, Y)
        
        If mblnIsShowCheck Then
            If UpdateCheckBox(X, Y) = False Then
                'û�и������Ͻǵ�checkbox
                For i = 1 To dcmViewer.Images.Count
                    Call DrawCheckBox(dcmViewer.Images(i), mlngSelColorStyle)
                Next
            End If
        End If
        
        If Shift <> 2 Then   '2��ʾctrl������
            For i = 1 To dcmViewer.Images.Count
                Call DrawBorder(dcmViewer.Images(i), mlngSelColorStyle)
            Next
        End If
        
        If mlngSelectIndex > 0 Then
            Call DrawBorder(dcmViewer.Images(mlngSelectIndex), mlngSelColorStyle, True)
        End If
    End If
    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
End Sub

Private Function UpdateCheckBox(ByVal X As Long, ByVal Y As Long) As Boolean
    Dim i As Long
    Dim lngImgIndex As Long
    Dim objLabs As DicomLabels
    
On Error GoTo errhandle
    Set objLabs = dcmViewer.LabelHits(X, Y, False, True, True)
    lngImgIndex = dcmViewer.ImageIndex(X, Y)
    
    UpdateCheckBox = False
    
    For i = 1 To objLabs.Count
        If objLabs(i).tag = "CHECKBOX" And objLabs(i).Visible Then
            '��objLabs(i).Visible=false��˵��ѡ�п��Ѿ������أ�����ѡ�д���
            objLabs(i).Transparent = Not objLabs(i).Transparent

            Call dcmViewer.Images(lngImgIndex).Refresh(False)
            
            UpdateCheckBox = True
            
            Exit For
        End If
    Next
Exit Function
errhandle:

End Function

Private Sub dcmViewer_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error GoTo errhandle
    If Delta > 0 Then
        If cbxPage.ListIndex <= 0 Then Exit Sub
        cbxPage.ListIndex = cbxPage.ListIndex - 1
    Else
        If cbxPage.ListIndex >= cbxPage.ListCount - 1 Then Exit Sub
        cbxPage.ListIndex = cbxPage.ListIndex + 1
    End If
Exit Sub
errhandle:
    Call MsgboxH(GetRootHwnd, err.Description, vbOKOnly, "��ʾ")
End Sub

Private Sub timerState_Timer()
On Error GoTo errhandle
    Dim strPath As String
    Dim objFileSys As New FileSystemObject
    Dim lngErrorCount As Long
    Dim lngProcessCount As Long
    
    mblnIsTimerWorking = True
    
    If mblnIsBGReadProcessing = False _
        Or timerState.Enabled = False Then
        mblnIsTimerWorking = False
        Exit Sub
    End If
    
    mlngServerTime = mlngServerTime + 1

    '�������Ƿ�����
    '�ж�transcmdĿ¼���Ƿ����ļ�
    If mlngServerTime >= VALIDTIME Then
        mlngServerTime = 0

        strPath = GetImgCmdPath()

        If objFileSys.GetFolder(strPath).Files.Count > 0 Then
            mblnBGServerStarted = StartBackGroundServer
        End If
    End If

    If timerState.Enabled = False Then
        mblnIsTimerWorking = False
        Exit Sub
    End If
    
    lngProcessCount = DetectionImgProcess(lngErrorCount)
    If lngProcessCount > 0 Then
        labState.Caption = String(3 - Len("" & lngProcessCount), "0") & lngProcessCount
        Call DrawRunState
    Else
        Call DrawRunState(True)
        Call DrawResultState(IIf(lngErrorCount > 0, rsFailed, rsOk))

        mblnIsBGReadProcessing = False
    End If

    If timerState.Enabled = False Then
        mblnIsTimerWorking = False
        Exit Sub
    End If
    
    'ˢ�¼�����ʾ��ͼ��״̬
    Call DrawImgStates
     
    
    mblnIsTimerWorking = False
Exit Sub
errhandle:
    mblnIsTimerWorking = False
    Set objFileSys = Nothing
    
    
End Sub

Private Sub DrawImgStates(Optional ByVal blnIsForce As Boolean = False)
'����ͼ����״̬
    Dim i As Long
    
    For i = 1 To dcmViewer.Images.Count
        If timerState.Enabled = False And blnIsForce = False Then
            Exit Sub
        End If
        
        Call DrawImgState(i, blnIsForce)
         
    Next
     
End Sub

Private Sub DrawImgState(ByVal lngImgIndex As Long, Optional ByVal blnIsForceDraw As Boolean = False)
    Dim objImgInfo As clsBgImgInfo
On Error GoTo errhandle
    If lngImgIndex <= 0 Or lngImgIndex > dcmViewer.Images.Count Then Exit Sub
    
    Set objImgInfo = dcmViewer.Images(lngImgIndex).tag
    
    If objImgInfo Is Nothing Then Exit Sub
    
    Select Case objImgInfo.LoadState
        Case lsLocal
            If objImgInfo.IsReDrawed Then Exit Sub
            Call DrawDicomFile(objImgInfo, lngImgIndex, blnIsForceDraw)
            
        Case lsRedo, lsError
            Call DrawErrorInfo(dcmViewer.Images(lngImgIndex), objImgInfo)
            
    End Select

    objImgInfo.IsReDrawed = True
Exit Sub
errhandle:
    
End Sub

Private Sub DrawResultState(ByVal lngResultState As TResultState)
'���ƴ�����
On Error GoTo errhandle
    Select Case lngResultState
        Case rsClear
            labState.ForeColor = picScroll.BackColor
            labState.Caption = " --"
        Case rsOk
            labState.ForeColor = &H8000&
            labState.Caption = " OK"
        Case rsFailed
            labState.ForeColor = vbRed
            labState.Caption = "ERR"
    End Select
Exit Sub
errhandle:
    
End Sub

Private Sub LockUpdateEx(ByVal lngHwnd As Long)
'    Call LockWindowUpdate(lngHwnd)
End Sub


Private Sub DrawDicomFile(objImgInfo As clsBgImgInfo, ByVal lngImgIndex As Long, Optional ByVal blnIsForceDraw As Boolean = False)
'����dicom�ļ�
    Dim strFile As String
    Dim objImg As DicomImage
    Dim strError As String
    
    LockUpdateEx dcmViewer.hwnd
On Error GoTo errhandle
    '�����������dicomͼ��
    Call dcmViewer.Images.Remove(lngImgIndex)
' Debug.Print GetTickCount & ":DrawDicomFile --" & lngImgIndex
    
    If objImgInfo.Format = ifAvi Then
        Set objImg = ReadMediaFile(sitAvi, strError)
    ElseIf objImgInfo.Format = ifWav Then
        Set objImg = ReadMediaFile(sitWav, strError)
    Else
        strFile = FormatFilePath(objImgInfo.FilePath & "\" & objImgInfo.Filename)
        Set objImg = ReadDicomFile(strFile, strError, IIf(objImgInfo.Format = ifDcm, True, False))
    End If
    
    If timerState.Enabled = False And blnIsForceDraw = False Then
        LockUpdateEx 0
        Exit Sub
    End If
    
    If objImg Is Nothing Then
        If objImgInfo.ErrorInfo = "" Then objImgInfo.ErrorInfo = strError
        Set objImg = ReadMediaFile(sitErr, strError)
        
        objImg.InstanceUID = objImgInfo.Key
        Set objImg.tag = objImgInfo
        Call AddImgToViewer(objImg, objImgInfo)
        
        If dcmViewer.Images.Count <> lngImgIndex Then
            Call dcmViewer.Images.Move(dcmViewer.Images.Count, lngImgIndex)
        End If
        
        Call DrawErrorInfo(objImg, objImgInfo)
    Else
        'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
        '���½�ú��DSAͼ����������ʾ
        '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
        '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
        If Not IsNull(objImg.Attributes(&H28, &H6100).value) Then
            objImg.Attributes.Remove &H28, &H6100
        End If
            
        Set objImg.tag = objImgInfo
        Call AddImgToViewer(objImg, objImgInfo)
        
        If dcmViewer.Images.Count <> lngImgIndex Then
            Call dcmViewer.Images.Move(dcmViewer.Images.Count, lngImgIndex)
        End If
    End If
    
    LockUpdateEx 0
Exit Sub
errhandle:
    LockUpdateEx 0
    
    
End Sub


Private Sub DrawRunState(Optional ByVal blnIsClear As Boolean = False)
'����״̬����ʾ��̨���ڴ���
On Error GoTo errhandle
    If blnIsClear Then
        labState.Caption = " --"
        labState.BackColor = picScroll.BackColor
        Exit Sub
    End If
    
    If mdtStartTime = CDate(0) Then
        Exit Sub
    End If
    
    If timerState.Enabled = False Then Exit Sub
    
    '�ж��Ƿ�ʱ����ʱ��ʹ�ú�ɫ��˸ͼ��
    If DateDiff("s", mdtStartTime, Now) > mlngTimeOut Then
        labState.BackColor = IIf(labState.BackColor <> picScroll.BackColor, picScroll.BackColor, vbRed)
    Else
        labState.BackColor = IIf(labState.BackColor <> picScroll.BackColor, picScroll.BackColor, &H8000&)
    End If
    
    labState.ForeColor = ColorConstants.vbBlack
Exit Sub
errhandle:
    
End Sub


Private Sub txtRecordCount_Change()
On Error GoTo errhandle
    If txtRecordCount.tag = "1" Then Exit Sub
    
    mlngPageRecord = Val(txtRecordCount.Text)
    
    If mlngPageRecord <= 0 Then
        mlngPageRecord = 8
        txtRecordCount.Text = mlngPageRecord
    End If
    
    Call ConfigPage(True)
    
    Call ReDrawImages(mlngPageIndex, True)
Exit Sub
errhandle:
    
End Sub

Private Sub UserControl_Initialize()
    mlngPageRecord = 8
    mlngPageIndex = 1
    mlngServerTime = VALIDTIME
    mlngTimeOut = 10
    mblnIsBGReadProcessing = False
    mblnIsDrawOrder = True
    mblnIsDrawHint = True
    mblnIsShowCheck = True
    mblnIsShowState = True
    mblnIsPageConfig = False
    mlngSelColorStyle = ColorConstants.vbRed
    mblnBGServerStarted = False
    mstrUploadCmdNames = ""
End Sub
 

Private Sub UserControl_Resize()
On Error Resume Next
    dcmViewer.Left = 0
    dcmViewer.Top = 0
    dcmViewer.Width = ScaleWidth
    dcmViewer.Height = ScaleHeight - picScroll.Height
    
    picScroll.Left = 0
    picScroll.Top = dcmViewer.Height
    picScroll.Width = ScaleWidth
     
    cbxPage.Left = picScroll.Width - cbxPage.Width
    txtRecordCount.Left = cbxPage.Left - txtRecordCount.Width
End Sub

Public Sub Destory()
On Error Resume Next
    Dim i As Long
    
    For i = 0 To ImgCount - 1
        Set maryImgInfo(i) = Nothing
    Next
     
    EraseAry maryImgInfo
    
    For i = 0 To ImgBufCount - 1 ' UBound(maryImgBuf())
        Set maryImgBuf(i) = Nothing
    Next
      
    EraseAry maryImgBuf
End Sub

Private Sub UserControl_Terminate()
    Call Destory
End Sub
