VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.2#0"; "DicomObjects.ocx"
Begin VB.Form frmImgSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "����ͼ��ѡ��"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   6840
      Width           =   7215
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5940
         TabIndex        =   5
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ѡ��(&S)"
         Height          =   350
         Left            =   4800
         TabIndex        =   4
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.Frame fraSplit1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   7110
   End
   Begin VB.PictureBox picView 
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   3375
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2415
         _Version        =   262146
         _ExtentX        =   4260
         _ExtentY        =   5953
         _StockProps     =   35
         BackColor       =   0
      End
   End
End
Attribute VB_Name = "frmImgSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pgbLoad As Object
Private AdviceID As Long, lngSendNO As Long
Private iPatientType As Integer, lngPatientID As Long, lngPatientDept As Long
Private lngPageID As Long, strCheckNo As String
Private int�Ʒ�״̬ As Integer, str�ѱ� As String, int��¼���� As Integer
Private intִ��״̬ As Integer, strNO As String, lng��������ID As Long
Private mstrPrivs As String

Private strURL As String
Private Inet1 As New clsFtp
Private Inet2 As New clsFtp
Private strDeviceNO1 As String
Private strDeviceNO2 As String
Private strVirtualPath As String

Private iCurImageIndex As Integer
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function zlRefresh(ByVal lngAdviceID As Long, ByVal SendNO As Long, strLocalPath As String) As String

    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo DBError
    strSQL = _
        " Select X.��¼���� as ��������,X.��¼״̬ as ����״̬," & _
        " A.ҽ��ID,A.���ͺ�,B.���ID,B.���,B.�������,B.������ĿID,A.����ʱ�� as ʱ��,A.NO," & _
        " A.��¼����,A.ִ��״̬,A.�Ʒ�״̬,B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,E.���� as ����,D.����," & _
        " Decode(B.������Դ,1,D.�����,2,D.סԺ��,NULL) as ��ʶ��,Nvl(F.�ѱ�,D.�ѱ�) as �ѱ�," & _
        " Decode(B.������Դ,1,'����',2,'סԺ',3,'����') as ��Դ,C.���� as ����,A.ִ�м�,A.ִ�в���ID" & _
        " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C,������Ϣ D,���ű� E,������ҳ F,���˷��ü�¼ X" & _
        " Where A.ҽ��ID=B.ID And B.������ĿID=C.ID And B.����ID=D.����ID" & _
        " And B.���˿���ID=E.ID And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+)" & _
        " And A.NO=X.NO(+) And A.��¼����=Decode(X.��¼����(+),0,1,X.��¼����(+))" & _
        " And X.��¼״̬(+)<>2 And X.ҽ�����(+)=A.ҽ��ID And X.���(+)=1 And C.���='D'" & _
        " And A.ҽ��ID=[1] And A.���ͺ�=[2]" & _
        " Order by A.����ʱ�� Desc,B.����ID,B.���"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, SendNO)
    
    AdviceID = lngAdviceID: lngSendNO = SendNO: iPatientType = 1
    lngPatientID = 0: lngPageID = 0: strCheckNo = "": lngPatientDept = 0
    int�Ʒ�״̬ = 0: str�ѱ� = "": int��¼���� = 1
    intִ��״̬ = 0: strNO = "": lng��������ID = 0
    If Not rsTmp.EOF Then
        iPatientType = IIf(rsTmp("��Դ") = "����", 1, 2)
        lngPatientID = rsTmp("����ID"): lngPageID = NVL(rsTmp("��ҳID"), 0): strCheckNo = NVL(rsTmp("�Һŵ�"), "")
        lngPatientDept = NVL(rsTmp("���˿���ID"), 0)
        int�Ʒ�״̬ = NVL(rsTmp!�Ʒ�״̬, 0): str�ѱ� = NVL(rsTmp!�ѱ�): int��¼���� = NVL(rsTmp!��¼����, 1)
        intִ��״̬ = NVL(rsTmp!ִ��״̬, 0): strNO = NVL(rsTmp!NO): lng��������ID = NVL(rsTmp!ִ�в���ID, 0)
    End If
    ShowMenu
    
    strSQL = " select Decode(��������,Null,'',to_Char(��������,'YYYYMMDD')||'/')||���UID As URL1 " & _
             " From Ӱ�����¼ where ҽ��ID = [1] and ���ͺ� = [2]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, SendNO)
    
    If Not rsTmp.EOF Then
        strLocalPath = IIf(IsNull(rsTmp("URL1")), "", rsTmp("URL1"))
    End If
    
    Me.Tag = "Loading": strURL = ""
    Me.Show vbModal
    
    zlRefresh = strURL
    
    Unload Me
    Exit Function
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If iCurImageIndex > 0 Then strURL = DViewer.Images(iCurImageIndex).Tag
    Me.Hide
End Sub

Private Sub DViewer_DblClick()
    cmdOK_Click
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(x, y)
        If i > 0 And i <= .Images.Count And iCurImageIndex > 0 And i <> iCurImageIndex Then
            .Images(iCurImageIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurImageIndex = i
        End If
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.Tag = "Loading" Then
        Me.Tag = ""
        ShowImages
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Load()
'
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.fraSplit1.Top > Me.ScaleHeight Then Me.fraSplit1.Top = Me.ScaleHeight / 2
    
    With Me.fraSplit1
        .Left = 0: .Width = Me.ScaleWidth - .Left
    End With
    With Me.Frame1
        .Left = -100: .Width = Me.ScaleWidth + 100 - .Left
        .Top = Me.ScaleHeight + 100 - .Height
    End With
    With cmdOK
        .Left = Me.Frame1.Width - 300 - Me.cmdCancel.Width - 60 - .Width
    End With
    With cmdCancel
        .Left = Me.Frame1.Width - 300 - Me.cmdCancel.Width
    End With
    With Me.picView
        .Left = 0: .Top = 0
        .Width = Me.ScaleWidth - .Left: .Height = Me.Frame1.Top - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If iCurImageIndex = 0 Then Cancel = True
End Sub

Private Sub fraSplit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    fraSplit1.BackColor = RGB(0, 0, 0)
    On Error Resume Next
    If fraSplit1.Top + y < 2000 Then
        fraSplit1.Top = 2000
    ElseIf Me.ScaleHeight - fraSplit1.Top - y < 4000 Then
        fraSplit1.Top = Me.ScaleHeight - 4000
    Else
        fraSplit1.Top = fraSplit1.Top + y
    End If
End Sub

Private Sub fraSplit1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraSplit1.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub ShowMenu()
'
End Sub

Private Sub ShowImages()
    Dim strSQL As String
    Dim strURL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double, lngRecID As Long
    Dim curImage As DicomImage
    Dim iRows As Integer, iCols As Integer
    Dim i As Integer, aImages() As String, iNum As Integer
    Dim strImages As String
    
    Dim strTempPath As String, lngBuffSize As Long
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    
    Dim strHost As String
    
    On Error GoTo DBError
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    
    strSQL = "Select ����ͼ��," & _
        "D.�û��� As User1,D.���� As Pwd1," & _
        "D.IP��ַ As Host1," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL1,d.�豸�� as �豸��1," & _
        "E.�û��� As User2,E.���� As Pwd2," & _
        "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL2 , e.�豸�� as �豸��2 " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And C.ҽ��ID=[1] And C.���ͺ�=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, AdviceID, lngSendNO)
    Screen.MousePointer = vbHourglass
    
    iCurImageIndex = 0
    With DViewer
        .Images.Clear
        If rsTmp.RecordCount > 0 Then
            strImages = Trim(Split(NVL(rsTmp(0), " "), "|")(0))
            If Len(strImages) = 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "û�б���ͼ��������Ƭվ���ã�", vbInformation, gstrSysName
                Me.Hide
                Exit Sub
            End If
            
            aImages = Split(strImages, ";"): iNum = UBound(aImages, 1)
'            Inet.strIPAddress = NVL(rsTmp("Host1")): Inet.strUser = NVL(rsTmp("User1")): Inet.strPsw = NVL(rsTmp("Pwd1"))
            
            ResizeRegion iNum + 1, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows

            lngRecID = 1
            For i = 0 To iNum
'                If Len(LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 And _
'                    InStr("bmp;jpg;jpeg;gif;ico", LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 Then
                
                    Set curImage = New DicomImage
'                    strTmpFile = strTempPath & objFileSystem.GetFileName(Trim(aImages(i)))
                    strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & rsTmp("URL1")
                    strTmpFile = Replace(strTmpFile, "/", "\")
                    MkLocalDir strTmpFile
                    strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(Trim(aImages(i)))
                
                If Dir(strTmpFile, vbDirectory) = "" Then
                    strHost = "ftp://" & NVL(rsTmp("User1")) & IIf(IsNull(rsTmp("Pwd1")), "", ":" & rsTmp("Pwd1")) & _
                        "@" & NVL(rsTmp("Host1"))
                    strVirtualPath = objFileSystem.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL1") & "/" & Trim(aImages(i)))
                    If strDeviceNO1 <> rsTmp("�豸��1") Then
                        strDeviceNO1 = rsTmp("�豸��1")
                        Inet1.FuncFtpConnect NVL(rsTmp("Host1")), NVL(rsTmp("User1")), NVL(rsTmp("Pwd1"))
                    End If
                    If strDeviceNO2 <> rsTmp("�豸��2") Then
                        strDeviceNO2 = rsTmp("�豸��2")
                        Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("User2")), NVL(rsTmp("Pwd2"))
                    End If
                    If Inet1.FuncDownloadFile(strVirtualPath, strTmpFile, Trim(aImages(i))) <> 0 Then
'                        Inet.strIPAddress = NVL(rsTmp("Host2")): Inet.strUser = NVL(rsTmp("User2")): Inet.strPsw = NVL(rsTmp("Pwd2"))
                        strHost = "ftp://" & NVL(rsTmp("User2")) & IIf(IsNull(rsTmp("Pwd2")), "", ":" & rsTmp("Pwd2")) & _
                            "@" & NVL(rsTmp("Host2"))
                        strVirtualPath = objFileSystem.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL2") & "/" & Trim(aImages(i)))
                        Call Inet2.FuncDownloadFile(strVirtualPath, strTmpFile, Trim(aImages(i)))
                    End If
                End If
                If Len(LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 And _
                    InStr("bmp;jpg;jpeg;gif;ico", LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 Then
                    curImage.FileImport strTmpFile, objFileSystem.GetExtensionName(Trim(aImages(i)))
'                    objFileSystem.DeleteFile strTmpFile, True
                    .Images.Add curImage: Set curImage = .Images(.Images.Count)
                Else
                    Set curImage = .Images.ReadFile(strTmpFile)
                End If
'                Else
'                    Set curImage = .Images.ReadURL(Inet.URL & rsTmp("URL") & Trim(aImages(i)))
'                End If
                With curImage
                    .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
                    .ShowLabels = True: .Tag = strHost & "," & strVirtualPath & "/" & Trim(aImages(i))
                End With
            Next
            
            iCurImageIndex = 1
            .Images(iCurImageIndex).BorderColour = vbRed
        Else
            Screen.MousePointer = vbDefault
            MsgBox "�ü��δ���У�û�б���ͼ��", vbInformation, gstrSysName
            Me.Hide
            Exit Sub
'            .MultiColumns = 1: .MultiRows = 1
        End If
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Screen.MousePointer = vbDefault
    Call SaveErrLog
End Sub

Private Sub picView_Resize()
    Dim iCols As Integer, iRows As Integer
    
    On Error Resume Next
    With DViewer
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
        
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
End Sub

Private Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub
Private Sub MkLocalDir(ByVal strDir As String)
'���ܣ���������Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub
