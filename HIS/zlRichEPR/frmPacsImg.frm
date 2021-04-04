VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSImg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6396
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6396
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   4875
      Left            =   105
      TabIndex        =   0
      Top             =   15
      Width           =   2535
      _Version        =   262147
      _ExtentX        =   4471
      _ExtentY        =   8599
      _StockProps     =   35
      BackColor       =   14737632
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   90
      Top             =   5535
      _Version        =   589884
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmPACSImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAdviceID As Long      '��ǰ�༭��ҽ��id
Private mlngCurIndex As Long      '��ǰѡ���ͼ�����
Private WithEvents mfrmShow As frmPacsImgShow
Attribute mfrmShow.VB_VarHelpID = -1
Private mlngModule As Long

'�����¼�
Public Event RequestRightMenu(ByRef cbsThis As Object)
Public Event InsertPicture(ByRef pic As StdPicture, ByVal strUid As String, ByVal lngAdviceID As Long)

Private Sub ConfigImgDisplayFormat(ByVal lngPageRecord As Long)
'����ͼ����ʾ��ʽ
    Dim iRows As Integer
    Dim iCols As Integer
    
    ResizeRegion lngPageRecord, DViewer.Width, DViewer.Height, iRows, iCols

    DViewer.MultiColumns = iCols
    DViewer.MultiRows = iRows
End Sub

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
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

Public Function GetCacheDir() As String
'��ȡ����Ŀ¼
    GetCacheDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
End Function

Private Function LoadAllCaptureImage(ByVal lngAdviceID As Long, dcmViewer As DicomViewer) As Boolean
'����Ԥ��ͼ�񵽽���
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim Inet1 As New cFTP
    Dim Inet2 As New cFTP
    
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim objImgInfo As Object
    
    Dim strSQL As String
    Dim rsCurImageData As ADODB.Recordset

    strSQL = "Select rownum as ˳���,A.ͼ��UID,c.����,c.�Ա�,c.����, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
            "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,D.����Ŀ¼ as ����Ŀ¼1,D.����Ŀ¼�û��� as ����Ŀ¼�û���1,D.����Ŀ¼���� as ����Ŀ¼����1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/') " & _
            "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1," & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
            "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,E.����Ŀ¼ as ����Ŀ¼2,E.����Ŀ¼�û��� as ����Ŀ¼�û���2,E.����Ŀ¼���� as ����Ŀ¼����2," & _
            "E.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And nvl(A.��̬ͼ,0) = 0 and c.ҽ��ID = [1]"
    Set rsCurImageData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    blnIsAddImage = False
    
    LoadAllCaptureImage = False
    
    If rsCurImageData.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    strCurInstanceUids = ""
        
    '����ͼ����ʾ��ʽ
    If rsCurImageData.RecordCount <> dcmViewer.Images.Count Then
        Call ConfigImgDisplayFormat(rsCurImageData.RecordCount)
    End If
        
    '��������ͼ�񻺴�Ŀ¼
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
    Do While Not rsCurImageData.EOF
        'ѭ������ͼ��DicomViewer��
        strImgInstanceUid = NVL(rsCurImageData!ͼ��UID)
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 Then
            
            blnIsAddImage = True
            
            '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
            If NVL(rsCurImageData!��̬ͼ, 0) = 0 Then
                strTmpFile = strCachePath & NVL(rsCurImageData("URL"))
            End If
            
            If Dir(strTmpFile) = vbNullString Then
                '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                '����FTP����
                If NVL(rsCurImageData("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(NVL(rsCurImageData("Host1")), NVL(rsCurImageData("User1")), NVL(rsCurImageData("Pwd1"))) = 0 Then
                        If NVL(rsCurImageData("�豸��2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))) = 0 Then
                                Exit Function
                            End If
                        Else
                            Exit Function
                        End If
                    End If
                End If
                
                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL"))) <> 0 Then
                    '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                    If NVL(rsCurImageData("�豸��2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")))
                    End If
                End If
            End If
    
            If Dir(strTmpFile) <> vbNullString Then
                If NVL(rsCurImageData!��̬ͼ, 0) = 0 Then
                    Err.Clear
                    On Error Resume Next
                    Set curImage = dcmViewer.Images.ReadFile(strTmpFile)
                    
                    If Err.Number <> 0 Then
                        Set curImage = dcmViewer.Images.AddNew
                        Call curImage.FileImport(strTmpFile, "JPG")
                    End If
                    
                    curImage.Tag = NVL(rsCurImageData("ͼ��UID")) & ".jpg"
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                End If
                
                'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
                '���½�ú��DSAͼ����������ʾ
                '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
                '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
                If Not IsNull(curImage.Attributes(&H28, &H6100).Value) Then
                    curImage.Attributes.Remove &H28, &H6100
                End If
            End If
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    UpdateSelectIndex 1
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    
    LoadAllCaptureImage = True
End Function

Private Sub UpdateSelectIndex(ByVal lngSelectIndex As Long)
'����ͼ���ѡ������
    Dim blnIsValidIndex As Boolean
    
    blnIsValidIndex = IIf(lngSelectIndex > 0 And lngSelectIndex <= Me.DViewer.Images.Count, True, False)
    
    If Not blnIsValidIndex Then Exit Sub

    If blnIsValidIndex Then DViewer.Images(lngSelectIndex).BorderColour = vbRed
    If mlngCurIndex = lngSelectIndex Then Exit Sub

    If mlngCurIndex > 0 And mlngCurIndex <= DViewer.Images.Count Then
        DViewer.Images(mlngCurIndex).BorderColour = vbWhite
    End If

    mlngCurIndex = lngSelectIndex
End Sub

Private Function LoadSelectReportImage(ByVal lngAdviceID As Long, dcmViewer As DicomViewer) As Boolean
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim aryFiles() As String    '����ͼ������
    Dim strFiles As String      '���ֺŷָ��ĳɹ����ص��ļ�
    
    Dim cFtpNet As New cFTP
    Dim strVirtualPath As String
    Dim strLocalPath As String
    Dim intCount As Integer
     
    LoadSelectReportImage = False
    
    '�ȸ��ݱ���ͼ���ֶλ�ȡ����ͼ��Ϣ��û��ֵ���ȡ���м��ͼ��
    strSQL = "Select To_Char(L.��������, 'yyyymmdd') As ��Ŀ¼, L.���uid, L.����ͼ��, A1.FtpĿ¼ As Root1, A1.Ip��ַ As Ip1," & vbNewLine & _
            "       A1.FTP�û��� As Usr1, A1.FTP���� As Pwd1, A2.FtpĿ¼ As Root2, A2.Ip��ַ As Ip2, A2.FTP�û��� As Usr2, A2.FTP���� As Pwd2" & vbNewLine & _
            "From Ӱ�����¼ L, Ӱ���豸Ŀ¼ A1, Ӱ���豸Ŀ¼ A2" & vbNewLine & _
            "Where L.λ��һ = A1.�豸��(+) And L.λ�ö� = A2.�豸��(+) And L.����ͼ�� Is Not Null And L.ҽ��id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ϣ", lngAdviceID)

    If rsTemp.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    aryFiles = Split("" & rsTemp!����ͼ��, ";")
    If UBound(aryFiles) < 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
        
    '�������ش洢Ŀ¼
    Err = 0: On Error Resume Next
    strLocalPath = App.Path & "\TmpImage\" & rsTemp!��Ŀ¼
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then Exit Function
    
    strLocalPath = strLocalPath & "\" & rsTemp!���uid
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then Exit Function
        
    '�ж����ӵ���Ч�ԣ��������������ļ�
    strFiles = ""
    If "" & rsTemp!Ip1 <> "" Then
        If cFtpNet.FuncFtpConnect("" & rsTemp!Ip1, "" & rsTemp!Usr1, "" & rsTemp!pwd1) <> 0 Then
            strVirtualPath = rsTemp!Root1 & "/" & rsTemp!��Ŀ¼ & "/" & rsTemp!���uid
            For intCount = 0 To UBound(aryFiles)
                If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
                    strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
                    aryFiles(intCount) = ""
                Else
                    If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(intCount)), Trim(aryFiles(intCount))) = 0 Then
                        If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
                            strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
                            aryFiles(intCount) = ""
                        End If
                    End If
                End If
            Next
        End If
        cFtpNet.FuncFtpDisConnect
    End If
    
    If strFiles <> "" Then strFiles = Mid(strFiles, 2)
    If UBound(Split(strFiles, ";")) <> UBound(aryFiles) And "" & rsTemp!Ip2 <> "" Then
        If cFtpNet.FuncFtpConnect("" & rsTemp!Ip2, "" & rsTemp!Usr2, "" & rsTemp!pwd2) <> 0 Then
            strVirtualPath = rsTemp!Root2 & "/" & rsTemp!��Ŀ¼ & "/" & rsTemp!���uid
            For intCount = 0 To UBound(aryFiles)
                If aryFiles(intCount) <> "" Then
                    If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(intCount)), Trim(aryFiles(intCount))) = 0 Then
                        If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
                            strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
                        End If
                    End If
                End If
            Next
        End If
        cFtpNet.FuncFtpDisConnect
    End If
    
    If strFiles <> "" Then
        If Left(strFiles, 1) = ";" Then strFiles = Mid(strFiles, 2)
    End If
    
    '����õ��ļ�װ��
    Dim curImage As DicomImage, iRows As Integer, iCols As Integer, strCurName As String
    aryFiles = Split(strFiles, ";")
    With Me.DViewer
        If .Images.Count > 0 Then strCurName = .Images(mlngCurIndex - 1).Tag '���µ�ǰ��ѡ
        mlngCurIndex = 0
        .Images.Clear
        For intCount = 0 To UBound(aryFiles)
            Set curImage = New DicomImage
            curImage.FileImport aryFiles(intCount), "JPG"
            curImage.BorderStyle = 6: curImage.BorderWidth = 1: curImage.BorderColour = vbWhite
            .Images.Add curImage
            .Images(intCount + 1).Tag = gobjFSO.GetFileName(aryFiles(intCount))                    '�����ļ��������
            If strCurName = aryFiles(intCount) Then mlngCurIndex = intCount + 1
        Next
        
        If .Images.Count > 0 Then
            If mlngCurIndex = 0 Then
                mlngCurIndex = 1
                If strCurName <> "" Then Unload mfrmShow                'ˢ��֮ǰ�е�ͼ,ˢ��֮��û��,�������ܱ�ɾ��
            End If
            .CurrentIndex = 1
            .Images(mlngCurIndex).BorderColour = vbRed
            
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols '��������
            .MultiColumns = iCols: .MultiRows = iRows
        Else
            Unload mfrmShow
        End If
    End With
    
End Function

Public Function zlRefresh(ByVal lngAdviceID As Long, ByVal lngModule As Long) As Boolean
    '��ȡ����ͼ�񵽱��أ���ˢ����ʾ
On Error GoTo errHand
        
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim aryFiles() As String    '����ͼ������
    Dim strFiles As String      '���ֺŷָ��ĳɹ����ص��ļ�
    
    Dim cFtpNet As New cFTP
    Dim strVirtualPath As String
    Dim strLocalPath As String
    Dim intCount As Integer
     
    mlngAdviceID = lngAdviceID
    mlngModule = lngModule
    
    '�������ظ�Ŀ¼
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then objFileSystem.CreateFolder App.Path & "\TmpImage\"
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then zlRefresh = False: Exit Function
     
    '���ݲ�ͬģ�飬ʹ�ò�ͬ�ı���ͼ��ȡ��ʽ
    If mlngModule = 1290 Then
        zlRefresh = LoadSelectReportImage(lngAdviceID, DViewer)
    Else
        zlRefresh = LoadAllCaptureImage(lngAdviceID, DViewer)
    End If

    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'---------------------------------------------------
'�����Ǵ���ռ��¼�����
'---------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Insert
        Call DViewer_DblClick
    Case conMenu_View_Refresh
        Call Me.zlRefresh(mlngAdviceID, mlngModule)
    Case conMenu_View_Option
        If mfrmShow.Visible Then
            Unload mfrmShow
        Else
            mfrmShow.Show , Me
            Call DViewer_Click
        End If
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Left = -120
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Dim iRows As Integer, iCols As Integer
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    With Me.DViewer
        iRows = .MultiRows: iCols = .MultiColumns
        .Left = lngScaleLeft + 120: .Top = lngScaleTop
        .Width = lngScaleRight - .Left: .Height = lngScaleBottom - .Top
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Insert
        Control.Enabled = (mlngCurIndex > 0)
    Case conMenu_View_Option
        Control.Enabled = (Me.DViewer.Images.Count > 0)
        If (Control.Enabled = False Or Me.Visible = False) And mfrmShow.Visible Then Unload mfrmShow
        Control.Checked = mfrmShow.Visible
    End Select
End Sub

Private Sub DViewer_Click()
    Dim pic As StdPicture, picUid As String
    If mfrmShow.Visible = False Then Exit Sub
    If mlngCurIndex > 0 Then
        Set pic = Me.DViewer.Images(mlngCurIndex).Picture
        picUid = Me.DViewer.Images(mlngCurIndex).Tag
    End If
    
    If Not (pic Is Nothing) Then
        Set mfrmShow.imgShow.Picture = pic
        mfrmShow.imgShow.Tag = picUid
    Else
        Set mfrmShow.imgShow.Picture = Nothing
        mfrmShow.imgShow.Tag = ""
    End If
End Sub

Private Sub DViewer_DblClick()
    Dim pic As StdPicture, picUid As String
    If mlngCurIndex > 0 Then
        Set pic = Me.DViewer.Images(mlngCurIndex).Picture
        picUid = Me.DViewer.Images(mlngCurIndex).Tag
    End If
    If Not pic Is Nothing Then RaiseEvent InsertPicture(pic, picUid, mlngAdviceID)
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    With DViewer
        i = .ImageIndex(X, Y)
        If i > 0 And i <= .Images.Count And i <> mlngCurIndex Then
            .Images(mlngCurIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            mlngCurIndex = i
        End If
    End With
End Sub

Private Sub Form_Load()
    Set mfrmShow = New frmPacsImgShow
    
    Dim cbrControl As CommandBarControl
    '-----------------------------------------------------
    '�ڲ��˵�����������
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.Position = xtpBarBottom
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Insert, "���뱨��(&S)"): cbrControl.STYLE = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.STYLE = xtpButtonCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "��ͼ(&B)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.STYLE = xtpButtonCaption
        cbrControl.Checked = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmShow Is Nothing Then Unload mfrmShow
    Set mfrmShow = Nothing
End Sub
 




Private Sub mfrmShow_DblClick(pic As stdole.StdPicture, ByVal strUid As String)
    RaiseEvent InsertPicture(pic, strUid, mlngAdviceID)
End Sub
