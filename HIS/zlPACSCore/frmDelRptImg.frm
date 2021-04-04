VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "dicomobjects.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDelRptImg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ɾ������ͼ��"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   Icon            =   "frmDelRptImg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6360
      TabIndex        =   3
      Top             =   6600
      Width           =   1100
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   600
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4080
      TabIndex        =   2
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   1680
      TabIndex        =   1
      Top             =   6600
      Width           =   1100
   End
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _Version        =   262147
      _ExtentX        =   16113
      _ExtentY        =   11033
      _StockProps     =   35
      BackColor       =   -2147483640
   End
End
Attribute VB_Name = "frmDelRptImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As frmViewer
Public pSeriesUID As String      '��¼��ǰ����ͼ����������ͼ�������UID��
                                '��Ϊͼ�������UID�����ݿ���û�б��޸ģ�����ʹ������UID�����ұ���ͼ��ȼ��UID���Ҹ���׼ȷ
Private iCurImageIndex As Integer
Private bDelete As Boolean
Private aImages() As String
Private aDelRecord As String
Private strRepImgField As String      '�������ݿ���ԭ���ı���ͼ���ֶ�ֵ
Private strURL As String                '����FTP����·��

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    If iCurImageIndex > 0 Then
        aDelRecord = IIf(aDelRecord = vbNullString, " ", aDelRecord & ",") & Me.DViewer.Images(iCurImageIndex).Tag
        Me.DViewer.Images.Remove iCurImageIndex
        iCurImageIndex = iCurImageIndex - 1
        If iCurImageIndex <= 0 And Me.DViewer.Images.Count > 0 Then
            iCurImageIndex = 1
        End If
        If iCurImageIndex > 0 Then
            Me.DViewer.Images(iCurImageIndex).BorderColour = vbRed
            Me.DViewer.Refresh
        End If
        bDelete = True
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim aDelImages() As String
    Dim aRptImg As String       '��¼����ͼ���ļ�����
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim aDelNum() As String
    
    If bDelete Then '����ɾ�����
        aDelNum = Split(aDelRecord, ",")
        
        If UBound(aDelNum) < 0 Then
            Unload Me
            Exit Sub
        End If
        
        ReDim aDelImages(UBound(aDelNum, 1)) As String
        For i = 0 To UBound(aDelNum, 1)
            aDelImages(i) = aImages(aDelNum(i))
            aImages(aDelNum(i)) = "DEL"
        Next
        For i = 0 To UBound(aImages, 1) '��ϱ���ͼ���ļ�����
            If aImages(i) <> "DEL" Then
                aRptImg = IIf(aRptImg = vbNullString, " ", aRptImg & ";") & aImages(i)
            End If
        Next
        '������ͼ�񱣴浽���ݿ���
        If InStr(strRepImgField, "|") <> 0 Then     '��¼������
            aRptImg = aRptImg & Right(strRepImgField, Len(strRepImgField) - InStr(strRepImgField, "|") + 1)
        End If
                 
        strSQL = "ZL_Ӱ���鱨��_UPDATE('" & PstrCheckUID & "','" & aRptImg & "')"
        zlDatabase.ExecuteProcedure strSQL, App.ProductName
        
        'ɾ��FTP�б�ɾ���ı���ͼ��
        For i = 0 To UBound(aDelImages, 1)
            Inet.Execute , "DELETE " & strURL & Trim(aDelImages(i))
            Do While Inet.StillExecuting
                DoEvents
            Loop
        Next
    End If
    Unload Me
End Sub


Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(x, y)
        If i <> iCurImageIndex Then
            If iCurImageIndex > 0 And iCurImageIndex <= .Images.Count Then
                .Images(iCurImageIndex).BorderColour = vbWhite
            End If
            If i > 0 And i <= .Images.Count Then
                .Images(i).BorderColour = vbRed
                iCurImageIndex = i
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
     ShowImages pSeriesUID
     aDelRecord = vbNullString
End Sub

Public Sub ShowImages(SeriesUID As String)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double, lngRecID As Long
    Dim curImage As DicomImage
    Dim iRows As Integer, iCols As Integer
    Dim i As Integer, iNum As Integer
    Dim aTemp() As String
    
    Dim strTempPath As String, lngBuffSize As Long
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    
    If gcnOracle Is Nothing Then Exit Sub
    On Error GoTo DBError
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    '�ָ�������UID�����Ҷ�Ӧ�ļ��UID
    strSQL = "select ���UID FROM Ӱ��������  where ����UID =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SeriesUID)
    If rsTmp.RecordCount = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "��ͼ��û�б��浽�����У�û�ж�Ӧ�ı���ͼ��", vbInformation, gstrSysName
        Me.Hide
        Exit Sub
    End If
    PstrCheckUID = Nvl(rsTmp!���UID)
    
    '�Ȳ��豸һ��û���ٲ��豸��
    strSQL = "Select ����ͼ��," & _
        "'ftp://'||Decode(FTP�û���,Null,'',FTP�û���||Decode(FTP����,Null,'',':'||FTP����))" & _
        "||'@'||IP��ַ As Host,'/'||FtpĿ¼||'/'" & _
        "||Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/' As URL " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where C.λ��һ=D.�豸��  " & _
        "And C.���UID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, PstrCheckUID)
    If rsTmp.RecordCount = 0 Then
        '���豸��
        strSQL = "Select ����ͼ��," & _
            "'ftp://'||Decode(FTP�û���,Null,'',FTP�û���||Decode(FTP����,Null,'',':'||FTP����))" & _
            "||'@'||IP��ַ As Host,'/'||FtpĿ¼||'/'" & _
            "||Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
            "||C.���UID||'/' As URL " & _
            "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
            "Where C.λ�ö�=D.�豸��  " & _
            "And C.���UID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, PstrCheckUID)
    End If
    Screen.MousePointer = vbHourglass
    
    iCurImageIndex = 0
    With DViewer
        .Images.Clear
        If rsTmp.RecordCount > 0 Then
            If Len(Nvl(rsTmp(0))) = 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "û�б���ͼ�����ڹ�Ƭվ�б��汨��ͼ��", vbInformation, gstrSysName
                Me.Hide
                Exit Sub
            End If
            
            strRepImgField = Nvl(rsTmp(0))
            If InStr(strRepImgField, "|") <> 0 Then      '��¼������
                aTemp = Split(Nvl(rsTmp(0)), "|")
                If UBound(aTemp, 1) > 0 Then
                    aImages = Split(aTemp(0), ";")
                End If
            Else
                aImages = Split(Nvl(rsTmp(0)), ";")
            End If
            iNum = UBound(aImages, 1)
            If iNum = -1 Then
                Screen.MousePointer = vbDefault
                MsgBox "û�б���ͼ��������Ƭվ���ã�", vbInformation, gstrSysName
                Me.Hide
                Exit Sub
            End If
            Inet.AccessType = icUseDefault: Inet.URL = rsTmp("Host")
            strURL = rsTmp("URL")
            
            ResizeRegion iNum + 1, .width, .height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows

            lngRecID = 1
            For i = 0 To iNum
                If Len(LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 And _
                    InStr("bmp;jpg;jpeg;gif;ico", LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 Then
                
                    Set curImage = New DicomImage
                    strTmpFile = strTempPath & objFileSystem.GetFileName(Trim(aImages(i)))
                    Inet.Execute , "Get " & rsTmp("URL") & Trim(aImages(i)) & " " & strTmpFile
                    Do While Inet.StillExecuting
                        DoEvents
                    Loop
                    If Dir(strTmpFile) <> "" Then
                        curImage.FileImport strTmpFile, ""
                        objFileSystem.DeleteFile strTmpFile, True
                        .Images.Add curImage: Set curImage = .Images(.Images.Count)
                    End If
                Else
                    Set curImage = .Images.ReadURL(Inet.URL & rsTmp("URL") & Trim(aImages(i)))
                End If
                If Not curImage Is Nothing Then
                    With curImage
                        .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
                        .ShowLabels = True: .Tag = i
                    End With
                End If
            Next
            If .Images.Count > 0 Then
                iCurImageIndex = 1
                .Images(iCurImageIndex).BorderColour = vbRed
            End If
        Else
            Screen.MousePointer = vbDefault
            MsgBox "�ü��δ���У�û�б���ͼ��", vbInformation, gstrSysName
            Me.Hide
            Exit Sub
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

