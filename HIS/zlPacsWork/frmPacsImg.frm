VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSImg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   65535
      Left            =   3000
      Top             =   3480
   End
   Begin VB.Frame fraSplit1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   1
      Top             =   2520
      Width           =   7110
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsImg.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picView 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   2760
      Width           =   5295
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   3375
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   2415
         _Version        =   262146
         _ExtentX        =   4260
         _ExtentY        =   5953
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin MSComctlLib.ListView lvwSeq 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmPACSImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents frmParent As Form
Attribute frmParent.VB_VarHelpID = -1
Private pgbLoad As Object
Private AdviceID As Long, lngSendNO As Long
Private iPatientType As Integer, lngPatientID As Long, lngPatientDept As Long
Private lngPageId As Long, strCheckNo As String
Private mblnShowPic As Boolean, mDispImgs As Integer
Private int�Ʒ�״̬ As Integer, str�ѱ� As String, int��¼���� As Integer
Private intִ��״̬ As Integer, strNO As String, lng��������ID As Long
Private strCheckUID As String
Private mstrPrivs As String
Private mblnMoved As Boolean
Private mblnAddImage As Boolean                             '�Ƿ�׷��ͼ��

Private strCachePath As String

Private iCurImageIndex As Integer, strFtpHost As String, strDicomPath As String, strLocalPath As String
Private strFtpUser As String, strFtpPwd As String

Public Function zlRefresh(objParent As Object, ByVal lngAdviceID As Long, ByVal SendNO As Long, _
    ByVal strPrivs As String, Optional objpgbLoad As Object, Optional blnShowPic As Boolean = True, Optional ByVal blnMoved As Boolean = False, Optional ByVal iDispImgs As Integer) As Boolean

    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo DBError
    mblnMoved = blnMoved: mDispImgs = iDispImgs
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
        " And X.��¼״̬(+)<>2 And X.ҽ�����(+)=A.ҽ��ID And X.���(+)=1" & _
        " And A.ҽ��ID= [1]  And A.���ͺ�= [2] " & _
        " Order by A.����ʱ�� Desc,B.����ID,B.���"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
    End If
        
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, SendNO)

    mblnShowPic = blnShowPic
    Set frmParent = objParent
    Set pgbLoad = objpgbLoad
    AdviceID = lngAdviceID: lngSendNO = SendNO: iPatientType = 1
    lngPatientID = 0: lngPageId = 0: strCheckNo = "": lngPatientDept = 0
    int�Ʒ�״̬ = 0: str�ѱ� = "": int��¼���� = 1: mstrPrivs = strPrivs
    intִ��״̬ = 0: strNO = "": lng��������ID = 0
    
    'ת����Ӱ���ܱ��汨��
    If mblnMoved Then
        mstrPrivs = Replace(mstrPrivs, "ͼ���������", "")
        mstrPrivs = Replace(mstrPrivs, "ͼ���ע����", "")
        mstrPrivs = Replace(mstrPrivs, "������ͼ��", "")
    End If
    
    If Not rsTmp.EOF Then
        iPatientType = Decode(rsTmp("��Դ"), "����", 1, 2)
        lngPatientID = rsTmp("����ID"): lngPageId = Nvl(rsTmp("��ҳID"), 0): strCheckNo = Nvl(rsTmp("�Һŵ�"), "")
        lngPatientDept = Nvl(rsTmp("���˿���ID"), 0)
        int�Ʒ�״̬ = Nvl(rsTmp!�Ʒ�״̬, 0): str�ѱ� = Nvl(rsTmp!�ѱ�): int��¼���� = Nvl(rsTmp!��¼����, 1)
        intִ��״̬ = Nvl(rsTmp!ִ��״̬, 0): strNO = Nvl(rsTmp!NO): lng��������ID = Nvl(rsTmp!ִ�в���ID, 0)
    End If
        
    If frmParent.Visible Then
        ShowSeqList
        
        If lvwSeq.SelectedItem Is Nothing Then
            DViewer.Images.Clear
        Else
            lvwSeq_ItemClick lvwSeq.SelectedItem
            zlRefresh = True
        End If
    Else
        Me.Tag = "Loading":
    End If
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'ִ�в˵�����
Public Sub zlMenuClick(mnuClick As Menu)
    Dim strMenu As String
    
    If mnuClick.Caption Like "*(&*)*" Then
        strMenu = Split(mnuClick.Caption, "(&")(0)
    Else
        strMenu = mnuClick.Caption
    End If
    mblnAddImage = False
    Select Case strMenu
        Case "Ӱ����"
            DViewer_DblClick
        Case "Ӱ��Ա�"
            mblnAddImage = True
            DViewer_DblClick
        Case "��ʾ��ǰ����ͼ��"
            mblnShowPic = Not mblnShowPic
            If Not lvwSeq.SelectedItem Is Nothing Then lvwSeq_ItemClick lvwSeq.SelectedItem
        Case "ѡ����������"
            SelectAll True
        Case "���ѡ���־"
            SelectAll False
    End Select
End Sub

Public Sub zlButtonClick(ByVal Button As MSComctlLib.Button)
    mblnAddImage = False
    Select Case Button.Key
        Case "��Ƭ"
            DViewer_DblClick
        Case "��ʾ"
            mblnShowPic = Not mblnShowPic
            If Not lvwSeq.SelectedItem Is Nothing Then lvwSeq_ItemClick lvwSeq.SelectedItem
        Case "ȫѡ"
            SelectAll True
        Case "ȫ��"
            SelectAll False
    End Select
End Sub

Private Sub SelectAll(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwSeq
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
    End With
End Sub

Private Sub DViewer_DblClick()
    Dim tmpImages As DicomImages, aFiles() As String
    Dim objPacsCore As Object
    Dim strSerials As String, lngSeqUID As String
    Dim Item As MSComctlLib.ListItem
    Dim i As Integer
    
    On Error GoTo CallError
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
'    If InStr(mstrPrivs, "Ӱ����") = 0 Or (InStr(mstrPrivs, "��д����") = 0 And InStr(mstrPrivs, "�������") = 0) Then Exit Sub
    If InStr(mstrPrivs, "Ӱ����") = 0 Then Exit Sub
    If Not lvwSeq.SelectedItem.Checked Then lvwSeq.SelectedItem.Checked = True
    
    strSerials = ""
    For Each Item In lvwSeq.ListItems
        lngSeqUID = Mid(Item.Key, 2)
        If Item.Checked Then
            strSerials = strSerials & ",'" & lngSeqUID & "'"
            i = i + 1
        End If
    Next
    If Len(strSerials) > 0 Then strSerials = Mid(strSerials, 2)
    
    aFiles = GetAllImageFiles(strCheckUID, strSerials, mblnMoved, strFtpHost, strDicomPath, _
        strLocalPath, strFtpUser, strFtpPwd)
    If UBound(aFiles) > 0 Then
        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
        objPacsCore.CallOpenViewerCache aFiles, frmParent, strCachePath & strLocalPath, strFtpHost & strDicomPath, mstrPrivs, strCheckUID, strFtpHost, strDicomPath, gcnOracle, strFtpUser, strFtpPwd, mblnAddImage, i
        Set objPacsCore = Nothing
    End If
    Exit Sub
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(x, y)
        If i > 0 And i <= .Images.Count And i <> iCurImageIndex Then
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
        pgbLoad.Visible = True
        
        ShowSeqList
        If lvwSeq.SelectedItem Is Nothing Then
            DViewer.Images.Clear
        Else
            lvwSeq_ItemClick lvwSeq.SelectedItem
        End If
        
        pgbLoad.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim objFileSystem As New Scripting.FileSystemObject
    
    strCachePath = App.Path & "\TmpImage\"
    If Not objFileSystem.FolderExists(strCachePath) Then objFileSystem.CreateFolder strCachePath
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.fraSplit1.Top > Me.ScaleHeight Then Me.fraSplit1.Top = Me.ScaleHeight / 2
    
    With lvwSeq
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
        .Height = Me.fraSplit1.Top - .Top
    End With
    With Me.fraSplit1
        .Left = 0: .Width = Me.ScaleWidth - .Left
    End With
    With Me.picView
        .Left = 0: .Top = fraSplit1.Top + fraSplit1.Height
        .Width = Me.ScaleWidth - .Left: .Height = Me.ScaleHeight - .Top
    End With
End Sub

Private Sub frmParent_Unload(Cancel As Integer)
    Unload Me
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

Private Sub ShowSeqList()
'��ѯ�������

    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim strCurKey As String
    
    On Error GoTo DBError
    If Not lvwSeq.SelectedItem Is Nothing Then strCurKey = lvwSeq.SelectedItem.Key
    
    With lvwSeq
        With .ColumnHeaders
            .Clear
        
            .Add , , "Ӱ�����", 2000
            .Add , , "����", 800, 1
            .Add , , "���к�", 800, 1
            .Add , , "ͼ����", 800, 1
            .Add , , "˵��", 2500
            .Add , , "�ɼ�ʱ��", 1800
            .Add , , "����", 600
            .Add , , "��Ƭ", 600
        End With
        .ListItems.Add , , "Temp", , 1
        .ListItems.Clear
    End With
    
    strSQL = "Select A.����UID,A.���к�,A.��������,A.�ɼ�ʱ��,B.Ӱ�����,B.����," & _
        "Decode(B.������,1,'��','  ') As ����,Decode(B.���Ž�Ƭ,1,'��','  ') As ��Ƭ,B.���UID,Sum(1) As ͼ���� " & _
        "From Ӱ�������� A,Ӱ�����¼ B,Ӱ���豸Ŀ¼ C,Ӱ����ͼ�� D " & _
        "Where B.ҽ��ID= [1]  And B.���ͺ�= [2] " & _
        " And A.���UID=B.���UID And B.λ��һ=C.�豸��(+) And A.����UID=D.����UID " & _
        "Group By A.����UID,A.���к�,A.��������,A.�ɼ�ʱ��,B.Ӱ�����,B.����," & _
        "Decode(B.������,1,'��','  '),Decode(B.���Ž�Ƭ,1,'��','  '),B.���UID " & _
        "Order By B.Ӱ�����,B.����,A.���к�"
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
    End If
        
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, AdviceID, lngSendNO)
   
    strCheckUID = ""
    If Not rsTmp.EOF Then
        strCheckUID = Nvl(rsTmp("���UID"))
        Do While Not rsTmp.EOF
            
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTmp("����UID"), rsTmp("Ӱ�����"))
            With tmpItem
                .SubItems(1) = rsTmp("����")
                .SubItems(2) = Nvl(rsTmp("���к�"))
                .SubItems(3) = Nvl(rsTmp("ͼ����"), 0)
                .SubItems(4) = Nvl(rsTmp("��������"))
                .SubItems(5) = Nvl(rsTmp("�ɼ�ʱ��"), Date)
                .SubItems(6) = rsTmp("����")
                .SubItems(7) = rsTmp("��Ƭ")
                
                If .Key = strCurKey Then .Selected = True
            End With
            
            rsTmp.MoveNext
        Loop
    End If
    
    DViewer.Images.Clear: iCurImageIndex = 0
    ShowMenu
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwSeq_DblClick()
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    DViewer_DblClick
End Sub

Private Sub lvwSeq_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strSQL As String, lngSeqUID As String
    Dim strURL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double, lngRecID As Long
    Dim curImage As DicomImage, i As Integer, iFrameCount As Integer
    Dim iCols As Integer, iRows As Integer
    Dim bln1stDev As Boolean, objFile As New Scripting.FileSystemObject, strFileName As String, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim ShowPhotoNumber As Long
    
    If Not mblnShowPic Then Exit Sub
        
    ShowPhotoNumber = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ʾͼ����", 20))
        
    bln1stDev = True
    
    On Error GoTo DBError
    Timer1.Enabled = False
    
    lngSeqUID = Mid(Item.Key, 2)
    strSQL = "Select A.ͼ���,D.�û��� As User1,D.���� As Pwd1," & _
        "D.IP��ַ As Host1," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL1,d.�豸�� as �豸��1, " & _
        "E.�û��� As User2,E.���� As Pwd2," & _
        "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL2, e.�豸�� as �豸��2 " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And Rownum<=[2] " & _
        "And A.����UID= [1]  Order By A.ͼ���"
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
            
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngSeqUID, mDispImgs)
    
    Screen.MousePointer = vbHourglass
    pgbLoad.Visible = True: pgbLoad.Value = 10: dblInit = pgbLoad.Value

    With DViewer
        .Images.Clear
        If rsTmp.RecordCount > 0 Then
            .MultiColumns = 1: .MultiRows = 1

            ResizeRegion IIf(ShowPhotoNumber > rsTmp.RecordCount, rsTmp.RecordCount, ShowPhotoNumber), .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows

            lngRecID = 1
            
            ClearCacheFolder strCachePath
            MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
            Do While Not rsTmp.EOF
                If i > ShowPhotoNumber Then Exit Do
                If strDeviceNO1 <> rsTmp("�豸��1") Then
                    strDeviceNO1 = rsTmp("�豸��1")
                    Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
                End If
                
                If strDeviceNO2 <> rsTmp("�豸��2") Then
                    strDeviceNO2 = rsTmp("�豸��2")
                    Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
                End If
                i = i + 1
                If Dir(strCachePath & Nvl(rsTmp("URL1"))) = vbNullString Then
                    strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
'                    Inet.strIPAddress = Nvl(rsTmp("Host1")): Inet.strUser = Nvl(rsTmp("User1")): Inet.strPsw = Nvl(rsTmp("Pwd1"))
                    If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFile.GetFileName(rsTmp("URL1"))) <> 0 Then
                        strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
'                        Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFile.GetFileName(rsTmp("URL2")))
                    End If
                End If
                Set curImage = .Images.ReadFile(strCachePath & Nvl(rsTmp("URL1")))
                DoEvents
                
                With curImage
                    .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
                End With

                pgbLoad.Value = dblInit + (lngRecID / rsTmp.RecordCount) * (100 - dblInit)
                lngRecID = lngRecID + 1

                rsTmp.MoveNext
                
            Loop

            iCurImageIndex = 1: .CurrentIndex = 1
            .Images(iCurImageIndex).BorderColour = vbRed
        Else
            .MultiColumns = 1: .MultiRows = 1: iCurImageIndex = 0
        End If
    End With

    pgbLoad.Visible = False
    Screen.MousePointer = vbDefault
    
    '�ر�FTP����
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    
    Timer1.Enabled = True
    Exit Sub

ReadURLError:
    If bln1stDev Then
        bln1stDev = False
        Resume
    Else
        If ErrCenter() = 1 Then Resume
        pgbLoad.Visible = False
        Screen.MousePointer = vbDefault
        Timer1.Enabled = True
        Call SaveErrLog
    End If
    Exit Sub

DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    pgbLoad.Visible = False
    Screen.MousePointer = vbDefault
    Timer1.Enabled = True
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

Private Sub Timer1_Timer()
    ShowSeqList
    If lvwSeq.SelectedItem Is Nothing Then
        DViewer.Images.Clear
    Else
'        lvwSeq_ItemClick lvwSeq.SelectedItem
    End If
End Sub

Private Sub ShowMenu()
    On Error Resume Next
    If lvwSeq.SelectedItem Is Nothing Then
        frmParent.mnuImageView(0).Enabled = False
        frmParent.mnuImageView(1).Enabled = False
        frmParent.mnuImageView(2).Enabled = False
        frmParent.tbrMain.Buttons("��Ƭ").Enabled = False
        frmParent.tbrMain.Buttons("ȫѡ").Enabled = False
        frmParent.tbrMain.Buttons("ȫ��").Enabled = False
        frmParent.mnuViewPic.Enabled = False
        frmParent.tbrMain.Buttons("��ʾ").Enabled = False
    Else
        frmParent.mnuImageView(0).Enabled = True
        frmParent.mnuImageView(1).Enabled = True
        frmParent.mnuImageView(2).Enabled = True
        frmParent.tbrMain.Buttons("��Ƭ").Enabled = True
        frmParent.tbrMain.Buttons("ȫѡ").Enabled = True
        frmParent.tbrMain.Buttons("ȫ��").Enabled = True
        frmParent.mnuViewPic.Enabled = True
        frmParent.tbrMain.Buttons("��ʾ").Enabled = True
    End If
End Sub
