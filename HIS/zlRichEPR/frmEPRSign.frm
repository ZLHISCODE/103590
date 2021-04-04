VERSION 5.00
Begin VB.Form frmEPRSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��дǩ��"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6330
   Icon            =   "frmEPRSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -270
      TabIndex        =   13
      Top             =   1785
      Width           =   6555
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -270
      TabIndex        =   11
      Top             =   510
      Width           =   6555
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "����ǩ��(&E)"
      Height          =   195
      Left            =   3105
      TabIndex        =   7
      Top             =   1013
      Width           =   1365
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1605
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1387
      Width           =   1365
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   1365
   End
   Begin VB.OptionButton optName 
      Caption         =   "ָ���û�(&U)"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   1320
   End
   Begin VB.OptionButton optName 
      Caption         =   "��ǰ�û�(&C)"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5010
      TabIndex        =   9
      Top             =   1875
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3765
      TabIndex        =   8
      Top             =   1875
      Width           =   1095
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   4110
   End
   Begin VB.PictureBox picǩ��ͼƬ 
      AutoRedraw      =   -1  'True
      Height          =   810
      Left            =   5265
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   12
      Top             =   690
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û�����(&P)"
      Height          =   180
      Left            =   510
      TabIndex        =   5
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1605
      TabIndex        =   10
      Top             =   660
      Width           =   360
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ������(&L)"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmEPRSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '������
Private Sign As cEPRSign                    'ǩ������

Private mlngPassType As Long                 '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�
Private mblnOk As Boolean
Private msSource As String                 '����ǩ����Դ�ַ���
Private mpicSign  As StdPicture
Private morgSign  As StdPicture             'ǩ��ԭʼͼ(��Ա��.ǩ��ͼƬ)
Public Function ShowMe(ByRef edtThis As Editor, ByRef fParent As Object, ByVal sSource As String, ByRef picSign As StdPicture) As cEPRSign
Dim bytFileKind As Byte    '�Ƿ�����
Dim lS As Long, rsTemp As ADODB.Recordset, strUserKind As String

    bytFileKind = fParent.Document.EPRPatiRecInfo.��������
    Set mpicSign = Nothing
    Set morgSign = Nothing
    Set Sign = New cEPRSign
    Set frmParent = fParent
    msSource = sSource
    
    lblUserName.Caption = gstrUserName
    '����ǩ����������ʼ����ǩ������
    Select Case bytFileKind
    Case cpr���Ʊ���
        cmbLevel.AddItem "1 - ҽ��"
        cmbLevel.AddItem "2 - ����"
        cmbLevel.AddItem "3 - ����"
        cmbLevel.ListIndex = 0
        If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 1
        If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 2
    Case Else
        gstrSQL = "Select ��Ա���� From ��Ա����˵�� Where ��Աid = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ա����", glngUserId)
        Do Until rsTemp.EOF
            strUserKind = strUserKind & "," & rsTemp!��Ա����
            rsTemp.MoveNext
        Loop
        
        If InStr(strUserKind, "ҽ��") > 0 And InStr(strUserKind, "��ʿ") = 0 Then '����Աֻ��ҽ��
            cmbLevel.AddItem "1 - ����ҽʦ"
            cmbLevel.AddItem "2 - ����ҽʦ"
            cmbLevel.AddItem "3 - ������ҽʦ"
            cmbLevel.AddItem "4 - ����ҽʦ"
            cmbLevel.ListIndex = 0
            If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 1
            If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 2
            If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 3
        ElseIf InStr(strUserKind, "ҽ��") = 0 And InStr(strUserKind, "��ʿ") > 0 Then '����Աֻ�ǻ�ʿ
            cmbLevel.AddItem "1 - ��ʿ"
            cmbLevel.AddItem "3 - ��ʿ��"
            cmbLevel.ListIndex = 0
            If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 1
        ElseIf bytFileKind = cpr������ Then ''����Ա����ҽ�����ǻ�ʿ �򶼲���ʱ�������ļ���������
            cmbLevel.AddItem "1 - ��ʿ"
            cmbLevel.AddItem "3 - ��ʿ��"
            cmbLevel.ListIndex = 0
            If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 1
        Else
            cmbLevel.AddItem "1 - ����ҽʦ"
            cmbLevel.AddItem "2 - ����ҽʦ"
            cmbLevel.AddItem "3 - ������ҽʦ"
            cmbLevel.AddItem "4 - ����ҽʦ"
            cmbLevel.ListIndex = 0
            If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 1
            If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 2
            If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 3

        End If
    End Select
    
    '��ȡ��ǰǩ����ʽ��ϵͳ����26��
    Select Case bytFileKind
    Case cpr���ﲡ��
        lS = 1
    Case cprסԺ����
        lS = 2
    Case cpr���Ʊ���
        Select Case fParent.Document.EPRFileInfo.lngModule
            Case 1290, 1291, 1294
                lS = 7
            Case Else
                lS = 3
        End Select
    Case cpr������
        lS = 4
    Case Else
        Select Case fParent.Document.EPRPatiRecInfo.������Դ
        Case cprPF_����
            lS = 1
        Case cprPF_סԺ
            lS = 2
        Case Else
            lS = 2  '������סԺΪ׼
        End Select
    End Select
    
    mlngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), lS, 1)) '����,סԺ,ҽ��,����,ҩƷ,LIS,PACS (1111111),Ϊ��Ĭ�ϲ�������ģʽ
    If mlngPassType = 1 Then
        If gstrESign = "" Or (lS = 3 And gstrESign = "0") Then 'ҽ������վ��д����û�е���clsDockxx��,�����ˢ��"סԺ����"ҳ�棬����д�������clsDockInEPR�в���gstrESign = "0"
            gstrESign = getPassESign(3, fParent.Document.EPRPatiRecInfo.����ID)
        End If
        mlngPassType = Val(gstrESign)
    End If
    
    Call optName_Click(0)
    
    Me.Show vbModal, frmParent
    If mblnOk Then
        Set ShowMe = Sign
        If Sign.ǩ��ͼƬ Then
            Set picSign = mpicSign
        Else
            Set picSign = Nothing
        End If
    Else
        Set picSign = Nothing
        Set ShowMe = Nothing
    End If
    Set mpicSign = Nothing
    Set morgSign = Nothing
End Function

'################################################################################################################
'## ���ܣ�  ����ǩ�����ڲ�ǩ���鲢ˢ����ʾ����֤�����������ǩ����
'################################################################################################################
Private Function Validation() As Boolean
    Dim blnSpecify As Boolean, strSpecifySign, lngSpecifyId As Long, lngSpecifyLevel As Long, intSign As Integer
    Dim lngCertID As Long, strSign As String, strʱ��� As String, objSignPic As Object, strʱ��Base64 As String
    Dim rsTemp As ADODB.Recordset, l As Long, strFile As String, strErr As String
    
    On Error GoTo errHand
    intSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0) '��ʾ����������ʾǩ�� ����ͬ�����
    
    If optName(1).Value Then  'ָ���ʺ�ǩ��
        blnSpecify = True
        txtName = Trim(txtName)
        txtPass = Trim(txtPass)
        
        If frmParent.Document.EPRPatiRecInfo.�������� = cprסԺ���� Or frmParent.Document.EPRPatiRecInfo.�������� = cpr���ﲡ�� Or frmParent.Document.EPRPatiRecInfo.�������� = cpr������ Then
            gstrSQL = "Select 1 From �ϻ���Ա�� A, ������Ա B Where a.�û��� = [1] And a.��Աid = b.��Աid And b.����id = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ǩ���û��뵱ǰ�û��Ƿ�ͬ����", UCase(txtName.Text), frmParent.Document.EPRPatiRecInfo.����ID)
            If rsTemp.EOF Then
                MsgBox "ָ��ǩ���û��뵱ǰ������Ա������ͬһ���ң���ֹ�����ÿ��Ҳ��˲�����", vbExclamation, gstrSysName: Exit Function
            End If
        End If
        
        If chkEsign.Value = vbUnchecked Then '����ǩ��
            If Trim(txtPass) = "" Then MsgBox "ָ���ʺ����벻��Ϊ�գ����飡", vbExclamation: Exit Function
            If gobjRegister Is Nothing Then Set gobjRegister = DynamicCreate("zlRegister.clsRegister", "������֤���")
            If Not gobjRegister.LoginValidate("", txtName, txtPass, strErr) Then
                MsgBox "ָ���ʺ�/�������,�����������¼�ʺź����룡" & strErr, vbInformation + vbOKOnly, gstrSysName: Exit Function
            End If
        End If
        
        gstrSQL = "Select b.Id, b.����, b.ǩ��" & vbNewLine & _
                    "From �ϻ���Ա�� A, ��Ա�� B" & vbNewLine & _
                    "Where a.�û��� =[1] And a.��Աid = b.Id And" & vbNewLine & _
                    "      Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'YYYY-MM-DD')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo", UCase(txtName))
        If rsTemp.EOF Then MsgBox "ָ���ʺŲ����ڣ������������¼�ʺź�����!", vbInformation, gstrSysName: Exit Function
        
        If intSign = 0 Then
            strSpecifySign = rsTemp!����
        Else
            strSpecifySign = NVL(rsTemp!ǩ��, rsTemp!����)         '��ʾǩ��
        End If
        lngSpecifyId = rsTemp.Fields("ID")   '�û�ID
        
        lngSpecifyLevel = GetUserSignLevel(lngSpecifyId, rsTemp!����, frmParent.Document.EPRPatiRecInfo.����ID, frmParent.Document.EPRPatiRecInfo.��ҳID) '��ȡָ���û���ǩ������
        If lngSpecifyLevel = cprSL_�հ� Then MsgBox "ָ���ʺ���δ����ǩ������������Ա�����е���Ƹ��ְ��", vbInformation, gstrSysName: Exit Function
        For l = 1 To frmParent.Document.Signs.Count
            If frmParent.Document.Signs(l).ǩ������ > lngSpecifyLevel Then
                MsgBox "��ǰ�������и��߼����ǩ��,��ǰǩ��������Ȩ��ǩ������", vbInformation, gstrSysName: Exit Function
            End If
        Next
    End If
    
    If Not (IIf(blnSpecify, lngSpecifyLevel, frmParent.Document.�û�ǩ������) >= Val(cmbLevel.Text)) Then '
        MsgBox "�û�ӵ�е�ǩ���������ѡ����ǩ������,������ѡ��ǩ������", vbInformation, gstrSysName: Exit Function
    End If

    If chkEsign.Value = vbChecked Then '����ǩ��,�ڴ˴����ж�ǩ��������г�ʼ�����˴��ڹرպ����ݱ��棬��ȡ��������Դ���ݽ���ǩ������ǩ�������ʼ��ʧ���򲻱���
        If gobjESign Is Nothing Then
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            If gobjESign.Initialize(gcnOracle, glngSys) = False Then Exit Function
        End If
        
		If gobjESign.CheckCertificate(IIf(blnSpecify, UCase(txtName), gstrDBUser)) = False Then Exit Function

        If Not gobjESign.CertificateStoped(IIf(blnSpecify, strSpecifySign, gstrUserName)) Then
            strSign = gobjESign.signature(msSource, IIf(chkEsign.Value = vbChecked, IIf(blnSpecify, UCase(txtName), gstrDBUser), ""), lngCertID, strʱ���, objSignPic, strʱ��Base64) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
            If strSign = "" Then MsgBox "����ǩ��ʧ�ܣ����ٴ�ǩ����", vbInformation + vbOKOnly, gstrSysName: Exit Function
        Else
            chkEsign.Value = vbUnchecked
        End If
    End If
    
    Sign.���� = IIf(blnSpecify, strSpecifySign, IIf(intSign = 0, gstrUserName, gstrSignName))
    Sign.ǩ����ID = IIf(blnSpecify, lngSpecifyId, glngUserId)
    Sign.ǩ������ = Val(cmbLevel.Text)
    If Sign.ǩ������ > cprSL_���� Then Sign.ǩ������ = cprSL_����
    
    If zlDatabase.GetPara("��ǩ��������Ϊǰ׺����", glngSys, 1070, "0") = 1 Then
        Sign.ǰ������ = Trim(Mid(Me.cmbLevel.Text, 4)) & "��"
    Else
        Sign.ǰ������ = ""
    End If
    Sign.��ʾ��ǩ = (zlDatabase.GetPara("��ʾ��ǩλ��", glngSys, 1070, "0") = 1)
    Sign.ǩ����ʽ = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.ǩ��ʱ�� = zlDatabase.Currentdate()
    Select Case Val(zlDatabase.GetPara("ǩ��ʱ��", glngSys, 1070, "0"))
        Case 1: Sign.��ʾʱ�� = "yyyy-MM-dd hh:mm"
        Case 2: Sign.��ʾʱ�� = "yyyy��MM��dd�� hh:mm"
        Case Else: Sign.��ʾʱ�� = ""
    End Select
    
    'ǩ������=2 ʹ��RTF.Text��Ϊ����ǩ��ԭ�� ��cEPRSignע��
    Sign.ǩ������ = 2
    Sign.ǩ����Ϣ = strSign
    Sign.֤��ID = lngCertID
    Sign.ʱ��� = strʱ���
    Sign.ʱ�����Ϣ = strʱ��Base64
'    'ǩ������=3 ʹ�ñ������ݿ��������ı�������ǩ��Ҫ�أ�ǩ������,ͼƬ������Ӷ���Ϊ����ǩ��ԭ��
'    '����ǩ����Ϣ�ڱ�����������ǩ���󷵻ز���������
'    Sign.ǩ������ = 3
'    Sign.ǩ����Ϣ = IIf(chkEsign.Value = vbChecked, IIf(blnSpecify, UCase(txtName), gstrDBUser), "") '�������ǩ�����ȴ�ǩ���ʺţ���������ǩ���������,ǩ����ɺ����
'    Sign.֤��ID = 0
'    Sign.ʱ��� = ""

    picǩ��ͼƬ.Cls: Set picǩ��ͼƬ.Picture = Nothing
    picǩ��ͼƬ.Move picǩ��ͼƬ.Left, picǩ��ͼƬ.Top, 810, 810
    
    If zlDatabase.GetPara("ǩ��ʹ��ͼƬ", glngSys, 1070, "0", , , , frmParent.Document.EPRPatiRecInfo.����ID) = 1 Then
        picǩ��ͼƬ.Visible = True
        strFile = zlBlobRead(15, IIf(optName(0).Value, glngUserId, lngSpecifyId), "", False)
        
        If strFile = "" Then
            MsgBox IIf(optName(0).Value, "��ǰ", "ָ��") & "�ʺ�û�п��õ�ǩ��ͼ������ʹ��ͼƬǩ�����ܣ�����ϵ����Ա��", vbExclamation, gstrSysName
            Exit Function
        Else
            Set morgSign = LoadPicture(strFile)
            DrawSignPicture
            Kill strFile
            
            Sign.ǩ��ͼƬ = True
            Set mpicSign = picǩ��ͼƬ.Picture
        End If
        'ʹ��ͼƬǩ���� ����ʾǩ��ǰ׺������ʾǩ��ʱ�䡢����ʾ��ǩ
        Sign.ǰ������ = ""
        Sign.��ʾʱ�� = ""
        Sign.��ʾ��ǩ = False
    Else
        picǩ��ͼƬ.Visible = False
        Set morgSign = Nothing
        Sign.ǩ��ͼƬ = False
        Set mpicSign = Nothing
    End If
    
    Validation = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Validation Then
        mblnOk = True
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmParent = Nothing
End Sub
Private Sub optName_Click(Index As Integer)
    Select Case mlngPassType
    Case 0 '����ǩ��
        chkEsign.Value = vbUnchecked
        chkEsign.Visible = False: chkEsign.Enabled = True
        txtName.Enabled = (Index = 1): txtName.Visible = True
        txtPass.Enabled = (Index = 1): txtPass.Visible = True
        Label2.Visible = True
    Case 1 '1������
        chkEsign.Value = vbChecked
        chkEsign.Move txtPass.Left, txtPass.Top
        chkEsign.Visible = True: chkEsign.Enabled = False
        txtName.Enabled = (Index = 1): txtName.Visible = True
        txtPass.Enabled = False: txtPass.Visible = False
        Label2.Visible = False
    End Select

    If txtName.Enabled And txtName.Visible Then
        txtName.SetFocus
    End If
End Sub
Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkEsign.Value = vbUnchecked Then
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus:  Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
    End If
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Sub txtNames_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < 32 Or KeyAscii > 126 Then KeyAscii = 0
    If InStr("""@\ ", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Sub DrawSignPicture()
Dim lngHeight As Long, lngWidth As Long
    On Error Resume Next
    If Not morgSign Is Nothing Then
        picǩ��ͼƬ.Appearance = 0: picǩ��ͼƬ.BorderStyle = 0
        If zlDatabase.GetPara("ǩ��ʹ��ԭͼ", glngSys, 1070, "1") = 1 Then
            Set picǩ��ͼƬ.Picture = morgSign
            If picǩ��ͼƬ.Width <> 810 Then picǩ��ͼƬ.Move picǩ��ͼƬ.Left, picǩ��ͼƬ.Top, 810, 810
            picǩ��ͼƬ.PaintPicture picǩ��ͼƬ.Picture, 0, 0, picǩ��ͼƬ.ScaleX(picǩ��ͼƬ.Width, vbTwips, vbPixels), picǩ��ͼƬ.ScaleY(picǩ��ͼƬ.Height, vbTwips, vbPixels)
        Else
            lngHeight = zlDatabase.GetPara("ǩ��ͼƬ�߶�", glngSys, 1070, "50")
            lngWidth = CLng(lngHeight * (morgSign.Width / morgSign.Height))
            picǩ��ͼƬ.Move picǩ��ͼƬ.Left, picǩ��ͼƬ.Top, picǩ��ͼƬ.ScaleX(lngWidth, vbPixels, vbTwips), picǩ��ͼƬ.ScaleY(lngHeight, vbPixels, vbTwips)
            picǩ��ͼƬ.PaintPicture morgSign, 0, 0, lngWidth, lngHeight
            Set picǩ��ͼƬ.Picture = picǩ��ͼƬ.Image
        End If
    End If
    Err.Clear
End Sub
