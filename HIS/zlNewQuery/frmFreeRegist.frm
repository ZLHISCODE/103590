VERSION 5.00
Begin VB.Form frmFreeRegist 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   " lblHospital.Caption = GetUnitName + Chr(10) + Chr(13) + ""���������Һ�ϵͳ"""
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrReload 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   6240
   End
   Begin VB.Timer tmrReadCardState 
      Left            =   2400
      Top             =   7560
   End
   Begin VB.PictureBox picReg 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   -720
      ScaleHeight     =   4935
      ScaleWidth      =   9975
      TabIndex        =   6
      Top             =   1080
      Width           =   9975
      Begin zl9NewQuery.ctlButton ctlBack 
         Height          =   915
         Left            =   1920
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1614
         Caption         =   "����"
         AutoSize        =   0   'False
         ButtonHeight    =   800
      End
      Begin zl9NewQuery.ctlButton ctlOK 
         Height          =   915
         Left            =   1920
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1614
         Caption         =   "ȡ��"
         AutoSize        =   0   'False
         ButtonHeight    =   800
      End
      Begin VB.Timer Time 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   2040
         Top             =   1680
      End
      Begin VB.Label lblCard 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   720
         Left            =   3960
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label lblCardID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ˢ�ڶ������֤ȡ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   42
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   840
         Left            =   -120
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   8550
      End
      Begin VB.Image imgBackgroundImg 
         Height          =   4095
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.PictureBox picLine 
      BackColor       =   &H0000C000&
      Height          =   45
      Index           =   1
      Left            =   960
      ScaleHeight     =   45
      ScaleWidth      =   9135
      TabIndex        =   5
      Top             =   6240
      Width           =   9135
   End
   Begin VB.PictureBox picLine 
      BackColor       =   &H0000C000&
      Height          =   45
      Index           =   0
      Left            =   360
      ScaleHeight     =   45
      ScaleWidth      =   9135
      TabIndex        =   4
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   4680
   End
   Begin VB.Image imgExit 
      Height          =   1380
      Left            =   8040
      MouseIcon       =   "frmFreeRegist.frx":0000
      MousePointer    =   4  'Icon
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2700
   End
   Begin VB.Label lblHospital 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ���׹Һ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2445
   End
   Begin VB.Label lblNoBIll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Բ���Ʊ���Ѿ�ʹ���꣬�뵽���ڹҺš�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   5040
      Width           =   15960
   End
   Begin VB.Label lblNoBIll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Բ���Ʊ���Ѿ�ʹ���꣬�뵽���ڹҺš�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Index           =   1
      Left            =   -3360
      TabIndex        =   0
      Top             =   2040
      Width           =   15960
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   525
      Left            =   1320
      TabIndex        =   3
      Top             =   6840
      Width           =   285
   End
End
Attribute VB_Name = "frmFreeRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlng����ID As Long
Private mlngˢ��ʱ�� As Long
Private mStrBillNo As String
'Private mrsReg As ADODB.Recordset
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mdblUpBgColor As Double, mdblDownBgColor As Double
Private mlngTime As Long
Private Type FreeRegistInfo
    lng����ID As Long
    strҽ������ As String
    lngҽ��id As Long
    lng����   As Long
    lng�ű� As Long
    str���� As String
    str��Ŀ As String
    lng��ĿID  As Long
End Type
Private mRegistInfo As FreeRegistInfo
Private Type PatientIdCard
        Name As String
        str���� As String
        Sex As String
        Address As String
        CardId As String
        Nation As String
        Birthday As String
End Type
Private Type PATIENTINFO
    PatientID As String          '���˵�ID
    Name As String
    Sex As String                '��¼���˵��������Ա�
    DoorPost As String
    Age As String                '��¼���˵�����ź�����
    FareClass As String
    strIDCard As String           '���֤��
    str�������� As String
    str������ַ As String
    str���� As String
End Type
Private mIdCardInfo As PatientIdCard
Private mPatient As PATIENTINFO

Private Sub ctlBack_CommandClick()
    Me.Time.Enabled = False
    Call setControlEnabled(False)
End Sub
Private Sub Form_Paint()
 If mdblUpBgColor = 0 And mdblDownBgColor = 0 Then
    Call DrawColorToColor(Me, Me.BackColor, &HFF8080, , True)
 Else
    Call DrawColorToColor(Me, mdblUpBgColor, mdblDownBgColor, , True)
 End If
End Sub

Private Sub Form_Resize()
        On Error Resume Next
        Me.lblHospital.Left = 0
        Me.lblHospital.Top = 20 * Screen.TwipsPerPixelY
        Me.lblHospital.Width = Me.ScaleWidth
        Me.picLine(0).Left = Me.lblHospital.Left
        Me.picLine(0).Top = lblHospital.Top + Me.lblHospital.Height + 20 * Screen.TwipsPerPixelY
        Me.picLine(0).Width = Me.ScaleWidth
        Me.Lblinfo.Left = 0
        Me.Lblinfo.Top = Me.ScaleHeight - 20 * Screen.TwipsPerPixelY - Me.Lblinfo.Height
        Me.Lblinfo.Width = Me.ScaleWidth
        Me.picLine(1).Left = Me.Lblinfo.Left
        Me.picLine(1).Top = Lblinfo.Top - 20 * Screen.TwipsPerPixelY
        Me.picLine(1).Width = Me.ScaleWidth
        With picReg
            .Left = 0
            .Top = picLine(0).Top + picLine(0).Height
            .Height = picLine(1).Top - picLine(0).Top - picLine(0).Height
            .Width = Me.ScaleWidth
        End With
        With lblCardID
            .Left = (Me.picReg.ScaleWidth - .Width) / 2
            .Top = (Me.picReg.ScaleHeight - .Height) / 2
        End With
        Me.imgExit.Top = Me.Lblinfo.Top
     '   Me.imgExit.Height = Me.ScaleHeight - Me.picLine(1).Top - Me.picLine(1).Height - 1 * Screen.TwipsPerPixelY
        Me.imgExit.Left = Me.ScaleWidth - Me.imgExit.Width
        Call SetMsgState
End Sub

Private Sub SetMsgState()
    On Error Resume Next
    Dim wd As Long
    With Me.lblCard
        .Left = IIf(.Width < Me.picReg.ScaleWidth, (Me.picReg.ScaleWidth - .Width) / 2, 0)
        .Top = lblCardID.Top - .Height
    End With
    wd = ctlOK.Width + 20 * Screen.TwipsPerPixelX + ctlBack.Width
    With Me.ctlOK
        .Left = (Me.picReg.ScaleWidth - wd) / 2
        .Top = lblCardID.Top + lblCardID.Height
    End With
    With Me.ctlBack
        .Left = ctlOK.Left + 20 * Screen.TwipsPerPixelX + ctlOK.Width
        .Top = lblCardID.Top + lblCardID.Height
    End With
End Sub
 
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
   Call InitPicture
   Call Reload
End Sub

Private Sub Reload()
    On Error GoTo hErr
    Call InitFreeRegistMsg
    Call Form_Resize
    mlngˢ��ʱ�� = Val(GetPara("������ˢ������", 0)) * 60
    If mlngˢ��ʱ�� = 0 Then mlngˢ��ʱ�� = 3000
    mlngTime = Val(GetPara("������֤����ͣ��ʱ��")) / 2
    Me.tmrReload.Interval = 1000: Me.tmrReload.Enabled = True
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hwnd)
    End If
    ctlBack.ShowPicture = False
    ctlOK.ShowPicture = False
     If InitBill() = False Then Exit Sub
     If InitFreeRegist() = False Then Exit Sub
     If Me.Lblinfo.Caption = "" Then Lblinfo.Caption = " �Һ���ĿΪ-" & _
                    Nvl(mRegistInfo.str����) & "/" & Nvl(mRegistInfo.str��Ŀ) & IIf(IsNull(mRegistInfo.strҽ������), "", "/" & Nvl(mRegistInfo.strҽ������))
 
     Call setControlEnabled(False)
     Me.tmrReadCardState.Enabled = False
     Me.tmrReadCardState.Interval = 300
  
 Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub
Public Sub ShowMe(objParent As Object)
    Me.Hide
    Me.Show , objParent
End Sub

Private Function setControlEnabled(blnAllowReg As Boolean)
    If mRegistInfo.lng�ű� < 0 And (Not mobjIDCard Is Nothing) Then mobjIDCard.SetEnabled False
    Me.lblCardID.Visible = Not blnAllowReg
    lblCard.Visible = blnAllowReg
    Time.Tag = Val(mlngTime)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Not blnAllowReg)
    Time.Enabled = blnAllowReg
    lblCard.Visible = blnAllowReg
    Me.ctlOK.Visible = blnAllowReg
    Me.ctlBack.Visible = blnAllowReg
    Me.tmrReadCardState.Enabled = blnAllowReg
     
End Function

 

 

Private Sub imgExit_DblClick()
    If frmExitPsw.ShowPsw(Me, True) Then
         Unload frmMainQuery
    End If
End Sub

Private Sub picReg_Paint()
     If Me.imgBackgroundImg.Picture = 0 Then
            Call DrawColorToColor(Me.picReg, mdblUpBgColor, mdblDownBgColor, , True)
    End If
'   Call InitPicture
  'Call ShowPicture
End Sub

Private Sub picReg_Resize()
 On Error Resume Next
 Me.imgBackgroundImg.Left = 0
 Me.imgBackgroundImg.Top = 0
 Me.imgBackgroundImg.Width = Me.picReg.ScaleWidth
 Me.imgBackgroundImg.Height = Me.picReg.ScaleHeight
' Me.picBack.Left = Me.imgBackgroundImg.Left
' Me.picBack.Top = Me.imgBackgroundImg.Top
' Me.picBack.Width = Me.imgBackgroundImg.Width
' Me.picBack.Height = Me.imgBackgroundImg.Height
' LoadBackGroundPicture
' Me.picShow.Width = picReg.Width
' Me.picShow.Height = Me.picReg.Height
End Sub

 Private Sub Time_Timer()
 '-----------------------
 'ˢ���֤�� ��ϢĬ����ʾʱ��
 '����ʱ�� Ĭ��Ϊȡ���Һ�
 '-----------------------
   
    Time.Tag = Val(Time.Tag) - 1
    If Val(Time.Tag) <= 0 Then
       setControlEnabled False: Time.Enabled = False
    End If
End Sub
Private Function InitBill() As Boolean
'Ʊ�����ü�鼰��ʼ
     Dim i As Integer
      mlng����ID = CheckUsedBill(4, IIf(mlng����ID > 0, mlng����ID, glng�Һ�ID))
      If mlng����ID <= 0 Then
          picReg.Visible = False
          ShowErrMsg "�Բ���Ʊ���Ѿ�ʹ���꣬�뵽���ڹҺš�"
          InitBill = False
          Exit Function
      End If
      LblNoBill(0).Visible = False
      LblNoBill(1).Visible = False
      picReg.Visible = True
      InitBill = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not (mobjIDCard Is Nothing) Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    tmrReadCardState.Enabled = False
    Time.Enabled = False
    Timer1.Enabled = False
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
     On Error GoTo hErr
    If Me.Time.Enabled Then Exit Sub
    Me.lblCard.Caption = "��ӭ" & strName & "(" & IIf(strSex = "Ů", "Ůʿ", "����") & ") ����Ժ���" & _
              IIf(Screen.Width / Screen.TwipsPerPixelX <= 800, vbCrLf, "") & "��ȡ��"
    If GetControlTextWidth(lblCard, Me.lblCard.Caption) > Me.ScaleWidth Then
            Me.lblCard.Caption = GetNewLineString(Me.lblCard.Caption, TextWidth(Me.lblCard.Caption))
    End If
    With mIdCardInfo
           .CardId = strID
           .Name = strName
           .Sex = strSex
           .Address = strAddress
           .Birthday = Format(datBirthday, "yyyy-mm-dd")
           .Nation = strNation
    End With
    Call SetMsgState
    setControlEnabled True
    Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub ctlOK_CommandClick()
    On Error GoTo hErr
    Me.ctlOK.Enabled = False
    Me.ctlBack.Enabled = False
    Time.Enabled = False
    If Save������Ϣ() = False Then frmShowMessage.ShowMe Me, "������Ϣ����ʧ�ܣ�":   GoTo ctlState
    If SaveData() = False Then frmShowMessage.ShowMe Me, "�Һ�ʧ���ˣ�": GoTo ctlState
    frmClose.ShowForm Me, mPatient.Name, mStrBillNo
    Call InitBill
   GoTo ctlState
Exit Sub
hErr:
If ErrCenter() = 1 Then Resume
    SaveErrLog
 Exit Sub
ctlState:
   setControlEnabled False: Me.ctlOK.Enabled = True: Me.ctlBack.Enabled = True:
End Sub
Private Sub Timer1_Timer()
'���������ʾ��
 On Error Resume Next
    If LblNoBill(0).Left + LblNoBill(0).Width > 0 Then
        LblNoBill(0).Left = LblNoBill(0).Left - 100
    Else
        LblNoBill(0).Left = LblNoBill(1).Left + LblNoBill(1).Width
    End If
    If LblNoBill(1).Left + LblNoBill(1).Width > 0 Then
        LblNoBill(1).Left = LblNoBill(1).Left - 100
    Else
        LblNoBill(1).Left = LblNoBill(0).Left + LblNoBill(0).Width
    End If
End Sub

Private Sub ShowErrMsg(strMsg As String)
    Me.picReg.Visible = True
    Me.lblCardID.Caption = strMsg
    With lblCardID
        .Left = (Me.picReg.ScaleWidth - .Width) / 2
        .Top = (Me.picReg.ScaleHeight - .Height) / 2
    End With
    Me.lblCardID.Visible = True
End Sub

 Private Function GetNewLineString(ByVal strMsg As String, ByVal lngWidth As Long) As String
  Dim strTmp As String, lngFontWidth As Long
  Dim lngTmp As Long
  If TextWidth(strMsg) < lngWidth Then GetNewLineString = strMsg: Exit Function
  lngFontWidth = TextWidth("��")
  For lngTmp = 1 To TextWidth(strMsg) / lngWidth
      strTmp = strTmp & IIf(lngTmp = 1, "", vbCrLf) & Mid$(strMsg, 1, lngWidth / lngFontWidth - IIf(lngWidth / lngFontWidth > 2, 1, 0))
      strMsg = Mid$(strMsg, lngWidth / lngFontWidth + IIf(lngWidth / lngFontWidth > 2, 0, 1))
  Next
   If strMsg <> "" Then strTmp = strTmp & vbCrLf & strMsg
   GetNewLineString = strTmp
 End Function
 
Private Function InitFreeRegist() As Boolean
   Dim strSQL As String, strMsg As String
   Dim rsReg As ADODB.Recordset
      mRegistInfo.lng�ű� = Val(GetPara("�򵥹Һźű�", -1))
   If mRegistInfo.lng�ű� < 0 Then
       ShowErrMsg "��ǰû�п���ʹ�õĹҺ���Ŀ��"
        InitFreeRegist = False
        Exit Function
   End If
   
    strSQL = "" & _
    "      Select a.Id, a.����, a.���� As �ű�, a.����id, a.��Ŀid,d.���� as ����, b.���� as ��Ŀ, a.ҽ������, a.ҽ��id, b.����, c.�ּ�,Nvl(b.��Ŀ����,0) as ����" & _
    "      From �ҺŰ��� A, �շ���ĿĿ¼ B, �շѼ�Ŀ C,���ű� D " & _
    "      Where a.���� = [1] And a.��Ŀid = b.ID And b.ID = c.�շ�ϸĿid And a.����Id=d.Id And Nvl(a.ͣ������, Sysdate + 1) > Sysdate " & vbNewLine & _
    "        And Not Exists(Select 1 From �ҺŰ���ͣ��״̬ Where ����ID=A.ID and Sysdate between ��ʼֹͣʱ�� and ����ֹͣʱ�� )" & vbNewLine & _
    "        And sysDate Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "      Union All" & vbNewLine & _
    "      Select a.Id, a.����, a.���� As �ű�, a.����id, a.��Ŀid, e.���� as ����, b.���� as ��Ŀ, a.ҽ������, a.ҽ��id, b.����, c.�ּ�,nvl(b.��Ŀ����,0) as ����" & _
    "      From �ҺŰ��� A, �շ���ĿĿ¼ B, �շѼ�Ŀ C, �շѴ�����Ŀ D,���ű� E " & _
    "      Where a.���� = [1] And a.����ID=E.Id And a.��Ŀid = d.����id And d.����id = b.ID And b.ID = c.�շ�ϸĿid And Nvl(a.ͣ������, Sysdate + 1) > Sysdate" & vbNewLine & _
    "        And Not Exists(Select 1 From �ҺŰ���ͣ��״̬ Where ����ID=A.ID and Sysdate between ��ʼֹͣʱ�� and ����ֹͣʱ�� )" & vbNewLine & _
    "        And sysDate Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))"
   On Error GoTo hErr
    Set rsReg = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mRegistInfo.lng�ű�)
     If rsReg.EOF Then
          ShowErrMsg "�Բ��𣬵�ǰ�ű��Ѿ�ͣ�ã��뵽���ڹҺš�": mRegistInfo.lng�ű� = -1: Exit Function
     End If
     With mRegistInfo
        .lng�ű� = Val(Nvl(rsReg!�ű�, 0))
        .lng���� = Val(rsReg!����)
        .lng����ID = Val(Nvl(rsReg!����Id, 0))
        .lngҽ��id = Val(Nvl(rsReg!ҽ��ID, 0))
         .strҽ������ = Nvl(rsReg!ҽ������)
        .str���� = Nvl(rsReg!����)
        .lng��ĿID = Val(Nvl(rsReg!��ĿId))
        .str��Ŀ = Nvl(rsReg!��Ŀ)
     End With
     InitFreeRegist = True
  Exit Function
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

 


Private Function Save������Ϣ() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ���� �����ڱ�����˽�������
    '���:
    '����:
    '����:�ɹ�,����true,����False
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    Dim strSQL As String, lng����ID As Long, intType As Integer, str����� As String
    Dim rsTemp As New ADODB.Recordset, str�������� As String, str���� As String
    Dim strNow As String, strҽ�Ƹ��ʽ As String, blnUpdatePatient As Boolean
    Dim lng���� As Long, strҽ�� As String
    'δˢ��,������
    If mIdCardInfo.CardId = "" Then Exit Function
    strSQL = " " & _
   "   Select " & _
   "          ����ID,����,�Ա�,�����,����,�ѱ�,Trunc(��������) as ��������,��ͥ��ַ,����,����,�ѱ�,ҽ�Ƹ��ʽ,ҽ����  " & _
   "   From ������Ϣ  where ���֤��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mIdCardInfo.CardId)
    If rsTemp.EOF = False Then
        mPatient.Name = zlCommFun.Nvl(rsTemp!����)
        mPatient.DoorPost = zlCommFun.Nvl(rsTemp!�����, 0)
        mPatient.Sex = zlCommFun.Nvl(rsTemp!�Ա�)
        mPatient.Age = zlCommFun.Nvl(rsTemp!����)
        mPatient.FareClass = zlCommFun.Nvl(rsTemp!�ѱ�)
        mPatient.strIDCard = mIdCardInfo.CardId
        mPatient.PatientID = zlCommFun.Nvl(rsTemp!����id)
        mPatient.str�������� = Format(rsTemp!��������, "yyyy-mm-dd")
        mPatient.str������ַ = Nvl(rsTemp!��ͥ��ַ)
        mPatient.str���� = Nvl(rsTemp!����)
        '���ڵĻ����Ͳ�������
        If mPatient.DoorPost <> 0 Then Save������Ϣ = True: Exit Function
        lng����ID = zlCommFun.Nvl(rsTemp!����id, 0)
        strҽ�Ƹ��ʽ = Nvl(rsTemp!ҽ�Ƹ��ʽ)
        mPatient.FareClass = Nvl(rsTemp!�ѱ�)
        mPatient.PatientID = lng����ID
        mIdCardInfo.str���� = zlCommFun.Nvl(rsTemp!����)
        str���� = mIdCardInfo.str����
        lng���� = zlCommFun.Nvl(rsTemp!����, 0)
        strҽ�� = zlCommFun.Nvl(rsTemp!ҽ����, "")
        blnUpdatePatient = True
          
    End If
    strNow = "To_Date('" & CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:MM:ss")) & _
                             "','YYYY-MM-DD HH24:MI:SS')"
    If Not blnUpdatePatient Then   '�½�������Ϣ
         '��ȡĬ�ϵķѱ�
         strSQL = "Select ���� From �ѱ� Where ȱʡ��־ = 1 And ������� In (1, 3) Order By �������,����"
         Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
         If Not rsTemp.EOF Then
             mPatient.FareClass = Nvl(rsTemp!����)
         End If
         strSQL = "Select ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1"
         Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
         If Not rsTemp.BOF Then
             strҽ�Ƹ��ʽ = Nvl(rsTemp!����)
         End If
         '�²���,�Ƚ���
         lng����ID = zlDatabase.GetNextNo(1): mPatient.PatientID = lng����ID
          If IsDate(mIdCardInfo.Birthday) Then
             strSQL = "Select (Sysdate-to_date([1],'yyyy-mm-dd'))/365 As �� From dual"
             Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mIdCardInfo.Birthday)
             str���� = Format(Int(Val(Nvl(rsTemp!��))), "###0") & "��"
             mIdCardInfo.str���� = str����
         Else
             str���� = ""
         End If
         If Exist����ID(lng����ID) = False Then
            '֮���Լ�� �Ƿ������ͬ��ID ��Ϊ�˼��ٲ�����ԭ��
            '��ɵ���ͬ�Ĳ���ID
            lng����ID = zlDatabase.GetNextNo(1): mPatient.PatientID = lng����ID
         End If
    End If
    str����� = Nvl(zlDatabase.GetNextNo(3), 0)
    'Ϊ�˱��Ⲣ����ԭ��
    If Exist�����(str�����, lng����ID) Then str����� = Nvl(zlDatabase.GetNextNo(3), 0)
    
    '  --�������ͣ�
    '  --             1=�½�������Ϣ�����ﲡ��(�����¹ҺŲ���)
    '  --             2=�޸Ĳ�����Ϣ���½����ﲡ��(�����޲����Ĳ���)
    '  --             3=�޸Ĳ�����Ϣ�����������ﲡ��(�����в����Ĳ���,�������޸��˲����������)
    '  --����ҩ��ָ���ʽ��"ID~����~~ID~����...",�������޸Ĳ�����Ϣʱ�á�
    
    'Zl_�ҺŲ��˲���_Insert
    strSQL = "Zl_�ҺŲ��˲���_Insert("
    '  ��������_In     Number,
    strSQL = strSQL & "" & IIf(blnUpdatePatient, 2, 1) & ","
    '  ����id_In       ������Ϣ.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �����_In       ������Ϣ.�����%Type,
    strSQL = strSQL & "" & str����� & ","
    '  ���￨��_In     ������Ϣ.���￨��%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����֤��_In     ������Ϣ.����֤��%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����_In         ������Ϣ.����%Type,
    strSQL = strSQL & "'" & mIdCardInfo.Name & "',"
    '  �Ա�_In         ������Ϣ.�Ա�%Type,
    strSQL = strSQL & "'" & mIdCardInfo.Sex & "',"
    '  ����_In         ������Ϣ.����%Type,
    strSQL = strSQL & "" & IIf(str���� = "", "NULL", "'" & str���� & "'") & ","
    '  �ѱ�_In         ������Ϣ.�ѱ�%Type,
    strSQL = strSQL & "'" & mPatient.FareClass & "',"
    '  ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    strSQL = strSQL & "" & IIf(strҽ�Ƹ��ʽ = "", "NULL", "'" & strҽ�Ƹ��ʽ & "'") & ","
    '  ����_In         ������Ϣ.����%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����_In         ������Ϣ.����%Type,
    strSQL = strSQL & "'" & mIdCardInfo.Nation & "',"
    '  ����_In         ������Ϣ.����״��%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ְҵ_In         ������Ϣ.ְҵ%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ���֤��_In     ������Ϣ.���֤��%Type,
    strSQL = strSQL & "'" & mIdCardInfo.CardId & "',"
    '  ������λ_In     ������Ϣ.������λ%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��ͬ��λid_In   ������Ϣ.��ͬ��λid%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��λ�绰_In     ������Ϣ.��λ�绰%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��λ�ʱ�_In     ������Ϣ.��λ�ʱ�%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    strSQL = strSQL & "'" & mIdCardInfo.Address & "',"
    '  ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  �����ʱ�_In     ������Ϣ.�����ʱ�%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  �Ǽ�ʱ��_In     ������Ϣ.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "" & strNow & ","
    '  ����ҩ��_In     Varchar2,
    strSQL = strSQL & "" & "NULL" & ","
    '  �Һŵ�_In       ���˹Һż�¼.NO%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��������_In     ������Ϣ.��������%Type := Null,
    If IsDate(mIdCardInfo.Birthday) Then
        strSQL = strSQL & "to_date('" & mIdCardInfo.Birthday & "','yyyy-mm-dd'),"
    Else
        strSQL = strSQL & "" & "null" & ","
    End If
    '  ҽ����_In       ������Ϣ.ҽ����%Type := Null,
    strSQL = strSQL & "" & IIf(strҽ�� = "", "NULL", "'" & strҽ�� & "'") & ","
    '  Ic����_In       ������Ϣ.Ic����%Type := Null
    strSQL = strSQL & "" & "NULL" & ","
   '  ����_In         ������Ϣ.����%Type := Null
    strSQL = strSQL & IIf(blnUpdatePatient, IIf(lng���� = 0, "null", lng����), "null") & ")"
    
    Err = 0: On Error GoTo errHand:
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    With mPatient
        .Name = mIdCardInfo.Name
        .PatientID = lng����ID
        .Sex = mIdCardInfo.Sex
        .Age = str����
        .str�������� = Format(mIdCardInfo.Birthday, "yyyy-mm-dd")
        .strIDCard = mIdCardInfo.CardId
        .str������ַ = mIdCardInfo.Address
        .str���� = mIdCardInfo.Nation
        .DoorPost = str�����
        
    End With
    Save������Ϣ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function


Private Function SaveData() As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset, rsPati As New ADODB.Recordset, rs As New ADODB.Recordset
    Dim aryItem As Variant, strNo As String, i As Integer, str�վݷ�Ŀ As String
    Dim StrRoom As String, strBed As String, str�ѱ� As String, strTmp As String, strNow As String
    Dim cllReg As Collection, strSQL As String
    Err = 0: On Error GoTo ErrHandle:
   ' If mblnCanCommit = False Then Exit Sub
  
    
    strNow = "To_Date('" & CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:MM:ss")) & "','YYYY-MM-DD HH24:MI:SS')"
    '�����ǰ��Ͽ���
    '------------------------------------------------------------------------------------------------------------------
    StrRoom = GetRoom(Nvl(mRegistInfo.lng�ű�))
    If StrRoom = "" Then
        StrRoom = "null"
    Else
        StrRoom = "'" + StrRoom + "'"
    End If
        
    Set cllReg = New Collection
    gstrSQL = "Select C.���� as ������" & _
                " From ������Ϣ A,ҽ�Ƹ��ʽ C" & _
                " Where A.����ID=[1] " & _
                " And A.ҽ�Ƹ��ʽ=C.����(+)"
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mPatient.PatientID))
    If rsPati.BOF = False Then strBed = zlCommFun.Nvl(rsPati("������").Value)
    
    '�ϸ�ȡ��һ������
     mStrBillNo = GetNextBill(mlng����ID)
    
     If mStrBillNo = "" Then
        'û�п��õ�Ʊ��
     End If
     strNo = zlDatabase.GetNextNo(12)
     aryItem = GetFreeRegistPrice(Val(mRegistInfo.lng��ĿID))
     'On Error GoTo ErrHandle
            
    '------------------------------------------------------------------------------------------------------------------
    For i = 0 To UBound(aryItem)
        
        gstrSQL = "Select �վݷ�Ŀ From ������Ŀ where ID =[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(aryItem(i, 1)))
        If Not rsTmp.BOF And Not IsNull(rsTmp("�վݷ�Ŀ")) Then str�վݷ�Ŀ = CStr(rsTmp("�վݷ�Ŀ"))
        '��ȡ�ѱ�ʵ�ս��
        str�ѱ� = mPatient.FareClass
        gstrSQL = VB_���˹Һż�¼_Insert(mPatient.PatientID, mPatient.DoorPost, mPatient.Name, mPatient.Sex, mPatient.Age, Val(strBed), mPatient.FareClass, strNo, mStrBillNo, _
                                   i + 1, Val(aryItem(i, 4)), CLng(aryItem(i, 5)), Format(Nvl(aryItem(i, 0)), "0.00"), CLng(aryItem(i, 1)), _
                                   str�վݷ�Ŀ, "", Val(aryItem(i, 0)), 0, _
                                 Val(mRegistInfo.lng����ID), Val(mRegistInfo.lng����ID), strNow, strNow, Nvl(mRegistInfo.strҽ������), Val(Nvl(mRegistInfo.lngҽ��id)), Val(Nvl(mRegistInfo.lng����)), Nvl(mRegistInfo.lng�ű�), StrRoom, 0, mlng����ID, 0, _
                                0, 0, 0, 0, Val(aryItem(i, 6)), Val(aryItem(i, 7)))
        zlAddArray cllReg, gstrSQL
    Next
    '����:31187:��Ҫ�ǽ��ҺŻ��ܵ�������
    If mRegistInfo.lng�ű� >= 0 Then
        strSQL = "zl_���˹ҺŻ���_Update("
        '  ҽ������_In   �ҺŰ���.ҽ������%Type,
        strSQL = strSQL & "'" & Nvl(mRegistInfo.strҽ������) & "',"
        '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
        strSQL = strSQL & "" & IIf(Val(Nvl(mRegistInfo.lngҽ��id)) = 0, "NULL", mRegistInfo.lngҽ��id) & ","
        '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
        strSQL = strSQL & "" & Nvl(mRegistInfo.lng��ĿID, 0) & ","
        '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
        strSQL = strSQL & "" & Nvl(mRegistInfo.lng����ID, 0) & ","
        '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
        strSQL = strSQL & "" & strNow & ","
        '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����
        strSQL = strSQL & "" & 0 & ","
        ' ����_In       �ҺŰ���.����%Type := Null
        strSQL = strSQL & mRegistInfo.lng�ű� & ")"
        
        Call zlAddArray(cllReg, strSQL)
    End If
    gblnBeginTrans = True
    zlExecuteProcedureArrAy cllReg, Me.Caption, False, False
    gblnBeginTrans = False
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1111", Me, "NO=" & strNo, 2)
    'Call frmClose.ShowForm(Me, mPatient.Name, strNo)
    SaveData = True
    Exit Function
    '-----------------------------------------------------------------------------------------------------------------
 
ErrHandle:
    If gblnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    gblnBeginTrans = False
    Call SaveErrLog
End Function

Private Function VB_���˹Һż�¼_Insert(ByVal lng����ID As String, ByVal lng����� As String, ByVal str���� As String, ByVal str�Ա� As String, ByVal str���� As String, _
    ByVal str���� As String, ByVal str�ѱ� As String, ByVal str���ݺ� As String, ByVal strƱ�ݺ� As String, ByVal int��� As String, ByVal lng���� As Long, ByVal lng�շ�ϸĿid As String, _
    ByVal db��׼���� As String, ByVal lng������Ŀid As String, ByVal str�վݷ�Ŀ As String, ByVal str���㷽ʽ As String, ByVal dbӦ�ս�� As String, ByVal dbʵ�ս�� As String, _
    ByVal lng���˿���id As String, ByVal lngִ�в���id As String, ByVal str����ʱ�� As String, ByVal str�Ǽ�ʱ�� As String, ByVal strҽ������ As String, ByVal lngҽ��id As String, _
    ByVal lng���� As Long, ByVal str�ű� As String, ByVal str��ҩ���� As String, ByVal lng����id As String, ByVal lng����ID As String, _
    ByVal CurԤ��֧�� As String, ByVal Curҽ��֧�� As String, ByVal str���մ���ID As String, ByVal str������Ŀ�Ƿ� As String, _
    ByVal strͳ���� As String, ByVal int�۸񸸺� As Integer, ByVal int�������� As Integer) As String
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    
    Dim strSQL As String, bln���ɶ��� As Boolean
    bln���ɶ��� = Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, 1113)) <> 0
 
    'Zl_���˹Һż�¼_Insert
    strSQL = "zl_���˹Һż�¼_Insert("
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �����_In     ������ü�¼.��ʶ��%Type,
    strSQL = strSQL & "" & lng����� & ","
    '  ����_In       ������ü�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  �Ա�_In       ������ü�¼.�Ա�%Type,
    strSQL = strSQL & "'" & str�Ա� & "',"
    '  ����_In       ������ü�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ���ʽ_In   ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
    strSQL = strSQL & "'" & Val(str����) & "',"
    '  �ѱ�_In       ������ü�¼.�ѱ�%Type,
    strSQL = strSQL & "'" & str�ѱ� & "',"
    '  ���ݺ�_In     ������ü�¼.NO%Type,
    strSQL = strSQL & "'" & str���ݺ� & "',"
    '  Ʊ�ݺ�_In     ������ü�¼.ʵ��Ʊ��%Type,
    strSQL = strSQL & "'" & strƱ�ݺ� & "',"
    '  ���_In       ������ü�¼.���%Type,
    strSQL = strSQL & "" & int��� & ","
    '  �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
    strSQL = strSQL & "" & IIf(int�۸񸸺� = 0, "NULL", int�۸񸸺�) & ","
    '  ��������_In   ������ü�¼.��������%Type,
    strSQL = strSQL & "" & IIf(int�������� = 0, "NULL", int��������) & ","
    '  �շ����_In   ������ü�¼.�շ����%Type,
    strSQL = strSQL & "'1',"
    '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
    strSQL = strSQL & "" & lng�շ�ϸĿid & ","
    '  ����_In       ������ü�¼.����%Type,
    strSQL = strSQL & "" & lng���� & ","
    '  ��׼����_In   ������ü�¼.��׼����%Type,
    strSQL = strSQL & "" & db��׼���� & ","
    '  ������Ŀid_In ������ü�¼.������Ŀid%Type,
    strSQL = strSQL & "" & lng������Ŀid & ","
    '  �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
    strSQL = strSQL & "'" & str�վݷ�Ŀ & "',"
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
    strSQL = strSQL & "" & dbӦ�ս�� & ","
    '  ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & dbʵ�ս�� & ","
    '  ���˿���id_In ������ü�¼.���˿���id%Type,
    strSQL = strSQL & "" & lng���˿���id & ","
    '  ��������id_In ������ü�¼.��������id%Type,
    strSQL = strSQL & "" & UserInfo.����ID & ","
    '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
    strSQL = strSQL & "" & lngִ�в���id & ","
    '  ����Ա���_In ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
    strSQL = strSQL & "" & str����ʱ�� & ","
    '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "" & str�Ǽ�ʱ�� & ","
    '  ҽ������_In   �ҺŰ���.ҽ������%Type,
    strSQL = strSQL & "'" & strҽ������ & "',"
    '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
    strSQL = strSQL & "" & IIf(lngҽ��id = 0, "NULL", lngҽ��id) & ","
    '  ������_In Number, --������¼�Ƿ���������
    strSQL = strSQL & "0,"
    '  ����_In       Number,
    strSQL = strSQL & lng���� & ","
    '  �ű�_In       �ҺŰ���.����%Type,
    strSQL = strSQL & "'" & str�ű� & "',"
    '  ����_In       ������ü�¼.��ҩ����%Type,
    strSQL = strSQL & "" & str��ҩ���� & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & IIf(lng����id = 0, "NULL", lng����id) & ","
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
    '  Ԥ��֧��_In   ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
    strSQL = strSQL & "" & Round(Val(CurԤ��֧��), 2) & ","
    '  �ֽ�֧��_In   ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
    strSQL = strSQL & "" & 0 & ","
    '  ����֧��_In   ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
    strSQL = strSQL & "" & Round(Val(Curҽ��֧��), 2) & ","
    '  ���մ���id_In ������ü�¼.���մ���id%Type,
    strSQL = strSQL & "" & IIf(Val(str���մ���ID) = 0, "NULL", str���մ���ID) & ","
    '  ������Ŀ��_In ������ü�¼.������Ŀ��%Type,
    strSQL = strSQL & "" & Val(str������Ŀ�Ƿ�) & ","
    '  ͳ����_In   ������ü�¼.ͳ����%Type,
    strSQL = strSQL & "" & Val(strͳ����) & ","
    '  ժҪ_In       ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
    strSQL = strSQL & "NULL,"
    '  ԤԼ�Һ�_In   Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
    strSQL = strSQL & "0,"
    '  �շ�Ʊ��_In   Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
    strSQL = strSQL & "0,"
    '  ���ձ���_In   ������ü�¼.���ձ���%Type,
    strSQL = strSQL & "NULL,"
    '  ����_In       ���˹Һż�¼.����%Type := 0,
    strSQL = strSQL & "0,"
    '  ����_In       �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
    strSQL = strSQL & "NULL,"
    '  ����_In       ���˹Һż�¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ԤԼ����_In   Number := 0,
    strSQL = strSQL & "0,"
    '  ԤԼ��ʽ_In   ԤԼ��ʽ.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ���ɶ���_In Number:=0
    strSQL = strSQL & IIf(bln���ɶ���, 1, 0) & ")"
    VB_���˹Һż�¼_Insert = strSQL
    
End Function

 
Private Sub tmrReadCardState_Timer()
    '----------------------------------
    '�����������֤�������Ƿ��Զ�����
    '��Idcard����ShowIDCardInfo�¼���
    '�������Զ�ˢ��
    '----------------------------------
    If Me.ctlOK.Visible Then
        mobjIDCard.SetEnabled False
    End If
End Sub

Private Sub tmrReload_Timer()
'---------------------------------
'ˢ�¹Һ���Ϣ
'---------------------------------
    Static lngTime As Double
    If mlngˢ��ʱ�� = 0 Then mlngˢ��ʱ�� = 6000
    If lngTime >= mlngˢ��ʱ�� Then
        If Me.Time.Enabled Then Exit Sub
         'frmFlash.Show , Me
          Call Reload
         lngTime = 0
         'Unload frmFlash
    Else
        If Me.Time.Enabled = False Then lngTime = lngTime + 1
    End If
End Sub

Private Function GetControlTextWidth(objControl As Control, strTxt As String) As Double
    '--------------------------------------------------------------------------------
    '��ȡ�ؼ����ı���Ӧ���еĿ��
    '--------------------------------------------------------------------------------
    
    Dim lngFont As Long
    lngFont = Me.Font.Size
    Me.Font.Size = objControl.Font.Size
    GetControlTextWidth = TextWidth(strTxt)
    If lngFont <> 0 Then Me.Font.Size = lngFont
End Function

Private Function GetFreeRegistPrice(ByVal lng��ĿID) As Variant
    '******************************************************************************************************************
    '���ܣ�����ָ���Һ����ͣ���ָ��ʱ��ļ۸��ά�����У����顣
    '   ��һ��Ϊ�۸񣬵ڶ��б�ʾ������ĿID����������д������Ŀ,������Ϊ���㵥λ,������Ϊ����,������Ϊ�շ�ϸĿID,������(�۸����),�ڰ���(��������)
    '������lng��ĿID=�Һ���ĿID(�շ�ϸĿID)
    '���أ�����
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim aryTmp(), i As Integer
    Dim int���� As Integer, int���� As Integer, lng������Ŀid As Long
    On Error GoTo errH

    gstrSQL = "Select 1 as ����,A.���,A.ID as ��ĿID,A.���㵥λ,B.������ĿID,1 as ����,C.�վݷ�Ŀ,B.�ּ�" & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=[1] " & _
        " And ((To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS') Between To_Char(B.ִ������,'YYYY-MM-DD HH24:MI:SS') And To_Char(B.��ֹ����,'YYYY-MM-DD HH24:MI:SS')) or (To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS')>=To_Char(B.ִ������,'YYYY-MM-DD HH24:MI:SS') And (B.��ֹ���� is NULL Or B.��ֹ����=To_Date('3000-01-01','YYYY-MM-DD'))))"
    gstrSQL = gstrSQL & " Union ALL " & _
        "Select 2 as ����,A.���,A.ID as ��ĿID,A.���㵥λ,C.ID as ������ĿID,D.�������� as ����,C.�վݷ�Ŀ,B.�ּ�" & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շѴ�����Ŀ D" & _
        " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=[1]" & _
        "        And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
        ""
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", lng��ĿID)
    If rs.EOF Then
        GetFreeRegistPrice = Null
    Else
        ReDim aryTmp(rs.RecordCount - 1, 8)
        int���� = 0: lng������Ŀid = 0
        For i = 1 To rs.RecordCount
            If lng��ĿID = Val(Nvl(rs!��ĿId)) Then
                If lng������Ŀid <> Val(Nvl(rs!������ĿID)) Then
                    int���� = 1: int���� = i:
                     lng������Ŀid = Val(Nvl(rs!������ĿID))
                End If
            Else
                int���� = 2
            End If
            
            aryTmp(i - 1, 0) = zlCommFun.Nvl(rs("�ּ�").Value, 0)
            aryTmp(i - 1, 1) = zlCommFun.Nvl(rs("������ĿID").Value, 0)
            aryTmp(i - 1, 2) = zlCommFun.Nvl(rs("�վݷ�Ŀ").Value)
            aryTmp(i - 1, 3) = zlCommFun.Nvl(rs("���㵥λ").Value)
            aryTmp(i - 1, 4) = zlCommFun.Nvl(rs("����").Value)
            aryTmp(i - 1, 5) = zlCommFun.Nvl(rs("��ĿID").Value)
            aryTmp(i - 1, 6) = IIf(int���� = 1 And i <> int����, int����, 0)
            aryTmp(i - 1, 7) = IIf(int���� = 2 And i <> int����, int����, 0)
            rs.MoveNext
        Next
        GetFreeRegistPrice = aryTmp
    End If
  Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    GetFreeRegistPrice = Null
End Function

Private Sub InitPicture()
    Dim rsPic As New ADODB.Recordset
    On Error GoTo hErr

    gstrSQL = "select ���,����,���,�߶�,���� from ��ѯͼƬԪ�� where ����=7 order by �޸����� desc,��� desc "
    Set rsPic = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsPic.BOF = False Then
      On Error Resume Next
       Me.imgBackgroundImg.Picture = ReadPicByFieldNew(rsPic!���)
    End If

    If rsPic.State <> adStateClosed Then rsPic.Close
    Set rsPic = Nothing
    
Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub
 
 

Private Sub InitFreeRegistMsg()
    Dim strFontName As String, strMsg As String, dblColor As Double, dblSize As Double
    Dim blnBold As Boolean, blnItalic As Boolean
    '��ʾ��Ϣ
    If GetRegistParaFont("�򵥹Һ���ʾ��Ϣ", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
        With Me.lblCardID
            .Caption = strMsg
            .Font.Name = strFontName
            .Font.Size = IIf(dblSize > 0, dblSize, .Font.Size)
            .ForeColor = dblColor
            .FontBold = blnBold
            .FontItalic = blnItalic
        End With
        With Me.lblCard
            .Caption = strMsg
            .Font.Name = strFontName
            .Font.Size = IIf(dblSize > 0, dblSize, .Font.Size)
            .ForeColor = dblColor
            .FontBold = blnBold
            .FontItalic = blnItalic
        End With
        
    Else
        lblCardID.Caption = "��ˢ�ڶ������֤ȡ��"
    End If
    ctlOK.Font.Name = "����"
    ctlOK.Font.Size = 40
    ctlOK.Font.Bold = True
    ctlBack.Font.Name = "����"
    ctlBack.Font.Bold = True
    ctlBack.Font.Size = 40
    If GetRegistParaFont("�򵥹Һ��ϱ���", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
        With Me.lblHospital
            .Caption = strMsg
            .Font.Name = strFontName
            .Font.Size = IIf(dblSize > 0, dblSize, .Font.Size)
            .ForeColor = dblColor
            .FontBold = blnBold
            .FontItalic = blnItalic
        End With
    Else
        lblHospital.Caption = GetUnitName & "-���׹Һ�"
    End If
    If GetRegistParaFont("�򵥹Һ��±���", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
      With Me.Lblinfo
            .Caption = strMsg
            .Font.Name = strFontName
            .Font.Size = IIf(dblSize > 0, dblSize, .Font.Size)
            .ForeColor = dblColor
            .FontBold = blnBold
            .FontItalic = blnItalic
      End With
    End If
      mdblUpBgColor = CDbl(Me.BackColor): mdblDownBgColor = CDbl(&HFFC0C0)
      Call GetFreeRegistBGColor(mdblUpBgColor, mdblDownBgColor)
    
End Sub
 
 
