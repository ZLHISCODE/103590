VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatholSpecimen_AcceptOrReject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ձ걾"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   Icon            =   "frmPatholSpecimen_AcceptOrReject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame framStudyInf 
      Height          =   3480
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   6135
      Begin VB.ComboBox cbxReject 
         Height          =   300
         ItemData        =   "frmPatholSpecimen_AcceptOrReject.frx":179A
         Left            =   1080
         List            =   "frmPatholSpecimen_AcceptOrReject.frx":17B6
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2160
         Width           =   4815
      End
      Begin VB.ComboBox cbxSubmitDoctor 
         Height          =   300
         Left            =   4200
         TabIndex        =   16
         Top             =   240
         Width           =   1665
      End
      Begin VB.TextBox txtRejectNotify 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   15
         Top             =   1680
         Width           =   1785
      End
      Begin VB.TextBox txtRegisterDoctor 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   4200
         TabIndex        =   14
         Top             =   1680
         Width           =   1785
      End
      Begin VB.TextBox txtContactWay 
         Height          =   300
         Left            =   1080
         TabIndex        =   13
         Top             =   1200
         Width           =   1785
      End
      Begin VB.TextBox txtFormDepart 
         Height          =   300
         Left            =   4200
         TabIndex        =   12
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtUnitName 
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Text            =   "��Ժ"
         Top             =   720
         Width           =   1785
      End
      Begin VB.TextBox txtRejectReason 
         Height          =   780
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2520
         Width           =   4770
      End
      Begin VB.ComboBox cbxStudyType 
         ForeColor       =   &H00FF0000&
         Height          =   300
         ItemData        =   "frmPatholSpecimen_AcceptOrReject.frx":18E0
         Left            =   1080
         List            =   "frmPatholSpecimen_AcceptOrReject.frx":18E2
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtpSubmitTime 
         Height          =   300
         Left            =   4200
         TabIndex        =   17
         Top             =   1200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   121700355
         CurrentDate     =   40646.4399652778
      End
      Begin VB.Label labRejectNotify 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͨ ֪ �ˣ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   1740
         Width           =   900
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �ˣ�"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   28
         Top             =   1740
         Width           =   900
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ����ڣ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   27
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ��ʽ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label labSubmitDoctor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �ˣ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   25
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ���ң�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   24
         Top             =   780
         Width           =   900
      End
      Begin VB.Label labUnitName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ쵥λ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   780
         Width           =   900
      End
      Begin VB.Label labRejectReason 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������ɣ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label labStudyType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ű����ƣ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   20
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   19
         Top             =   765
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   18
         Top             =   2520
         Width           =   255
      End
   End
   Begin VB.TextBox txtPatholNum 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   4890
   End
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3495
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdReject_Cancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   5040
      TabIndex        =   2
      Top             =   5115
      Width           =   1215
   End
   Begin VB.CommandButton cmdReject_Sure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   3720
      TabIndex        =   1
      Top             =   5115
      Width           =   1215
   End
   Begin VB.Label labPatholNumNeed 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label labPatholNum 
      Caption         =   "����ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmPatholSpecimen_AcceptOrReject.frx":18E4
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6360
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "    ����ϸ�˶��ͼ�걾������ȷ¼��걾��/���յ���ϸ��Ϣ�����걾�����պ󣬽����ܶ����޸Ļ�ɾ����"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   195
      Width           =   5175
   End
End
Attribute VB_Name = "frmPatholSpecimen_AcceptOrReject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsRejectSpecimen As Boolean
Private mblnIsSucceed As Boolean
Private mlngCurAdviceId As Long
Private mlngPatholStationMoneyExeModle As Long
'Public mlngCurStudyProcedure As Long
Public mstrCurDepartmentId As String


Public mtxtAcceptHistory As RichTextBox

Public mstrPatholNum As String
Private mlngStudyType As Long

Public mlngPatholSerialNum As Long
Public mstrPatholInitNum As String

Public mobjSquareCard As Object    'һ��ͨ�������㲿��

Public mfrmParent As Form

Public mstrPrivs As String          '�����ߵ�Ȩ��

Property Get IsRejectSpecimen() As Boolean
    IsRejectSpecimen = mblnIsRejectSpecimen
End Property

Property Let IsRejectSpecimen(value As Boolean)
    mblnIsRejectSpecimen = value
End Property


Property Get AdviceId() As Long
    AdviceId = mlngCurAdviceId
End Property

Property Let AdviceId(value As Long)
    mlngCurAdviceId = value
End Property




Property Get IsSucceed() As Boolean
    IsSucceed = mblnIsSucceed
End Property


Public Function ShowAcceptOrRejectSpecimenWindow(lngAdviceID As Long, _
    ByVal lngCurDepartmentId As String, txtAcceptHis As RichTextBox, blnIsReject As Boolean, owner As Form, _
    strPrivs As String, ByVal blnShowSentInfo As Boolean) As Boolean
    
    Dim frmAOR As New frmPatholSpecimen_AcceptOrReject
    
    On Error GoTo errFree
    
    With frmAOR
        .AdviceId = lngAdviceID
        
'        .mlngCurStudyProcedure = lngCurStudyProcedure
        .mstrCurDepartmentId = lngCurDepartmentId
        .mlngPatholSerialNum = 0
        .mstrPatholInitNum = ""
        .mstrPrivs = strPrivs
        
        Set .mtxtAcceptHistory = txtAcceptHis
        Set .mfrmParent = owner
        
        .IsRejectSpecimen = blnIsReject
        
        .txtRejectReason.Text = ""
        .dtpSubmitTime.value = zlDatabase.Currentdate
        .txtRegisterDoctor.Text = UserInfo.����
        
        If blnIsReject Then
            frmAOR.Caption = "���ձ걾"
            
            .labSubmitDoctor.Left = .labStudyType.Left
            .cbxSubmitDoctor.Left = .cbxStudyType.Left
            .cbxSubmitDoctor.Width = .txtRejectReason.Width
            
            .labStudyType.Visible = False
            .cbxStudyType.Visible = False
            
            .framStudyInf.Top = .txtPatholNum.Top
            .picShow.Top = .framStudyInf.Top + .framStudyInf.Height + 2400
            
            .cmdReject_Sure.Top = .framStudyInf.Top + .framStudyInf.Height + 120
            .cmdReject_Cancel.Top = .cmdReject_Sure.Top
            
            .Height = .cmdReject_Sure.Top + .cmdReject_Sure.Height + 120 + 430
        Else
            frmAOR.Caption = "���ձ걾"
            
            .txtRejectReason.Visible = False
            .cbxReject.Visible = False
            
            .labRejectNotify.Visible = False
            .txtRejectNotify.Visible = False
            
            .Label24.Left = .labRejectNotify.Left
            .txtRegisterDoctor.Left = .txtRejectNotify.Left
            
            .framStudyInf.Height = 2160
            .Height = 4655
            
            .cmdReject_Sure.Top = .ScaleHeight - .cmdReject_Sure.Height - 120
            .cmdReject_Cancel.Top = .cmdReject_Sure.Top
            .picShow.Top = .cmdReject_Sure.Top - 120
        End If
        
        '��ȡ�������
        Call .GetStudyAcceptInf(lngAdviceID)
        Call .ConfigStudyType
        Call .ConfigSubmitInf(lngAdviceID, blnShowSentInfo)
        
        .txtPatholNum.Visible = Not blnIsReject
        .labPatholNum.Visible = Not blnIsReject
        .labPatholNumNeed.Visible = Not blnIsReject
        
        .cbxReject.Enabled = blnIsReject
        .cbxReject.BackColor = IIf(blnIsReject, &H80000005, &H8000000F)
        
        .txtRejectReason.Enabled = blnIsReject
        .txtRejectReason.BackColor = IIf(blnIsReject, &H80000005, &H8000000F)
        
        .txtRejectNotify.Enabled = blnIsReject
        .txtRejectNotify.BackColor = IIf(blnIsReject, &H80000005, &H8000000F)
        
        .labRejectReason.Enabled = blnIsReject
        .labRejectNotify.Enabled = blnIsReject
        
        Call .CloseProcessHint
        
        If Trim(.mstrPatholNum) = "" Then
            '����Ĭ�ϼ��صļ�����͵õ���صĲ����
            .txtPatholNum.Text = GetPatholNum(Val(.cbxStudyType.Text))
        End If
    End With
    

    Call frmAOR.Show(1, owner)
    
    ShowAcceptOrRejectSpecimenWindow = frmAOR.IsSucceed
        
errFree:
    Unload frmAOR
    Set frmAOR = Nothing
End Function




Public Sub ConfigSubmitInf(lngAdviceID As Long, ByVal blnShowSentInfo As Boolean)
'�����ͼ���Ϣ
    If lngAdviceID <= 0 Then Exit Sub
    
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strSubmitDoctor As String
    Dim strUnit As String     '�����ͼ���Ϣ �ͼ쵥λ
    Dim strSentPerson As String '�����ͼ���Ϣ �ͼ���
    Dim strSentDepart As String '�����ͼ���Ϣ �ͼ����
    
    If blnShowSentInfo Then
    
        strSql = "select �ͼ쵥λ,�ͼ����,�ͼ��� from �����ͼ���Ϣ where ҽ��ID=[1] and �ͼ�����=to_date('1000/10/10 10:10:10','yyyy/mm/dd hh24:mi:ss')"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
        If rsData.RecordCount = 1 Then
            strUnit = Trim(Nvl(rsData!�ͼ쵥λ))
            strSentPerson = Trim(Nvl(rsData!�ͼ���))
            strSentDepart = Trim(Nvl(rsData!�ͼ����))
        End If
    Else
        strUnit = "��Ժ"
    End If
    
    txtUnitName = strUnit
    txtFormDepart = strSentDepart
    
    If Len(strSentDepart) = 0 Then
        strSql = "select ���� from ���ű� a, ����ҽ����¼ b where a.id =b.��������id and b.id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
        
        If rsData.RecordCount <= 0 Then Exit Sub
        txtFormDepart.Text = rsData!����
    End If
    
    '��ȡ�ͼ���Ա��Ϣ
    strSql = "select case when c.����ҽ��=a.���� then 1 else 0 end as �Ƿ��ͼ�, a.���� from ��Ա�� a, ������Ա b, ����ҽ����¼ c where a.id=b.��Աid and b.����Id=c.��������Id and c.id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    strSubmitDoctor = ""
    Call cbxSubmitDoctor.Clear
    
    If Len(strSentPerson) > 0 Then Call cbxSubmitDoctor.AddItem(strSentPerson)
    While Not rsData.EOF
        Call cbxSubmitDoctor.AddItem(Nvl(rsData!����))
        
        strSubmitDoctor = IIf(Val(Nvl(rsData!�Ƿ��ͼ�)) = 1, Nvl(rsData!����), strSubmitDoctor)
        
        rsData.MoveNext
    Wend
    
    If strSubmitDoctor <> "" Then
        cbxSubmitDoctor.Text = strSubmitDoctor
        If Len(strSentPerson) > 0 Then cbxSubmitDoctor.Text = strSentPerson
    End If
End Sub


Public Sub GetStudyAcceptInf(ByVal lngAdviceID As Long)
'��ȡ��������Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select �����,�������ID from ��������Ϣ where ҽ��ID=[1]"
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    mstrPatholNum = ""
    mlngStudyType = -1
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    mstrPatholNum = Nvl(rsData!�����)
    mlngStudyType = Val(Nvl(rsData!�������ID))
End Sub


Public Sub ConfigStudyType()
'���ü���¼������
    If mlngStudyType < 0 Then Exit Sub
    
    'cbxStudyType.ListIndex = mlngStudyType
    txtPatholNum.Text = mstrPatholNum
    
    cbxStudyType.BackColor = &H8000000F
    cbxStudyType.Enabled = False
    
    txtPatholNum.BackColor = &H8000000F
    txtPatholNum.Enabled = False
    
    labPatholNum.Enabled = False
End Sub

Private Sub LoadStudyType()
    '����������
    On Error GoTo errH
    Dim i As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset

    strSql = "select ID,���� from ����������"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ò���������")

    If rsData.RecordCount > 0 Then
        With cbxStudyType
        .Clear
            rsData.MoveFirst
            Do While Not rsData.EOF
                If Nvl(rsData!����, "  ") <> "  " Then
                    .AddItem Nvl(rsData!ID, 0) & "-" & rsData!����
                End If
                rsData.MoveNext
            Loop

        End With
    End If
    
    Call GetStudyAcceptInf(mlngCurAdviceId)
    
    If mlngStudyType >= 0 Then
        '�Ѿ�������
        strSql = "select B.ID as ִ�з��� from  ��������Ϣ A,���������� B where A.�������ID=B.ID and ҽ��id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurAdviceId)
    Else
        '��δ����
        strSql = "select ִ�з��� from ������ĿĿ¼ where ID= (select ������ĿID from ����ҽ����¼ where id=[1])"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurAdviceId)
    End If

    If rsData.RecordCount > 0 Then
        For i = 0 To cbxStudyType.ListCount - 1
            If Val(Mid(cbxStudyType.list(i), 1, InStr(cbxStudyType.list(i), "-") - 1)) = Val(Nvl(rsData!ִ�з���)) Then
                cbxStudyType.ListIndex = i
                Exit Sub
            End If
        Next
        cbxStudyType.ListIndex = 0
    Else
        cbxStudyType.ListIndex = 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function CheckSubmitInfoIsValid() As String
    '����ͼ���Ϣ�Ƿ���Ч
    CheckSubmitInfoIsValid = ""
    
    If Trim(txtFormDepart.Text) = "" Then
        CheckSubmitInfoIsValid = "�ͼ���Ҳ���Ϊ�ա�"
        
        Call txtFormDepart.SetFocus
        Exit Function
    End If
    
    If Trim(cbxSubmitDoctor.Text) = "" Then
        CheckSubmitInfoIsValid = "�ͼ��˲���Ϊ�ա�"
        
        Call cbxSubmitDoctor.SetFocus
        Exit Function
    End If
    
    If txtRejectReason.Enabled Then
        If Trim(txtRejectReason.Text) = "" Then
            CheckSubmitInfoIsValid = "����ԭ����Ϊ�ա�"
            
            Call txtRejectReason.SetFocus
            Exit Function
        End If
    Else
        If Trim(txtPatholNum.Text) = "" Then
            CheckSubmitInfoIsValid = "����Ų���Ϊ�ա�"
            
            Call txtPatholNum.SetFocus
            Exit Function
        End If
    End If
End Function



Private Sub ShowProcessHint(ByVal strHint As String)
'��ʾ������Ϣ
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Public Sub CloseProcessHint()
'�رմ�����ʾ
    picShow.Visible = False
End Sub

Private Sub cbxReject_Click()
On Error GoTo errHandle
    If Trim(txtRejectReason.Text) <> "" Then txtRejectReason.Text = txtRejectReason.Text & vbCrLf
    txtRejectReason.Text = txtRejectReason.Text & Mid(cbxReject.Text, InStr(cbxReject.Text, "-") + 1, 100)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function GetPatholNum(ByVal longID As Long) As String
'����ID��ȡ�����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetPatholNum = ""
    
    strSql = "select Zl_�������_��Ż�ȡ([1]) as ������� from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, longID)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    mlngPatholSerialNum = Val(Nvl(rsData!�������))
    
    strSql = "select Zl_�������_����([1],[2]) as ����� from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, longID, mlngPatholSerialNum)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    mstrPatholInitNum = Nvl(rsData!�����)
    
    GetPatholNum = mstrPatholInitNum
End Function


Private Sub cbxStudyType_Click()
On Error GoTo errHandle
    If Trim(mstrPatholNum) = "" Then
        txtPatholNum.Text = GetPatholNum(Val(cbxStudyType.Text))
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdReject_Cancel_Click()
    mblnIsSucceed = False
    
    Call Me.Hide
End Sub


Private Function AutoRegister() As Boolean
'�Զ�����ע��
'ȡ�����˵�ǰ���۷��ã���ִ�к��Զ���˻��۵�����Чʱ��
    Dim curMoney As Currency
    Dim str��� As String
    Dim str����� As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngSourceType As Long
    Dim lngPatientID As Long
    Dim arrSQL() As Variant
    Dim i As Integer
    Dim blnTran As Boolean
    Dim rsOneCard As ADODB.Recordset
    Dim int��¼���� As Integer     '����ҽ������.��¼���ʣ�����ҽ���ļ�¼���ʣ�1-�շѼ�¼��2-���ʼ�¼
    Dim int������� As Integer     '����ҽ������.������ʣ������סԺҽ��վ����Ϊ�������ʱ��Ϊ1,��������������ʺ�סԺ���ʣ������Ķ���Ϊ��
    Dim str������� As String
    Dim lng���ͺ� As Long
    Dim str���ݺ� As String
    Dim strҽ��IDs As String
    

On Error GoTo errHandle

    AutoRegister = True


    strSql = "select A.������Դ,A.ID,A.����,A.�Ա�,A.����,A.����ID,A.��ҳID,B.��������,B.��ǰ����ID, decode(c.ҽ��id, null, '0',1) as ����״̬, D.���ͺ� " & _
            " from ����ҽ����¼ A, ������Ϣ B, Ӱ�����¼ C, ����ҽ������ D " & _
            " where a.����id = b.����id and a.id = c.ҽ��id(+) and a.ID =D.ҽ��ID and a.id=[1]"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurAdviceId)
    
    '����Ѿ�д������Ϣ���򲻽��б�������
    If Nvl(rsTemp!����״̬, 0) = 1 Then Exit Function
    
    
    lngSourceType = Nvl(rsTemp!������Դ, 3)
    lngPatientID = rsTemp!����ID
    
    
    '�������Լ�һ��ͨ�Ĵ���
        'ҵ���߼��ǣ�
        '1�������߼�û���շѵĲ��ܱ�������������С�δ�ɷѱ�����Ȩ�޵ģ�������û���շѵ�����±�����
        '   ��ˢ����Ϣ��ʱ���Ѿ����Ʊ�����ȷ����ť��
        '2���Թ�������������֧�֣�
        '       ������28--����һ��ͨ�����Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
        '       ������81--ִ�к��Զ����
        '       ������163--����һ��ͨ����Ŀִ��ǰ�������շѻ��ȼ������
        '3���ȴ�����Ҫһ��ͨ����ȷ�ϵģ�����������֮һ
        '       ��1����¼����=1
        '       ��2��ִ�к��Զ����=False����¼����=2���� ����Դ<>סԺ��  ���� ����Դ=סԺ��������ʡ���
        '   ���һ��ͨ����ȷ�ϳɹ�������Ա��������һ��ͨ����ȷ�ϲ��ɹ��������Ȩ�ޡ�δ�ɷѱ�������ʾ�Ƿ����������
        '4���ٴ���һ��ͨ���ü�����֤�ģ�ֻ������˵ģ������ǣ�
        '       ��1����¼����=2��ִ�к��Զ����=True
        '       ��2����δ��˷���
        '
        '
        '
        gstrSQL = "Select A.��¼����,A.�������,A.���ͺ�,A.NO,B.������� from ����ҽ������ A,����ҽ����¼ B  where A.ҽ��ID=B.ID and  B.ID =[1]"
        Set rsOneCard = zlDatabase.OpenSQLRecord(gstrSQL, "PACS�������Ҽ�¼����", mlngCurAdviceId)
        If rsOneCard.EOF = False Then
            int��¼���� = Nvl(rsOneCard!��¼����, 0)
            int������� = Nvl(rsOneCard!�������, 0)
            str������� = Nvl(rsOneCard!�������)
            lng���ͺ� = rsOneCard!���ͺ�
            str���ݺ� = Nvl(rsOneCard!NO)
        End If
        
        If int��¼���� = 1 Or _
            (gblnִ�к���� = False And int��¼���� = 2 And (lngSourceType <> 2 Or (lngSourceType = 2 And int������� = 1))) Then
            
            If Not ItemHaveCash(lngSourceType, False, mlngCurAdviceId, 0, lng���ͺ�, str�������, str���ݺ�, int��¼����, _
                int�������, 0) Then
                If gblnִ��ǰ�Ƚ��� Then
                    '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������,�������ݺţ�����ҽ��ID��ȡ����δ�շѵ��ݻ�δ��˵ļ��ʵ�
                    '��ȡҽ��ID��
                    strҽ��IDs = mlngCurAdviceId
                    gstrSQL = "Select Id  from ����ҽ����¼ where ���ID = [1]"
                    Set rsOneCard = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ��ID��", mlngCurAdviceId)
                    While rsOneCard.EOF = False
                        strҽ��IDs = strҽ��IDs & "," & rsOneCard!ID
                        rsOneCard.MoveNext
                    Wend
                    
                    If mobjSquareCard.zlSquareAffirm(Me, 1294, mstrPrivs, lngPatientID, 0, False, , , strҽ��IDs) = False Then
                        '����С�δ�ɷѱ�����Ȩ�ޣ�����ʾ�Ƿ�ȷ��δ�շѿ��Ա�����
                        If CheckPopedom(mstrPrivs, "δ�ɷѱ���") Then
                            If MsgBoxD(Me, "�ɷѲ��ɹ����ò��˻�����δ�շѵķ��ã��Ƿ����������", vbYesNo, "�ɷ�ʧ��") = vbNo Then
                                AutoRegister = False
                                Exit Function
                            End If
                        Else
                            MsgBoxD Me, "�ɷѲ��ɹ����ò��˻�����δ�շѵķ��ã��޷����������顣", vbOKOnly, "�ɷ�ʧ��"
                            AutoRegister = False
                            Exit Function
                        End If
                    End If
                Else
                    '����С�δ�ɷѱ�����Ȩ�ޣ�����ʾ�Ƿ�ȷ��δ�շѿ��Ա�����
                    If CheckPopedom(mstrPrivs, "δ�ɷѱ���") Then
                        If MsgBoxD(Me, "�ò��˻�����δ�շѵķ��ã��Ƿ����������", vbYesNo, "��ʾ��Ϣ") = vbNo Then
                            AutoRegister = False
                            Exit Function
                        End If
                    Else
                        MsgBoxD Me, "�ò��˻�����δ�շѵķ��ã����顣", vbOKOnly, "��ʾ��Ϣ"
                        AutoRegister = False
                        Exit Function
                    End If
                End If
            End If
        End If
        
    
    If gblnִ�к���� And int��¼���� = 2 Then
        curMoney = GetAdviceMoney(mlngCurAdviceId, lngSourceType, str���, str�����)
        
        '�����ò�Ϊ0ʱ������Ƿ�һ��ͨˢ�����Ƿ���Ҫ���˱���
        If curMoney <> 0 Then
            '���˱���
            If Not FinishBillingWarn(Me, "", lngPatientID, rsTemp!��ҳID, Val(Nvl(rsTemp!��ǰ����ID)), curMoney, str���, str�����) Then
                AutoRegister = False
                Exit Function
            End If
    
            '���⣺34856
            '����һ��ͨ���������֤
            '����28--����һ��ͨ���Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
            '����81--ִ�к��Զ����
            If glng������֤ <> 0 And gblnִ�к���� _
                And curMoney > 0 And lngSourceType = 1 Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, lngPatientID, curMoney, , , , , , , IIf(glng������֤ = 2, True, False)) Then
                    AutoRegister = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    arrSQL = Array()

    '��ʼ���
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    
    'Ӱ�����"DG"��ʾ����
    arrSQL(UBound(arrSQL)) = "ZL_Ӱ����_BEGIN(Null,Null," & mlngCurAdviceId & "," & rsTemp!���ͺ� & ",'DG','" & _
        Nvl(rsTemp!����, "") & "','','" & Nvl(rsTemp!�Ա�, "") & "','" & _
        Nvl(rsTemp!����, "") & "'," & zlStr.To_Date(Nvl(rsTemp!��������, "")) & ",Null,Null,Null,Null,Null,Null,Null,Null,Null," & _
        mstrCurDepartmentId & ")"
    
    '����Ӱ�����¼--ִ�й���Ϊ-�ѱ���������ʱ������˵ķ���
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_Ӱ����_State(" & mlngCurAdviceId & "," & rsTemp!���ͺ� & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mstrCurDepartmentId & ")"
    
    
    'ִ�з���
    If mlngPatholStationMoneyExeModle = 0 Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_Ӱ�����ִ��(" & mlngCurAdviceId & "," & rsTemp!���ͺ� & ",2,Null,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mstrCurDepartmentId & ")"
    End If
    
    gcnOracle.BeginTrans
    
    blnTran = True
    
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "д������")
    Next
    gcnOracle.CommitTrans
    
    Exit Function
errHandle:
    AutoRegister = False
    
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function


Private Function IsNewPatholStudy() As Boolean
'�����Ƿ��µĲ�����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    IsNewPatholStudy = True
    
    strSql = "select ����� from ��������Ϣ where ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If Nvl(rsData!�����) <> "" Then IsNewPatholStudy = False
End Function


Private Function IsHasPatholNum(ByVal strCurPatholNum As String) As Boolean
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ����ҽ��ID from ��������Ϣ where upper(�����)=upper([1])"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strCurPatholNum)
    
    IsHasPatholNum = IIf(rsData.RecordCount > 0, True, False)
End Function


Private Sub cmdReject_Sure_Click()
On Error GoTo errHandle
    '��/���ձ걾
    Dim rsPathol As ADODB.Recordset
    Dim i As Integer
    Dim strSql As String
    Dim strErr As String
    Dim strPatholNum As String

    strErr = CheckSubmitInfoIsValid
    If Trim(strErr) <> "" Then
        Call ShowProcessHint(strErr)
        Exit Sub
    End If

    If mblnIsRejectSpecimen Then
        '���ձ걾
        strSql = "Zl_����걾_����(" & mlngCurAdviceId & ",'" & _
                                    txtUnitName.Text & "','" & _
                                    txtFormDepart.Text & "','" & _
                                    cbxSubmitDoctor.Text & "'," & _
                                    zlStr.To_Date(dtpSubmitTime.value) & ",'" & _
                                    txtContactWay.Text & "','" & _
                                    txtRegisterDoctor.Text & "','" & _
                                    txtRejectReason.Text & "','" & _
                                    txtRejectNotify.Text & "')"

        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Else
        '����걾��Ϊ���գ���ִ���Զ�����
        If IsNewPatholStudy Then
        
            '�жϲ�����Ƿ��ظ�
            If IsHasPatholNum(txtPatholNum.Text) Then
                Call MsgBoxD(Me, "������ظ����޸ġ�", vbInformation, Me.Caption)
                txtPatholNum.SetFocus
                
                Exit Sub
            End If
        
            If Not AutoRegister Then Exit Sub
        End If
        
        '�ȱ���걾��Ϣ
        Call mfrmParent.SaveSpecimenData(False)
    
        '���ձ걾
        strSql = "Zl_����걾_����(" & mlngCurAdviceId & ",'" & _
                                    txtPatholNum.Text & "'," & _
                                    Val(cbxStudyType.Text) & ",'" & _
                                    txtUnitName.Text & "','" & _
                                    txtFormDepart.Text & "','" & _
                                    cbxSubmitDoctor.Text & "'," & _
                                    zlStr.To_Date(dtpSubmitTime.value) & ",'" & _
                                    txtContactWay.Text & "','" & _
                                    txtRegisterDoctor.Text & "')"

        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        
        If Trim(mstrPatholNum) = "" And mstrPatholInitNum = Trim(txtPatholNum.Text) Then
            '���²������
            strSql = "ZL_�������_��Ÿ���(" & Val(cbxStudyType.Text) & "," & mlngPatholSerialNum & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
    End If

    Call mfrmParent.LoadSpecimenAcceptOrRejectHistoryData

    mblnIsSucceed = True
    
    Call Me.Hide

    Exit Sub
errHandle:
    Call ShowProcessHint(err.Description)
End Sub



Private Sub Form_Initialize()
    mblnIsSucceed = False
    mblnIsRejectSpecimen = False
    mlngCurAdviceId = 0
    
    Set mtxtAcceptHistory = Nothing
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    
    
    '���������㲿��
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    '��ʼ�������㲿��
    mobjSquareCard.zlInitComponents Me, 1294, glngSys, gstrDBUser, gcnOracle
    
    mlngPatholStationMoneyExeModle = Val(zlDatabase.GetPara("�������ִ��ģʽ", glngSys, G_LNG_PATHOLSYS_NUM, 0))
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadStudyType
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Set mobjSquareCard = Nothing
End Sub
