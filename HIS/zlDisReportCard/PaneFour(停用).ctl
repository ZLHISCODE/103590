VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl PaneFour 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   ScaleHeight     =   2400
   ScaleWidth      =   9825
   Begin MSComCtl2.MonthView MView 
      Height          =   2220
      Left            =   7920
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12648447
      Appearance      =   1
      StartOfWeek     =   187105282
      CurrentDate     =   42010
   End
   Begin VB.TextBox txtEnter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   7500
      TabIndex        =   20
      Tag             =   "577,1004"
      ToolTipText     =   "�ʱ�������ʱ�ɳ����Զ�����"
      Top             =   1650
      Width           =   450
   End
   Begin VB.TextBox txtEnter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   6210
      TabIndex        =   16
      Tag             =   "508,1004"
      ToolTipText     =   "�ʱ�������ʱ�ɳ����Զ�����"
      Top             =   1650
      Width           =   1095
   End
   Begin VB.TextBox txtEnter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   8130
      TabIndex        =   15
      Tag             =   "620,1004"
      ToolTipText     =   "�ʱ�������ʱ�ɳ����Զ�����"
      Top             =   1650
      Width           =   525
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   14
      Tag             =   "113,1027"
      Top             =   2025
      Width           =   9060
   End
   Begin VB.TextBox txtDoctor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1050
      TabIndex        =   10
      Tag             =   "143,1004"
      ToolTipText     =   "�ҽ�������ʱ�ɳ����Զ�����"
      Top             =   1650
      Width           =   3255
   End
   Begin VB.TextBox txtDocNumber 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6210
      TabIndex        =   8
      Tag             =   "479,979"
      Top             =   1290
      Width           =   2520
   End
   Begin VB.TextBox txtUnit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1050
      TabIndex        =   6
      Tag             =   "137,979"
      Top             =   1290
      Width           =   3255
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6210
      TabIndex        =   4
      Tag             =   "479,957"
      Top             =   945
      Width           =   2520
   End
   Begin VB.TextBox txtIName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1050
      TabIndex        =   2
      Tag             =   "137,957"
      Top             =   945
      Width           =   3255
   End
   Begin VB.TextBox txtImportant 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   1
      Tag             =   "79,773"
      Top             =   270
      Width           =   9500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8865
      Picture         =   "PaneFour.ctx":0000
      Top             =   795
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line5 
      Tag             =   "611,1014,644"
      X1              =   8130
      X2              =   8565
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line Line4 
      Tag             =   "569,1014,599"
      X1              =   7500
      X2              =   7935
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line Line1 
      Index           =   5
      Tag             =   "485,1014,555"
      X1              =   6240
      X2              =   7305
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��"
      Height          =   180
      Index           =   0
      Left            =   7335
      TabIndex        =   19
      Tag             =   "558,1004"
      Top             =   1650
      Width           =   180
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��"
      Height          =   180
      Index           =   1
      Left            =   7965
      TabIndex        =   18
      Tag             =   "600,1004"
      Top             =   1650
      Width           =   180
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��"
      Height          =   180
      Index           =   2
      Left            =   8640
      TabIndex        =   17
      Tag             =   "644,1004"
      Top             =   1650
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��ע��"
      Height          =   180
      Index           =   11
      Left            =   105
      TabIndex        =   13
      Tag             =   "78,1027"
      Top             =   2010
      Width           =   540
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�����*��"
      Height          =   180
      Index           =   10
      Left            =   5265
      TabIndex        =   12
      Tag             =   "420,1004"
      ToolTipText     =   "�ʱ�������ʱ�ɳ����Զ�����"
      Top             =   1650
      Width           =   990
   End
   Begin VB.Line Line1 
      Index           =   4
      Tag             =   "143,1014,244"
      X1              =   1035
      X2              =   4350
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�ҽ��*��"
      Height          =   180
      Index           =   9
      Left            =   105
      TabIndex        =   11
      Tag             =   "78,1004"
      ToolTipText     =   "�ҽ�������ʱ�ɳ����Զ�����"
      Top             =   1650
      Width           =   990
   End
   Begin VB.Line Line1 
      Index           =   3
      Tag             =   "479,990,652"
      X1              =   6150
      X2              =   8785
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��ϵ�绰��"
      Height          =   180
      Index           =   8
      Left            =   5265
      TabIndex        =   9
      Tag             =   "420,979"
      Top             =   1290
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   2
      Tag             =   "137,990,358"
      X1              =   1020
      X2              =   4335
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "���浥λ��"
      Height          =   180
      Index           =   7
      Left            =   105
      TabIndex        =   7
      Tag             =   "79,979"
      Top             =   1290
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   1
      Tag             =   "479,967,652"
      X1              =   6150
      X2              =   8755
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�˿�ԭ��"
      Height          =   180
      Index           =   6
      Left            =   5265
      TabIndex        =   5
      Tag             =   "420,957"
      Top             =   945
      Width           =   900
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "����������"
      Height          =   180
      Index           =   13
      Left            =   105
      TabIndex        =   3
      Tag             =   "78,957"
      Top             =   945
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   9
      Tag             =   "137,967,358"
      X1              =   1020
      X2              =   4335
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�������������Լ��ص��⴫Ⱦ����"
      Height          =   180
      Index           =   3
      Left            =   105
      TabIndex        =   0
      Tag             =   "79,750"
      Top             =   30
      Width           =   2880
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   9825
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   9825
      Y1              =   675
      Y2              =   675
   End
End
Attribute VB_Name = "PaneFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolLoadData As Collection  '����ؼ���ʾ��Ϣ
Private mstrVL2014 As String

Public Function HaveChanged() As Boolean
'�жϿؼ���ʾ��Ϣ�Ƿ����仯
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    HaveChanged = False
    If mcolLoadData Is Nothing Then
        Set mcolLoadData = New Collection
    End If
    If mcolLoadData.Count <= 0 Then
        Exit Function
    End If
    For Each objCtl In UserControl.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                If objCtl.Text <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
            Case "uCheckNorm"
                If IIf(objCtl.Checked = True, 1, 0) <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
        End Select
        i = i + 1
    Next
End Function

Private Sub SaveLoadData()
'���ܣ�����ؼ���ʾ��Ϣ
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    Set mcolLoadData = New Collection
    For Each objCtl In UserControl.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                Call mcolLoadData.Add(objCtl.Text, "K" & i)
            Case "uCheckNorm"
                Call mcolLoadData.Add(IIf(objCtl.Checked = True, 1, 0), "K" & i)
        End Select
        i = i + 1
    Next
End Sub

Public Sub ClearMe()
    Dim objCtl As Control
    
    On Error GoTo errHand
    For Each objCtl In UserControl.Controls
        Call ClearInfo(objCtl)
    Next
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub PrintFour()
    Dim objCtl As Control
    For Each objCtl In UserControl.Controls
        Call PrintInfo(objCtl)
    Next
End Sub

Public Sub LoadData(colData As Collection, bytType As Byte, ByVal strChkType As String)
    Dim strTmp As String
    Dim i As Integer
    Dim strInfo() As String
    Dim objCtl As Control
    
    On Error GoTo errHand
    If bytType = 1 Then
        '�������������Լ��ص��⴫Ⱦ��
        txtImportant.Text = CStr(colData("K31"))
        If glngVersion = VL_2014 Then
            '��������
            txtIName.Text = CStr(colData("K38"))
            '�˿�ԭ��
            txtReason.Text = CStr(colData("K39"))
            '���浥λ
            txtUnit.Text = CStr(colData("K40"))
            '��ϵ�绰
            txtDocNumber.Text = CStr(colData("K41"))
            '�ҽ��
            txtDoctor.Text = CStr(colData("K42"))
            '��ע
            txtRemarks.Text = CStr(colData("K44"))
            '�����
            strTmp = CStr(colData("K43"))
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtEnter(i) = strInfo(i)
            Next
            mstrVL2014 = CStr(colData("K32")) & "$" & CStr(colData("K33")) & "$" & CStr(colData("K34")) & "$" & CStr(colData("K35")) & "$" & CStr(colData("K36")) & "$" & CStr(colData("K37")) & "$"
        ElseIf glngVersion = VL_2016 Then
            '��������
            txtIName.Text = CStr(colData("K32"))
            '�˿�ԭ��
            txtReason.Text = CStr(colData("K33"))
            '���浥λ
            txtUnit.Text = CStr(colData("K34"))
            '��ϵ�绰
            txtDocNumber.Text = CStr(colData("K35"))
            '�ҽ��
            txtDoctor.Text = CStr(colData("K36"))
            '��ע
            txtRemarks.Text = CStr(colData("K38"))
            '�����
            strTmp = CStr(colData("K37"))
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtEnter(i) = strInfo(i)
            Next
        End If
    Else
        txtUnit.Text = CStr(colData("K11"))
    End If
    Call SaveLoadData
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Function MakeSaveSql(arrSql() As Variant, colCls As Collection, strFileId As String) As Boolean
    Dim strObjNo_2014 As String
    Dim strObjNo_2016 As String

    Dim strContent As String
    Dim strReportInfo As String
    Dim i As Integer

    Dim strTmp As String
    Dim strTmp1 As String

    On Error GoTo errHand
    strObjNo_2014 = "31$32$33$34$35$36$37$38$39$40$41$42$43$44"
    strObjNo_2016 = "31$32$33$34$35$36$37$38"
    
    '������Ⱦ��
    strContent = Trim(txtImportant.Text) & "$"
    If glngVersion = VL_2014 Then
        strContent = strContent & mstrVL2014
    End If
    '�����������˿�ԭ�򡢱��浥λ����ϵ�绰���ҽ��
    strContent = strContent & txtIName.Text & "$" & txtReason.Text & "$" & txtUnit.Text & "$" & txtDocNumber.Text & "$" & txtDoctor.Text & "$"

    '�����
    strTmp = txtEnter(0).Text & "-" & txtEnter(1).Text & "-" & txtEnter(2).Text
    If Trim(strTmp) = "--" Then
        strTmp = ""
    End If
    strContent = strContent & strTmp & "$"

    '��ע
    strContent = strContent & txtRemarks.Text & "$"
    
    If glngVersion = VL_2014 Then
        strReportInfo = strObjNo_2014 & "|" & strContent
    Else
        strReportInfo = strObjNo_2016 & "|" & strContent
    End If
    
    MakeSaveSql = GetSaveSql(arrSql, colCls, strFileId, strReportInfo)
    Call SaveLoadData
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Sub SetEnterInfo(ByVal strDoctor As String, ByVal strDate As String)
    Dim strDateInfo() As String
    Dim strCurTime() As String
    strDateInfo = Split(Format(strDate, "yyyy-mm-dd"), "-")
    strCurTime = Split(Format(zlDatabase.Currentdate, "yyyy-mm-dd"), "-")
    txtDoctor.Text = strDoctor
    If UBound(strDateInfo) < 2 Then
        txtEnter(0).Text = strCurTime(0)
        txtEnter(1).Text = strCurTime(1)
        txtEnter(2).Text = strCurTime(2)
    Else
        txtEnter(0).Text = strDateInfo(0)
        txtEnter(1).Text = strDateInfo(1)
        txtEnter(2).Text = strDateInfo(2)
    End If
End Sub

Public Sub ClearEnterInfo()
    txtDoctor.Text = ""
    txtEnter(0).Text = ""
    txtEnter(1).Text = ""
    txtEnter(2).Text = ""
End Sub

Private Sub lblAttack_Click(Index As Integer)
    MView.Top = txtEnter(0).Top - MView.Height - 10
    MView.Left = txtEnter(0).Left
    MView.Visible = True
    Call MView.SetFocus
End Sub

Private Sub lblAttack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lblAttack(Index).MouseIcon = Image1.Picture
    lblAttack(Index).MousePointer = vbCustom
End Sub

Private Sub MView_DateClick(ByVal DateClicked As Date)
    txtEnter(0).Text = MView.Year
    txtEnter(1).Text = MView.Month
    txtEnter(2).Text = MView.Day
    MView.Visible = False
End Sub

Private Sub MView_LostFocus()
    MView.Visible = False
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackColor = vbWindowBackground
End Sub
