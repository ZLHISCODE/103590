VERSION 5.00
Begin VB.Form frmParaModiSet 
   Caption         =   "��������ֵ��������������"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4845
   Icon            =   "frmParaModiSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4845
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraSplit 
      BackColor       =   &H80000012&
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   2400
      Width           =   6700
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4845
      TabIndex        =   10
      Top             =   2475
      Width           =   4845
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   3555
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   4995
      TabIndex        =   12
      Top             =   0
      Width           =   5000
      Begin VB.TextBox txtOld 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1365
         MaxLength       =   4000
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Frame fraSplit 
         BackColor       =   &H80000012&
         Height          =   30
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   6700
      End
      Begin VB.CommandButton cmdPC 
         Caption         =   "��"
         Height          =   300
         Left            =   4065
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1755
         Width           =   270
      End
      Begin VB.CommandButton cmdUser 
         Caption         =   "��"
         Height          =   300
         Left            =   4065
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1440
         Width           =   270
      End
      Begin VB.TextBox txtPC 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1365
         Locked          =   -1  'True
         MaxLength       =   4000
         TabIndex        =   4
         Top             =   1740
         Width           =   2730
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1365
         Locked          =   -1  'True
         MaxLength       =   4000
         TabIndex        =   1
         Top             =   1440
         Width           =   2730
      End
      Begin VB.TextBox txtValue 
         Height          =   300
         Left            =   1365
         MaxLength       =   4000
         TabIndex        =   7
         Top             =   2040
         Width           =   2970
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         Caption         =   "ԭ����ֵ "
         Height          =   180
         Left            =   480
         TabIndex        =   17
         Top             =   1020
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblChangeInfo 
         AutoSize        =   -1  'True
         Caption         =   "��ǰѡ��������������Ϣ"
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "��ȷ�ϲ���ֵ��Χ�����ֵ��ɹ���Ȼ��ʹ�øù��ܡ�"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   4500
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Left            =   480
         TabIndex        =   0
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "����ֵ(&V)"
         Height          =   180
         Left            =   480
         TabIndex        =   6
         Top             =   2100
         Width           =   810
      End
      Begin VB.Label lblPC 
         AutoSize        =   -1  'True
         Caption         =   "������(&M)"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmParaModiSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintFunType As Integer '0-�޸Ĳ���ֵ��1-������������
'ͨ����ر���
Private mstrValue As String
Private mblnOk As Boolean
'��������������ر���
Private mint˽�� As Integer '�Ƿ���˽�в���
Private mint���� As Integer '�Ƿ��Ǳ�������
Private mstrUsers As String '�û�����ɵ��ַ���
Private mstrPCs As String '��������ɵ��ַ���
Private mstrSysOwner As String '��ǰϵͳ������
Private mstrNote As String '����ֵ�޸�ʱ����ʾ
Private mlngParaID As Long '����ID

Public Function ShowMe(ByVal frmParent As Object, ByVal intFunType As Integer, ByVal strParInfo As String, ByVal strNote As String, ByVal strSysOwner As String, ByVal lngParaID As Long, ByRef strValue As String, Optional ByRef strUsers As String, Optional ByRef strPCs As String) As Boolean
'���ܣ��ô�������
'������frmParent=������
'          intFunType=���ܣ�0-�޸Ĳ���ֵ��1-������������
'          strParInfo=��ʽ���Ƿ񱾻�,�Ƿ�˽�С�����˽�в���������1,1��ʶ��������������1,0��
'          strSysOwner=ϵͳ������
'          strNote=����ֵ�޸�ʱ����ʾ
'����=True:ȷ�ϲ�����False-ȡ������
'          strValue=�µĲ���ֵ����������
'          strUsers=�û�����ɵ��ַ������ö��ŷָ��������˽�����Ͳ�������ʱ����
'          strPCs=��������ɵ��ַ������ö��ŷָ���������������Ͳ�������ʱ����
    Dim arrTmp As Variant
    
    arrTmp = Split(strParInfo & ",", ",")
    mint���� = Val(arrTmp(0))
    mint˽�� = Val(arrTmp(1))
    mintFunType = intFunType
    mstrSysOwner = strSysOwner
    mlngParaID = lngParaID
    mstrNote = strNote
    mstrValue = strValue
    If strSysOwner = "" And intFunType = 1 Then
        MsgBox "��ǰ������û�а�װ������Ա���ݣ����������������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    mstrPCs = ""
    mstrUsers = ""
    mblnOk = False
    
    Me.Show vbModal
    
    ShowMe = mblnOk
    strUsers = mstrUsers
    strPCs = mstrPCs
    strValue = mstrValue
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    mstrValue = ""
    mstrPCs = ""
    mstrUsers = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strParas As String, strMsg As String
    
    If InStr(txtUser.Text, "^") > 0 Then
        MsgBox "�û������зǷ��ַ�""^""�����飡", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    ElseIf InStr(txtUser.Text, "#") > 0 Then
        MsgBox "�û������зǷ��ַ�""#""�����飡", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    ElseIf InStr(txtUser.Text, "'") > 0 Then
        MsgBox "�û������зǷ��ַ�""'""�����飡", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    End If
    
    If InStr(txtPC.Text, "^") > 0 Then
        MsgBox "���������зǷ��ַ�""^""�����飡", vbInformation, gstrSysName
        txtPC.SetFocus
        Exit Sub
    ElseIf InStr(txtPC.Text, "#") > 0 Then
        MsgBox "���������зǷ��ַ�""#""�����飡", vbInformation, gstrSysName
        txtPC.SetFocus
        Exit Sub
    ElseIf InStr(txtPC.Text, "'") > 0 Then
        MsgBox "���������зǷ��ַ�""'""�����飡", vbInformation, gstrSysName
        txtPC.SetFocus
        Exit Sub
    End If
    
    If InStr(txtValue.Text, "^") > 0 Then
        MsgBox "����ֵ���зǷ��ַ�""^""�����飡", vbInformation, gstrSysName
        txtValue.SetFocus
        Exit Sub
    ElseIf InStr(txtValue.Text, "#") > 0 Then
        MsgBox "����ֵ���зǷ��ַ�""#""�����飡", vbInformation, gstrSysName
        txtValue.SetFocus
        Exit Sub
    ElseIf InStr(txtValue.Text, "'") > 0 Then
        MsgBox "����ֵ���зǷ��ַ�""'""�����飡", vbInformation, gstrSysName
        txtValue.SetFocus
        Exit Sub
    End If

    If txtValue.Text = "" Then
        If MsgBox("����ֵΪ�գ��Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            txtValue.SetFocus
            Exit Sub
        End If
    End If
    If txtUser.Visible And txtUser.Text = "" Then
        MsgBox "�������û�����", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    End If
    If txtPC.Visible And txtPC.Text = "" Then
        MsgBox "�������������", vbInformation, gstrSysName
        txtPC.SetFocus
        Exit Sub
    End If
    '����Ѿ����ڵĲ�������
    strSQL = "Select ����id, c.�û���, c.������" & vbNewLine & _
                "From (Select a.�û���, b.������" & vbNewLine & _
                "       From (Select Distinct Column_Value �û��� From Table(f_Str2list(Nvl([2], ',')))) a," & vbNewLine & _
                "            (Select Distinct Column_Value ������ From Table(f_Str2list(Nvl([3], ',')))) b) c, Zluserparas d" & vbNewLine & _
                "Where d.����id = [1] And Nvl(d.�û���, '�տ�') = Nvl(c.�û���, '�տ�') And Nvl(d.������, '�տ�') = Nvl(c.������, '�տ�')"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mlngParaID, txtUser.Tag, txtPC.Text)
    If rsTmp.RecordCount <> 0 Then
        If rsTmp.RecordCount = 1 Then
            If mint���� = 1 And mint˽�� = 1 Then
                 strMsg = rsTmp!�û��� & "��" & rsTmp!������ & "�ϵĲ��������Ѿ�����"
            Else
                strMsg = IIf(mint˽�� = 1, rsTmp!�û��� & "", rsTmp!������ & "") & "�Ĳ��������Ѿ�����"
            End If
        Else
            strMsg = "����" & rsTmp.RecordCount & "�����������Ѿ�����"
        End If
        '����ѯ���Ƿ񸲸ǣ������ǣ���ɾ��ԭ�в�������
        If MsgBox(strMsg & "���Ƿ񸲸�ԭ�����ã�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Do While Not rsTmp.EOF
                strTmp = rsTmp!�û��� & "^" & rsTmp!������
                If ActualLen(strParas & "#" & strTmp) >= 2000 Then
                    Call ExecuteProcedure("Zlparameters_Del_Details(" & mlngParaID & ",'" & strParas & "')", "ɾ����������")
                    strParas = strTmp
                Else
                    strParas = IIf(strParas = "", strTmp, strParas & "#" & strTmp)
                End If
                rsTmp.MoveNext
            Loop
            If strParas <> "" Then
                Call ExecuteProcedure("Zlparameters_Del_Details(" & mlngParaID & ",'" & strParas & "')", "ɾ����������")
            End If
        End If
    End If
    
    mblnOk = True
    mstrValue = txtValue.Text
    mstrPCs = txtPC.Text
    mstrUsers = txtUser.Tag
    Unload Me
End Sub

Private Sub cmdPC_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strTmp As String
    Dim i As Long

    strSQL = "Select a.Id, a.�ϼ�id, a.����, a.����, 0 ĩ��" & vbNewLine & _
                    "From " & mstrSysOwner & ".���ű� a" & vbNewLine & _
                    "Where a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000/1/1', 'yyyy-mm-dd')" & vbNewLine & _
                    "Start With �ϼ�id Is Null" & vbNewLine & _
                    "Connect By Prior Id = �ϼ�id" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select RowNum ID, a.Id, a.����, b.����վ ����, 1 ĩ��" & vbNewLine & _
                    "From " & mstrSysOwner & ".���ű� a, Zlclients b" & vbNewLine & _
                    "Where a.���� = b.���� And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000/1/1', 'yyyy-mm-dd'))"

    Set rsTmp = gclsBase.ShowSQLSelectEx(gcnOracle, Me, txtPC, strSQL, 2, "����վѡ��", False, "", "", True, True, False, blnCancel, True, True, True, "NotShowNon=1")
    If Not blnCancel And Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            strTmp = strTmp & "," & rsTmp!����
            rsTmp.MoveNext
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        txtPC.Text = strTmp
    ElseIf Not blnCancel Then
        txtPC.Text = ""
    End If
End Sub

Private Sub cmdUser_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strTmp As String, strTmp1 As String
    Dim i As Long
    
    strSQL = "Select a.Id, a.�ϼ�id, a.����, a.���� ����, ' ' �û���, 0 ĩ��" & vbNewLine & _
                    "From " & mstrSysOwner & ".���ű� a" & vbNewLine & _
                    "Where a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000/1/1', 'yyyy-mm-dd')" & vbNewLine & _
                    "Start With �ϼ�id Is Null" & vbNewLine & _
                    "Connect By Prior Id = �ϼ�id" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select c.Id, a.Id, c.���, c.����, d.�û���, 1 ĩ��" & vbNewLine & _
                    "From " & mstrSysOwner & ".���ű� a, " & mstrSysOwner & ".������Ա b, " & mstrSysOwner & ".��Ա�� c, " & mstrSysOwner & ".�ϻ���Ա�� d" & vbNewLine & _
                    "Where a.Id = b.����id And b.��Աid = c.Id And c.Id = d.��Աid And" & vbNewLine & _
                    "      (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000/1/1', 'yyyy-mm-dd')) And b.ȱʡ=1"

    Set rsTmp = gclsBase.ShowSQLSelectEx(gcnOracle, Me, txtUser, strSQL, 2, "��Աѡ����", False, "", "", False, True, False, blnCancel, True, True, True, "NotShowNon=1")
    If Not blnCancel And Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            If InStr("," & strTmp & ",", "," & rsTmp!�û��� & ",") = 0 Then
                strTmp = strTmp & "," & rsTmp!�û���
            End If
            If InStr("," & strTmp1 & ",", "," & rsTmp!���� & ",") = 0 Then
                strTmp1 = strTmp1 & "," & rsTmp!����
            End If
            rsTmp.MoveNext
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        If strTmp1 <> "" Then strTmp1 = Mid(strTmp1, 2)
        txtUser.Text = strTmp1
        txtUser.Tag = strTmp
    ElseIf Not blnCancel Then
        txtUser.Text = ""
        txtUser.Tag = ""
    End If
End Sub

Private Sub Form_Activate()
    If mintFunType = 0 Then
        txtValue.SetFocus
    Else
        If mint˽�� = 0 Then
            txtPC.SetFocus
        Else
            txtUser.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("^") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("#") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    lblPC.Visible = mintFunType = 1 And mint���� = 1
    txtPC.Visible = lblPC.Visible
    cmdPC.Visible = lblPC.Visible
    lblUser.Visible = mintFunType = 1 And mint˽�� = 1
    txtUser.Visible = lblUser.Visible
    cmdUser.Visible = lblUser.Visible
    If mintFunType = 0 Then
        lblChangeInfo.Caption = mstrNote
        lblChangeInfo.Visible = True
        lblOld.Visible = True: txtOld.Visible = True
        txtOld.Text = mstrValue
        txtValue.Text = mstrValue
        Call SetCtrlSameDistance(True, 0, 2, fraSplit(1), lblChangeInfo, lblOld, lblValue, fraSplit(0))
        Call SetCtrlPosOnLine(False, 0, lblOld, 60, txtOld)
        Call SetCtrlPosOnLine(False, 0, lblValue, 60, txtValue)
        Me.Caption = "��������ֵ"
    Else
        Me.Caption = "������������"
        If mint˽�� = 0 Then
            Call SetCtrlSameDistance(True, 0, 2, fraSplit(1), lblPC, lblValue, fraSplit(0))
            Call SetCtrlPosOnLine(False, 0, lblPC, 60, txtPC, 0, cmdPC)
            Call SetCtrlPosOnLine(False, 0, lblValue, 60, txtValue)
        ElseIf mint���� = 0 Then
            Call SetCtrlSameDistance(True, 0, 2, fraSplit(1), lblUser, lblValue, fraSplit(0))
            Call SetCtrlPosOnLine(False, 0, lblUser, 60, txtUser, 0, cmdUser)
            Call SetCtrlPosOnLine(False, 0, lblValue, 60, txtValue)
        Else
            Call SetCtrlSameDistance(True, 0, 2, fraSplit(1), lblUser, lblPC, lblValue, fraSplit(0))
            Call SetCtrlPosOnLine(False, 0, lblUser, 60, txtUser, 0, cmdUser)
            Call SetCtrlPosOnLine(False, 0, lblPC, 60, txtPC, 0, cmdPC)
            Call SetCtrlPosOnLine(False, 0, lblValue, 60, txtValue)
        End If
    End If
End Sub

Private Sub Form_Resize()
    Me.Height = 3660
    Me.Width = 5085
End Sub

Private Sub txtPC_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyDelete Then
        txtPC.Text = ""
        txtPC.Tag = ""
     End If
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyDelete Then
        txtUser.Text = ""
        txtUser.Tag = ""
     End If
End Sub

Private Sub txtValue_GotFocus()
    Call SelAll(txtValue)
End Sub
