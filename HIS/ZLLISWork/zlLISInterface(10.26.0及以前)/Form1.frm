VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   8550
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdGetAllDev 
      Caption         =   "ȡ���������б�"
      Height          =   360
      Left            =   6675
      TabIndex        =   17
      Top             =   270
      Width           =   1650
   End
   Begin VB.CommandButton cmdUnAudit 
      Caption         =   "ȡ�����󱨸�"
      Height          =   360
      Left            =   5235
      TabIndex        =   16
      Top             =   4200
      Width           =   1605
   End
   Begin VB.CommandButton cmdWritLIS 
      Caption         =   "д�����󱨸�"
      Height          =   360
      Left            =   3285
      TabIndex        =   15
      Top             =   4230
      Width           =   1605
   End
   Begin VB.CommandButton cmdȡ������ 
      Caption         =   "ȡ������"
      Height          =   360
      Left            =   6555
      TabIndex        =   14
      Top             =   3435
      Width           =   990
   End
   Begin VB.TextBox txt����ID 
      Height          =   270
      Left            =   5100
      TabIndex        =   12
      Text            =   "41"
      Top             =   3525
      Width           =   525
   End
   Begin VB.TextBox txt�걾�� 
      Height          =   270
      Left            =   5835
      TabIndex        =   11
      Text            =   "1"
      Top             =   3510
      Width           =   645
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "���ձ걾��LIS"
      Height          =   480
      Index           =   1
      Left            =   3285
      TabIndex        =   10
      Top             =   3360
      Width           =   1620
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "��Ϊ�����˷�"
      Height          =   480
      Index           =   0
      Left            =   345
      TabIndex        =   9
      Top             =   3315
      Width           =   1620
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "3����ȡ������ϸ"
      Height          =   480
      Left            =   300
      TabIndex        =   8
      Top             =   1410
      Width           =   1620
   End
   Begin VB.TextBox txtItem 
      Height          =   315
      Left            =   2085
      TabIndex        =   7
      Top             =   1485
      Width           =   1980
   End
   Begin VB.TextBox txtClinic 
      Height          =   315
      Left            =   2085
      TabIndex        =   6
      Top             =   930
      Width           =   1980
   End
   Begin VB.CommandButton cmdGetClinic 
      Caption         =   "2����ȡ��������"
      Height          =   480
      Left            =   300
      TabIndex        =   5
      Top             =   855
      Width           =   1620
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Left            =   2085
      TabIndex        =   4
      Top             =   390
      Width           =   1980
   End
   Begin VB.CommandButton cmdGetA 
      Caption         =   "1����ȡ��������"
      Height          =   480
      Left            =   300
      TabIndex        =   3
      Top             =   300
      Width           =   1620
   End
   Begin VB.CommandButton cmdSaveTiJian 
      Caption         =   "���������"
      Height          =   480
      Left            =   225
      TabIndex        =   2
      Top             =   4125
      Width           =   1620
   End
   Begin VB.CommandButton cmdDelRtf 
      Caption         =   "ɾ��Rtf����"
      Height          =   480
      Left            =   4035
      TabIndex        =   1
      Top             =   5565
      Width           =   1620
   End
   Begin VB.CommandButton cmdInsRtf 
      Caption         =   "����Rtf����"
      Height          =   480
      Left            =   1860
      TabIndex        =   0
      Top             =   5565
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����סԺ�ŵȣ�"
      Height          =   180
      Left            =   4170
      TabIndex        =   18
      Top             =   390
      Width           =   1260
   End
   Begin VB.Label lblIDID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ID  �걾��"
      Height          =   180
      Left            =   5115
      TabIndex        =   13
      Top             =   3225
      Width           =   1260
   End
   Begin VB.Line Line10 
      X1              =   2760
      X2              =   2505
      Y1              =   2655
      Y2              =   2895
   End
   Begin VB.Line Line9 
      X1              =   2130
      X2              =   2430
      Y1              =   2625
      Y2              =   2895
   End
   Begin VB.Line Line8 
      X1              =   2760
      X2              =   2610
      Y1              =   5400
      Y2              =   5490
   End
   Begin VB.Line Line7 
      X1              =   2430
      X2              =   2580
      Y1              =   5370
      Y2              =   5490
   End
   Begin VB.Line Line6 
      X1              =   2580
      X2              =   2580
      Y1              =   4980
      Y2              =   5490
   End
   Begin VB.Line Line4 
      Index           =   3
      X1              =   3810
      X2              =   3810
      Y1              =   4590
      Y2              =   4965
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   1305
      X2              =   1305
      Y1              =   4575
      Y2              =   4950
   End
   Begin VB.Line Line5 
      X1              =   1320
      X2              =   3840
      Y1              =   4995
      Y2              =   4995
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   3795
      X2              =   3795
      Y1              =   3825
      Y2              =   4200
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   1350
      X2              =   1350
      Y1              =   3825
      Y2              =   4200
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   3840
      X2              =   3840
      Y1              =   2940
      Y2              =   3315
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   1335
      X2              =   1335
      Y1              =   2940
      Y2              =   3285
   End
   Begin VB.Line Line2 
      X1              =   1335
      X2              =   3870
      Y1              =   2925
      Y2              =   2925
   End
   Begin VB.Line Line1 
      X1              =   2460
      X2              =   2460
      Y1              =   1905
      Y2              =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lis As New ZLLISInterface.clsLisInterface
Private mblnReg As Boolean

Private Sub cmdGetAllDev_Click()
    Dim strReturn As String, strErr As String
    strReturn = Lis.gGetAllDevice(strErr)
    If strReturn = "" Then
        MsgBox strErr
    Else
        MsgBox strReturn
    End If
End Sub

Private Sub cmdGetClinic_Click()
    Dim lngID As Long, strReturn As String
    lngID = Val(txtClinic)
    If lngID > 0 Then
        strReturn = Lis.gGetClinicItem(lngID)
        MsgBox strReturn
    End If
End Sub

Private Sub cmdGetA_Click()
    '��ȡ��������
    Dim strInput As String, strReturn As String
    strInput = txtInput
    strReturn = Lis.gGetApplication(strInput)
    MsgBox strReturn
    
End Sub

Private Sub cmdInsRtf_Click()
    Dim strTmp As String
    Dim astrTmp() As String
    
    Call Lis.gInsertReport(CLng(txtClinic.Text), "c:\Temp\test.rtf", strTmp)

    MsgBox strTmp
End Sub

Private Sub cmdDelRtf_Click()
    Dim strTmp As String
    
    rsTmp = Lis.gDeleteReport(1380633)
    MsgBox rsTmp
End Sub

Private Sub cmdItem_Click()
    Dim lngID As Long, strReturn As String
    lngID = Val(txtItem)
    If lngID > 0 Then
        strReturn = Lis.gGetItemList(lngID)
        MsgBox strReturn
    End If
    
End Sub

Private Sub cmdReg_Click(Index As Integer)
    Dim lng����id As Long, lngID As Long, strInfo As String, str�걾�� As String
    lng����id = Val(txt����ID)
    lngID = Val(txtClinic)
    str�걾�� = Val(txt�걾��)
    
    If Lis.gzlLisRegister(lng����id, lngID, str�걾��, strInfo) = True Then
        MsgBox "���ճɹ���" & vbNewLine & strInfo
    Else
        MsgBox "����ʧ�ܣ�" & vbNewLine & strInfo
    End If
End Sub

Private Sub cmdSaveLIS_Click()

End Sub

Private Sub cmdSaveTiJian_Click()
       '������Ŀid;������1;��λ1;�����1��;�����־1|������Ŀid;������2;��λ2;����ο�2;�����־2
    Dim strResult As String, varTmp As Variant
    Dim strErr As String, strInfo As String
    strResult = Lis.gTestResults(8252908, "���鼼ʦ", Format(Now, "yyyy-MM-dd HH:mm:ss"), _
                "548;3.72;X10E+12/L;3.5��5.5;|" & _
                "443;10.10;10E+9/L;4��10;ƫ��|" & _
                "721;107.00;g/L;110��160;ƫ��|" & _
                "554;0.37;%;.37��.49;ƫ��|" & _
                "648;98.1;fL;80��100;|" & _
                "550;28.8;pg;27��31;|" & _
                "551;293.20;g/L;320��360;ƫ��|" & _
                "735;216.00;10E+9/L;100��450;|" & _
                "610;16.1;%;20��40;ƫ��|" & _
                "500;6.9;��;3��8;|" & _
                "792;74.9;��;50��70;ƫ��|" & _
                "673;2.0;��;.5��5;|" & _
                "670;0.1;��;0��1;|" & _
                "611;1.63;10E+9/L;.8��4;|" & _
                "501;0.70;10E+9/L;.12��.8;|" & _
                "793;7.56;10E+9/L;2��7.5;ƫ��|" & _
                "674;0.20;10E+9/L;0��.5;|" & _
                "671;0.01;10E+9/L;0��.1;|" & _
                "544;0.15;��;.11��.16;|" & _
                "545;53.1;fL;37��54;|" & _
                "738;12.3;fL;9��17;|" & _
                "649;10.1;fL;8��13;|" & _
                "497;0.26;��;.13��.43;|" & _
                "734;0.22;��;.18��.22;")
    If strResult <> "" Then
        varTmp = Split(strResult, vbNewLine)
        For i = LBound(varTmp) To UBound(varTmp)
            If varTmp(i) Like "0|*" Then
                strErr = strErr & Mid(varTmp(i), 3) & vbNewLine
            ElseIf varTmp(i) Like "1|*" Then
                strInfo = strInfo & Mid(varTmp(i), 3) & vbNewLine
            End If
        Next
        MsgBox IIf(strErr <> "", "������ʾ��" & vbNewLine & strErr, "") & IIf(strInfo <> "", "��ʾ��Ϣ��" & vbNewLine & strInfo, "")
    Else
        MsgBox "OK"
    End If
End Sub

Private Sub cmdUnAudit_Click()
    Dim strErrinfo As String
    If Lis.gzlLisUnAudit(Val(txtClinic), strErrinfo) Then
        MsgBox "Ok" & strErrinfo
    Else
        MsgBox "ʧ��" & strErrinfo
    End If
End Sub

Private Sub cmdWritLIS_Click()
    Dim lngID As Long, strItem As String, strErrinfo As String, blnOK As Boolean
    Dim varItem As Variant
    lngID = Val("" & txtClinic)
    If Val(txtItem) = 0 Then
        MsgBox "�����������ĿID��"
        txtItem.SetFocus
        Exit Sub
    End If
    strItem = Lis.gGetItemList(Val(txtItem))
    varItem = Split(strItem, "|")
    strItem = ""
    '������ģ�����ݣ�������Ϊ��ţ�ʵ�ʽӿ�����Ҫ��Ϊ����ļ�����
    For i = LBound(varItem) To UBound(varItem)
        strItem = strItem & "|" & Split(varItem(i), "^")(0) & "^" & i
    Next
    blnOK = Lis.gZLLisInsterReport(lngID, strItem, strErrinfo)
    If blnOK Then
        MsgBox "���" & strErrinfo
    Else
        MsgBox "ʧ��" & vbNewLine & strErrinfo
    End If
    
End Sub

Private Sub cmdȡ������_Click()
    Dim strErr As String
    If Lis.gzlLisUnRegister(Val(txt�ɼ�ID), strErr) Then
        MsgBox "��ȡ��" & strErr
    Else
        MsgBox "ʧ��" & vbNewLine & strErr
    End If
End Sub

Private Sub Form_Activate()
    Dim ctl As Object
    For Each ctl In Me.Controls
        If TypeName(ctl) = "CommandButton" Then
            ctl.Enabled = mblnReg
        End If
    Next
End Sub

Private Sub Form_Load()
  If Lis.gOpenDataBase("txyy118", "zlhis", "aqa") = False Then
    mblnReg = False
  Else
    mblnReg = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Lis.gOraDataClose
End Sub
