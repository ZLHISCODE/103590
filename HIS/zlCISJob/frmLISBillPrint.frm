VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLISBillPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�걾�����ӡ"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboCapture 
      Height          =   300
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1500
      Width           =   2865
   End
   Begin VB.CommandButton cmdInter 
      Caption         =   "�ж�(F9)"
      Height          =   350
      Left            =   90
      TabIndex        =   11
      Top             =   3105
      Width           =   1100
   End
   Begin VB.CheckBox chkMachine 
      Caption         =   "�������ֱ��ӡ(&S)"
      Height          =   225
      Left            =   405
      TabIndex        =   5
      ToolTipText     =   "ѡ�д�ѡ���ʱ�����ֻҪ��ͬһִ�п��ҡ�ͬ�ֱ걾ֻ��ӡһ�����롣����ÿһ���ɼ����ֱ��ӡ��"
      Top             =   2250
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkRetry 
      Caption         =   "�Ѵ�ӡ�����´�ӡ(&R)"
      Height          =   225
      Left            =   135
      TabIndex        =   6
      Top             =   2610
      Width           =   3315
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   8
      Top             =   3105
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   2700
      TabIndex        =   7
      Top             =   3105
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   2955
      Width           =   4965
   End
   Begin VB.CheckBox chkOnly 
      Caption         =   "ͬһ���˵�ͬ���걾�ϲ���ӡ(&O)"
      Height          =   225
      Left            =   135
      TabIndex        =   4
      ToolTipText     =   "ѡ�д�ѡ���ʱ�����ֻҪ��ͬһִ�п��ҡ�ͬ�ֱ걾ֻ��ӡһ�����롣����ÿһ���ɼ����ֱ��ӡ��"
      Top             =   1920
      Value           =   1  'Checked
      Width           =   3645
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   97386499
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   97386499
      CurrentDate     =   38082
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ӡ����Ĳɼ���ʽ"
      Height          =   180
      Left            =   150
      TabIndex        =   12
      Top             =   1560
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��                      ��"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   2880
   End
   Begin VB.Label lblDesc 
      Appearance      =   0  'Flat
      Caption         =   $"frmLISBillPrint.frx":0000
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   750
      TabIndex        =   9
      Top             =   120
      Width           =   4170
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmLISBillPrint.frx":008C
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmLISBillPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strRecentDate As String '�����ӡ��ҽ������ʱ��
Private strPatiSource As String '������Դ
Private lngDeptID As Long
Private blnCancel As Boolean '�Ƿ�ȡ����ӡ��ҵ

Public Sub ShowMe(objParent As Object, ByVal PatiSource As String, DeptID As Long)
    strPatiSource = PatiSource: lngDeptID = DeptID
    blnCancel = False
    
    Me.Show vbModal, objParent
    Unload Me
End Sub

Private Sub chkOnly_Click()
    Me.chkMachine.Enabled = (Me.chkOnly.Value = 1)
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdInter_Click()
    blnCancel = True
End Sub

Private Sub cmdOK_Click()
    '�����ӡ����
        
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\�����ӡ", "���ҽ��ʱ��", Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\�����ӡ", "�ɼ���ʽ", cboCapture.Text
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\�����ӡ", "��ӡ��ʽ", IIf(Me.chkMachine, "������", IIf(Me.chkOnly, "���걾", ""))
    If PrintBill Then Me.Hide
End Sub

Private Function PrintBill() As Boolean
    Dim strSQL As String
    Dim strDateFilter As String
    Dim rsTmp As New ADODB.Recordset
    Dim strNO As String, int���� As Integer
    
    PrintBill = False
    On Error GoTo DataError
    Me.MousePointer = vbHourglass
    
    If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
        strDateFilter = " And A.����ʱ�� Between [2] And Sysdate"
    Else
        strDateFilter = " And A.����ʱ�� Between [2] And [3]"
    End If

    If chkOnly.Value = 1 Then
        If chkMachine.Value = 0 Then
            'ͬһ�걾ֻ��һ��
            strSQL = "Select ����ID,�걾,ִ�в���,NO," & _
                " Trim(����1||' '||����2||' '||����3||' '||����4||' '||����5) As ��Ŀ,���" & _
                " From" & _
                " (Select B.����ID,B.�걾��λ As �걾,F.���� As ִ�в���,S.���," & _
                "  Max(Decode(Mod(Rownum,5),0,B.ҽ������,'')) As ����1," & _
                "  Max(Decode(Mod(Rownum,5),1,B.ҽ������,'')) As ����2," & _
                "  Max(Decode(Mod(Rownum,5),2,B.ҽ������,'')) As ����3," & _
                "  Max(Decode(Mod(Rownum,5),3,B.ҽ������,'')) As ����4," & _
                "  Max(Decode(Mod(Rownum,5),4,B.ҽ������,'')) As ����5," & _
                "  Max(S.NO||','||S.��¼����) As NO" & _
                "  From ����ҽ����¼ B,���ű� F," & _
                "   (Select A.ҽ��ID,A.NO,A.��¼����,B.������ĿID," & _
                "    'ZLCISBILL'||trim(to_Char(F.���, '00000'))||'-1' AS ���" & _
                "    From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C, ��������Ӧ�� E,�����ļ��б� F" & _
                "    Where a.ҽ��ID = B.ID And B.������ĿID = C.ID AND B.������ĿID=E.������ĿID AND B.������Դ=E.Ӧ�ó��� AND E.�����ļ�ID=F.ID" & _
                "     And instr([4],','||B.������Դ||',') > 0  And A.ִ�в���ID+0= [1] " & _
                "     And C.���='E' And Nvl(C.��������,'0')='6'" & IIf(cboCapture.ItemData(cboCapture.ListIndex) = 0, "", " And C.ID+0=[5]") & _
                strDateFilter & " And Nvl(A.ִ��״̬,0)=0" & IIf(chkRetry, "", " And A.������ Is Null") & ") S" & _
                "  Where B.ִ�п���ID = F.ID And B.���ID = S.ҽ��ID" & _
                "  Group By B.����ID, B.�걾��λ,F.����,S.������ĿID,S.���)" & _
                " Order By ����ID"
        Else
            'ͬһ�걾�ٰ������ֱ��ӡ
            strSQL = "Select ����ID,�걾,ִ�в���,NO," & _
                " Trim(����1||' '||����2||' '||����3||' '||����4||' '||����5) As ��Ŀ,����,���" & _
                " From" & _
                " (Select B.����ID,B.�걾��λ As �걾,F.���� As ִ�в���,S.����,S.���," & _
                "  Max(Decode(Mod(Rownum,5),0,B.ҽ������,'')) As ����1," & _
                "  Max(Decode(Mod(Rownum,5),1,B.ҽ������,'')) As ����2," & _
                "  Max(Decode(Mod(Rownum,5),2,B.ҽ������,'')) As ����3," & _
                "  Max(Decode(Mod(Rownum,5),3,B.ҽ������,'')) As ����4," & _
                "  Max(Decode(Mod(Rownum,5),4,B.ҽ������,'')) As ����5," & _
                "  Max(S.NO||','||S.��¼����) As NO" & _
                "  From ����ҽ����¼ B,���ű� F," & _
                "   (Select DISTINCT ҽ��ID,NO,��¼����,����,������ĿID,��� FROM " & _
                "    (Select A.ҽ��ID,A.NO,A.��¼����,B.������ĿID,I.������ĿID," & _
                "     'ZLCISBILL'||trim(to_Char(F.���, '00000'))||'-1' AS ���,MAX(Decode(M.����,NULL,'�ֹ�',M.����)) AS ���� " & _
                "     From ����ҽ������ A,����ҽ����¼ B,����ҽ����¼ D,������ĿĿ¼ C,���鱨����Ŀ I,����������Ŀ J,�������� M, ��������Ӧ�� E,�����ļ��б� F" & _
                "     Where a.ҽ��ID = B.ID And B.������ĿID = C.ID" & _
                "      AND D.���ID = B.ID AND D.������ĿID=I.������ĿID(+) AND I.������ĿID=J.��ĿID(+) AND J.����ID=M.ID(+) AND B.������ĿID=E.������ĿID AND B.������Դ=E.Ӧ�ó��� AND E.�����ļ�id=F.ID" & _
                "      And instr([4],','||B.������Դ||',') > 0 And A.ִ�в���ID+0= [1] " & _
                "      And C.���='E' And Nvl(C.��������,'0')='6'" & IIf(cboCapture.ItemData(cboCapture.ListIndex) = 0, "", " And C.ID+0=[5]") & _
                strDateFilter & " And Nvl(A.ִ��״̬,0)=0" & IIf(chkRetry, "", " And A.������ Is Null") & _
                "     GROUP BY A.ҽ��ID,A.NO,A.��¼����,B.������ĿID,I.������ĿID,F.���)" & _
                "   ) S" & _
                "  Where B.ִ�п���ID = F.ID And B.���ID = S.ҽ��ID" & _
                "  Group By B.����ID, B.�걾��λ,F.����,S.����,S.������ĿID,S.���)" & _
                " Order By ����ID"
        End If
    Else
        '�ֱ��ӡ
        strSQL = "Select B.����ID,B.�걾��λ as �걾,F.���� As ִ�в���, B.ҽ������ As ��Ŀ,S.NO||','||S.��¼���� As NO,���" & _
            " From ����ҽ����¼ B,���ű� F," & _
            " (Select A.ҽ��ID,A.NO,A.��¼����,'ZLCISBILL'||trim(to_Char(F.���, '00000'))||'-1' AS ��� From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C, ��������Ӧ�� E,�����ļ��б� F" & _
            "  Where a.ҽ��ID = B.ID And B.������ĿID = C.ID AND B.������ĿID=E.������ĿID AND B.������Դ=E.Ӧ�ó��� AND E.�����ļ�id=F.ID" & _
            "   And instr([4],','||B.������Դ||',') > 0 And A.ִ�в���ID+0= [1] " & _
            "   And C.���='E' And Nvl(C.��������,'0')='6'" & IIf(cboCapture.ItemData(cboCapture.ListIndex) = 0, "", " And C.ID+0=[5]") & _
            strDateFilter & " And Nvl(A.ִ��״̬,0)=0" & IIf(chkRetry, "", " And A.������ Is Null") & ") S" & _
            " Where B.ִ�п���ID = F.ID And B.���ID = S.ҽ��ID" & _
            " Order By ����ID"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID, CDate(Format(dtpBegin.Value, "yyyy-MM-dd hh:mm:ss")), _
                    CDate(Format(dtpEnd.Value, "yyyy-MM-dd hh:mm:ss")), "," & strPatiSource & ",", cboCapture.ItemData(cboCapture.ListIndex))
    
    If rsTmp.EOF Then
        Me.MousePointer = vbDefault
        MsgBox "�ڸ�ʱ����û����Ҫ��ӡ�ı걾���롣", vbInformation, gstrSysName
        Exit Function
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, Nvl(rsTmp("���")), Me) Then
        cmdInter.Enabled = True
        Do While Not rsTmp.EOF
            If blnCancel Then PrintBill = True: Exit Function
            strNO = Split(rsTmp("NO"), ",")(0)
            int���� = Split(rsTmp("NO"), ",")(1)
            DoEvents
            Call ReportOpen(gcnOracle, glngSys, Nvl(rsTmp("���")), Me, "NO=" & strNO, "����=" & int����, "��Ŀ=" & Nvl(rsTmp("��Ŀ")), 2)
            
            rsTmp.MoveNext
        Loop
        cmdInter.Enabled = False
        '��д�����ˡ�����ʱ�䣬��ʾ�Ѿ���ӡ
        strSQL = "ZL_����ҽ��ִ��_��������('" & strPatiSource & "'," & lngDeptID & ",'" & _
            Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','" & _
            IIf(Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm"), "", _
                Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm")) & "','" & UserInfo.���� & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Name
    End If
    PrintBill = True
    
    Me.MousePointer = vbDefault
    Exit Function
DataError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
        
    Me.MousePointer = vbDefault
End Function

Private Sub Form_Activate()
    Dim curDate As Date
    
    cmdInter.Enabled = False
    On Error GoTo DataError
    
    curDate = zlDatabase.Currentdate
    dtpEnd.MaxDate = curDate: dtpBegin.MaxDate = curDate
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd HH:mm")
    dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
    dtpBegin.Value = Format(strRecentDate, "yyyy-MM-dd HH:mm")
        
    dtpBegin.SetFocus
    Exit Sub
DataError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
    'F9�ж�
    If KeyCode = 120 And cmdInter.Enabled Then cmdInter_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String
    On Error GoTo DataError
    
    '��ȡ����
    strRecentDate = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\�����ӡ", "���ҽ��ʱ��", Format(Date, "yyyy-MM-dd HH:mm"))
    If Not IsDate(strRecentDate) Then strRecentDate = Format(Date, "yyyy-MM-dd HH:mm")
    
    '��ʼ�ɼ���ʽ
    strSQL = "Select Distinct A.ID,A.����" & _
        " From ������ĿĿ¼ A,����ִ�п��� B " & _
        " Where A.���='E' AND A.��������='6'" & _
        " And (A.����ʱ�� IS NULL Or A.����ʱ��=To_Date('3000-01-01','yyyy-mm-dd')) " & _
        " And A.ID=B.������ĿID And B.ִ�п���ID=[1]" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID)
    With cboCapture
        .AddItem "���з�ʽ"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
        Do While Not rsTmp.EOF
            .AddItem rsTmp("����")
            .ItemData(.NewIndex) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        On Error Resume Next
        strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\�����ӡ", "�ɼ���ʽ", "���з�ʽ")
        .Text = strTmp
    End With
    strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\�����ӡ", "��ӡ��ʽ", "")
    If strTmp = "������" Then
        Me.chkOnly.Value = 1
        Me.chkMachine.Value = 1
    ElseIf strTmp = "���걾" Then
        Me.chkOnly.Value = 1
        Me.chkMachine.Value = 0
    Else
        Me.chkOnly.Value = 0
        Me.chkMachine.Value = 0
    End If
    
    Exit Sub
DataError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
