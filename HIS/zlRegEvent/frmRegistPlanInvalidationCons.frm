VERSION 5.00
Begin VB.Form frmRegistPlanInvalidationCons 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "����ͣ����������"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   1
      Left            =   -75
      TabIndex        =   15
      Top             =   3540
      Width           =   9345
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   0
      Left            =   15
      TabIndex        =   14
      Top             =   840
      Width           =   9345
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   315
      Index           =   3
      Left            =   5460
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2925
      Width           =   345
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   3
      Left            =   1095
      TabIndex        =   7
      Tag             =   "����"
      Top             =   2880
      Width           =   4365
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   315
      Index           =   2
      Left            =   5460
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2445
      Width           =   345
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   2
      Left            =   1095
      TabIndex        =   5
      Tag             =   "����"
      Top             =   2430
      Width           =   4365
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   315
      Index           =   1
      Left            =   5460
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1980
      Width           =   345
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   1
      Left            =   1095
      TabIndex        =   3
      Tag             =   "����"
      Top             =   1965
      Width           =   4365
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   315
      Index           =   0
      Left            =   5460
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1485
      Width           =   345
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   0
      Left            =   1095
      TabIndex        =   1
      Tag             =   "����"
      Top             =   1530
      Width           =   4365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3990
      TabIndex        =   8
      Top             =   3870
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5130
      TabIndex        =   9
      Top             =   3870
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "    ��ϸ����ָ��ͣ�����ڵĸ��ҺŰ���;����Ϊͣ�õ�ָ�����ҺŰ��ŵ��������,����֮��Ĺ�ϵΪ�ҹ�ϵ."
      Height          =   540
      Left            =   945
      TabIndex        =   16
      Top             =   420
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   300
      Picture         =   "frmRegistPlanInvalidationCons.frx":0000
      Top             =   330
      Width           =   480
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��"
      Height          =   180
      Index           =   3
      Left            =   705
      TabIndex        =   6
      Top             =   2985
      Width           =   360
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ"
      Height          =   180
      Index           =   2
      Left            =   705
      TabIndex        =   4
      Top             =   2505
      Width           =   360
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   705
      TabIndex        =   2
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   705
      TabIndex        =   0
      Top             =   1605
      Width           =   360
   End
End
Attribute VB_Name = "frmRegistPlanInvalidationCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mblnOk As Boolean
Private mstrType As String, mstrDept As String, mstr��Ŀ As String, mstrҽ�� As String
Public Function ShowCons(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    strType As String, strDept As String, str��Ŀ As String, strҽ�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�������ô���(���)
    '���:lngModule -ģ���
    '       strPrivs-Ȩ�޴�
    '����:strType -����(����ö��ŷָ�)
    '       strDept-������Ϣ(����ö��ŷָ�)
    '       str��Ŀ -�Һ���Ŀ(����ö��ŷָ�)
    '       strҽ��-ҽ��(��ʽ:Ժ��ҽ��(ID:�ö��ŷָ�)||Ժ��ҽ��(����:�ö��ŷָ�)
    '����:��ȷ��,����true,���򷵻�False
    '����:���˺�
    '����:2010-09-07 11:52:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrType = "": mstrDept = "": mstr��Ŀ = "": mstrҽ�� = ""
    mlngModule = lngModule: mstrPrivs = strPrivs: mblnOk = False
    txtEdit(0).Tag = "": txtEdit(1).Tag = "": txtEdit(2).Tag = "": txtEdit(3).Tag = ""
    Me.Show 1, frmMain
    strType = mstrType: strDept = mstrDept: str��Ŀ = mstr��Ŀ: strҽ�� = mstrҽ��
    ShowCons = mblnOk
End Function
Public Function SelectItem(ByVal intIndex As Integer, ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ֵ��ѡ����ص�����(���ڶ�ѡ)
    '���:intIndex-����
    '       strInput-�����ֵ
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-09-07 10:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCode As String, blnCancel As Boolean, rsTemp As ADODB.Recordset
    Dim strDept As String, strDeptWhere As String, strTable As String
    Dim strLike As String, strWhere As String, bytCode As Byte
    Dim strTittle As String
    Dim vRect  As RECT
    On Error GoTo Hd
    bytCode = Val(zlDatabase.GetPara("���뷽ʽ", , , 0)) + 1
    strLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    '���ܣ��๦��ѡ����,ʹ��ADO.Command��,����ʹ��[x]����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '     arrInput=��Ӧ�ĸ���SQL����ֵ,��˳����,����Ϊ��ȷ����
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б��
    If strInput <> "" Then
        strCode = strLike & strInput & "%"
        If zlCommFun.IsCharAlpha(strInput) Then
                strWhere = "(A.���� Like upper([1]) Or A.���� Like upper([1]))"
        ElseIf IsNumeric(strInput) Or zlCommFun.IsNumOrChar(strInput) Then
            strWhere = "A.���� Like upper([1])"
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strWhere = "A.���� Like [1]"
        Else
            strWhere = "(A.���� Like [1] Or A.���� Like upper([1]) Or A.���� Like upper([1]))"
        End If
    Else
        strWhere = ""
    End If
    
    Select Case intIndex
    Case 0   '����
        If strWhere <> "" Then strWhere = " WHERE " & strWhere
        strSQL = "" & _
        "   Select rownum as ID,����,����,����,ȱʡ��־,˵�� " & _
        "   From ���� A" & _
            strWhere & _
        "   Order by ����"
        strTittle = "����"
    Case 1   ' ����
        strTittle = "����"
        'ȡ�������ٴ�����
        strSQL = _
            " Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
            " From ���ű� A,��������˵�� B " & IIf(Not zlStr.IsHavePrivs(mstrPrivs, "���п���"), ",������Ա C", "") & _
            " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            "           And B.����ID=A.ID And Instr(',1,3,',',' || B.������� || ',')>0 And B.�������� = '�ٴ�'" & _
                        IIf(Not zlStr.IsHavePrivs(mstrPrivs, "���п���"), "  And A.id=C.����ID and C.��Աid =[2]", "") & _
            "           And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Case 2   '�Һ���Ŀ
         strCode = strLike & strInput & "%"
         If strInput <> "" Then
                If zlCommFun.IsCharAlpha(strInput) Then
                        strWhere = "(A.���� Like upper([1]) Or B.���� Like upper([1]) and B.���� in (3," & bytCode & "))"
                ElseIf IsNumeric(strInput) Or zlCommFun.IsNumOrChar(strInput) Then
                    strWhere = "A.���� Like upper([1])"
                ElseIf zlCommFun.IsCharChinese(strInput) Then
                    strWhere = "A.���� Like [1]"
                Else
                    strWhere = "(A.���� Like [1] Or A.���� Like upper([1]) Or B.���� Like upper([1]) and B.���� in (3," & bytCode & ") )"
                End If
                strWhere = " And " & strWhere
          Else
            strWhere = ""
          End If
            strSQL = "" & _
            "   Select Distinct A.ID, A.����, B.���� ,A.���, A.����, A.���㵥λ " & _
            "   From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
            "   Where ���='1' and A.id=B.�շ�ϸĿID  " & _
            "           And  (A.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') Or A.����ʱ�� Is Null)  " & strWhere & _
            "           And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "           And rownum<101 " & _
            "   Order by ����"
        strTittle = "�Һ���Ŀ"
    Case 3  'ҽ��
        strTittle = "ҽ��"
        strDept = "": strTable = ""
        If txtEdit(1).Tag <> "" Then
            strDept = Trim(txtEdit(1).Tag)
            If InStr(1, strDept, ",") > 0 Then
                If zlCommFun.ActualLen(strDept) > 1990 Then
                    strTable = "Select Column_Value as ID from Table(Cast(f_Num2list([4]) As zlTools.t_Numlist))  "
                Else
                    strTable = " Select ID From ���ű� where id in (" & strDept & ") "
                End If
                strTable = ",(" & strTable & ") E"
                strDeptWhere = " C.����ID=E.ID"
            Else
                strDeptWhere = " And C.����id  =[3]"
            End If
        End If
        strWhere = Replace(strWhere, "����", "����")
        If strWhere <> "" Then strWhere = " And " & strWhere
        strSQL = _
        "   Select /*+ rule */ distinct A.ID,A.��� as ����,A.���� as ����,A.���� " & _
        "   From ��Ա�� A ,��Ա����˵�� B, ������Ա C" & strTable & vbCrLf & _
        "   Where A.ID=B.��Աid And A.id=C.��Աid  " & _
        "           And  B.��Ա����='ҽ��' And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & vbCrLf & _
        "           And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    strDeptWhere & Replace(strWhere, "����", "���") & _
        "   Union ALL  " & _
        "   Select ID,����,���� as ����,����  " & _
        "   From ( " & _
        "               Select Distinct -1*rownum  as ID,'' as ����,A.ҽ������ as ����, zlspellcode(ҽ������) as ����" & _
        "               From �ҺŰ��� A" & strTable & _
        "               where A.ҽ��ID is null " & Replace(UCase(strDeptWhere), "C.����ID", "A.����ID") & _
        "           ) A " & _
        "  Where 1=1 " & strWhere
    Case Else
        Exit Function
    End Select
    
    vRect = zlcontrol.GetControlRect(txtEdit(intIndex).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, strTittle & "ѡ��", False, "", "��ѡ��", False, False, True, vRect.Left, vRect.Top, txtEdit(intIndex).Height, blnCancel, True, True, strCode, UserInfo.ID, Val(strDept), strDept)
    If blnCancel = True Then
        If txtEdit(intIndex).Enabled And txtEdit(intIndex).Visible Then txtEdit(intIndex).SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "û���ҵ�����������" & strTittle & "������!", vbInformation + vbOKOnly, gstrSysName
        If txtEdit(intIndex).Enabled And txtEdit(intIndex).Visible Then txtEdit(intIndex).SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "û���ҵ�����������" & strTittle & "������!", vbInformation + vbOKOnly, gstrSysName
        If txtEdit(intIndex).Enabled And txtEdit(intIndex).Visible Then txtEdit(intIndex).SetFocus
        Exit Function
    End If
    Dim strText As String, strValues As String, strValues1 As String
    With rsTemp
        Do While Not .EOF
            strText = strText & ";" & Nvl(rsTemp!����)
            If intIndex <> 0 Then
                If intIndex = 3 And Val(Nvl(rsTemp!ID)) < 0 Then
                    strValues1 = strValues1 & "," & Nvl(rsTemp!����)
                Else
                    strValues = strValues & "," & Nvl(rsTemp!ID)
                End If
            Else
                strValues = strValues & "," & Nvl(rsTemp!����)
            End If
            .MoveNext
        Loop
        If strText <> "" Then strText = Mid(strText, 2)
        If strValues <> "" Then strValues = Mid(strValues, 2)
        If strValues1 <> "" Then strValues1 = "||" & Mid(strValues1, 2)
        txtEdit(intIndex).Text = strText: txtEdit(intIndex).Tag = strValues & strValues1
    End With
    zlCommFun.PressKey vbKeyTab
    SelectItem = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrType = txtEdit(0).Tag: mstrDept = txtEdit(1).Tag: mstr��Ŀ = txtEdit(2).Tag: mstrҽ�� = txtEdit(3).Tag
    If mstrType = "" And mstrDept = "" And mstr��Ŀ = "" And mstrҽ�� = "" Then
        MsgBox "δѡ��һ�����������ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If txtEdit(0).Text <> "" And mstrType = "" Then
        MsgBox "ע��:" & vbCrLf & "    ����ѡ������(������δ���س�����ѡ��)������!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If txtEdit(1).Text <> "" And mstrDept = "" Then
        MsgBox "ע��:" & vbCrLf & "    ����ѡ������(������δ���س�����ѡ��)������!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If txtEdit(2).Text <> "" And mstr��Ŀ = "" Then
        MsgBox "ע��:" & vbCrLf & "    �Һ���Ŀѡ������(������δ���س�����ѡ��)������!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If txtEdit(3).Text <> "" And mstrҽ�� = "" Then
        MsgBox "ע��:" & vbCrLf & "    ҽ��ѡ������(������δ���س�����ѡ��)������!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdSel_Click(Index As Integer)
    If SelectItem(Index, "") = False Then
        Exit Sub
    End If
End Sub
Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    zlcontrol.TxtSelAll txtEdit(Index)
End Sub
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txtEdit(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txtEdit(Index).Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If SelectItem(Index, Trim(txtEdit(Index).Text)) = False Then
        Exit Sub
    End If
End Sub
