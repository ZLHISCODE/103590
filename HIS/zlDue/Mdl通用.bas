Attribute VB_Name = "Mdlͨ��"
Option Explicit
Public Enum gEditType
     g���� = 0
     g�޸� = 1
     g��� = 2
     gȡ�� = 3
     g�鿴 = 4
     gԤ�� = 7
End Enum
Public Enum RecBillStatus  '��¼״̬��Ϣ
    ������¼ = 1
    ������¼ = 2
    ��������¼ = 3
End Enum
Public Enum ErrBillStatusInfor  '����״̬��Ϣ
    ������� = 1
    �Ѿ�ɾ��
    �Ѿ����
    �Ѿ�����
End Enum
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

Public Const glngGetFocus As Long = &HA87B82                    '����ʱ��ѡ����ɫ
Public Const glngGetFocus_Font As Long = &H80000005             '����ʱ��������ɫ
Public Const glngLostFocus As Long = &HC0C0C0                   '�뿪ʱ��ѡ��ɫ
Public Const glngLostFocus_Font As Long = &H80000008            '�뿪ʱ������ɫ

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '����:���ؼ�ƥ�䴮%dd%,�����Ǵ�д
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper = False Then
        GetMatchingSting = strLeft & strString & strRight
    Else
        GetMatchingSting = strLeft & UCase(strString) & strRight
    End If
End Function

Public Function MulitSelectPersion(ByVal frmParent As Form, ByVal objCtl As Object, _
    ByVal strKey As String, Optional lng����ID As Long = 0, _
    Optional ByRef lng��ԱID As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ������Ա
    '���:frmParent-���õĸ�����
    '     objCtl-�ؼ�(Ŀǰֻ֧���ı���)
    '     strKey-����Ľ�ֵ
    '     lng����ID-�����Ϊ��,��������Ա,����, ��ָ�������µ���Ա
    '����:lng��Աid-������ԱID
    '����:���ҳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/23
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    
    'zlDatabase.ShowSQLSelect
    '���ܣ��๦��ѡ����
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
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    Err = 0: On Error GoTo ErrHand:
    
     
     If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
        If lng����ID = 0 Then
            gstrSQL = "" & _
                "   Select ID,���,����,����,����,�Ա�,����,��������,�칫�ҵ绰" & _
                "   From ��Ա�� " & _
                "   Where (���� like [1] or ��� like [1] or ���� like [1] or ���� like [1]) " & zl_��ȡվ������ & "" & _
                "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
                "   order by ���"
        Else
            gstrSQL = "" & _
                "   Select distinct a.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰" & _
                "   From ��Ա�� a,������Ա C " & _
                "   Where a.id=c.��Աid and c.����Id=[2]   " & zl_��ȡվ������(True, "a") & " and (a.���� like [1] or a.��� like [1] or a.���� like [1] or a.���� like [1]) " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & _
                "   order by ���"
        End If
     Else
        If lng����ID = 0 Then
            gstrSQL = "" & _
                "   Select ID,���,����,����,����,�Ա�,����,��������,�칫�ҵ绰" & _
                "   From ��Ա�� " & _
                "   Where (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) " & zl_��ȡվ������ & "" & _
                "   order by ���"
        Else
            gstrSQL = "" & _
                "   Select distinct a.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰" & _
                "   From ��Ա�� a,������Ա C " & _
                "   Where a.id=c.��Աid and c.����Id=[2] " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)  " & zl_��ȡվ������(True, "a") & "" & _
                "   order by ���"
        End If
    End If
    
    If UCase(TypeName(objCtl)) = "TEXTBOX" Then
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(frmParent, gstrSQL, 0, "��Աѡ����", False, "", "��Աѡ��", False, False, True, vRect.Left - 15, vRect.Top, objCtl.Height, blnCancel, False, False, strKey, lng����ID)
    Else
        Dim sngX As Single, sngY As Single
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        Set rsTemp = zlDatabase.ShowSQLSelect(frmParent, gstrSQL, 0, "��Աѡ����", False, "", "��Աѡ��", False, False, True, sngX, sngY - objCtl.MsfObj.CellHeight, objCtl.MsfObj.CellHeight, blnCancel, False, False, strKey, lng����ID)
    End If
    lng��ԱID = 0
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then
        ShowMsgbox "δ�ҵ�ָ������Ա,����!"
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = "TEXTBOX" Then
        objCtl.Text = Nvl(rsTemp!����)
    Else
        objCtl.TextMatrix(objCtl.Row, objCtl.Col) = Nvl(rsTemp!����)
        objCtl.Text = Nvl(rsTemp!����)
    End If
    lng��ԱID = Val(Nvl(rsTemp!ID))
    rsTemp.Close
    MulitSelectPersion = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


