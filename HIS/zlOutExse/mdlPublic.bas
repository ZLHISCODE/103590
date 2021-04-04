Attribute VB_Name = "mdlPublic"
Option Explicit 'Ҫ���������
'ϵͳ������ʱ����
Public glngSys As Long
Public glngModul As Long
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���

Public gstrSQL As String
Public gblnOK As Boolean

Public gstrSysName As String                'ϵͳ����
Public gstrUnitName As String '�û���λ����
Public gstrDBUser As String '��ǰ���ݿ��û���
Public gfrmMain As Object
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��

'-----------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    �������� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO


Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
'----------------------------------------
Public Const LONG_MAX = 2147483647 'Long�����ֵ
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public glngOld As Long, glngFormW As Long, glngFormH As Long

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Windows���----------------------------------
Public Const ETO_OPAQUE = 2
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'---------------------------------------------
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'���뷨����API----------------------------------------------------------------------------------------------
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type TY_WindowsRect
    MaxW As Long
    MaxH As Long
    MinW  As Long
    MinH As Long
End Type
Public gWinRect As TY_WindowsRect

Public Function SetWindowResizeWndMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = gWinRect.MinW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = gWinRect.MinH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = gWinRect.MaxW \ Screen.TwipsPerPixelX
        MinMax.ptMaxTrackSize.Y = gWinRect.MaxH \ Screen.TwipsPerPixelY
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SetWindowResizeWndMessage = 1
        Exit Function
    End If
    SetWindowResizeWndMessage = CallWindowProc(glngOld, hWnd, msg, wp, lp)
End Function

Public Function SetWidth(cboHwnd As Long, NewWidthPixel As Long) As Boolean
'���ܣ����� Combo �����Ŀ��,��λΪ pixels
    '�ȸ�Ϊtwip��λ
    cbo.SetListWidth cboHwnd, NewWidthPixel * Screen.TwipsPerPixelX
    
    SetWidth = True
End Function

Public Function GetFirstRow(objBill As ExpenseBill, intPage As Long, Optional ByVal strClass As String) As Integer
'���ܣ���ȡ��ǰ�����е�һ��ΪҩƷ���к�
'������strClass=�Ƿ�ֻȡָ�����ҩƷ��
'���أ�0=û��ҩƷ�շ���
    Dim i As Integer
    
    For i = 1 To objBill.Pages(intPage).Details.Count
        If strClass = "" Then
            If InStr(",5,6,7,", objBill.Pages(intPage).Details(i).�շ����) > 0 Then
                GetFirstRow = i: Exit Function
            End If
        Else
            If objBill.Pages(intPage).Details(i).�շ���� = strClass Then
                GetFirstRow = i: Exit Function
            End If
        End If
    Next
End Function

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'���ܣ���ָ�����ֱҴ��������д���,���ش����Ľ��
'������curMoney=Ҫ���зֱҴ���Ľ��(ΪӦ�ɽ��,2λС��)
'      gBytMoney=
'         0.������
'         1.��ȡ�������뷨,eg:0.51=0.50;0.56=0.60
'         2.�����շ�,eg:0.51=0.60,0.56=0.60
'         3.����շ�,eg:0.51=0.50,0.56=0.50
'         4.�����������˫,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           �����������˫,����ҹ���ѧ����ίԱ����ʽ�䲼�ġ�������Լ����,������vb��Round����,�������������ְ�����λ����ʱ�����Ը����ֽ���������Լ
'           �����м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ
'         5.�������塢�������,�Խǽ��д�������Ҫ�ȶԷֱҽ�������,��0.29(��)���¶�����ǣ�0.80(��)���϶����ǣ�0.3-0.79����Ϊ0.5��
'         6.��������:eg:0.15=0.10:0.16=0.2:���˺�:34519
'91385,������5.�������塢������롱�����ȶԷֱҽ����������룬��0.24(��)���¶�����ǣ�0.75(��)���϶����ǣ�0.25-0.74������Ϊ0.5
'       �ֱ����������룬��ô0.00��0.24=0��0.25��0.5=0.50, 0.50��0.74=0.50��0.75��1.00=1������������ռ50%�ı���
    
    Dim intSign As Integer, curTmp As Currency

    If gBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf gBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '��ȡ��λ���,�ٴ���ֱ�,��:0.248 ��0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf gBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf gBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf gBytMoney = 4 Then
        CentMoney = Format(Round(curMoney, 1), "0.00")
    ElseIf gBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf gBytMoney = 6 Then
         '���˺� ����:34519 ��������:eg:0.15=0.10:0.16=0.2:    ����:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, msg, wp, lp)
End Function

Public Function strPad(ByVal strPre As String, ByVal intLen As Integer, ByVal strFill As String, ByVal bytAlign As Byte, Optional ByVal blnTrim As Boolean) As String
'���ܣ�����ַ���
'������
'     strPre=Ҫ�����ַ���
'     intLen=����ĳ���
'     strFill=Ҫ�����ַ�
'     bytAlign=1,2/��,�Ҷ��룬�����ʱ����ԭ�ַ����ұ����
'     blnTrim=���ַ�������ʱ���Ƿ�ǿ�а�ָ�����Ƚ�ȡ��
'���أ��������ַ���
'˵����һ�����ֵ��������ַ����ȴ���
    Dim i As Integer
    
    If LenB(StrConv(strPre, vbFromUnicode)) >= intLen Then
        If blnTrim Then
            For i = 1 To Len(strPre)
                strPad = strPad & Mid(strPre, i, 1)
                If LenB(StrConv(strPad, vbFromUnicode)) >= intLen Then Exit For
            Next
        Else
            strPad = strPre
        End If
    Else
        If Len(strFill) > 1 Then strFill = Left(strFill, 1)
        If bytAlign = 1 Then
            strPad = strPre
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
        ElseIf bytAlign = 2 Then
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
            strPad = strPad & strPre
        End If
    End If
End Function

Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'���ܣ��ж��������Ƿ���ԭ�ۺ��ִ��޶��ķ�Χ��
'������varL=ԭ��,varR=�ּ�,varI=������
'���أ�������ڷ�Χ��,��Ϊ��ʾ��Ϣ,����Ϊ�մ�
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '�����ֵ������ͬ,���þ���ֵ�ж�
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "����ļ۸����ֵ���ڷ�Χ(" & FormatEx(Abs(varL), 5) & "-" & FormatEx(Abs(varR), 5) & ")��."
        End If
    Else
        '������Ų���ͬ,����ԭʼ��Χ�ж�
        If varI < varL Or varI > varR Then
            CheckScope = "����ļ۸�ֵ���ڷ�Χ(" & FormatEx(varL, 5) & "-" & FormatEx(varR, 5) & ")��."
        End If
    End If
End Function

Public Sub ExChangeLocate(objA As Object, objB As Object)
'���ܣ�����ҽ���Ϳ������ҵ�����λ��
    Dim x1 As Long, y1 As Long, w1 As Long, t1 As Integer
    Dim x2 As Long, y2 As Long, w2 As Long, t2 As Integer
    Dim obj1 As Object, obj2 As Object
    
    x1 = objA.Left
    y1 = objA.Top
    w1 = objA.Width
    t1 = objA.TabIndex
    Set obj1 = objA.Container

    x2 = objB.Left
    y2 = objB.Top
    w2 = objB.Width
    t2 = objB.TabIndex
    Set obj2 = objB.Container
    
    Set objB.Container = obj1
    objB.Left = x1
    objB.Top = y1
    objB.Width = w1
    objB.TabIndex = t1
    
    Set objA.Container = obj2
    objA.Left = x2
    objA.Top = y2
    objA.Width = w2
    objA.TabIndex = t2
End Sub

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function MakeBillRecord(objBill As ExpenseBill, ByVal bln���� As Boolean, ByVal intPage As Integer, _
    ByVal strDate As String, ByVal str�ѱ� As String, ByVal strInvoice As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݶ������ݴ���һ����¼��Ϣ(���ۼ۵�λ)
    '���:intPage=�൥���շ�ģʽʱ��ָ���ĵ���,���Ϊ��,��ʾȫ������
    '        strDate=����ʱ��,
    '        strInvoice=Ʊ�ݺ�
    '����:
    '����:ҽ��������ݵ����ݼ�(�������(1--n),����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��)
    '����:���˺�
    '����:2011-08-15 16:40:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, intStartPage As Integer, intPages As Integer
    Dim p As Integer, strSQL As String
    Dim dbl���� As Double, curʵ�� As Currency, curͳ�� As Currency
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand:
    rsTmp.Fields.Append "�������", adBigInt, 50, adFldIsNullable
    rsTmp.Fields.Append "�ѱ�", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsTmp.Fields.Append "���", adBigInt, , adFldIsNullable '����:42961
    rsTmp.Fields.Append "ʵ��Ʊ��", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "����ʱ��", adDBTimeStamp, , adFldIsNullable
    rsTmp.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "���㵥λ", adVarChar, 50, adFldIsNullable
    '69788:���ϴ�,2014-6-5,�����������ֶδ�С����20��Ϊ100
    rsTmp.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "ͳ����", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "����֧������ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�Ƿ�ҽ��", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "���ձ���", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "ժҪ", adVarChar, 2000, adFldIsNullable
    rsTmp.Fields.Append "�Ƿ���", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��������ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "ִ�в���ID", adBigInt, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    intStartPage = IIf(intPage <= 0, 1, intPage)
    intPages = IIf(intPage <= 0, objBill.Pages.Count, intPage)
    For p = intStartPage To intPages
         If objBill.Pages(p).NO <> "" Then       '��ȡ���ǻ��۵�
                '��ȡ�Ļ��۵�(�ۼ۵�λ)
                strSQL = _
                "Select '" & strInvoice & "' as ʵ��Ʊ��,NO,Nvl( �۸񸸺�, ���) as ���,To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
                        objBill.����ID & " As ����ID,'" & str�ѱ� & "' As �ѱ�,�շ����,�վݷ�Ŀ,���㵥λ,������," & _
                "       �շ�ϸĿID,���մ���ID As ����֧������ID,Nvl(������Ŀ��,0) As �Ƿ�ҽ��,���ձ���," & _
                "       Avg(Nvl(����,0)*����) As ����,Avg(��׼����) As ����," & _
                "       Sum(ʵ�ս��) As ʵ�ս��,Sum(ͳ����) As ͳ����,ժҪ," & _
                        IIf(bln����, "1", "0") & " as �Ƿ���,��������ID,ִ�в���ID From ������ü�¼" & _
                " Where ��¼����=1 And ��¼״̬=0 And NO=[1]" & _
                " Group By Nvl(�۸񸸺�,���),�շ����,�վݷ�Ŀ,���㵥λ,������," & _
                "       �շ�ϸĿID,���մ���ID,Nvl(������Ŀ��,0),���ձ���,ժҪ,��������ID,ִ�в���ID,NO" & _
                " Order by  ��� "
                Set rsNo = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���۵�����-ҽ��", objBill.Pages(p).NO)
                If rsNo.RecordCount <> 0 Then rsNo.MoveFirst
                Do While Not rsNo.EOF
                    rsTmp.AddNew
                    rsTmp!������� = p
                    rsTmp!�ѱ� = str�ѱ�
                    rsTmp!NO = NVL(rsNo!NO)   '����ȡ���۵�ʱ����ֵ
                    rsTmp!��� = Val(NVL(rsNo!���))   '����ȡ���۵�ʱ����ֵ
                    rsTmp!ʵ��Ʊ�� = strInvoice
                    rsTmp!����ʱ�� = CDate(strDate)
                    rsTmp!����ID = IIf(objBill.����ID = 0, Null, objBill.����ID)
                    rsTmp!�շ���� = NVL(rsNo!�շ����)
                    rsTmp!�վݷ�Ŀ = NVL(rsNo!�վݷ�Ŀ)
                    rsTmp!������ = NVL(rsNo!������)
                    rsTmp!�շ�ϸĿID = Val(NVL(rsNo!�շ�ϸĿID))
                    rsTmp!���㵥λ = NVL(rsNo!���㵥λ)
                    rsTmp!���� = Val(NVL(rsNo!����))
                    rsTmp!���� = Val(NVL(rsNo!����))
                    rsTmp!ʵ�ս�� = Val(NVL(rsNo!ʵ�ս��))
                    rsTmp!ͳ���� = Val(NVL(rsNo!ͳ����))
                    rsTmp!����֧������ID = IIf(Val(NVL(rsNo!����֧������ID)) = 0, Null, Val(NVL(rsNo!����֧������ID)))
                    rsTmp!�Ƿ�ҽ�� = Val(NVL(rsNo!�Ƿ�ҽ��))
                    rsTmp!���ձ��� = NVL(rsNo!���ձ���)
                    rsTmp!ժҪ = NVL(rsNo!ժҪ)
                    rsTmp!�Ƿ��� = IIf(bln����, 1, 0)
                    rsTmp!��������ID = Val(NVL(rsNo!��������ID))
                    rsTmp!ִ�в���ID = Val(NVL(rsNo!ִ�в���ID))
                    rsTmp.Update
                    rsNo.MoveNext
                Loop
         Else
            For i = 1 To objBill.Pages(p).Details.Count
                dbl���� = 0: curʵ�� = 0: curͳ�� = 0
                With objBill.Pages(p).Details(i)
                    For j = 1 To .InComes.Count
                        dbl���� = dbl���� + .InComes(j).��׼����
                        curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                        curͳ�� = curͳ�� + .InComes(j).ͳ����
                    Next
                    rsTmp.AddNew
                    rsTmp!������� = p
                    rsTmp!�ѱ� = str�ѱ�
                    rsTmp!NO = ""   '����ȡ���۵�ʱ����ֵ
                    rsTmp!��� = i
                    rsTmp!ʵ��Ʊ�� = strInvoice
                    rsTmp!����ʱ�� = CDate(strDate)
                    rsTmp!����ID = IIf(objBill.����ID = 0, Null, objBill.����ID)
                    rsTmp!�շ���� = .�շ����
                    If .InComes.Count > 0 Then
                        rsTmp!�վݷ�Ŀ = .InComes(1).�վݷ�Ŀ
                    Else
                        rsTmp!�վݷ�Ŀ = Null
                    End If
                    rsTmp!������ = objBill.Pages(p).������
                    
                    rsTmp!�շ�ϸĿID = .�շ�ϸĿID
                    
                    rsTmp!���㵥λ = .���㵥λ
                    If InStr(",5,6,7,", .�շ����) > 0 And gblnҩ����λ Then
                        '��ҩ����λת��Ϊ�ۼ۵�λ
                        rsTmp!���� = IIf(.���� = 0, 1, .����) * .���� * .Detail.ҩ����װ
                        rsTmp!���� = Format(dbl���� / .Detail.ҩ����װ, gstrFeePrecisionFmt)
                    Else
                        rsTmp!���� = IIf(.���� = 0, 1, .����) * .����
                        rsTmp!���� = Format(dbl����, gstrFeePrecisionFmt)
                    End If
                    rsTmp!ʵ�ս�� = Format(curʵ��, gstrDec)
                    rsTmp!ͳ���� = Format(curͳ��, gstrDec)
                    rsTmp!����֧������ID = IIf(.���մ���ID = 0, Null, .���մ���ID)
                    rsTmp!�Ƿ�ҽ�� = IIf(.������Ŀ��, 1, 0)
                    rsTmp!���ձ��� = .���ձ���
                    rsTmp!ժҪ = .ժҪ
                    rsTmp!�Ƿ��� = IIf(bln����, 1, 0)
                    rsTmp!��������ID = objBill.Pages(p).��������ID
                    rsTmp!ִ�в���ID = .ִ�в���ID
                    rsTmp.Update
                End With
            Next
        End If
    Next
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeBillRecord = rsTmp
    Exit Function
Errhand::
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlCreateFeeListStruc(ByRef rsFeelists As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������صķ��ü�¼���ṹ
    '���:
    '����:rsFeelists-���ر��ؼ�¼���ṹ,ͬʱ���˼�¼����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-05 16:18:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set rsFeelists = New ADODB.Recordset
    
    rsFeelists.Fields.Append "�������", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "�ѱ�", adVarChar, 50, adFldIsNullable
    rsFeelists.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsFeelists.Fields.Append "ʵ��Ʊ��", adVarChar, 20, adFldIsNullable
    rsFeelists.Fields.Append "����ʱ��", adDBTimeStamp, , adFldIsNullable
    rsFeelists.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsFeelists.Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
    rsFeelists.Fields.Append "���㵥λ", adVarChar, 50, adFldIsNullable
    '69788:���ϴ�,2014-6-5,�����������ֶδ�С����20��Ϊ100
    rsFeelists.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsFeelists.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "����", adDouble, , adFldIsNullable
    rsFeelists.Fields.Append "����", adDouble, , adFldIsNullable
    rsFeelists.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    rsFeelists.Fields.Append "ͳ����", adCurrency, , adFldIsNullable
    rsFeelists.Fields.Append "����֧������ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "�Ƿ�ҽ��", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "���ձ���", adVarChar, 50, adFldIsNullable
    rsFeelists.Fields.Append "ժҪ", adVarChar, 2000, adFldIsNullable
    rsFeelists.Fields.Append "�Ƿ���", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "��������ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "ִ�в���ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "���ν���", adDouble, , adFldIsNullable
    rsFeelists.CursorLocation = adUseClient
    rsFeelists.LockType = adLockOptimistic
    rsFeelists.CursorType = adOpenStatic
    rsFeelists.Open
    zlCreateFeeListStruc = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlBuldingFeeListdata(objBill As ExpenseBill, ByVal bln���� As Boolean, ByVal intPage As Integer, _
    ByVal strDate As String, ByVal str�ѱ� As String, ByVal strInvoice As String, ByRef rsFeelists As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݶ������ݴ���һ����¼��Ϣ(���ۼ۵�λ)
    '���:intPage=�൥���շ�ģʽʱ��ָ���ĵ���
    '     strDate=����ʱ��,
    '     strInvoice=Ʊ�ݺ�
    '����:rsFeeLists-���ط��ü�¼��( �������(�Ե���Ϊ׼),����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��,���ν���(����))
    '����:
    '����:���˺�
    '����:2010-01-05 16:11:44
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim dbl���� As Double, curʵ�� As Currency, curͳ�� As Currency
    Err = 0: On Error GoTo Errhand:
    For i = 1 To objBill.Pages(intPage).Details.Count
        dbl���� = 0: curʵ�� = 0: curͳ�� = 0
        With objBill.Pages(intPage).Details(i)
            For j = 1 To .InComes.Count
                dbl���� = dbl���� + .InComes(j).��׼����
                curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                curͳ�� = curͳ�� + .InComes(j).ͳ����
            Next
            rsFeelists.AddNew
            rsFeelists!������� = intPage
            rsFeelists!�ѱ� = str�ѱ�
            rsFeelists!NO = ""   '����ȡ���۵�ʱ����ֵ
            rsFeelists!ʵ��Ʊ�� = strInvoice
            rsFeelists!����ʱ�� = CDate(strDate)
            rsFeelists!����ID = IIf(objBill.����ID = 0, Null, objBill.����ID)
            rsFeelists!�շ���� = .�շ����
            
            If .InComes.Count > 0 Then
                rsFeelists!�վݷ�Ŀ = .InComes(1).�վݷ�Ŀ
            Else
                rsFeelists!�վݷ�Ŀ = Null
            End If
            rsFeelists!������ = objBill.Pages(intPage).������
            
            rsFeelists!�շ�ϸĿID = .�շ�ϸĿID
            
            rsFeelists!���㵥λ = .���㵥λ
            If InStr(",5,6,7,", .�շ����) > 0 And gblnҩ����λ Then
                '��ҩ����λת��Ϊ�ۼ۵�λ
                rsFeelists!���� = IIf(.���� = 0, 1, .����) * .���� * .Detail.ҩ����װ
                rsFeelists!���� = Format(dbl���� / .Detail.ҩ����װ, gstrFeePrecisionFmt)
            Else
                rsFeelists!���� = IIf(.���� = 0, 1, .����) * .����
                rsFeelists!���� = Format(dbl����, gstrFeePrecisionFmt)
            End If
            rsFeelists!ʵ�ս�� = Format(curʵ��, gstrDec)
            rsFeelists!ͳ���� = Format(curͳ��, gstrDec)
            rsFeelists!����֧������ID = IIf(.���մ���ID = 0, Null, .���մ���ID)
            rsFeelists!�Ƿ�ҽ�� = IIf(.������Ŀ��, 1, 0)
            rsFeelists!���ձ��� = .���ձ���
            rsFeelists!ժҪ = .ժҪ
            rsFeelists!�Ƿ��� = IIf(bln����, 1, 0)
            rsFeelists!��������ID = objBill.Pages(intPage).��������ID
            rsFeelists!ִ�в���ID = .ִ�в���ID
            rsFeelists!���ν��� = 0
            rsFeelists.Update
        End With
    Next
    If rsFeelists.RecordCount > 0 Then rsFeelists.MoveFirst
    zlBuldingFeeListdata = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetBill�ѱ�(objBill As ExpenseBill) As String
'���ܣ���������������зѱ�һ�£��򷵻ظ÷ѱ�,���򷵻ؿ�
    Dim i As Integer, p As Integer, strTmp As String
    
    For p = 1 To objBill.Pages.Count
        For i = 1 To objBill.Pages(p).Details.Count
            If i = 1 Then
                strTmp = objBill.Pages(p).Details(i).�ѱ�
            ElseIf objBill.Pages(p).Details(i).�ѱ� <> strTmp Then
                Exit Function
            End If
        Next
    Next
    GetBill�ѱ� = strTmp
End Function

Public Function GetDrugTotal(ByVal objBill As ExpenseBill, ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long, Optional ByVal intPage As Integer) As Double
'���ܣ���ȡ������ָ��ҩƷ��ͬһҩ�����е�������
'������ lngҩ��ID-0��ʾ���뷢ҩʱ,���޶�ҩ�����
    Dim i As Integer, p As Integer, dblCount As Double
    
    For p = 1 To objBill.Pages.Count
        If intPage = 0 Or p = intPage Then
            For i = 1 To objBill.Pages(p).Details.Count
                If objBill.Pages(p).Details(i).�շ�ϸĿID = lngҩƷID And _
                    IIf(lngҩ��ID <> 0, objBill.Pages(p).Details(i).ִ�в���ID = lngҩ��ID, 1 = 1) Then
                    dblCount = dblCount + objBill.Pages(p).Details(i).���� * objBill.Pages(p).Details(i).����
                End If
            Next
        End If
    Next
    GetDrugTotal = dblCount
End Function

Public Sub SaveRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub
Public Sub GetRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
End Sub

Public Function MoneyOverFlow(objBill As ExpenseBill) As Boolean
'���ܣ���鵥�ݺϼƽ���Ƿ����
'˵������Currency����922337203685477Ϊ׼
    Dim dblӦ�� As Double, dblʵ�� As Double
    Dim i As Integer, j As Integer, k As Integer
    
    'Ҫ��VALתΪDouble��������
    For i = 1 To objBill.Pages.Count
        For j = 1 To objBill.Pages(i).Details.Count
            For k = 1 To objBill.Pages(i).Details(j).InComes.Count
                If Abs(dblӦ�� + Val(objBill.Pages(i).Details(j).InComes(k).Ӧ�ս��)) > 922337203685477# Then
                    MoneyOverFlow = True: Exit Function
                End If
                If Abs(dblʵ�� + Val(objBill.Pages(i).Details(j).InComes(k).ʵ�ս��)) > 922337203685477# Then
                    MoneyOverFlow = True: Exit Function
                End If
                dblӦ�� = dblӦ�� + Val(objBill.Pages(i).Details(j).InComes(k).Ӧ�ս��)
                dblʵ�� = dblʵ�� + Val(objBill.Pages(i).Details(j).InComes(k).ʵ�ս��)
            Next
        Next
    Next
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function CollectionExitsValue(ByVal coll As Collection, _
    ByVal strKey As String) As Boolean
    '���ݹؼ����ж�Ԫ���Ƿ�����ڼ�����
    Dim blnExits As Boolean
    
    If coll Is Nothing Then Exit Function
    CollectionExitsValue = True
    Err = 0: On Error Resume Next
    blnExits = IsObject(coll(strKey))
    If Err <> 0 Then Err = 0: CollectionExitsValue = False
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'���ܣ��������뷽ʽ��ʽ������
'������intBit=���С��λ��
'����ţ�94552
'˵����VB�Դ���Round�����м����뷨,��ʵ�ʲ�һ�¡���Round(57.575,2)=57.58,Round(57.565,2)=57.56
    If intBit > 0 Then
        RoundEx = Val(Format(dblNumber, "0." & String(intBit, "0")))
    Else
        RoundEx = dblNumber
    End If
End Function
