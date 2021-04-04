Attribute VB_Name = "mEditor"
Option Explicit
'#########################################################################
'   ��������
'#########################################################################

Public Type PageInfo
    PageNumber As Long      'ҳ��
    Start As Long           '�ַ���ʼλ��
    End As Long             '�ַ���ֹλ��
    ActualHeight As Long    '��ҳʵ�ʴ�ӡ�߶�
End Type

'#########################################################################
'   ��������
'#########################################################################

Public AllPages() As PageInfo   'ҳ��Ϣ
Public PubInfo As New cEditor   '��ǰҳ����ͼ�Ĺ������ԡ�

Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

'############################################################################################################
'## ���ܣ�  �ı�RichTextBox������(��0��ʼ)ΪIndex��OLE�����dwFlags
'##
'## ������  TOM         ��TOM����
'##         NewFlag     ��REO_FLAGS����ʾOLE�������ʾ����
'##         Index       ��OLE�����˳��ֵ��-1��ʾ���ж���
'############################################################################################################
Public Function ChangeReObjectsFlag(ByVal TOM As cTextDocument, ByVal NewFlag As REO_FLAGS, Optional Index As Long = -1) As Boolean
    On Error GoTo LL
    Dim mIRichEditOle As IRichEditOle
    Dim mReObject As REOBJECT
    Dim mILockBytes As ILockBytes
    Dim lS As Long, lE As Long
    Dim OldCharRange As CHARRANGE
    Dim NewCharRange As CHARRANGE
    Dim objCount As Long
    Dim mIStorage As IStorage
    Dim mIOleClientSite As IOleClientSite
    Dim mIOleObject As olelib.IOleObject
    Dim mUUID As UUID
    
    '��ȡIRichEditOle�ӿ�
    SendMessage TOM.hwnd, EM_GETOLEINTERFACE, 0, mIRichEditOle
    If ObjPtr(mIRichEditOle) = 0 Then
        '��ȡIRichEditOle�ӿ�ʧ��
        'MsgBox "��ȡIRichEditOle�ӿ�ʧ��"
        Exit Function
    End If
   
    '���RichTextBox��OLE���������
    objCount = mIRichEditOle.GetObjectCount
    If objCount = 0 Then
        'RichTextBox��û�а���OLE����
        Set mIRichEditOle = Nothing
        Exit Function
    End If
    If Index <= -1 Then 'ȫ���ı�
        '��¼��RichTextBox��ǰѡ��������
'        SendMessage TOM.hwnd, EM_EXGETSEL, 0, OldCharRange
        lS = TOM.TextDocument.Selection.Start
        lE = TOM.TextDocument.Selection.End
        Dim i As Long
        For i = 0 To objCount - 1
            '���OLEObject����Ϣ
            mReObject.cbStruct = LenB(mReObject)    '���ýṹ��ߴ�
            mIRichEditOle.GetObject i, mReObject, REO_GETOBJ_ALL_INTERFACES     '��ȡ����i��OLE��������нӿ�
            Set mIOleObject = mReObject.poleobj
            With NewCharRange
                .cpMin = mReObject.cP
                .cpMax = mReObject.cP
            End With
            'ɾ����ǰ��oleobject
            'ֻ���Բ���selstart֮������Կ��ƣ�����ΪReObject.cp�ǻ����ֽڵ�
            'PutFocus TOM.hwnd
            'ѡ�и÷�Χ
            SendMessage TOM.hwnd, EM_EXSETSEL, 0, NewCharRange
            TOM.TextDocument.Selection.Delete tomCharacter, 1
            'SendKeys "{DEL}", True
           
            '�ı�dwflags�����²���oleobject
            Set mILockBytes = CreateILockBytesOnHGlobal(0&, True)
            If ObjPtr(mILockBytes) = 0 Then
                'MsgBox "����ȫ�ֶѳ���"
                Exit Function
            End If
            '����Storage��ʵ����mIStorage
            Set mIStorage = StgCreateDocfileOnILockBytes(mILockBytes, STGM_SHARE_EXCLUSIVE _
                            Or STGM_CREATE Or STGM_READWRITE, 0)
            If ObjPtr(mIStorage) = 0 Then
                'MsgBox "�����洢�������"
                Exit Function
            End If
           
            '����GetClientSite������ʵ����mIOleClientSite
            Set mIOleClientSite = mIRichEditOle.GetClientSite
            If ObjPtr(mIOleClientSite) = 0 Then
                'MsgBox "��ȡ�ͻ�������"
                Exit Function
            End If
            '֪ͨһ��OLE����Ƕ�뵽�����У�ȡ����ȷ���á�
            OleSetContainedObject mIOleObject, True
            mIOleObject.GetUserClassID mUUID
            With mReObject
                .cbStruct = LenB(mReObject)
                .clsid = mUUID
                .cP = REO_CP_SELECTION
                .dwFlags = NewFlag              '�����µ�״̬��־
                Set .poleobj = mIOleObject
                Set .polesite = mIOleClientSite
                Set .pStg = mIStorage
            End With
            '�ָ�OLE����
            mIRichEditOle.InsertObject mReObject
        Next
        '�ָ�RichTextBoxԭ��ѡ��������
        'SendMessage TOM.hwnd, EM_EXSETSEL, 0, OldCharRange
        TOM.TextDocument.Range(lS, lE).Select
    Else
        If Index > objCount - 1 Then
            'MsgBox "��Ч������������Index����ֵ(Index=0,1,2,...)"
            Set mIRichEditOle = Nothing
            Exit Function
        Else
            '��¼��RichTextBox��ǰѡ��������
            'SendMessage TOM.hwnd, EM_EXGETSEL, 0, OldCharRange
            lS = TOM.TextDocument.Selection.Start
            lE = TOM.TextDocument.Selection.End
            '���oleobject����Ϣ
            mReObject.cbStruct = LenB(mReObject)
            mIRichEditOle.GetObject Index, mReObject, REO_GETOBJ_ALL_INTERFACES
            Set mIOleObject = mReObject.poleobj
            
            With NewCharRange
                .cpMin = mReObject.cP
                .cpMax = mReObject.cP
            End With
            'ɾ����ǰ��oleobject
            'ֻ���Բ���selstart֮������Կ��ƣ�����ΪReObject.cp�ǻ����ֽڵ�
            'PutFocus TOM.hwnd
            SendMessage TOM.hwnd, EM_EXSETSEL, 0, NewCharRange
            TOM.TextDocument.Selection.Delete tomCharacter, 1
            'SendKeys "{DEL}", True
            
            
            '�ı�dwflags�����²���oleobject
            Set mILockBytes = CreateILockBytesOnHGlobal(0&, True)
            If ObjPtr(mILockBytes) = 0 Then
                'MsgBox "����ȫ�ֶѳ���"
                Exit Function
            End If
            '����storage��ʵ����mIStorage
            Set mIStorage = StgCreateDocfileOnILockBytes(mILockBytes, STGM_SHARE_EXCLUSIVE _
                            Or STGM_CREATE Or STGM_READWRITE, 0)
            If ObjPtr(mIStorage) = 0 Then
                'MsgBox "�����洢�������"
                Exit Function
            End If
           
            '����GetClientSite������ʵ����mIOleClientSite
            Set mIOleClientSite = mIRichEditOle.GetClientSite
            If ObjPtr(mIOleClientSite) = 0 Then
                'MsgBox "��ȡ�ͻ�������"
                Exit Function
            End If
            '֪ͨһ��OLE����Ƕ�뵽�����У�ȡ����ȷ���á�
            OleSetContainedObject mIOleObject, True
            mIOleObject.GetUserClassID mUUID
            With mReObject
                .cbStruct = LenB(mReObject)
                .clsid = mUUID
                .cP = REO_CP_SELECTION
                .dwFlags = NewFlag              '�����µ�״̬��־
                Set .poleobj = mIOleObject
                Set .polesite = mIOleClientSite
                Set .pStg = mIStorage
            End With
            mIRichEditOle.InsertObject mReObject
            '�ָ�RichTextBoxԭ��ѡ��������
            'SendMessage TOM.hwnd, EM_EXSETSEL, 0, OldCharRange
            TOM.TextDocument.Range(lS, lE).Select
         End If
     End If
    '�ͷ���Դ
    Set mIRichEditOle = Nothing
    Set mILockBytes = Nothing
    Set mIStorage = Nothing
    Set mIOleClientSite = Nothing
    Set mIOleObject = Nothing
    ChangeReObjectsFlag = True
    Exit Function
LL:
    ChangeReObjectsFlag = False
End Function

'############################################################################################################
'## ���ܣ�  �ı�RichTextBox������(��0��ʼ)ΪIndex��ͼƬ�ߴ�
'##
'## ������  hWnd        ��RTB����ľ��
'############################################################################################################
Public Function ResizeReObject(ByVal rtbThis As RichTextBox, _
    ByVal lngWidth As Long, ByVal lngHeight As Long) As Boolean
    
    On Error GoTo LL
    Dim mIRichEditOle As IRichEditOle
    Dim mReObject As REOBJECT
    Dim mILockBytes As ILockBytes
    Dim OldCharRange As CHARRANGE
    Dim NewCharRange As CHARRANGE
    Dim objCount As Long
    Dim mIStorage As IStorage
    Dim mIOleClientSite As IOleClientSite
    Dim mIOleObject As olelib.IOleObject
    Dim mUUID As UUID
    
    Dim Index As Long
    Index = 0
    
    '��ȡIRichEditOle�ӿ�
    SendMessage rtbThis.hwnd, EM_GETOLEINTERFACE, 0, mIRichEditOle
    If ObjPtr(mIRichEditOle) = 0 Then Exit Function '��ȡIRichEditOle�ӿ�ʧ��
   
    '���RichTextBox��OLE���������
    objCount = mIRichEditOle.GetObjectCount
    If objCount = 0 Or Index > objCount - 1 Then
        'RichTextBox��û�а���OLE����
        Set mIRichEditOle = Nothing
        Exit Function
    End If

     '���oleobject����Ϣ
     mReObject.cbStruct = LenB(mReObject)
     mIRichEditOle.GetObject Index, mReObject, REO_GETOBJ_ALL_INTERFACES
     Set mIOleObject = mReObject.poleobj
     
     With NewCharRange
         .cpMin = mReObject.cP
         .cpMax = mReObject.cP
     End With
     'ɾ����ǰ��oleobject
     rtbThis.Text = ""
     
     '���²���oleobject
     Set mILockBytes = CreateILockBytesOnHGlobal(0&, True)
     If ObjPtr(mILockBytes) = 0 Then Exit Function  '����ȫ�ֶѳ���
     
     '����storage��ʵ����mIStorage
     Set mIStorage = StgCreateDocfileOnILockBytes(mILockBytes, STGM_SHARE_EXCLUSIVE _
                     Or STGM_CREATE Or STGM_READWRITE, 0)
     If ObjPtr(mIStorage) = 0 Then Exit Function    '�����洢�������
    
     '����GetClientSite������ʵ����mIOleClientSite
     Set mIOleClientSite = mIRichEditOle.GetClientSite
     If ObjPtr(mIOleClientSite) = 0 Then Exit Function  '��ȡ�ͻ�������
     
     '֪ͨһ��OLE����Ƕ�뵽�����У�ȡ����ȷ���á�
     OleSetContainedObject mIOleObject, True
     mIOleObject.GetUserClassID mUUID
     With mReObject
         .cbStruct = LenB(mReObject)
         .clsid = mUUID
         .cP = REO_CP_SELECTION
         Set .poleobj = mIOleObject
         Set .polesite = mIOleClientSite
         Set .pStg = mIStorage
         .sizel.cx = lngWidth * 26.4541015625 / 15#
         .sizel.cy = lngHeight * 26.4544270833333 / 15#
     End With
     mIRichEditOle.InsertObject mReObject

    '�ͷ���Դ
    Set mIRichEditOle = Nothing
    Set mILockBytes = Nothing
    Set mIStorage = Nothing
    Set mIOleClientSite = Nothing
    Set mIOleObject = Nothing
    ResizeReObject = True
    Exit Function
LL:
    ResizeReObject = False
End Function

