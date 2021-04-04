Attribute VB_Name = "mEditor"
Option Explicit
'#########################################################################
'   公共类型
'#########################################################################

Public Type PageInfo
    PageNumber As Long      '页码
    Start As Long           '字符起始位置
    End As Long             '字符终止位置
    ActualHeight As Long    '本页实际打印高度
End Type

'#########################################################################
'   公共变量
'#########################################################################

Public AllPages() As PageInfo   '页信息
Public PubInfo As New cEditor   '当前页面视图的公共属性。

Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

'############################################################################################################
'## 功能：  改变RichTextBox中索引(从0开始)为Index的OLE对象的dwFlags
'##
'## 参数：  TOM         ：TOM对象
'##         NewFlag     ：REO_FLAGS，表示OLE对象的显示特性
'##         Index       ：OLE对象的顺序值，-1表示所有对象
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
    
    '获取IRichEditOle接口
    SendMessage TOM.hwnd, EM_GETOLEINTERFACE, 0, mIRichEditOle
    If ObjPtr(mIRichEditOle) = 0 Then
        '获取IRichEditOle接口失败
        'MsgBox "获取IRichEditOle接口失败"
        Exit Function
    End If
   
    '获得RichTextBox中OLE对象的数量
    objCount = mIRichEditOle.GetObjectCount
    If objCount = 0 Then
        'RichTextBox中没有包含OLE对象
        Set mIRichEditOle = Nothing
        Exit Function
    End If
    If Index <= -1 Then '全部改变
        '记录下RichTextBox当前选定的内容
'        SendMessage TOM.hwnd, EM_EXGETSEL, 0, OldCharRange
        lS = TOM.TextDocument.Selection.Start
        lE = TOM.TextDocument.Selection.End
        Dim i As Long
        For i = 0 To objCount - 1
            '获得OLEObject的信息
            mReObject.cbStruct = LenB(mReObject)    '设置结构体尺寸
            mIRichEditOle.GetObject i, mReObject, REO_GETOBJ_ALL_INTERFACES     '获取索引i的OLE对象的所有接口
            Set mIOleObject = mReObject.poleobj
            With NewCharRange
                .cpMin = mReObject.cP
                .cpMax = mReObject.cP
            End With
            '删除当前的oleobject
            '只所以不用selstart之类的属性控制，是因为ReObject.cp是基于字节的
            'PutFocus TOM.hwnd
            '选中该范围
            SendMessage TOM.hwnd, EM_EXSETSEL, 0, NewCharRange
            TOM.TextDocument.Selection.Delete tomCharacter, 1
            'SendKeys "{DEL}", True
           
            '改变dwflags后重新插入oleobject
            Set mILockBytes = CreateILockBytesOnHGlobal(0&, True)
            If ObjPtr(mILockBytes) = 0 Then
                'MsgBox "创建全局堆出错"
                Exit Function
            End If
            '创建Storage，实例化mIStorage
            Set mIStorage = StgCreateDocfileOnILockBytes(mILockBytes, STGM_SHARE_EXCLUSIVE _
                            Or STGM_CREATE Or STGM_READWRITE, 0)
            If ObjPtr(mIStorage) = 0 Then
                'MsgBox "创建存储对象出错"
                Exit Function
            End If
           
            '调用GetClientSite函数，实例化mIOleClientSite
            Set mIOleClientSite = mIRichEditOle.GetClientSite
            If ObjPtr(mIOleClientSite) = 0 Then
                'MsgBox "获取客户区出错"
                Exit Function
            End If
            '通知一个OLE对象嵌入到容器中，取保正确引用。
            OleSetContainedObject mIOleObject, True
            mIOleObject.GetUserClassID mUUID
            With mReObject
                .cbStruct = LenB(mReObject)
                .clsid = mUUID
                .cP = REO_CP_SELECTION
                .dwFlags = NewFlag              '设置新的状态标志
                Set .poleobj = mIOleObject
                Set .polesite = mIOleClientSite
                Set .pStg = mIStorage
            End With
            '恢复OLE对象
            mIRichEditOle.InsertObject mReObject
        Next
        '恢复RichTextBox原来选定的内容
        'SendMessage TOM.hwnd, EM_EXSETSEL, 0, OldCharRange
        TOM.TextDocument.Range(lS, lE).Select
    Else
        If Index > objCount - 1 Then
            'MsgBox "无效的索引，请检查Index属性值(Index=0,1,2,...)"
            Set mIRichEditOle = Nothing
            Exit Function
        Else
            '记录下RichTextBox当前选定的内容
            'SendMessage TOM.hwnd, EM_EXGETSEL, 0, OldCharRange
            lS = TOM.TextDocument.Selection.Start
            lE = TOM.TextDocument.Selection.End
            '获得oleobject的信息
            mReObject.cbStruct = LenB(mReObject)
            mIRichEditOle.GetObject Index, mReObject, REO_GETOBJ_ALL_INTERFACES
            Set mIOleObject = mReObject.poleobj
            
            With NewCharRange
                .cpMin = mReObject.cP
                .cpMax = mReObject.cP
            End With
            '删除当前的oleobject
            '只所以不用selstart之类的属性控制，是因为ReObject.cp是基于字节的
            'PutFocus TOM.hwnd
            SendMessage TOM.hwnd, EM_EXSETSEL, 0, NewCharRange
            TOM.TextDocument.Selection.Delete tomCharacter, 1
            'SendKeys "{DEL}", True
            
            
            '改变dwflags后重新插入oleobject
            Set mILockBytes = CreateILockBytesOnHGlobal(0&, True)
            If ObjPtr(mILockBytes) = 0 Then
                'MsgBox "创建全局堆出错"
                Exit Function
            End If
            '创建storage，实例化mIStorage
            Set mIStorage = StgCreateDocfileOnILockBytes(mILockBytes, STGM_SHARE_EXCLUSIVE _
                            Or STGM_CREATE Or STGM_READWRITE, 0)
            If ObjPtr(mIStorage) = 0 Then
                'MsgBox "创建存储对象出错"
                Exit Function
            End If
           
            '调用GetClientSite函数，实例化mIOleClientSite
            Set mIOleClientSite = mIRichEditOle.GetClientSite
            If ObjPtr(mIOleClientSite) = 0 Then
                'MsgBox "获取客户区出错"
                Exit Function
            End If
            '通知一个OLE对象嵌入到容器中，取保正确引用。
            OleSetContainedObject mIOleObject, True
            mIOleObject.GetUserClassID mUUID
            With mReObject
                .cbStruct = LenB(mReObject)
                .clsid = mUUID
                .cP = REO_CP_SELECTION
                .dwFlags = NewFlag              '设置新的状态标志
                Set .poleobj = mIOleObject
                Set .polesite = mIOleClientSite
                Set .pStg = mIStorage
            End With
            mIRichEditOle.InsertObject mReObject
            '恢复RichTextBox原来选定的内容
            'SendMessage TOM.hwnd, EM_EXSETSEL, 0, OldCharRange
            TOM.TextDocument.Range(lS, lE).Select
         End If
     End If
    '释放资源
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
'## 功能：  改变RichTextBox中索引(从0开始)为Index的图片尺寸
'##
'## 参数：  hWnd        ：RTB对象的句柄
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
    
    '获取IRichEditOle接口
    SendMessage rtbThis.hwnd, EM_GETOLEINTERFACE, 0, mIRichEditOle
    If ObjPtr(mIRichEditOle) = 0 Then Exit Function '获取IRichEditOle接口失败
   
    '获得RichTextBox中OLE对象的数量
    objCount = mIRichEditOle.GetObjectCount
    If objCount = 0 Or Index > objCount - 1 Then
        'RichTextBox中没有包含OLE对象
        Set mIRichEditOle = Nothing
        Exit Function
    End If

     '获得oleobject的信息
     mReObject.cbStruct = LenB(mReObject)
     mIRichEditOle.GetObject Index, mReObject, REO_GETOBJ_ALL_INTERFACES
     Set mIOleObject = mReObject.poleobj
     
     With NewCharRange
         .cpMin = mReObject.cP
         .cpMax = mReObject.cP
     End With
     '删除当前的oleobject
     rtbThis.Text = ""
     
     '重新插入oleobject
     Set mILockBytes = CreateILockBytesOnHGlobal(0&, True)
     If ObjPtr(mILockBytes) = 0 Then Exit Function  '创建全局堆出错
     
     '创建storage，实例化mIStorage
     Set mIStorage = StgCreateDocfileOnILockBytes(mILockBytes, STGM_SHARE_EXCLUSIVE _
                     Or STGM_CREATE Or STGM_READWRITE, 0)
     If ObjPtr(mIStorage) = 0 Then Exit Function    '创建存储对象出错
    
     '调用GetClientSite函数，实例化mIOleClientSite
     Set mIOleClientSite = mIRichEditOle.GetClientSite
     If ObjPtr(mIOleClientSite) = 0 Then Exit Function  '获取客户区出错
     
     '通知一个OLE对象嵌入到容器中，取保正确引用。
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

    '释放资源
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

