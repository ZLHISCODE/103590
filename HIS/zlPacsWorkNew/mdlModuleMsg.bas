Attribute VB_Name = "mdlModuleMsg"
Option Explicit

Public Enum TMsgModuleType
    mtImage = 0
    mtVideo
    mtPathol
End Enum


Public Const WM_XWREPORT_IMG As Long = 5120         '接收报告图像消息的API


'列表相关消息
Public Const WM_LIST_SYNCROW As Long = 5001         '同步列表选择行
Public Const WM_LIST_REFRESH As Long = 5002         '刷新数据列表
Public Const WM_LIST_MOVEUP As Long = 5003          '上移
Public Const WM_LIST_MOVEDOWN As Long = 5004        '下移
Public Const WM_LIST_GETLASTADVICE As Long = 5005     '获取上一条医嘱
Public Const WM_LIST_GETNEXTADVICE As Long = 5006   '获取下一条医嘱

Public Const WM_IMG_OPENVIEW As Long = 5101         '打开观片
Public Const WM_IMG_CONTRASTVIEW As Long = 5102         '对比观片

Public Const WM_REPORT_VIEW As Long = 5201          '报告预览
Public Const WM_REPORT_PRINT As Long = 5202          '报告打印


'Public Const WM_VIEW_REPORT As Long = 0             '预览报告
'Public Const WM_VIEW_IMAGE As Long = 0              '预览图像
'
'Public Const WM_EDITOR_LOCK As Long = 0             '锁定编辑
'Public Const WM_EDITOR_UNLOCK As Long = 0           '解锁编辑

'主菜单执行
Public Const BM_SYS__EVENT_MENU As Long = 1001

'RIS相关消息
Public Const BM_RIS_EVENT_REGISTER As Long = 4001         '检查登记
Public Const BM_RIS_EVENT_RECEVIE As Long = 4002          '检查报到
Public Const BM_RIS_EVENT_COMPLETE  As Long = 4003      '检查完成
Public Const BM_RIS_EVENT_CANCELREG As Long = 4004      '取消登记
Public Const BM_RIS_EVENT_CANCELREC As Long = 4005      '取消报到
Public Const BM_RIS_EVENT_CANCELCOMP As Long = 4006 '取消完成

'报告相关消息
Public Const BM_REPORT_EVENT_PRINT As Long = 6101           '报告打印事件(plugin...)
Public Const BM_REPORT_EVENT_SAVE As Long = 6102            '报告保存事件(plugin...)
Public Const BM_REPORT_EVENT_POPUPEXIT As Long = 6103   '弹出窗口触发的退出事件
Public Const BM_REPORT_EVENT_SIGN As Long = 6104            '报告签名事件(plugin...)
Public Const BM_REPORT_EVENT_AUDIT As Long = 6105           '预留，报告审核
Public Const BM_REPORT_EVENT_REJECT As Long = 6106         '报告驳回事件(plugin...)
Public Const BM_REPORT_EVENT_DELETE As Long = 6107          '报告删除事件(plugin...)
Public Const BM_REPORT_EVENT_BACK As Long = 6108            '报告回退事件(plugin...)
Public Const BM_REPORT_EVENT_Verify As Long = 6109          '报告验证事件(plugin...)
Public Const BM_REPORT_EVENT_REJHISTORY As Long = 6110      '驳回历史查看(plugin...)
Public Const BM_REPORT_EVENT_OPEN As Long = 6111            '报告打开事件
Public Const BM_REPORT_EVENT_IMGCHANGE As Long = 6112       '报告图改变事件
Public Const BM_REPORT_EVENT_QUALITY As Long = 6113         '报告质量标记事件
Public Const BM_REPORT_EVENT_ADDIMG As Long = 6114          '报告图添加事件
Public Const BM_REPORT_EVENT_CLOSEEPR As Long = 6115        '报告窗口关闭事件
Public Const BM_REPORT_EVENT_REFWCHR As Long = 6116         '刷新常用词句字符
Public Const BM_REPORT_EVENT_DELREPIMG As Long = 6117       '报告图删除事件
Public Const BM_REPORT_EVENT_REFFRAGMENT As Long = 6118     '刷新词句片段

'图像相关消息
Public Const BM_IMAGE_EVENT_DEL As Long = 6200              '删除图像
Public Const BM_IMAGE_EVENT_CAPTURE As Long = 6201    '采集图像
Public Const BM_IMAGE_EVENT_FIRST As Long = 6202   '采集首张图像

Public Const BM_IMAGE_EVENT_QUALITYTAG As Long = 6203   '标记影像质量
Public Const BM_IMAGE_EVENT_XWFILMPRINT   As Long = 6204  '胶片打印
Public Const BM_IMAGE_EVENT_GETIMAGE    As Long = 6205  '获取影像
Public Const BM_IMAGE_EVENT_TECHDO      As Long = 6206  '技师执行
Public Const BM_IMAGE_EVENT_CHANGEDEVICE As Long = 6207 '更换设备



'病理相关消息，兼容以前版本处理
Public Const BM_PATHOL_EVENT_BASE As Long = 7000

Private mlngImageProcHwnd As Long
Private mlngVideoProcHwnd As Long
Private mlngPatholProcHwnd As Long


Public gobjImageMainWindow As Object                  '用来接收报告图消息的窗体指针
Public gobjVideoMainWindow As Object
Public gobjPatholMainWindow As Object


Public Sub AttachModuleMsgProc(moduleType As TMsgModuleType, objMainWindow As Object)
    '指定自定义的窗口过程
    '返回并保存原来默认的窗口过程指针
    Dim lngOldProcHwnd As Long
On Error GoTo errhandle:
        
    If App.LogMode = 0 Then Exit Sub
    
    lngOldProcHwnd = SetWindowLong(objMainWindow.hwnd, GWL_WNDPROC, AddressOf MainWindowProc)
    
    Select Case moduleType
        Case mtImage
            mlngImageProcHwnd = lngOldProcHwnd
            Set gobjImageMainWindow = objMainWindow
        Case mtVideo
            mlngVideoProcHwnd = lngOldProcHwnd
            Set gobjVideoMainWindow = objMainWindow
        Case mtPathol
            mlngPatholProcHwnd = lngOldProcHwnd
            Set gobjPatholMainWindow = objMainWindow
    End Select
     
    Exit Sub
errhandle:
    
End Sub

Public Sub UnAttachModuleMsgProc(ByVal hwnd As Long, moduleType As TMsgModuleType)
On Error GoTo errhandle
    Dim temp As Long
    Dim lpWndProc As Long
    
    If hwnd = 0 Then Exit Sub
        
    Select Case moduleType
        Case mtImage
            lpWndProc = mlngImageProcHwnd
        Case mtVideo
            lpWndProc = mlngVideoProcHwnd
        Case mtPathol
            lpWndProc = mlngPatholProcHwnd
    End Select

    temp = SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc)
    
    Exit Sub
errhandle:

End Sub


Function MainWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'消息处理程序
Dim lngProcHwnd As Long
Dim objProc As Object

On Error GoTo errhandle
    lngProcHwnd = 0
    Set objProc = Nothing
    
    If Not gobjImageMainWindow Is Nothing Then
        If hw = gobjImageMainWindow.hwnd Then
            Set objProc = gobjImageMainWindow
            lngProcHwnd = mlngImageProcHwnd
        End If
    End If
    
    If Not gobjVideoMainWindow Is Nothing Then
        If hw = gobjVideoMainWindow.hwnd Then
            Set objProc = gobjVideoMainWindow
            lngProcHwnd = mlngVideoProcHwnd
        End If
    End If
    
    If Not gobjPatholMainWindow Is Nothing Then
        If hw = gobjPatholMainWindow.hwnd Then
            Set objProc = gobjPatholMainWindow
            lngProcHwnd = mlngPatholProcHwnd
        End If
    End If
    
    If Not objProc Is Nothing Then
        Call objProc.MainWindowProc(hw, uMsg, wParam, lParam)
    End If
 
    '调用原来的窗口过程
    MainWindowProc = CallWindowProc(lngProcHwnd, hw, uMsg, wParam, lParam)
Exit Function
errhandle:
    If lngProcHwnd <> 0 Then
        MainWindowProc = CallWindowProc(lngProcHwnd, hw, uMsg, wParam, lParam)
    End If
    
End Function

