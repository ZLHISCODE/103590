VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'不用菜单设计器，用API自已创建简单的弹出菜单。
'在工程中新增一个类，名称为cPopupMenu

Option Explicit
'
Private Type POINT
    X As Long
    Y As Long
End Type
'
Private Const MF_ENABLED = &H0&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_NONOTIFY = &H80&
Private Const TPM_RETURNCMD = &H100&

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal sCaption As String) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, nIgnored As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long

Private mSelMenuString As String

Public Property Get SelMenuString() As String
    SelMenuString = mSelMenuString
End Property

Public Function Popup(ParamArray param()) As Long
    Dim iMenu As Long
    Dim hMenu As Long
    Dim nMenus As Long
    Dim p As POINT
    Dim result As Long
    Dim buffer As String
    
    Const MF_BYPOSITION = &H400&
    

    GetCursorPos p


    hMenu = CreatePopupMenu()

    ' determine # of strings in paramarray
    nMenus = 1 + UBound(param)

    For iMenu = 1 To nMenus
        If Trim(CStr(param(iMenu - 1))) = "-" Then
            lngmenu = AppendMenu(hMenu, MF_SEPARATOR, iMenu, "")
        Else
            lngmenu = AppendMenu(hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(param(iMenu - 1)))
            
            ModifyMenu lngmenu, 0, &H0&, 2, ""
        End If
        
        
    Next
    
    '在当前光标位置显示弹出菜单,返回位置
    iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, GetForegroundWindow(), 0)

    buffer = Space(255)
   
    result = GetMenuString(hMenu, (iMenu - 1), buffer, Len(buffer), MF_BYPOSITION)
    
    mSelMenuString = Trim(buffer)
    
    DestroyMenu hMenu
    Popup = iMenu
    
End Function



