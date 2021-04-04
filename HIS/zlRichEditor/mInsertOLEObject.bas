Attribute VB_Name = "mInsertOLEObject"
'######################################################################################
'##模 块 名：mInsertOLEObject.bas
'##创 建 人：吴庆伟
'##日    期：2005年4月1日
'##修 改 人：
'##日    期：
'##描    述：在RTF中插入OLE对象的相关声明。
'##版    本：
'######################################################################################

Option Explicit

Public Declare Function OleUIInsertObject Lib "oledlg.dll" Alias "OleUIInsertObjectA" (inParam As Any) As Long
Public Declare Function ProgIDFromCLSID Lib "ole32.dll" (clsid As Any, strAddess As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pvoid As Long)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
 
Public Const IOF_SHOWHELP = &H1
Public Const IOF_SELECTCREATENEW = &H2
Public Const IOF_SELECTCREATEFROMFILE = &H4
Public Const IOF_CHECKLINK = &H8
Public Const IOF_CHECKDISPLAYASICON = &H10
Public Const IOF_CREATENEWOBJECT = &H20
Public Const IOF_CREATEFILEOBJECT = &H40
Public Const IOF_CREATELINKOBJECT = &H80
Public Const IOF_DISABLELINK = &H100
Public Const IOF_VERIFYSERVERSEXIST = &H200
Public Const IOF_DISABLEDISPLAYASICON = &H400
Public Const IOF_HIDECHANGEICON = &H800
Public Const IOF_SHOWINSERTCONTROL = &H1000
Public Const IOF_SELECTCREATECONTROL = &H2000

Public Const OLEUI_FALSE = 0
Public Const OLEUI_SUCCESS = 1                                 '  No  error,  same  as  OLEUI_OK.
Public Const OLEUI_OK = 1                                           '  OK  button  pressed.
Public Const OLEUI_CANCEL = 2                                   '  Cancel  button  pressed.

'  Main  UDT  used  in  OleUIInsertObject.
Public Type OleUIInsertObjectType
'  These  IN  fields  are  standard  across  all  OLEUI  dialog  box  functions.
        cbStruct  As Long
        dwFlags  As Long
        hWndOwner  As Long
        lpszCaption    As String         '  LPCSTR
        lpfnHook  As Long                     '  LPFNOLEUIHOOK
        lCustData  As Long                   '  LPARAM
        hInstance    As Long
        lpszTemplate  As String         '  LPCSTR
        hResource  As Long                   '  HRSRC
        clsid  As GUID

'  Specifics  for  OLEUIINSERTOBJECT.
        lpszFile  As String                 '  LPTSTR
        cchFile  As Long
        cClsidExclude  As Long
        lpClsidExclude  As Long         '  LPCLSID
        IID  As GUID

'  Specifics  to  create  objects  if  flags  say  so.
        oleRender  As Long
        lpFormatEtc  As Long               '  LPFORMATETC
        lpIOleClientSite  As Long     '  LPOLECLIENTSITE
        lpIStorage  As Long                 '  LPSTORAGE
        ppvObj  As Long                         '  LPVOID  FAR  *
        sc  As Long                                 '  SCODE
        hMetaPict  As Long                   '  HGLOBAL
End Type
 
