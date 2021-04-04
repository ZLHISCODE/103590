VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'######################################################################################################################
'常量定义

Private Enum PageSize
    信笺 = 1
    小型信笺 = 2
    小型报 = 3
    分类帐 = 4
    法律文件 = 5
    声明书 = 6
    行政文件 = 7
    A3 = 8
    A4 = 9
    A4小号 = 10
    A5 = 11
    B4 = 12
    B5 = 13
    对开本 = 14
    四开本 = 15
    便条 = 18
    信封9号 = 19
    信封10号 = 20
    信封11号 = 21
    信封12号 = 22
    信封14号 = 23
    C尺寸工作单 = 24
    D尺寸工作单 = 25
    E尺寸工作单 = 26
    DL型信封 = 27
    C5型信封 = 28
    C3型信封 = 29
    C4型信封 = 30
    C6型信封 = 31
    C65型信封 = 32
    B4型信封 = 33
    B5型信封 = 34
    B6型信封 = 35
    信封1 = 36
    信封大王 = 37
    信封2 = 38
    US标准复写簿 = 39
    德国标准复写簿 = 40
    德国法律复写簿 = 41
    自定义 = 256
End Enum

'打印纸张常量(256=自定义)
Private mrsPaper As ADODB.Recordset

Private Const HORZRES = 8                           'Horizontal width in pixels
Private Const VERTRES = 10                          'Vertical width in pixels
Private Const LOGPIXELSX = 88                       'Logical pixels/inch in X
Private Const LOGPIXELSY = 90                       'Logical pixels/inch in Y
Private Const PHYSICALOFFSETX = 112                 'Physical Printable Area x margin
Private Const PHYSICALOFFSETY = 113                 'Physical Printable Area y margin
Private Const PHYSICALHEIGHT = 111                  'Physical Height in device units
Private Const PHYSICALWIDTH = 110                   'Physical Width in device units
Private Const DC_PAPERNAMES = 16                    '纸张名称(每64字符为一段,以Chr(0)结束)
Private Const DC_PAPERS = 2                         '纸张编号(Array or Word)
Private Const DC_BINNAMES = 12                      '进纸方式(每24字符为一段,以Chr(0)结束)
Private Const DC_BINS = 6                           '进纸编号(Array or Word)

'Window版本函数
Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

'######################################################################################################################

Public Function GetPaperName(intSize As Long) As String
    '******************************************************************************************************************
    '功能：根据当前打印机的设置，获取纸张名称
    '参数：
    '返回：纸张名称
    '******************************************************************************************************************

    If intSize = 256 Then
        GetPaperName = "用户自定义 ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        
        mrsPaper.Filter = ""
        mrsPaper.Filter = "纸张序号=" & intSize
        If mrsPaper.RecordCount > 0 Then
            GetPaperName = mrsPaper("纸张名称").Value
        Else
            GetPaperName = "不可测的纸张 ..."
        End If
        
    Else
        GetPaperName = "不可测的纸张 ..."
    End If
End Function

Public Function IsWindowsNT() As Boolean
    '功能：是否WindowNT操作系统
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows95() As Boolean
    '功能：是否Window95操作系统
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

Private Function GetWinPlatform() As Long
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

Private Sub Class_Initialize()
    
    Set mrsPaper = New ADODB.Recordset
    With mrsPaper
        .Fields.Append "纸张序号", adInteger
        .Fields.Append "纸张名称", adVarChar, 100
        .Open
    End With
    
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 1: mrsPaper("纸张名称").Value = "信笺， 8 1/2 x 11 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 2: mrsPaper("纸张名称").Value = "+A611 小型信笺， 8 1/2 x 11 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 3: mrsPaper("纸张名称").Value = "小型报， 11 x 17 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 4: mrsPaper("纸张名称").Value = "分类帐， 17 x 11 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 5: mrsPaper("纸张名称").Value = "法律文件， 8 1/2 x 14 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 6: mrsPaper("纸张名称").Value = "声明书，5 1/2 x 8 1/2 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 7: mrsPaper("纸张名称").Value = "行政文件，7 1/2 x 10 1/2 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 8: mrsPaper("纸张名称").Value = "A3, 297 x 420 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 9: mrsPaper("纸张名称").Value = "A4, 210 x 297 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 10: mrsPaper("纸张名称").Value = "A4小号， 210 x 297 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 11: mrsPaper("纸张名称").Value = "A5, 148 x 210 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 12: mrsPaper("纸张名称").Value = "B4, 250 x 354 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 13: mrsPaper("纸张名称").Value = "B5, 182 x 257 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 14: mrsPaper("纸张名称").Value = "对开本， 8 1/2 x 13 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 15: mrsPaper("纸张名称").Value = "四开本， 215 x 275 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 16: mrsPaper("纸张名称").Value = "10 x 14 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 17: mrsPaper("纸张名称").Value = "11 x 17 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 18: mrsPaper("纸张名称").Value = "便条，8 1/2 x 11 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 19: mrsPaper("纸张名称").Value = "#9 信封， 3 7/8 x 8 7/8 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 20: mrsPaper("纸张名称").Value = "#10 信封， 4 1/8 x 9 1/2 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 21: mrsPaper("纸张名称").Value = "#11 信封， 4 1/2 x 10 3/8 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 22: mrsPaper("纸张名称").Value = "#12 信封， 4 1/2 x 11 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 23: mrsPaper("纸张名称").Value = "#14 信封， 5 x 11 1/2 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 24: mrsPaper("纸张名称").Value = "C 尺寸工作单"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 25: mrsPaper("纸张名称").Value = "D 尺寸工作单"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 26: mrsPaper("纸张名称").Value = "E 尺寸工作单"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 27: mrsPaper("纸张名称").Value = "DL 型信封， 110 x 220 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 28: mrsPaper("纸张名称").Value = "C5 型信封， 162 x 229 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 29: mrsPaper("纸张名称").Value = "C3 型信封， 324 x 458 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 30: mrsPaper("纸张名称").Value = "C4 型信封， 229 x 324 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 31: mrsPaper("纸张名称").Value = "C6 型信封， 114 x 162 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 32: mrsPaper("纸张名称").Value = "C65 型信封，114 x 229 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 33: mrsPaper("纸张名称").Value = "B4 型信封， 250 x 353 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 34: mrsPaper("纸张名称").Value = "B5 型信封，176 x 250 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 35: mrsPaper("纸张名称").Value = "B6 型信封， 176 x 125 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 36: mrsPaper("纸张名称").Value = "信封， 110 x 230 毫米"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 37: mrsPaper("纸张名称").Value = "信封大王， 3 7/8 x 7 1/2 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 38: mrsPaper("纸张名称").Value = "信封， 3 5/8 x 6 1/2 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 39: mrsPaper("纸张名称").Value = "U.S. 标准复写簿， 14 7/8 x 11 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 40: mrsPaper("纸张名称").Value = "德国标准复写簿， 8 1/2 x 12 英寸"
    mrsPaper.AddNew: mrsPaper("纸张序号").Value = 41: mrsPaper("纸张名称").Value = "德国法律复写簿， 8 1/2 x 13 英寸"

End Sub
