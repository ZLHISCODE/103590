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
'��������

Private Enum PageSize
    �ż� = 1
    С���ż� = 2
    С�ͱ� = 3
    ������ = 4
    �����ļ� = 5
    ������ = 6
    �����ļ� = 7
    A3 = 8
    A4 = 9
    A4С�� = 10
    A5 = 11
    B4 = 12
    B5 = 13
    �Կ��� = 14
    �Ŀ��� = 15
    ���� = 18
    �ŷ�9�� = 19
    �ŷ�10�� = 20
    �ŷ�11�� = 21
    �ŷ�12�� = 22
    �ŷ�14�� = 23
    C�ߴ繤���� = 24
    D�ߴ繤���� = 25
    E�ߴ繤���� = 26
    DL���ŷ� = 27
    C5���ŷ� = 28
    C3���ŷ� = 29
    C4���ŷ� = 30
    C6���ŷ� = 31
    C65���ŷ� = 32
    B4���ŷ� = 33
    B5���ŷ� = 34
    B6���ŷ� = 35
    �ŷ�1 = 36
    �ŷ���� = 37
    �ŷ�2 = 38
    US��׼��д�� = 39
    �¹���׼��д�� = 40
    �¹����ɸ�д�� = 41
    �Զ��� = 256
End Enum

'��ӡֽ�ų���(256=�Զ���)
Private mrsPaper As ADODB.Recordset

Private Const HORZRES = 8                           'Horizontal Width in pixels
Private Const VERTRES = 10                          'Vertical Width in pixels
Private Const LOGPIXELSX = 88                       'Logical pixels/inch in X
Private Const LOGPIXELSY = 90                       'Logical pixels/inch in Y
Private Const PHYSICALOFFSETX = 112                 'Physical Printable Area x margin
Private Const PHYSICALOFFSETY = 113                 'Physical Printable Area y margin
Private Const PHYSICALHEIGHT = 111                  'Physical Height in device units
Private Const PHYSICALWidth = 110                   'Physical Width in device units
Private Const DC_PAPERNAMES = 16                    'ֽ������(ÿ64�ַ�Ϊһ��,��Chr(0)����)
Private Const DC_PAPERS = 2                         'ֽ�ű��(Array or Word)
Private Const DC_BINNAMES = 12                      '��ֽ��ʽ(ÿ24�ַ�Ϊһ��,��Chr(0)����)
Private Const DC_BINS = 6                           '��ֽ���(Array or Word)

'Window�汾����
Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'######################################################################################################################

Public Function GetPaperName(intSize As Long) As String
    '******************************************************************************************************************
    '���ܣ����ݵ�ǰ��ӡ�������ã���ȡֽ������
    '������
    '���أ�ֽ������
    '******************************************************************************************************************

    If intSize = 256 Then
        GetPaperName = "�û��Զ��� ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        
        mrsPaper.Filter = ""
        mrsPaper.Filter = "ֽ�����=" & intSize
        If mrsPaper.RecordCount > 0 Then
            GetPaperName = mrsPaper("ֽ������").Value
        Else
            GetPaperName = "���ɲ��ֽ�� ..."
        End If
        
    Else
        GetPaperName = "���ɲ��ֽ�� ..."
    End If
End Function

Public Function IsWindowsNT() As Boolean
    '���ܣ��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows95() As Boolean
    '���ܣ��Ƿ�Window95����ϵͳ
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

Private Function GetWinPlatform() As Long
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformID
End Function

Private Sub Class_Initialize()
    
    Set mrsPaper = New ADODB.Recordset
    With mrsPaper
        .Fields.Append "ֽ�����", adInteger
        .Fields.Append "ֽ������", adVarChar, 100
        .Open
    End With
    
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 1: mrsPaper("ֽ������").Value = "�ż㣬 8 1/2 x 11 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 2: mrsPaper("ֽ������").Value = "+A611 С���ż㣬 8 1/2 x 11 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 3: mrsPaper("ֽ������").Value = "С�ͱ��� 11 x 17 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 4: mrsPaper("ֽ������").Value = "�����ʣ� 17 x 11 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 5: mrsPaper("ֽ������").Value = "�����ļ��� 8 1/2 x 14 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 6: mrsPaper("ֽ������").Value = "�����飬5 1/2 x 8 1/2 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 7: mrsPaper("ֽ������").Value = "�����ļ���7 1/2 x 10 1/2 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 8: mrsPaper("ֽ������").Value = "A3, 297 x 420 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 9: mrsPaper("ֽ������").Value = "A4, 210 x 297 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 10: mrsPaper("ֽ������").Value = "A4С�ţ� 210 x 297 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 11: mrsPaper("ֽ������").Value = "A5, 148 x 210 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 12: mrsPaper("ֽ������").Value = "B4, 250 x 354 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 13: mrsPaper("ֽ������").Value = "B5, 182 x 257 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 14: mrsPaper("ֽ������").Value = "�Կ����� 8 1/2 x 13 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 15: mrsPaper("ֽ������").Value = "�Ŀ����� 215 x 275 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 16: mrsPaper("ֽ������").Value = "10 x 14 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 17: mrsPaper("ֽ������").Value = "11 x 17 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 18: mrsPaper("ֽ������").Value = "������8 1/2 x 11 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 19: mrsPaper("ֽ������").Value = "#9 �ŷ⣬ 3 7/8 x 8 7/8 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 20: mrsPaper("ֽ������").Value = "#10 �ŷ⣬ 4 1/8 x 9 1/2 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 21: mrsPaper("ֽ������").Value = "#11 �ŷ⣬ 4 1/2 x 10 3/8 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 22: mrsPaper("ֽ������").Value = "#12 �ŷ⣬ 4 1/2 x 11 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 23: mrsPaper("ֽ������").Value = "#14 �ŷ⣬ 5 x 11 1/2 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 24: mrsPaper("ֽ������").Value = "C �ߴ繤����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 25: mrsPaper("ֽ������").Value = "D �ߴ繤����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 26: mrsPaper("ֽ������").Value = "E �ߴ繤����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 27: mrsPaper("ֽ������").Value = "DL ���ŷ⣬ 110 x 220 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 28: mrsPaper("ֽ������").Value = "C5 ���ŷ⣬ 162 x 229 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 29: mrsPaper("ֽ������").Value = "C3 ���ŷ⣬ 324 x 458 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 30: mrsPaper("ֽ������").Value = "C4 ���ŷ⣬ 229 x 324 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 31: mrsPaper("ֽ������").Value = "C6 ���ŷ⣬ 114 x 162 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 32: mrsPaper("ֽ������").Value = "C65 ���ŷ⣬114 x 229 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 33: mrsPaper("ֽ������").Value = "B4 ���ŷ⣬ 250 x 353 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 34: mrsPaper("ֽ������").Value = "B5 ���ŷ⣬176 x 250 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 35: mrsPaper("ֽ������").Value = "B6 ���ŷ⣬ 176 x 125 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 36: mrsPaper("ֽ������").Value = "�ŷ⣬ 110 x 230 ����"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 37: mrsPaper("ֽ������").Value = "�ŷ������ 3 7/8 x 7 1/2 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 38: mrsPaper("ֽ������").Value = "�ŷ⣬ 3 5/8 x 6 1/2 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 39: mrsPaper("ֽ������").Value = "U.S. ��׼��д���� 14 7/8 x 11 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 40: mrsPaper("ֽ������").Value = "�¹���׼��д���� 8 1/2 x 12 Ӣ��"
    mrsPaper.AddNew: mrsPaper("ֽ�����").Value = 41: mrsPaper("ֽ������").Value = "�¹����ɸ�д���� 8 1/2 x 13 Ӣ��"

End Sub
