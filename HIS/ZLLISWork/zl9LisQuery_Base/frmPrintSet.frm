VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ����"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   Icon            =   "frmPrintSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cbo��ӡ�� 
      Height          =   300
      Left            =   1275
      TabIndex        =   31
      Text            =   "cbo��ӡ��"
      Top             =   2865
      Width           =   3660
   End
   Begin VB.TextBox txtPageSize 
      Height          =   300
      Index           =   1
      Left            =   3525
      TabIndex        =   29
      Top             =   3705
      Width           =   900
   End
   Begin VB.TextBox txtPageSize 
      Height          =   300
      Index           =   0
      Left            =   900
      TabIndex        =   27
      Top             =   3705
      Width           =   900
   End
   Begin VB.Frame Frame4 
      Height          =   135
      Left            =   135
      TabIndex        =   26
      Top             =   2655
      Width           =   4950
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   25
      Top             =   1275
      Width           =   4950
   End
   Begin VB.ComboBox cboPageSize 
      Height          =   300
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3210
      Width           =   3660
   End
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   1065
      Left            =   345
      TabIndex        =   20
      Top             =   1515
      Width           =   1320
      Begin VB.OptionButton opt���� 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   315
         TabIndex        =   22
         Top             =   690
         Width           =   885
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   315
         TabIndex        =   21
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ҳ�߾�(����)"
      Height          =   1065
      Left            =   1875
      TabIndex        =   11
      Top             =   1515
      Width           =   2880
      Begin VB.TextBox txt�߾� 
         Height          =   300
         Index           =   3
         Left            =   2040
         TabIndex        =   18
         Top             =   615
         Width           =   600
      End
      Begin VB.TextBox txt�߾� 
         Height          =   300
         Index           =   2
         Left            =   735
         TabIndex        =   16
         Top             =   615
         Width           =   600
      End
      Begin VB.TextBox txt�߾� 
         Height          =   300
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Top             =   270
         Width           =   600
      End
      Begin VB.TextBox txt�߾� 
         Height          =   300
         Index           =   0
         Left            =   735
         TabIndex        =   12
         Top             =   270
         Width           =   600
      End
      Begin VB.Label lbl�߾� 
         Caption         =   "��(R)"
         Height          =   195
         Index           =   3
         Left            =   1515
         TabIndex        =   19
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl�߾� 
         Caption         =   "��(L)"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   17
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl�߾� 
         Caption         =   "��(B)"
         Height          =   195
         Index           =   1
         Left            =   1515
         TabIndex        =   15
         Top             =   315
         Width           =   495
      End
      Begin VB.Label lbl�߾� 
         Caption         =   "��(T)"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.OptionButton optҳ�� 
      Caption         =   "��"
      Height          =   225
      Index           =   2
      Left            =   3570
      TabIndex        =   10
      Top             =   1020
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optҳ�� 
      Caption         =   "��"
      Height          =   225
      Index           =   1
      Left            =   2925
      TabIndex        =   9
      Top             =   1020
      Width           =   615
   End
   Begin VB.OptionButton optҳ�� 
      Caption         =   "��"
      Height          =   225
      Index           =   0
      Left            =   2280
      TabIndex        =   8
      Top             =   1020
      Width           =   615
   End
   Begin VB.CheckBox chkҳ�� 
      Caption         =   "��ӡҳ��"
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   1020
      Width           =   1125
   End
   Begin VB.TextBox txt����Font 
      Enabled         =   0   'False
      Height          =   300
      Left            =   945
      TabIndex        =   5
      Top             =   555
      Width           =   3630
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "��"
      Height          =   300
      Left            =   4605
      TabIndex        =   4
      Top             =   555
      Width           =   300
   End
   Begin MSComDlg.CommonDialog cmdigFont 
      Left            =   5010
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   945
      TabIndex        =   2
      Top             =   210
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3795
      TabIndex        =   1
      Top             =   4155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2550
      TabIndex        =   0
      Top             =   4155
      Width           =   1100
   End
   Begin VB.Label Label3 
      Caption         =   "��ӡ��(D)"
      Height          =   195
      Left            =   195
      TabIndex        =   32
      Top             =   2895
      Width           =   1050
   End
   Begin VB.Label lblPageSize 
      Caption         =   "�߶�(H)            ����"
      Height          =   255
      Index           =   1
      Left            =   2820
      TabIndex        =   30
      Top             =   3750
      Width           =   2070
   End
   Begin VB.Label lblPageSize 
      Caption         =   "���(W)            ����"
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   28
      Top             =   3750
      Width           =   2070
   End
   Begin VB.Label Label2 
      Caption         =   "ֽ�Ŵ�С(Z)"
      Height          =   195
      Left            =   195
      TabIndex        =   24
      Top             =   3255
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   270
      Left            =   165
      TabIndex        =   6
      Top             =   615
      Width           =   870
   End
   Begin VB.Label lbl���� 
      Caption         =   "��ӡ����"
      Height          =   270
      Left            =   165
      TabIndex        =   3
      Top             =   255
      Width           =   750
   End
End
Attribute VB_Name = "frmPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngIndex As Long
Const DC_PAPERS = 2          '��ʾҪ��ȡ��ӡֽ�Ŵ�С
Const DC_PAPERNAMES = 16     '��ʾҪ��ȡ��ӡֽ������

'��ӡ�����豸�ṹ
Private Type DEVMODE
    dmdevicename As String * 64
    dmspecversion As Integer
    dmdriverversion As Integer
    dmsize As Integer
    dmdriverextra As Integer
    dmfields As Long
End Type

'API����������?
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" ( _
ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, _
ByVal lpOutput As Long, lpDevMode As DEVMODE) As Long

Private Sub cboPageSize_Click()
    Dim intPage As Integer
    
    txtPageSize(0).Enabled = False
    txtPageSize(1).Enabled = False
    If cboPageSize.ListIndex >= 0 Then
        intPage = cboPageSize.ItemData(cboPageSize.ListIndex)
        If intPage > 0 And intPage < 256 Then
            vsPrint.vp.PaperSize = intPage
            txtPageSize(0).Text = Format(vsPrint.vp.PageWidth / (1440 / 25.4), "0.00")
            txtPageSize(1).Text = Format(vsPrint.vp.PageHeight / (1440 / 25.4), "0.00")
        ElseIf intPage = 256 Then
            vsPrint.vp.PaperSize = intPage
            txtPageSize(0).Text = Format(vsPrint.vp.PageWidth / (1440 / 25.4), "0.00")
            txtPageSize(1).Text = Format(vsPrint.vp.PageHeight / (1440 / 25.4), "0.00")
            txtPageSize(0).Enabled = True
            txtPageSize(1).Enabled = True
        End If
    End If
End Sub


Private Sub cbo��ӡ��_Click()
    If cbo��ӡ��.ListIndex >= 0 Then
        vsPrint.vp.Device = cbo��ӡ��.List(cbo��ӡ��.ListIndex)
        Call ReadSetup
    End If
End Sub

Private Sub chkҳ��_Click()
    If chkҳ��.Value = 1 Then
        optҳ��(0).Enabled = True
        optҳ��(1).Enabled = True
        optҳ��(2).Enabled = True
    Else
        optҳ��(0).Enabled = False
        optҳ��(1).Enabled = False
        optҳ��(2).Enabled = False
    End If
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click()
    Dim strFont As String
    
    strFont = Trim(txt����Font)
    If strFont = "" Then strFont = "����|18"
    With cmdigFont
        .Flags = &H3 Or &H2 Or &H1 Or &H100
        .FontName = Split(strFont, "|")(0)
        .FontSize = Split(strFont, "|")(1)

        .ShowFont
        
        txt����Font = .FontName & "|" & .FontSize
    End With
    
End Sub

Private Sub cmdOK_Click()
    If txtPageSize(0).Text <> "" Then
        If Not IsNumeric(txtPageSize(0).Text) Then
            MsgBox "ҳ��ȴ���", vbInformation, Me.Caption
            txtPageSize(0).SetFocus
            Exit Sub
        End If
    End If
    
    If txtPageSize(1).Text <> "" Then
        If Not IsNumeric(txtPageSize(1).Text) Then
            MsgBox "ҳ�߶ȴ���", vbInformation, Me.Caption
            txtPageSize(1).SetFocus
            Exit Sub
        End If
    End If
    
    If txt�߾�(0).Text <> "" Then
        If Not IsNumeric(txt�߾�(0).Text) Then
            MsgBox "�ϱ߾����", vbInformation, Me.Caption
            txt�߾�(0).SetFocus
            Exit Sub
        End If
    End If
    
    If txt�߾�(1).Text <> "" Then
        If Not IsNumeric(txt�߾�(1).Text) Then
            MsgBox "�±߾����", vbInformation, Me.Caption
            txt�߾�(1).SetFocus
            Exit Sub
        End If
    End If
    
    If txt�߾�(2).Text <> "" Then
        If Not IsNumeric(txt�߾�(2).Text) Then
            MsgBox "��߾����", vbInformation, Me.Caption
            txt�߾�(2).SetFocus
            Exit Sub
        End If
    End If
    
    If txt�߾�(3).Text <> "" Then
        If Not IsNumeric(txt�߾�(3).Text) Then
            MsgBox "�ұ߾����", vbInformation, Me.Caption
            txt�߾�(3).SetFocus
            Exit Sub
        End If
    End If
    
    Call SaveSetup
    Unload Me
End Sub

Private Sub Form_Load()

    Call LoadPrint
End Sub

Public Sub PageSetup(ByVal Index As Long)
    mlngIndex = Index
    Me.Show vbModal
    
End Sub

Private Sub SaveSetup()
    Dim strValue As String
    
    If cbo��ӡ��.ListIndex < 0 Then Exit Sub
    
    strValue = cbo��ӡ��.List(cbo��ӡ��.ListIndex)
    WriteIni "Report" & mlngIndex, "��ӡ��", strValue, App.Path & "\PrintSetup.ini"
    
    strValue = txt����
    WriteIni "Report" & mlngIndex, "����", strValue, App.Path & "\PrintSetup.ini"
    strValue = txt����Font
    WriteIni "Report" & mlngIndex, "��������", strValue, App.Path & "\PrintSetup.ini"
    
    If chkҳ��.Value = 1 Then
        If optҳ��(0).Value = True Then
            WriteIni "Report" & mlngIndex, "��ӡҳ��", "1", App.Path & "\PrintSetup.ini"
        ElseIf optҳ��(1).Value = True Then
            WriteIni "Report" & mlngIndex, "��ӡҳ��", "2", App.Path & "\PrintSetup.ini"
        ElseIf optҳ��(2).Value = True Then
            WriteIni "Report" & mlngIndex, "��ӡҳ��", "3", App.Path & "\PrintSetup.ini"
        End If
    Else
        WriteIni "Report" & mlngIndex, "��ӡҳ��", "0", App.Path & "\PrintSetup.ini"
    End If
    
    'vp.Orientation
    If opt����(0).Value = True Then
        WriteIni "Report" & mlngIndex, "��ӡ����", "0", App.Path & "\PrintSetup.ini"
    ElseIf opt����(1).Value = True Then
        WriteIni "Report" & mlngIndex, "��ӡ����", "1", App.Path & "\PrintSetup.ini"
    End If
    
    strValue = Val(txt�߾�(0).Text)
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        WriteIni "Report" & mlngIndex, "�ϱ߾�", Val(strValue), App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "�ϱ߾�", "25.4", App.Path & "\PrintSetup.ini"
    End If
    
    strValue = Val(txt�߾�(1).Text)
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        WriteIni "Report" & mlngIndex, "�±߾�", Val(strValue), App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "�±߾�", "25.4", App.Path & "\PrintSetup.ini"
    End If
    
    strValue = Val(txt�߾�(2).Text)
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        WriteIni "Report" & mlngIndex, "��߾�", Val(strValue), App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "��߾�", "25.4", App.Path & "\PrintSetup.ini"
    End If

    strValue = Val(txt�߾�(3).Text)
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        WriteIni "Report" & mlngIndex, "�ұ߾�", Val(strValue), App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "�ұ߾�", "25.4", App.Path & "\PrintSetup.ini"
    End If

    strValue = cboPageSize.ItemData(cboPageSize.ListIndex)
    
    If Val(strValue) <> 256 Then
        vsPrint.vp.PaperSize = Val(strValue)
        WriteIni "Report" & mlngIndex, "ֽ�Ŵ�С", Val(strValue), App.Path & "\PrintSetup.ini"
        WriteIni "Report" & mlngIndex, "ֽ�ſ��", vsPrint.vp.PageWidth, App.Path & "\PrintSetup.ini"
        WriteIni "Report" & mlngIndex, "ֽ�Ÿ߶�", vsPrint.vp.PageHeight, App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "ֽ�Ŵ�С", Val(strValue), App.Path & "\PrintSetup.ini"
        WriteIni "Report" & mlngIndex, "ֽ�ſ��", Val(txtPageSize(0).Text) * (1440 / 25.4), App.Path & "\PrintSetup.ini"
        WriteIni "Report" & mlngIndex, "ֽ�Ÿ߶�", Val(txtPageSize(1).Text) * (1440 / 25.4), App.Path & "\PrintSetup.ini"
    End If
End Sub

Private Sub LoadPrint()
    Dim i As Integer, strValue As String
    
    cbo��ӡ��.Clear
    For i = 0 To vsPrint.vp.NDevices - 1
        cbo��ӡ��.AddItem vsPrint.vp.Devices(i)
    Next
    
    If cbo��ӡ��.ListCount <= 0 Then
        MsgBox "�밲װ��ӡ������ʹ�ô˹��ܣ�", vbInformation, Me.Caption
        Unload Me
        Exit Sub
    End If
    
    strValue = ReadIni("Report" & mlngIndex, "��ӡ��", App.Path & "\PrintSetup.ini")

    For i = 0 To cbo��ӡ��.ListCount - 1
        If strValue = cbo��ӡ��.List(i) Then
            cbo��ӡ��.ListIndex = i
            vsPrint.vp.Device = strValue
            Exit For
        End If
    Next
    
    If cbo��ӡ��.ListIndex < 0 Then
        cbo��ӡ��.ListIndex = 0
        vsPrint.vp.Device = cbo��ӡ��.List(cbo��ӡ��.ListIndex)
    End If
End Sub

Private Sub ReadSetup()

    Dim strValue As String, i As Integer
    
'    cboPageSize.Clear
'    For i = 1 To 256
'        If vsPrint.vp.PaperSizes(i) Then
'            strValue = getPageName(i)
'            If strValue <> "" Then
'                cboPageSize.AddItem strValue
'                cboPageSize.ItemData(cboPageSize.NewIndex) = i
'            End If
'        End If
'    Next
    Call FillPaperSize
    
    txt���� = ReadIni("Report" & mlngIndex, "����", App.Path & "\PrintSetup.ini")
    txt����Font = ReadIni("Report" & mlngIndex, "��������", App.Path & "\PrintSetup.ini")
    
    strValue = ReadIni("Report" & mlngIndex, "��ӡҳ��", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 4 Then
        chkҳ��.Value = 1
        optҳ��(Val(strValue) - 1).Value = True
        optҳ��(0).Enabled = True
        optҳ��(1).Enabled = True
        optҳ��(2).Enabled = True
    Else
        chkҳ��.Value = 0
        optҳ��(0).Enabled = False
        optҳ��(1).Enabled = False
        optҳ��(2).Enabled = False
    End If
    
    '����
    strValue = ReadIni("Report" & mlngIndex, "��ӡ����", App.Path & "\PrintSetup.ini")
    'vp.Orientation
    
    If Val(strValue) = 0 Then
        opt����(0).Value = True
    Else
        opt����(1).Value = True
    End If
    
    '�߾�
    strValue = ReadIni("Report" & mlngIndex, "�ϱ߾�", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        txt�߾�(0).Text = Format(Val(strValue), "0.0")
    Else
        txt�߾�(0).Text = "25.4"
    End If
    strValue = ReadIni("Report" & mlngIndex, "�±߾�", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        txt�߾�(1).Text = Format(Val(strValue), "0.0")
    Else
        txt�߾�(1).Text = "25.4"
    End If
    strValue = ReadIni("Report" & mlngIndex, "��߾�", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        txt�߾�(2).Text = Format(Val(strValue), "0.0")
    Else
        txt�߾�(2).Text = "25.4"
    End If
    strValue = ReadIni("Report" & mlngIndex, "�ұ߾�", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        txt�߾�(3).Text = Format(Val(strValue), "0.0")
    Else
        txt�߾�(3).Text = "25.4"
    End If
    
    'ֽ�Ŵ�С
    strValue = ReadIni("Report" & mlngIndex, "ֽ�Ŵ�С", App.Path & "\PrintSetup.ini")
    For i = 0 To cboPageSize.ListCount - 1
        If cboPageSize.ItemData(i) = Val(strValue) Then
            cboPageSize.ListIndex = i
            Exit For
        End If
    Next
    
    If cboPageSize.ListIndex < 0 Then
        For i = 0 To cboPageSize.ListCount - 1
            If cboPageSize.ItemData(i) = 9 Then
                cboPageSize.ListIndex = i 'A4
                Exit For
            End If
        Next
    End If
    'ֽ�ſ��
    If Val(strValue) <> 256 Then
        If Val(strValue) > 0 And Val(strValue) <= 256 Then vsPrint.vp.PaperSize = Val(strValue)
    Else
        strValue = ReadIni("Report" & mlngIndex, "ֽ�ſ��", App.Path & "\PrintSetup.ini")
        txtPageSize(0).Text = Val(strValue) / (1440 / 25.4)
        strValue = ReadIni("Report" & mlngIndex, "ֽ�Ÿ߶�", App.Path & "\PrintSetup.ini")
        txtPageSize(1).Text = Val(strValue) / (1440 / 25.4)
    End If
End Sub

Private Sub FillPaperSize()

    '���ݵ�ǰʹ�õĴ�ӡ��ȡ�����д�ӡֽ�źʹ�С��䵽ComBoBox��ʾ�ؼ���(cboPaper)��
    
    On Error Resume Next
    
    Dim devname As String
    Dim devoutput As String
    Dim papercount As Long
    Dim bytepapernames() As Byte
    Dim bytepapersizes() As Byte
    Dim sinfo As String
    Dim X As Long
    Dim di As Long
    Dim spapersizes As String
    Dim dv As DEVMODE
    
    Screen.MousePointer = vbHourglass
    
    devname = cbo��ӡ��.List(cbo��ӡ��.ListIndex)   ' DeviceName
    
    '��ǰ��ӡ��������
    
    devoutput = vsPrint.vp.Port
    
    '��ǰ��ӡ��������˿����Ƶõ���ǰ��ӡ���Ĵ�ӡֽ����?
    
    papercount = DeviceCapabilities(devname, devoutput, DC_PAPERNAMES, 0&, dv)
    
    If papercount = 0 Then
    
        'MsgBox "�ô�ӡ������Ч�Ĵ�ӡֽ�Ŵ�С?", vbInformation, "��ӡ����"
        Exit Sub
    
    End If
    
    'Ϊ�����ӡ��ֽ��������ռ�
    
    ReDim bytepapernames(1 To 64 * papercount)
    
    'һ��ֽ��������Ҫ64���ַ��ռ����洢
    
    'ȡ����ӡ���ϵ�������ֽ����
    
    DeviceCapabilities devname, devoutput, DC_PAPERNAMES, VarPtr(bytepapernames(1)), dv
    
    'Ϊ�����ӡ��ֽ����Ӧ��PaperSize���ַ�������ռ�
    
    ReDim bytepapersizes(1 To 2 * papercount)
    
    'һ��PaperSize��Ҫ2 ���ַ��ռ����洢ȡ����ӡ���ϵ�������ֽ����Ӧ��PaperSize
    
    DeviceCapabilities devname, devoutput, DC_PAPERS, VarPtr(bytepapersizes(1)), dv
    
    cboPageSize.Clear
    
    'Ϊ����ȷ��ȡ����?��StrConv������PaperName����ת��?
    
    For X = 1 To papercount
    
    'һ��ȡ��һ����ӡ��ֽ������
    
        sinfo = StrConv(MidB(bytepapernames, (X - 1) * 64 + 1, 64), vbUnicode)
        
        cboPageSize.AddItem Left(sinfo, InStr(sinfo, Chr(0)) - 1)  'ֽ������
        
        cboPageSize.ItemData(cboPageSize.NewIndex) = bytepapersizes((X - 1) * 2 + 1)  'ֽ�Ŵ�С
    
    Next X
    
    'If cboPageSize.ListCount > 0 Then cboPageSize.ListIndex = 0
    
    Screen.MousePointer = vbDefault

End Sub
