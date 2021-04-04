VERSION 5.00
Begin VB.Form frmCISAduitPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ����ѡ��"
   ClientHeight    =   5445
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7320
   Icon            =   "frmCISAduitPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraPrinter 
      Caption         =   "��ӡ��"
      Height          =   1470
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5775
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   3915
      End
      Begin VB.Label lblPrinterInfo2 
         AutoSize        =   -1  'True
         Caption         =   "Ĭ�ϴ�ӡ��:��"
         Height          =   180
         Left            =   1695
         TabIndex        =   4
         Top             =   990
         Width           =   1170
      End
      Begin VB.Label lblPrinterName 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   945
         TabIndex        =   1
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lblPrinterInfo 
         AutoSize        =   -1  'True
         Caption         =   "λ��:���ӵ�LTP1:"
         Height          =   180
         Left            =   1680
         TabIndex        =   3
         Top             =   645
         Width           =   1440
      End
      Begin VB.Image imgPrinter 
         Height          =   360
         Left            =   270
         Picture         =   "frmCISAduitPrint.frx":000C
         Top             =   270
         Width           =   360
      End
   End
   Begin VB.Frame fraPageScope 
      Caption         =   "��ӡ��Χ(&R)"
      Height          =   3795
      Left            =   30
      TabIndex        =   5
      Top             =   1590
      Width           =   5775
      Begin VB.ListBox lst 
         Height          =   3420
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   285
         Width           =   3600
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "ȫ��(&D)"
         Height          =   350
         Left            =   3840
         TabIndex        =   8
         Top             =   660
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   3855
         TabIndex        =   7
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ȫ�壺Shift+Delete"
         Height          =   180
         Left            =   3870
         TabIndex        =   10
         Top             =   1560
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ȫѡ��Ctrl+A"
         Height          =   180
         Left            =   3870
         TabIndex        =   9
         Top             =   1245
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6045
      TabIndex        =   11
      Top             =   135
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6045
      TabIndex        =   12
      Top             =   600
      Width           =   1100
   End
End
Attribute VB_Name = "frmCISAduitPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mstrPrinterDeviceName As String
Private mstrPrintRange As String

Public Function ShowDialog(ByVal frmMain As Object, ByRef strPrinterDeviceName As String, ByRef strPrintRange As String) As Boolean
    Dim intCount As Integer
    Dim arySerial As Variant
    Dim strTmp As String
    Dim strSelect As String
    
    mblnOK = False
    mstrPrintRange = strPrintRange
    mstrPrinterDeviceName = GetRegister(˽��ģ��, "��ӡ����", "��ӡ��", Printer.DeviceName)
    If mstrPrinterDeviceName = "" Then mstrPrinterDeviceName = Printer.DeviceName
    With cboPrinterName
        .Clear
        For intCount = 0 To Printers.count - 1
            .AddItem Printers(intCount).DeviceName
            If Printers(intCount).DeviceName = mstrPrinterDeviceName Then .ListIndex = intCount
        Next
    End With
    
    '----------------------------------------------------------------------------------------
    '1-סԺҽ��;2-סԺ����;3-������;4-�����¼;5-��ҳ��¼;6-ҽ������;7-����֤��;8-֪���ļ�
    
    strSelect = "," & GetRegister(˽��ģ��, "��ӡ����", "��ӡ����", "1,2,3,4,5,6,7,8,9") & ","
    
    strTmp = Trim(zlDatabase.GetPara("��������˳��", ParamInfo.ϵͳ��, 1560, "5;1;6;2;3;4;8;7;9"))
    If strTmp = "" Then strTmp = "5;1;6;2;3;4;8;7;9"
    arySerial = Split(strTmp, ";")
    
    With lst
        For intCount = 0 To UBound(arySerial)
            Select Case Val(arySerial(intCount))
            Case 1
                .AddItem "סԺҽ��": .ItemData(.NewIndex) = 1
                If InStr(strSelect, ",1,") > 0 Then .Selected(.NewIndex) = True
            Case 2
                .AddItem "סԺ����": .ItemData(.NewIndex) = 2
                If InStr(strSelect, ",2,") > 0 Then .Selected(.NewIndex) = True
            Case 3
                .AddItem "������": .ItemData(.NewIndex) = 3
                If InStr(strSelect, ",3,") > 0 Then .Selected(.NewIndex) = True
            Case 4
                .AddItem "�����¼": .ItemData(.NewIndex) = 4
                If InStr(strSelect, ",4,") > 0 Then .Selected(.NewIndex) = True
            Case 5
                .AddItem "��ҳ��¼": .ItemData(.NewIndex) = 5
                If InStr(strSelect, ",5,") > 0 Then .Selected(.NewIndex) = True
            Case 6
                .AddItem "ҽ������": .ItemData(.NewIndex) = 6
                If InStr(strSelect, ",6,") > 0 Then .Selected(.NewIndex) = True
            Case 7
                .AddItem "����֤��": .ItemData(.NewIndex) = 7
                If InStr(strSelect, ",7,") > 0 Then .Selected(.NewIndex) = True
            Case 8
                .AddItem "֪���ļ�": .ItemData(.NewIndex) = 8
                If InStr(strSelect, ",8,") > 0 Then .Selected(.NewIndex) = True
            Case 9
                .AddItem "�ٴ�·��": .ItemData(.NewIndex) = 9
                If InStr(strSelect, ",9,") > 0 Then .Selected(.NewIndex) = True
            End Select
        Next

        .ListIndex = 0
    End With
    
    Me.Show 1, frmMain
    
    If mblnOK Then
        strPrintRange = mstrPrintRange
        strPrinterDeviceName = mstrPrinterDeviceName
    End If
    
    ShowDialog = mblnOK
    
End Function

Private Sub cboPrinterName_Click()
    
    lblPrinterInfo.Caption = "λ��:���ӵ�" & Printers(cboPrinterName.ListIndex).Port
    lblPrinterInfo2.Caption = "Ĭ�ϴ�ӡ��:" & IIf(Printers(cboPrinterName.ListIndex).DeviceName = Printer.DeviceName, "��", "��")

End Sub

Private Sub cboPrinterName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    Dim intCount As Integer
    
    For intCount = 0 To lst.ListCount - 1
        lst.Selected(intCount) = False
    Next
    
End Sub

Private Sub cmdOK_Click()
    Dim strSelect As String
    Dim intCount As Integer
    
    mstrPrintRange = ""
    For intCount = 0 To lst.ListCount - 1
        If lst.Selected(intCount) = True Then
            mstrPrintRange = mstrPrintRange & "," & lst.ItemData(intCount)
        End If
    Next
    If mstrPrintRange <> "" Then mstrPrintRange = Mid(mstrPrintRange, 2)
    
    mstrPrinterDeviceName = cboPrinterName.Text
        
    Call SetRegister(˽��ģ��, "��ӡ����", "��ӡ��", cboPrinterName.Text)
    
    Call SetRegister(˽��ģ��, "��ӡ����", "��ӡ����", mstrPrintRange)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim intCount As Integer
    
    For intCount = 0 To lst.ListCount - 1
        lst.Selected(intCount) = True
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
    Case 1
    
        If KeyCode = vbKeyDelete Then
            Call cmdClearAll_Click
        End If
    Case 2
        If KeyCode = vbKeyA Then
            Call cmdSelectAll_Click
        End If
    End Select
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
