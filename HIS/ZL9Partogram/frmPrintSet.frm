VERSION 5.00
Begin VB.Form frmPrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ����"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmPrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Caption         =   "��ӡ��"
      Height          =   1485
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5850
      Begin VB.ComboBox cboBin 
         Height          =   300
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3885
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   3885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ֽ����Դ"
         Height          =   180
         Left            =   825
         TabIndex        =   4
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   1185
         TabIndex        =   1
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblLoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "λ��"
         Height          =   180
         Left            =   1185
         TabIndex        =   3
         Top             =   660
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmPrintSet.frx":058A
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3600
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4800
      TabIndex        =   7
      Top             =   1710
      Width           =   1100
   End
End
Attribute VB_Name = "frmPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnWinNT As Boolean
Private mdblW As Double  '��߲��ɴ�ӡ����
Private mdblH As Double  '�ϱ߲��ɴ�ӡ����

'��ӡ��������
Private mstrPrinter As String '��ӡ��
Private mstrBin As String '��ֽ��ʽ

'�¼�����
Private mblnChange As Boolean
Private mbytMode As Byte

Public Sub ShowMe(ByVal frmParent As Object, Optional ByVal bytMode As Byte = 1)
'----------------------------------------------------
'
'---------------------------------------------------
    mbytMode = bytMode
    Me.Show 1, frmParent
End Sub


Private Sub cboPrinter_Click()
    Dim i As Integer, j As Integer
    Dim lngCount As Long, strtmp As String
    Dim strPaperBinName As String * 1000
    Dim strPaperbins As String, strTemp As String, strCount As String
    
    Set Printer = Printers(cboPrinter.ItemData(cboPrinter.ListIndex))
    mstrPrinter = Printer.DeviceName
    lblLoc.Caption = "λ��: " & Printer.Port
    

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '���֧��,�򱣳�ԭ�н�ֽ��ʽ
    On Error Resume Next
    Printer.PaperBin = mstrBin
    On Error GoTo 0
    mstrBin = Printer.PaperBin
    
   '���ÿ��ý�ֽ��ʽ
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBinName, 0)
    For i = 1 To lngCount
        j = Asc(Mid(strPaperBinName, i * 2, 1)) * 256# + Asc(Mid(strPaperBinName, i * 2 - 1, 1))
        If j >= 1 And j <= 11 Then 'ֻ�г���׼֧�ֵĽ�ֽ��С
            If j = mstrBin Then
                strPaperbins = strPaperbins & "," & j & "*" 'ԭ�е�
            Else
                strPaperbins = strPaperbins & "," & j
            End If
        End If
    Next
    Err = 0
    
    If Printer.PaperBin = 14 Then
        strPaperbins = strPaperbins & ",14" _
            & IIf(mstrBin = 14, "*", "")
    End If
    
    strPaperbins = Mid(strPaperbins, 2)
'    'ֽ����Դ
    With cboBin
        .Clear
        strTemp = strPaperbins
        Do While InStr(1, strTemp, ",") > 0
            strCount = Left(strTemp, InStr(1, strTemp, ",") - 1)
            If Right(strCount, 1) = "*" Then
                .AddItem getPaperBin(CInt(Left(strCount, Len(strCount) - 1)))
                .ItemData(.NewIndex) = CInt(Left(strCount, Len(strCount) - 1))
                .ListIndex = .NewIndex
            Else
                .AddItem getPaperBin(CInt(strCount))
                .ItemData(.NewIndex) = CInt(strCount)
            End If
            strTemp = Mid(strTemp, InStr(1, strTemp, ",") + 1)
        Loop
        strCount = strTemp
        If Right(strCount, 1) = "*" Then
            .AddItem getPaperBin(CInt(Left(strCount, Len(strCount) - 1)))
            .ItemData(.NewIndex) = CInt(Left(strCount, Len(strCount) - 1))
            .ListIndex = .NewIndex
        Else
            If IsNumeric(strCount) Then
                .AddItem getPaperBin(CInt(strCount))
                .ItemData(.NewIndex) = CInt(strCount)
            End If
        End If

    End With
End Sub

Public Function getPaperBin(mBin As Integer) As String
    '------------------------------------------------
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡ��ֽ��ʽ����
    '���أ� ��ֽ��ʽ�ַ���
    '------------------------------------------------
    Err = 0
    On Error GoTo errHand
    
    If mBin = 14 Then
        getPaperBin = "���ӵĿ�ʽֽ�н�ֽ"
        Exit Function
    End If
    If mBin >= 1 And mBin <= 11 Then
        getPaperBin = Switch( _
            mBin = 1, conBin1, mBin = 2, conBin2, mBin = 3, conBin3, mBin = 4, conBin4, mBin = 5, conBin5, _
            mBin = 6, conBin6, mBin = 7, conBin7, mBin = 8, conBin8, mBin = 9, conBin9, mBin = 10, conBin10, _
            mBin = 11, conBin11)
        Exit Function
    End If
errHand:
    getPaperBin = "�Զ�ѡ��..."
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    mstrBin = ""
    mstrBin = Me.cboBin.ItemData(Me.cboBin.ListIndex)
    '�����ӡ����
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", mstrPrinter)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperBin", "")
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    If Not ExistsPrinter Then
        MsgBox "ϵͳ��û�а�װ�κδ�ӡ��,���Ȱ�װ��ӡ����", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    mblnChange = True
    
'    ��ʼ����ӡ����
    mstrPrinter = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printers(0).DeviceName)
    mstrBin = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperBin", "")
    
    '��ʼ��ӡ���б�
    With cboPrinter
        .Clear
        For i = 0 To Printers.Count - 1
            .AddItem Printers(i).DeviceName
            .ItemData(.ListCount - 1) = i '��ӡ������
            
            '��ȡ�洢�Ĵ�ӡ��Ϊ��ǰ��ӡ��,����ʼ������ҳ��
            If mstrPrinter = Printers(i).DeviceName Then .ListIndex = .NewIndex
        Next
        
        'ȱʡ��ʼ��Ϊ��ǰ��ӡ��
        If .ListIndex = -1 Then
            For i = 0 To .ListCount - 1
                '��ȡϵͳ��ǰ�Ĵ�ӡ��Ϊ��ǰ��ӡ��,����ʼ������ҳ��
                If .List(i) = Printer.DeviceName Then .ListIndex = i: Exit For
            Next
        End If
    End With
End Sub

