VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ����"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmPrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame5 
      Caption         =   "�߾�(mm)"
      Height          =   1065
      Left            =   120
      TabIndex        =   30
      Top             =   2265
      Width           =   2385
      Begin VB.TextBox txt�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1455
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   600
         Width           =   540
      End
      Begin MSComCtl2.UpDown UD�� 
         Height          =   315
         Left            =   2010
         TabIndex        =   9
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt��"
         BuddyDispid     =   196611
         OrigLeft        =   3750
         OrigTop         =   255
         OrigRight       =   3990
         OrigBottom      =   525
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD�� 
         Height          =   315
         Left            =   915
         TabIndex        =   7
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt��"
         BuddyDispid     =   196612
         OrigLeft        =   2385
         OrigTop         =   240
         OrigRight       =   2625
         OrigBottom      =   540
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD�� 
         Height          =   315
         Left            =   915
         TabIndex        =   11
         Top             =   615
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt��"
         BuddyDispid     =   196613
         OrigLeft        =   1080
         OrigTop         =   240
         OrigRight       =   1320
         OrigBottom      =   540
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1455
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   270
         Width           =   540
      End
      Begin VB.TextBox txt�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   270
         Width           =   540
      End
      Begin VB.TextBox txt�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   615
         Width           =   525
      End
      Begin MSComCtl2.UpDown UD�� 
         Height          =   300
         Left            =   2010
         TabIndex        =   13
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt��"
         BuddyDispid     =   196610
         OrigLeft        =   1080
         OrigTop         =   240
         OrigRight       =   1320
         OrigBottom      =   540
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   1245
         TabIndex        =   36
         Top             =   660
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   1245
         TabIndex        =   33
         Top             =   330
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   150
         TabIndex        =   32
         Top             =   330
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   150
         TabIndex        =   31
         Top             =   675
         Width           =   180
      End
   End
   Begin VB.Frame fraOrient 
      Caption         =   "ֽ��"
      Height          =   1065
      Left            =   2520
      TabIndex        =   34
      Top             =   2265
      Width           =   1425
      Begin VB.OptionButton opt���� 
         Caption         =   "����"
         Height          =   285
         Left            =   675
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "����"
         Height          =   285
         Left            =   675
         TabIndex        =   15
         Top             =   600
         Width           =   660
      End
      Begin VB.Image img���� 
         Height          =   480
         Left            =   120
         Picture         =   "frmPrintSet.frx":058A
         Top             =   330
         Width           =   480
      End
      Begin VB.Image img���� 
         Height          =   480
         Left            =   120
         Picture         =   "frmPrintSet.frx":0E54
         Top             =   330
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ӡ��"
      Height          =   1005
      Left            =   120
      TabIndex        =   27
      Top             =   90
      Width           =   5850
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   3885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   1185
         TabIndex        =   29
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblLoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "λ��"
         Height          =   180
         Left            =   1185
         TabIndex        =   28
         Top             =   660
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmPrintSet.frx":171E
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ֽ��"
      Height          =   1065
      Left            =   120
      TabIndex        =   21
      Top             =   1155
      Width           =   3825
      Begin VB.ComboBox cboPage 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2955
      End
      Begin VB.TextBox txtWidth 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   3
         TabIndex        =   2
         Top             =   630
         Width           =   480
      End
      Begin VB.TextBox txtHeight 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2415
         MaxLength       =   3
         TabIndex        =   4
         Top             =   630
         Width           =   480
      End
      Begin MSComCtl2.UpDown UDHeight 
         Height          =   285
         Left            =   2895
         TabIndex        =   5
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196631
         OrigLeft        =   2985
         OrigTop         =   630
         OrigRight       =   3225
         OrigBottom      =   930
         Max             =   460
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDWidth 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196630
         OrigLeft        =   1200
         OrigTop         =   645
         OrigRight       =   1440
         OrigBottom      =   945
         Max             =   460
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��С"
         Height          =   180
         Left            =   285
         TabIndex        =   26
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   300
         TabIndex        =   25
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�߶�"
         Height          =   180
         Left            =   2010
         TabIndex        =   24
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   1515
         TabIndex        =   23
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   3210
         TabIndex        =   22
         Top             =   690
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4455
      TabIndex        =   17
      Top             =   1215
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4455
      TabIndex        =   18
      Top             =   1665
      Width           =   1100
   End
   Begin VB.Frame Frame4 
      Caption         =   "��ֽ��"
      Height          =   765
      Left            =   120
      TabIndex        =   19
      Top             =   3375
      Width           =   3825
      Begin VB.ComboBox cboBin 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   285
         Width           =   2505
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ֽ����Դ"
         Height          =   180
         Left            =   210
         TabIndex        =   20
         Top             =   345
         Width           =   720
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   3975
      ScaleHeight     =   491.128
      ScaleMode       =   0  'User
      ScaleWidth      =   491.128
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2055
      Width           =   2130
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   405
         ScaleHeight     =   1455
         ScaleMode       =   0  'User
         ScaleWidth      =   1140
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   270
         Width           =   1170
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   450
         ScaleHeight     =   1485
         ScaleWidth      =   1170
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   315
         Width           =   1170
      End
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
Private mintPage As Integer 'ֽ��
Private mlngWidth As Long '�Զ���ֽ�ſ��,Twip
Private mlngHeight As Long '�Զ���ֽ�Ÿ߶�'Twip
Private mintOrient As Integer   'ֽ��
Private mintBin As Integer '��ֽ��ʽ
Private mlngLeft As Long '��߾�'mm
Private mlngRight As Long '�ұ߾�'mm
Private mlngTop As Long '�ϱ߾�'mm
Private mlngBottom As Long '�±߾�'mm

'�¼�����
Private mblnChange As Boolean

Private Sub cboBin_Click()
    If cboBin.ListIndex <> -1 Then
        mintBin = cboBin.ItemData(cboBin.ListIndex)
    End If
End Sub

Private Sub cboPage_Click()
    Dim blnOK As Boolean
    Dim dblRight As Double
    Dim dblDown As Double
    
    
    'ֽ��
    If cboPage.ItemData(cboPage.ListIndex) <> 256 Then
        Printer.PaperSize = cboPage.ItemData(cboPage.ListIndex)
        mintPage = Printer.PaperSize
    Else
        'ǿ�������Զ���ֽ�ſ���,�����
        mintPage = 256
    End If
        
    'ֽ��
    If mintPage <> 256 Then
        On Error Resume Next
        Printer.Orientation = mintOrient
        mintOrient = Printer.Orientation
    Else
        mintOrient = 1
    End If
    On Error GoTo 0
    fraOrient.Enabled = mintPage <> 256
        
    '���ʵ������ֽ�Ŵ�С(ֽ��Ӱ��֮��)
    If mintPage <> 256 Then
        'ȡ�ô�ӡ��֧�ָ÷������ʵ�ߴ�
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
        
        '���ɴ�ӡ�������
        mdblW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
        mdblH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
    Else
        '�Զ���ֽ����Ϊȫ�����Դ�ӡ
        mdblW = 0
        mdblH = 0
    End If
    
    '��ʾֽ�ųߴ�
    mblnChange = False
    txtWidth.Tag = mlngWidth
    txtWidth.Text = CLng(mlngWidth / 56.7)
    txtHeight.Tag = mlngHeight
    txtHeight.Text = CLng(mlngHeight / 56.7)
    mblnChange = True
    
    '��ʾ���ñ߾�
    '��С�ڿɴ�ӡ����֮��
    '��󲻳�����ߵ�1/4
    UD��.Min = mlngWidth / 56.7 * mdblW
    UD��.Max = mlngWidth / 56.7 / 4
    UD��.Min = UD��.Min
    UD��.Max = UD��.Max
    
    UD��.Min = mlngHeight / 56.7 * mdblH
    UD��.Max = mlngHeight / 56.7 / 4
    UD��.Min = UD��.Min
    UD��.Max = UD��.Max
    
    If mlngLeft >= UD��.Min And mlngLeft <= UD��.Max Then
        UD��.Value = mlngLeft
    Else
        UD��.Value = UD��.Min
    End If
    If mlngRight >= UD��.Min And mlngRight <= UD��.Max Then
        UD��.Value = mlngRight
    Else
        UD��.Value = UD��.Min
    End If
    If mlngTop >= UD��.Min And mlngTop <= UD��.Max Then
        UD��.Value = mlngTop
    Else
        UD��.Value = UD��.Min
    End If
    If mlngBottom >= UD��.Min And mlngBottom <= UD��.Max Then
        UD��.Value = mlngBottom
    Else
        UD��.Value = UD��.Min
    End If
    
    mlngLeft = UD��.Value
    mlngRight = UD��.Value
    mlngTop = UD��.Value
    mlngBottom = UD��.Value
    
    '��ʾֽ��
    mblnChange = False
    If mintOrient = 1 Then
        opt����.Value = True: opt����_Click
    Else
        opt����.Value = True: opt����_Click
    End If
    mblnChange = True
    
    '��ʾԤ��ֽ��
    Call ShowPaper
End Sub

Private Sub cboPrinter_Click()
    Dim i As Integer, j As Integer
    Dim lngCount As Long, strTmp As String
    Dim strPaperSize As String * 300
    Dim strPaperBin As String * 100
    Dim strPaperBinName As String * 1000
    
    Set Printer = Printers(cboPrinter.ItemData(cboPrinter.ListIndex))
    mstrPrinter = Printer.DeviceName
    lblLoc.Caption = "λ��: " & Printer.Port
    
    '���֧��,�򱣳�ԭ��ֽ��
    If mintPage <> 256 Then
        On Error Resume Next
        Printer.PaperSize = mintPage
        On Error GoTo 0
        mintPage = Printer.PaperSize
    End If
    
    '���ÿ���ֽ��
    '���ظ�ʽ������˵����MSDN
    cboPage.Clear
    '------------------------------------------------------------------------------------------
    'ֽ�Ŵ�С
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, strPaperSize, 0)
    For i = 1 To lngCount
        j = Asc(Mid(strPaperSize, i * 2, 1)) * 256 + Asc(Mid(strPaperSize, i * 2 - 1, 1))
        If j >= 1 And j <= 41 Then 'ֻ�г���׼֧�ֵ�ֽ��
            cboPage.AddItem GetPaperName(j)
            cboPage.ItemData(cboPage.ListCount - 1) = j
            If j = mintPage Then cboPage.ListIndex = cboPage.NewIndex
        End If
    Next
    '------------------------------------------------------------------------------------------
    '�Զ���ֽ�Ŵ���
    i = 256
    cboPage.AddItem GetPaperName(i)
    cboPage.ItemData(cboPage.ListCount - 1) = i
    If mintPage = 256 Then cboPage.ListIndex = cboPage.NewIndex
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '���֧��,�򱣳�ԭ�н�ֽ��ʽ
    On Error Resume Next
    Printer.PaperBin = mintBin
    On Error GoTo 0
    mintBin = Printer.PaperBin
    
    '���ÿ��ý�ֽ��ʽ
    cboBin.Clear
    '------------------------------------------------------------------------------------------
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINNAMES, strPaperBinName, 0)
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBin, 0)
    j = 1
    For i = 1 To lngCount
        '��ֽ����
        Do
            If Mid(strPaperBinName, j, 1) = Chr(0) Then
                cboBin.AddItem Trim(strTmp)
                
                '��ֽ���
                cboBin.ItemData(cboBin.ListCount - 1) = Asc(Mid(strPaperBin, i * 2, 1)) * 256 + Asc(Mid(strPaperBin, i * 2 - 1, 1))
                If cboBin.ItemData(cboBin.ListCount - 1) = mintBin Then
                    cboBin.ListIndex = cboBin.NewIndex
                End If

                j = 24 + j - LenB(StrConv(strTmp, vbFromUnicode))
                strTmp = ""
                Exit Do
            Else
                strTmp = strTmp & Mid(strPaperBinName, j, 1)
                j = j + 1
            End If
        Loop
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txtWidth.Text) Then
        MsgBox "��ȷ�������ֽ�ſ�ȣ�", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Sub
    End If
    If CInt(txtWidth.Text) > UDWidth.Max Then
        MsgBox "�����ֽ�ſ�Ȳ��ܳ���" & UDWidth.Max & "���ף�", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Sub
    End If
    
    If Not IsNumeric(txtHeight.Text) Then
        MsgBox "��ȷ�������ֽ�Ÿ߶ȣ�", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Sub
    End If
    If CInt(txtHeight.Text) > UDHeight.Max Then
        MsgBox "�����ֽ�Ÿ߶Ȳ��ܳ���" & UDHeight.Max & "���ף�", vbExclamation, App.Title
        txtHeight.SetFocus: Exit Sub
    End If
    
    '�����ӡ����
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��ӡ��", mstrPrinter
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", mintPage
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "���", mlngWidth
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�߶�", mlngHeight
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", mintOrient
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��ֽ", mintBin
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��߾�", mlngLeft
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ұ߾�", mlngRight
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ϱ߾�", mlngTop
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�±߾�", mlngBottom
    
    gblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    If Printers.Count = 0 Then
        MsgBox "ϵͳ��û�а�װ�κδ�ӡ��,���Ȱ�װ��ӡ����", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    gblnOK = False
    mblnChange = True
    
    mblnWinNT = IsWindowsNT
    
    '��ʼ����ӡ����
    mstrPrinter = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��ӡ��", Printer.DeviceName)
    mintPage = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", Printer.PaperSize)
    mlngWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "���", Printer.Width)
    mlngHeight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�߶�", Printer.Height)
    mintOrient = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", Printer.Orientation)
    mintBin = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��ֽ", Printer.PaperBin)
    mlngLeft = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��߾�", OFFSET_LEFT)
    mlngRight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ұ߾�", OFFSET_RIGHT)
    mlngTop = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ϱ߾�", OFFSET_TOP)
    mlngBottom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�±߾�", OFFSET_BOTTOM)
    
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
    
    '�߾�
    txt��.Text = mlngLeft
    txt��.Text = mlngRight
    txt��.Text = mlngTop
    txt��.Text = mlngBottom
End Sub

Private Sub opt����_Click()
    Dim lngL As Long, lngR As Long
    Dim lngT As Long, lngB As Long
    
    If opt����.Value Then
        img����.Visible = False
        img����.Visible = True
        
        If mintOrient = 1 Then
            lngL = mlngLeft
            lngR = mlngRight
            lngT = mlngTop
            lngB = mlngBottom
            
            mlngLeft = lngB
            mlngRight = lngT
            mlngTop = lngL
            mlngBottom = lngR
        End If
        
        mintOrient = 2
        
        If mblnChange Then Call cboPage_Click
    End If
End Sub

Private Sub opt����_Click()
    Dim lngL As Long, lngR As Long
    Dim lngT As Long, lngB As Long
    
    If opt����.Value Then
        img����.Visible = True
        img����.Visible = False
        
        If mintOrient = 2 Then
            lngL = mlngLeft
            lngR = mlngRight
            lngT = mlngTop
            lngB = mlngBottom
              
            mlngLeft = lngT
            mlngRight = lngB
            mlngTop = lngR
            mlngBottom = lngL
        End If
        
        mintOrient = 1
        
        If mblnChange Then Call cboPage_Click
    End If
End Sub

Private Sub txtHeight_Change()
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtHeight.Text) Then
        txtHeight.Tag = CLng(txtHeight.Text * 56.7)
        mlngHeight = CLng(txtHeight.Text * 56.7)
        
        cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ShowPaper
End Sub

Private Sub txtWidth_Change()
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtWidth.Text) Then
        txtWidth.Tag = CLng(txtWidth.Text * 56.7)
        mlngWidth = CLng(txtWidth.Text * 56.7)
        
        cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ShowPaper
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
End Sub

Private Sub txtHeight_GotFocus()
    txtHeight.SelStart = 0: txtHeight.SelLength = Len(txtHeight.Text)
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = 0: txtWidth.SelLength = Len(txtWidth.Text)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
End Sub

Private Sub txt��_GotFocus()
    zlControl.TxtSelAll txt��
End Sub

Private Sub txt��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��_GotFocus()
    zlControl.TxtSelAll txt��
End Sub

Private Sub txt��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��_GotFocus()
    zlControl.TxtSelAll txt��
End Sub

Private Sub txt��_GotFocus()
    zlControl.TxtSelAll txt��
End Sub

Private Sub txt��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub UD��_Change()
    mlngTop = UD��.Value
    Call ShowPaper
End Sub

Private Sub UD��_Change()
    mlngBottom = UD��.Value
    Call ShowPaper
End Sub

Private Sub UD��_Change()
    mlngRight = UD��.Value
    Call ShowPaper
End Sub

Private Sub UD��_Change()
    mlngLeft = UD��.Value
    Call ShowPaper
End Sub

Private Sub ShowPaper()
'���ܣ���ʾ���õ�ֽ�ŵ�Ԥ��
    On Error Resume Next
    
    picPaper.Cls
    
    picPaper.Width = mlngWidth / 56.7
    picPaper.Height = mlngHeight / 56.7
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth
    picPaper.ScaleHeight = mlngHeight
    
    picPaper.Line (0, mlngTop * 56.7)-(picPaper.ScaleWidth, mlngTop * 56.7), &H808080
    picPaper.Line (0, picPaper.ScaleHeight - (mlngBottom + 2) * 56.7)-(picPaper.ScaleWidth, picPaper.ScaleHeight - (mlngBottom + 2) * 56.7), &H808080
    
    picPaper.Line (mlngLeft * 56.7, 0)-(mlngLeft * 56.7, picPaper.ScaleHeight), &H808080
    picPaper.Line (picPaper.ScaleWidth - (mlngRight + 2) * 56.7, 0)-(picPaper.ScaleWidth - (mlngRight + 2) * 56.7, picPaper.ScaleHeight), &H808080
    
    Me.Refresh
End Sub
