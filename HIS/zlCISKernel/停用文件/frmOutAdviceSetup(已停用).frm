VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOutAdviceSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ҽ��ѡ��"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmOutAdviceSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabPar 
      Height          =   6090
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   10742
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   617
      WordWrap        =   0   'False
      TabCaption(0)   =   "ҽ���´�(&1)"
      TabPicture(0)   =   "frmOutAdviceSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl����ҩ��"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "vsfDrugStore"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbo����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraLine"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraPurMed"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "ҽ������(&2)"
      TabPicture(1)   =   "frmOutAdviceSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPBPSet"
      Tab(1).Control(1)=   "frmPoint"
      Tab(1).Control(2)=   "fraSendNO"
      Tab(1).Control(3)=   "chk�ر�ҽ��"
      Tab(1).Control(4)=   "chkִ��"
      Tab(1).Control(5)=   "fraBillPrint"
      Tab(1).Control(6)=   "fra���"
      Tab(1).Control(7)=   "Frame4"
      Tab(1).ControlCount=   8
      Begin VB.Frame fraPBPSet 
         Height          =   1140
         Left            =   -70110
         TabIndex        =   42
         Top             =   1560
         Width           =   4560
         Begin VB.CommandButton cmdPBPSet 
            Caption         =   "֧��Ʊ�ݴ�ӡ����"
            Height          =   300
            Left            =   390
            TabIndex        =   44
            Top             =   450
            Width           =   1620
         End
         Begin VB.CheckBox chkSendPay 
            Caption         =   "����ʱ���п�����֧��(���֧��)"
            Height          =   360
            Left            =   135
            TabIndex        =   43
            Top             =   -60
            Width           =   3015
         End
      End
      Begin VB.Frame frmPoint 
         Caption         =   "���ͺ�,ָ����"
         Height          =   1350
         Left            =   -67680
         TabIndex        =   38
         Top             =   4590
         Width           =   2145
         Begin VB.OptionButton optPoint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   41
            Top             =   900
            Width           =   1560
         End
         Begin VB.OptionButton optPoint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   40
            Top             =   600
            Width           =   1560
         End
         Begin VB.OptionButton optPoint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   39
            Top             =   300
            Value           =   -1  'True
            Width           =   1560
         End
      End
      Begin VB.Frame fraPurMed 
         Caption         =   "����ҩ��ȱʡ��ҩĿ��"
         Height          =   765
         Left            =   5250
         TabIndex        =   33
         Top             =   2190
         Width           =   4215
         Begin VB.OptionButton optPurMed 
            Caption         =   "�´�ʱȷ��"
            Height          =   180
            Index           =   0
            Left            =   270
            TabIndex        =   45
            Top             =   360
            Width           =   1560
         End
         Begin VB.OptionButton optPurMed 
            Caption         =   "Ԥ��"
            Height          =   180
            Index           =   1
            Left            =   1890
            TabIndex        =   35
            Top             =   360
            Width           =   680
         End
         Begin VB.OptionButton optPurMed 
            Caption         =   "����"
            Height          =   180
            Index           =   2
            Left            =   3120
            TabIndex        =   34
            Top             =   360
            Value           =   -1  'True
            Width           =   680
         End
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Left            =   5280
         TabIndex        =   31
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cbo���� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   6540
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1335
         Width           =   2310
      End
      Begin VB.Frame fraSendNO 
         Caption         =   "���ݲ�������"
         Height          =   4395
         Left            =   -74880
         TabIndex        =   22
         Top             =   1545
         Width           =   4605
         Begin VB.CheckBox chkTimeDef 
            Caption         =   "��ʼʱ�䲻��ͬһ��ķֱ��������"
            Height          =   180
            Left            =   240
            TabIndex        =   46
            Top             =   600
            Width           =   3480
         End
         Begin VB.CheckBox chkNOType 
            Caption         =   "��ͬ��ϵ�ҽ���ֱ��������"
            Height          =   180
            Left            =   240
            TabIndex        =   37
            Top             =   315
            Width           =   2760
         End
         Begin VB.OptionButton optSendNO 
            Caption         =   "�������ҽ������ִͬ�п���ֻ����һ�ŵ���"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   36
            Top             =   1200
            Width           =   4140
         End
         Begin VB.CheckBox chkһ����ҩ���� 
            Caption         =   "һ����ҩ�ļ�ʹ�����㲻ͬҲ����Ϊһ�ŵ���"
            Height          =   255
            Left            =   465
            TabIndex        =   32
            Top             =   3045
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.OptionButton optSendNO 
            Caption         =   "ÿ�η���ҽ��ֻ����һ�ŵ���"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   915
            Width           =   3060
         End
         Begin VB.OptionButton optSendNO 
            Caption         =   "����ͬһ���ҽ����ִͬ�п���ֻ����һ�ŵ���"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   1515
            Width           =   4140
         End
         Begin VB.ListBox lstSendNO 
            Columns         =   4
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   1110
            IMEMode         =   3  'DISABLE
            Left            =   465
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   1815
            Width           =   3660
         End
         Begin VB.Label lblPrompt 
            Caption         =   $"frmOutAdviceSetup.frx":0044
            Height          =   825
            Left            =   465
            TabIndex        =   26
            Top             =   3360
            Width           =   3735
         End
      End
      Begin VB.CheckBox chk�ر�ҽ�� 
         Caption         =   "�������֮���Զ��رշ��ʹ���"
         Height          =   195
         Left            =   -69945
         TabIndex        =   13
         Top             =   1095
         Width           =   2940
      End
      Begin VB.CheckBox chkִ�� 
         Caption         =   "����ʱ������ִ�е���Ϊ��ִ��"
         Height          =   195
         Left            =   -69945
         TabIndex        =   12
         Top             =   675
         Width           =   2820
      End
      Begin VB.Frame fraBillPrint 
         Caption         =   "���ͺ�,���Ƶ���"
         Height          =   1350
         Left            =   -70080
         TabIndex        =   18
         Top             =   4590
         Width           =   2235
         Begin VB.OptionButton optPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   9
            Top             =   300
            Width           =   1560
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   10
            Top             =   600
            Value           =   -1  'True
            Width           =   1560
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   11
            Top             =   900
            Width           =   1560
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " ҽ������ "
         Height          =   2565
         Left            =   5280
         TabIndex        =   17
         Top             =   3360
         Width           =   4215
         Begin VB.CommandButton cmdBloodTip 
            Caption         =   "��Ѫ����ע����������"
            Height          =   350
            Left            =   105
            TabIndex        =   47
            Top             =   2100
            Width           =   2490
         End
         Begin VB.CheckBox chkMustAddAgent 
            Caption         =   "�´ﶾ��͵�һ�ྫ��ҩƷʱ����ǼǴ�����"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1770
            Width           =   3960
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�´�ҩƷҽ��ʱ����¼��ҩƷ����"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   3360
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�´�ҩƷҽ��ʱ����ָ����ҩ����"
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Top             =   915
            Width           =   3360
         End
         Begin VB.CheckBox chkƤ�� 
            Caption         =   "�Զ�����Ƥ�Բ����ݽ������ҽ������"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   1335
            Width           =   3360
         End
      End
      Begin VB.Frame fra��� 
         Height          =   1530
         Left            =   -70080
         TabIndex        =   20
         Top             =   2850
         Width           =   4530
         Begin VB.ListBox lst��� 
            Columns         =   3
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   360
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   390
            Width           =   2580
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "�����������ʱ��������д"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   2640
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " ���͵��� "
         Height          =   960
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   4605
         Begin VB.CheckBox chk��λ���� 
            Caption         =   "ֻ�к�Լ��λ���˵�ҽ���ſ��Է���Ϊ���ʵ�"
            Height          =   195
            Left            =   255
            TabIndex        =   6
            Top             =   630
            Width           =   3960
         End
         Begin VB.OptionButton optSend 
            Caption         =   "����ʱ��ȷ��"
            Height          =   180
            Index           =   2
            Left            =   2565
            TabIndex        =   5
            Top             =   330
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton optSend 
            Caption         =   "���ʵ���"
            Height          =   180
            Index           =   1
            Left            =   1395
            TabIndex        =   4
            Top             =   330
            Width           =   1020
         End
         Begin VB.OptionButton optSend 
            Caption         =   "�շѵ���"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   330
            Width           =   1020
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   5445
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   5055
         _cx             =   8916
         _cy             =   9604
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutAdviceSetup.frx":00E8
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ���ϲ���"
         Height          =   180
         Left            =   5280
         TabIndex        =   30
         Top             =   1380
         Width           =   1080
      End
      Begin VB.Label lbl����ҩ�� 
         Caption         =   $"frmOutAdviceSetup.frx":0195
         Height          =   615
         Left            =   5280
         TabIndex        =   28
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8625
      TabIndex        =   15
      Top             =   6345
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7530
      TabIndex        =   14
      Top             =   6345
      Width           =   1100
   End
End
Attribute VB_Name = "frmOutAdviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mMainPrivs As String
Public mblnҽ��վ As Boolean
Private Const VsPubBackColor = &HFAEADA

Private Sub chkSendPay_Click()
    cmdPBPSet.Enabled = chkSendPay.value
End Sub

Private Sub chk���_Click()
    lst���.Enabled = chk���.value = 1 And lst���.Tag = ""
End Sub

Private Sub cmdBloodTip_Click()
    Dim strPar As String
    strPar = cmdBloodTip.Tag
    Call frmInputBox.InputBox(Me, "��Ѫ����ע������", "���ݣ�", 4000, 6, True, True, strPar)
    cmdBloodTip.Tag = strPar
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim str��� As String, strSendNO As String
    Dim i As Long, bytType As Long
    Dim arr����ҩ��(3) As String, arrȱʡҩ��(3) As String, arrTmp() As String
    Dim blnSetup As Boolean
    Dim str��ҩ������ As String, str��ҩ������ As String, str��ҩ������ As String
    
    '������Ƿ�ָ����ȱʡҩ������Ϊ����û�в�������Ȩ�ޣ����������ǿ��Զ���ġ�
    
    If mblnҽ��վ = False Then
        If chk���.value = 1 Then
            For i = 0 To lst���.ListCount - 1
                If lst���.Selected(i) Then
                    str��� = str��� & Chr(lst���.ItemData(i))
                End If
            Next
            If str��� = "" Then
                MsgBox "������ѡ��һ��Ҫ�����ϵ�ҽ�����", vbInformation, gstrSysName
                tabPar.Tab = 1: lst���.SetFocus: Exit Sub
            End If
        End If
    End If
        
    '����ѡ��
    strSendNO = ""
    For i = 0 To lstSendNO.ListCount - 1
        If lstSendNO.Selected(i) Then
            strSendNO = strSendNO & Chr(lstSendNO.ItemData(i))
        End If
    Next
    
    '----------------------------------------------------------------------------------------------------
    'ҩ��
    With vsfDrugStore
        For i = .FixedRows To .Rows - 1
            Select Case .TextMatrix(i, .ColIndex("���"))
            Case "��ҩ��"
                bytType = 0
                If .TextMatrix(i, .ColIndex("��ҩ����")) <> "�Զ�����" Then
                    str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("��ҩ����"))
                End If
            Case "��ҩ��"
                bytType = 1
                If .TextMatrix(i, .ColIndex("��ҩ����")) <> "�Զ�����" Then
                    str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("��ҩ����"))
                End If
            Case "��ҩ��"
                bytType = 2
                If .TextMatrix(i, .ColIndex("��ҩ����")) <> "�Զ�����" Then
                    str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("��ҩ����"))
                End If
            End Select
            If .TextMatrix(i, .ColIndex("����")) <> 0 Then arr����ҩ��(bytType) = arr����ҩ��(bytType) & "," & .RowData(i)
            If .TextMatrix(i, .ColIndex("ȱʡ")) = "��" Then arrȱʡҩ��(bytType) = .RowData(i)
        Next
    End With
    
    blnSetup = InStr(GetInsidePrivs(p����ҽ���´�), ";ҽ��ѡ������;") > 0
    arrTmp = Split("��ҩ��,��ҩ��,��ҩ��", ",")
    For bytType = 0 To UBound(arrTmp)
        Call zlDatabase.SetPara("�������" & arrTmp(bytType), Mid(arr����ҩ��(bytType), 2), glngSys, p����ҽ���´�, blnSetup)
        Call zlDatabase.SetPara("����ȱʡ" & arrTmp(bytType), arrȱʡҩ��(bytType), glngSys, p����ҽ���´�, blnSetup)
    Next
    Call zlDatabase.SetPara("��ҩ������", Mid(str��ҩ������, 2), glngSys, p����ҽ���´�, blnSetup)
    Call zlDatabase.SetPara("��ҩ������", Mid(str��ҩ������, 2), glngSys, p����ҽ���´�, blnSetup)
    Call zlDatabase.SetPara("��ҩ������", Mid(str��ҩ������, 2), glngSys, p����ҽ���´�, blnSetup)
          
    Call zlDatabase.SetPara("����ȱʡ���ϲ���", IIF(cbo����.ListIndex = 0, "0", cbo����.ItemData(cbo����.ListIndex)), glngSys, p����ҽ���´�, blnSetup)
    
    '����¼��ҩƷ����
    Call zlDatabase.SetPara("����¼��ҩƷ����", chk����.value, glngSys, p����ҽ���´�, blnSetup)
    
    'ҽ��ִ������
    Call zlDatabase.SetPara("ҽ��ִ������", chk����.value, glngSys, p����ҽ���´�, blnSetup)
    
    '����ҩ��ȱʡ��ҩĿ��
    For i = 0 To 2
        If optPurMed(i).value Then
            Call zlDatabase.SetPara("����ҩ��ȱʡ��ҩĿ��", i & "", glngSys, p����ҽ���´�, blnSetup)
            Exit For
        End If
    Next
    
    '----------------------------------------------------------------------------------------------------
    '����ѡ��
    Call zlDatabase.SetPara("���͵�������", IIF(optSend(0).value, 0, IIF(optSend(1).value, 1, 2)), glngSys, p����ҽ���´�, blnSetup)
        
    '����Լ��λ���˷���Ϊ���ʵ�
    Call zlDatabase.SetPara("��λ����", chk��λ����.value, glngSys, p����ҽ���´�, blnSetup)

    '�´ﶾ��͵�һ�ྫ��ҩƷҽ��ʱ����ǼǴ�����
    Call zlDatabase.SetPara("Ҫ��ǼǴ�����", chkMustAddAgent.value, glngSys, p����ҽ���´�, blnSetup)
        
    '����ִ���Զ����
    Call zlDatabase.SetPara("���ﱾ���Զ�ִ��", chkִ��.value, glngSys, p����ҽ���´�, blnSetup)

    '�ر�ҽ������
    Call zlDatabase.SetPara("������ɺ�ر�ҽ������", chk�ر�ҽ��.value, glngSys, p����ҽ���´�, blnSetup)
    
    If mblnҽ��վ = False Then
        '�Զ�����Ƥ��
        Call zlDatabase.SetPara("�Զ�����Ƥ��", chkƤ��.value, glngSys, p����ҽ���´�, blnSetup)
        
        '���ݴ�ӡ:0-����ӡ,1-�ֹ���ӡ,2-�Զ���ӡ
        Call zlDatabase.SetPara("���﷢�͵��ݴ�ӡ", IIF(optPrint(0).value, 0, IIF(optPrint(1).value, 1, 2)), glngSys, p����ҽ���´�, blnSetup)
        
        'Ҫ�������������
        Call zlDatabase.SetPara("Ҫ�������������", str���, glngSys, p����ҽ���´�, blnSetup)
    End If
     
    '��ͬ��ϵ�ҽ���ֱ��������
    Call zlDatabase.SetPara("��ͬ��ϵ�ҽ���ֱ��������", chkNOType.value, glngSys, p����ҽ���´�, blnSetup)
    '��ʼʱ�䲻��ͬһ��ķֱ��������
    Call zlDatabase.SetPara("��ʼʱ�䲻��ͬһ��ķֱ��������", chkTimeDef.value, glngSys, p����ҽ���´�, blnSetup)
    
    '���͵��ݺ�
    Call zlDatabase.SetPara("���͵��ݺŹ���", IIF(optSendNO(0).value, 1, IIF(optSendNO(2).value, 2, 0)), glngSys, p����ҽ���´�, blnSetup) '0-���,1-����,2-����
    
    '����Ϊͬһ���ݵ�ҽ�����
    Call zlDatabase.SetPara("����Ϊͬһ���ݵ�ҽ�����", strSendNO, glngSys, p����ҽ���´�, blnSetup)
    
    Call zlDatabase.SetPara("һ����ҩ����Ϊһ��", chkһ����ҩ����.value, glngSys, p����ҽ���´�, blnSetup)

    'ָ������ӡ
    Call zlDatabase.SetPara("ָ������ӡ��ʽ", IIF(optPoint(0).value, 0, IIF(optPoint(1).value, 1, 2)), glngSys, p����ҽ���´�, blnSetup)
    
    '���֧��
    Call zlDatabase.SetPara("�������֧��", chkSendPay.value, glngSys, p����ҽ���´�, blnSetup)
    
    Call zlDatabase.SetPara("��Ѫ����ע������", cmdBloodTip.Tag, glngSys, p����ҽ���´�, blnSetup)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPBPSet_Click()
    On Error Resume Next
    If gobjSquareCard Is Nothing Then
        Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If gobjSquareCard.zlInitComponents(Me, p����ҽ���´�, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set gobjSquareCard = Nothing
            MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
            err.Clear: Exit Sub
        End If
    End If
    Call gobjSquareCard.zlCliniqueRoomPayPrintSet(Me)
    err.Clear: On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strPar As String
    Dim blnSetup As Boolean, arrTmp() As String
    Dim strDSIDs As String, strDefault As String, lngBackColor As Long, bytLockEdit As Byte
    Dim intType1 As Integer, intType2 As Integer, lngRow As Long
    Dim str���� As String, j As Integer
    
    On Error GoTo errH
    
    gblnOK = False
    
    If mblnҽ��վ Then
        chkƤ��.Visible = False
        fraBillPrint.Visible = False
        fra���.Visible = False
    End If
    
    blnSetup = InStr(GetInsidePrivs(p����ҽ���´�), "ҽ��ѡ������") > 0
    '------------------------------------------------------------------------------------------------------------------------
    'ҩ���뷢�ϲ���
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.�������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " AND B.����ID=A.ID And B.������� IN(1,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��','���ϲ���')" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by ��������,����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With vsfDrugStore
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        .MergeCol(.ColIndex("���")) = True
        .MergeCells = flexMergeFixedOnly
        
        rsTmp.Filter = "��������<>'���ϲ���'"
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = .FixedRows
            arrTmp = Split("��ҩ��,��ҩ��,��ҩ��", ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "��������='" & arrTmp(i) & "'"
                strDefault = zlDatabase.GetPara("����ȱʡ" & arrTmp(i), glngSys, p����ҽ���´�, , , , intType1)
                strDSIDs = "," & zlDatabase.GetPara("�������" & arrTmp(i), glngSys, p����ҽ���´�, , , , intType2) & ","
                '��ҩ����
                str���� = zlDatabase.GetPara(arrTmp(i) & "����", glngSys, p����ҽ���´�, , , blnSetup)
                Do While Not rsTmp.EOF
                    .TextMatrix(lngRow, .ColIndex("���")) = arrTmp(i)
                    .TextMatrix(lngRow, .ColIndex("ҩ��")) = rsTmp!����
                    .RowData(lngRow) = Val(rsTmp!ID)
                    
                    If Val(rsTmp!ID) = Val(strDefault) Then
                        .TextMatrix(lngRow, .ColIndex("ȱʡ")) = "��"
                        .TextMatrix(lngRow, .ColIndex("����")) = -1   'true
                    Else
                        .TextMatrix(lngRow, .ColIndex("ȱʡ")) = ""
                        .TextMatrix(lngRow, .ColIndex("����")) = IIF(InStr(strDSIDs, "," & rsTmp!ID & ",") > 0, -1, 0)
                    End If
                    
                    'ȱʡ��Ԫ��
                    'intType-'���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType1 & ",") > 0 Then
                        lngBackColor = IIF(blnSetup, VsPubBackColor, &H8000000F)      '��Ȩ�޿���
                        bytLockEdit = IIF(blnSetup, 0, 1)
                    ElseIf intType1 = 5 Then
                        lngBackColor = VsPubBackColor       '����ģ��,������Ȩ�޿���
                    Else
                        lngBackColor = &H80000005     '�����༭
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("ȱʡ")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("ȱʡ")) = bytLockEdit
                     
                    '���õ�Ԫ��
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType2 & ",") > 0 Then
                        lngBackColor = IIF(blnSetup, VsPubBackColor, &H8000000F)      '��Ȩ�޿���
                        bytLockEdit = IIF(blnSetup, 0, 1)
                    ElseIf intType2 = 5 Then
                        lngBackColor = VsPubBackColor       '����ģ��,������Ȩ�޿���
                    Else
                        lngBackColor = &H80000005     '�����༭
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("����")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("����")) = bytLockEdit
                    
                    '��ҩ����
                    For j = 0 To UBound(Split(str����, ","))
                        If Val(.RowData(lngRow)) = Val(Split(Split(str����, ",")(j), ":")(0)) Then
                            .TextMatrix(lngRow, .ColIndex("��ҩ����")) = Split(Split(str����, ",")(j), ":")(1)
                            Exit For
                        End If
                    Next
                    If .TextMatrix(lngRow, .ColIndex("��ҩ����")) = "" Then .TextMatrix(lngRow, .ColIndex("��ҩ����")) = "�Զ�����"
                    .Cell(flexcpBackColor, lngRow, .ColIndex("��ҩ����")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("��ҩ����")) = bytLockEdit
                    
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Loop
                If lngRow < .Rows - 1 Then  '���ָ���
                    .Select lngRow, .FixedCols, lngRow, .Cols - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
            Next
        End If
    End With
    
    cbo����.AddItem "�˹�ѡ��"
    rsTmp.Filter = "��������='���ϲ���'"
    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!����
        cbo����.ItemData(cbo����.ListCount - 1) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    strPar = zlDatabase.GetPara("����ȱʡ���ϲ���", glngSys, p����ҽ���´�, , Array(lbl����, cbo����), blnSetup)
    zlControl.CboLocate cbo����, strPar, True
        
    '����¼��ҩƷ����
    chk����.value = Val(zlDatabase.GetPara("����¼��ҩƷ����", glngSys, p����ҽ���´�, , Array(chk����), blnSetup))
    
    'ҽ��ִ������
    chk����.value = Val(zlDatabase.GetPara("ҽ��ִ������", glngSys, p����ҽ���´�, , Array(chk����), blnSetup))
    
    '����ҩ��ȱʡ��ҩĿ��
    strPar = zlDatabase.GetPara("����ҩ��ȱʡ��ҩĿ��", glngSys, p����ҽ���´�, "0")
    If strPar = "3" Then strPar = "0"
    optPurMed(Val(strPar)).value = True
    
    '------------------------------------------------------------------------------------------------------------------------
    '����ѡ��
    optSend(Val(zlDatabase.GetPara("���͵�������", glngSys, p����ҽ���´�, , Array(optSend(0), optSend(1), optSend(2)), blnSetup))).value = True
        
    '����Լ��λ���˷���Ϊ���ʵ�
    chk��λ����.value = Val(zlDatabase.GetPara("��λ����", glngSys, p����ҽ���´�, , Array(chk��λ����), blnSetup))
    
    'Ҫ��ǼǴ�����
    chkMustAddAgent.value = Val(zlDatabase.GetPara("Ҫ��ǼǴ�����", glngSys, p����ҽ���´�, "1", Array(chkMustAddAgent), blnSetup))
    
    '����ִ���Զ����
    chkִ��.value = Val(zlDatabase.GetPara("���ﱾ���Զ�ִ��", glngSys, p����ҽ���´�, , Array(chkִ��), blnSetup))
    
    '�ر�ҽ������
    chk�ر�ҽ��.value = Val(zlDatabase.GetPara("������ɺ�ر�ҽ������", glngSys, p����ҽ���´�, , Array(chk�ر�ҽ��), blnSetup))
    
    'ָ������ӡ
    optPoint(Val(zlDatabase.GetPara("ָ������ӡ��ʽ", glngSys, p����ҽ���´�, , Array(optPoint(0), optPoint(1), optPoint(2)), blnSetup))).value = True
    
    '���֧��
    chkSendPay.value = Val(zlDatabase.GetPara("�������֧��", glngSys, p����ҽ���´�, , Array(chkSendPay), blnSetup))
    '���֧������Ҫ���÷�ҩ����
    If chkSendPay.value = 0 Then
        vsfDrugStore.ColHidden(vsfDrugStore.ColIndex("��ҩ����")) = True
        vsfDrugStore.ColWidth(vsfDrugStore.ColIndex("ҩ��")) = vsfDrugStore.ColWidth(vsfDrugStore.ColIndex("ҩ��")) + vsfDrugStore.ColWidth(vsfDrugStore.ColIndex("��ҩ����"))
    End If
    
    cmdPBPSet.Enabled = chkSendPay.value
    
    If mblnҽ��վ = False Then
            
        '�Զ�����Ƥ��
        chkƤ��.value = Val(zlDatabase.GetPara("�Զ�����Ƥ��", glngSys, p����ҽ���´�, , Array(chkƤ��), blnSetup))
                        
        '���ݴ�ӡ:0-����ӡ,1-�ֹ���ӡ,2-�Զ���ӡ
        optPrint(Val(zlDatabase.GetPara("���﷢�͵��ݴ�ӡ", glngSys, p����ҽ���´�, , Array(optPrint(0), optPrint(1), optPrint(2)), blnSetup))).value = True
        
        'Ҫ�������������
        strPar = zlDatabase.GetPara("Ҫ�������������", glngSys, p����ҽ���´�, , Array(chk���, lst���), blnSetup)
        If Not chk���.Enabled Then lst���.Tag = "1" '�̶���ʶΪ������
        If strPar <> "" Then
            chk���.value = 1
            Call chk���_Click
        End If
        strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('4','5','6','7','8','9') Union ALL Select '5','ҩƷ' From Dual Order by ����"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        With lst���
            Do While Not rsTmp.EOF
                .AddItem rsTmp!���� & "-" & rsTmp!����
                .ItemData(.NewIndex) = Asc(rsTmp!����)
                
                If strPar <> "" Then
                    If InStr(strPar, Chr(.ItemData(.NewIndex))) > 0 Then
                        .Selected(.NewIndex) = True
                    End If
                End If
                rsTmp.MoveNext
            Loop
            .ListIndex = 0
        End With
        cmdBloodTip.Tag = zlDatabase.GetPara("��Ѫ����ע������", glngSys, p����ҽ���´�, , Array(cmdBloodTip), blnSetup)
    Else
        strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('4','5','6','7','8','9') Order by ����"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        
    End If
    
    '�ؼ�bug�����30������ʾ�������У����˺�߶�690��Ȼû�б䣩
    lstSendNO.Height = lstSendNO.Height + 30
    
    '��ͬ��ϵ�ҽ���ֱ��������
    chkNOType.value = Val(zlDatabase.GetPara("��ͬ��ϵ�ҽ���ֱ��������", glngSys, p����ҽ���´�, 0, Array(chkNOType), blnSetup))
    
    '��ʼʱ�䲻��ͬһ��ķֱ��������
    chkTimeDef.value = Val(zlDatabase.GetPara("��ʼʱ�䲻��ͬһ��ķֱ��������", glngSys, p����ҽ���´�, 0, Array(chkTimeDef), blnSetup))
    '���͵��ݺ�
    i = Val(zlDatabase.GetPara("���͵��ݺŹ���", glngSys, p����ҽ���´�, , Array(optSendNO(0), optSendNO(1), optSendNO(2), lstSendNO), blnSetup)) '0-���,1-������2-����
    i = IIF(i = 0, 1, IIF(i = 2, 2, 0))
    optSendNO(i).value = True
    Call optSendNO_Click(i)
    
    chkһ����ҩ����.value = Val(zlDatabase.GetPara("һ����ҩ����Ϊһ��", glngSys, p����ҽ���´�, 1, Array(chkһ����ҩ����), blnSetup))
    
    'ִ�п�����ͬʱ����Ϊͬһ���ݵ�ҽ�����
    strPar = zlDatabase.GetPara("����Ϊͬһ���ݵ�ҽ�����", glngSys, p����ҽ���´�, , Array(lstSendNO), blnSetup)
    With lstSendNO
        If rsTmp.RecordCount > 0 Then rsTmp.Filter = "����<>'5'"
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & "-" & rsTmp!����
            .ItemData(.NewIndex) = Asc(rsTmp!����)
            
            If strPar <> "" Then
                If InStr(strPar, Chr(.ItemData(.NewIndex))) > 0 Then
                    .Selected(.NewIndex) = True
                End If
            End If
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    cmdCancel.Left = Me.Left + Me.Width - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
    mblnҽ��վ = False
End Sub

Private Sub optSend_Click(Index As Integer)
    chk��λ����.Enabled = Index <> 0
End Sub

Private Sub optSend_GotFocus(Index As Integer)
    tabPar.Tab = 1
End Sub

Private Sub optSendNO_Click(Index As Integer)
    lstSendNO.Enabled = optSendNO(1).value
    chkһ����ҩ����.Enabled = optSendNO(1).value
End Sub

Private Sub vsfDrugStore_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDrugStore.ColIndex("����") Then
        Call Set����ҩ��(Row, True)
    ElseIf Col = vsfDrugStore.ColIndex("����") Then
        Call Setȱʡҩ��
    End If
    If Col <> vsfDrugStore.ColIndex("��ҩ����") Then Cancel = True
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("ȱʡ")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("��ҩ����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    With vsfDrugStore
        If .MouseCol = .ColIndex("ȱʡ") Then
            Call Setȱʡҩ��
        ElseIf .MouseCol = .ColIndex("ҩ��") Then
            Call Set����ҩ��(.Row, True)
        ElseIf .MouseCol = .ColIndex("����") And .MouseRow = .FixedRows - 1 Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                Call Set����ҩ��(i)
            Next
        End If
    End With
End Sub

Private Sub vsfDrugStore_EnterCell()
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore
        If .Row > 0 Then
            If .Col = .ColIndex("��ҩ����") Then
                Set rsTmp = Read��ҩ����(.RowData(.Row))
                .ColComboList(.Col) = "�Զ�����|" & .BuildComboList(rsTmp, "����")
                .FocusRect = flexFocusSolid
            Else
                .FocusRect = flexFocusLight
            End If
        End If
    End With
End Sub

Private Sub vsfDrugStore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If vsfDrugStore.Col = vsfDrugStore.ColIndex("ȱʡ") Then
            Call Setȱʡҩ��
        End If
    End If
End Sub

Private Sub Setȱʡҩ��()
'���ܣ����õ�ǰ�е�ȱʡҩ����ͬʱ������ͬ���͵������е�ȱʡҩ��
    Dim i As Long
    
    With vsfDrugStore
        If Val("" & .Cell(flexcpData, .Row, .ColIndex("ȱʡ"))) = 0 Then  '�ò��������޸ĵ������
            If .TextMatrix(.Row, .ColIndex("ȱʡ")) = "��" Then
                .TextMatrix(.Row, .ColIndex("ȱʡ")) = ""
            Else
                '��û����Ȩ���޸Ŀ���ʱ�ҿ���Ϊ0��false)ʱ����������ȱʡ
                If Not (Val(.TextMatrix(.Row, .ColIndex("����"))) = 0 And Val("" & .Cell(flexcpData, .Row, .ColIndex("����"))) = 1) Then
                    'ͬ����������ȡ��ȱʡ
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(.Row, .ColIndex("���")) = .TextMatrix(i, .ColIndex("���")) Then
                            If .TextMatrix(i, .ColIndex("ȱʡ")) = "��" Then .TextMatrix(i, .ColIndex("ȱʡ")) = ""
                        End If
                    Next
                    .TextMatrix(.Row, .ColIndex("����")) = -1    '�Զ�����Ϊ����
                    .TextMatrix(.Row, .ColIndex("ȱʡ")) = "��"
                Else
                    MsgBox "���õ�ǰҩ��Ϊȱʡʱ����ͬʱ����ǰҩ������Ϊ���ã�" & vbNewLine & "��û���޸Ŀ���ҩ����Ȩ�ޡ�", vbInformation, gstrSysName
                End If
            End If
        Else
            MsgBox "��û���޸�ȱʡҩ����Ȩ�ޡ�", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Set����ҩ��(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = False)
'���ܣ����õ�ǰ�еĿ���ҩ����ͬʱ����ǰ�е�ȱʡҩ��

    With vsfDrugStore
        If Val("" & .Cell(flexcpData, lngRow, .ColIndex("����"))) = 0 Then   '�ò��������޸ĵ������
            If Val(.TextMatrix(lngRow, .ColIndex("����"))) = -1 Then
                '��ǰ���ҹ�ѡ����
                If Not (Val("" & .Cell(flexcpData, lngRow, .ColIndex("ȱʡ"))) = 1 And .TextMatrix(lngRow, .ColIndex("ȱʡ")) = "��") Then
                    .TextMatrix(lngRow, .ColIndex("����")) = 0
                    .TextMatrix(lngRow, .ColIndex("ȱʡ")) = ""
                Else
                    If blnAsk Then
                        MsgBox "ȡ����ǰҩ������ʱ����ͬʱȡ����ǰҩ��ȱʡ��" & vbNewLine & "��û���޸�ȱʡҩ����Ȩ�ޡ�", vbInformation, gstrSysName
                    End If
                End If
            Else
                .TextMatrix(lngRow, .ColIndex("����")) = -1    '�Զ�����Ϊ����
            End If
        Else
            If blnAsk Then
                MsgBox "��û���޸Ŀ���ҩ����Ȩ�ޡ�", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

