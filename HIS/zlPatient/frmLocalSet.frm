VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocalSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab sTab 
      Height          =   4755
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   8387
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "�������(&1)"
      TabPicture(0)   =   "frmLocalSet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkAutoRefresh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ҽ�ƿ�Ʊ�ݿ���(&2)"
      TabPicture(1)   =   "frmLocalSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboType"
      Tab(1).Control(1)=   "chkBrushCardVerfy"
      Tab(1).Control(2)=   "chkBruhCardBackCard"
      Tab(1).Control(3)=   "fraTitle"
      Tab(1).Control(4)=   "cmdDeviceSetup(0)"
      Tab(1).Control(5)=   "img16"
      Tab(1).Control(6)=   "lblDefaultPayCard"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Ԥ��������(&3)"
      TabPicture(2)   =   "frmLocalSet.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPrepay"
      Tab(2).Control(1)=   "cmdDeviceSetup(1)"
      Tab(2).Control(2)=   "chkLedWelcome"
      Tab(2).Control(3)=   "cboDefaultBalance"
      Tab(2).Control(4)=   "lblEdit"
      Tab(2).ControlCount=   5
      Begin VB.Frame fraPrepay 
         Caption         =   "���ع���Ʊ��"
         Height          =   3315
         Left            =   -74925
         TabIndex        =   11
         Top             =   525
         Width           =   5865
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   2925
            Left            =   90
            TabIndex        =   12
            Top             =   270
            Width           =   5670
            _cx             =   10001
            _cy             =   5159
            Appearance      =   0
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmLocalSet.frx":0054
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
            ExplorerBar     =   2
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
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Index           =   1
         Left            =   -70560
         TabIndex        =   16
         Top             =   4260
         Width           =   1500
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED��ʾ��ӭ��Ϣ"
         Height          =   225
         Left            =   -74835
         TabIndex        =   13
         ToolTipText     =   "�շѴ������벡�˺�,�Ƿ���ʾ��ӭ��Ϣ������"
         Top             =   4020
         Value           =   1  'Checked
         Width           =   1710
      End
      Begin VB.CheckBox chkAutoRefresh 
         Caption         =   "�л���������ѡ�ʱ���Զ�ˢ�²�������"
         Height          =   180
         Left            =   285
         TabIndex        =   3
         Top             =   555
         Width           =   3840
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   -73665
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4050
         Width           =   2580
      End
      Begin VB.CheckBox chkBrushCardVerfy 
         Caption         =   "�˿���ȡ���ݺź�ˢ����֤�˿�"
         Height          =   180
         Left            =   -74835
         TabIndex        =   6
         Top             =   3540
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.CheckBox chkBruhCardBackCard 
         Caption         =   "���������ˡ�ˢ���˿�"
         Height          =   240
         Left            =   -74835
         TabIndex        =   7
         Top             =   3795
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.ComboBox cboDefaultBalance 
         Height          =   300
         Left            =   -73725
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4290
         Width           =   1875
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ع���..."
         Height          =   2880
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   5745
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2445
            Left            =   60
            TabIndex        =   5
            Top             =   300
            Width           =   5595
            _cx             =   9869
            _cy             =   4313
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmLocalSet.frx":0131
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
            ExplorerBar     =   2
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
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Index           =   0
         Left            =   -70605
         TabIndex        =   10
         Top             =   4020
         Width           =   1500
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   -71145
         Top             =   855
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLocalSet.frx":0212
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblDefaultPayCard 
         Caption         =   "ȱʡ��������"
         Height          =   210
         Left            =   -74835
         TabIndex        =   8
         Top             =   4095
         Width           =   1290
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡ���㷽ʽ"
         Height          =   180
         Left            =   -74850
         TabIndex        =   14
         Top             =   4350
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6240
      TabIndex        =   2
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6240
      TabIndex        =   1
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   0
      Top             =   360
      Width           =   1100
   End
End
Attribute VB_Name = "frmLocalSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlngModul As Long, mstrPrivs As String, mbln���� As Boolean
Private mstrClass As String, mstrDeposit As String
Private mblnOK As Boolean
Public Function zlSetPara(ByVal frmMain As Object, ByVal strPrivs As String, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:mlngModul-1101-������Ϣ����,1102-���￨����,1103-Ԥ�������
    '����:
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 14:22:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False: mstrPrivs = strPrivs: mlngModul = lngModule
    mbln���� = InStr(mstrPrivs, ";������Ϣ;") > 0 And mlngModul = 1101
    Me.Show 1, frmMain
    zlSetPara = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDeviceSetup_Click(Index As Integer)
    Call zlCommFun.DeviceSetup(Me, 100, mlngModul)
End Sub

Private Sub cmdHelp_Click()
    Select Case mlngModul
        Case 1101 '������Ϣ
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet1"
        Case 1102 '���￨
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet2"
        Case 1103 'Ԥ����
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet3"
    End Select
End Sub

Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str��� As String
    IsValied = False
    
    On Error GoTo errHandle
    If mlngModul <> 1103 Then
        '���ÿ��ʹ����ʽֻ��һ��ѡ��
        With vsBill
            str��� = "-"
            For i = 1 To vsBill.Rows - 1
                If str��� <> Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����"))) Then
                   str��� = Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����")))
                   lngSelCount = 0
                    For j = 1 To vsBill.Rows - 1
                        If Trim(.TextMatrix(i, .ColIndex("ҽ�ƿ����"))) = Trim(.TextMatrix(j, .ColIndex("ҽ�ƿ����"))) Then
                            If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                                lngSelCount = lngSelCount + 1
                            End If
                        End If
                    Next
                    If lngSelCount > 1 Then
                        MsgBox "ע��:" & vbCrLf & "    ҽ�ƿ����Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                        Exit Function
                    End If
                End If
            Next
        End With
    End If
    If mlngModul = 1102 Then IsValied = True: Exit Function
  '���ÿ��ʹ��Ԥ��ֻ��һ��ѡ��
    With vsPrepay
        str��� = "-"
        For i = 1 To .Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("Ԥ������"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("Ԥ������")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("Ԥ������"))) = Trim(.TextMatrix(j, .ColIndex("Ԥ������"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    Ԥ������Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ʊ��
    '����:���˺�
    '����:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    If mlngModul <> 1103 Then
        '���湲��Ʊ��
        strValue = ""
        With vsBill
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                    strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("ҽ�ƿ����")))
                End If
            Next
        End With
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "����ҽ�ƿ�����", strValue, glngSys, mlngModul, blnHavePrivs
    End If
    If mlngModul = 1102 Then Exit Sub
    
    
    '����Ԥ��Ʊ��
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Val(.Cell(flexcpData, i, .ColIndex("Ԥ������")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "����Ԥ��Ʊ������", strValue, glngSys, mlngModul, blnHavePrivs
End Sub
Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-07-06 18:41:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intTYPE As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String, rsҽ�ƿ���� As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim strȱʡҽ�ƿ� As String, lngȱʡҽ�ƿ� As Long
    Dim strBillFormat As String
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    On Error GoTo errHandle
    '�ָ��п��
    If mlngModul <> 1103 Then
            lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, , , True, intTYPE))
            '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
            gstrSQL = "Select ID,����,����, nvl(�Ƿ�̶�,0) as �Ƿ�̶�  from ҽ�ƿ����  Where nvl(�Ƿ�����,0)=1 And nvl(�Ƿ�֤��,0)=0 "
            Set rsҽ�ƿ���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            rsҽ�ƿ����.Filter = "����='���￨' and �Ƿ�̶�=1"
            If rsҽ�ƿ����.EOF = False Then
                strȱʡҽ�ƿ� = rsҽ�ƿ����!����: lngȱʡҽ�ƿ� = Val(rsҽ�ƿ����!ID)
            End If
            With rsҽ�ƿ����
                cboType.Clear
                rsҽ�ƿ����.Filter = 0
                If rsҽ�ƿ����.RecordCount <> 0 Then rsҽ�ƿ����.MoveFirst
                Do While Not .EOF
                    cboType.AddItem NVL(!����)
                    cboType.ItemData(cboType.NewIndex) = NVL(!ID)
                    If NVL(!����) = "���￨" Then cboType.ListIndex = cboType.NewIndex
                    If lngCardTypeID = Val(NVL(!ID)) Then
                        cboType.ListIndex = cboType.NewIndex
                    End If
                    .MoveNext
                Loop
            End With
            
            zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "����ҽ��Ʊ���б�", False, False
            strShareInvoice = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModul, , , True, intTYPE)
            '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
            vsBill.Tag = ""
            Select Case intTYPE
            Case 1, 3, 5, 15
                vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
                fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
                If intTYPE = 5 Then vsBill.Tag = ""
            Case Else
                vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
                fraTitle.ForeColor = &H80000008
            End Select
            With vsBill
                .Editable = flexEDKbdMouse
                If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
            End With
            
            '��ʽ:����ID1,ҽ�ƿ����ID1|����IDn,ҽ�ƿ����IDn|...
            varData = Split(strShareInvoice, "|")
    
            '1.���ù���Ʊ��
            Set rsTemp = GetShareInvoiceGroupID(5)
            With vsBill
                .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
                lngRow = 1
                .MergeCells = flexMergeRestrictRows
                .MergeCellsFixed = flexMergeFixedOnly
                .MergeCol(0) = True
                Do While Not rsTemp.EOF
                    .RowData(lngRow) = Val(NVL(rsTemp!ID))
                    '105985:���ϴ�,2017/4/10,��ҽ�ƿ���������Ʊ��
                    If Val(NVL(rsTemp!ʹ�����ID)) = 0 Then
                        .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = strȱʡҽ�ƿ�
                        .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = lngȱʡҽ�ƿ�
                    Else
                        rsҽ�ƿ����.Filter = "ID=" & Val(NVL(rsTemp!ʹ�����ID))
                        If Not rsҽ�ƿ����.EOF Then
                            .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = NVL(rsҽ�ƿ����!����)
                        Else
                            .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = NVL(rsTemp!ʹ�����)
                        End If
                        .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = Val(NVL(rsTemp!ʹ�����ID))
                    End If
                    .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
                    .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
                    .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
                    .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(NVL(rsTemp!ʣ������)), "##0;-##0;;")
                    For i = 0 To UBound(varData)
                        varTemp = Split(varData(i) & ",", ",")
                        lngTemp = Val(varTemp(0))
                        If Val(.RowData(lngRow)) = lngTemp _
                            And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����"))) Then
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                        End If
                    Next
                    .MergeRow(lngRow) = True
                    lngRow = lngRow + 1
                    rsTemp.MoveNext
                Loop
            End With
    End If
    If mlngModul = 1102 Then Exit Sub
    '����Ԥ��Ʊ������
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
    
    strShareInvoice = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, mlngModul, , , True, intTYPE)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intTYPE
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intTYPE = 5 Then vsBill.Tag = ""
    Case Else
        vsPrepay.ForeColor = &H80000008: vsPrepay.ForeColorFixed = &H80000008
        fraPrepay.ForeColor = &H80000008
    End Select
    With vsPrepay
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,Ԥ�����ID1|����IDn,Ԥ�����IDn|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsPrepay
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            '58071
            Select Case Val(NVL(rsTemp!ʹ�����, ""))
            Case 0 '�����������סԺƱ��
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = ""
                .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = 0
            Case 1  '����Ʊ��
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ������Ʊ��"
                .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = 1
            Case Else   'סԺƱ��
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ��סԺƱ��"
                .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = 2
            End Select
            
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(NVL(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ԥ������"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
 
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, strTmp As String
    
    '���ع��þ��￨
    If IsValied = False Then Exit Sub
    Call SaveInvoice
    
    Select Case mlngModul
    Case 1101 '������Ϣ
        '76824�����ϴ���2014/8/19��ҽ�ƿ������
        If cboType.ListIndex >= 0 Then
            zlDatabase.SetPara "ȱʡҽ�ƿ����", cboType.ItemData(cboType.ListIndex), glngSys, mlngModul, IIf(cboType.Enabled = True, True, False)
        Else
            zlDatabase.SetPara "ȱʡҽ�ƿ����", 0, glngSys, mlngModul, IIf(cboType.Enabled = True, True, False)
        End If
        '54701:������,2012-09-19
        zlDatabase.SetPara "�Զ�ˢ������", chkAutoRefresh.Value, glngSys, mlngModul, IIf(chkAutoRefresh.Enabled = True, True, False)
    Case 1102   '���￨
        '����28130��27929
        If chkBruhCardBackCard.Value And chkBrushCardVerfy.Value Then
            strTmp = "3"
        ElseIf chkBruhCardBackCard.Value Then
            strTmp = "1"
        ElseIf chkBrushCardVerfy.Value Then
            strTmp = "2"
        Else
            strTmp = "0"
        End If
        Call zlDatabase.SetPara("�˿�ˢ��", strTmp, glngSys, mlngModul, IIf(chkBruhCardBackCard.Enabled = True, True, False))
    Case 1103
        zlDatabase.SetPara "ȱʡԤ�����㷽ʽ", Trim(cboDefaultBalance.Text), glngSys, glngModul, IIf(cboDefaultBalance.Enabled = True, True, False)
    End Select
    'LED�豸
    zlDatabase.SetPara "LED��ʾ��ӭ��Ϣ", chkLedWelcome.Value, glngSys, mlngModul, IIf(chkLedWelcome.Enabled = True, True, False)

    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

 Private Sub LoadȱʡԤ�����㷽ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ش��տ�
    '����:���˺�
    '����:2011-07-19 15:13:59
    '����:  34705
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, varData As Variant, varTemp As Variant, j As Long, strTmp As String
    
    str���㷽ʽ = zlDatabase.GetPara("ȱʡԤ�����㷽ʽ", glngSys, glngModul, , Array(cboDefaultBalance), InStr(mstrPrivs, ";��������;") > 0)
     
     On Error GoTo errHandle
    '���㷽ʽ
    strSQL = _
    " Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
    " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
    " Where A.Ӧ�ó���='Ԥ����' And B.����=A.���㷽ʽ And Nvl(B.����,1) In(1,2,3,5,8)" & _
    " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboDefaultBalance
        Do While Not rsTmp.EOF
            .AddItem NVL(rsTmp!����)
            If .ListIndex < 0 And Val(NVL(rsTmp!ȱʡ)) = 1 Then .ListIndex = .NewIndex
            If str���㷽ʽ = NVL(rsTmp!����) Then .ListIndex = .NewIndex
            rsTmp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
 End Sub

Private Sub Form_Load()
    Dim i As Long, lngCardTypeID As Long
    Dim strPrintMode As String '�����:50656
    Dim strArr��ӡ��ʽ() As String '�����:50656
    Dim strTmp As String
    gblnOK = False
    Me.sTab.TabVisible(2) = False   '34705
    sTab.TabVisible(0) = mlngModul = 1101
    sTab.TabVisible(2) = mlngModul = 1103    '34705
    sTab.TabVisible(1) = mlngModul <> 1103    '34705
    If mlngModul = 1103 Then Call LoadȱʡԤ�����㷽ʽ
    Call InitShareInvoice   '���ع�����Ʊ����Ϣ
    
    'LED�豸
    chkLedWelcome.Value = zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModul, 1, Array(chkLedWelcome), InStr(mstrPrivs, ";��������;") > 0)

    Select Case mlngModul
    Case 1101 ''������Ϣ
        lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, , Array(cboType), InStr(mstrPrivs, ";��������;") > 0))
        For i = 0 To cboType.ListCount - 1
            If cboType.ItemData(i) = lngCardTypeID Then cboType.ListIndex = i: Exit For
        Next
        
        '54701:������,2012-09-19
        chkAutoRefresh.Value = zlDatabase.GetPara("�Զ�ˢ������", glngSys, mlngModul, 1, Array(chkAutoRefresh), InStr(mstrPrivs, ";��������;") > 0)
    
    Case 1102   '���￨
        '����28130
        Select Case Val(zlDatabase.GetPara("�˿�ˢ��", glngSys, mlngModul, "0", Array(chkBruhCardBackCard, chkBrushCardVerfy), InStr(mstrPrivs, ";��������;") > 0))
        Case 0: chkBruhCardBackCard.Value = 0: chkBrushCardVerfy.Value = 0
        Case 1: chkBruhCardBackCard.Value = 1
        Case 2: chkBrushCardVerfy.Value = 1
        Case 3: chkBruhCardBackCard.Value = 1: chkBrushCardVerfy.Value = 1
        End Select
        chkBruhCardBackCard.Visible = True: chkBrushCardVerfy.Visible = True
    Case 1103  'Ԥ����
    End Select
    chkLedWelcome.Visible = mlngModul = 1103
    Exit Sub
errH:
    If ErrCenter() = 1 Then
         Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln���� = False
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����ҽ��Ʊ���б�", False, False
    zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub
 
Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub
Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("ҽ�ƿ����"))) = Trim(.Cell(flexcpData, i, .ColIndex("ҽ�ƿ����"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub
Private Sub vsPrepay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsPrepay
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("Ԥ������"))) = Trim(.Cell(flexcpData, i, .ColIndex("Ԥ������"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsPrepay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsPrepay
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub
