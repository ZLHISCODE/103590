VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmInterfaceManager 
   BackColor       =   &H80000005&
   Caption         =   "�����ӿ���Ȩ����"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmInterfaceManager.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   0
      Left            =   120
      ScaleHeight     =   5775
      ScaleWidth      =   8895
      TabIndex        =   3
      Top             =   2400
      Width           =   8895
      Begin VB.CommandButton cmdEdit 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   350
         Index           =   2
         Left            =   7440
         TabIndex        =   8
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�༭(&E)"
         Enabled         =   0   'False
         Height          =   350
         Index           =   1
         Left            =   6270
         TabIndex        =   7
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "����(&N)"
         Height          =   350
         Index           =   0
         Left            =   5100
         TabIndex        =   6
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdStateChange 
         Caption         =   "ͣ��(&S)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2640
         TabIndex        =   5
         Top             =   5280
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   4935
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   8535
         _cx             =   15055
         _cy             =   8705
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInterfaceManager.frx":04F9
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
         ExplorerBar     =   5
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
   Begin VB.CommandButton cmdInterface 
      Caption         =   "�޸Ľӿ��û�����(&M)"
      Height          =   350
      Left            =   8280
      TabIndex        =   1
      Top             =   1510
      Width           =   2055
   End
   Begin XtremeSuiteControls.TabControl tbcMain 
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   9135
      _Version        =   589884
      _ExtentX        =   16113
      _ExtentY        =   10821
      _StockProps     =   64
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   1
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   8895
      TabIndex        =   9
      Top             =   2280
      Width           =   8895
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   5055
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   6855
         _cx             =   12091
         _cy             =   8916
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmInterfaceManager.frx":05FD
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
         ExplorerBar     =   5
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
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   14
         Top             =   480
         Width           =   1770
      End
      Begin VB.CommandButton cmdRepaire 
         Caption         =   "����Ȩ������(&A)"
         Height          =   350
         Left            =   7080
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblSearch 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7080
         TabIndex        =   13
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   120
      Picture         =   "frmInterfaceManager.frx":068A
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblExplain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmInterfaceManager.frx":0C39
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   840
      TabIndex        =   12
      Top             =   600
      Width           =   9660
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ӿ���Ȩ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   1920
   End
End
Attribute VB_Name = "frmInterfaceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Tab_Index
    TI_��Ȩ���� = 0
    TI_����Ȩ�� = 1
End Enum

Private Enum EditMode
    EM_New = 0
    EM_Modi = 1
    EM_Del = 2
End Enum

Private Enum AppGrant
    AG_�к� = 0
    AG_�ӿ����� = 1
    AG_��Ȩ�� = 2
    AG_��Чʱ�� = 3
    AG_ʧЧʱ�� = 4
    AG_״̬ = 5
    AG_��Ȩ˵�� = 6
End Enum

Private mlngCurPos  As Long         '��ǰ����λ��
Private mstrFind    As String       '�����ַ���
Private mblnReturn  As Boolean      '�Ƿ����˻س�
'===========================================================================
'==�����ӿ�
'===========================================================================
Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
End Sub
'===========================================================================
'==�¼�
'===========================================================================
Private Sub cmdEdit_Click(Index As Integer)
    Dim lngAPPNo    As Long
    Dim strRemarks  As String
    
    On Error GoTo errH
    If Index <> EM_New Then
        If vsfMain(TI_��Ȩ����).Row < vsfMain(TI_��Ȩ����).FixedRows Then
            Exit Sub
        End If
        lngAPPNo = vsfMain(TI_��Ȩ����).RowData(vsfMain(TI_��Ȩ����).Row)
    End If
    
    If cmdInterface.Caption <> "�ӿ���������((&M)" Then
        MsgBox "����δ����ZLInterface��ZLInterface�˻���������,����" & Mid(cmdInterface.Caption, 1, InStr(cmdInterface.Caption, "(") - 1) & "��", vbInformation, gstrSysName
        Exit Sub
    End If

    
    If Index = EM_Del Then
        If MsgBox("��ȷ��Ҫɾ��""" & vsfMain(TI_��Ȩ����).TextMatrix(vsfMain(TI_��Ȩ����).Row, AG_�ӿ�����) & """����Ȩ��Ϣ��", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            Exit Sub
        End If
        If Not CheckAuditStatus("0316", "�ӿ���Ȩ����", strRemarks) Then Exit Sub
        Call ExecuteProcedure("Zl_Zlinterface_Edit(1," & lngAPPNo & ")", Me.Caption, gcnOracle)
        Call SaveAuditLog(3, "�ӿ���Ȩ����", "ɾ���ӿ���Ȩ��Ϣ""" & vsfMain(TI_��Ȩ����).TextMatrix(vsfMain(TI_��Ȩ����).Row, AG_�ӿ�����) & """(" & vsfMain(TI_��Ȩ����).TextMatrix(vsfMain(TI_��Ȩ����).Row, AG_��Ȩ��) & "):" & strRemarks)
    ElseIf Index = EM_New Then
        If Not frmInterfaceEdit.ShowMe(lngAPPNo) Then
            Exit Sub
        End If
    Else
        If Not frmInterfaceEdit.ShowMe(lngAPPNo) Then
            Exit Sub
        End If
    End If
    Call LoadData(TI_��Ȩ����, lngAPPNo)
    Exit Sub
errH:
    MsgBox err.Description, vbCritical, gstrSysName
    err.Clear
End Sub

Private Sub cmdInterface_Click()
    Dim strPass     As String
    Dim lngType     As Long
    Dim strError    As String
    If cmdInterface.Tag = "" Then
        If frmInterfaceUser.ShowMe(Mid(cmdInterface.Caption, 1, InStr(cmdInterface.Caption, "(") - 1)) Then
        End If
        Call CheckZLinterface
    Else
        lngType = Val(Mid(cmdInterface.Tag, 1, 1))
        strPass = Mid(cmdInterface.Tag, 3)
        If lngType = 3 Then
            If MsgBox("ZLInterface�û��������ϵͳ���ָ�ȱʡ���룬�Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        If Not RepairGeneralAccount(gcnOracle, "ZLINTERFACE", strPass, strError) Then
            MsgBox IIf(lngType = 0, "����", "�޸�") & "ZLInterface�û�ʧ�ܡ���Ϣ��" & strError, vbInformation, gstrSysName
        Else
            MsgBox IIf(lngType = 0, "����", "�޸�") & "ZLInterface�û��ɹ���", vbInformation, gstrSysName
        End If
        Call CheckZLinterface
    End If
End Sub

Private Sub cmdRepaire_Click()
    On Error GoTo errH
    If MsgBox("��ZLInterface����Ȩ��ȱʧʱ��ͨ���ù��ܿ����Զ���ȱʧ�Ķ�����в�����Ȩ��" & vbCrLf & _
            "(ZL_THIRD_����̺�����B_THIRDSERVICE�����ZL_ҵ����Ϣ�嵥_INSERT��ZL_MSG_TODO),�Ƿ�ȷ��ִ�У�", vbInformation + vbYesNo, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Call ExecuteProcedure("Zl_Granttointerface()", Me.Caption, gcnOracle)
    MsgBox "�޸�����Ȩ�޳ɹ���", vbInformation, gstrSysName
    Call LoadData(TI_����Ȩ��)
    Exit Sub
errH:
    MsgBox "�޸�����Ȩ��ʧ��,��Ϣ��" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub cmdStateChange_Click()
    Dim lngAPPNo    As Long
    
    On Error GoTo errH
    If vsfMain(TI_��Ȩ����).Row < vsfMain(TI_��Ȩ����).FixedRows Then
        Exit Sub
    End If
    lngAPPNo = vsfMain(TI_��Ȩ����).RowData(vsfMain(TI_��Ȩ����).Row)
    If cmdInterface.Caption <> "�ӿ���������((&M)" Then
        MsgBox "����δ����ZLInterface��ZLInterface�˻���������,����" & Mid(cmdInterface.Caption, 1, InStr(cmdInterface.Caption, "(") - 1) & "��", vbInformation, gstrSysName
        Exit Sub
    End If
    Call ExecuteProcedure("Zl_Zlinterface_Edit(2," & lngAPPNo & ",NULL,NULL,NULL,NULL,NULL," & Val(cmdStateChange.Tag) & ")", Me.Caption, gcnOracle)
    Call LoadData(TI_��Ȩ����, lngAPPNo)
    Exit Sub
errH:
    MsgBox "�����Ȩ��״̬ʧ�ܣ���Ϣ��" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub Form_Load()
    '��ʼ������
    cmdStateChange.Tag = 1
    cmdStateChange.Caption = "ͣ��(&S)"
    tbcMain.Tag = "δ����"
    '��ʼ������
    tbcMain.InsertItem TI_��Ȩ����, "��Ȩ����", picMain(TI_��Ȩ����).hwnd, 0
    tbcMain.InsertItem TI_����Ȩ��, "����Ȩ��", picMain(TI_����Ȩ��).hwnd, 0
    tbcMain.Tag = ""
    Call CheckZLinterface
End Sub

Private Sub Form_Resize()
    Dim i       As Integer
    On Error Resume Next
    tbcMain.Move tbcMain.Left, tbcMain.Top, Me.ScaleWidth - tbcMain.Left - 60, Me.ScaleHeight - tbcMain.Top
    For i = TI_��Ȩ���� To TI_����Ȩ��
        picMain(i).Move picMain(i).Left, picMain(i).Top + 60, tbcMain.Width - 60, tbcMain.Height - picMain(i).Top
    Next
    cmdInterface.Left = Me.ScaleWidth - 120 - cmdInterface.Width
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub picMain_Resize(Index As Integer)
    On Error Resume Next
    vsfMain(Index).Left = 0
    vsfMain(Index).Top = 30
    vsfMain(Index).Height = picMain(Index).ScaleHeight - vsfMain(Index).Top - 500
    If Index = TI_����Ȩ�� Then
        cmdRepaire.Left = picMain(Index).ScaleWidth - cmdRepaire.Width - 120
        vsfMain(Index).Width = cmdRepaire.Left - vsfMain(Index).Left - 150
        txtSearch.Left = cmdRepaire.Left
        lblSearch.Left = cmdRepaire.Left
    Else
        vsfMain(Index).Width = picMain(Index).ScaleWidth - vsfMain(Index).Left - 120
        cmdEdit(EM_Del).Top = vsfMain(Index).Height + vsfMain(Index).Top + 60
        cmdEdit(EM_Del).Left = vsfMain(Index).Left + vsfMain(Index).Width - cmdEdit(EM_Del).Width
        Call SetCtrlPosOnLine(False, 0, cmdEdit(EM_Del), -1 * (cmdEdit(EM_Del).Width + cmdEdit(EM_Modi).Width + 30), cmdEdit(EM_Modi), -1 * (cmdEdit(EM_Modi).Width + cmdEdit(EM_New).Width + 30), cmdEdit(EM_New), -1 * (cmdEdit(EM_New).Width + cmdStateChange.Width + 240), cmdStateChange)
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call LoadData(Item.Index)
End Sub

Private Sub txtSearch_Change()
    mlngCurPos = 0
    mblnReturn = False
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        mblnReturn = True
        mstrFind = txtSearch.Text
        If Not mblnReturn Then
            mlngCurPos = 0
            mlngCurPos = FindItem(mlngCurPos)
        Else
            mlngCurPos = FindItem(mlngCurPos)
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub vsfMain_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnEditable As Boolean, blnStop     As Boolean
    If vsfMain(Index).Redraw = flexRDNone Then Exit Sub
    If Index = TI_��Ȩ���� Then
        If NewRow >= vsfMain(TI_��Ȩ����).FixedRows Then
            blnEditable = vsfMain(TI_��Ȩ����).RowData(NewRow) <> 0
            blnStop = vsfMain(TI_��Ȩ����).TextMatrix(NewRow, AG_״̬) = "ͣ��"
        End If
        cmdEdit(EM_Del).Enabled = blnEditable
        cmdEdit(EM_Modi).Enabled = blnEditable
        cmdStateChange.Enabled = blnEditable
        cmdStateChange.Tag = IIf(blnEditable And Not blnStop, 1, 0)
        cmdStateChange.Caption = IIf(blnEditable And Not blnStop, "ͣ��(&S)", "����(&S)")
    End If
End Sub
'===========================================================================
'==˽�з���
'===========================================================================
Private Sub LoadData(ByVal tiCur As Tab_Index, Optional ByVal lngNo As Long = -1)
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    Dim i       As Long, lngRow     As Long
    Dim lngCurRow   As Long
    On Error GoTo errH
    If tiCur = TI_��Ȩ���� Then
        '����һһ��Ӧ�������ĵ�һ���ֶδ���Rowdata
        strSQL = "Select Appno NO, Appname, Key, To_Char(Starttime, 'YYYY-MM-DD hh24:mi:ss') Starttime," & vbNewLine & _
                "       To_Char(Stoptime, 'YYYY-MM-DD hh24:mi:ss') Stoptime," & vbNewLine & _
                "       Decode(State, 0, Decode(Sign(Nvl(Stoptime, Starttime + 2) - Starttime), 1, '����', '����'), 1, 'ͣ��') State, Note," & vbNewLine & _
                "       Appno" & vbNewLine & _
                "From Zlinterface" & vbNewLine & _
                "Order By Appno"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    Else
        '����һһ��Ӧ�������ĵ�һ���ֶδ���Rowdata
        strSQL = "Select Rownum NO, Table_Schema, Table_Name, Privilege, Rownum ID" & vbNewLine & _
                "From (Select Rownum Rn, a.Table_Schema, a.Table_Name, a.Privilege, Rownum Rn" & vbNewLine & _
                "       From All_Tab_Privs A" & vbNewLine & _
                "       Where a.Grantee = 'ZLINTERFACE'" & vbNewLine & _
                "       Order By a.Table_Schema, a.Table_Name)"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    End If
    
    With vsfMain(tiCur)
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rsTmp.RecordCount + 1
        For lngRow = 1 To rsTmp.RecordCount
            For i = .FixedCols To .Cols - 1
                If tiCur = TI_��Ȩ���� And i = AG_��Ȩ�� Then
                    .TextMatrix(lngRow, i) = Sm4DecryptEcb(rsTmp.Fields(i).value & "", GetGeneralAccountKey(G_APP_KEY))
                Else
                    .TextMatrix(lngRow, i) = rsTmp.Fields(i).value & ""
                End If
            Next
            If lngNo = Val(rsTmp.Fields(.Cols).value & "") Then
                lngCurRow = lngRow
            End If
            .RowData(lngRow) = Val(rsTmp.Fields(.Cols).value & "")
            rsTmp.MoveNext
        Next
        If .Rows > .FixedRows And lngCurRow = 0 Then
            If lngNo = 0 Then
                lngCurRow = .Rows - 1
            Else
                lngCurRow = .FixedRows
            End If
        End If
        .Row = lngCurRow
        .Redraw = flexRDDirect
        Call vsfMain_AfterRowColChange(Val(tiCur), -1, -1, lngCurRow, .FixedCols)
    End With
    Exit Sub
errH:
    MsgBox "��������ʧ�ܣ���Ϣ��" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

'************************************************************************************************************
'����:��齨ZLinterface�û�
'����:�Ƿ����ZLInterFace�û���
'************************************************************************************************************
Private Sub CheckZLinterface()
    Dim strSQL  As String, strErr       As String
    Dim rsTmp   As ADODB.Recordset
    Dim strTmp  As String
    Dim connTmp As New ADODB.Connection
    On Error GoTo errH
    strSQL = "Select 1 From All_Users A Where a.Username = 'ZLINTERFACE'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    strTmp = GetZLInterfacePWD
    If rsTmp.RecordCount = 0 Then
        If strTmp <> "" Then
            cmdInterface.Caption = "�޸��ӿ��û�(&F)"
            cmdInterface.Tag = "2:" & strTmp
        Else
            cmdInterface.Caption = "�����ӿ��û�(&C)"
            cmdInterface.Tag = "0:ZL2018Soft."
        End If
    Else
        If strTmp = "" Then
            cmdInterface.Caption = "�޸��ӿ��û�(&F)"
            cmdInterface.Tag = "1:" & "ZL2018Soft."
        Else
            Set connTmp = gobjRegister.GetConnection(gstrServer, "ZLINTERFACE", strTmp, False, MSODBC, strErr, False)
            If connTmp.State = adStateClosed Then
                cmdInterface.Caption = "�޸��ӿ��û�(&F)"
                If InStr(strErr, "ORA-01017") > 0 Then
                    cmdInterface.Tag = "3:ZL2018Soft."
                Else
                    cmdInterface.Tag = "2:" & strTmp
                End If
            Else
                Call connTmp.Close
                Set connTmp = Nothing
                cmdInterface.Caption = "�ӿ���������((&M)"
                cmdInterface.Tag = ""
            End If
        End If
    End If
    Exit Sub
errH:
    MsgBox "���ZLinterfaceʧ�ܣ���Ϣ��" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

'************************************************************************************************************
'����:��ȡZLInterface�û�����
'************************************************************************************************************
Private Function GetZLInterfacePWD() As String
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset

    On Error GoTo errH
    strSQL = "Select Max(����) ���� From zlRegInfo A Where a.��Ŀ = '�����ӿ�����'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    If Trim(rsTmp!���� & "") <> "" Then
        GetZLInterfacePWD = Sm4DecryptEcb(rsTmp!���� & "", GetGeneralAccountKey(G_INTERFACE_KEY))
    End If
    Exit Function
errH:
    MsgBox "��ȡZLInterface����ʧ�ܡ���Ϣ��" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Private Function FindItem(Optional ByVal intCurPosition As Long) As Long
'���ܣ�����ģ�����
'������intCurPosition=��ǰλ�ã�<=1��ʾ��ͷ��β��ʼ���ң�����ӵ�ǰλ�ÿ�ʼ����
'���أ�ƥ����Ŀλ��
    Dim i As Integer
    Dim blnFind As Boolean
    Dim strLike As String
    Dim strMsg As String
    
    On Error Resume Next
    If intCurPosition < 0 Then FindItem = -1: Exit Function
    
    '�����ַ�������
    strLike = "*" & UCase(mstrFind) & "*"
    '���в���
    
    For i = intCurPosition + 1 To vsfMain(TI_����Ȩ��).Rows - 1
        If vsfMain(1).TextMatrix(i, vsfMain(TI_����Ȩ��).ColIndex("����")) Like strLike Then
            blnFind = True
            Exit For
        End If
    Next
    'δ���ҵ�ԭ����ʾ
    If Not blnFind Then
        If mlngCurPos <= 0 Then
            MsgBox "δ�ҵ�ƥ����Ѿ���Ȩ����", vbInformation, Me.Caption
            FindItem = -1
            '��ʾ�Ƿ��ͷ��ʼ����
        Else
            If MsgBox("δ�ҵ�ƥ����Ѿ���Ȩ�����Ƿ����½��в���", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                mlngCurPos = 0
                mlngCurPos = FindItem(mlngCurPos)
                FindItem = mlngCurPos
            Else
                FindItem = -1
            End If
        End If
    Else
        FindItem = i
        vsfMain(TI_����Ȩ��).Select i, vsfMain(TI_����Ȩ��).ColIndex("�к�")
        vsfMain(TI_����Ȩ��).ShowCell i, vsfMain(TI_����Ȩ��).ColIndex("�к�")
        Call vsfMain_AfterRowColChange(TI_����Ȩ��, -1, -1, i, vsfMain(TI_����Ȩ��).ColIndex("�к�"))
    End If
End Function
