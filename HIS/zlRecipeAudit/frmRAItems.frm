VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRAItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������Ŀ"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10260
   Icon            =   "frmRAItems.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkPass 
      Caption         =   "������ҩ������"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   6480
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   8880
      TabIndex        =   16
      Top             =   7920
      Width           =   990
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   7680
      TabIndex        =   15
      Top             =   7920
      Width           =   990
   End
   Begin VB.Frame fraPass 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   9975
      Begin MSComctlLib.ListView lvwPass 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   1085
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CheckBox chkCheckInPat 
         Caption         =   "���סԺ"
         Height          =   180
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkCheckOutPat 
         Caption         =   "�������"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵�������¹�ѡ������PASS�����ͬʱ���Զ����Ϊ�����ϸ�"
         Height          =   180
         Left            =   4800
         TabIndex        =   13
         Top             =   240
         Width           =   5040
      End
   End
   Begin VB.Frame fraItems 
      Caption         =   "�����Ŀ"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   360
         Left            =   8760
         TabIndex        =   8
         Top             =   5760
         Width           =   990
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�༭(&E)"
         Height          =   360
         Left            =   7560
         TabIndex        =   7
         Top             =   5760
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   360
         Left            =   6360
         TabIndex        =   6
         Top             =   5760
         Width           =   990
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfItems 
         Height          =   4935
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   9735
         _cx             =   17171
         _cy             =   8705
         Appearance      =   2
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.CheckBox chkInPat 
         Caption         =   "סԺȫѡ(&I)"
         Height          =   180
         Left            =   8520
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkOutPat 
         Caption         =   "����ȫѡ(&O)"
         Height          =   180
         Left            =   7080
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optAccord 
         Caption         =   "���ݡ�������������淶��28�����"
         Height          =   180
         Index           =   1
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.OptionButton optAccord 
         Caption         =   "���ݡ���������취��7�����"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmRAItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private mblnParaAccord As Boolean       '�����ݣ�True������������淶28�False-��������취7��
Private mlngModule As Long
Private mstrPrivs As String
Private mintParaPass As Integer         '������ҩ����
Private mblnEnter As Boolean            '�Ƿ���ɳ�ʼ�����̣�True��ɣ�False���ڳ�ʼ������
Private mblnMemory As Boolean           'ʹ�ø��Ի����True���ã�Falseδ����
Private mrsItems As ADODB.Recordset

Private Const MSTR_VSF As String = "ID,,1,0|����,,1,0|�������,,3,900|���סԺ,,3,900|�������,,1,0|���,,3,1000|����,,3,1000|���,,3,1500|��������,,3,4000"

Private Sub Form_Load()
    Dim strPASS As String
    Dim arrVal As Variant
    Dim i As Integer
    Dim strTmp As String
    Dim limTmp As ListItem

    mblnEnter = False
    
    mlngModule = glngModule
    mstrPrivs = zlStr.FormatString(";[1];", GetPrivFunc(glngSys, mlngModule))

    '��ȡ����ֵ
    mblnMemory = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1
    mintParaPass = Val(zlDatabase.GetPara("������ҩ���ӿ�", glngSys))      '0-��ʾδʹ��,1-�����ӿ�,2-��ͨ�ӿ�,3-̫Ԫͨ�ӿ�,4-����

    '��ʼ���ؼ�
    If Val(zlDatabase.GetPara("�����������", glngSys)) = 2 Then
        optAccord(0).Value = True
        fraItems.Tag = "0"
    Else
        optAccord(1).Value = True
        fraItems.Tag = "1"
    End If
    
    InitVSF vsfItems
'    If mblnMemory Then
'        strTmp = GetSetting("ZLSOFT", FormatString("˽��ģ��\[1]\��������\[2]\[3]\[4]", UserInfo.�û���, App.ProductName, Me.Name), vsfItems.Name)
'        If strTmp = "" Then
'            strTmp = MSTR_VSF
'        Else
'            MergeVSFHead strTmp, MSTR_VSF, strTmp
'        End If
'    Else
'        strTmp = MSTR_VSF
'    End If
    SetVSFHead vsfItems, MSTR_VSF   'strTmp
    
    With vsfItems
        .ColDataType(.ColIndex("�������")) = flexDTBoolean
        .ColDataType(.ColIndex("���סԺ")) = flexDTBoolean
    End With
    
    '����PASS
    chkPass.Width = 2500
    Select Case mintParaPass
        Case 1
            chkPass.Caption = zlStr.FormatString("������ҩ�����飨[1]��", "����")
            strPASS = "4-��(�ϸ߶ȹ�ע)|3-��(���ع�ע)|2-��(�߶ȹ�ע)|1-��(�ʶȹ�ע)|0-��(����)"
        Case 2
            chkPass.Caption = zlStr.FormatString("������ҩ�����飨[1]��", "��ͨ")
            chkPass.Enabled = False
        Case 3
            chkPass.Caption = zlStr.FormatString("������ҩ�����飨[1]��", "̫Ԫͨ")
            chkPass.Width = 2650
            strPASS = "1-��(����)|2-��(����)|3-��(����)"
        Case 4
            chkPass.Caption = zlStr.FormatString("������ҩ�����飨[1]��", "����")
            strPASS = "3-��(����)|2-��(����)|1-��(����)"
        Case Else
            chkPass.Caption = "������ҩ������"
            chkPass.Width = 1800
            chkPass.Enabled = False
    End Select
    Call chkPass_Click
    
    On Error GoTo errHandle
    gstrSQL = "Select ID, Decode(���, 1, '1-7��', 2, '2-28��', 3, '3-�̶�', 4, '4-�Զ���') ���, ����, ���, ���� ��������, " & vbCr & _
              "  �Ƿ���������*-1 �������, �Ƿ�סԺ����*-1 ���סԺ, �������, PASS���, ������, ����ʱ�� " & vbCr & _
              "From ���������Ŀ " & vbCr & _
              "Where ����ʱ�� Is Null " & vbCr & _
              "Order By ���, ����"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ŀ����")
    
    '��ȡPASS��Ŀ
    mrsItems.Filter = zlStr.FormatString("���='[1]'", "3-�̶�")
    If mrsItems.RecordCount > 0 Then
        strTmp = zlCommFun.NVL(mrsItems!PASS���)
        lvwPass.Tag = CStr(mrsItems!ID)
        chkCheckOutPat.Value = Abs(Val(zlCommFun.NVL(mrsItems!�������)))
        chkCheckInPat.Value = Abs(Val(zlCommFun.NVL(mrsItems!���סԺ)))
        chkPass.Value = IIf(chkCheckOutPat.Value = 1 Or chkCheckInPat.Value = 1, 1, 0)
    End If
    
    lvwPass.ListItems.Clear
    If strPASS <> "" Then
        With lvwPass
            .View = lvwSmallIcon
            arrVal = Split(strPASS, "|")
            For i = LBound(arrVal) To UBound(arrVal)
                Set limTmp = .ListItems.Add(, "K_" & Val(arrVal(i)), arrVal(i))
                If strTmp <> "" Then
                    limTmp.Checked = IIf(InStr(";" & strTmp & ";", Left(arrVal(i), 1)) > 0, True, False)
                End If
            Next
        End With
    End If
    
    RestoreWinState Me, App.ProductName
    
    mblnEnter = True
    
    '��������
    mrsItems.Filter = zlStr.FormatString("���='[1]' or ���='4-�Զ���'", IIf(optAccord(0).Value, "1-7��", "2-28��"))     '1-7�2-28�3-�̶�1�4-�Զ���
    mdlDefine.FillVSFData vsfItems, mrsItems
    If mrsItems.RecordCount > 0 Then
        '�������������סԺ��Ԫ����ɫ
        Call SetVSFColor
        Call vsfItems_AfterRowColChange(0, 0, 1, 0)
    End If
    
    Exit Sub

errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub chkCheckInPat_Click()
    lvwPass.Enabled = chkCheckOutPat.Enabled And chkCheckOutPat.Value Or chkCheckInPat.Enabled And chkCheckInPat.Value
    If mblnEnter Then chkPass.Tag = "1"
End Sub

Private Sub chkCheckOutPat_Click()
    lvwPass.Enabled = chkCheckOutPat.Enabled And chkCheckOutPat.Value Or chkCheckInPat.Enabled And chkCheckInPat.Value
    If mblnEnter Then chkPass.Tag = "1"
End Sub

Private Sub chkInPat_Click()
    Dim lngRow As Long
    Dim strVal As String
    
    With vsfItems
        If .Rows < 2 Then Exit Sub
        .Redraw = False
        
        For lngRow = 1 To .Rows - 1
            strVal = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
            If strVal = "1" Or strVal = "2" Then .Cell(flexcpChecked, lngRow, .ColIndex("���סԺ")) = IIf(chkInPat.Value, "1", "")
        Next
        
        .Redraw = True
    End With
End Sub

Private Sub chkOutPat_Click()
    Dim lngRow As Long
    Dim strVal As String
    
    With vsfItems
        If .Rows < 2 Then Exit Sub
        .Redraw = False
        
        For lngRow = 1 To .Rows - 1
            strVal = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
            If strVal = "0" Or strVal = "2" Then .Cell(flexcpChecked, lngRow, .ColIndex("�������")) = IIf(chkOutPat.Value, "1", "")
        Next
        
        .Redraw = True
    End With
End Sub

Private Sub chkPass_Click()
    If mblnEnter Then chkPass.Tag = "1" '��ʾ���޸�
    
    chkCheckOutPat.Enabled = chkPass.Enabled And chkPass.Value
    chkCheckInPat.Enabled = chkPass.Enabled And chkPass.Value
    If chkPass.Value = 0 Then
        chkCheckOutPat.Value = 0
        chkCheckInPat.Value = 0
    End If
    
    Call chkCheckOutPat_Click
    Call chkCheckInPat_Click
End Sub

Private Sub InitVSF(ByVal vsfVar As VSFlexGrid)
'���ܣ���ʼ�������VSFlexGrid�ؼ��ķ��
'������
'  vsfVar��Ҫ��ʼ����VSFlexGrid�ؼ�

    With vsfVar
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
        .BackColorBkg = .BackColor
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub cmdAdd_Click()
    frmRAItemsEdit.ShowMe 1, 0, Me
End Sub

Private Sub cmdEdit_Click()
    frmRAItemsEdit.ShowMe 2, Val(vsfItems.TextMatrix(vsfItems.Row, vsfItems.ColIndex("ID"))), Me
End Sub

Private Sub cmdDel_Click()
    With vsfItems
        If MsgBox(zlStr.FormatString("�Ƿ�ȷ��ɾ����[1]�������Ŀ��", .TextMatrix(.Row, .ColIndex("���"))), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            If .TextMatrix(.Row, .ColIndex("����")) = "1" Then
                .RemoveItem .Row
            Else
                .RowHidden(.Row) = True
            End If
        End If
        .SetFocus
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    '����
    If Save() Then
        zlDatabase.SetPara "�����������", IIf(optAccord(0).Value, "2", "1"), glngSys
        Unload Me
    End If
End Sub

Private Function Save() As Boolean
'���ܣ���������
'���أ�True�ɹ���Falseʧ��

    Dim colSQL As New Collection
    Dim lngRow As Long
    Dim strSQL As String, strPASS As String
    
    With vsfItems
        If .Rows < 2 Then Exit Function
        
        For lngRow = 1 To .Rows - 1
            strSQL = "ZL_���������Ŀ_UPDATE("
            '1.ID
            strSQL = strSQL & Trim(.TextMatrix(lngRow, .ColIndex("ID"))) & ","
            '2.���
            strSQL = strSQL & Val(.TextMatrix(lngRow, .ColIndex("���"))) & ","
            '3.����
            strSQL = strSQL & zlStr.FormatString("'[1]',", .TextMatrix(lngRow, .ColIndex("����")))
            '4.���
            strSQL = strSQL & zlStr.FormatString("'[1]',", .TextMatrix(lngRow, .ColIndex("���")))
            '5.��������
            strSQL = strSQL & zlStr.FormatString("'[1]',", .TextMatrix(lngRow, .ColIndex("��������")))
            '6.�Ƿ���������
            strSQL = strSQL & IIf(Val(.TextMatrix(lngRow, .ColIndex("�������"))) = -1, "1", "0") & ","
            '7.�Ƿ�סԺ����
            strSQL = strSQL & IIf(Val(.TextMatrix(lngRow, .ColIndex("���סԺ"))) = -1, "1", "0") & ","
            '8.�������
            strSQL = strSQL & Trim(.TextMatrix(lngRow, .ColIndex("�������"))) & ","
            '9.PASS���
            strSQL = strSQL & "Null,"
            '10.������
            strSQL = strSQL & zlStr.FormatString("'[1]',", UserInfo.����)
            '11.�Ƿ�����
            If .RowHidden(lngRow) Then
                strSQL = strSQL & "1)"
            Else
                strSQL = strSQL & "Null)"
            End If
            
            'SQL���뼯�϶���
            AddArray colSQL, strSQL
            'Debug.Print strSQL
        Next
    End With
    
    If Val(chkPass.Tag) = 1 Then
        For lngRow = 1 To lvwPass.ListItems.Count
            If lvwPass.ListItems(lngRow).Checked Then
                strPASS = strPASS & Mid(lvwPass.ListItems(lngRow).Key, 3) & ";"
            End If
        Next
        If strPASS <> "" Then
            strPASS = Left(strPASS, Len(strPASS) - 1)
        End If
        
        strSQL = "ZL_���������Ŀ_UPDATE("
        '1.ID
        strSQL = strSQL & lvwPass.Tag & ","
        '2.��� = 3-PASS
        strSQL = strSQL & "3,"
        '3.���룻4.��ƣ�5.��������
        strSQL = strSQL & "Null,Null,Null,"
        '6.�Ƿ���������
        strSQL = strSQL & IIf(chkCheckOutPat.Value, "1", "0") & ","
        '7.�Ƿ�סԺ����
        strSQL = strSQL & IIf(chkCheckInPat.Value, "1", "0") & ","
        '8.�������
        strSQL = strSQL & "2,"
        '9.PASS���
        strSQL = strSQL & IIf(strPASS = "", "Null", zlStr.FormatString("'[1]'", strPASS)) & ","
        '10.������
        strSQL = strSQL & zlStr.FormatString("'[1]',", UserInfo.����)
        '11.�Ƿ�����
        strSQL = strSQL & "Null)"
            
        'SQL���뼯�϶���
        AddArray colSQL, strSQL
    End If
    
    On Error GoTo errHandle
    ExecuteProcedureArray colSQL, "�����Զ�����Ŀ"

    Save = True

    Exit Function

errHandle:
    If zl9ComLib.ErrCenter = 1 Then
        Resume
    Else
        gcnOracle.RollbackTrans
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim strTmp As String
    
    If Not mrsItems Is Nothing Then
        If mrsItems.State = adStateOpen Then mrsItems.Close
    End If
    
    SaveWinState Me, App.ProductName
    
'    If mblnMemory Then
'        strTmp = GetCurrentVSFHead(vsfItems)
'        SaveSetting "ZLSOFT", FormatString("˽��ģ��\[1]\��������\[2]\[3]\[4]", UserInfo.�û���, App.ProductName, Me.Name), vsfItems.Name, strTmp
'    End If
End Sub

Private Sub lvwPass_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If mblnEnter Then chkPass.Tag = "1"
End Sub

Private Sub optAccord_Click(Index As Integer)
    '��ֹ���δ����¼�
    If Val(fraItems.Tag) = Index Or mblnEnter = False Then Exit Sub
    
    '����Ƿ����δ���ļ�¼
    If GetRecipeAuditBills(0) Then
        optAccord(Val(fraItems.Tag)).Value = 1
        MsgBox "�������ϵͳ�������δ���ļ�¼�����飡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If vsfItems.Rows < 2 Then 'Or vsfItems.Rows = 2 And vsfItems.TextMatrix(1, 0) = "" Then
        '��������
        mrsItems.Filter = zlStr.FormatString("���='[1]' or ���='4-�Զ���'", IIf(optAccord(0).Value, "1-7��", "2-28��"))     '1-7�2-28�3-�̶�1�4-�Զ���
        mdlDefine.FillVSFData vsfItems, mrsItems
        If mrsItems.RecordCount > 0 Then
            '�������������סԺ��Ԫ����ɫ
            Call SetVSFColor
            Call vsfItems_AfterRowColChange(0, 0, 1, 0)
        End If
        fraItems.Tag = CStr(Index)
        Exit Sub
    End If
    
    If MsgBox("�л������Ŀ���ݻὫԭ��������������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        optAccord(Val(fraItems.Tag)).Value = 1
        optAccord(Val(fraItems.Tag)).SetFocus
        Exit Sub
    End If
    
    '��������
    mrsItems.Filter = zlStr.FormatString("���='[1]' or ���='4-�Զ���'", IIf(optAccord(0).Value, "1-7��", "2-28��"))     '1-7�2-28�3-�̶�1�4-�Զ���
    mdlDefine.FillVSFData vsfItems, mrsItems
    If mrsItems.RecordCount > 0 Then
        '�������������סԺ��Ԫ����ɫ
        Call SetVSFColor
        Call vsfItems_AfterRowColChange(0, 0, 1, 0)
    End If
    
    fraItems.Tag = CStr(Index)
End Sub

Private Sub vsfItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnEnter = False Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    With vsfItems
        cmdEdit.Enabled = Val(.TextMatrix(NewRow, .ColIndex("���"))) = 4
        cmdDel.Enabled = cmdEdit.Enabled
    End With
End Sub

Private Sub vsfItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfItems
        If Col = .ColIndex("�������") Then
            If Trim(.TextMatrix(Row, .ColIndex("�������"))) = "0" Or Trim(.TextMatrix(Row, .ColIndex("�������"))) = "2" Then
            Else
                'MsgBox "����Ŀֻ������סԺ��", vbInformation, gstrSysName
                Cancel = True
            End If
        ElseIf Col = .ColIndex("���סԺ") Then
            If Trim(.TextMatrix(Row, .ColIndex("�������"))) = "1" Or Trim(.TextMatrix(Row, .ColIndex("�������"))) = "2" Then
            Else
                'MsgBox "����Ŀֻ���������", vbInformation, gstrSysName
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End With
End Sub

'Private Sub FillVSFData(ByRef vsfVar As VSFlexGrid, ByRef rsVar As ADODB.Recordset)
''���ܣ�����¼����������������vsf�ؼ���
''������
''  vsfVar��Ҫ������ݵ�Vsf�ؼ�
''  rsVar����¼������
'
'    If rsVar Is Nothing Then Exit Sub
'    If rsVar.State <> adStateOpen Then Exit Sub
'    If vsfVar Is Nothing Then Exit Sub
'
'    Dim i As Integer, intCol As Integer
'    Dim lngRow As Long
'
'    With rsVar
'        vsfVar.Redraw = flexRDNone
'        vsfVar.Rows = .RecordCount + 1
'        vsfVar.Clear 1
'
'        lngRow = 1
'        If .RecordCount > 0 Then .MoveFirst
'        Do While .EOF = False
'            For i = 0 To .Fields.Count - 1
'                intCol = vsfVar.ColIndex(.Fields(i).Name)
'                If intCol >= 0 Then
'                    'vsf�д��ڸ��ֶ�
'                    vsfVar.TextMatrix(lngRow, intCol) = zlCommFun.NVL(.Fields(i).Value)
'                End If
'            Next
'
'            lngRow = lngRow + 1
'            .MoveNext
'        Loop
'        vsfVar.Redraw = flexRDDirect
'    End With
'
'End Sub

Private Sub SetVSFColor()
    Dim i As Integer
    
    With Me.vsfItems
        If .Rows <= 1 Then Exit Sub
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("�������"))) = "0" Then
                .Cell(flexcpBackColor, i, .ColIndex("���סԺ"), i, .ColIndex("���סԺ")) = &H8000000C
            ElseIf Trim(.TextMatrix(i, .ColIndex("�������"))) = "1" Then
                .Cell(flexcpBackColor, i, .ColIndex("�������"), i, .ColIndex("�������")) = &H8000000C
            End If
        Next
    End With
End Sub
