VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEInvoiceCheck 
   BorderStyle     =   0  'None
   Caption         =   "����Ʊ�ݺ˶�"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   8748
      Left            =   600
      ScaleHeight     =   8745
      ScaleWidth      =   10335
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   576
      Width           =   10332
      Begin VB.Frame fraSplit 
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   2088
         MousePointer    =   7  'Size N S
         TabIndex        =   13
         Top             =   3624
         Width           =   1005
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   444
         Left            =   72
         ScaleHeight     =   450
         ScaleWidth      =   9990
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   144
         Width           =   9996
         Begin VB.OptionButton opt��Ʊ 
            Caption         =   "��Ʊ"
            Height          =   285
            Left            =   8190
            TabIndex        =   10
            Top             =   92
            Width           =   705
         End
         Begin VB.OptionButton opt��Ʊ����Ʊ 
            Caption         =   "��Ʊ����Ʊ"
            Height          =   285
            Left            =   6810
            TabIndex        =   9
            Top             =   92
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.CommandButton cmdCheck 
            Caption         =   "�˶�(&C)"
            Height          =   300
            Left            =   9015
            TabIndex        =   11
            Top             =   84
            Width           =   1000
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   276
            Left            =   3576
            TabIndex        =   6
            Top             =   96
            Width           =   1356
            _ExtentX        =   2381
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115539971
            CurrentDate     =   43941
         End
         Begin VB.ComboBox cbo��Ʊ�� 
            Height          =   276
            Left            =   672
            TabIndex        =   4
            Top             =   96
            Width           =   1812
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   276
            Left            =   5232
            TabIndex        =   8
            Top             =   96
            Width           =   1356
            _ExtentX        =   2381
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115539971
            CurrentDate     =   43941
         End
         Begin VB.Label lblҵ��ʱ��_ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   4992
            TabIndex        =   7
            Top             =   144
            Width           =   180
         End
         Begin VB.Label lblҵ������ 
            AutoSize        =   -1  'True
            Caption         =   "�շ�����"
            Height          =   180
            Left            =   2760
            TabIndex        =   5
            Top             =   144
            Width           =   720
         End
         Begin VB.Label lbl��Ʊ�� 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ��"
            Height          =   180
            Left            =   72
            TabIndex        =   3
            Top             =   144
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfTotalCheck 
         Height          =   1356
         Left            =   192
         TabIndex        =   12
         Top             =   864
         Width           =   6108
         _cx             =   1983064598
         _cy             =   1983056216
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
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
         FixedCols       =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetailCheck 
         Height          =   1404
         Left            =   816
         TabIndex        =   14
         Top             =   4104
         Width           =   4404
         _cx             =   1983061592
         _cy             =   1983056300
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
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
         FixedCols       =   0
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
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1032
      Left            =   0
      Top             =   888
      Width           =   528
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10848
      _Version        =   589884
      _ExtentX        =   19135
      _ExtentY        =   529
      _StockProps     =   6
      Caption         =   "����Ʊ�ݺ˶�"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmEInvoiceCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Object, mlngSys As Long, mlngModule As Long, mstrDBUser As String
Private mcbsMain   As Object          'CommandBar�ؼ�
Private mobjEInvoice As clsEInvoiceModule
Private mblnPrinting As Boolean
Private mrs��Ʊ�� As ADODB.Recordset
Private mrs�շ�Ա As ADODB.Recordset
Private mbytƱ�ݺ˶�ʱ������ As Byte '0-Ʊ�ݿ���ʱ�䣬1-����ҵ����ʱ��
Private mstrEInvoiceNodeCode As String '��Ʊ��

Public Event ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
Public Event ShowInfo(ByVal strInfo As String)

Public Sub InitCommVariable(frmParent As Object, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String, _
    objEInvoice As Object)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
    Set mobjEInvoice = objEInvoice
    mbytƱ�ݺ˶�ʱ������ = mobjEInvoice.ZLCheckTimeMode
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom

    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�˶���ϸ(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FlatAccount, "ƽ������(&M)"): cbrControl.BeginGroup = True
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
    End With

    '����������
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�˶���ϸ", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FlatAccount, "ƽ������", cbrControl.index + 1): cbrControl.BeginGroup = True
    End With

    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        '.Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Select Case Control.ID
    Case conMenu_File_Preview 'Ԥ��
        Call OutputList(2)
    Case conMenu_File_Print '��ӡ
        Call OutputList(1)
    Case conMenu_File_Excel '�����Excel��
        Call OutputList(3)
    Case conMenu_Edit_Audit '�˶�
        Call DetailCheck
    Case conMenu_Edit_FlatAccount 'ƽ������
        Call FlatAccount
    End Select
End Sub

Private Sub FlatAccount()
    '��ƽ�ʴ���
    Dim i As Long, byt���� As Byte
    Dim lngEInvoiceID As Long, lng����ID As Long, bln�ѻ���ֽ�� As Boolean
    Dim rsEInvoice As ADODB.Recordset, strErrMsg As String
    Dim strSQL As String, blnTrans As Boolean
    Dim strҵ������ As String, strDate As String
    Dim bln�����ɹ� As Boolean, strMsg As String
    Dim cllPro As New Collection, lng����ID As Long
    
    Dim strSysSouceName_Out As String, strExtend As String
    Dim strEInvoiceCode_out As String, strEInvoiceNo_Out As String
    Dim strCheckCode_out As String, strCreateTime_Out As String, strEInvQRCode_Out As String, strEInvUrl_Out As String, strEInvUrl1_Out As String
    Dim strEinvRemark_Out As String
    
    On Error GoTo ErrHandler
    strҵ������ = vsfTotalCheck.TextMatrix(vsfTotalCheck.Row, vsfTotalCheck.ColIndex("ҵ������"))
    If strҵ������ = "" Then Exit Sub
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With vsfDetailCheck
        For i = .FixedRows To .Rows - 1
            '.RowData(i) = "":ƽ̨��his�޵�Ʊ�ݣ�����Ʊ����his�޼�¼����������
            If .TextMatrix(i, .ColIndex("�˶Խ��")) = "�˶�ʧ��" And InStr(1, .RowData(i), "_") > 0 Then
                Set cllPro = New Collection
                lng����ID = Split(.RowData(i), "_")(1): lngEInvoiceID = Split(.RowData(i), "_")(2)
                
                Select Case Val(zlStr.NeedCode(.TextMatrix(i, .ColIndex("������ʽ"))))
                Case 1 '1-����HIS����
                    Set rsEInvoice = GetEInvoiceInfo(lngEInvoiceID, strErrMsg)
                    If Val(Nvl(rsEInvoice!�Ƿ񻻿�)) = 1 Then '�ѻ���ֽ��Ʊ�ݣ��޷������ջ�
                        bln�����ɹ� = False
                        strMsg = "�ѻ���ֽ��Ʊ�ݣ��޷����������ջ�"
                    Else
                        bln�����ɹ� = True
                        lng����ID = zlDatabase.GetNextId("����Ʊ��ʹ�ü�¼")
                        'Zl_����Ʊ��ʹ�ü�¼_Delete
                        strSQL = "Zl_����Ʊ��ʹ�ü�¼_Delete("
                        '  Id_In           In ����Ʊ��ʹ�ü�¼.Id%Type,
                        strSQL = strSQL & "" & lng����ID & ","
                        '  ��Ʊ��_In       In ����Ʊ��ʹ�ü�¼.��Ʊ��%Type,
                        strSQL = strSQL & "'" & .Cell(flexcpData, i, .ColIndex("HIS����-��Ʊ��")) & "',"
                        '  ϵͳ��Դ_In     In ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  ����ʱ��_In     In ����Ʊ��ʹ�ü�¼.����ʱ��%Type,
                        strSQL = strSQL & "'" & Format(strDate, "yyyyMMddHHmmss000") & "',"
                        '  ��ע_In         In ����Ʊ��ʹ�ü�¼.��ע%Type,
                        strSQL = strSQL & "'" & "ƽ������������HIS����" & "',"
                        '  ����Ա���_In   In ����Ʊ��ʹ�ü�¼.����Ա���%Type,
                        strSQL = strSQL & "'" & UserInfo.��� & "',"
                        '  ����Ա����_In   In ����Ʊ��ʹ�ü�¼.����Ա����%Type,
                        strSQL = strSQL & "'" & UserInfo.���� & "',"
                        '  �Ǽ�ʱ��_In     In ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type,
                        strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                        '  ԭ����Ʊ��id_In In ����Ʊ��ʹ�ü�¼.Id%Type
                        strSQL = strSQL & "" & lngEInvoiceID & ")"
                        cllPro.Add strSQL
                        
                        '���ϼ�¼����������¼
                        'Zl_����Ʊ��������¼_Update
                        strSQL = "Zl_����Ʊ��������¼_Update("
                        '  ҵ������_In   ����Ʊ��������¼.ҵ������%Type,
                        strSQL = strSQL & "To_Date('" & strҵ������ & "','yyyy-mm-dd'),"
                        '  ����Ʊ��id_In ����Ʊ��������¼.����Ʊ��id%Type,
                        strSQL = strSQL & "" & lng����ID & ","
                        '  ҵ����ˮ��_In ����Ʊ��������¼.ҵ����ˮ��%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  His��Ʊ��_In    ����Ʊ��������¼.His��Ʊ��%Type,
                        strSQL = strSQL & "'" & .Cell(flexcpData, i, .ColIndex("HIS����-��Ʊ��")) & "',"
                        '  His��Ʊ���_In  ����Ʊ��������¼.His��Ʊ���%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("HIS����-��Ʊ���"))) & ","
                        '  HisƱ��״̬_In  ����Ʊ��������¼.HisƱ��״̬%Type,
                        strSQL = strSQL & "" & 3 & ","
                        '  ƽ̨��Ʊ��_In   ����Ʊ��������¼.ƽ̨��Ʊ��%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  ƽ̨��Ʊ���_In ����Ʊ��������¼.ƽ̨��Ʊ���%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  ƽ̨Ʊ��״̬_In ����Ʊ��������¼.ƽ̨Ʊ��״̬%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  ������ʽ_In   ����Ʊ��������¼.������ʽ%Type,
                        strSQL = strSQL & "" & 4 & ","
                        '  ������_In     ����Ʊ��������¼.������%Type,
                        strSQL = strSQL & "'" & UserInfo.���� & "',"
                        '  ����ʱ��_In   ����Ʊ��������¼.����ʱ��%Type,
                        strSQL = strSQL & "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                        '  �������_In   ����Ʊ��������¼.�������%Type,
                        strSQL = strSQL & "" & 1 & ","
                        '  ����˵��_In   ����Ʊ��������¼.����˵��%Type
                        strSQL = strSQL & "'" & "ƽ������������HIS����ʱ���������ϼ�¼" & "')"
                        cllPro.Add strSQL
                    End If
                    
                Case 2 '2-����ƽ̨����
                    strExtend = GetJsonNodeString("bustype", .Cell(flexcpData, i, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݴ���")), Json_Text)
                    strExtend = strExtend & "," & GetJsonNodeString("billbatchcode", .TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݴ���")), Json_Text)
                    strExtend = strExtend & "," & GetJsonNodeString("billno", .TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݺ���")), Json_Text)
                    strExtend = "{" & strExtend & "}"
                    
                    bln�����ɹ� = mobjEInvoice.zlCancelEInvoice(Me, lngEInvoiceID, mstrEInvoiceNodeCode, strSysSouceName_Out, _
                        strEInvoiceCode_out, strEInvoiceNo_Out, strCheckCode_out, strCreateTime_Out, strEInvQRCode_Out, strEInvUrl_Out, strEInvUrl1_Out, strEinvRemark_Out, strMsg, strExtend)
                    
                    If bln�����ɹ� Then '����¼����������¼
                         'Zl_����Ʊ��������¼_Update
                        strSQL = "Zl_����Ʊ��������¼_Update("
                        '  ҵ������_In   ����Ʊ��������¼.ҵ������%Type,
                        strSQL = strSQL & "To_Date('" & strҵ������ & "','yyyy-mm-dd'),"
                        '  ����Ʊ��id_In ����Ʊ��������¼.����Ʊ��id%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  ҵ����ˮ��_In ����Ʊ��������¼.ҵ����ˮ��%Type,
                        strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-ҵ����ˮ��")) & "',"
                        '  His��Ʊ��_In    ����Ʊ��������¼.His��Ʊ��%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  His��Ʊ���_In  ����Ʊ��������¼.His��Ʊ���%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  HisƱ��״̬_In  ����Ʊ��������¼.HisƱ��״̬%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  ƽ̨��Ʊ��_In   ����Ʊ��������¼.ƽ̨��Ʊ��%Type,
                        strSQL = strSQL & "'" & mstrEInvoiceNodeCode & "',"
                        '  ƽ̨��Ʊ���_In ����Ʊ��������¼.ƽ̨��Ʊ���%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-��Ʊ���"))) & ","
                        '  ƽ̨Ʊ��״̬_In ����Ʊ��������¼.ƽ̨Ʊ��״̬%Type,
                        strSQL = strSQL & "" & 3 & ","
                        '  ������ʽ_In   ����Ʊ��������¼.������ʽ%Type,
                        strSQL = strSQL & "" & 4 & ","
                        '  ������_In     ����Ʊ��������¼.������%Type,
                        strSQL = strSQL & "'" & UserInfo.���� & "',"
                        '  ����ʱ��_In   ����Ʊ��������¼.����ʱ��%Type,
                        strSQL = strSQL & "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                        '  �������_In   ����Ʊ��������¼.�������%Type,
                        strSQL = strSQL & "" & 1 & ","
                        '  ����˵��_In   ����Ʊ��������¼.����˵��%Type
                        strSQL = strSQL & "'" & "ƽ������������ƽ̨����ʱ�����ĳ���¼" & "')"
                        cllPro.Add strSQL
                    End If
                    
                Case 3 '3-����HIS��ƽ̨�����ؿ�Ʊ��
                    '�ݲ�����Ӧ�ò�������������
                    bln�����ɹ� = False
                    strMsg = "�ݲ�֧������HIS��ƽ̨�����ؿ�Ʊ��"
                    
                Case 4 '4-�����������
                    bln�����ɹ� = True
                End Select
                
                'Zl_����Ʊ��������¼_Update
                strSQL = "Zl_����Ʊ��������¼_Update("
                '  ҵ������_In   ����Ʊ��������¼.ҵ������%Type,
                strSQL = strSQL & "To_Date('" & strҵ������ & "','yyyy-mm-dd'),"
                '  ����Ʊ��id_In ����Ʊ��������¼.����Ʊ��id%Type,
                strSQL = strSQL & "" & ZVal(lngEInvoiceID) & ","
                '  ҵ����ˮ��_In ����Ʊ��������¼.ҵ����ˮ��%Type,
                strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-ҵ����ˮ��")) & "',"
                '  His��Ʊ��_In    ����Ʊ��������¼.His��Ʊ��%Type,
                strSQL = strSQL & "'" & .Cell(flexcpData, i, .ColIndex("HIS����-��Ʊ��")) & "',"
                '  His��Ʊ���_In  ����Ʊ��������¼.His��Ʊ���%Type,
                strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("HIS����-��Ʊ���"))) & ","
                '  HisƱ��״̬_In  ����Ʊ��������¼.HisƱ��״̬%Type,
                strSQL = strSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("HIS����-Ʊ��״̬"))) & ","
                '  ƽ̨��Ʊ��_In   ����Ʊ��������¼.ƽ̨��Ʊ��%Type,
                strSQL = strSQL & "'" & .Cell(flexcpData, i, .ColIndex("Ʊ��ƽ̨����-��Ʊ��")) & "',"
                '  ƽ̨��Ʊ���_In ����Ʊ��������¼.ƽ̨��Ʊ���%Type,
                strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-��Ʊ���"))) & ","
                '  ƽ̨Ʊ��״̬_In ����Ʊ��������¼.ƽ̨Ʊ��״̬%Type,
                strSQL = strSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("Ʊ��ƽ̨����-Ʊ��״̬"))) & ","
                '  ������ʽ_In   ����Ʊ��������¼.������ʽ%Type,
                strSQL = strSQL & "" & Val(zlStr.NeedCode(.TextMatrix(i, .ColIndex("������ʽ")))) & ","
                '  ������_In     ����Ʊ��������¼.������%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  ����ʱ��_In   ����Ʊ��������¼.����ʱ��%Type,
                strSQL = strSQL & "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                '  �������_In   ����Ʊ��������¼.�������%Type,
                strSQL = strSQL & "" & IIf(bln�����ɹ�, 1, 0) & ","
                '  ����˵��_In   ����Ʊ��������¼.����˵��%Type
                strSQL = strSQL & "'" & strMsg & "')"
                cllPro.Add strSQL
                
                gcnOracle.BeginTrans: blnTrans = True
                ExecuteProcedureArrAy cllPro, Me.Caption, True, True
                gcnOracle.CommitTrans: blnTrans = False
                
                .TextMatrix(i, .ColIndex("������")) = UserInfo.����
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(strDate, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, .ColIndex("�������")) = IIf(bln�����ɹ�, "�����ɹ�", "����ʧ��")
                .TextMatrix(i, .ColIndex("����˵��")) = strMsg
                
                If bln�����ɹ� Then
                    .TextMatrix(i, .ColIndex("�˶Խ��")) = "�˶Գɹ�"
                    .TextMatrix(i, .ColIndex("�˶�˵��")) = ""
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                Else
                    .TextMatrix(i, .ColIndex("�˶Խ��")) = "�˶�ʧ��"
                    .TextMatrix(i, .ColIndex("�˶�˵��")) = strMsg
                End If
            End If
        Next
    End With
    
    Call DetailCheck
    Exit Sub
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte
    Dim intCurrentRow As Integer, vsfGrid As VSFlexGrid
    
    On Error GoTo ErrHandler
    '��ͷ
    Set objOut = New zlPrint1Grd
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    If Me.ActiveControl Is vsfDetailCheck Then
        Set vsfGrid = vsfDetailCheck
        objOut.Title.Text = "����Ʊ����ϸ�˶��嵥"
    Else
        Set vsfGrid = vsfTotalCheck
        objOut.Title.Text = "����Ʊ�ݻ��ܺ˶��嵥"
    End If
    
    '����
    If Me.ActiveControl Is vsfDetailCheck Then
        Set objRow = New zlTabAppRow
        objRow.Add "��Ʊ�㣺" & cbo��Ʊ��.Text
        objRow.Add "ҵ�����ڣ�" & vsfTotalCheck.TextMatrix(vsfTotalCheck.Row, vsfTotalCheck.ColIndex("ҵ������"))
        objOut.UnderAppRows.Add objRow
    Else
        Set objRow = New zlTabAppRow
        objRow.Add "��Ʊ�㣺" & cbo��Ʊ��.Text
        objRow.Add "ҵ��ʱ�䣺" & Format(dtp��ʼʱ��, "yyyy-mm-dd") & " �� " & Format(dtp����ʱ��, "yyyy-mm-dd")
        objOut.UnderAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    vsfGrid.Redraw = False
    intCurrentRow = vsfGrid.Row
    mblnPrinting = True
    
    '����
    Set objOut.Body = vsfGrid
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mblnPrinting = False
    vsfGrid.Row = intCurrentRow
    vsfGrid.Redraw = True
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    mblnPrinting = False
    vsfGrid.Row = intCurrentRow
    vsfGrid.Redraw = True
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel��
        If Me.ActiveControl Is vsfDetailCheck Then
            Control.Enabled = vsfDetailCheck.TextMatrix(2, 0) <> ""
        Else
            Control.Enabled = vsfTotalCheck.TextMatrix(2, 0) <> ""
        End If
    
    Case conMenu_Edit_Audit '�˶���ϸ
        Control.Enabled = vsfTotalCheck.TextMatrix(2, 0) <> ""
    Case conMenu_Edit_FlatAccount 'ƽ������
        If vsfTotalCheck.Row > 0 Then
            Control.Enabled = vsfDetailCheck.TextMatrix(2, 0) <> "" And vsfTotalCheck.TextMatrix(vsfTotalCheck.Row, vsfTotalCheck.ColIndex("�˶Խ��")) = "�˶�ʧ��"
        Else
            Control.Enabled = False
        End If
    
    Case conMenu_View_Refresh 'ˢ��
        Control.Visible = False
        Control.Enabled = Control.Visible
    End Select
End Sub

Private Sub cbo��Ʊ��_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
     
    If cbo��Ʊ��.ListIndex <> -1 Then
        '�����б�ʱ,�����ı�������������
        If UCase(cbo��Ʊ��.Text) <> UCase(cbo��Ʊ��.List(cbo��Ʊ��.ListIndex)) Then Call zlControl.CboSetIndex(cbo��Ʊ��.hWnd, -1)
    End If
    
    If cbo��Ʊ��.Text = "" Then
        cbo��Ʊ��.ListIndex = -1
    ElseIf cbo��Ʊ��.ListIndex = -1 Then
        If mrs�շ�Ա Is Nothing Then
            If Select��Ʊ��(Me, mlngSys, mlngModule, cbo��Ʊ��, mrs��Ʊ��) = False Then
                KeyAscii = 0: zlControl.TxtSelAll cbo��Ʊ��: Exit Sub
            End If
        Else
            If Select�շ�Ա(Me, mlngSys, mlngModule, cbo��Ʊ��, mrs�շ�Ա) = False Then
                KeyAscii = 0: zlControl.TxtSelAll cbo��Ʊ��: Exit Sub
            End If
        End If
    End If
    
    If cbo��Ʊ��.ListIndex = -1 Then cbo��Ʊ��.Text = ""
End Sub

Private Sub cbo��Ʊ��_LostFocus()
    If cbo��Ʊ��.Text <> "" And cbo��Ʊ��.ListIndex < 0 Then cbo��Ʊ��.Text = ""
End Sub

Private Sub cmdCheck_Click()
    Call TotalCheck
End Sub

Private Sub TotalCheck()
    '���ܺ˶�
    Dim dtBegin As Date, dtEnd As Date, strErrMsg As String
    Dim str��Ʊ�� As String, bytMode As Byte '1-�˶Կ�Ʊ����Ʊ��2-���˶���Ʊ
    
    '1.���ݼ��
    On Error GoTo ErrHandler
    dtBegin = Format(dtp��ʼʱ��.Value, "yyyy-MM-dd"): dtEnd = Format(dtp����ʱ��.Value, "yyyy-MM-dd 23:59:59")
    If dtp��ʼʱ�� > dtp����ʱ�� Then
        MsgBox "��ʼʱ�䲻�ܴ��ڽ���ʱ�䣡", vbInformation, gstrSysName
        zlControl.ControlSetFocus dtp����ʱ��:  Exit Sub
    End If
    
    If DateDiff("m", dtp��ʼʱ��, dtp����ʱ��) > 6 Then
        If MsgBox("�Ե�ǰʱ�䷶Χ�ڵ����ݽ��к˶Կ�����Ҫ�ϳ�ʱ�䣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
     
    bytMode = IIf(opt��Ʊ����Ʊ.Value, 1, 2)
    str��Ʊ�� = zlStr.NeedCode(cbo��Ʊ��.Text)
    
    '2.��ȡ����
    Dim strSQL As String, rsHISEInvoice As ADODB.Recordset, strWhere As String, strSqlSub As String
    If mbytƱ�ݺ˶�ʱ������ = 1 Then
        If bytMode = 2 Then strWhere = " And a.��¼״̬ = 2"
        If str��Ʊ�� <> "" Then strWhere = strWhere & " And a.��Ʊ�� = [3]"
        
        '1)Ԥ����
        strSQL = _
            " Select Distinct a.ID" & _
            " From ����Ʊ��ʹ�ü�¼ A,����Ԥ����¼ B" & _
            " Where a.����ID =b.ID And a.Ʊ��=2 And b.��¼����=1 And b.�տ�ʱ�� Between [1] And [2]" & strWhere
        '����˿�
        strSQL = strSQL & " Union All " & _
            " Select Distinct a.ID" & _
            " From ����Ʊ��ʹ�ü�¼ A,����Ԥ����¼ B" & _
            " Where a.�˿�ID =b.ID And a.Ʊ��=2 And b.��¼����=11 And b.�տ�ʱ�� Between [1] And [2]" & strWhere
        '2)���￨
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select Distinct a.ID " & _
            " From ����Ʊ��ʹ�ü�¼ A,סԺ���ü�¼ B" & _
            " Where a.����ID =b.����ID And a.Ʊ��=5 And b.��¼����=5 And b.��¼״̬ In(1,3) And b.�Ǽ�ʱ�� Between [1] And [2]" & strWhere
        '3)����
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                " Select Distinct a.ID" & _
                " From ����Ʊ��ʹ�ü�¼ A,���˽��ʼ�¼ B" & _
                " Where a.����ID =b.ID And a.Ʊ��=3 And b.��¼״̬ In(1,3) And b.�շ�ʱ�� Between [1] And [2]" & strWhere
        '4)�Һš��շ�
        strSqlSub = _
            " Select Distinct a.ID" & _
            " From ����Ʊ��ʹ�ü�¼ A,������ü�¼ B" & _
            " Where a.����ID =b.����ID And a.Ʊ��=[Ʊ��] And b.��¼����=[��¼����] And b.��¼״̬ In(1,3) And b.�Ǽ�ʱ�� Between [1] And [2]" & strWhere
        
        '���ղ������
        strSqlSub = strSqlSub & " Union All " & _
            " Select Distinct a.ID" & _
            " From ����Ʊ��ʹ�ü�¼ A,���ò����¼ B" & _
            " Where a.����ID =b.����ID And a.Ʊ��=[Ʊ��]  And b.��¼����=[��¼����] And Nvl(b.���ӱ�־,0)=[���ӱ�־]" & _
            "           And b.��¼״̬ In(1,3) And b.�Ǽ�ʱ�� Between [1] And [2]" & strWhere
            
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            Replace(Replace(Replace(strSqlSub, "[��¼����]", 1), "[���ӱ�־]", 0), "[Ʊ��]", 1)

        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            Replace(Replace(Replace(strSqlSub, "[��¼����]", 4), "[���ӱ�־]", 1), "[Ʊ��]", 4)
        
        strSQL = _
            " Select To_Char(To_Date(Substr(a.����ʱ��, 1, 8), 'yyyymmdd'),'yyyy-mm-dd') As ҵ������, Count(1) As ��Ʊ��, Sum(Decode(a.��¼״̬, 2, -1, 1) * a.Ʊ�ݽ��) As ��Ʊ���" & _
            " From ����Ʊ��ʹ�ü�¼ A,(" & strSQL & ") B" & _
            " Where a.ID =b.ID" & _
            " Group By To_Date(Substr(a.����ʱ��, 1, 8), 'yyyymmdd')"
        
    Else
        strSQL = _
            " Select To_Char(To_Date(Substr(a.����ʱ��, 1, 8), 'yyyymmdd'),'yyyy-mm-dd') As ҵ������, Count(1) As ��Ʊ��, Sum(Decode(a.��¼״̬, 2, -1, 1) * a.Ʊ�ݽ��) As ��Ʊ���" & _
            " From ����Ʊ��ʹ�ü�¼ A" & _
            " Where To_Date(Substr(a.����ʱ��, 1, 8), 'yyyymmdd') Between [1] And [2]" & _
                        IIf(bytMode = 2, " And a.��¼״̬ = 2", " And a.��¼״̬ In(1,2,3)") & _
                        IIf(str��Ʊ�� = "", "", " And a.��Ʊ��=[3]") & _
            " Group By To_Date(Substr(a.����ʱ��, 1, 8), 'yyyymmdd')"
    End If
    Set rsHISEInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtBegin, dtEnd, str��Ʊ��)
            
    Dim clldata As Collection, cllDatas As Collection '����(ҵ������,�ܱ���,��Ʊ��,��Ʊ���,���ؽ��,����ԭ��),Key=_ҵ������
    If mobjEInvoice.ZlGetTotalCheckData(dtBegin, dtEnd, cllDatas, bytMode, str��Ʊ��, strErrMsg) Then
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    
    Dim rsPreCheck As ADODB.Recordset '�ϴκ˶Լ�¼
    strSQL = _
        " Select To_Char(ҵ������,'yyyy-mm-dd') As ҵ������, �˶���, �˶�ʱ��, �˶Խ��, �˶�˵��" & _
        " From (Select a.ҵ������, a.�˶���, a.�˶�ʱ��, a.�˶Խ��, a.�˶�˵��, Row_Number() Over(Partition By a.ҵ������ Order By a.�˶�ʱ�� Desc) As ���" & _
        "           From ����Ʊ�ݺ˶Լ�¼ A" & _
        "           Where ҵ������ Between [1] And [2] And �˶�����=[3]" & _
                                IIf(str��Ʊ�� = "", "", " And a.��Ʊ��=[4]") & _
        "           )" & _
        " Where ��� = 1"
    Set rsPreCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtBegin, dtEnd, bytMode, str��Ʊ��)
    
    Dim rsUpdate As ADODB.Recordset '������¼
    strSQL = _
        " Select 1 As ����, To_Char(ҵ������,'yyyy-mm-dd') As ҵ������, Count(1) As ��Ʊ��, Sum(HIS��Ʊ���) As ��Ʊ���" & _
        " From ����Ʊ��������¼" & _
        " Where ҵ������ Between [1] And [2] And ������� = 1 And ����Ʊ��id Is Not Null" & _
                    IIf(bytMode = 2, " And HISƱ��״̬ = 2", " And HISƱ��״̬ In(1,2,3)") & _
                    IIf(str��Ʊ�� = "", "", " And HIS��Ʊ��=[3]") & _
        " Group By ҵ������" & _
        " Union All" & _
        " Select 2 As ����, To_Char(ҵ������,'yyyy-mm-dd') As ҵ������, Count(1) As ��Ʊ��, Sum(ƽ̨��Ʊ���) As ��Ʊ���" & _
        " From ����Ʊ��������¼" & _
        " Where ҵ������ Between [1] And [2] And ������� = 1 And ҵ����ˮ�� Is Not Null" & _
                    IIf(bytMode = 2, " And ƽ̨Ʊ��״̬ = 2", " And ƽ̨Ʊ��״̬ In(1,2,3)") & _
                    IIf(str��Ʊ�� = "", "", " And ƽ̨��Ʊ��=[3]") & _
        " Group By ҵ������"
    Set rsUpdate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtBegin, dtEnd, str��Ʊ��)

    '3.��ʼ�˶�
    Dim dtCurrent As Date, blnChecked As Boolean, strCheckMsg As String
    Dim lngHIS��Ʊ�� As Long, lngƽ̨��Ʊ�� As Long, dblHIS��Ʊ�� As Double, dblƽ̨��Ʊ�� As Double
    Dim lngOldRow As Long, lngOldCol As Long
    
    lngOldRow = vsfTotalCheck.Row: lngOldCol = vsfTotalCheck.Col
    vsfTotalCheck.Clear 1
    vsfTotalCheck.Rows = vsfTotalCheck.FixedRows + 1
    
    vsfDetailCheck.Clear 1
    vsfDetailCheck.Rows = vsfDetailCheck.FixedRows + 1
    
    'ҵ������,HIS����-��Ʊ��,HIS����-��Ʊ���,Ʊ��ƽ̨����-��Ʊ��,Ʊ��ƽ̨����-��Ʊ���,Ʊ��ƽ̨����-�ܱ���,
    '�˶Խ��,�˶�˵��,�ϴκ˶���,�ϴκ˶�ʱ��,�ϴκ˶Խ��,�ϴκ˶�˵��
    With vsfTotalCheck
        .Redraw = flexRDNone
        
        dtCurrent = dtBegin
        Do While dtCurrent <= dtEnd
            
            If .TextMatrix(.Rows - 1, .ColIndex("ҵ������")) <> "" Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("ҵ������")) = Format(dtCurrent, "yyyy-MM-dd")
            
            lngHIS��Ʊ�� = 0: dblHIS��Ʊ�� = 0
            rsHISEInvoice.Filter = "ҵ������='" & Format(dtCurrent, "yyyy-MM-dd") & "'"
            If Not rsHISEInvoice.EOF Then
                lngHIS��Ʊ�� = Val(Nvl(rsHISEInvoice!��Ʊ��)): dblHIS��Ʊ�� = Val(Nvl(rsHISEInvoice!��Ʊ���))
                rsUpdate.Filter = "����=1 And ҵ������='" & Format(dtCurrent, "yyyy-MM-dd") & "'"
                If Not rsUpdate.EOF Then '�ų������������¼
                    lngHIS��Ʊ�� = lngHIS��Ʊ�� - Val(Nvl(rsUpdate!��Ʊ��))
                    dblHIS��Ʊ�� = dblHIS��Ʊ�� - Val(Nvl(rsUpdate!��Ʊ���))
                End If
                
                .TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊ��")) = lngHIS��Ʊ��
                .TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊ���")) = FormatEx(dblHIS��Ʊ��, 6, , , 2)
            End If
            
            lngƽ̨��Ʊ�� = 0: dblƽ̨��Ʊ�� = 0
            Set clldata = Nothing
            If CollectionExitsValue(cllDatas, "_" & Format(dtCurrent, "yyyy-MM-dd")) Then
                Set clldata = cllDatas("_" & Format(dtCurrent, "yyyy-MM-dd"))
            End If
            
            blnChecked = False
            If Not clldata Is Nothing Then
                lngƽ̨��Ʊ�� = Val(Nvl(clldata("��Ʊ��"))): dblƽ̨��Ʊ�� = Val(Nvl(clldata("��Ʊ��")))
                rsUpdate.Filter = "����=2 And ҵ������='" & Format(dtCurrent, "yyyy-MM-dd") & "'"
                If Not rsUpdate.EOF Then '�ų������������¼
                    lngHIS��Ʊ�� = lngHIS��Ʊ�� - Val(Nvl(rsUpdate!��Ʊ��))
                    dblHIS��Ʊ�� = dblHIS��Ʊ�� - Val(Nvl(rsUpdate!��Ʊ���))
                End If
                
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ��")) = lngƽ̨��Ʊ��
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ���")) = FormatEx(dblƽ̨��Ʊ��, 6, , , 2)
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-�ܱ���")) = Nvl(clldata("�ܱ���"))
                If Nvl(clldata("���ؽ��")) = "ʧ��" Then
                    blnChecked = True
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶�ʧ��"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = Nvl(clldata("����ԭ��"))
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                End If
            End If
            
            If Not blnChecked Then
                blnChecked = True: strCheckMsg = ""
                '�˶Թ���HIS��Ʊ�� = ƽ̨��Ʊ�� And HIS��Ʊ��� = ƽ̨��Ʊ���
                If lngHIS��Ʊ�� <> lngƽ̨��Ʊ�� Then
                    strCheckMsg = strCheckMsg & "  HIS��Ʊ����" & lngHIS��Ʊ�� & "��/ƽ̨��Ʊ����" & lngƽ̨��Ʊ�� & "��"
                End If
                If dblHIS��Ʊ�� <> dblƽ̨��Ʊ�� Then
                    strCheckMsg = strCheckMsg & "  HIS��Ʊ��" & FormatEx(dblHIS��Ʊ��, 6, , , 2) & "/ƽ̨��Ʊ��" & FormatEx(dblƽ̨��Ʊ��, 6, , , 2)
                End If
                If strCheckMsg <> "" Then strCheckMsg = Mid(strCheckMsg, 2)
                blnChecked = strCheckMsg = ""
                
                If blnChecked Then
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶Գɹ�"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = ""
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶�ʧ��"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = strCheckMsg
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                End If
            End If
            
            rsPreCheck.Filter = "ҵ������='" & Format(dtCurrent, "yyyy-MM-dd") & "'"
            If Not rsPreCheck.EOF Then
                .TextMatrix(.Rows - 1, .ColIndex("�ϴκ˶���")) = Nvl(rsPreCheck!�˶���)
                .TextMatrix(.Rows - 1, .ColIndex("�ϴκ˶�ʱ��")) = Format(Nvl(rsPreCheck!�˶�ʱ��), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("�ϴκ˶Խ��")) = IIf(Val(Nvl(rsPreCheck!�˶Խ��)) = 1, "�˶Գɹ�", "�˶�ʧ��")
                .TextMatrix(.Rows - 1, .ColIndex("�ϴκ˶�˵��")) = Nvl(rsPreCheck!�˶�˵��)
            End If
        
            dtCurrent = DateAdd("d", 1, dtCurrent)
        Loop
        
        If .Rows > .FixedRows And .Cols > .FixedCols Then     'ȱʡ��λ��
            .Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
            .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows - 1 > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
            .Col = IIf(lngOldCol = 0 Or lngOldCol > .Cols - 1, .FixedCols, lngOldCol)
            .ShowCell .Row, .Col  '������ʾ��ָ����Ԫ
        End If
        .Redraw = flexRDBuffered
    End With
    
    If SaveTotalCheckData(bytMode, str��Ʊ��) = False Then
        vsfTotalCheck.Clear 1
        vsfTotalCheck.Rows = vsfTotalCheck.FixedRows + 1
    End If
    Call ShowTotalRow
    Exit Sub
ErrHandler:
    vsfTotalCheck.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowTotalRow()
    '��ʾ������
    Dim i As Long
    
    On Error GoTo ErrHandler
    With vsfTotalCheck
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("ҵ������")) = "�ϼ�"
        For i = .FixedRows To .Rows - 1
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊ��")) = Val(.TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊ��"))) + Val(.TextMatrix(i, .ColIndex("HIS����-��Ʊ��")))
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊ���")) = FormatEx(Val(.TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊ���"))) + Val(.TextMatrix(i, .ColIndex("HIS����-��Ʊ���"))), 6, , , 2)
            .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ��")) = Val(.TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ��"))) + Val(.TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-��Ʊ��")))
            .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ���")) = FormatEx(Val(.TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ���"))) + Val(.TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-��Ʊ���"))), 6, , , 2)
            .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-�ܱ���")) = Val(.TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-�ܱ���"))) + Val(.TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-�ܱ���")))
        Next
        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveTotalCheckData(ByVal byt�˶����� As Byte, ByVal str��Ʊ�� As String) As Boolean
    '������ܺ˶�����
    '��Σ�
    '   byt�˶����� 1-�˶Կ�Ʊ����Ʊ��2-���˶���Ʊ
    Dim strSQL As String, cllPro As New Collection
    Dim blnTran As Boolean, i As Long, strDate As String
    
    On Error GoTo ErrHandler
    
    'ҵ������,HIS����-��Ʊ��,HIS����-��Ʊ���,Ʊ��ƽ̨����-��Ʊ��,Ʊ��ƽ̨����-��Ʊ���,Ʊ��ƽ̨����-�ܱ���,
    '�˶Խ��,�˶�˵��,�ϴκ˶���,�ϴκ˶�ʱ��,�ϴκ˶Խ��,�ϴκ˶�˵��
    With vsfTotalCheck
        If .TextMatrix(.FixedRows, .ColIndex("ҵ������")) = "" Then SaveTotalCheckData = True: Exit Function
        
        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        For i = .FixedRows To .Rows - 1
            'Zl_����Ʊ�ݺ˶Լ�¼_Update
            strSQL = "Zl_����Ʊ�ݺ˶Լ�¼_Update("
            '  �˶�����_In     ����Ʊ�ݺ˶Լ�¼.�˶�����%Type,
            strSQL = strSQL & "" & byt�˶����� & ","
            '  ҵ������_In     ����Ʊ�ݺ˶Լ�¼.ҵ������%Type,
            strSQL = strSQL & "To_Date('" & Format(.TextMatrix(i, .ColIndex("ҵ������")), "yyyy-MM-dd") & "','yyyy-mm-dd'),"
            '  ��Ʊ��_In       ����Ʊ�ݺ˶Լ�¼.��Ʊ��%Type,
            strSQL = strSQL & "'" & str��Ʊ�� & "',"
            '  His��Ʊ��_In    ����Ʊ�ݺ˶Լ�¼.His��Ʊ��%Type,
            strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("HIS����-��Ʊ��"))) & ","
            '  His��Ʊ���_In  ����Ʊ�ݺ˶Լ�¼.His��Ʊ���%Type,
            strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("HIS����-��Ʊ���"))) & ","
            '  ƽ̨��Ʊ��_In   ����Ʊ�ݺ˶Լ�¼.ƽ̨��Ʊ��%Type,
            strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-��Ʊ��"))) & ","
            '  ƽ̨��Ʊ���_In ����Ʊ�ݺ˶Լ�¼.ƽ̨��Ʊ���%Type,
            strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("Ʊ��ƽ̨����-��Ʊ���"))) & ","
            '  �˶���_In       ����Ʊ�ݺ˶Լ�¼.�˶���%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  �˶�ʱ��_In     ����Ʊ�ݺ˶Լ�¼.�˶�ʱ��%Type,
            strSQL = strSQL & "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '  �˶Խ��_In     ����Ʊ�ݺ˶Լ�¼.�˶Խ��%Type,
            strSQL = strSQL & "" & IIf(.TextMatrix(i, .ColIndex("�˶Խ��")) = "�˶Գɹ�", 1, 0) & ","
            '  �˶�˵��_In     ����Ʊ�ݺ˶Լ�¼.�˶�˵��%Type
            strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("�˶�˵��")) & "')"
            cllPro.Add strSQL
        Next
    End With
    
    gcnOracle.BeginTrans: blnTran = True
    ExecuteProcedureArrAy cllPro, Me.Caption, True, True
    gcnOracle.CommitTrans: blnTran = False
    
    SaveTotalCheckData = True
    Exit Function
ErrHandler:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DetailCheck()
    '��ϸ�˶�
    Dim dtBegin As Date, dtEnd As Date, strErrMsg As String
    Dim str��Ʊ�� As String, bytMode As Byte '1-�˶Կ�Ʊ����Ʊ��2-���˶���Ʊ
    Dim strҵ������ As String, strSQL As String
    
    On Error GoTo ErrHandler
    '1.���ݼ��
    If vsfDetailCheck.Row < vsfDetailCheck.FixedRows Or vsfDetailCheck.Row > vsfDetailCheck.Rows - 1 Then Exit Sub
    
    strҵ������ = vsfTotalCheck.TextMatrix(vsfTotalCheck.Row, vsfTotalCheck.ColIndex("ҵ������"))
    If strҵ������ = "" Then Exit Sub
    
    dtBegin = strҵ������: dtEnd = Format(strҵ������, "yyyy-MM-dd 23:59:59")
     
    bytMode = IIf(opt��Ʊ����Ʊ.Value, 1, 2)
    str��Ʊ�� = zlStr.NeedCode(cbo��Ʊ��.Text)
    
    '2.��ȡ����
    Dim rsHISEInvoice As ADODB.Recordset
    If GetEInvoiceData(0, dtBegin, dtEnd, rsHISEInvoice, IIf(bytMode = 2, 2, 0), mbytƱ�ݺ˶�ʱ������, 0, "", str��Ʊ��) = False Then Exit Sub
    
    Dim clldata As Collection, cllDatas As Collection '����(ҵ������,ҵ����ˮ��,��Ʊ��,Ʊ����������,Ʊ�ݴ���,Ʊ�ݺ���,��Ʊ���,��Ʊʱ��,��������,����Ʊ�ݴ���,����Ʊ�ݺ���),Key=_ҵ����ˮ��
    If Not mobjEInvoice.ZlGetDetailCheckData(dtBegin, dtEnd, cllDatas, bytMode, str��Ʊ��, strErrMsg) Then
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Sub
    End If

    Dim rsUpdate As ADODB.Recordset '������¼
    strSQL = _
        " Select ����Ʊ��id, ҵ����ˮ��, ������ʽ, ������, ����ʱ��, �������, ����˵��, ƽ̨Ʊ��״̬ as Ʊ��״̬" & _
        " From ����Ʊ��������¼" & _
        " Where ҵ������ = [1]" & IIf(str��Ʊ�� = "", "", " And HIS��Ʊ��=[2]")
    Set rsUpdate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtBegin, str��Ʊ��)
    
    '3.��ʼ�˶�
    Dim strҵ����ˮ�� As String, blnChecked As Boolean, strCheckMsg As String
    Dim strHISƱ�ݺ��� As String, strƽ̨Ʊ�ݺ��� As String, dblHIS��Ʊ�� As Double, dblƽ̨��Ʊ�� As Double
    vsfDetailCheck.Clear 1
    vsfDetailCheck.Rows = 2
    vsfDetailCheck.Rows = vsfDetailCheck.FixedRows + 1
    
    With vsfDetailCheck
        .Redraw = flexRDNone
        
        Do While Not rsHISEInvoice.EOF
            'HIS����-�շ�ʱ��,HIS����-Ʊ������,HIS����-Ʊ��״̬,HIS����-����,HIS����-Ʊ�ݴ���,HIS����-Ʊ�ݺ���,HIS����-��Ʊ���,HIS����-��Ʊʱ��,
            'Ʊ��ƽ̨����-ҵ����ˮ��,Ʊ��ƽ̨����-��Ʊ��,Ʊ��ƽ̨����-Ʊ�ݴ���,Ʊ��ƽ̨����-Ʊ�ݺ���,Ʊ��ƽ̨����-��Ʊ���,
            'Ʊ��ƽ̨����-��Ʊʱ��,Ʊ��ƽ̨����-��������,Ʊ��ƽ̨����-Ʊ��״̬,Ʊ��ƽ̨����-����Ʊ�ݴ���,Ʊ��ƽ̨����-����Ʊ�ݺ���,
            '�˶Խ��,�˶�˵��,������ʽ,������,����ʱ��,�������,����˵��
            
            strҵ����ˮ�� = Nvl(rsHISEInvoice!����ID) & "_" & IIf(Val(Nvl(rsHISEInvoice!ԭƱ��ID)) = 0, Nvl(rsHISEInvoice!ID), Nvl(rsHISEInvoice!ԭƱ��ID))  '����Ʊ��ʹ�ü�¼.����ID_����Ʊ��ʹ�ü�¼.ID
            If .TextMatrix(.Rows - 1, .ColIndex("HIS����-�շ�ʱ��")) <> "" Then .Rows = .Rows + 1
            .RowData(.Rows - 1) = strҵ����ˮ��
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-�շ�ʱ��")) = Format(Nvl(rsHISEInvoice!�շ�ʱ��), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊ��")) = Nvl(rsHISEInvoice!��Ʊ��)
            .Cell(flexcpData, .Rows - 1, .ColIndex("HIS����-��Ʊ��")) = Nvl(rsHISEInvoice!��Ʊ��)
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-Ʊ������")) = Decode(Val(Nvl(rsHISEInvoice!Ʊ��)), 1, "�շ�", 2, "Ԥ��", 3, "����", 4, "�Һ�", 5, "���￨")
            .Cell(flexcpData, .Rows - 1, .ColIndex("HIS����-Ʊ������")) = Val(Nvl(rsHISEInvoice!Ʊ��))
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-Ʊ��״̬")) = Decode(Val(Nvl(rsHISEInvoice!Ʊ��״̬)), 1, "����", 2, "���", 3, "����")
            .Cell(flexcpData, .Rows - 1, .ColIndex("HIS����-Ʊ��״̬")) = Val(Nvl(rsHISEInvoice!Ʊ��״̬))
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-���ݺ�")) = Nvl(rsHISEInvoice!No)
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-Ʊ�ݴ���")) = Nvl(rsHISEInvoice!Ʊ�ݴ���)
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-Ʊ�ݺ���")) = Nvl(rsHISEInvoice!Ʊ�ݺ���)
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊ���")) = FormatEx(Val(Nvl(rsHISEInvoice!Ʊ�ݽ��)), 6, , , 2)
            .TextMatrix(.Rows - 1, .ColIndex("HIS����-��Ʊʱ��")) = Format(Nvl(rsHISEInvoice!��Ʊʱ��), "yyyy-MM-dd HH:mm:ss")
            strHISƱ�ݺ��� = Nvl(rsHISEInvoice!Ʊ�ݺ���): dblHIS��Ʊ�� = Val(Nvl(rsHISEInvoice!Ʊ�ݽ��))
            
            strҵ����ˮ�� = IIf(Val(Nvl(rsHISEInvoice!Ʊ��״̬)) = 2, 2, 1) & "_" & strҵ����ˮ��
            If CollectionExitsValue(cllDatas, strҵ����ˮ��) Then
                Set clldata = cllDatas(strҵ����ˮ��)
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-ҵ����ˮ��")) = Nvl(clldata("ҵ����ˮ��"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ��")) = Nvl(clldata("��Ʊ��"))
                .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ��")) = Nvl(clldata("��Ʊ��"))
                '.TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ����������")) = Nvl(clldata("Ʊ����������"))
                .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݴ���")) = Nvl(clldata("ҵ���ʶ"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݴ���")) = Nvl(clldata("Ʊ�ݴ���"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݺ���")) = Nvl(clldata("Ʊ�ݺ���"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ���")) = FormatEx(Val(Nvl(clldata("��Ʊ���"))), 6, , , 2)
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊʱ��")) = Format(CDateEx(Nvl(clldata("��Ʊʱ��"))), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��������")) = Decode(Val(Nvl(clldata("��������"))), 1, "��������", 2, "���Ӻ�Ʊ", 3, "����ֽ��", 4, "����ֽ�ʺ�Ʊ", 5, "�հ�ֽ��")
                .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ��ƽ̨����-��������")) = Val(Nvl(clldata("��������")))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ��״̬")) = Decode(Val(Nvl(clldata("Ʊ��״̬"))), 1, "����", 2, "����", 3, "���")
                .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ��״̬")) = Val(Nvl(clldata("Ʊ��״̬")))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-����Ʊ�ݴ���")) = Nvl(clldata("����Ʊ�ݴ���"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-����Ʊ�ݺ���")) = Nvl(clldata("����Ʊ�ݺ���"))
                strƽ̨Ʊ�ݺ��� = Nvl(clldata("Ʊ�ݺ���")): dblƽ̨��Ʊ�� = Val(Nvl(clldata("��Ʊ���")))
                
                blnChecked = False
                rsUpdate.Filter = "����Ʊ��id=" & Val(Nvl(rsHISEInvoice!ID)) & " And ҵ����ˮ��='" & strҵ����ˮ�� & "'"
                If Not rsUpdate.EOF Then
                    blnChecked = Val(Nvl(rsUpdate!�������)) = 1
                    .TextMatrix(.Rows - 1, .ColIndex("������ʽ")) = Decode(Val(Nvl(rsUpdate!������ʽ)), 1, "1-����HIS����", 2, "2-����ƽ̨����", 3, "3-���������ؿ�Ʊ��", 4, "4-�����������")
                    .TextMatrix(.Rows - 1, .ColIndex("������")) = Nvl(rsUpdate!������)
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Format(Nvl(rsUpdate!����ʱ��), "yyyy-MM-dd HH:mm:ss")
                    .TextMatrix(.Rows - 1, .ColIndex("�������")) = IIf(Val(Nvl(rsUpdate!�������)) = 1, "�����ɹ�", "����ʧ��")
                    .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = Nvl(rsUpdate!����˵��)
                End If
                
                strCheckMsg = ""
                If Not blnChecked Then
                    '�˶Թ���HISƱ�ݺ��� = ƽ̨Ʊ�ݺ��� And HIS��Ʊ��� = ƽ̨��Ʊ���
                    If strHISƱ�ݺ��� <> strƽ̨Ʊ�ݺ��� Then
                        strCheckMsg = strCheckMsg & "  HISƱ�ݺ��룺" & strHISƱ�ݺ��� & "/ƽ̨��Ʊ����" & strƽ̨Ʊ�ݺ���
                    End If
                    If dblHIS��Ʊ�� <> dblƽ̨��Ʊ�� Then
                        strCheckMsg = strCheckMsg & "  HIS��Ʊ��" & FormatEx(dblHIS��Ʊ��, 6, , , 2) & "/ƽ̨��Ʊ��" & FormatEx(dblƽ̨��Ʊ��, 6, , , 2)
                    End If
                    If strCheckMsg <> "" Then strCheckMsg = Mid(strCheckMsg, 2)
                    blnChecked = strCheckMsg = ""
                End If
                
                If blnChecked Then
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶Գɹ�"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = ""
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶�ʧ��"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = strCheckMsg
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    .TextMatrix(.Rows - 1, .ColIndex("������ʽ")) = IIf(Val(Nvl(rsHISEInvoice!Ʊ��״̬)) = 1, "3-���������ؿ�Ʊ��", "4-�����������")
                End If
                
                cllDatas.Remove strҵ����ˮ��
            Else
                blnChecked = False
                rsUpdate.Filter = "����Ʊ��id=" & Val(Nvl(rsHISEInvoice!ID))
                If Not rsUpdate.EOF Then
                    blnChecked = Val(Nvl(rsUpdate!�������)) = 1
                    .TextMatrix(.Rows - 1, .ColIndex("������ʽ")) = Decode(Val(Nvl(rsUpdate!������ʽ)), 1, "1-����HIS����", 2, "2-����ƽ̨����", 3, "3-���������ؿ�Ʊ��", 4, "4-�����������")
                    .TextMatrix(.Rows - 1, .ColIndex("������")) = Nvl(rsUpdate!������)
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Format(Nvl(rsUpdate!����ʱ��), "yyyy-MM-dd HH:mm:ss")
                    .TextMatrix(.Rows - 1, .ColIndex("�������")) = IIf(Val(Nvl(rsUpdate!�������)) = 1, "�����ɹ�", "����ʧ��")
                    .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = Nvl(rsUpdate!����˵��)
                End If
                
                If blnChecked Then
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶Գɹ�"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = ""
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶�ʧ��"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = "HISƱ�ݺ��룺" & strHISƱ�ݺ��� & "/ƽ̨Ʊ�ݺ��룺��"
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    .TextMatrix(.Rows - 1, .ColIndex("������ʽ")) = IIf(Val(Nvl(rsHISEInvoice!Ʊ��״̬)) = 1, "1-����HIS����", "4-�����������")
                End If
            End If
             
            rsHISEInvoice.MoveNext
        Loop
        
        If Not cllDatas Is Nothing Then
            For Each clldata In cllDatas
                If .TextMatrix(.Rows - 1, .ColIndex("HIS����-�շ�ʱ��")) <> "" Or .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-ҵ����ˮ��")) <> "" Then .Rows = .Rows + 1
                .RowData(.Rows - 1) = Nvl(clldata("ҵ����ˮ��"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-ҵ����ˮ��")) = Nvl(clldata("ҵ����ˮ��"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ��")) = Nvl(clldata("��Ʊ��"))
                .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ��")) = Nvl(clldata("��Ʊ��"))
                '.TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ����������")) = Nvl(clldata("Ʊ����������"))
                .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݴ���")) = Nvl(clldata("ҵ���ʶ"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݴ���")) = Nvl(clldata("Ʊ�ݴ���"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ�ݺ���")) = Nvl(clldata("Ʊ�ݺ���"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊ���")) = FormatEx(Val(Nvl(clldata("��Ʊ���"))), 6, , , 2)
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��Ʊʱ��")) = Format(CDateEx(Nvl(clldata("��Ʊʱ��"))), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-��������")) = Decode(Val(Nvl(clldata("��������"))), 1, "��������", 2, "���Ӻ�Ʊ", 3, "����ֽ��", 4, "����ֽ�ʺ�Ʊ", 5, "�հ�ֽ��")
                .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ��ƽ̨����-��������")) = Val(Nvl(clldata("��������")))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ��״̬")) = Decode(Val(Nvl(clldata("Ʊ��״̬"))), 1, "����", 2, "����", 3, "���")
                .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ��ƽ̨����-Ʊ��״̬")) = Val(Nvl(clldata("Ʊ��״̬")))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-����Ʊ�ݴ���")) = Nvl(clldata("����Ʊ�ݴ���"))
                .TextMatrix(.Rows - 1, .ColIndex("Ʊ��ƽ̨����-����Ʊ�ݺ���")) = Nvl(clldata("����Ʊ�ݺ���"))
                
                blnChecked = False
                rsUpdate.Filter = "ҵ����ˮ��='" & Nvl(clldata("ҵ����ˮ��")) & "' And Ʊ��״̬=" & Val(Nvl(clldata("Ʊ��״̬")))
                If Not rsUpdate.EOF Then
                    blnChecked = Val(Nvl(rsUpdate!�������)) = 1
                    .TextMatrix(.Rows - 1, .ColIndex("������ʽ")) = Decode(Val(Nvl(rsUpdate!������ʽ)), 1, "1-����HIS����", 2, "2-����ƽ̨����", 3, "3-���������ؿ�Ʊ��", 4, "4-�����������")
                    .TextMatrix(.Rows - 1, .ColIndex("������")) = Nvl(rsUpdate!������)
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Format(Nvl(rsUpdate!����ʱ��), "yyyy-MM-dd HH:mm:ss")
                    .TextMatrix(.Rows - 1, .ColIndex("�������")) = IIf(Val(Nvl(rsUpdate!�������)) = 1, "�����ɹ�", "����ʧ��")
                    .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = Nvl(rsUpdate!����˵��)
                End If
                
                If blnChecked Then
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶Գɹ�"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = ""
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("�˶Խ��")) = "�˶�ʧ��"
                    .TextMatrix(.Rows - 1, .ColIndex("�˶�˵��")) = "HISƱ�ݺ��룺��/ƽ̨Ʊ�ݺ��룺" & Nvl(clldata("Ʊ�ݺ���"))
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    .TextMatrix(.Rows - 1, .ColIndex("������ʽ")) = IIf(Val(Nvl(clldata("Ʊ��״̬"))) = 1 And Val(Nvl(clldata("��������"))) = 1, "2-����ƽ̨����", "4-�����������")
                End If
            Next
        End If
        
        If .Rows > .FixedRows And .Cols > .FixedCols Then     'ȱʡ��λ��
            .Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
            .Row = .FixedRows
        End If
            
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    vsfDetailCheck.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    lblҵ������.Caption = IIf(mbytƱ�ݺ˶�ʱ������ = 0, "��Ʊ����", "�շ�����")
    
    Call InitTotalCheckGrid
    Call InitDetailCheckGrid
    
    Call load��Ʊ��(cbo��Ʊ��, mrs��Ʊ��, mrs�շ�Ա)
    Call Get��Ʊ�����(UserInfo.ID, OS.ComputerName, mstrEInvoiceNodeCode)
    
    dtp����ʱ��.Value = zlDatabase.Currentdate
    dtp��ʼʱ��.Value = DateAdd("d", -7, dtp����ʱ��.Value)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    picMain.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If Button <> vbLeftButton Then Exit Sub
    If vsfTotalCheck.Height + Y < 1200 Or vsfDetailCheck.Height - Y < 1200 Then Exit Sub

    fraSplit.Top = fraSplit.Top + Y
    
    vsfTotalCheck.Height = vsfTotalCheck.Height + Y
    vsfDetailCheck.Top = vsfDetailCheck.Top + Y
    vsfDetailCheck.Height = vsfDetailCheck.Height - Y
    Me.Refresh
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    picFilter.Move 0, 0, picMain.ScaleWidth
    vsfTotalCheck.Move 0, picFilter.Top + picFilter.Height, picMain.ScaleWidth + 20, picMain.ScaleHeight * 2 / 3
    fraSplit.Move 0, vsfTotalCheck.Top + vsfTotalCheck.Height, picMain.ScaleWidth + 20
    vsfDetailCheck.Move 0, fraSplit.Top + fraSplit.Height, picMain.ScaleWidth + 20, picMain.ScaleHeight - (fraSplit.Top + fraSplit.Height) + 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    Set mcbsMain = Nothing
    Set mobjEInvoice = Nothing
    
    Set mrs��Ʊ�� = Nothing
    Set mrs�շ�Ա = Nothing
End Sub

Private Function InitTotalCheckGrid() As Boolean
    '��ʼ��VSFGrid���ؼ�
    Dim strHead As String, varData As Variant
    Dim strHead0 As String, varData0 As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '����1,���뷽ʽ1,�п�1|����2,���뷽ʽ2,�п�2|...
    strHead = "ҵ������,4,1000|��Ʊ��,7,1000|��Ʊ���,7,1200" & _
                    "|��Ʊ��,7,1000|��Ʊ���,7,1200|�ܱ���,7,1000" & _
                    "|�˶Խ��,1,1000|�˶�˵��,1,6000" & _
                    "|�ϴκ˶���,1,1000|�ϴκ˶�ʱ��,1,2000|�ϴκ˶Խ��,1,1200|�ϴκ˶�˵��,1,6000"
    strHead0 = "*|HIS����|HIS����|Ʊ��ƽ̨����|Ʊ��ƽ̨����|Ʊ��ƽ̨����|*|*|*|*|*|*"
    With vsfTotalCheck
        .Redraw = flexRDNone '��ͣ�����ʾˢ��
        .Clear
        .Rows = 3
        .FixedRows = 2: .FixedCols = 0

        varData = Split(strHead, "|"): varData0 = Split(strHead0, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = IIf(varData0(i) = "*", Split(varData(i), ",")(0), varData0(i))
            .TextMatrix(1, i) = Split(varData(i), ",")(0)
            .ColKey(i) = IIf(varData0(i) = "*", "", varData0(i) & "-") & Split(varData(i), ",")(0) '����Keyֵ,���ڸ��� ColIndex() ȷ����
            .ColWidth(i) = Split(varData(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = Split(varData(i), ",")(1)
        Next
        .Cell(flexcpText, 0, .ColIndex("ҵ������"), 1, .ColIndex("ҵ������")) = IIf(mbytƱ�ݺ˶�ʱ������ = 0, "��Ʊ����", "�շ�����")

        .AllowSelection = False '�������ѡ
        .AllowBigSelection = False '���������̶���/��ѡ������/����
        .SelectionMode = flexSelectionByRow '����ѡ��
        .AllowUserResizing = flexResizeColumns '�����û������п�
        .BackColorSel = &HE0E0E0
        .ForeColorSel = vbBlack
        
        .MergeCellsFixed = flexMergeFree
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeRow(1) = True
        
        .MergeCol(.ColIndex("ҵ������")) = True
        .MergeCol(.ColIndex("�˶Խ��")) = True
        .MergeCol(.ColIndex("�˶�˵��")) = True
        .MergeCol(.ColIndex("�ϴκ˶���")) = True
        .MergeCol(.ColIndex("�ϴκ˶�ʱ��")) = True
        .MergeCol(.ColIndex("�ϴκ˶Խ��")) = True
        .MergeCol(.ColIndex("�ϴκ˶�˵��")) = True

        .RowHeightMin = 300
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, .ColIndex("HIS����-��Ʊ���")) = .BackColorFixed
        
        Call ShowTotalRow
        
        .Redraw = flexRDBuffered 'ˢ�±����ʾ
    End With
    InitTotalCheckGrid = True
    Exit Function
ErrHandler:
    vsfTotalCheck.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitDetailCheckGrid() As Boolean
    '��ʼ��VSFGrid���ؼ�
    Dim strHead As String, varData As Variant
    Dim strHead0 As String, varData0 As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '����1,���뷽ʽ1,�п�1|����2,���뷽ʽ2,�п�2|...
    strHead = "�շ�ʱ��,4,1900|��Ʊ��,1,1000|Ʊ������,4,1000|Ʊ��״̬,4,1000|���ݺ�,1,1000|Ʊ�ݴ���,1,2000|Ʊ�ݺ���,1,2000|��Ʊ���,7,1200|��Ʊʱ��,4,1900" & _
                    "|ҵ����ˮ��,1,1000|��Ʊ��,1,1000|Ʊ�ݴ���,1,2000|Ʊ�ݺ���,1,2000|��Ʊ���,7,1200" & _
                    "|��Ʊʱ��,4,1000|��������,1,1000|Ʊ��״̬,1,1000|����Ʊ�ݴ���,1,2000|����Ʊ�ݺ���,1,2000" & _
                    "|�˶Խ��,1,1000|�˶�˵��,1,6000|������ʽ,1,2000|������,1,1000|����ʱ��,4,1900|�������,1,2000|����˵��,1,6000"
    strHead0 = "HIS����|HIS����|HIS����|HIS����|HIS����|HIS����|HIS����|HIS����|HIS����" & _
                    "|Ʊ��ƽ̨����|Ʊ��ƽ̨����|Ʊ��ƽ̨����|Ʊ��ƽ̨����|Ʊ��ƽ̨����" & _
                    "|Ʊ��ƽ̨����|Ʊ��ƽ̨����|Ʊ��ƽ̨����|Ʊ��ƽ̨����|Ʊ��ƽ̨����|*|*|*|*|*|*|*"
    With vsfDetailCheck
        .Redraw = flexRDNone '��ͣ�����ʾˢ��
        .Clear
        .Rows = 3
        .FixedRows = 2: .FixedCols = 0

        varData = Split(strHead, "|"): varData0 = Split(strHead0, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = IIf(varData0(i) = "*", Split(varData(i), ",")(0), varData0(i))
            .TextMatrix(1, i) = Split(varData(i), ",")(0)
            .ColKey(i) = IIf(varData0(i) = "*", "", varData0(i) & "-") & Split(varData(i), ",")(0) '����Keyֵ,���ڸ��� ColIndex() ȷ����
            .ColWidth(i) = Split(varData(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = Split(varData(i), ",")(1)
        Next

        .AllowSelection = False '�������ѡ
        .AllowBigSelection = False '���������̶���/��ѡ������/����
        .SelectionMode = flexSelectionByRow '����ѡ��
        .AllowUserResizing = flexResizeColumns '�����û������п�
        .BackColorSel = &HE0E0E0
        .ForeColorSel = vbBlack

        .MergeCellsFixed = flexMergeFree
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(.ColIndex("�˶Խ��")) = True
        .MergeCol(.ColIndex("�˶�˵��")) = True
        .MergeCol(.ColIndex("������ʽ")) = True
        .MergeCol(.ColIndex("������")) = True
        .MergeCol(.ColIndex("����ʱ��")) = True
        .MergeCol(.ColIndex("�������")) = True
        .MergeCol(.ColIndex("����˵��")) = True

        .RowHeightMin = 300
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, .ColIndex("HIS����-��Ʊ���")) = .BackColorFixed

        .Redraw = flexRDBuffered 'ˢ�±����ʾ
    End With
    InitDetailCheckGrid = True
    Exit Function
ErrHandler:
    vsfDetailCheck.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfDetailCheck_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    On Error Resume Next
    vsfDetailCheck.ForeColorSel = vsfDetailCheck.CellForeColor
End Sub

Private Sub vsfDetailCheck_GotFocus()
    Call SetActiveList(vsfDetailCheck)
End Sub

Private Sub vsfDetailCheck_LostFocus()
    Call SetActiveList(vsfDetailCheck, False)
End Sub

Private Sub vsfDetailCheck_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not (Me.ActiveControl Is vsfDetailCheck And Button = vbRightButton) Then Exit Sub
    RaiseEvent ShowPopupMenu(True)
End Sub

Private Sub vsfTotalCheck_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Or NewRow < vsfTotalCheck.FixedRows Then Exit Sub
    vsfDetailCheck.Clear 1
    vsfDetailCheck.Rows = vsfDetailCheck.FixedRows + 1
    
    On Error Resume Next
    vsfTotalCheck.ForeColorSel = vsfTotalCheck.CellForeColor
End Sub

Private Sub vsfTotalCheck_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not (Me.ActiveControl Is vsfTotalCheck And Button = vbRightButton) Then Exit Sub
    RaiseEvent ShowPopupMenu(True)
End Sub

Private Sub vsfTotalCheck_GotFocus()
    Call SetActiveList(vsfTotalCheck)
End Sub

Private Sub vsfTotalCheck_LostFocus()
    Call SetActiveList(vsfTotalCheck, False)
End Sub

Private Sub SetActiveList(vsfGrid As VSFlexGrid, Optional ByVal blnGetFocus As Boolean = True)
    '���ÿؼ�ѡ���б�������ɫ
    If blnGetFocus Then
        vsfTotalCheck.BackColorSel = &HE0E0E0
        vsfDetailCheck.BackColorSel = &HE0E0E0

        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &H8000000D '&HC0C0C0
    Else
        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &HE0E0E0
    End If
End Sub

Private Function Get��Ʊ�����(ByVal lng��Աid As Long, ByVal str�ͻ��� As String, ByRef str��Ʊ��_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ����
    '���:
    '����:str��Ʊ��_Out-���ؿ�Ʊ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-20 15:13:36
    '˵���������Ʊ��δ���ö��룬���Բ���Ա����Ϊ׼
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp  As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select 1 From Ʊ�ݿ�Ʊ����� where Rownum <2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ʊ�������Ϣ", lng��Աid, str�ͻ���)
     
    If rsTemp.RecordCount = 0 Then  'δ���ã�ȱʡΪ��ǰ����Ա���
        str��Ʊ��_Out = UserInfo.���
        Get��Ʊ����� = UserInfo.��� <> "": Exit Function
    End If
    

    strSQL = "" & _
    "   Select  nvl(A.��ԱID,0) as ��ԱID,nvl(A.�ͻ���,'-')  as �ͻ���,A.��Ʊ��ID,b.���� as ��Ʊ�����,B.����  " & _
    "   From Ʊ�ݿ�Ʊ����� A,����Ʊ�ݿ�Ʊ�� B " & _
    "   Where A.��Ʊ��ID=B.ID And nvl(B.����ʱ��,sysdate+1)>=SysDate And (a.��ԱID=[1] Or a.�ͻ���=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ʊ�������Ϣ", UserInfo.ID, str�ͻ���)
    
    '��ԱID+�ͻ���
    rsTemp.Filter = "��ԱID=" & lng��Աid & " And �ͻ���='" & str�ͻ��� & "'"
    If rsTemp.EOF = False Then
        str��Ʊ��_Out = Nvl(rsTemp!��Ʊ�����)
        Get��Ʊ����� = str��Ʊ��_Out <> ""
        Exit Function
    End If

    '���շ�Ա
    rsTemp.Filter = "��ԱID=" & lng��Աid & " And �ͻ���='-'"
    If rsTemp.EOF = False Then
        str��Ʊ��_Out = Nvl(rsTemp!��Ʊ�����)
        Get��Ʊ����� = str��Ʊ��_Out <> ""
        Exit Function
    End If
    
    '�ͻ���
    rsTemp.Filter = "�ͻ���='" & str�ͻ��� & "' And ��ԱID=0"
    If rsTemp.EOF = False Then
        str��Ʊ��_Out = Nvl(rsTemp!��Ʊ�����)
        Get��Ʊ����� = str��Ʊ��_Out <> ""
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

