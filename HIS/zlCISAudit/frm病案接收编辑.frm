VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm�������ձ༭ 
   Caption         =   "�������ձ༭"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11430
   Icon            =   "frm�������ձ༭.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   11430
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraCmd 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   5400
      Width           =   11415
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   3120
         TabIndex        =   18
         Top             =   285
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10200
         TabIndex        =   17
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "ȷ��(&O)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   9120
         TabIndex        =   16
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ(&P)"
         Height          =   350
         Left            =   8040
         TabIndex        =   13
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2520
         TabIndex        =   19
         Top             =   325
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H80000004&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton cmd������ 
         Height          =   300
         Left            =   11010
         Picture         =   "frm�������ձ༭.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "����סԺ��"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txtOuter 
         Height          =   300
         Left            =   6960
         TabIndex        =   5
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtInput 
         Height          =   300
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   4
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox txtSongMen 
         Height          =   300
         Left            =   8805
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboOutDept 
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgInDetail 
         Height          =   3735
         Left            =   0
         TabIndex        =   6
         Top             =   960
         Width           =   12495
         _cx             =   22040
         _cy             =   6588
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
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
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
         TabBehavior     =   1
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
      Begin MSComCtl2.DTPicker dtpOuterDate 
         Height          =   300
         Left            =   9120
         TabIndex        =   7
         Top             =   4920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483643
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   88145923
         CurrentDate     =   39799
      End
      Begin VB.Label lblOuter 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   6360
         TabIndex        =   21
         Top             =   4920
         Width           =   540
      End
      Begin VB.Label lblApplyDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   8130
         TabIndex        =   10
         Top             =   645
         Width           =   540
      End
      Begin VB.Label lblPurveyDept 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblOuterDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   8280
         TabIndex        =   8
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "�������ձ༭"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   11295
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   6240
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm�������ձ༭.frx":6C94
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15108
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm�������ձ༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
'Private mstrNo As String                    '����ĵ��ݺ�;
Private mintEditState As Integer            '1.������2���޸ģ�3���鿴��
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnChange As Boolean
Private mlngApplyId As Long                  '����ID
Private mstrPatientSum As String
Private mstrPrivs As String
Private mstrOldName As String
Private mlngCount As Long
Private mdtLend As Date
Private mblnInTo As Boolean
Private mstrDeptName As String
Private mintDblick As Integer
Private mlngModule  As Long

Public Sub ShowCard(frmMain As Form, ByVal intEditState As Integer, ByVal strPatientSum As String, Optional lngDeptId As Long = 0, Optional blnSuccess As Boolean = False, Optional ByVal lngModule As Long = 201)
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--����:��ʾ�ͱ༭��Ƭ
    '--����:frmMain-������
    '       intEditState-�༭״̬
    '       lngDeptId -��Ժ����ID
    '--����:blnSuccess-����ɹ�,true,����false
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Set mfrmMain = frmMain
    mintEditState = intEditState
    mlngApplyId = lngDeptId
    mstrPatientSum = strPatientSum
    mblnSuccess = blnSuccess
    mblnChange = False
    mlngModule = lngModule
    
    If mintEditState = 1 Then
        lblOuter.Enabled = True
        txtOuter.Enabled = True
        lblOuterDate.Enabled = True
        dtpOuterDate.Enabled = True
        With vfgInDetail
            .Editable = flexEDKbdMouse
        End With
        txtInput.Enabled = True
        chkInput.Enabled = True
        chkInput.Value = 1
    ElseIf mintEditState = 2 Then
        lblOuter.Enabled = True
        txtOuter.Enabled = True
        lblOuterDate.Enabled = True
        dtpOuterDate.Enabled = True
'        With vfgInDetail
'            .Editable = flexEDKbdMouse
'        End With
        cboOutDept.Enabled = False
        txtInput.Enabled = False
        chkInput.Enabled = False
    ElseIf mintEditState = 3 Then
        txtSongMen.Enabled = False
        lblOuter.Enabled = False
        txtOuter.Enabled = False
        lblOuterDate.Enabled = False
        dtpOuterDate.Enabled = False
        cboOutDept.Enabled = False
        txtSongMen.Enabled = False
        cmdSave.Caption = "�鿴(&V)"
        txtInput.Enabled = False
        chkInput.Enabled = False
    End If
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
End Sub

Private Sub cboOutDept_Change()
    mblnChange = True
End Sub

Private Sub cboOutDept_Click()
    Dim lngApplyId As Long
    
    If Me.cboOutDept.ListCount = 0 Then Exit Sub
    If Me.cboOutDept.ListIndex = -1 Then Exit Sub
    If cboOutDept.ItemData(cboOutDept.ListIndex) = 1 And cboOutDept.Text = "���в���" Then
        lngApplyId = 0
    Else
        lngApplyId = cboOutDept.ItemData(cboOutDept.ListIndex)
    End If
    
    If lngApplyId <> mlngApplyId Then
        If Not mblnInTo Then
            mlngApplyId = lngApplyId
            Exit Sub
        End If
        If ExaminData(vfgInDetail) Then
            If MsgBox("���ڳ�Ժ���ҷ����ı�,�Ƿ�Ҫ��������е�����(����ȡ���ı�)?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                mlngApplyId = lngApplyId
                mblnChange = False
                cmdSave.Enabled = False
                Call LoadvfgInDetailData(mintEditState)
            Else
                cboOutDept.ItemData(cboOutDept.ListIndex) = mlngApplyId
                cboOutDept.Text = mstrDeptName
            End If
        Else
            mlngApplyId = lngApplyId
            mblnChange = False
            cmdSave.Enabled = False
            Call LoadvfgInDetailData(mintEditState)
        End If
    End If
End Sub

Private Sub cboOutDept_GotFocus()
    With cboOutDept
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub

Private Sub cboOutDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(cboOutDept.Text) = "" Then
            If vfgInDetail.Enabled Then vfgInDetail.SetFocus
            Exit Sub
        End If
        cboOutDept.Text = Replace(UCase(cboOutDept.Text), "'", "")
        vRect = GetControlRect(cboOutDept.hWnd)
        
        strSQL = "" & _
        "   SELECT A.����, A.����, A.����,A.id " & _
        "   FROM  ���ű� A" & _
        "   Where ( TO_CHAR (A.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or A.����ʱ�� is null) AND A.ID in (" & _
        "         Select B.����ID From ��������˵�� B" & _
        "         Where (B.��������='�ٴ�' or B.��������='����') and (B.�������=2 or B.�������=3)) And " & _
        "         (A.���� like [1] or A.���� like [1] or A.���� like [1] or A.����||'-'||A.���� like [1] ) " & zl_��ȡվ������(True, "A") & _
        "         start with A.�ϼ�id is null connect by prior A.id=A.�ϼ�id"
        
        strTemp = Trim(cboOutDept.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        
        lngHeigth = cboOutDept.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ѡ��", False, cboOutDept.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp, glngUserId)
               
        If rsTemp Is Nothing Then
            If Not blnCancel Then MsgBox "û�����������Ŀ���,����[������Ϣ]!", vbInformation, gstrSysName
            If cboOutDept.Enabled Then
                cboOutDept.SetFocus
                cboOutDept.SelStart = 0
                cboOutDept.Text = mstrDeptName
                cboOutDept.SelLength = Len(cboOutDept.Text)
                Exit Sub
            End If
        End If
        With rsTemp
            If UCase(TypeName(cboOutDept)) = "COMBOBOX" Then
                cboOutDept = !���� & "-" & IIf(IsNull(!����), "", !����)
                mlngApplyId = !ID
                Call GetInitDept
                zlCommFun.PressKey vbKeyTab
'                If vfgInDetail.Enabled Then Me.vfgInDetail.SetFocus
            Else
                cboOutDept.SetFocus
                cboOutDept.SelStart = 0
                cboOutDept.SelLength = Len(cboOutDept.Text)
                If cboOutDept.Enabled Then cboOutDept.SetFocus
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub cboOutDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub chkInput_Click()
    txtInput.Enabled = IIf(chkInput.Value = 1, True, False)
End Sub

Private Sub chkInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdPrint_Click()
    printbill
End Sub

Private Sub cmdSave_Click()
     '�������ݱ��洦��
    Dim blnSuccess As Boolean
    Dim strBillPrint As String

    If mintEditState = 3 Then '�鿴
        '�����ӡ
        Unload Me
        Exit Sub
    End If
    
    If Not ExamineMtlBeData(vfgInDetail) Then Exit Sub
    If ExamineMtlDataRepeat(vfgInDetail) Then Exit Sub
    If Not ValidData Then Exit Sub
    
            
    blnSuccess = SaveInCard
    
    If blnSuccess = True Then
'        '�޸Ĺ���:����ɹ�:��Ҫ����Ƿ��Զ����
'        strBillPrint = "���̴�ӡ"
'

        If mlngModule = 201 Then
'           If IIf(Val(zlDatabase.GetPara(strBillPrint, glngSys, mlngModule)) = 1, 1, 0) = 1 Then
'               '��ӡ
'               printbill
'           End If
        Else
            Dim lngRow As Long
            Dim lngCurRow As Long
            For lngRow = 1 To vfgInDetail.Rows - 1
                If Val(vfgInDetail.TextMatrix(lngRow, vfgInDetail.ColIndex("����ID"))) > 0 Then
                    lngCurRow = lngCurRow + 1
                End If
            Next
            
            MsgBox "��ǰ�ܹ����ղ���: " & lngCurRow & " ��", vbInformation, "��ʾ"
            
            If IIf(Val(zlDatabase.GetPara("��ӡ�����嵥", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                Call zlRptPrint(1)
            End If
        End If
        
        If mintEditState = 2 Then  '�޸�
            Unload Me
            Exit Sub
        End If
'        stbThis.Panels(2).Text = "��һ�ŵĵ��ݺţ�" & mstrNo
    Else
        Exit Sub
    End If
    txtSongMen = ""
    txtInput = ""
    mstrPatientSum = ""
    Call LoadvfgInDetailData(mintEditState)
    mblnChange = False
    cmdSave.Enabled = False
End Sub
  
Private Function SaveInCard() As Boolean
    '----------------------------------------------------------------------------
    '--����:��������
    '--����:����ɹ�,����true,���򷵻�false
    '----------------------------------------------------------------------------
    Dim lngOutDeptId As Long
    Dim strSongMen As String
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strRecDate As String
    Dim strOuter As String
    Dim strOutDate As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim cllTemp As New Collection
    Dim strNow As String
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln���� As Boolean
    
    SaveInCard = False
    
    lngOutDeptId = cboOutDept.ItemData(cboOutDept.ListIndex)
    
    strRecDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:mm:ss")
    strSongMen = Trim(txtSongMen.Text)
    strOuter = Trim(txtOuter.Text)
    If strOuter = "" Then
        strOutDate = ""
    Else
        strOutDate = Format(dtpOuterDate, "yyyy-mm-dd HH:mm:ss")
    End If
    
    If mlngModule = 201 Then
        bln���� = (glngHIS����� > 0)
    Else
        bln���� = False
    End If
    
    If strNow = "" Then strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With vfgInDetail
        For intRow = 1 To .Rows - 1
            If Trim(.TextMatrix(intRow, .ColIndex("����ID"))) <> "" Then
                lngPatientlId = Val(.TextMatrix(intRow, .ColIndex("����ID")))
                lngMtyId = Val(.TextMatrix(intRow, .ColIndex("��ҳID")))
                'Create Or Replace Procedure Zl_�������ռ�¼_Insert
'                If mintEditState = 1 Then
'                    strSQL = "   Zl_�������ռ�¼_Insert("
'                Else
'                    strSQL = "   Zl_�������ռ�¼_Update("
'                End If
                '51584:������,2012-12-5,����ͬʱ��ɹ鵵
                If bln���� = True Then
                    '����°滤ʿվ���л����ļ��鵵
                    gstrSQL = "Select distinct nvl(Ӥ��,0) ��� From ���˻����ļ� where ����ID=[1] And ��ҳID=[2]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���˻����ļ�", lngPatientlId, lngMtyId)
                    Do While Not rsTemp.EOF
                        strSQL = "  ZL_���˻����ļ�_ARCHIVE("
                        strSQL = strSQL & "" & lngPatientlId & ","
                        strSQL = strSQL & "" & lngMtyId & ","
                        strSQL = strSQL & "" & Val(NVL(rsTemp!���)) & ",1)"
                        AddArray cllTemp, strSQL
                    rsTemp.MoveNext
                    Loop
                    '����ϰ滤ʿվ���л����ļ��鵵
                    gstrSQL = "Select distinct nvl(Ӥ��,0) ��� From ���˻����¼ where ����ID=[1] And ��ҳID=[2] And ������Դ = 2"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���˻����¼", lngPatientlId, lngMtyId)
                    Do While Not rsTemp.EOF
                        strSQL = "  Zl_���ӻ����¼_Archive("
                        strSQL = strSQL & "" & lngPatientlId & ","
                        strSQL = strSQL & "" & lngMtyId & ","
                        strSQL = strSQL & "" & Val(NVL(rsTemp!���)) & ","
                        strSQL = strSQL & "'" & UserInfo.���� & "',"
                        strSQL = strSQL & "To_Date('" & strNow & "','YYYY-MM-DD hh24:mi:ss'))"
                        AddArray cllTemp, strSQL
                    rsTemp.MoveNext
                    Loop
                    '�������סԺ�����ļ��鵵
                    gstrSQL = "select ID from ���Ӳ�����¼ where ����ID=[1] and ��ҳID=[2] And ��������=2 And ������Դ=2 And RowNum<2"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���Ӳ�����¼", lngPatientlId, lngMtyId)
                    If Not rsTemp.EOF Then
                        strSQL = "   Zl_���Ӳ�����¼_Archive(" & rsTemp!ID & ",0,1)"
                        AddArray cllTemp, strSQL
                    End If
                End If
                
                strSQL = "   Zl_�������ռ�¼_Insert("
                strSQL = strSQL & "" & lngPatientlId & ","
                strSQL = strSQL & "" & lngMtyId & ","
                strSQL = strSQL & "" & IIf(strSongMen = "", "NULL", "'" & strSongMen & "'") & ","
                strSQL = strSQL & "" & IIf(strOuter = "", "NULL", "'" & strOuter & "'") & ","
                strSQL = strSQL & "" & IIf(strOutDate = "", "NULL", "to_date('" & strOutDate & "','yyyy-mm-dd hh24:mi:ss')") & ","
                strSQL = strSQL & "" & IIf(strRecDate = "", "NULL", "to_date('" & strRecDate & "','yyyy-mm-dd hh24:mi:ss')") & ")"
                AddArray cllTemp, strSQL
                
                If mlngModule <> 201 Then
                '����ǵ��Ӳ�������,��Ҫ�ύ�����ύ��¼
                    strSQL = "zl_�����ύ��¼_Receive('" & Val(.TextMatrix(intRow, .ColIndex("�ύID"))) & "','" & strOuter & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                    AddArray cllTemp, strSQL
                End If
                
            End If
        Next
    End With
    
    Err = 0: On Error GoTo errHand:
    blnTrans = True
    ExecuteProcedureArrAy cllTemp, Me.Caption
    mblnSuccess = True
    mblnChange = False
    SaveInCard = True
    Exit Function
errHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    SaveInCard = False
End Function

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, 4)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdFind_Click()
    If lblName.Visible = False Then
        lblName.Visible = True
        txtName.Visible = True
        txtName.SetFocus
    Else
        txtName.Text = Replace(txtName.Text, "'", "")
        SearchRow vfgInDetail, vfgInDetail.ColIndex("סԺ��"), txtName.Text, True
        lblName.Visible = False
        txtName.Visible = False
    End If
End Sub

Private Sub cmd������_Click()
    Call SelectDoctor
End Sub

Private Sub Form_Load()
    mlngCount = 0
    mstrPrivs = gstrPrivs
    mblnInTo = False
    mintDblick = 0
    lblTitle = GetUnitName & lblTitle
    If Not GetInitDept Then Exit Sub
    
    If mintEditState = 1 Or mintEditState = 2 Then
        Me.txtOuter = gstrUserName
        Me.dtpOuterDate = Format(zlDatabase.Currentdate, "yyyy-MM-DD HH:mm:ss")
    End If
    
    If mlngModule = 201 Then
        Call LoadvfgInDetailData(mintEditState)
        cmd������.Visible = False
    Else
        Call LoadvfgInDetailAuditData(mintEditState)
        cmd������.Visible = True
    End If
    
    Me.cmdPrint.Visible = False
'    If mintEditState >= 3 Then
''        Me.cmdPrint.Visible = InStr(1, mstrPrivs, ";���ݴ�ӡ;") <> 0
'    Else
'        Me.cmdPrint.Visible = False
'    End If
    mblnInTo = True
    
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 7230 Then Me.Height = 7230
    If Me.Width < 11595 Then Me.Width = 11595
    
    With PicMain
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - fraCmd.Height  '- 100
    End With
    
    With lblTitle
        .Top = 120
        .Left = 0
        .Width = PicMain.Width
    End With
    
    With vfgInDetail
        .Top = 960
        .Left = 0
        .Width = PicMain.Width
        .Height = PicMain.Height - .Top - 720
    End With
    
    With txtSongMen
        .Left = vfgInDetail.Width - .Width - 420
        lblApplyDept.Left = .Left - lblApplyDept.Width - 60
        cmd������.Left = .Left + .Width + 60
    End With
       
    With dtpOuterDate
        .Top = vfgInDetail.Top + vfgInDetail.Height + 225
        .Left = PicMain.Width - PicMain.Left - dtpOuterDate.Width - 120
        lblOuterDate.Left = .Left - lblOuterDate.Width - 60
    End With
    
    With txtOuter
        .Top = vfgInDetail.Top + vfgInDetail.Height + 225
        .Left = lblOuterDate.Left - .Width - 120
        lblOuter.Left = .Left - lblOuter.Width - 60
    End With
    
    txtInput.Top = dtpOuterDate.Top
    
    lblOuter.Top = dtpOuterDate.Top + 60
    lblOuterDate.Top = lblOuter.Top
    chkInput.Top = lblOuter.Top
    
    With fraCmd
        .Top = PicMain.Top + PicMain.Height
        .Left = PicMain.Left
        .Width = PicMain.Width
    End With
    
    With cmdCancel
        .Left = fraCmd.Width - .Width - 375
    End With
    
    With cmdSave
        .Left = cmdCancel.Left - .Width - 105
    End With
    
    With cmdPrint
        .Left = cmdSave.Left - .Width - 105
    End With
End Sub

Private Function GetInitDept() As Boolean
    '----------------------------------------------------------------------------
    '����:��ȡ��Ժ����
    '����:���г�Ժ����,�򷵻�True,���򷵻�False
    '----------------------------------------------------------------------------
    Dim strSQL As String
    Dim i As Long
    Dim blnHaving As Boolean
    Dim rsApplys As New ADODB.Recordset
    
    strSQL = "" & _
    "   SELECT A.����, A.����, A.����,A.id " & _
    "   FROM  ���ű� A" & _
    "   Where ( TO_CHAR (A.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or A.����ʱ�� is null) AND A.ID in (" & _
    "         Select B.����ID From ��������˵�� B" & _
    "         Where (B.��������='�ٴ�' or B.��������='����') and (B.�������=2 or B.�������=3)) " & zl_��ȡվ������(True, "A") & _
    "         start with A.�ϼ�id is null connect by prior A.id=A.�ϼ�id"
    
    On Error GoTo errHandle
    Set rsApplys = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With rsApplys

        If .EOF Then
            GetInitDept = False
            Exit Function
        End If
    End With
    With Me.cboOutDept
        .Clear
        
        .AddItem "���п���"
        'װ������
        blnHaving = False
        mlngCount = rsApplys.RecordCount
        For i = 1 To rsApplys.RecordCount
            .AddItem rsApplys!���� & "-" & rsApplys!����
            .ItemData(.NewIndex) = rsApplys!ID
            If rsApplys!ID = mlngApplyId Then
                .ListIndex = .NewIndex
                blnHaving = True
            End If
            rsApplys.MoveNext
        Next
        rsApplys.Close
        If Not blnHaving Then
            .ListIndex = 0
        End If
    End With
    GetInitDept = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initVfgInHeadTitle()
    Dim strHead As String
    strHead = "���,600,1,1;סԺ��,1500,1,1;����,900,1,0;�Ա�,500,4,0;����,500,7,0;סԺ����,900,7,0;��Ժ����,1200,1,0;��Ժʱ��,1100,1,0;" & _
              "��Ժ����,1200,1,0;��Ժʱ��,1100,1,0;��������,1100,1,0;��ͥ��ַ,1350,1,0;����ID,0,7,-1;��ҳid,0,7,-1;�ύid,0,7,-1"
    Call SetVsFlexGridChangeHead(strHead, vfgInDetail, 1)
End Sub

Private Sub LoadvfgInDetailData(ByVal intEditState As Long)
    '��ȡ�����������
    Dim strSQL As String
    Dim strBillHead As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngApplyId As Long
    Dim lngGeneralId As Long
    Dim i As Long
    ' " From ������ҳ U, ������Ϣ X,�������ռ�¼ A,Table(Cast(f_Str2list('" & mstrPatientSum & "') As zlTools.t_Strlist)) B" & _
    '
    On Error GoTo errHandle
    If mintEditState = 1 Then
        strBillHead = " " & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����" & _
        " From ������ҳ U, ������Ϣ X,Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) B" & _
        " Where U.����id = X.����id And U.��ҳID <> 0 And U.��Ժ����id =[1] And U.����id = substr(B.Column_Value,1,instr(B.Column_Value,'_')-1) And " & _
        "       U.��ҳID = substr(B.Column_Value,instr(B.Column_Value,'_')+1)"
        strSQL = "" & _
        "   Select distinct A.����id,A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��ĿԱ, A.��Ŀ���� " & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, mstrPatientSum)
    Else
        strBillHead = " " & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', '�ѽ���', '�ѱ�Ŀ') As ״̬" & _
        " From ������ҳ U, ������Ϣ X,�������ռ�¼ A,Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) B" & _
        " Where U.����id = X.����id And U.��ҳID <> 0 And U.��Ժ����id =[1] And A.����id = substr(B.Column_Value,1,instr(B.Column_Value,'_')-1) And " & _
        "       A.��ҳID = substr(B.Column_Value,instr(B.Column_Value,'_')+1) And A.����id = U.����id And A.��ҳID = U.��ҳID"
        strSQL = "" & _
        "   Select distinct A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��ĿԱ, A.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��,A.״̬" & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, mstrPatientSum)
        
    End If
      
    With vfgInDetail
        .Clear
        Call initVfgInHeadTitle
        .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
        If Not rsTemp.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("���")) = i
                .TextMatrix(i, .ColIndex("סԺ��")) = IIf(IsNull(rsTemp!סԺ��), 0, rsTemp!סԺ��)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("�Ա�")) = IIf(IsNull(rsTemp!�Ա�), "", rsTemp!�Ա�)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("סԺ����")) = IIf(IsNull(rsTemp!��סԺ����), "", rsTemp!��סԺ����)
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                .TextMatrix(i, .ColIndex("��������")) = IIf(IsNull(rsTemp!��������), "", Format(rsTemp!��������, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("��ͥ��ַ")) = IIf(IsNull(rsTemp!��ͥ��ַ), "", rsTemp!��ͥ��ַ)
                .TextMatrix(i, .ColIndex("����ID")) = IIf(IsNull(rsTemp!����ID), 0, rsTemp!����ID)
                .TextMatrix(i, .ColIndex("��ҳid")) = IIf(IsNull(rsTemp!��ҳID), 0, rsTemp!��ҳID)
                If mintEditState <> 1 Then
                    txtSongMen = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                    txtOuter = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                    dtpOuterDate.Value = IIf(IsNull(rsTemp!����ʱ��), Format(zlDatabase.Currentdate, "yyyy-MM-DD HH:mm"), Format(rsTemp!����ʱ��, "yyyy-MM-DD HH:mm"))
''                    lngApplyId = IIf(IsNull(rsTemp!��Ժ����), 0, rsTemp!��Ժ����)
                End If
                rsTemp.MoveNext
            Next
            If intEditState = 1 Then
                .Rows = .Rows + 1
            End If
            cmdSave.Enabled = True
        Else
            Select Case intEditState
                Case 1
                    .Rows = .Rows + 1
            End Select
        End If
        If .Rows > 1 Then
            .Select 1, .ColIndex("סԺ��")
        End If
        If intEditState = 1 Then
            stbThis.Panels(2).Text = "�����ڡ�����סԺ�š����벡��סԺ��¼�룬�س����Ӳ���������Ϣ��"
        End If
        .ExplorerBar = flexExSortShowAndMove
        '��ѡ��
'        .FixedCols = 1
        .SelectionMode = flexSelectionByRow
    End With
    rsTemp.Close
    Call RestoreHead(vfgInDetail)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadvfgInDetailAuditData(ByVal intEditState As Long)
    '��ȡ�����������
    Dim strSQL As String
    Dim strBillHead As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngApplyId As Long
    Dim lngGeneralId As Long
    Dim i As Long
    ' " From ������ҳ U, ������Ϣ X,�������ռ�¼ A,Table(Cast(f_Str2list('" & mstrPatientSum & "') As zlTools.t_Strlist)) B" & _
    '
    On Error GoTo errHandle
    If mintEditState = 1 Then
        strBillHead = " " & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����,A.ID as �ύID" & _
        " From ������ҳ U, ������Ϣ X,�����ύ��¼ A,Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) B" & _
        " Where U.����id = X.����id And U.��ҳID <> 0 And " & IIf(mlngApplyId = 0, "0=[1]", "U.��Ժ����id =[1]") & " And U.����ID = A.����ID And U.��ҳID = A.��ҳID And A.��¼״̬=1 And U.����id = substr(B.Column_Value,1,instr(B.Column_Value,'_')-1) And " & _
        "       U.��ҳID = substr(B.Column_Value,instr(B.Column_Value,'_')+1)"
        strSQL = "" & _
        "   Select distinct A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��ĿԱ, A.��Ŀ����,A.�ύID " & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
           Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, mstrPatientSum)
    Else
        strBillHead = " " & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', '�ѽ���', '�ѱ�Ŀ') As ״̬" & _
        " From ������ҳ U, ������Ϣ X,�������ռ�¼ A,Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) B" & _
        " Where U.����id = X.����id And U.��ҳID <> 0 And " & IIf(mlngApplyId = 0, "0=[1]", "U.��Ժ����id =[1]") & "  And A.����id = substr(B.Column_Value,1,instr(B.Column_Value,'_')-1) And " & _
        "       A.��ҳID = substr(B.Column_Value,instr(B.Column_Value,'_')+1) And A.����id = U.����id And A.��ҳID = U.��ҳID"
        strSQL = "" & _
        "   Select distinct A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��ĿԱ, A.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��,A.״̬" & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, mstrPatientSum)
        
    End If
      
    With vfgInDetail
        .Clear
        Call initVfgInHeadTitle
        .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
        If Not rsTemp.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("���")) = i
                .TextMatrix(i, .ColIndex("סԺ��")) = IIf(IsNull(rsTemp!סԺ��), 0, rsTemp!סԺ��)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("�Ա�")) = IIf(IsNull(rsTemp!�Ա�), "", rsTemp!�Ա�)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("סԺ����")) = IIf(IsNull(rsTemp!��סԺ����), "", rsTemp!��סԺ����)
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                .TextMatrix(i, .ColIndex("��������")) = IIf(IsNull(rsTemp!��������), "", Format(rsTemp!��������, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("��ͥ��ַ")) = IIf(IsNull(rsTemp!��ͥ��ַ), "", rsTemp!��ͥ��ַ)
                .TextMatrix(i, .ColIndex("����ID")) = IIf(IsNull(rsTemp!����ID), 0, rsTemp!����ID)
                .TextMatrix(i, .ColIndex("��ҳid")) = IIf(IsNull(rsTemp!��ҳID), 0, rsTemp!��ҳID)
                .TextMatrix(i, .ColIndex("�ύid")) = IIf(IsNull(rsTemp!�ύId), 0, rsTemp!�ύId)

                If mintEditState <> 1 Then
                    txtSongMen = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                    txtOuter = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                    dtpOuterDate.Value = IIf(IsNull(rsTemp!����ʱ��), Format(zlDatabase.Currentdate, "yyyy-MM-DD HH:mm"), Format(rsTemp!����ʱ��, "yyyy-MM-DD HH:mm"))
''                    lngApplyId = IIf(IsNull(rsTemp!��Ժ����), 0, rsTemp!��Ժ����)
                End If
                rsTemp.MoveNext
            Next
            If intEditState = 1 Then
                .Rows = .Rows + 1
            End If
            cmdSave.Enabled = True
        Else
            Select Case intEditState
                Case 1
                    .Rows = .Rows + 1
            End Select
        End If
        If .Rows > 1 Then
            .Select 1, .ColIndex("סԺ��")
        End If
        If intEditState = 1 Then
            stbThis.Panels(2).Text = "�����ڡ�����סԺ�š����벡��סԺ��¼�룬�س����Ӳ���������Ϣ��"
        End If
        .ExplorerBar = flexExSortShowAndMove
        '��ѡ��
'        .FixedCols = 1
        .SelectionMode = flexSelectionByRow
    End With
    rsTemp.Close
    Call RestoreHead(vfgInDetail)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mblnChange Or mintEditState = 2 Or mintEditState = 4 Or mintEditState = 3 Then
        Call SaveHead(vfgInDetail)
        SaveWinState Me, App.ProductName
        Exit Sub
    End If
    
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        Call SaveHead(vfgInDetail)
        SaveWinState Me, App.ProductName
    End If
End Sub

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤���ݵ���Ч��
    '����:��֤�ɹ�����true,����false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ValidData = False
    Dim intLop As Integer
    
    If Trim(txtSongMen.Text) = "" Then
        MsgBox "�����˱�������!", vbInformation + vbOKOnly, gstrSysName
        If txtSongMen.Enabled Then txtSongMen.SetFocus
        Exit Function
    End If
    
    If InStr(1, txtSongMen.Text, "'") > 0 Then
        MsgBox "�����˴��ڷǷ��ַ�!", vbInformation + vbOKOnly, gstrSysName
        If txtSongMen.Enabled Then txtSongMen.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtSongMen.Text)) > 20 Then
        MsgBox "�����˳���,���������10�����ֻ�20���ַ�!", vbInformation + vbOKOnly, gstrSysName
        If txtSongMen.Enabled Then txtSongMen.SetFocus
        Exit Function
    End If
    
    If Trim(txtOuter.Text) = "" Then
        MsgBox "�����˱�������!", vbInformation + vbOKOnly, gstrSysName
        If txtOuter.Enabled Then txtOuter.SetFocus
        Exit Function
    End If
    
    If InStr(1, txtOuter.Text, "'") > 0 Then
        MsgBox "�����˴��ڷǷ��ַ�!", vbInformation + vbOKOnly, gstrSysName
        If txtOuter.Enabled Then txtOuter.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtOuter.Text)) > 20 Then
        MsgBox "�����˳���,���������10�����ֻ�20���ַ�!", vbInformation + vbOKOnly, gstrSysName
        If txtOuter.Enabled Then txtOuter.SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Sub txtInput_GotFocus()
    With txtInput
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strBillHead As String
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    Dim lngPatientId As Long
    Dim lngMtalId As Long
    Dim i As Long
    Dim j As Long
    Dim strMsg As String
    Dim strSection As String
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtInput.Text) = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txtInput.Text = Replace(UCase(txtInput.Text), "'", "")
        vRect = GetControlRect(txtInput.hWnd)
        
'        Rownum as ID
        If cboOutDept.Text = "���п���" Then
            strSection = ""
        Else
            strSection = " And U.��Ժ����id =[1] "
        End If
        
        If mlngModule = 201 Then
            strBillHead = " " & _
            " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
            "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
            "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
            "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����" & _
            " From ������ҳ U, ������Ϣ X" & _
            " Where U.����id = X.����id And U.�������� = 0 And U.��ҳID <> 0 And U.��Ŀ���� is null And U.��Ժ���� is not null " & strSection & " And U.סԺ�� = [2]"
            '62940:������,2013-06-24,�ѽ��վͲ����ٴν���
            If mintEditState = 1 Then
                strBillHead = strBillHead & _
                    " And NOT Exists (Select ID From �������ռ�¼ Where ����ID=U.����ID And ��ҳID=U.��ҳID)"
            Else
                strBillHead = strBillHead & _
                    " And Exists (Select ID From �������ռ�¼ Where ����ID=U.����ID And ��ҳID=U.��ҳID)"
            End If
            strSQL = "" & _
            "   Select distinct  Rownum as ID,A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
            "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
            "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����," & _
            "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��Ժ����id, A.��Ժ����id " & _
            "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
            "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
            "    Order by A.��Ժ���� desc "
        Else
            strBillHead = " " & _
            " Select Distinct X.����id, U.��ҳid, C.ID as �ύID,U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
            "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
            "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id,U.����״̬,Decode(Nvl(U.����״̬,1),1,'�ύ����',10,'���մ���',2,'�ܾ�����',3,'�������',4,'��鷴��',5,'���鵵',6,'�������',13,'���ڳ��',14,'��鷴��',16,'�������') as ����״ֵ̬," & _
            "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����" & _
            " From ������ҳ U, ������Ϣ X,�����ύ��¼ C" & _
            " Where U.����id = X.����id And U.����id = C.����ID And U.��ҳID = C.��ҳID And  C.��¼״̬ <>2 And U.��ҳID <> 0   " & strSection & "  And U.סԺ�� = [2]"
            strSQL = "" & _
            "   Select distinct  Rownum as ID,A.����id, A.��ҳid,A.�ύID, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
            "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
            "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����," & _
            "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��Ժ����id, A.��Ժ����id,A.����״̬,A.����״ֵ̬ " & _
            "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
            "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
            "    Order by A.��Ժ���� desc "
        End If
        
            
        strTemp = Trim(txtInput.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
'        strTemp = LfPBF & strTemp & RgPbf
        strTemp = strTemp
        lngHeigth = txtInput.Height
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyId, strTemp)
        If rsTemp.RecordCount <> 1 Then
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ѡ��", False, txtInput.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, mlngApplyId, strTemp)
        End If
        
        If mlngModule <> 201 Then
            
            If rsTemp Is Nothing Then
                MsgBox "��ǰ������δ�ύ��ǰָ������û�иò���,������Ϣ!", vbInformation, gstrSysName
                If txtInput.Enabled Then
                    txtInput.SetFocus
                    txtInput.SelStart = 0
                    txtInput.SelLength = Len(txtInput.Text)
                    
                    Exit Sub
                End If
            End If
            
            If rsTemp.RecordCount = 1 Then

                If rsTemp!����״̬ <> 1 Then
                    strMsg = "��ǰ����״̬Ϊ:[" & rsTemp!����״ֵ̬ & "],�����ڽ��н���!" & vbCrLf & vbCrLf
                    strMsg = strMsg & ChkStrUniCode("������" & IIf(IsNull(rsTemp!����), "", rsTemp!����) & "                    ", 20) & "סԺ�ţ�" & IIf(IsNull(rsTemp!סԺ��), 0, rsTemp!סԺ��)

                    MsgBox strMsg, vbInformation, gstrSysName
                    If txtInput.Enabled Then
                        txtInput.SetFocus
                        txtInput.SelStart = 0
                        txtInput.SelLength = Len(txtInput.Text)
                    End If

                    Exit Sub
                End If
            End If
        Else
        
            If rsTemp Is Nothing Then
                MsgBox "û�����������Ĳ���,������Ϣ!", vbInformation, gstrSysName
                If txtInput.Enabled Then
                    txtInput.SetFocus
                    txtInput.SelStart = 0
                    txtInput.SelLength = Len(txtInput.Text)
                    Exit Sub
                End If
            End If
        End If

        
        With rsTemp
            If UCase(TypeName(txtInput)) = "TEXTBOX" Then
                lngPatientId = IIf(IsNull(!����ID), 0, !����ID)
                lngMtalId = IIf(IsNull(!��ҳID), 0, !��ҳID)
                If Not ExamineInputRepeat(vfgInDetail, lngPatientId, lngMtalId) Then
                    i = 0
                    For j = 1 To vfgInDetail.Rows - 1
                        If IsNull(vfgInDetail.TextMatrix(j, vfgInDetail.ColIndex("����"))) Or vfgInDetail.TextMatrix(j, vfgInDetail.ColIndex("����")) = "" Then
                            i = j
                            j = vfgInDetail.Rows
                        End If
                    Next
                    If i = 0 Then
                        vfgInDetail.Rows = vfgInDetail.Rows + 1
                        vfgInDetail.Row = vfgInDetail.Rows - 1
                    End If
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("���")) = vfgInDetail.Rows - 1
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("סԺ��")) = IIf(IsNull(!סԺ��), 0, !סԺ��)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("����")) = IIf(IsNull(!����), "", !����)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("�Ա�")) = IIf(IsNull(!�Ա�), "", !�Ա�)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("����")) = IIf(IsNull(!����), "", !����)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("סԺ����")) = IIf(IsNull(!��סԺ����), "", !��סԺ����)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("��Ժ����")) = IIf(IsNull(!��Ժ����), "", !��Ժ����)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("��Ժʱ��")) = IIf(IsNull(!��Ժ����), "", Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("��Ժ����")) = IIf(IsNull(!��Ժ����), "", !��Ժ����)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("��Ժʱ��")) = IIf(IsNull(!��Ժ����), "", Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("��������")) = IIf(IsNull(!��������), "", Format(!��������, "yyyy-MM-dd"))
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("��ͥ��ַ")) = IIf(IsNull(!��ͥ��ַ), "", !��ͥ��ַ)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("����ID")) = IIf(IsNull(!����ID), 0, !����ID)
                    vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("��ҳid")) = IIf(IsNull(!��ҳID), 0, !��ҳID)
                    If mlngModule = 201 Then
                        vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("�ύid")) = 0
                    Else
                        vfgInDetail.TextMatrix(i, vfgInDetail.ColIndex("�ύid")) = IIf(IsNull(!�ύId), 0, !�ύId)
                    End If
                    
                    vfgInDetail.Rows = vfgInDetail.Rows + 1
                    vfgInDetail.Select i, vfgInDetail.ColIndex("סԺ��")
                    vfgInDetail.Row = i
                    mblnChange = True
                    cmdSave.Enabled = True
                    
                    txtInput.SetFocus
                    txtInput.SelStart = 0
                    txtInput.SelLength = Len(txtInput.Text)
                    If txtInput.Enabled Then txtInput.SetFocus
                    
                    
                Else
                    txtInput.SetFocus
                    txtInput.SelStart = 0
                    txtInput.SelLength = Len(txtInput.Text)
                    If txtInput.Enabled Then txtInput.SetFocus
                End If
            Else
                txtInput.SetFocus
                txtInput.SelStart = 0
                txtInput.SelLength = Len(txtInput.Text)
                If txtInput.Enabled Then txtInput.SetFocus
                zlCommFun.PressKey vbKeyTab
            End If
            
            
            
            
            
            .Close
        End With
        
        Call ShowReceiveNum
    
    End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtInput, KeyAscii, m����ʽ
'    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'        Exit Sub
'    Else
'        If KeyAscii = vbKeyReturn Then
'        Else
'            KeyAscii = 0
'        End If
'    End If
End Sub

Private Sub txtOuter_GotFocus()
    txtOuter.SelStart = 0
    txtOuter.SelLength = Len(txtOuter)
End Sub

Private Sub txtOuter_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtOuter.Text) = "" Then
            txtOuter.Text = gstrUserName
            If dtpOuterDate.Enabled Then dtpOuterDate.SetFocus
            If cmdSave.Enabled Then cmdSave.SetFocus
            Exit Sub
        End If
        txtOuter.Text = Replace(UCase(txtOuter.Text), "'", "")
        vRect = GetControlRect(txtOuter.hWnd)
        
        strSQL = "" & _
            "   Select ���,����,����,id " & _
            "   From ��Ա�� " & _
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) " & zl_��ȡվ������(True) & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
            "   order by ���"
            
        strTemp = Trim(txtOuter.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        
        lngHeigth = txtOuter.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Աѡ��", False, txtOuter.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp)
               
        If rsTemp Is Nothing Then
            MsgBox "û����������������,����[��Ա��Ϣ]!", vbInformation, gstrSysName
            If txtOuter.Enabled Then
                txtOuter.SetFocus
                txtOuter.SelStart = 0
                txtOuter.SelLength = Len(txtOuter.Text)
                Exit Sub
            End If
        End If
       
        With rsTemp
            If UCase(TypeName(txtOuter)) = "TEXTBOX" Then
                txtOuter = IIf(IsNull(!����), "", !����)
                mblnChange = True
                If cmdSave.Enabled Then Me.cmdSave.SetFocus
            Else
                txtOuter.SetFocus
                txtOuter.SelStart = 0
                txtOuter.SelLength = Len(txtOuter.Text)
                If txtOuter.Enabled Then txtOuter.SetFocus
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub txtOuter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtSongMen_GotFocus()
    With txtSongMen
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub

Private Sub txtSongMen_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtSongMen.Text) = "" Then
            txtSongMen.Text = gstrUserName
'            If dtpOuterDate.Enabled Then dtpOuterDate.SetFocus
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txtSongMen.Text = Replace(UCase(txtSongMen.Text), "'", "")
        vRect = GetControlRect(txtSongMen.hWnd)
        
        strSQL = "" & _
            "   Select ���,����,����,id " & _
            "   From ��Ա�� " & _
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) " & zl_��ȡվ������(True) & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
            "   order by ���"
            
        strTemp = Trim(txtSongMen.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        
        lngHeigth = txtSongMen.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Աѡ��", False, txtSongMen.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp)
               
        If rsTemp Is Nothing Then
'            MsgBox "û����������������,����[��Ա��Ϣ]!", vbInformation, gstrSysName
            If txtSongMen.Enabled Then
                txtSongMen.SetFocus
                txtSongMen.SelStart = 0
                txtSongMen.SelLength = Len(txtSongMen.Text)
                Exit Sub
            End If
        End If
       
        With rsTemp
            If UCase(TypeName(txtOuter)) = "TEXTBOX" Then
                txtSongMen = IIf(IsNull(!����), "", !����)
                mblnChange = True
                zlCommFun.PressKey vbKeyTab
            Else
                txtSongMen.SetFocus
                txtSongMen.SelStart = 0
                txtSongMen.SelLength = Len(txtSongMen.Text)
                If txtSongMen.Enabled Then txtSongMen.SetFocus
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub txtSongMen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub vfgInDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    With vfgInDetail
        Select Case Col
           Case vfgInDetail.ColIndex("סԺ��")
                strValue = Trim(.TextMatrix(Row, .ColIndex("סԺ��")))
                If Not IsNull(strValue) Then
                    If Not GetSelectMuchPurvey(vfgInDetail, strValue, 1) Then
                        .Select Row, Col
                        .TextMatrix(Row, .ColIndex("סԺ��")) = mstrOldName
                        mstrOldName = ""
                        Exit Sub
                    End If
                End If

        End Select
    End With
End Sub

Private Sub vfgInDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfgInDetail.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    If mintEditState = 1 Then
        Select Case Col
            Case vfgInDetail.ColIndex("סԺ��")
                mstrOldName = Trim(vfgInDetail.TextMatrix(Row, vfgInDetail.ColIndex("סԺ��")))
                Cancel = False
                Exit Sub
            Case Else
                Cancel = True
                Exit Sub
        End Select
    End If
End Sub

Private Sub vfgInDetail_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vfgInDetail.ColIndex("סԺ��")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Then
        Position = Col
    End If
End Sub

Private Sub vfgInDetail_BeforeSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridBeforeSort(vfgInDetail, Col, Order)
End Sub

Private Sub vfgInDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim i As Long
    Dim j As Long
    Dim strTemp As String
    Dim lngRow As Long
    Dim lngCurrRow As Long
    Dim blnRow As Boolean
    Dim strValue As String
    
    strValue = ""
    If InStr(vfgInDetail.Cell(flexcpText, 0, Col), "סԺ��") > 0 Then ' And mintDblick = 0
         Err = 0: On Error GoTo errHand:
        If Not GetSelectMuchPurvey(vfgInDetail, strValue, 2) Then
            vfgInDetail.Select Row, Col
            Exit Sub
        End If
        Exit Sub
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'

Private Sub vfgInDetail_Click()
    mintDblick = 0
End Sub

Private Sub vfgInDetail_DblClick()
    mintDblick = 1
End Sub

Private Sub vfgInDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If mintEditState = 1 Or mintEditState = 2 Then
        If KeyCode = vbKeyDelete Then
            If vfgInDetail.Row > 0 Then
                If MsgBox("��Ҫɾ����ǰ��¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    vfgInDetail.RemoveItem (vfgInDetail.Row)
                    If vfgInDetail.Row = 0 Then
                        vfgInDetail.Rows = vfgInDetail.Rows + 1
                        vfgInDetail.Select vfgInDetail.Rows - 1, vfgInDetail.Col
                    End If
                End If
            End If
            
            Call ShowReceiveNum
        End If
        
        If KeyCode = vbKeyInsert Then
            With vfgInDetail
                If MsgBox("��Ҫ���Ӽ�¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                   vfgInDetail.Rows = vfgInDetail.Rows + 1
                   .Select vfgInDetail.Rows - 1, vfgInDetail.Col
                End If
                
                If mintEditState = 1 Then
                    stbThis.Panels(2).Text = "�����ڡ�����סԺ�š����벡��סԺ��¼�룬�س����Ӳ���������Ϣ��" & " ��ǰ���ղ���: " & vfgInDetail.Rows - 2 & " ��"
                End If
            End With
        End If
    End If
    If KeyCode = vbKeyReturn Then
        If mintEditState = 1 Then
            Call zlPvVsMoveGridCell(vfgInDetail, vfgInDetail.ColIndex("סԺ��"), vfgInDetail.ColIndex("��ͥ��ַ"), True, lngRow, SetHeadCodeData(vfgInDetail))
        Else
            Call zlPvVsMoveGridCell(vfgInDetail, vfgInDetail.ColIndex("סԺ��"), vfgInDetail.ColIndex("��ͥ��ַ"), False, lngRow, SetHeadCodeData(vfgInDetail))
        End If
    End If
    
    If KeyCode <> vbKeyReturn Then
        vfgInDetail.ColComboList(vfgInDetail.ColIndex("סԺ��")) = ""
    End If
End Sub

Private Sub vfgInDetail_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngRow As Long
    If KeyCode = vbKeyReturn Then
        If mintEditState = 1 Then
            Call zlPvVsMoveGridCell(vfgInDetail, vfgInDetail.ColIndex("סԺ��"), vfgInDetail.ColIndex("��ͥ��ַ"), True, lngRow, SetHeadCodeData(vfgInDetail))
        Else
            Call zlPvVsMoveGridCell(vfgInDetail, vfgInDetail.ColIndex("סԺ��"), vfgInDetail.ColIndex("��ͥ��ַ"), False, lngRow, SetHeadCodeData(vfgInDetail))
        End If
    End If
End Sub

Private Sub vfgInDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0 'Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub vfgInDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     Select Case KeyAscii
        Case vbKeyReturn
        Case vbKeyBack, vbKeyEscape, 3, 22: Exit Sub
        Case Else
'            Select Case Col
'                Case vfgInDetail.ColIndex("����")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
''            End Select
''            Select Case Col
'                Case vfgInDetail.ColIndex("��������")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
''            End Select
''            Select Case Col
'                Case vfgInDetail.ColIndex("��ϴ����")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
'            End Select
    End Select
End Sub

Private Sub vfgInDetail_KeyUp(KeyCode As Integer, Shift As Integer)
     vfgInDetail.ColComboList(vfgInDetail.ColIndex("סԺ��")) = "..."
End Sub

Private Sub vfgInDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     vfgInDetail.ColComboList(vfgInDetail.ColIndex("סԺ��")) = "..."
End Sub

Private Function GetSelectMuchPurvey(ByRef vsGrid As VSFlexGrid, ByVal strSearch As String, ByVal intFlag As Integer) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '����:��������,������
    '����:strSearch-��������ֵ,
    '����:��ֻ����һ��ֵʱ����True,���򷵻�False
    '--------------------------------------------------------------------------------------------------------------
    Dim LfPBF As String
    Dim RgPbf As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strBillHead As String
    Dim strTemp As String
    Dim lngHeigth As Long
    Dim lngTop As Long
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim StrCodeName As String
    Dim lngPatientId As Long
    Dim lngMtalId As Long
    
    Dim lngRow As Long
    Dim i As Long
    Dim j As Long
    Dim strSection As String
    
    If intFlag = 1 Then
        If strSearch = "" Then Exit Function
        If InStr(1, strSearch, "'") <> 0 Then
            MsgBox "����ֵ�к��зǷ��ַ���", vbInformation, gstrSysName
            Exit Function
        End If
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        
        strSearch = Replace(UCase(strSearch), "'", "")
        
'        strTemp = LfPBF & strSearch & RgPbf
        strTemp = strSearch
    End If
    
    If cboOutDept.Text = "���п���" Then
        strSection = ""
    Else
        strSection = " And U.��Ժ����id =[1] "
    End If
        
    vRect = GetControlRect(vsGrid.hWnd)
    '62940:������,2013-06-24
    If mlngModule = 201 Then
        If intFlag = 1 Then
            strBillHead = " " & _
            " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
            "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
            "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
            "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����" & _
            " From ������ҳ U, ������Ϣ X" & _
            " Where U.����id = X.����id And U.�������� = 0 And U.��ҳID <> 0 And U.��Ŀ���� is null And U.��Ժ���� is not null " & strSection & " And U.סԺ�� = [2]"
        Else
            strBillHead = " " & _
            " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
            "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
            "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
            "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����" & _
            " From ������ҳ U, ������Ϣ X" & _
            " Where U.����id = X.����id And U.�������� = 0 And U.��ҳID <> 0 And U.��Ŀ���� is null And U.��Ժ���� is not null " & strSection
        End If
        '62940:������,2013-06-24,�ѽ��վͲ����ٴν���
        If mintEditState = 1 Then
            strBillHead = strBillHead & _
            " And NOT Exists (Select ID From �������ռ�¼ Where ����ID=U.����ID And ��ҳID=U.��ҳID)"
        Else
            strBillHead = strBillHead & _
            " And  Exists (Select ID From �������ռ�¼ Where ����ID=U.����ID And ��ҳID=U.��ҳID)"
        End If
        strSQL = "" & _
            "   Select distinct  Rownum as ID,A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
            "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
            "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����," & _
            "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��Ժ����id, A.��Ժ����id " & _
            "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
            "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
            "    Order by A.��Ժ���� desc "
    Else
        If intFlag = 1 Then
            strBillHead = " " & _
            " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
            "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
            "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
            "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����,A.ID as �ύID" & _
            " From ������ҳ U, ������Ϣ X,�����ύ��¼ A" & _
            " Where U.����id = X.����id And U.����ID = A.����ID And U.��ҳID = A.��ҳID And A.��¼״̬=1 And U.��ҳID <> 0 And U.��Ŀ���� is null " & strSection & " And U.סԺ�� = [2]"
        Else
            strBillHead = " " & _
            " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
            "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
            "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
            "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����,A.ID as �ύID" & _
            " From ������ҳ U, ������Ϣ X,�����ύ��¼ A" & _
            " Where U.����id = X.����id And U.��ҳID <> 0 And  U.����ID = A.����ID And U.��ҳID = A.��ҳID And A.��¼״̬=1 And U.��Ŀ���� is null " & strSection
        End If
        strSQL = "" & _
        "   Select distinct  Rownum as ID,A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��Ժ����id, A.��Ժ����id,A.�ύID " & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
    
    End If
    Err = 0
    On Error GoTo errHand:
    
    lngTop = vRect.Top + vsGrid.CellTop + vsGrid.CellHeight

    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ѡ��", False, vsGrid.Tag, "", False, False, True, vRect.Left - 15, lngTop, lngHeigth, blnCancel, False, False, mlngApplyId, strTemp)
    
    If rsTemp Is Nothing Then
        If Not blnCancel Then MsgBox "û�����������Ĳ�����Ϣ!", vbInformation, gstrSysName
        If vsGrid.Enabled Then
            vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("סԺ��")) = mstrOldName
            vsGrid.SetFocus
            GetSelectMuchPurvey = False
            Exit Function
        End If
    End If

    i = 1
    With rsTemp
        If UCase(TypeName(vsGrid)) = "VSFLEXGRID" Then
            With vsGrid
                lngPatientId = IIf(IsNull(rsTemp!����ID), 0, rsTemp!����ID)
                lngMtalId = IIf(IsNull(rsTemp!��ҳID), 0, rsTemp!��ҳID)
                If Not ExamineInputRepeat(vfgInDetail, lngPatientId, lngMtalId) Then
                    i = 0
                    For j = 1 To .Rows - 1
                        If IsNull(.TextMatrix(j, .ColIndex("����"))) Or .TextMatrix(j, .ColIndex("����")) = "" Then
                            i = j
                            j = .Rows
                        End If
                    Next
                    If i = 0 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    .TextMatrix(i, .ColIndex("���")) = .Rows - 1
                    .TextMatrix(i, .ColIndex("סԺ��")) = IIf(IsNull(rsTemp!סԺ��), 0, rsTemp!סԺ��)
                    .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(i, .ColIndex("�Ա�")) = IIf(IsNull(rsTemp!�Ա�), "", rsTemp!�Ա�)
                    .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(i, .ColIndex("סԺ����")) = IIf(IsNull(rsTemp!��סԺ����), "", rsTemp!��סԺ����)
                    .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                    .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                    .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                    .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss"))

                    .TextMatrix(i, .ColIndex("��������")) = IIf(IsNull(rsTemp!��������), "", Format(rsTemp!��������, "yyyy-MM-dd"))
                    .TextMatrix(i, .ColIndex("��ͥ��ַ")) = IIf(IsNull(rsTemp!��ͥ��ַ), "", rsTemp!��ͥ��ַ)
                    .TextMatrix(i, .ColIndex("����ID")) = IIf(IsNull(rsTemp!����ID), 0, rsTemp!����ID)
                    .TextMatrix(i, .ColIndex("��ҳid")) = IIf(IsNull(rsTemp!��ҳID), 0, rsTemp!��ҳID)
                    If mlngModule = 201 Then
                        .TextMatrix(i, vfgInDetail.ColIndex("�ύid")) = 0
                    Else
                        .TextMatrix(i, .ColIndex("�ύID")) = IIf(IsNull(rsTemp!�ύId), 0, rsTemp!�ύId)
                    End If
                   
                    
                    .Rows = .Rows + 1
                    .Select i, .ColIndex("סԺ��")
                    mblnChange = True
                    cmdSave.Enabled = True
                End If
            End With
            
            Call ShowReceiveNum
            
       Else
            .Close
            If vsGrid.Enabled Then vsGrid.SetFocus
            zlCommFun.PressKey vbKeyTab
        End If
        .Close
    End With
    GetSelectMuchPurvey = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RightHead(ByVal vsGrid As VSFlexGrid)

    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = GetControlRect(vsGrid.hWnd)
    lngLeft = vRect.Left + vsGrid.Left
    lngTop = vRect.Top + vsGrid.RowHeight(0) 'vsGrid.CellTop + vsGrid.CellHeight  '
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, vsGrid.RowHeight(0))
    Call SaveHead(vsGrid)
End Sub

Private Sub SaveHead(ByVal vsGrid As VSFlexGrid)
    zl_VsGrid_SaveToPara vsGrid, Me.Caption, mlngModule, "�������ձ༭��ͷ��Ϣ", True, False
End Sub

Private Sub RestoreHead(ByVal vsGrid As VSFlexGrid)
    zl_VsGrid_FromParaRestore vsGrid, Me.Caption, mlngModule, "�������ձ༭��ͷ��Ϣ", True, False
End Sub

Private Sub vfgInDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intGetHeight As Integer
    Dim intGetWidth As Integer
    
    intGetWidth = vfgInDetail.ColWidth(0)
    intGetHeight = vfgInDetail.RowHeight(0)
    If (Button = 2) Then
        If X < intGetWidth And Y < intGetHeight Then
            Call RightHead(vfgInDetail)
        End If
    End If
End Sub

Private Function ExamineMtlDataRepeat(ByRef vsGrid As VSFlexGrid) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤�����ݵ��Ƿ����ظ����Լ����˵ĳ�Ժʱ���Ƿ���ڽ���ʱ��
    '����:���ظ�����true,����false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim lngMeterlId As Long
    Dim strValue As String
        
    ExamineMtlDataRepeat = True
    With vsGrid
        For i = 1 To .Rows - 1
            For j = i To .Rows - 1
                If i <> j Then
                    If .TextMatrix(i, .ColIndex("����ID")) = .TextMatrix(j, .ColIndex("����ID")) And .TextMatrix(i, .ColIndex("��ҳid")) = .TextMatrix(j, .ColIndex("��ҳid")) Then
                        MsgBox "�ڵ�" & i & "�еĲ������" & j & "�Ĳ�����ͬ����ɾ������һ�е����ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            Next
        Next
    End With
    
    '62939:������,2013-06-24,��������ʱ�䲻��С�ڲ��˳�Ժʱ��
    With vsGrid
        For i = 1 To .Rows - 1
            If IsDate(.TextMatrix(i, .ColIndex("��Ժʱ��"))) Then
                If CDate(Format(.TextMatrix(i, .ColIndex("��Ժʱ��")), "YYYY-MM-DD HH:mm:ss")) > CDate(Format(dtpOuterDate.Value, "YYYY-MM-DD HH:mm:ss")) Then
                    MsgBox "�ڵ�" & i & "�еĲ�����Ժʱ����ڲ�������ʱ��,���飡", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    
    ExamineMtlDataRepeat = False
End Function

Private Function ExamineInputRepeat(ByRef vsGrid As VSFlexGrid, ByVal lngPatientId As Long, ByVal lngMtalId As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤�����ݵ��Ƿ����ظ�
    '����:���ظ�����true,����false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
           
    ExamineInputRepeat = True
    With vsGrid
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("����ID"))) = lngPatientId And Val(.TextMatrix(i, .ColIndex("��ҳid"))) = lngMtalId Then
                MsgBox "¼�벡�����б��е�" & i & "�Ĳ�����ͬ����¼���������������ݣ�", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End With
    
    ExamineInputRepeat = False
End Function

Private Function ExamineMtlBeData(ByRef vsGrid As VSFlexGrid) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤�Ƿ��������
    '����:�з���true,����false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    ExamineMtlBeData = False
    With vsGrid
        For i = 1 To .Rows - 1
            If Not IsNull(.TextMatrix(i, .ColIndex("����ID"))) And .TextMatrix(i, .ColIndex("����ID")) <> "" Then
                ExamineMtlBeData = True
                Exit Function
            End If
        Next
    End With
    MsgBox "�����ڵ����ݣ����ܽ��б��棡", vbInformation, gstrSysName
End Function

Private Function ExaminData(ByRef vsGrid As VSFlexGrid) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤�Ƿ��������
    '����:�з���true,����false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    ExaminData = False
    With vsGrid
        For i = 1 To .Rows - 1
            If Not IsNull(.TextMatrix(i, .ColIndex("����ID"))) And .TextMatrix(i, .ColIndex("����ID")) <> "" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                ExaminData = True
                Exit Function
            End If
        Next
    End With
End Function

'��ӡ����
Private Sub printbill()
    Dim strGetData As String
    strGetData = GetChoiceData(vfgInDetail)
'    ReportOpen gcnOracle, glngSys, "ZL4_BILL_361_1", Me, "���ݱ��=" & strNo, "��¼״̬=" & mintRecordState, "��λ=0", 2
End Sub

Private Function SetHeadCodeData(ByRef vsGrid As VSFlexGrid) As String
    Dim i As Long
    Dim strTemp As String
    
    SetHeadCodeData = ""
    With vsGrid
        For i = 0 To .Cols - 1
            If mintEditState = 1 Then
                If i = .ColIndex("סԺ��") Then
                    If IsNull(strTemp) Or strTemp = "" Then
                        strTemp = i & "||0"
                    Else
                        strTemp = strTemp & ";" & i & "||0"
                    End If
                End If
            End If
        Next
    End With
    SetHeadCodeData = strTemp
End Function

Private Function GetChoiceData(ByRef vsGrid As VSFlexGrid) As String
    Dim lngApplyId As String
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim lngRows As Long
    Dim strTemp As String
    Dim i As Long
    Dim j As Long
    Dim intCount As Integer
                    
    intCount = 0
    strTemp = ""
    GetChoiceData = ""
    With vsGrid
        If .Rows > 1 Then
            lngRows = .Rows - 1
            For i = 1 To lngRows
                lngPatientlId = Val(.TextMatrix(i, .ColIndex("����ID")))
                lngMtyId = Val(.TextMatrix(i, .ColIndex("��ҳID")))
                If Not IsNull(.TextMatrix(i, .ColIndex("����ID"))) And .TextMatrix(i, .ColIndex("����ID")) <> "" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                    intCount = intCount + 1
                    If intCount > 100 Then
                        GetChoiceData = strTemp
                        MsgBox "������Ĳ�����̫���ˣ�ֻ����ǰ��ѡ�е�100�ݡ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If strTemp = "" Then
                        strTemp = lngPatientlId & "_" & lngMtyId
                    Else
                        strTemp = strTemp & "," & lngPatientlId & "_" & lngMtyId
                    End If
                End If
            Next
        End If
    End With
    GetChoiceData = strTemp
End Function

'�����룬���ƣ���������ĳһ��
Private Function SearchRow(ByVal vfgBill As VSFlexGrid, ByVal intColIndex As Integer, _
    ByVal strColValue As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim strSQL As String
    Dim rsCode As New Recordset
    
    SearchRow = True
    With vfgBill
        If .Rows = 2 Then Exit Function
        If strColValue = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, intColIndex) <> "" Then
                strCode = .TextMatrix(intRow, intColIndex)
                If InStr(1, UCase(strCode), UCase(strColValue)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = intColIndex
                    .Select .Row, .Col
                    Exit Function
                End If
            End If
        Next
        
        On Error GoTo errHandle
        strSQL = "SELECT סԺ�� " _
                 & "FROM ������ҳ " _
               & " Where upper(����) LIKE '" & IIf(gstrMatchMethod = "0", "%", "") & strColValue & "%' "
        Set rsCode = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsCode.EOF Then
            SearchRow = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, intColIndex) <> "" Then
                strCode = .TextMatrix(intRow, intColIndex)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!סԺ��)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = intColIndex
                        rsCode.Close
                        .Select .Row, .Col
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    SearchRow = False
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub zlPvVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng���� As Long = -1, Optional lngβ�� As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1, Optional strValue As String)
    ', Optional strHeadMove As String
    '-----------------------------------------------------------------------------------------------------------

    '����:�ƶ���Ԫ�����
    '���:blnEdit-��ǰ�����ڱ༭״̬,����������
    '     lng����-����,���<0,������Ϊ0��,����Ϊָ������
    '     lngβ��-β��,���<0,������Ϊ.cols-1,����Ϊָ������
    '����:lngRow-������ڲ�����,�򷵻ر�������к�,���򷵻�-1
    '����:
    '����:���˺�
    '����:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------

    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    Dim lngValue As Long
    Dim arrHead As Variant
    Dim j As Long
    Dim lngColValue As Long
    
    Err = 0: On Error GoTo errHand:
    'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)

    If lng���� <> -1 Then
        lngCol = lng����
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lngβ�� < 0, vsGrid.Cols - 1, lngβ��)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                If IsNull(strValue) Or strValue = "" Then
                    arrSplit = Split(.ColData(i) & "||", "||")
                    If IsNull(arrSplit(1)) Or Trim(arrSplit(1)) = "" Then
                        lngValue = 0
                    Else
                        lngValue = Val(arrSplit(1))
                    End If
                Else
                    arrHead = Split(strValue, ";")
                    For j = 0 To UBound(arrHead)
                        lngValue = 1
                        lngColValue = Val(Split(arrHead(j), "||")(0))
                        If i = lngColValue Then
                            lngValue = Val(Split(arrHead(j), "||")(1))
                            Exit For
                        End If
                    Next
                End If
                If .ColHidden(i) Or lngValue >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
errHand:
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim objRow1 As New zlTabAppRow
    Dim strRange As String
    Dim intCol As Long
    Dim strListTitle As String
    
    strListTitle = "�����������Ͳ��˲������"
    With vfgInDetail
        '���ѡ���е���ɫ
''        For intCol = 0 To .Cols - 1
''            .Col = intCol
''            .CellBackColor = glngGetFocus_Font
'''            .CellForeColor = glngLostFocus_Font
''        Next
        .GridLines = flexGridInset
    End With
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strListTitle
        
    Set objRow = New zlTabAppRow

    If cboOutDept.Visible Then
        objRow.Add "��Ժ����:" & cboOutDept.Text
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow1 = New zlTabAppRow
    objRow1.Add "������:" & txtSongMen.Text
    objRow1.Add "������:" & txtOuter.Text
    objPrint.BelowAppRows.Add objRow1
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = vfgInDetail
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    With vfgInDetail
        .GridLines = flexGridNone
    End With
End Sub

Private Sub ShowReceiveNum()
'��ʾ��ǰ�����ղ����ķ���
    Dim lngRow As Long
    Dim lngCurRow As Long
    If mintEditState = 1 Then
        For lngRow = 1 To vfgInDetail.Rows - 1
            
            
            If Val(vfgInDetail.TextMatrix(lngRow, vfgInDetail.ColIndex("����ID"))) > 0 Then
                vfgInDetail.TextMatrix(lngRow, vfgInDetail.ColIndex("���")) = lngRow
                lngCurRow = lngCurRow + 1
            End If
        Next
    
        stbThis.Panels(2).Text = "�����ڡ�����סԺ�š����벡��סԺ��¼�룬�س����Ӳ���������Ϣ��" & " ��ǰ�ȴ����ղ���: " & lngCurRow & " ��"
    End If
End Sub

'ѡ��ҽ��
Private Sub SelectDoctor(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
    Dim strSQL As String
On Error GoTo errH
    strSQL = ""
    If strShortName <> "" Then
        strSQL = strSQL & vbCrLf & "Select c.ID,c.���,c.���� As ����"
        strSQL = strSQL & vbCrLf & "From ��Ա�� C"
        strSQL = strSQL & vbCrLf & "Where  c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null "
        strSQL = strSQL & vbCrLf & "And (c.���� like '%'||[1]||'%' or ���� like '%'||[1]||'%')"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strShortName))
        
        bytRet = ShowPubSelectTest(Me, txtSongMen, 2, "���,1200,0,;����,1200,0,", Me.Name & "\������ѡ��", "����±���ѡ��һ��ҽ��", rsTmp, rsResult, 8790, 4500, False)
    Else
        strSQL = strSQL & vbCrLf & "Select Id,�ϼ�id,0 As ĩ��,���� as ���,���� From ���ű�"
        strSQL = strSQL & vbCrLf & "Start With �ϼ�id Is Null"
        strSQL = strSQL & vbCrLf & "Connect By Prior ID = �ϼ�id"
        strSQL = strSQL & vbCrLf & "Union All"
        strSQL = strSQL & vbCrLf & "Select c.id,b.����id As �ϼ�Id,1 As ĩ��,c.���,c.���� As ����"
        strSQL = strSQL & vbCrLf & "From ������Ա b,��Ա�� C"
        strSQL = strSQL & vbCrLf & "Where c.Id=b.��Աid And b.ȱʡ=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
      
        bytRet = ShowPubSelectTest(Me, txtSongMen, 3, "���,1200,0,;����,1200,0,", Me.Name & "\������ѡ��", "����±���ѡ��һ��ҽ��", rsTmp, rsResult, 8790, 4500, False)
 
    End If
    
    If rsResult Is Nothing Then
        txtSongMen.Text = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txtSongMen.Text = ""
    Else
        txtSongMen.Text = rsResult!����
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub


