VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInTendEPR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3420
      Index           =   0
      Left            =   1035
      ScaleHeight     =   3420
      ScaleWidth      =   4650
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   390
      Width           =   4650
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmDockInTendEPR.frx":0000
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   3480
         Left            =   135
         TabIndex        =   1
         Top             =   945
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   6138
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockInTendEPR.frx":054E
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         Editable        =   2
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
      Begin VSFlex8Ctl.VSFlexGrid vfgWrit 
         Height          =   2310
         Left            =   45
         TabIndex        =   3
         Top             =   75
         Width           =   3735
         _cx             =   6588
         _cy             =   4075
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
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
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00FFEBD7&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   0
            Picture         =   "frmDockInTendEPR.frx":059C
            ScaleHeight     =   225
            ScaleMode       =   0  'User
            ScaleWidth      =   283.333
            TabIndex        =   4
            Top             =   0
            Width           =   250
         End
         Begin MSComctlLib.ImageList imgWrit 
            Left            =   1860
            Top             =   1005
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":6DEE
                  Key             =   "��д"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":7388
                  Key             =   "�޶�"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":7922
                  Key             =   "�鵵"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":7EBC
                  Key             =   "ת��"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":8256
                  Key             =   "��ӡ"
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   720
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   195
      Top             =   540
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmDockInTendEPR.frx":EAB8
      Left            =   120
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockInTendEPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private Enum mCol
    w��־ = 0: wID: wҳ����: wҳ������: w��������: w������: w����ʱ��: w������: w���ʱ��: w��ǰ�汾: wǩ������: w��ǰ���: w�鵵��: w�鵵����: w����ID: w������: w����״̬: w�༭��ʽ: wӤ��: w��ӡ
End Enum

Private mstrColWidthConfig As String
'
Private mstrPrivs As String                             '��ǰʹ���߶Ա�����(1255)��Ȩ�޴�
Private mblnSearch As Boolean                           '��ǰʹ�����Ƿ�߱���������(1273)Ȩ
Private mlngPatiId As Long                              '����id
Private mlngPageId As Long                              '��ҳid
Private mlngDeptId As Long                              '��ǰ��������id���粡�˿��Һ͵�ǰ���Ҳ�һ�£����ܲ����鵵��Ĺ���
Private mblnEdit As Boolean                             '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˲���������
Private mblnDoctorStation As Boolean                    '�Ƿ�ҽ��վ����
Private mblnMoved As Boolean                            '�Ƿ�ת��
Private mblnInsideTools As Boolean
Private WithEvents mfrmNew As frmDockEPRNew
Attribute mfrmNew.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1
Private mfrmMonitor As New frmDockEPRMonitor
Private mObjTabEpr As cTableEPR
Attribute mObjTabEpr.VB_VarHelpID = -1
Private mObjTabEprView As cTableEPR
Public Event Activate()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private mfrmTipInfo As New frmTipInfo
Private mblnViewTag As Boolean   'vfgWrit���б任�¼�ִ�б�־��true����ִ�У�falseû��ִ��
Private mblnViewNow As Boolean  'vfgWrit˫���¼���־��ture����ִ�У�falseû��ִ��
Private mlngCurId As Long
Public Function GetFormOperation() As String
'��¼����ѡ����Ϣ����Ϊ����վ���л�ҳ��ʱ���ͷ��˶��󣬻�����ʱ���³�ʼ��ˢ�µġ�
    GetFormOperation = mlngCurId
End Function

Public Sub RestoreFormOperation(ByVal strValue As String)
'�ָ�����ѡ����Ϣ������վ��ˢ��֮ǰ����
    mlngCurId = Val(strValue)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-С(ȱʡ)��1-��
Dim bytFontSize As Byte

    bytFontSize = Decode(bytSize, 0, 9, 1, 12, bytSize)
    Call mPublic.SetFontSize(Me, bytFontSize)
    Call mPublic.SetFontSize(mfrmNew, bytFontSize)
End Sub
Public Function InitData(ByVal strPrivs As String) As Boolean
    mstrPrivs = strPrivs
End Function

Public Function RefreshData(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal blnDoctorStation As Boolean, _
                            ByVal blnEdit As Boolean, Optional ByVal blnForce As Boolean, Optional ByVal blnMoved As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�ˢ������
    '������
    '���أ�
    '******************************************************************************************************************
    If mlngPatiId = lngPatiID And mlngPageId = lngPageId And blnForce = False Then Exit Function '��ǿ��ˢ�£�������ͬ��ˢ��
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '��ȡ�Ƿ񱾲������õ���ǩ��,���ұ����ûȡ��ʱ��ȡ
        gstrESign = getPassESign(4, lngDeptId)
    End If
    
    mlngPatiId = lngPatiID: mlngPageId = lngPageId: mblnEdit = blnEdit: mlngDeptId = lngDeptId
    
    mblnDoctorStation = blnDoctorStation: mblnMoved = blnMoved
    Call zlRefWrit
    
End Function
Public Sub zlDefCommandBars(ByVal cbsThis As Object, ByVal blnInsideTools As Boolean)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

    mblnInsideTools = blnInsideTools
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '�������:���ڵ�һ��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��(&O)��", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��", cbrControl.Index + 1)
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "�鵵(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "��������(&S)"): cbrControl.BeginGroup = True
    End With

    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", cbrMenuBar.Index, False)
        cbrMenuBar.ID = conMenu_ToolPopup
    End If
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "�����������(&M)"): cbrControl.BeginGroup = True
    End With
    
    '����������
    cbrMain.DeleteAll
    If mblnInsideTools Then
        Set cbrToolBar = cbrMain.Add("������", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        cbrToolBar.ContextMenuPresent = False
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "�鵵"): cbrControl.STYLE = xtpButtonIconAndCaption
        End With
    Else
        Set cbrToolBar = cbsThis(2)
        For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
            If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
                Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
            End If
        Next
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "�鵵", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��", 1)
            .Item(cbrControl.Index + 1).BeginGroup = True
        End With
    End If
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
    End With

    '-----------------------------------------------------
    '����Ȩ��״̬����ʾ���Ӵ���
    '-----------------------------------------------------
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "��������д") > 0) Then
        Me.dkpMain.Panes(3).Select
        Call mfrmNew.zlRefList(3, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs)
    End If
End Sub
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim strInfo As String, lFileId As Long
Dim bFinded As Boolean, frmThis As Form, bEditor As Byte
    If mblnMoved And (Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or _
                        Control.ID = conMenu_Edit_Audit Or Control.ID = conMenu_Edit_Archive * 10 + 1 Or _
                        Control.ID = conMenu_File_Open Or Control.ID = conMenu_File_ExportToXML) Then  '��ת������,�޸�,ɾ��,���,�鵵,�򿪲��������
        MsgBox "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                        "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lFileId = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID))
    bEditor = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w�༭��ʽ))
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open        '�����Ķ�
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If bEditor = 0 Then
            Dim fViewDoc As New frmEPRView, blnCanPrint As Boolean
            blnCanPrint = (InStr(1, mstrPrivs, "��������ӡ") > 0) And (Trim(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w�鵵��)) = "" Or InStr(1, mstrPrivs, "�鵵�������") > 0)
            fViewDoc.ShowMe Me, lFileId, , blnCanPrint
        Else
            If Not mObjTabEprView Is Nothing Then
                bFinded = mObjTabEprView.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
            End If
            If Not bFinded Then
                mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, True, 0, cprPF_סԺ, mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "��������ӡ") > 0, Val(gstrESign)
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) And InStr(mstrPrivs, "ȡ����ӡ") = 0 Then '�Ѿ���ӡ����û��ȡ����ӡȨ��,�������ظ���ӡ
            MsgBox "��ǰ�����Ѵ�ӡ���������ظ���ӡ��", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(True)
        Call zlRefWrit
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) And InStr(mstrPrivs, "ȡ����ӡ") = 0 Then '�Ѿ���ӡ����û��ȡ����ӡȨ��,�������ظ���ӡ
            MsgBox "��ǰ�����Ѵ�ӡ���������ظ���ӡ��", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(False)
        Call zlRefWrit
    Case conMenu_Edit_NoPrint 'ȡ����ӡ���
        If Split(EprIsCommit, "|")(0) = 0 Then
            MsgBox "�ò��˲������ύ��飬���ܳ�����ӡ����ȡ���������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        Call PrintCancel(CLng(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)))
        Call zlRefWrit
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_ExportToXML
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        '������XML�ļ�
        Dim strF As String
        dlgThis.Filename = "����_" & Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.w��������) & _
            "(" & Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID) & ").xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        strF = dlgThis.Filename
        On Error GoTo errHand
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If bEditor = 1 Then
                '���ʽ����
            mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, False, 0, cprPF_סԺ, _
                    mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs
            If mObjTabEprView.zlExportXML(strF) Then
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            Dim DocXML As New cEPRDocument '��ͨסԺ����
            DocXML.InitAndOpenEPR Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID), 0, , True
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem

        dkpMain.Panes(3).Select
        Call mfrmNew.zlRefList(3, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify                    '�޸Ļ����¼����
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
        lFileId = CLng(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID))
        If vfgWrit.TextMatrix(vfgWrit.Row, mCol.w�༭��ʽ) = 1 Then
            '���ʽ����
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, True, 0, cprPF_סԺ, _
                    mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "��������ӡ") > 0, Val(gstrESign)
                    mObjTabEpr.EPRPatiRecInfo.Ӥ�� = CLng(Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wӤ��)))
            End If
        Else
            '�������༭ģʽ
            Dim Doc As New cEPRDocument
            With Me.vfgWrit
                Doc.InitEPRDoc cprEM_�޸�, cprET_�������༭, .TextMatrix(.Row, mCol.wID), cprPF_סԺ, mlngPatiId, CStr(mlngPageId)
                Doc.EPRPatiRecInfo.Ӥ�� = CLng(Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wӤ��)))
                Doc.ShowEPREditor Me
            End With
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        If Split(EprIsCommit, "|")(1) = 0 Then
            MsgBox "�ò��˲������ύ��飬����ɾ������ȡ���������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
        With Me.vfgWrit
            strInfo = "���ɾ����ݡ�" & .TextMatrix(.Row, mCol.w��������) & "����"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "Zl_���Ӳ�����¼_Delete(" & .TextMatrix(.Row, mCol.wID) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call zlRefWrit
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
        If bEditor = 1 Then
            '���ʽ����
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_���������, lFileId, True, 0, cprPF_סԺ, _
                    mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "��������ӡ") > 0, Val(gstrESign)
            End If
        Else
            '���������ģʽ
            Dim frmAudit As Form, bFindedAudit As Boolean
            For Each frmAudit In Forms
                If frmAudit.Name = "frmMain" Then
                    If frmAudit.Document.EPRPatiRecInfo.ID = Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, 1) _
                        And frmAudit.Document.EPRPatiRecInfo.������Դ = cprPF_סԺ And frmAudit.Document.EPRPatiRecInfo.����ID = mlngPatiId _
                        And frmAudit.Document.EPRPatiRecInfo.��ҳID = mlngPageId And frmAudit.ChildMode = False Then
                        frmAudit.Show
                        bFindedAudit = True
                    End If
                End If
            Next
            If bFindedAudit = False Then
                '�״����
                Dim DocAudit As New cEPRDocument
                DocAudit.InitEPRDoc cprEM_�޸�, cprET_���������, Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, 1), cprPF_סԺ, mlngPatiId, CStr(mlngPageId)
                DocAudit.ShowEPREditor Me
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10 + 1

        With vfgWrit
            If Trim(.TextMatrix(.Row, mCol.w�鵵��)) = "" Then
                If Trim(.TextMatrix(.Row, mCol.w����״̬)) = "��Ժ" Then
                    strInfo = "��Ľ��÷ݡ�" & .TextMatrix(.Row, mCol.w��������) & "���鵵��"
                    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    gstrSQL = "Zl_���Ӳ�����¼_Archive(" & lFileId & ",0)"
                Else
                    strInfo = "�����Ѿ�" & Trim(.TextMatrix(.Row, mCol.w����״̬)) & "��Ҫ�����˱���סԺȫ���������鵵��" & vbCrLf _
                            & "  ѡ���ǡ����鵵���˱���ȫ����������" & vbCrLf _
                            & "  ѡ�񡰷񡱣����鵵�÷ݡ�" & .TextMatrix(.Row, mCol.w��������) & "����"
                    Select Case MsgBox(strInfo, vbQuestion + vbYesNoCancel + vbDefaultButton3, gstrSysName)
                    Case vbYes: gstrSQL = "Zl_���Ӳ�����¼_Archive(" & lFileId & ",0,1)"
                    Case vbNo: gstrSQL = "Zl_���Ӳ�����¼_Archive(" & lFileId & ",0)"
                    Case Else: Exit Sub
                    End Select
                End If
            Else
                strInfo = "��Ҫ�����ò��˱���סԺ�����ѹ鵵��������" & vbCrLf _
                        & "  ѡ���ǡ��������ò��˱���סԺ�����ѹ鵵��������" & vbCrLf _
                        & "  ѡ�񡰷񡱣��������÷ݡ�" & .TextMatrix(.Row, mCol.w��������) & "���Ĺ鵵��"
                Select Case MsgBox(strInfo, vbQuestion + vbYesNoCancel + vbDefaultButton3, gstrSysName)
                Case vbYes: gstrSQL = "Zl_���Ӳ�����¼_Archive(" & lFileId & ",1,1)"
                Case vbNo: gstrSQL = "Zl_���Ӳ�����¼_Archive(" & lFileId & ",1)"
                Case Else: Exit Sub
                End Select
            End If
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call zlRefWrit
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Sort
        '����
        Dim frmSort As New frmEPRSort
        If frmSort.ShowMe(Me, mlngPatiId, mlngPageId, cpr������, Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wҳ����)) = True Then
            'ˢ����ʾ
            Call zlRefWrit
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh

        Call zlRefWrit

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Monitor
        If mfrmMonitor.Visible = False Then mfrmMonitor.Show vbModeless, Me
        Call mfrmMonitor.zlRefList(mlngPatiId, mlngPageId, 4, mlngDeptId, 1, 1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search
        Call frmEPRSearchMan.ShowSearchClinic(Me, mlngDeptId)
    Case conMenu_Tool_SignVerify
        If bEditor = 0 Then
            Call VerifySignature(Me, lFileId, mblnMoved)
        Else '���ʽ������28δ��������ǩ�����
            'call
        End If
    End Select
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
LL:
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    Dim lngCount As Long, blnFinished As Boolean, lngMaxVersion As Long, eSignLevel As EPRSignLevelEnum

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open
        Control.Visible = True
        Control.Enabled = (Val(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID)) <> 0 And mblnEdit)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML

        Control.Enabled = (Val(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID)) <> 0)
        Control.Enabled = (Val(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID)) <> 0 And InStr(1, mstrPrivs, "��������ӡ") > 0)
        If Control.Enabled Then Control.Enabled = (Trim(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.w�鵵��)) = "" Or InStr(1, mstrPrivs, "�鵵�������") > 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NoPrint
        Control.Enabled = InStr(mstrPrivs, "ȡ����ӡ") > 0 And (Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) <> 0)
        If Control.Enabled Then Control.Enabled = Trim(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w��ӡ)) <> ""
        If Control.Enabled Then Control.Enabled = mblnEdit
    Case conMenu_File_Excel

        Control.Enabled = (Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) <> 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem
        Control.Enabled = (mblnEdit And mlngPatiId > 0)
        Control.Visible = (InStr(1, mstrPrivs, "��������д") > 0 And mblnDoctorStation = False)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "��������д") > 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
    
        Control.Enabled = (mblnEdit And mlngPatiId > 0)

        With Me.vfgWrit
            Control.Visible = (InStr(1, mstrPrivs, "��������д") > 0 Or InStr(1, mstrPrivs, "���˻�����") > 0) And mblnDoctorStation = False
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "��������д") > 0)
            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.w����ID)))   '���Ʋ����ſ��Ը�
            If Control.Enabled Then
                If Trim(.TextMatrix(.Row, mCol.w���ʱ��)) = "" Then
                    Control.Enabled = (InStr(1, mstrPrivs, "���˻�����") > 0 Or Trim(.TextMatrix(.Row, mCol.w������)) = Trim(gstrUserName))
                ElseIf Trim(.TextMatrix(.Row, mCol.w�鵵��)) = "" And Val(.TextMatrix(.Row, mCol.w��ǰ�汾)) <= 1 And InStr(1, ",1,2,4,", Val(.TextMatrix(.Row, mCol.wǩ������))) > 0 Then
                    Control.Enabled = (InStr(1, mstrPrivs, "���˻�����") > 0 Or InStr(1, .TextMatrix(.Row, mCol.w������), Trim(gstrUserName)) > 0)
                Else
                    Control.Enabled = False
                End If
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        Control.Enabled = (mblnEdit And mlngPatiId > 0)

        With vfgWrit

            Control.Visible = (InStr(1, mstrPrivs, "ǿ��ɾ������") > 0 Or InStr(1, mstrPrivs, "��������д") > 0 Or InStr(1, mstrPrivs, "���˻�����") > 0) And mblnDoctorStation = False

            Control.Enabled = (Val(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID)) <> 0)
            If Control.Enabled And InStr(1, mstrPrivs, "ǿ��ɾ������") > 0 Then Exit Sub '�߱�ǿ��ɾ��Ȩ�ޣ��򲻽��к������ж�
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "��������д") > 0)
            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.w����ID)))   '���Ʋ����ſ���ɾ
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.w���ʱ��)) = "")        'δ��ɲ�������ɾ
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "���˻�����") > 0 Or Trim(.TextMatrix(.Row, mCol.w������)) = Trim(gstrUserName))
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit

        Control.Visible = (InStr(1, mstrPrivs, "����������") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible)
        With vfgWrit
'                If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.w����ID)))   '���Ʋ����ſ�����
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.w���ʱ��)) <> "")       '��ɲ����ſ�����
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.w�鵵��)) = "")          'δ�鵵����������
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10 + 1

        Control.Visible = (InStr(1, mstrPrivs, "�������鵵") > 0 And mblnDoctorStation = False)

        'ֻ���Ѿ���ɵ�δ�鵵�Ĳ���,���ܽ��й鵵����
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible)
        With Me.vfgWrit
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.wǩ������)) <> 0)         '��ǰ�汾�Ѿ�ǩ����ɲſ��Թ鵵
            If Trim(.TextMatrix(.Row, mCol.w�鵵��)) = "" Then
                Control.Caption = "�����鵵": Control.Checked = False
            Else
                Control.Caption = "��������": Control.Checked = True
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup

        Control.Visible = (mblnDoctorStation = False And (InStr(1, mstrPrivs, "�������鵵") > 0 _
                                                        Or InStr(1, mstrPrivs, "����������") > 0 _
                                                        Or InStr(1, mstrPrivs, "��������д") > 0 _
                                                        Or InStr(1, mstrPrivs, "ǿ��ɾ������") > 0))
        Control.Enabled = Control.Visible
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Monitor
        Control.Visible = True
        Control.Enabled = (mlngPatiId > 0 And InStr(1, mstrPrivs, "���������") > 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search
        Control.Visible = True
        Control.Enabled = mblnSearch
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Sort
        '����ֻ�ж��ĵ�����ҳ��ʱ�ſ��Ե�����ţ�
        Dim R1&, C1&, R2&, C2&
        vfgWrit.GetMergedRange vfgWrit.Row, mCol.wҳ������, R1, C1, R2, C2
        Control.Enabled = (R1 <> R2)
        Control.Visible = Control.Enabled
    Case conMenu_Tool_SignVerify
        Control.Enabled = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) <> 0 And Trim(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w���ʱ��)) <> ""
    End Select
End Sub

Public Sub RefreshList()
    Call zlRefWrit
End Sub
Private Function InitColumnSelect() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    On Error Resume Next
    '���ܣ�����ԭʼ����ʾ״̬��ʼ����ѡ����
    Dim lngRow As Long, i As Long

    vsColumn.Rows = vsColumn.FixedRows
    With vfgWrit
        For i = .FixedCols To .Cols - 1
            Select Case i
            Case mCol.w��������, mCol.w������, mCol.w����ʱ��, mCol.w������, mCol.w���ʱ��, mCol.w��ǰ���, mCol.w������
                 vsColumn.Rows = vsColumn.Rows + 1
                 lngRow = vsColumn.Rows - 1
                 vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                 vsColumn.RowData(lngRow) = i

                 '�̶���ʾ��
                 If InStr(",ҳ������,��������,", "," & .TextMatrix(0, i) & ",") > 0 Then
                     vsColumn.TextMatrix(lngRow, 0) = 1
                     vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                 End If
            End Select
        Next
    End With
    vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 130
    vsColumn.Row = 1

    InitColumnSelect = True

End Function

Private Sub zlEPRPrint(blnPreview As Boolean)
Dim lFileId As Long
Dim frmP As New frmPrintPreview, r As String, blnOrigMode As Boolean '�Ƿ���ʾԭʼ״̬
    If vfgWrit.TextMatrix(vfgWrit.Row, mCol.w�༭��ʽ) = 0 Then
        r = zlCommFun.ShowMsgBox("����Ԥ��/��ӡ", "��ѡ����Ԥ��/��ӡ�ĸ�ʽ��", "!���ո�ʽ(&F),ԭʼ��ʽ(&O),ȡ��(&C)", Nothing)
        If r = "���ո�ʽ" Then
            blnOrigMode = False
        ElseIf r = "ԭʼ��ʽ" Then
            blnOrigMode = True
        Else
            Exit Sub
        End If
        frmP.DoMultiDocPreview Me, cpr������, mlngPatiId, mlngPageId, cpr������, Me.vfgWrit.Cell(flexcpText, Me.vfgWrit.Row, mCol.wҳ����), Me.vfgWrit.Cell(flexcpText, Me.vfgWrit.Row, mCol.wID), Not blnPreview, blnOrigMode, , mblnMoved
        Unload frmP 'ByZT:����Load��δ��ʾ��û����Ϊ�رյ������VB�����Զ�Unload
        Set frmP = Nothing
    Else
        lFileId = CLng(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID))
        mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_���������, lFileId, False, 0, cprPF_סԺ, mlngPatiId, mlngPageId, , mlngDeptId, , mstrPrivs, mblnMoved, InStr(mstrPrivs, "��������ӡ") > 0
        mObjTabEprView.zlPrintDoc Me, blnPreview
    End If
End Sub

Private Sub zlRefWrit()
'---------------------------------------------
'������ˢ��
'---------------------------------------------
Dim lngCurId As Long    'ˢ��ǰѡ�еĲ�����¼ID
Dim lngCurRow As Long
Dim lngCol As Long
Dim lngRow As Long
Dim rsTemp As New ADODB.Recordset
    
    vsColumn.Visible = False
    vfgWrit.Tag = ""
    Call mfrmContent.Clear
    
    gstrSQL = "Select r.Id, f.���, Decode(f.ҳ��, Null, r.��������, f.ҳ��) As ҳ��, r.��������, r.������ As ������," & _
            "        To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, r.������," & _
            "        To_Char(r.���ʱ��, 'yyyy-mm-dd hh24:mi') As ���ʱ��, r.���汾 As ��ǰ�汾, r.ǩ������," & _
            "        Decode(r.���汾, 1, '��д��', '�޶���') || r.������ || '��' || To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') ||" & _
            "         Decode(Nvl(r.ǩ������, 0), 0, '����(δ���)', 1, '���', '��ǩ') As ��ǰ���, r.�鵵��, r.�鵵����," & _
            "        r.����id As ����id, d.���� As ����, p.����״̬,r.�༭��ʽ,r.Ӥ��,r.��ӡ�� as ��ӡ" & _
            " From ���Ӳ�����¼ r, ���ű� d," & _
            "      (Select Decode(��Ժ����, Null, Decode(״̬, 3, 'Ԥ��Ժ', '��Ժ'), '��Ժ') As ����״̬" & vbNewLine & _
            "        From ������ҳ" & vbNewLine & _
            "        Where ����id = [1] And ��ҳid = [2]) p," & _
            "      (Select d.Id As �ļ�id, f.����, f.���, f.���� As ҳ��, d.����" & _
            "        From �����ļ��б� d, ����ҳ���ʽ f" & _
            "        Where d.���� = 4 And d.���� = f.���� And d.ҳ�� = f.���) f" & _
            " Where r.�ļ�id = f.�ļ�id(+) And r.������Դ = 2 And r.�������� = 4 And r.����id = d.Id And r.����id = [1] And r.��ҳid = [2]" & _
            " Order By r.��������, f.���, r.���, r.����ʱ��"
    Err = 0: On Error GoTo errHand
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
    
    With Me.vfgWrit
        Err = 0: On Error Resume Next
        lngCurId = Val(.TextMatrix(.Row, mCol.wID))
        If lngCurId = 0 Then lngCurId = mlngCurId
        .Clear
        Set .DataSource = rsTemp

        .MergeCells = flexMergeFree: .MergeCellsFixed = flexMergeFree
        .MergeCol(mCol.wҳ������) = True

        Dim T As Variant, i As Long
        On Error Resume Next
        T = Split(mstrColWidthConfig, ";")
        If UBound(T) < 18 Then
            mstrColWidthConfig = "270;0;0;1200;2000;800;1600;800;0;800;0;3300;0;0;0;1200;0;0;0"
        Else
            For i = 0 To .Cols - 1
                .ColWidth(i) = T(i)
                .ColHidden(i) = (.ColWidth(i) = 0)
            Next
        End If
        .TextMatrix(0, mCol.wҳ������) = .TextMatrix(0, mCol.w��������)
        .MergeRow(0) = True
        For lngCol = .FixedCols To .Cols - 1
            .FixedAlignment(lngCol) = flexAlignCenterCenter
        Next
        For lngRow = .FixedRows To .Rows - 1
            .MergeRow(lngRow) = True
            If Trim(.TextMatrix(lngRow, mCol.w�鵵��)) <> "" Then
                Set .Cell(flexcpPicture, lngRow, mCol.w��־) = imgWrit.ListImages("�鵵").Picture
            ElseIf Val(.TextMatrix(lngRow, mCol.w��ǰ�汾)) <= 1 Then
                Set .Cell(flexcpPicture, lngRow, mCol.w��־) = imgWrit.ListImages("��д").Picture
            Else
                Set .Cell(flexcpPicture, lngRow, mCol.w��־) = imgWrit.ListImages("�޶�").Picture
            End If
            If Trim(.TextMatrix(lngRow, mCol.w��ӡ)) <> "" Then
                Set .Cell(flexcpPicture, lngRow, mCol.wҳ������) = imgWrit.ListImages("��ӡ").Picture
            End If
            If lngCurId = Val(.TextMatrix(lngRow, mCol.wID)) Then lngCurRow = lngRow
        Next
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngCurRow = 0 Then
            .Row = 0 '��ʹvfgthis��ѡ���κ��У�����ʾ�κ����ݣ�����ѡ��ĳ��ʱ��ˢ��
        Else
            .Row = lngCurRow
        End If
        Call vfgWrit_RowColChange
    End With

    Call InitColumnSelect '��ѡ����
    
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "��������д") > 0) Then
        Me.dkpMain.Panes(3).Select
        Call mfrmNew.zlRefList(3, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs)
    End If
    'vfgWrit.Cell(flexcpWidth, mCol.w��ӡ) = 0
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    zlExecuteCommandBars Control
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    zlUpdateCommandBars Control
End Sub

'######################################################################################################################
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hwnd
    Case 2
        If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
        Item.Handle = mfrmContent.hwnd
    Case 3
        If mfrmNew Is Nothing Then Set mfrmNew = New frmDockEPRNew
        Item.Handle = mfrmNew.hwnd
    End Select
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    If vsColumn.Visible Then
        vsColumn.SetFocus '��ѡ����
    Else
        If Me.vfgWrit.Visible Then Me.vfgWrit.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    vsColumn.Visible = False '��ѡ����
End Sub

Private Sub Form_Load()
 Dim objPane As Pane, lngFontSize As Long
    On Error GoTo errHand
    
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "����") > 0)

    mstrColWidthConfig = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", _
        "270;0;0;1200;2000;800;1600;800;0;800;0;3300;0;0;0;1200;0;0;0")
    
    lngFontSize = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgWrit.Name, "FontSize", 9)
    vfgWrit.FontSize = lngFontSize
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With Me.cbrMain
        .VisualTheme = xtpThemeOffice2003
        Set .Icons = zlCommFun.GetPubIcons
        .ActiveMenuBar.Visible = False
        .EnableCustomization False
        With .Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '����VisualTheme����Ч
            .UseDisabledIcons = True
            .LargeIcons = True
            .SetIconSize True, 24, 24
        End With
    End With
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    dkpMain.SetCommandBars cbrMain
    
    Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing): objPane.Title = "�����б�": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 500, DockBottomOf, objPane): objPane.Title = "����Ԥ��": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(3, 100, 100, DockRightOf, Nothing): objPane.Title = "�½�����": objPane.Options = PaneNoCaption

    Set mObjTabEprView = New cTableEPR
    If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    Call RestoreWinState(Me, App.ProductName)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strCols As String, i As Long
    On Error Resume Next
    For i = 0 To vfgWrit.Cols - 1
        strCols = strCols & IIf(i = 0, "", ";") & vfgWrit.ColWidth(i)
    Next

    mstrColWidthConfig = strCols
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", mstrColWidthConfig
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgWrit.Name, "FontSize", vfgWrit.FontSize
    If Not mfrmContent Is Nothing Then Unload mfrmContent
    If Not mfrmNew Is Nothing Then Unload mfrmNew
    If Not mfrmMonitor Is Nothing Then Unload mfrmMonitor
    If Not mfrmTipInfo Is Nothing Then Unload mfrmTipInfo
    Set mfrmContent = Nothing
    Set mfrmNew = Nothing
    Set mfrmMonitor = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mfrmTipInfo = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmNew_NewClick(ByVal FileId As Long, ByVal babyNum As Long)
Dim frmThis As Form, bFinded As Boolean, strTmp As String
Dim rs As New ADODB.Recordset, strSQL As String
    
    If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        If Not gobjPlugIn.AddEMRBefore(glngSys, 1255, mlngPatiId, mlngPageId, FileId) Then Exit Sub
        Err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errHand
    If gstrPrivsEpr = ";;" Then
        MsgBox "�����߱������༭��ӦȨ�ޣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Split(EprIsCommit, "|")(0) = 0 Then
        MsgBox "�ò��˲������ύ��飬����������������ȡ���������ԣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TimeLimitOut Then Exit Sub
    
    strSQL = "Select ���� From �����ļ��б� Where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, FileId)
    If rs!���� < 0 Then
        '���ⲡ������������
        Exit Sub
    ElseIf rs!���� = 2 Then '���ʽ�༭��
        If Not mObjTabEpr Is Nothing Then
            bFinded = mObjTabEpr.Showfrm(FileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
        End If
        If Not bFinded Then
            Set mObjTabEpr = New cTableEPR
            mObjTabEpr.InitOpenEPR Me, cprEM_����, cprET_�������༭, FileId, True, 0, cprPF_סԺ, mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "��������ӡ") > 0, Val(gstrESign)
            dkpMain.Panes(3).Close
        End If
    Else
        'RichEPR����
        '�жϹ����ĵ��Ƿ��Ѿ���д��
        gstrSQL = "Select ID From �����ļ��б� Where ��� <> NVL(ҳ��,���) And ID =[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
        If rs.EOF = False Then '�ǹ����ĵ�
            gstrSQL = "Select M.ID,M.����" & vbNewLine & _
                        "       From �����ļ��б� L, �����ļ��б� M" & vbNewLine & _
                        "       Where M.���� = L.���� And M.��� = L.ҳ�� And L.ID =[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
            If rs.EOF Then MsgBox "�ò����Ĺ���������ʧЧ������ϵϵͳ����Ա��", vbInformation, gstrSysName: Exit Sub
            strTmp = rs!ID & "|" & rs!����
            gstrSQL = "Select ID" & vbNewLine & _
                        "From ���Ӳ�����¼" & vbNewLine & _
                        "Where ����id = [1] And ��ҳid =[2] And �ļ�id+0 =[3]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, Val(Split(strTmp, "|")(0)))
            If rs.EOF Then
                MsgBox "�ò����Ĺ����� [" & Split(strTmp, "|")(1) & "] ��δ��д�����顣", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    
        For Each frmThis In Forms
            If frmThis.Name = "frmMain" Then
                With frmThis.Document
                    If .EPRFileInfo.ID = FileId And .EPRPatiRecInfo.����ID = mlngPatiId _
                        And .EPRPatiRecInfo.������Դ = cprPF_סԺ And .EPRPatiRecInfo.��ҳID = mlngPageId _
                        And .EPRPatiRecInfo.����ID = mlngDeptId And frmThis.ChildMode = False Then
                        frmThis.Show
                        bFinded = True
                    End If
                End With
            End If
        Next
        If bFinded = False Then
            Dim Doc As New cEPRDocument
            
            Doc.InitEPRDoc cprEM_����, cprET_�������༭, FileId, cprPF_סԺ, mlngPatiId, CStr(mlngPageId), , mlngDeptId
            Doc.EPRPatiRecInfo.Ӥ�� = babyNum
            Doc.ShowEPREditor Me
            
            dkpMain.Panes(3).Close
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'��ʾָ�������б��е���ʷǩ����¼
Dim strTipInfo As String, lngRow As Long, strPrint As String
    If picInfo.Visible = False Then Exit Sub
    lngRow = vfgWrit.MouseRow
    If lngRow <= 0 Then Exit Sub
    
    strTipInfo = vfgWrit.Cell(flexcpData, lngRow, mCol.w��ǰ���)
    If strTipInfo = "" Then '���û�л�ȡ������������ȡ����¼���б���
        strTipInfo = GetEprSign(vfgWrit.TextMatrix(lngRow, mCol.wID))   '��ȡǩ��
        Call EprPrinted(vfgWrit.TextMatrix(lngRow, mCol.wID), strPrint) '��ȡ��ӡ��¼
        strTipInfo = "�� " & Rpad(vfgWrit.TextMatrix(lngRow, mCol.w������), 8) & _
                     "�� " & Rpad(vfgWrit.TextMatrix(lngRow, mCol.w����ʱ��), 19) & " ����" & vbCrLf & strTipInfo
        strTipInfo = strTipInfo & vbCrLf & strPrint
        vfgWrit.Cell(flexcpData, lngRow, mCol.w��ǰ���) = strTipInfo
    End If
    
    mfrmTipInfo.ShowTipInfo picInfo.hwnd, strTipInfo, True
End Sub
Private Function GetEprSign(ByVal lngFileID As Long)
'��ȡ������ʷǩ����¼
Dim rsTemp As ADODB.Recordset, strSign As String
    gstrSQL = "Select ��ʼ�� As �汾, Decode(Ҫ�ر�ʾ, 3, '����ҽʦ', 2, '����ҽʦ', '����ҽʦ') || '���' || Decode(��ʼ��, 1, 'ǩ��', '�޶�') As ����," & vbNewLine & _
                "       Decode(Nvl(Instr(�����ı�, ';'), 0), 0, �����ı�, Substr(�����ı�, 1, Instr(�����ı�, ';') - 1)) As ��Ա," & vbNewLine & _
                "       RTrim(Substr(��������, Instr(��������, ';', 1, 4) + 1)) As ʱ��" & vbNewLine & _
                "From ���Ӳ�������" & vbNewLine & _
                "Where �ļ�id = [1] And �������� = 8 Order By ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ����¼", lngFileID)
    Do Until rsTemp.EOF
        strSign = strSign & "�� " & Rpad(NVL(rsTemp!��Ա), 8) & "�� " & Rpad(NVL(rsTemp!ʱ��), 19) & " ��" & NVL(rsTemp!����) & vbCrLf
        rsTemp.MoveNext
    Loop
    GetEprSign = strSign
End Function
Private Function EprPrinted(ByVal lngRecordId As Long, Optional strPrintInfo As String) As Boolean
'��鵱ǰ������¼�Ƿ��Ѿ���ӡ��
Dim rsTemp As ADODB.Recordset
On Error GoTo errHand
    '��Ҫ�������Ӳ�����¼����ӡ�ˣ���ӡʱ�䣩��������ʷ���ݲ�ת�ƣ���¼�������ϲ�ѯ
    gstrSQL = "Select ��ӡ��, ��ӡʱ�� From ���Ӳ�����ӡ Where �ļ�id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select ��ӡ��, ��ӡʱ�� From ���Ӳ�����¼ Where ID = [1] And ��ӡ�� is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp.EOF Then Exit Function
    
    Do Until rsTemp.EOF
        strPrintInfo = strPrintInfo & vbCrLf & "��ӡ�ˣ�" & Rpad(rsTemp!��ӡ��, 8) & "��ӡʱ�䣺" & Format(rsTemp!��ӡʱ��, "yyyy-MM-dd hh:mm")
        rsTemp.MoveNext
    Loop
    strPrintInfo = Mid(strPrintInfo, 3)
    EprPrinted = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function EprIsCommit() As String
'��|�ָ���ʽ����,״̬Ϊ0 ������ 1 �����ֱ���� ����|ɾ��|����

Dim rsTemp As ADODB.Recordset, intNew As Integer, intDel As Integer, intMod As Integer
    gstrSQL = "Select ����״̬ From ������ҳ Where ����id = [1] And ��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)

    Select Case NVL(rsTemp!����״̬, 0)
        Case 0
            intNew = 1: intDel = 1: intMod = 1
        Case 1 '�ȴ����
            intNew = 0: intDel = 0: intMod = 0
        Case 2 '�ܾ����
            intNew = 1: intDel = 1: intMod = 1
        Case 3 '�������
            intNew = 0: intDel = 0: intMod = 0
        Case 4 '��鷴��
            intNew = 0: intDel = 0: intMod = 1
        Case 5 '���鵵
            intNew = 0: intDel = 0: intMod = 0
        Case 6 '�������
            intNew = 0: intDel = 0: intMod = 1
        Case 13 '���ڳ��
            intNew = 1: intDel = 1: intMod = 1
        Case 14 '��鷴��
            intNew = 1: intDel = 1: intMod = 1
        Case 16 '�������
            intNew = 1: intDel = 1: intMod = 1
        Case Else
            intNew = 0: intDel = 0: intMod = 0
    End Select
    EprIsCommit = CStr(intNew) & "|" & CStr(intDel) & "|" & CStr(intMod)
End Function
Private Sub PrintCancel(ByVal lngRecordId As Long)
'ȡ����Ǵ�ӡ
On Error GoTo errHand
    gstrSQL = "Zl_���Ӳ�����ӡ_Cancel(" & lngRecordId & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vfgWrit.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        fraColSel.Move Me.vfgWrit.Left + 50, Me.vfgWrit.Top + 50
        fraColSel.ZOrder 0
        vsColumn.Move fraColSel.Left, fraColSel.Top + fraColSel.Height
        vsColumn.ZOrder 0
        
    End Select
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim i As Long

    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsColumn
            If .Visible Then
                .Visible = False
                vfgWrit.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vfgWrit.ColHidden(.RowData(i)) Or vfgWrit.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next

                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub vfgWrit_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
If picInfo.Visible Then
    picInfo.Move vfgWrit.Cell(flexcpLeft, NewTopRow, mCol.w��ǰ���) + vfgWrit.Cell(flexcpWidth, NewTopRow, mCol.w��ǰ���) - picInfo.Width - 30
End If
End Sub

Private Sub vfgWrit_DblClick()
Dim lFileId As Long, bFinded As Boolean
    '��vfgWrit�б任ʱ������˫���¼�����˫���¼�ִ��ʱ������˫���¼�����
    If mblnViewTag = True Or mblnViewNow = True Then Exit Sub
    '˫���¼�ִ�б�־��true����ִ�У�falseû��ִ��
    mblnViewNow = True
    If Not mblnEdit Then Exit Sub
    lFileId = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID))
    If lFileId = 0 Then Exit Sub
    If vfgWrit.TextMatrix(vfgWrit.Row, mCol.w�༭��ʽ) = 0 Then
        Dim fViewDoc As New frmEPRView, blnCanPrint As Boolean
        blnCanPrint = (InStr(1, mstrPrivs, "��������ӡ") > 0) And (Trim(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w�鵵��)) = "" Or InStr(1, mstrPrivs, "�鵵�������") > 0)
        fViewDoc.ShowMe Me, CLng(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)), , blnCanPrint
    Else
        If Not mObjTabEprView Is Nothing Then
            bFinded = mObjTabEprView.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
        End If
        If Not bFinded And mblnEdit Then
            mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, True, 0, cprPF_סԺ, mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "��������ӡ") > 0, Val(gstrESign)
        End If
    End If
    mblnViewNow = False
    Call vfgWrit_RowColChange '��ֹѡ��������û��ˢ�£��ֶ�ˢ��
End Sub

Private Sub vfgWrit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngCol As Long, lngRow As Long
    lngCol = vfgWrit.MouseCol: lngRow = vfgWrit.MouseRow
    If lngRow <= 0 Then picInfo.Visible = False: Exit Sub
    
    If Not Me.ActiveControl Is Nothing Then
        If Me.ActiveControl.Name <> "vfgWrit" Then
            vfgWrit.SetFocus
        Else
            vfgWrit.SetFocus
        End If
    Else
        vfgWrit.SetFocus
    End If
    
    If Val(vfgWrit.TextMatrix(lngRow, mCol.wID)) <> 0 Then
        If Val(picInfo.Tag) = lngRow And picInfo.Visible Then Exit Sub
        picInfo.Tag = lngRow
        picInfo.Move vfgWrit.Cell(flexcpLeft, lngRow, mCol.w��ǰ���) + vfgWrit.Cell(flexcpWidth, lngRow, mCol.w��ǰ���) - picInfo.Width - 30, vfgWrit.Cell(flexcpTop, lngRow, mCol.w��ǰ���) + 15
        If vfgWrit.RowSel = lngRow Then
            picInfo.BackColor = vfgWrit.BackColorSel
        Else
            picInfo.BackColor = &H80000005
        End If
        picInfo.Visible = True
    Else
        picInfo.Visible = False
    End If
End Sub

Private Sub vfgWrit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If vfgWrit.MouseRow = -1 Then vfgWrit.Row = vfgWrit.Rows - 1
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
    'mblnInsideTools = True
    If Button = vbRightButton And mblnInsideTools Then
        Dim Popup As CommandBar
        Dim cbrControl As CommandBarControl
        
        Set Popup = cbrMain.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����(&U)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "�鵵(&I)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ(&P)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "��������(&S)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub vfgWrit_RowColChange()
    If vfgWrit.Cols < mCol.wID + 1 Then Exit Sub 'δ��ʼ��
    
    '��˫���¼�ִ��ʱ����ִ�б任��ȡ���ݲ�������ǰ�����¼�в����ʱ��ִ��ˢ��
    If mblnViewNow = True Then Exit Sub
    mblnViewTag = True
    
    If Not mfrmNew Is Nothing Then dkpMain.Panes(3).Close

    Err = 0
    On Error Resume Next
    mlngCurId = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) 'ѡ���е�ID
    If Val(vfgWrit.Tag) = mlngCurId Then Exit Sub 'δ�л���

    If Not mfrmContent Is Nothing Then Call mfrmContent.zlRefresh(mlngCurId, IIf(mblnEdit = False, "", mstrPrivs), , mblnMoved, , Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w�༭��ʽ)), True)
    vfgWrit.Tag = mlngCurId 'ˢ������¼��ǰ�е�ID
    mblnViewTag = False
    
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim lngCol As Long, T As Variant, i As Long

    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            T = Split("270;0;0;1200;2000;800;1600;800;1600;800;0;3300;0;0;0;1200;0;0", ";")
            vfgWrit.ColWidth(lngCol) = T(lngCol)
            vfgWrit.ColHidden(lngCol) = False
        Else
            vfgWrit.ColWidth(lngCol) = 0
            vfgWrit.ColHidden(lngCol) = True
        End If
    End If
    Dim strCols As String
    For i = 0 To vfgWrit.Cols - 1
        strCols = strCols & IIf(i = 0, "", ";") & vfgWrit.ColWidth(i)
    Next
    mstrColWidthConfig = strCols
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    On Error Resume Next
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub

Private Sub vsColumn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then '�ر���ѡ����
        If vsColumn.Visible Then
            vsColumn.Visible = False
            vfgWrit.SetFocus
        End If
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then '����ѡ����
        Call imgColSel_MouseUp(1, 0, 0, 0)
    End If
End Sub

Private Sub vfgWrit_KeyDown(KeyCode As Integer, Shift As Integer)
    vsColumn_KeyDown KeyCode, Shift
End Sub
Private Function TimeLimitOut() As Boolean
'����:����Ƿ���ת�ƣ���Ժ��Ԥ��Ժ�������������¼��Ͳ�¼ʱ��
Dim rsTemp As New ADODB.Recordset, lngTimeLimit As Long, strReturn As String
    
    gstrSQL = "Select Decode(��ֹԭ��, 1, '��Ժ', 3, 'ת��', 10, 'Ԥ��Ժ') �¼�, ��ֹʱ��,Trunc((Sysdate - ��ֹʱ��) * 24, 5) ��ǰʱ��" & vbNewLine & _
                "From ���˱䶯��¼" & vbNewLine & _
                "Where ID = (Select Nvl(Max(ID), 0)" & vbNewLine & _
                "            From ���˱䶯��¼" & vbNewLine & _
                "            Where ����id = [1] And ��ҳid = [2] And ��ֹʱ�� Is Not Null And ��ֹԭ�� In (1, 3, 10))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����䶯��¼", mlngPatiId, mlngPageId)
    If rsTemp.EOF Then Exit Function
    
    lngTimeLimit = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", 100))
    
    If rsTemp!��ǰʱ�� > lngTimeLimit Then
        If rsTemp!�¼� = "ת��" Then
            strReturn = rsTemp!�¼� & "|" & lngTimeLimit
            gstrSQL = "Select ��Ժ����id From ������ҳ Where ����id = [1] And ��ҳid = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", mlngPatiId, mlngPageId)
            If mlngDeptId = rsTemp!��Ժ����ID Then strReturn = "" 'ת�ƺ���ת�����������������ʱ������
        Else
            strReturn = rsTemp!�¼� & "|" & lngTimeLimit
        End If
    End If
    
    If strReturn <> "" Then
        MsgBox "�ò����Ѿ�" & Split(strReturn, "|")(0) & ",���ҳ����趨��" & Split(strReturn, "|")(1) & "Сʱ��¼ʱ��,������䶯������", vbInformation, gstrSysName
        TimeLimitOut = True
    End If
End Function
