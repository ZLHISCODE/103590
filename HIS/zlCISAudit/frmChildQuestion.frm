VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmChildQuestion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cbo������ 
      Height          =   300
      Left            =   5145
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   435
      Width           =   2625
   End
   Begin VB.Timer tmr 
      Interval        =   60000
      Left            =   2490
      Top             =   5430
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Left            =   3990
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   450
      Width           =   1140
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   3
      Left            =   810
      ScaleHeight     =   1935
      ScaleWidth      =   2340
      TabIndex        =   6
      Top             =   2880
      Width           =   2340
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   150
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   2
      Left            =   255
      ScaleHeight     =   1935
      ScaleWidth      =   2340
      TabIndex        =   4
      Top             =   1635
      Width           =   2340
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   150
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   1
      Left            =   4005
      ScaleHeight     =   3585
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   1155
      Width           =   3135
      Begin XtremeSuiteControls.TabControl tbcQuestion 
         Height          =   1830
         Left            =   255
         TabIndex        =   3
         Top             =   450
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   0
      Left            =   150
      ScaleHeight     =   1935
      ScaleWidth      =   2340
      TabIndex        =   0
      Top             =   510
      Width           =   2340
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   150
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   -30
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmChildQuestion.frx":0000
      Left            =   810
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChildQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mfrmMain            As Object
Private mlngKey             As Long
Private mstr�ļ�id          As String
Private mlngҽ��id          As Long
Private mlng����ID          As Long
Private mlngReferKey        As Long
Private mblnReading         As Boolean
Private mstrSQL             As String
Private mblnDataChanged     As Boolean
Private mblnAllowModify     As Boolean
Private mlngMoudal          As Long
Private mclsVsf(2)          As New clsVsf
Private mlng�ύId          As Long
Private mlng����ID          As Long
Private mlng��ҳID          As Long
Private mstrObject          As String
Private mlngTmp             As Long
Private mintIndex           As Integer
Private mintPreTime         As String
Private mstrStart           As String
Private mstrEnd             As String
Private mblnCurrentPatient  As Boolean
Private mlngIntenal         As Long
Private mlngLoop            As Long
Private mstrDepts           As String
Private mrsCondition        As ADODB.Recordset
Private mlngCurNum          As Long '��ǰ����
Private mstr����ѡ��        As String
Private mstr��鿪ʼʱ��    As String
Private mstr������ʱ��    As String
Private mstr������          As String
Private mblnRef             As Boolean
Private mblnAuditEnter  As Boolean              '�Ƿ�����¼��������
Private mstrPrivs       As String
Private mblnDataExecute As Boolean

Private WithEvents mfrmChildQuestionEdit As frmChildQuestionEdit
Attribute mfrmChildQuestionEdit.VB_VarHelpID = -1

Public Event AfterSaveQuestion(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
Public Event AfterDeleteQuestion(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
Public Event AfterDataChanged()
Public Event LocationDocument(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal byt�������� As Byte, ByVal lng�ļ�id As Long, ByVal lngҽ��id As Long, ByVal lng����ID As Long)
Public Event AfterQuestionType(ByVal blnQuestionType As Boolean)

'######################################################################################################################
Public Property Get �ύId() As Long
    �ύId = mlng�ύId
End Property

Public Property Get Depts() As String
    Depts = mstrDepts
End Property

Public Property Let Depts(ByVal vDepts As String)
    mstrDepts = vDepts
End Property

Public Property Get CurrentPatient() As Boolean
    CurrentPatient = mblnCurrentPatient
End Property

Public Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildQuestionEdit.DataChanged = blnData
End Property

Public Property Get DataChanged() As Boolean
    If Not (mfrmChildQuestionEdit Is Nothing) Then
        DataChanged = mfrmChildQuestionEdit.DataChanged
    End If
End Property

Public Property Let AllowModify(ByVal blnData As Boolean)
    mblnAllowModify = blnData
    
    mfrmChildQuestionEdit.AllowModify = blnData
    
End Property

Public Property Get AllowModify() As Boolean
    AllowModify = mblnAllowModify
End Property

Public Function InitData(ByVal frmMain As Object, ByVal lngMoudal As Long, ByVal blnAllowModify As Boolean, ByVal blnAuditEnter As Boolean, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAuditEnter = blnAuditEnter
    Set mfrmMain = frmMain
    mblnAllowModify = blnAllowModify
    mlngMoudal = lngMoudal
    mstrPrivs = strPrivs
    
    mstr����ѡ�� = "ǰһ��"
    mstr��鿪ʼʱ�� = GetDateTime("ǰһ��", 1)
    mstr������ʱ�� = GetDateTime("ǰһ��", 2)
    
    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Then Exit Function
    Call ExecuteCommand("��ע���")
    Call ExecuteCommand("�ؼ�״̬")
    Call ExecuteCommand("ˢ�´���")
    mintPreTime = cbo.Text
    If cbo.Text <> "[ָ��...]" Then
        mstrStart = GetDateTime(mintPreTime, 1)
        mstrEnd = GetDateTime(mintPreTime, 2)
    End If
    Call ExecuteCommand("��ȡ��ɷ���")
    
    DataChanged = False
    
End Function

Public Function SetParamter(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strObject As String, ByVal strParam As String, Optional ByVal lng�ύId As Long = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim varParam As Variant
    
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlng�ύId = lng�ύId
    mstrObject = strObject
    
    Select Case mstrObject
    Case "��ҳ��¼", "סԺҽ��"
                
        mstr�ļ�id = 0
        mlngҽ��id = 0
        mlng����ID = 0
            
    Case "סԺ����", "������", "֪���ļ�", "����֤��"
                
        If strParam <> "" Then
            varParam = Split(strParam, ";")
            mstr�ļ�id = varParam(0)
            mlngҽ��id = 0
            mlng����ID = 0
        End If
    
    Case "ҽ������"
        'strParam������id;ҽ��id
        If strParam <> "" Then
            varParam = Split(strParam, ";")
            If UBound(varParam) >= 1 Then
                mstr�ļ�id = varParam(0)
                mlngҽ��id = Val(varParam(1))
                mlng����ID = 0
            End If
        End If
    
    Case "�����¼"
        
        'strParam������id;����;��ʼ~��ֹ;�ļ�id
        
        If strParam <> "" Then
            varParam = Split(strParam, ";")
            If UBound(varParam) >= 1 Then
                mstr�ļ�id = Val(varParam(3))
                mlngҽ��id = 0
                mlng����ID = Val(varParam(0))
            End If
            
        End If
    Case Else
        mstr�ļ�id = 0
        mlngҽ��id = 0
        mlng����ID = 0
    End Select
    
    SetParamter = True
    
End Function

Public Function RefreshData(strDepts As String, rsCondition As ADODB.Recordset, ByVal blnAuditEnter As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAuditEnter = blnAuditEnter
    mstrDepts = strDepts
    Set mrsCondition = rsCondition
    Call ExecuteCommand("��ʼ����")
    Call ExecuteCommand("�ؼ�״̬")
    
    If ExecuteCommand("ˢ������") = False Then Exit Function
    
    DataChanged = False
    
    RefreshData = True
    
End Function

'######################################################################################################################
Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_ThingAdd, "���Ӵ���", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "ѡ�����", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "���ӷ���", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_CopyNewItem, "�ٴη���", , conMenu_Edit_NewItem, xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ������", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Send, "��������", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SendBack, "���˽���", , , xtpButtonIcon)
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "�������", True, , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ������", , , xtpButtonIcon)
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "��ǰ����", True, , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "ˢ�·���", , , xtpButtonIcon)
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, conMenu_View_Find, "ʱ��", , 1, xtpButtonCaption)
    objControl.Flags = xtpFlagRightAlign
    Set objCustom = NewToolBar(objBar, xtpControlCustom, conMenu_View_Find, "", , , xtpButtonCaption)
    objCustom.Handle = cbo.hWnd
    objCustom.Flags = xtpFlagRightAlign
    
    
    Set objBar = cbsMain.Add("���", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, conMenu_View_FindType, "������", , 1, xtpButtonCaption)
    objControl.Flags = xtpFlagAlignLeft
    
    Set objCustom = NewToolBar(objBar, xtpControlCustom, conMenu_View_FindType, "", , , xtpButtonCaption)
    objCustom.Handle = cbo������.hWnd
    objCustom.Flags = xtpFlagAlignLeft
    
    
    
    
End Function

Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs              As New ADODB.Recordset
    Dim rsSQL           As New ADODB.Recordset
    Dim blnAllowModify  As Boolean
    Dim intRow          As Integer
    Dim strTmp          As String
    Dim strDept         As String
    Dim i               As Integer
    Dim mlngTmpCurNum   As Long
    On Error GoTo errHand
    
    mblnReading = True
    Call SQLRecord(rsSQL)
    
    Select Case strCmd
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsf(0) = New clsVsf
        With mclsVsf(0)
            Call .Initialize(Me.Controls, vsf(0), True, False, frmPubResource.GetImageList(16))
            Call .ClearColumn
            
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[ͼ��]", False)
            Call .AppendColumn("�������", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 750, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("�ύid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("���id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("�ļ�id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("ҽ��id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��������id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("����", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            
            .AppendRows = True
        End With
        
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsf(1) = New clsVsf
        With mclsVsf(1)
            Call .Initialize(Me.Controls, vsf(1), True, False, frmPubResource.GetImageList(16))
            Call .ClearColumn
            
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("�ύid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("���id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("�ļ�id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("ҽ��id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[ͼ��]", False)
            Call .AppendColumn("����", 750, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("�������", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("����", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            
            .AppendRows = True
        End With
        
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsf(2) = New clsVsf
        With mclsVsf(2)
            Call .Initialize(Me.Controls, vsf(2), True, False, frmPubResource.GetImageList(16))
            Call .ClearColumn
            
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("�ύid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("���id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("�ļ�id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("ҽ��id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[ͼ��]", False)
            Call .AppendColumn("����", 750, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("�������", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("����", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            
            .AppendRows = True
        End With
        
        Call InitCommandBar
            
        '����ͣ������
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing): objPane.Title = "����": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, Nothing): objPane.Title = "��ϸ": objPane.Options = PaneNoCaption

        dkpMain.SetCommandBars cbsMain
        Call DockPannelInit(dkpMain)
        
        Call TabControlInit(tbcQuestion)
        With tbcQuestion
            .PaintManager.BoldSelected = True
                           
            .InsertItem 0, "δ��", picPane(0).hWnd, 6
            .InsertItem 1, "δ��", picPane(2).hWnd, 7
            .InsertItem 2, "����", picPane(3).hWnd, 8
            .Item(0).Selected = True
        End With
        
        With cbo
            .AddItem "��  ��"
            .AddItem "��  ��"
            .AddItem "��  ��"
            .AddItem "��  ��"
            .AddItem "��  ��"
            .AddItem "������"
            .AddItem "��  ��"
            .AddItem "ǰ����"
            .AddItem "ǰһ��"
            .AddItem "ǰ����"
            .AddItem "ǰһ��"
            .AddItem "ǰ����"
            .AddItem "ǰ����"
            .AddItem "ǰ����"
            .AddItem "ǰһ��"
            .AddItem "ǰ����"
            .AddItem "[ָ��...]"
            .ListIndex = 0
        End With
        
        mintPreTime = "��  ��"
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        mlngIntenal = Val(GetPara("δ����ˢ��Ƶ��", mfrmMain.ģ���, "5"))
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        If tbcQuestion.Enabled <> Not DataChanged Then
            
            tbcQuestion.Enabled = Not DataChanged
            vsf(0).Enabled = Not DataChanged
            vsf(1).Enabled = Not DataChanged
            vsf(2).Enabled = Not DataChanged
            
            vsf(0).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
            vsf(1).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
            vsf(2).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
            
            RaiseEvent AfterDataChanged

        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��״̬"
    
        With vsf(0)
            If tbcQuestion.ItemCount > 0 Then
                If Val(.RowData(.Row)) > 0 Then
                    tbcQuestion.Item(0).Caption = "δ��(" & .Rows - 1 & ")"
                Else
                    tbcQuestion.Item(0).Caption = "δ��"
                End If
            End If
        End With
        
        With vsf(1)
            If tbcQuestion.ItemCount > 0 Then
                If Val(.RowData(.Row)) > 1 Then
                    tbcQuestion.Item(1).Caption = "δ��(" & .Rows - 1 & ")"
                Else
                    tbcQuestion.Item(1).Caption = "δ��"
                End If
            End If
        End With

        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
        
        Call ExecuteCommand("��ȡδ�ķ���")
        Call ExecuteCommand("��ȡδ����")
        Call ExecuteCommand("��ȡ��ɷ���")
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("ˢ��״̬")
        
        GoTo endHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��ָ������"
                
        Set rs = gclsPackage.GetQuestion(mrsCondition, "", 1, mlngTmp)
        If rs.BOF = True Then Exit Function

        intRow = mclsVsf(0).FindRow(mlngTmp, -1)
        With vsf(0)
            If intRow > 0 Then
                '�Ѽ���
                .Row = intRow
            Else
                'δ����
                If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
            Call mclsVsf(0).LoadGridRow(.Row, rs)
        End With

        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("ˢ��״̬")
                
    '------------------------------------------------------------------------------------------------------------------
    Case "���ӷ�����¼"
        
        '����Ƿ��Ѿ������˷���
        If gclsPackage.GetExamineStartUse = False Then
            Call MsgBox("�����ڵ��Ӳ��������Ŀ������һ����鷽��,������Ӳ�������!", vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����)
            GoTo endHand
        End If
        
        With vsf(0)
            If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
            .Row = .Rows - 1
            .ShowCell .Row, IIf(.Col = -1, 1, .Col)
        End With

        Call ExecuteCommand("��ȡ��������")
        
        mlngCurNum = GetMaxNumNEW(mlng����ID, mlng��ҳID, 0)
        Call mfrmChildQuestionEdit.SetCurNum(mlngCurNum)
        
        
        Call mfrmChildQuestionEdit.NewData(mstrObject, mstr�ļ�id, mlngҽ��id, mlng����ID, mlngReferKey, mlngCurNum)
        mclsVsf(mintIndex).AppendRows = True
        
        GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
    Case "���Ʒ�����¼"
        
        With vsf(mintIndex)
            '��鲢��ȡ��صķ�����¼
            Set rs = gclsPackage.GetRelevanceID(Val(.RowData(.Row)))
            If Not rs.EOF Then
                If NVL(rs!���ID) = "" Then
                    mlngReferKey = -1
                Else
                    mlngReferKey = NVL(rs!���ID, 0)
                End If
            Else
                mlngReferKey = Val(.RowData(.Row))
            End If
            Call vsf_DblClick(mintIndex)
            If mlngReferKey > 0 Or mlngReferKey = -1 Then tbcQuestion.Item(0).Selected = True
        End With

        GoTo endHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ɾ��������¼"
        
        With vsf(0)
            If Val(.RowData(.Row)) = 0 Then GoTo endHand
            
            If MsgBox("���Ƿ����Ҫɾ����ǰ����������", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                gstrSQL = "zl_����������¼_Delete(" & Val(.RowData(.Row)) & ")"
                Call SQLRecordAdd(rsSQL, gstrSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
                Call ExecuteCommand("ˢ�´���")
                Call SetCob(mlngCurNum)
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ɷ�����¼"
        
        With vsf(mintIndex)
            If Val(.RowData(.Row)) = 0 Then GoTo endHand
            
            If MsgBox("���Ƿ����Ҫ������ǰ����������", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                gstrSQL = "zl_����������¼_Finish(" & Val(.RowData(.Row)) & ",'" & gstrUserName & "')"
                Call SQLRecordAdd(rsSQL, gstrSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
            
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "������ɷ���"
        
        With vsf(mintIndex)
            If Val(.RowData(.Row)) = 0 Then GoTo endHand
            
            If MsgBox("���Ƿ����Ҫ���˵�ǰ�ѽ����ķ���������", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                gstrSQL = "zl_����������¼_RollBackFinish(" & Val(.RowData(.Row)) & ")"
                Call SQLRecordAdd(rsSQL, gstrSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�Ƴ�������¼"
        
        With vsf(mintIndex)
            If .Rows > 2 Then
                .RemoveItem .Row
                mclsVsf(mintIndex).AppendRows = True
            Else
                Call mclsVsf(mintIndex).ClearGrid
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡδ�ķ���"
        
        mlngTmpCurNum = mlngCurNum
        mclsVsf(0).ClearGrid
        mlngCurNum = mlngTmpCurNum
        
        If mrsCondition Is Nothing Then GoTo endHand
        
        If mblnCurrentPatient Then
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 2, , , , mlng����ID, mlng��ҳID, mlngCurNum, mstr��鿪ʼʱ��, mstr������ʱ��, mstr������)
        Else
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 2, , , , , , mlngCurNum, mstr��鿪ʼʱ��, mstr������ʱ��, mstr������)
        End If
        
        If rs.BOF = False Then
            Call mclsVsf(0).LoadGrid(rs)
        End If
                
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡδ����"
        
        mclsVsf(1).ClearGrid
        
        If mrsCondition Is Nothing Then GoTo endHand
        
        If mblnCurrentPatient Then
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 3, , , , mlng����ID, mlng��ҳID, mlngCurNum, mstr��鿪ʼʱ��, mstr������ʱ��, mstr������)
        Else
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 3, , , , , , mlngCurNum, mstr��鿪ʼʱ��, mstr������ʱ��, mstr������)
        End If
        If rs.BOF = False Then
            Call mclsVsf(1).LoadGrid(rs)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��ɷ���"
        
        mclsVsf(2).ClearGrid
        '
        If mrsCondition Is Nothing Then GoTo endHand
        
        If mstrStart = "" Then
            mstrStart = GetDateTime("��  ��", 1)
            mstrEnd = GetDateTime("��  ��", 2)
        End If
        
        If mblnCurrentPatient Then
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 4, , mstrStart, mstrEnd, mlng����ID, mlng��ҳID, mlngCurNum, mstr��鿪ʼʱ��, mstr������ʱ��, mstr������)
        Else
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 4, , mstrStart, mstrEnd, , , mlngCurNum, mstr��鿪ʼʱ��, mstr������ʱ��, mstr������)
        End If
        
        If rs.BOF = False Then
            Call mclsVsf(2).LoadGrid(rs)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"
        
        With vsf(mintIndex)
            Call mfrmChildQuestionEdit.RefreshData(Val(.RowData(.Row)), mblnAuditEnter)
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ�´���"
        
        If mstr����ѡ�� = "�Զ���" Or mstr����ѡ�� = "����" Then
        
        Else
            mstr��鿪ʼʱ�� = GetDateTime(mstr����ѡ��, 1)
            mstr������ʱ�� = GetDateTime(mstr����ѡ��, 2)
        End If
        
        Call Init������(mstr��鿪ʼʱ��, mstr������ʱ��)
        Call SetCob(mlngCurNum)
'        call ExecuteCommand("ˢ������")
          
    '------------------------------------------------------------------------------------------------------------------
    Case "�ָ�����"
            
        If mfrmChildQuestionEdit.DataChanged Then
            With vsf(0)
                If Val(.RowData(.Row)) = 0 And .Rows > 2 Then
                    .Rows = .Rows - 1
                    .Row = .Rows - 1
                End If
            End With
            Call ExecuteCommand("��ȡ��������")
            mfrmChildQuestionEdit.DataChanged = False
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "У������"
    
        '1.
        '--------------------------------------------------
        If mfrmChildQuestionEdit.DataChanged Then
            If mfrmChildQuestionEdit.ValidData = False Then GoTo endHand
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
                    
        If mfrmChildQuestionEdit.DataChanged Then
        
            With vsf(0)
                mlngTmp = Val(.RowData(.Row))
                If mlngTmp > 0 Then
                    
                    '�޸ı���,����id,��ҳid,�ύid�õ�ǰ��¼��ֵ
                    If mfrmChildQuestionEdit.SaveData(rsSQL, mlngTmp, Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), Val(.TextMatrix(.Row, .ColIndex("�ύid"))), mlngCurNum) = False Then GoTo endHand
                
                Else
                
                    '��������ʱ,����id,��ҳid,�ύid�õ�ǰ��紫���ֵ
                    If mfrmChildQuestionEdit.SaveData(rsSQL, mlngTmp, mlng����ID, mlng��ҳID, mlng�ύId, mlngCurNum) = False Then GoTo endHand
                    
                End If
            End With
        
            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            Call ExecuteCommand("ˢ�´���")
            Call SetCob(mlngCurNum)
        End If

        GoTo endHand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ע���"
        
        On Error Resume Next
        
        strTmp = GetPara("������ⷶΧ", mfrmMain.ģ���, "��  ��")
        If Left(strTmp, 7) = "[ָ��...]" Then
            cbo.Text = "[ָ��...]"
            mstrStart = Split(strTmp, ";")(1)
            mstrEnd = Split(strTmp, ";")(2)
        Else
            cbo.Text = strTmp
        End If
        
        mblnCurrentPatient = (Val(zlDatabase.GetPara("��ǰ����", glngSys, mlngMoudal, "0")) = 1)
        
        On Error GoTo errHand

        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
            'ʹ�ø��Ի�����
            mclsVsf(0).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_0_20081113", ""))
            mclsVsf(1).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_1_20081113", ""))
            mclsVsf(2).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_2_20081113", ""))
        End If
        
        mlngCurNum = GetRegister(˽��ģ��, Me.Name, "��ǰ����", 1)
        mstr����ѡ�� = GetRegister(˽��ģ��, Me.Name, "����ѡ��", "ǰһ��")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "дע���"
        
        If cbo.Text = "[ָ��...]" Then
            Call SetPara("������ⷶΧ", cbo.Text & ";" & mstrStart & ";" & mstrEnd, mfrmMain.ģ���)
        Else
            Call SetPara("������ⷶΧ", cbo.Text, mfrmMain.ģ���)
        End If
        Call SetPara("��ǰ����", IIf(mblnCurrentPatient, 1, 0), mfrmMain.ģ���)
        
        Call SetRegister(˽��ģ��, Me.Name, "������_0_20081113", mclsVsf(0).SaveStateToString)
        Call SetRegister(˽��ģ��, Me.Name, "������_1_20081113", mclsVsf(1).SaveStateToString)
        Call SetRegister(˽��ģ��, Me.Name, "������_2_20081113", mclsVsf(2).SaveStateToString)
        Call SetRegister(˽��ģ��, Me.Name, "��ǰ����", mlngCurNum)
        Call SetRegister(˽��ģ��, Me.Name, "����ѡ��", mstr����ѡ��)
    End Select

    ExecuteCommand = True
    
    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
    mblnReading = False
End Function


Private Sub cbo_Click()
    Dim strTmp As String
    Dim dStart As Date
    Dim dEnd As Date
    
    If mblnReading Then Exit Sub
    
    If cbo.Text = "[ָ��...]" Then
        
        If mstrStart = "" Then
            mstrStart = GetDateTime("��  ��", 1)
            mstrEnd = GetDateTime("��  ��", 2)
        End If
        
        dStart = CDate(mstrStart)
        dEnd = CDate(mstrEnd)
        If Not frmQuestionTime.ShowMe(Me, dStart, dEnd) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
            Call zlControl.CboLocate(cbo, mintPreTime)
            vsf(mintIndex).SetFocus
            Exit Sub
        Else
            mstrStart = Format(dStart, "yyyy-MM-dd HH:mm:ss")
            mstrEnd = Format(dEnd, "yyyy-MM-dd HH:mm:ss")
            vsf(mintIndex).SetFocus
            mintPreTime = cbo.Text
        End If
        
    Else
        mintPreTime = cbo.Text
        mstrStart = GetDateTime(mintPreTime, 1)
        mstrEnd = GetDateTime(mintPreTime, 2)
    End If

    Call ExecuteCommand("��ȡ��ɷ���")
    
End Sub

Private Sub cbo������_Click()
    If mblnRef Then Exit Sub
    mlngCurNum = CLng(cbo������.ItemData(cbo������.ListIndex))
    
    If mlngCurNum > 0 Then
        '���·���ʱ��
         If cbo������.Text <> "" Then
            mstr��鿪ʼʱ�� = GetAnalyseTime(cbo������.Text, 1)
            mstr������ʱ�� = GetAnalyseTime(cbo������.Text, 2)
         End If
    Else
        If mstr����ѡ�� = "�Զ���" Or mstr����ѡ�� = "����" Then
            mstr��鿪ʼʱ�� = Format("2000-01-01 00:00:00", "yyyy-MM-dd HH:mm:SS")
            mstr������ʱ�� = Format("3000-01-01 23:59:59", "yyyy-MM-dd HH:mm:SS")

        Else
            mstr��鿪ʼʱ�� = GetDateTime(mstr����ѡ��, 1)
            mstr������ʱ�� = GetDateTime(mstr����ѡ��, 2)
        End If
    End If
    
    
    Call ExecuteCommand("��ȡδ�ķ���")
    Call ExecuteCommand("��ȡ��������")
    Call ExecuteCommand("ˢ��״̬")
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_ThingAdd        '���Ӵ���
    '��ȡ���� ������+1
    'ˢ���б�
        mlngCurNum = GetMaxNum(mstr��鿪ʼʱ��, mstr������ʱ��)
        Call ExecuteCommand("ˢ������")
        Call mfrmChildQuestionEdit.SetCurNum(mlngCurNum)
        
        mlngReferKey = 0
        Call ExecuteCommand("���ӷ�����¼")

        DataChanged = False
    Case conMenu_File_Preview              '��������
    '--------------------------------------------------------------------------------------------------------------
    '��ʾ��������
    '��������ˢ���б�
        Dim blnFilter As Boolean
        blnFilter = frmChildQuestionFilter.ShowPara(Me, mstr��鿪ʼʱ��, mstr������ʱ��, mstr����ѡ��, mlngCurNum, mstr������)
        If blnFilter Then
            Call Init������(mstr��鿪ʼʱ��, mstr������ʱ��)
            Call SetCob(mlngCurNum)
            If ExecuteCommand("ˢ������") = False Then Exit Sub
            DataChanged = False
        End If
    Case conMenu_Edit_NewItem               '���ӷ�����¼
    '--------------------------------------------------------------------------------------------------------------
        mlngReferKey = 0
        Call ExecuteCommand("���ӷ�����¼")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_CopyNewItem               '���Ʒ�����¼
        
        mlngReferKey = 0
        Call ExecuteCommand("���Ʒ�����¼")
        If mlngReferKey > 0 Or mlngReferKey = -1 Then Call ExecuteCommand("���ӷ�����¼")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                'ɾ��������¼

        If ExecuteCommand("ɾ��������¼") Then
            Call ExecuteCommand("�Ƴ�������¼")
            Call ExecuteCommand("ˢ��״̬")
'            RaiseEvent AfterDeleteQuestion(mlng����ID, mlng��ҳID)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Send            '��ɷ�����¼
        
        If ExecuteCommand("��ɷ�����¼") Then
            Call ExecuteCommand("�Ƴ�������¼")
            Call ExecuteCommand("��ȡ��ɷ���")
            Call ExecuteCommand("ˢ��״̬")
            RaiseEvent AfterDeleteQuestion(mlng����ID, mlng��ҳID)
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SendBack
                
        If ExecuteCommand("������ɷ���") Then
            Call ExecuteCommand("�Ƴ�������¼")
            Call ExecuteCommand("��ȡδ�ķ���")
            Call ExecuteCommand("��ȡδ����")
            Call ExecuteCommand("ˢ��״̬")
            
            RaiseEvent AfterDeleteQuestion(mlng����ID, mlng��ҳID)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save           '��������
    
        If ExecuteCommand("У������") And DataChanged Then
            If ExecuteCommand("��������") Then
                
                DataChanged = False
                
                Call ExecuteCommand("ˢ��ָ������")
                Call ExecuteCommand("ˢ��״̬")
                
'                RaiseEvent AfterSaveQuestion(mlng����ID, mlng��ҳID)
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle         '�ָ�����
    
        Call ExecuteCommand("�ָ�����")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Find                '��������
        
        If ExecuteCommand("��������") Then
            Call ExecuteCommand("ˢ������")
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter                  '��ǰ����
        
        mblnCurrentPatient = Not mblnCurrentPatient
        mlngCurNum = CLng(cbo������.ItemData(cbo������.ListIndex))
        Call ExecuteCommand("ˢ������")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               'ˢ������
        Call ExecuteCommand("ˢ������")
    End Select
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand
    
    With vsf(mintIndex)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Find
            Control.Visible = (mintIndex = 2)
            Control.Enabled = (Control.Visible And DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_ThingAdd
            Control.Visible = False ' (mintIndex = 0 And AllowModify)
'            Control.Enabled = (Control.Visible And DataChanged = False And AllowModify And mlng����ID > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview
            Control.Visible = (mintIndex = 0 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And AllowModify And mlng����ID > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem
            Control.Visible = (mintIndex = 0 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And AllowModify And mlng����ID > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_CopyNewItem
            Control.Visible = (mintIndex = 1 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And AllowModify And mlng����ID > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            Control.Visible = (mintIndex = 0 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And AllowModify)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Send
            Control.Visible = (mintIndex <> 2 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And AllowModify)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SendBack
            Control.Visible = (mintIndex = 2 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And AllowModify)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
            Control.Visible = ((mintIndex = 0 Or mintIndex = 1) And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = True And AllowModify)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Filter
            Control.Visible = (IsPrivs(mstrPrivs, "Ժ������") And IsPrivs(mstrPrivs, "�Ƽ�����")) Or (IsPrivs(mstrPrivs, "Ժ������") And IsPrivs(mstrPrivs, "�Ƽ�����") = False)
        
            Control.Checked = Control.Visible And mblnCurrentPatient
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_FindType '������
            Control.Enabled = (mintIndex = 0 And AllowModify)
'            If (mintIndex = 0 And AllowModify) Then
''                Me.cbsMain.ActiveMenuBar.Visible = False
'            Else
'
'            End If
'            Control.Checked = mblnCurrentPatient
        End Select
    End With
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(1).hWnd
    Case 2
        Set mfrmChildQuestionEdit = New frmChildQuestionEdit
        Call mfrmChildQuestionEdit.InitData(mfrmMain, AllowModify, mstrPrivs)
        Item.Handle = mfrmChildQuestionEdit.hWnd
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 2, 100, 325, Me.ScaleWidth, 325)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ExecuteCommand("дע���")
    Unload mfrmChildQuestionEdit
End Sub

Private Sub mfrmChildQuestionEdit_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildQuestionEdit_AfterQuestionType(ByVal blnQuestionType As Boolean)
    'blnQuestionType=True Ժ������ =Flase �Ƽ�����
    Dim lngCurNum As Long
    If blnQuestionType Then
        lngCurNum = GetMaxNumNEW(mlng����ID, mlng��ҳID, 0)
        Call mfrmChildQuestionEdit.SetCurNum(lngCurNum)
    Else
        lngCurNum = GetMaxNumNEW(mlng����ID, mlng��ҳID, 1)
        Call mfrmChildQuestionEdit.SetCurNum(lngCurNum)
    End If
    RaiseEvent AfterQuestionType(blnQuestionType)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(0).AppendRows = True
    Case 1
        tbcQuestion.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 2
        vsf(1).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(1).AppendRows = True
    Case 3
        vsf(2).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(2).AppendRows = True
    End Select
End Sub

Private Sub tbcQuestion_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mintIndex = Item.Index
    Call ExecuteCommand("��ȡ��������")
End Sub

Private Sub tmr_Timer()
        
    mlngLoop = mlngLoop + 1
    If mlngLoop = mlngIntenal And mlngIntenal > 0 Then
    
        '�Զ�ˢ��δ����
        
        Call ExecuteCommand("��ȡδ����")
        
        If tbcQuestion.Item(1).Selected Then Call ExecuteCommand("��ȡ��������")
        mlngIntenal = Val(GetPara("δ����ˢ��Ƶ��", mfrmMain.ģ���, "5"))
        
    End If
    
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        Call ExecuteCommand("��ȡ��������")
    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey)
    vsf(Index).ShowCell vsf(Index).Row, vsf(Index).Col
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mclsVsf(Index).SaveKey = Val(vsf(Index).RowData(vsf(Index).Row))
End Sub

Private Sub vsf_DblClick(Index As Integer)
    With vsf(Index)
        RaiseEvent LocationDocument(Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), Val(.TextMatrix(.Row, .ColIndex("��������id"))), Val(.TextMatrix(.Row, .ColIndex("�ļ�id"))), Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), Val(.TextMatrix(.Row, .ColIndex("����id"))))
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsf_DblClick(Index)
    End If
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar     As CommandBar
    Dim cbrPopupItem    As CommandBarControl
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)

        If Not mclsVsf(Index).MoveColumn Then
            
            '�����˵�����
            Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "���ӷ���")
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_CopyNewItem, "�ٴη���")
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������")
            
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Send, "��ɷ���"): cbrPopupItem.BeginGroup = True
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SendBack, "�������")
            
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Transf_Save, "�������"): cbrPopupItem.BeginGroup = True
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ������")
            
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Filter, "��ǰ����"): cbrPopupItem.BeginGroup = True
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "ˢ�·���")

            cbrPopupBar.ShowPopup
        End If

    End Select
End Sub

'��ȡ����������
Private Function GetMaxNum(ByVal str��鿪ʼʱ�� As String, ByVal str������ʱ�� As String) As Long
    On Error GoTo errH
    Dim rsData          As ADODB.Recordset
    Dim strSQL          As String
    Dim strStart        As String
    Dim strEnd          As String
    strSQL = "select max(��������) as ���� from ����������¼ where ����ʱ�� BetWeen [1] And [2]"
    strStart = str��鿪ʼʱ�� ' GetDateTime("��  ��", 1)
    strEnd = str������ʱ�� ' GetDateTime("��  ��", 2)
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����������", CDate(strStart), CDate(strEnd))

    If rsData.EOF = False Then
        GetMaxNum = IIf(IsNull(rsData!����), 1, rsData!���� + 1)
    Else
        GetMaxNum = 2
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
End Function

'�Զ���ȡ�ò��˵ĵ�������ִ���
Private Function GetMaxNumNEW(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng��ʽ As Long) As Long
'��һ����Ȼ����Ϊһ�����ִΣ���һ����Ȼ�������е��������ۼӺϼ�����(����+����+���ּ����������ִ�)
    On Error GoTo errH
    Dim rsData          As ADODB.Recordset
    Dim strSQL          As String
    Dim strNow          As String

    strNow = zlDatabase.Currentdate
    
    strSQL = "Select Sum(A.���մ���) as ���մ���,Sum(A.����������) as ���������� From (" & vbNewLine & _
        "select max(��������) as ���մ���,0 AS ���������� from ����������¼ where ����ID=[1] And ��ҳID=[2] And nvl(���ּ���,0)=[3]" & vbNewLine & _
        "And ����ʱ�� BetWeen To_Date([4], 'yyyy-mm-dd') And To_Date([4], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
        "Union All" & vbNewLine & _
        "select 0 AS ���մ���,max(��������) AS ���������� from ����������¼ where ����ID=[1] And ��ҳID=[2] And nvl(���ּ���,0)=[3]" & vbNewLine & _
        "And ����ʱ��< To_Date([4], 'yyyy-mm-dd')) A"

    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����������", lng����ID, lng��ҳID, lng��ʽ, Format(strNow, "yyyy-mm-dd"))

    If rsData.EOF = False Then
        If rsData!���մ��� > 0 Then
            GetMaxNumNEW = rsData!���մ���
        Else
            If rsData!���������� > 0 Then
                GetMaxNumNEW = rsData!���������� + 1
            Else
                GetMaxNumNEW = 1
            End If
        End If
    Else
        GetMaxNumNEW = 1
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
End Function


'��ȡ������������Ϣ
Private Sub Init������(ByVal str��鿪ʼʱ�� As String, ByVal str������ʱ�� As String)
    On Error GoTo errH
        Dim rs As ADODB.Recordset
        Dim lngCount As Long '��¼����
        mblnRef = True
        cbo������.Clear
        lngCount = 0
        gstrSQL = "select distinct(��������),Sum(A.��ֵ) as �ܿ۷���,Min(A.����ʱ��) as ���練��ʱ�� from ����������¼ A where A.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400 group by A.�������� order by A.�������� Asc"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(str��鿪ʼʱ��, "yyyy-mm-dd"), Format(str������ʱ��, "yyyy-mm-dd"))
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            cbo������.AddItem "����"
            
                Do Until rs.EOF
                If NVL(rs!��������, 0) = 0 Then
                    cbo������.AddItem "��" & NVL(rs!��������, 0) & "��-" & Format(NVL(rs!���練��ʱ��, Now()), "YYYY-MM-DD") & "(" & NVL(rs!�ܿ۷���, 0) & ")"
                    cbo������.ItemData(cbo������.NewIndex) = NVL(rs!��������, 0)
                End If
                rs.MoveNext
            Loop
            
            rs.MoveFirst
            Do Until rs.EOF
                    If lngCount >= 10 Then Exit Do
'                        Call AddComboData(cbo������, rs, "���練��ʱ��", "����", , False)
                        If NVL(rs!��������, 0) <> 0 Then
                            cbo������.AddItem "��" & NVL(rs!��������, 0) & "��-" & Format(NVL(rs!���練��ʱ��, Now()), "YYYY-MM-DD") & "(" & NVL(rs!�ܿ۷���, 0) & ")"
                            cbo������.ItemData(cbo������.NewIndex) = NVL(rs!��������, 0)
                        End If
                    lngCount = lngCount + 1
                    rs.MoveNext
            Loop
            cbo������.ListIndex = 0
        Else
            cbo������.AddItem "����"
            cbo������.ListIndex = 0
            cbo������.ItemData(cbo������.NewIndex) = 0
        End If
        mblnRef = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnRef = False
    Err.Clear
End Sub

Private Sub SetCob(ByVal lngCurNum As Long)
    Dim i As Integer
    mblnRef = True
    For i = 0 To cbo������.ListCount - 1
        If cbo������.ItemData(i) = lngCurNum Then
            cbo������.ListIndex = i
            mblnRef = False
            Exit Sub
        End If
    Next
    mblnRef = False
End Sub

Private Function GetAnalyseTime(ByVal strTime As String, ByVal lngMode As Long) As String
    Dim strTemp As String
    Dim i As Integer
    '��ȡʱ��ֵ
    i = InStrRev(strTime, "��")
    If i > 0 Then
        strTemp = Right(strTime, Len(strTime) - i - 1)
        i = InStrRev(strTemp, "(")
        If i > 0 Then
            strTemp = Left(strTemp, i - 1)
            
            Select Case lngMode
            Case 1
                GetAnalyseTime = Format(strTemp, "yyyy-MM-dd 00:00:00")
            Case 2
                GetAnalyseTime = Format(strTemp, "yyyy-MM-dd 23:59:59")
            End Select
        End If
    End If
    
End Function
