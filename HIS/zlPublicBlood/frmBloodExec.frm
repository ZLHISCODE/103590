VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBloodExec 
   BorderStyle     =   0  'None
   Caption         =   "ִ�еǼ�"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VSFlex8Ctl.VSFlexGrid vsExec 
      Height          =   1485
      Left            =   -30
      TabIndex        =   0
      Top             =   825
      Width           =   7125
      _cx             =   12568
      _cy             =   2619
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
      BackColorSel    =   16444122
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin ComctlLib.ImageList imgList 
      Left            =   1635
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBloodExec.frx":0000
            Key             =   "δִ��"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBloodExec.frx":059A
            Key             =   "��ִ��"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBloodExec.frx":0B34
            Key             =   "�ܾ�ִ��"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBloodExec.frx":10CE
            Key             =   "����ִ��"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsExec 
      Left            =   0
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBloodExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsVsf As clsVsf
Private mlngSys As Long
Private mlngModul As Long
Private mlngҽ��ID As Long
Private mlngҽ������ID As Long
Private mlng����ID As Long, mint���� As Integer
Private mlngFontSize As Long
'ҽ���������
Private mlng���ͺ� As Long, mintִ��״̬ As Integer
Private mint��¼���� As Integer, mint������� As Integer, mstrNO As String, mlngִ�в���ID As Long
Private mint�Ʒ�״̬ As Integer, mlng����ID As Long, mlng��ҳid As Long, mlng��ID As Long, mstr����ʱ�� As String

Private mstrPrivs As String
Private mblnMoved As Boolean
Private mblnLoad As Boolean
Private mfrmParent As Object

Private mblnShowExec As Boolean
Private mblnExecFresh As Boolean  '�Ƿ���ִ�й���ˢ��(��Ҫ�Ǳ����ظ�ˢ�£��ı����ڵ��ô��ڸ�ֵ)
Private Enum CMD_EXEC
    ID_��ʾִ�� = 1
    ID_���ִ�� = 2
    ID_ȡ����� = 3
    ID_ִ�м�¼ = 4
    ID_ִ�е��� = 5
    ID_ִ��ɾ�� = 6
    ID_ִ��ǰ�˶� = 7 '��Ѫǰ�˶�
    ID_ȡ��ִ��ǰ�˶� = 8
    ID_ִ���к˶� = 9 'ִ���к˶�
    ID_ȡ��ִ���к˶� = 10
End Enum

Private Enum Enum_ExecState
    E_���״̬ = 0
    E_��¼ִ�� = 1
    E_ɾ��ִ�� = 2
    E_ִ����� = 3
    E_ȡ����� = 4
    E_ִ�к˶� = 5
    E_ȡ���˶� = 6
End Enum
Private mintAdviceExecState As Enum_ExecState

Public Event ShowExec(ByVal blnShow As Boolean, ByVal lngHeight As Long)

Public Property Get AdviceExecState() As Integer
'ִ��״̬(���ٴ�����ˢ��ʹ��)
    AdviceExecState = mintAdviceExecState
End Property

Public Property Let AdviceExecState(intAdviceExecState As Integer)
    mintAdviceExecState = intAdviceExecState
End Property

Public Property Let ExecFresh(blnFresh As Boolean)
'�Ƿ���ִ�й���ˢ��
    mblnExecFresh = blnFresh
End Property

Public Property Get IsShowExec() As Boolean
    IsShowExec = mblnShowExec
End Property

Public Property Let IsShowExec(blnValue As Boolean)
    Call SetShowExec(blnValue)
    mblnShowExec = blnValue
    RaiseEvent ShowExec(mblnShowExec, Me.Height)
End Property

Public Function zlRefresh(ByVal frmParent As Object, ByVal lngSys As Long, ByVal lngModul As Enum_Inside_Program, ByVal lngҽ��ID As Long, ByVal lngҽ������ID As Long, ByVal strPrivs As String, _
   Optional ByVal int���� As Integer = 2, Optional ByVal lng����ID As Long, Optional ByVal blnMoved As Boolean = False, Optional ByVal lngFontSize As Long = 9) As Boolean
'���ܣ�ˢ�¶�Ӧҽ����ѪҺ��Ϣ
'frmParent ���ö��������壬��Ҫ�ǹ�����ˢ��ʹ��(�ô���Ҫ�����һ��timer�ؼ�������ΪtimBRefresh)
' int���� =1 �������,2-סԺ����, ������=2ʱ������lng����ID
    Dim lngCount As Long
    
    If mblnExecFresh = False Then  'ִ�й����б����ظ�ˢ�£�ӦΪ�������ڲ��Ѿ�������ˢ��
        Set mfrmParent = frmParent
        mlngSys = lngSys
        mlngModul = lngModul
        mlngҽ��ID = lngҽ��ID
        mlngҽ������ID = lngҽ������ID
        mint���� = int����
        mlng����ID = lng����ID
        mstrPrivs = strPrivs
        mblnMoved = blnMoved
        mlngFontSize = lngFontSize
        mlng���ͺ� = 0
        
        If mint���� = 2 Then
             'ɾ�����ڵĹ������������˵���
            For lngCount = cbsExec.ActiveMenuBar.Controls.Count To 1 Step -1
                cbsExec.ActiveMenuBar.Controls(lngCount).Delete
            Next
            For lngCount = cbsExec.Count To 2 Step -1
                cbsExec(lngCount).Delete
            Next
            Call InitExecBar
        End If
        Call RefreshData
        Call SetFontSize(mlngFontSize)
    End If
    mblnExecFresh = False
    zlRefresh = True
End Function

Private Sub InitTable()
'����ʼ��
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsExec, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)  '�շ�ID
        Call .AppendColumn("״̬", 810, flexAlignLeftCenter, flexDTString, , "ѪҺ״̬") '����ִ��״̬
        Call .AppendColumn("ѪҺ����", 1800, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("���", 810, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("ABO", 810, flexAlignLeftCenter, flexDTString, , "ABO", True)
        Call .AppendColumn("Rh(D)", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
        Call .AppendColumn("Ѫ�����", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("Ч��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", "ѪҺЧ��", True)
        Call .AppendColumn("����", 500, flexAlignRightCenter, flexDTDecimal, , , , , , False)
        
        'ִ�м�¼
        Call .AppendColumn("�˲���", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("������", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("�˲�ʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
'        Call .AppendColumn("ִ�п���", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��ʼִ����", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��ʼʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
'        Call .AppendColumn("ǰ15���ӵ���", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("��עǰ��Ѫ��Ӧ", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("��עǰ��Ӧʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
'        Call .AppendColumn("��ע��ִ����", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("��15���ӵ���", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("��ע����Ѫ��Ӧ", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("��ע��Ӧʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("����ִ����", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("����ʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        
        If Mid(gstrҽ���˶�, 1, 1) = "1" Then
            Call .AppendColumn("�˶���", 1200, flexAlignLeftCenter, flexDTString)
            Call .AppendColumn("�˶�ʱ��", 1500, flexAlignLeftCenter, flexDTString)
        Else
            Call .AppendColumn("�˶���", 0, flexAlignLeftCenter, flexDTString)
            Call .AppendColumn("�˶�ʱ��", 0, flexAlignLeftCenter, flexDTString)
        End If
        
        'ѪҺ�䷢��Ϣ
        Call .AppendColumn("��Ѫ����", 1500, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("��Ѫ����", 1500, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��Ѫ����", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��Ѫ��", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("ȡѪ��", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��Ѫʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("������", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("������", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("����ʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        
        
        '������
        Call .AppendColumn("ѪҺID", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("ѪҺЧ����ɫ", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("����״̬", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("ִ��״̬", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("��ִ�п���ID", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("ִ�п���ID", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        vsExec.FrozenCols = vsExec.ColIndex("���")
        .AppendRows = False
    End With
End Sub

Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
   
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    Set cbsExec.Icons = gobjCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
    
    Set objBar = cbsExec.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap '+ xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_��ʾִ��, "��ʾִ������")
        Set objControl = .Add(xtpControlButton, ID_���ִ��, "ִ�����")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_Manage_Complete
        Set objControl = .Add(xtpControlButton, ID_ȡ�����, "ȡ�����")
            objControl.IconId = conMenu_Edit_Untread
        
        Set objControl = .Add(xtpControlButton, ID_ִ��ǰ�˶�, "�˲�")
            objControl.IconId = conMenu_Manage_ThingAudit
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_ȡ��ִ��ǰ�˶�, "ȡ���˲�")
            objControl.IconId = conMenu_Manage_ThingDelAudit
            
        Set objControl = .Add(xtpControlButton, ID_ִ�м�¼, "��¼ִ�����")
            objControl.IconId = conMenu_Manage_ThingAdd
        Set objControl = .Add(xtpControlButton, ID_ִ��ɾ��, "ɾ��ִ�����")
            objControl.IconId = conMenu_Manage_ThingDel

        Set objControl = .Add(xtpControlButton, ID_ִ���к˶�, "�˶�")
            objControl.IconId = conMenu_Manage_ThingAudit
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_ȡ��ִ���к˶�, "ȡ���˶�")
            objControl.IconId = conMenu_Manage_ThingDelAudit
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsExec.KeyBindings
        '.Add FCONTROL, vbKeyH, 0
    End With
End Sub

Private Sub RefreshData(Optional ByVal blnRefreshBlood As Boolean = True)
    '����:ˢ�¶�Ӧҽ����Ӧ��ѪҺ��Ϣ
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim lngRow As Long
    Dim lngSelectRowID As Long
    Dim lng��ID As Long
    Dim arrData, arrTmp() As String, i As Integer
    Dim strID As String
    On Error GoTo ErrHand
    If mblnLoad = False Then Call InitTable
    strSQL = _
        " Select A.���ͺ�,B.���ID,A.����ʱ��,A.ִ��״̬,A.��¼����,A.NO,A.ִ�в���ID,A.�Ʒ�״̬,A.�������,B.����ID,B.��ҳID ,C.ִ��ʱ��,C.�˶���,C.�˶�ʱ��" & vbNewLine & _
        " From ����ҽ��ִ�� C,����ҽ������ A,����ҽ����¼ B" & vbNewLine & _
        " Where A.ҽ��ID=C.ҽ��ID(+) and a.���ͺ�=c.���ͺ�(+) and a.ҽ��ID=b.ID And b.id=[1]"
    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "�ѷ�ѪҺ��Ϣ��ȡ", mlngҽ��ID)
    If rsData.EOF Then
        MsgBox "��ҽ����δ���ͣ����ܽ���ִ�еǼǣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    lng��ID = Val("" & rsData!���ID)
    mlng���ͺ� = rsData!���ͺ�
    mstr����ʱ�� = Format(rsData!����ʱ�� & "", "YYYY-MM-DD HH:mm:ss")
    mintִ��״̬ = Val("" & rsData!ִ��״̬)
    mint��¼���� = Val("" & rsData!��¼����)
    mstrNO = "" & rsData!NO
    mlngִ�в���ID = Val("" & rsData!ִ�в���ID)
    mint�Ʒ�״̬ = Val("" & rsData!�Ʒ�״̬)
    mint������� = Val("" & rsData!�������)
    mlng����ID = Val("" & rsData!����id)
    mlng��ҳid = Val("" & rsData!��ҳid)
    mlng��ID = lng��ID
    
    arrData = Array()
    Do While Not rsData.EOF
        ReDim Preserve arrData(UBound(arrData) + 1)
        arrData(UBound(arrData)) = "" & rsData!ִ��ʱ�� & "'" & rsData!�˶��� & "'" & rsData!�˶�ʱ��
    rsData.MoveNext
    Loop
    
    'ˢ��ǰȷ��֮ǰѡ���ѪҺ
    If vsExec.Row >= vsExec.FixedRows And vsExec.Row < vsExec.Rows Then
        lngSelectRowID = Val(vsExec.RowData(vsExec.Row))
    End If
    
    If blnRefreshBlood = False Then Exit Sub
    strSQL = _
        " Select b.����id, a.Id, a.ѪҺid, a.Abo, a.Rh, To_Char(a.Ч��, 'YYYY-MM-DD hh24:mi') ѪҺЧ��, a.��ɫ ѪҺ��ɫ, a.��� Ѫ�����, a.��Ѫ��," & vbNewLine & _
        " Decode(Zl_ѪҺʧЧ_Check(k.Ч�ڱ���,k.Ч�ڵ�λ,a.Ч��),0," & COLOR.ԭʼ���� & ",1," & COLOR.���ɫ & ",2," & COLOR.��ɫ & ") ѪҺЧ����ɫ," & vbNewLine & _
        "       To_Char(a.��Ѫ����, 'YYYY-MM-DD hh24:mi') ��Ѫʱ��, Nvl(a.��Ѫ״̬, 0) ��Ѫ״̬, c.���� ��Ѫ����, a.Ѫ�����," & vbNewLine & _
        "       Decode(Nvl(h.ִ��״̬, 0)," & vbNewLine & _
        "               0," & vbNewLine & _
        "               " & IIf(gbln���պ����ִ�� = True, "Decode(Nvl(h.����״̬, 0), 0, '������', 2, '�ܾ�����', '�ѽ���'),", "'�ȴ�ִ��',") & vbNewLine & _
        "               1," & vbNewLine & _
        "               '����ִ��'," & vbNewLine & _
        "               2," & vbNewLine & _
        "               '���ִ��'," & vbNewLine & _
        "               3," & vbNewLine & _
        "               'ִֹͣ��') ѪҺ״̬, a.ʵ������ As ����, e.���� As ѪҺ����, e.���," & vbNewLine & _
        "       (Select f_List2str(Cast(Collect(g.����) As t_Strlist))" & vbNewLine & _
        "         From ������ĿĿ¼ g, ѪҺ��Ѫ���� f" & vbNewLine & _
        "         Where f.��Ѫ����id = g.Id(+) And f.�շ�id = a.Id) ��Ѫ����," & vbNewLine & _
        "       (Select Max(f.��Ѫ����) From ������ĿĿ¼ g, ѪҺ��Ѫ���� f Where f.��Ѫ����id = g.Id(+) And f.�շ�id = a.Id) ��Ѫ����, h.������ ��Ѫ��, h.ȡѪ��," & vbNewLine & _
        "       To_Char(h.����ʱ��, 'YYYY-MM-DD hh24:mi') ��Ѫʱ��, h.����״̬, h.������, h.����ʱ��, h.������, h.����ʱ��, h.ִ��״̬, h.ִ�п���id ��ִ�п���id," & vbNewLine & _
        "       g.���� ��ִ�в���,h.ִ�к˶��� �˲���, h.ִ�и����� ������, h.ִ�к˶�ʱ�� �˲�ʱ��" & vbNewLine & _
        " From ���ű� c, �շ���ĿĿ¼ e, ѪҺƷ�� k, ѪҺ��� l, ѪҺ�շ���¼ a, ���ű� g, ѪҺ���ͼ�¼ h, ѪҺ��Ѫ��¼ b" & vbNewLine & _
        " Where c.Id = a.�ⷿid And e.Id = a.ѪҺid And k.Ʒ��id = l.Ʒ��id And l.���id = a.ѪҺid And a.Id = h.�շ�id  And" & vbNewLine & _
        "      h.ִ�п���id = g.Id(+) And h.�䷢id = b.Id And b.����id = [1]" & vbNewLine & _
        " Order By a.��Ѫ����, a.���"

    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "�ѷ�ѪҺ��Ϣ��ȡ", lng��ID)
    Call mclsVsf.LoadGrid(rsData, "", True)
    strID = ""
    For lngRow = vsExec.FixedRows To vsExec.Rows - 1
        Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = Nothing
        If Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ID"))) > 0 Then
            strID = strID & "," & Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ID")))
            Select Case Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ִ��״̬")))
                '0-δִ��;1-����ִ��;2-���ִ��;3-ִֹͣ��
                Case 0
                   Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgList.ListImages("δִ��").Picture
                Case 1
                    Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgList.ListImages("����ִ��").Picture
                Case 2
                    Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgList.ListImages("��ִ��").Picture
                Case 3
                    Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgList.ListImages("�ܾ�ִ��").Picture
            End Select
            vsExec.Cell(flexcpForeColor, lngRow, mclsVsf.ColIndex("ѪҺЧ��"), lngRow, mclsVsf.ColIndex("ѪҺЧ��")) = Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ѪҺЧ����ɫ")))
            
            '��λ���ϴ�ѡ����д�
            If Val(vsExec.RowData(lngRow)) = lngSelectRowID And lngSelectRowID > 0 Then
                vsExec.Row = lngRow
            End If
        End If
    Next lngRow
    If Left(strID, 1) = "," Then strID = Mid(strID, 2)
    If strID <> "" Then
        '��ȡѪҺִ����Ϣ
        strSQL = "Select /*+ CARDINALITY(B 10)*/ �շ�id, ��¼����, ���, ִ��ʱ��, ִ����, ִ�п���id, ����, ��Ѫ��Ӧ, ��Ӧʱ��, ��Ѫ��λ�Ƿ���© �Ƿ���©, �Ƿ�ʹ��ҩ��, ����, ����, ����, ����ѹ, ����ѹ, ժҪ, �Ǽ���, �Ǽ�ʱ��," & vbNewLine & _
            "       ǩ����, ǩ��ʱ��" & vbNewLine & _
            " From ѪҺִ�м�¼ a,Table(f_str2list([1])) B" & vbNewLine & _
            " Where a.�շ�id = b.Column_Value" & vbNewLine & _
            " Order By �շ�id, ��¼����, ���"
        Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "�ѷ�ѪҺ��Ϣ��ȡ", strID)
        For lngRow = vsExec.FixedRows To vsExec.Rows - 1
            If Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ID"))) > 0 Then
                rsData.Filter = "�շ�ID=" & Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ID")))
                rsData.Sort = "��¼����,���"
                Do While Not rsData.EOF
                    Select Case Val("" & rsData!��¼����)
                        Case 1
                            vsExec.TextMatrix(lngRow, vsExec.ColIndex("��ʼִ����")) = rsData!ִ���� & ""
                            vsExec.TextMatrix(lngRow, vsExec.ColIndex("��ʼʱ��")) = Format("" & rsData!ִ��ʱ��, "YYYY-MM-DD HH:mm")
                        Case 3
                            vsExec.TextMatrix(lngRow, vsExec.ColIndex("����ִ����")) = rsData!ִ���� & ""
                            vsExec.TextMatrix(lngRow, vsExec.ColIndex("����ʱ��")) = Format("" & rsData!ִ��ʱ��, "YYYY-MM-DD HH:mm")
                    End Select
                rsData.MoveNext
                Loop
                '����ִ���к˶���Ϣ
                For i = 0 To UBound(arrData)
                    arrTmp = Split(CStr(arrData(i)), "'")
                    If Format(vsExec.TextMatrix(lngRow, vsExec.ColIndex("��ʼʱ��")), "YYYY-MM-DD HH:mm:ss") = Format(arrTmp(0), "YYYY-MM-DD HH:mm:ss") Then
                        vsExec.TextMatrix(lngRow, vsExec.ColIndex("�˶���")) = arrTmp(1)
                        vsExec.TextMatrix(lngRow, vsExec.ColIndex("�˶�ʱ��")) = arrTmp(2)
                        Exit For
                    End If
                Next i
            End If
        Next lngRow
    End If
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CheckDataMoved() As Boolean
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        CheckDataMoved = True
    End If
End Function

Private Function CheckItemOk() As Boolean
'���ܣ�������е���Ŀ�Ƿ��Ѿ�ִ�����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngTmp As Long
    
    strSQL = "Select a.�շ�id" & vbNewLine & _
    " From ѪҺ���ͼ�¼ a, ѪҺ��Ѫ��¼ b" & vbNewLine & _
    " Where a.�䷢id = b.Id And (Nvl(a.ִ��״̬, 0) = 0 or Nvl(a.ִ��״̬, 0) = 1) And b.����id = [1] And Rownum < 2"
    On Err GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "������˼��", mlng��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "��ҽ��������δִ�е�ѪҺ��¼���������ִ�С�", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Mid(gstrҽ���˶�, 1, 1)) = 1 Then
        strSQL = "Select �˶��� From ����ҽ��ִ�� Where ҽ��id = [1] And ���ͺ� = [2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�)
        If rsTmp.RecordCount = 1 Then
            If rsTmp!�˶��� & "" = "" Then
                MsgBox "��ҽ��������δ�˶Ե�ִ�еǼǣ�����˶��˲�����ɡ�", vbInformation, gstrSysName
                Exit Function
            End If
        ElseIf rsTmp.RecordCount > 1 Then
            lngTmp = rsTmp.RecordCount
            rsTmp.Filter = "�˶���<>''"
            If lngTmp <> rsTmp.RecordCount Then
                MsgBox "��ҽ��������δ�˶Ե�ִ�еǼǣ�����ȫ���˶��˲�����ɡ�", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "��ǰҽ����δ��¼ִ������������¼ִ�������˶��˲�����ɡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckItemOk = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckPatiIsAduit() As Boolean
'���ܣ���鲡���Ƿ�ʼ���
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim int��˱�־ As Integer
    
    If mlng��ҳid = 0 Then CheckPatiIsAduit = True: Exit Function
    strSQL = "Select a.��˱�־ From ������ҳ a" & _
                " Where a.����ID=[1] And a.��ҳID=[2]"
    On Err GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "������˼��", mlng����ID, mlng��ҳid)
    If rsTmp.RecordCount > 0 Then
        If Val("" & rsTmp!��˱�־) >= 1 And gbyt������˷�ʽ = 1 Then
            MsgBox "�ò��˵ķ���������˻��Ѿ���ˣ����������ҽ���ͷ��á�", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPatiIsAduit = True
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetShowExec(ByVal blnShow As Boolean)
    vsExec.Visible = blnShow
    If blnShow = True Then
        vsExec.Tag = "�ɼ�"
        Me.Height = vsExec.Top + vsExec.Height
    Else
        vsExec.Tag = ""
        Me.Height = vsExec.Top
    End If
End Sub

Public Sub SetFontSize(ByVal lngFontSize As Long)
'����:����ҽ���嵥�������С
    Dim bytSize As Byte
    bytSize = IIf(lngFontSize = 9, 0, 1)
    Call SetPublicFontSize(Me, bytSize)
End Sub

Private Function FuncExec(ByVal intExecId As CMD_EXEC) As Boolean
    Dim strSQL As String, rsTmp As New Recordset
    Dim byt��Դ As Integer, blnIsAbnormal As Boolean
    Dim curMoney As Currency, str��� As String, str����� As String
    Dim lngID As Long, strִ��ʱ�� As String, str�˶��� As String
    Dim blnTrans As Boolean, blnOk As Boolean
    Dim i As Integer
    Dim arrSQL As Variant
    Dim blnFinish As Boolean
    '�˶Խ��
    Dim strCheckOper As String, strCheckTime As String, strCheckResult As String
    On Error GoTo ErrHand
    Select Case intExecId
        Case ID_ִ��ǰ�˶�, conMenu_Manage_ThingAudit * 100# + 1
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˲���")) <> "" Then
                MsgBox "�ô�ѪҺ�Ѿ��˶ԣ��������ٴκ˶ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
            If gbln���պ����ִ�� = True Then
                If Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("����״̬"))) <> 1 And Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("����״̬"))) <> 3 Then
                    MsgBox "�ô�ѪҺ��δ���գ�������˶ԡ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            strִ��ʱ�� = Format(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("����ʱ��")), "YYYY-MM-DD HH:mm")
            lngID = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID")))
            blnOk = frmUserCheck.ShowMe(Me, mlngModul, mlngҽ������ID, mlngҽ������ID, strִ��ʱ��, "", True, ִ�к˶�)
            If blnOk = True Then
                strCheckOper = frmUserCheck.SendAndTakeOper
                strCheckTime = frmUserCheck.SendTime
                strCheckResult = frmUserCheck.CheckResult
                strSQL = "Zl_ѪҺִ�м�¼_Check(" & lngID & ",'" & Split(strCheckOper, "'")(0) & "','" & Split(strCheckOper, "'")(1) & "',To_Date('" & strCheckTime & "','YYYY-MM-DD HH24:MI:SS'),'" & strCheckResult & "')"
                Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Else
                Exit Function
            End If
        Case ID_ȡ��ִ��ǰ�˶�, conMenu_Manage_ThingDelAudit * 100# + 1
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˲���")) = "" Then
                MsgBox "�ô�ѪҺ��δ�˶ԣ�����ȡ���˶ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
            If Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ִ��״̬"))) <> 0 Then
                MsgBox "�ô�ѪҺ�Ѿ���ʼִ�У�����ȡ���˶ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
            strCheckOper = ""
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˲���")) <> UserInfo.���� And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("������")) <> UserInfo.���� Then
                strCheckOper = gobjDatabase.UserIdentifyByUser(Me, "��ȡ���˶�ǰ�������������û�����������������֤��", mlngSys, mlngModul, "ִ������Ǽ�", , True)
                If strCheckOper = "" Then Exit Function
                If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˲���")) <> strCheckOper And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("������")) <> strCheckOper Then
                    MsgBox "ֻ��ȡ���Լ��˶Ի򸴲��ѪҺ����ǰѪҺ�˶�����""" & vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˲���")) & """" & "��������""" & vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("������")) & """", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If MsgBox("��ȷ��Ҫȡ���˶���", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
            End If
            lngID = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID")))
            strSQL = "Zl_ѪҺִ�м�¼_Uncheck(" & lngID & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Case ID_ִ���к˶�, conMenu_Manage_ThingAudit 'ִ���к˶�
            If Not Mid(gstrҽ���˶�, 1, 1) = "1" Then
                MsgBox "���ܺ˶���Ѫҽ�������ڻ��������й�ѡ��˶���Ѫҽ��������", vbInformation, gstrSysName
                Exit Function
            End If
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˶���")) <> "" Then
                MsgBox "�ô�ѪҺ�Ѿ��˶ԣ������ٴκ˶ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
            If Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ִ��״̬"))) = 0 Or Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ִ��״̬"))) = 1 Then
                MsgBox "�ô�ѪҺ��δ���ִ������Ǽǣ����ܺ˶ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
            str�˶��� = gobjDatabase.UserIdentifyByUser(Me, "�ں˶�ִ�����ǰ�������������û�����������������֤��", 100, IIf(mint���� = 1, pҽ������վ, pסԺҽ������), "ִ������Ǽ�", , True)
            If str�˶��� = "" Then Exit Function
            
            If str�˶��� = vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("��ʼִ����")) Then
                MsgBox "ִ���˲��ܺ��������ͬ�����ܺ˶ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            strִ��ʱ�� = Format(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("��ʼʱ��")), "yyyy-MM-dd HH:mm:ss")
            strSQL = "Zl_����ҽ���˶�_Insert(" & mlngҽ��ID & "," & mlng���ͺ� & ",'" & str�˶��� & "',To_Date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS'))"
            Call gobjDatabase.ExecuteProcedure(strSQL, "ҽ���˶�")
            Call SetExecState(E_ִ�к˶�)
        Case ID_ȡ��ִ���к˶�, conMenu_Manage_ThingDelAudit 'ȡ��ִ���к˶�
            If Not Mid(gstrҽ���˶�, 1, 1) = "1" Then
                MsgBox "����ȡ����Ѫҽ���˶ԣ����ڻ��������й�ѡ��˶���Ѫҽ��������", vbInformation, gstrSysName
                Exit Function
            End If
            
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˶���")) = "" Then
                MsgBox "�ô�ѪҺ��δ�˶ԣ�����ȡ���˶ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˶���")) <> UserInfo.���� Then
                str�˶��� = gobjDatabase.UserIdentifyByUser(Me, "��ȡ���˶�ǰ�������������û�����������������֤��", 100, IIf(mint���� = 1, pҽ������վ, pסԺҽ������), "ִ������Ǽ�", , True)
                If str�˶��� = "" Then Exit Function
                If str�˶��� <> vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˶���")) Then
                    MsgBox "ֻ��ȡ���Լ��˶Ե�ѪҺִ�У���ǰѪҺ�˶�����""" & vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˶���")) & """", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If MsgBox("��ȷ��Ҫȡ���˶���", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
            End If
            strSQL = "Zl_����ҽ���˶�_Delete(" & mlngҽ��ID & "," & mlng���ͺ� & ",To_Date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS'))"
            Call gobjDatabase.ExecuteProcedure(strSQL, "ȡ��ҽ���˶�")
            Call SetExecState(E_ȡ���˶�)
        Case ID_ִ�м�¼, conMenu_Manage_ThingAdd
            If mintִ��״̬ = 1 Then
                MsgBox "��ҽ����ǰ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
                Exit Function
            End If
            If Val(Mid(gstrҽ���˶�, 1, 1)) > 0 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˶���")) <> "" Then
                MsgBox "�ô�ѪҺ�Ѿ��˶ԣ���ȡ���˶Ժ����ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
            If CheckDataMoved Then Exit Function
            lngID = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID")))
            blnFinish = False
            If frmBloodExecEdit.ShowEdit(Me, mlngModul, mlngҽ��ID, mlng���ͺ�, mlngҽ������ID, lngID, mlngִ�в���ID, mstrPrivs, , blnFinish) = False Then
                Exit Function
            End If
            If blnFinish = True Then
                Call SetExecState(E_ִ�����)
            Else
                Call SetExecState(E_��¼ִ��)
            End If
        Case ID_ִ��ɾ��, conMenu_Manage_ThingDel
            If mintִ��״̬ = 1 Then
                MsgBox "��ҽ����ǰ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
                Exit Function
            End If
            
            If Val(Mid(gstrҽ���˶�, 1, 1)) > 0 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˶���")) <> "" Then
                MsgBox "�ô�ѪҺ�Ѿ��˶ԣ���ȡ���˶Ժ����ԡ�", vbInformation, gstrSysName
                Exit Function
            End If
        
            If CheckDataMoved Then Exit Function
            If MsgBox("ȷʵҪɾ���ô�ѪҺ��ִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            lngID = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID")))
            strִ��ʱ�� = vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("��ʼʱ��"))
            '�Ƿ���ڸ�Ѫ����Ϣ����������Ϊ�Ѷ�
            If mint���� = 1 Then
                strSQL = "select a.id,a.���ͱ���,a.����id,a.ҵ���ʶ from ҵ����Ϣ�嵥 a,����ҽ����¼ b,���˹Һż�¼ c where a.����id = [1] and a.����id = c.id and c.no = b.�Һŵ� and b.id = [2]" & vbNewLine & _
                        "and a.�Ƿ����� = 0 and a.���ͱ��� in ('ZLHIS_BLOOD_006','ZLHIS_BLOOD_007')  "
            ElseIf mint���� = 2 Then
                strSQL = "select a.id,a.���ͱ���,a.����id,a.ҵ���ʶ from ҵ����Ϣ�嵥 a,����ҽ����¼ b where a.����id = [1] and a.����id = b.��ҳid and b.id = [2]" & vbNewLine & _
                        "and a.�Ƿ����� = 0 and a.���ͱ��� in ('ZLHIS_BLOOD_006','ZLHIS_BLOOD_007')  "
            End If
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "Ѫ�������Ϣ", mlng����ID, mlngҽ��ID)
            arrSQL = Array()
            For i = 0 To 1
                '��ZLHIS_BLOOD_006����Ϣ��Ϊ�Ѷ�
                If i = 0 Then rsTmp.Filter = "ҵ���ʶ = '" & mlngҽ��ID & ":" & Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("id"))) & "'"
                '��ZLHIS_BLOOD_007����Ϣ��Ϊ�Ѷ�
                If i = 1 Then rsTmp.Filter = "ҵ���ʶ = '" & mlng��ID & ":" & mlngҽ��ID & ":" & Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("id"))) & "'"
                If Not rsTmp.EOF Then
                    rsTmp.MoveFirst
                    Do While Not rsTmp.EOF
                        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & mlng����ID & "," & rsTmp!����id & ",'" & rsTmp!���ͱ��� & "',"
                        strSQL = strSQL & IIf(mint���� = 1, 4, 3) & ",'" & UserInfo.���� & "'," & mlng����ID & ",NULL,"
                        strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                        rsTmp.MoveNext
                    Loop
                End If
            Next
            gcnOracle.BeginTrans
            blnTrans = True
            strSQL = "ZL_����ҽ��ִ��_Delete(" & mlngҽ��ID & "," & mlng���ͺ� & ",To_Date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),0,0," & mlng����ID & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            strSQL = "Zl_ѪҺִ�м�¼_Delete(" & lngID & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            For i = 0 To UBound(arrSQL)
                Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            gcnOracle.CommitTrans
            blnTrans = False
            Call SetExecState(E_ɾ��ִ��)
        Case ID_���ִ��, conMenu_Manage_Complete
            If mintִ��״̬ = 1 Then
                MsgBox "��ҽ����ǰ�Ѿ�ִ����ɣ������ظ���ɡ�", vbInformation, gstrSysName
                Exit Function
            End If
            If CheckDataMoved Then Exit Function
            '������е���Ŀ�Ƿ��Ѿ����ִ�У������˺˶��Ƿ��Ѿ��˶�
            If Not CheckItemOk Then Exit Function
            '��鲡���Ƿ��������
            If Not CheckPatiIsAduit Then Exit Function
            
            '�Ƿ��������δ�շѲ��˵���Ŀ:���ܼ��ʻ���,��ΪҪִ�к����,�����ſ��ܷ��͵������շ�
            If mint��¼���� = 1 And mint�Ʒ�״̬ > 0 Then
                If Not ItemHaveCash(2, True, mlngҽ��ID, mlng��ID, mlng���ͺ�, "E", mstrNO, 1, 0, 0, mblnMoved, CDate(mstr����ʱ��), "", "", blnIsAbnormal) Then
                    If blnIsAbnormal Then
                        MsgBox "�ò��˻������쳣���ã����顣", vbInformation, gstrSysName
                    Else
                        MsgBox "�ò��˻�����δ�շѵķ��ã����顣", vbInformation, gstrSysName
                    End If
                    Exit Function
                End If
            End If
            If mint��¼���� = 2 Then
                curMoney = GetAdviceMoney(IIf(mlng��ID = 0, mlngҽ��ID, mlng��ID), mlngҽ��ID, mlng���ͺ�, str���, str�����, True, IIf(mint������� = 0, 2, 1))
                If curMoney > 0 Then
                    'סԺ��Ժ���˷��ÿ���
                    If Not PatiCanBilling(mlng����ID, mlng��ҳid, GetInsidePrivs(100, pסԺҽ������), pסԺҽ������) Then Exit Function
                    '���ʱ���
                    If InitObjPublicExpense(mlngSys) Then
                        If gobjPublicExpense.zlBillingWarn.zlBillingVerfyWarnCheck(Me, pסԺҽ������, "", mstrNO, GetInsidePrivs(mlngSys, pסԺҽ������), mlng����ID) = False Then Exit Function
                    End If
                    
                    '����һ��ͨ���������֤,ֻ���������ʷ���
                    If mint������� = 1 Then
                        If InitObjPublicExpense(mlngSys) Then
                            If Not gobjPublicExpense.zlPatiIdentify(mlngModul, Me, mlng����ID, curMoney) Then Exit Function
                        End If
                    End If
                End If
            End If
            
            If MsgBox("ȷ��Ҫ����ҽ��ִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            strSQL = "ZL_����ҽ��ִ��_Finish(" & mlngҽ��ID & "," & mlng���ͺ� & ",Null,0,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Call SetExecState(E_ִ�����)
        Case ID_ȡ�����, conMenu_Manage_Undone
            If mintִ��״̬ <> 1 Then
                MsgBox "��ҽ����ǰ��������ִ��״̬������ȡ��ִ�С�", vbInformation, gstrSysName
                Exit Function
            End If
            If CheckDataMoved Then Exit Function
            '��鲡���Ƿ��������
            If Not CheckPatiIsAduit Then Exit Function
            
            If mint��¼���� <> 1 Then
                If mint������� = 0 Then
                    byt��Դ = 2
                Else
                    byt��Դ = 1
                End If
                '���ý����ж�
                If Not ItemCanCancel(mlngҽ��ID, mlng���ͺ�, mlng��ID, "E", True, mblnMoved, byt��Դ) Then Exit Function
            End If
            
            If MsgBox("ȷʵҪ����ҽ��ȡ��ִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            strSQL = "ZL_����ҽ��ִ��_Cancel(" & mlngҽ��ID & "," & mlng���ͺ� & "," & "Null,0," & mlng����ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Call SetExecState(E_ȡ�����)
    End Select
    FuncExec = True
    Exit Function
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case ID_��ʾִ��
            mblnShowExec = Not mblnShowExec
            Call SetShowExec(mblnShowExec)
            RaiseEvent ShowExec(mblnShowExec, Me.Height)
        Case ID_���ִ��, ID_ȡ�����, conMenu_Manage_Complete, conMenu_Manage_Undone
            If FuncExec(Control.id) = True Then Call RefreshData(False)
        Case ID_ִ�м�¼, ID_ִ��ɾ��, ID_ִ��ǰ�˶�, ID_ȡ��ִ��ǰ�˶�, ID_ִ���к˶�, ID_ȡ��ִ���к˶�, conMenu_Manage_ThingAdd, conMenu_Manage_ThingDel, _
                conMenu_Manage_ThingAudit, conMenu_Manage_ThingDelAudit, conMenu_Manage_ThingAudit * 100# + 1, conMenu_Manage_ThingDelAudit * 100# + 1
            If FuncExec(Control.id) = True Then Call RefreshData
    End Select
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean
    Dim intѪҺִ��״̬ As Integer, int����״̬ As Integer, lng��ִ�п��� As Long, lngִ�п��� As Long
    Dim bln����ִ�� As Boolean
    
    blnEnable = False
    If vsExec.Row >= vsExec.FixedRows And mblnLoad = True Then
        blnEnable = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID"))) > 0
        If blnEnable = True Then
            intѪҺִ��״̬ = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ִ��״̬")))
            int����״̬ = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("����״̬")))
            lng��ִ�п��� = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("��ִ�п���ID")))
            lngִ�п��� = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ִ�п���ID")))
        End If
    End If
    
    bln����ִ�� = True
'    If intѪҺִ��״̬ > 0 Then '����Ѿ�ִ��,��ִ�п��ұ����ǵ�ǰִ�п���
'        bln����ִ�� = (lngִ�п��� = mlngҽ������ID)
'    Else
'        '����ִ��ʱ�������ǽ���ʱ��ִ�п���
'        bln����ִ�� = (lng��ִ�п��� = mlngҽ������ID) Or InStr(mstrPrivs, "ִ��������Ŀ")
'    End If
    
    Select Case Control.id
        Case ID_��ʾִ��
            Control.Checked = mblnShowExec And mint���� = 2 And Control.Visible
        Case ID_���ִ��, conMenu_Manage_Complete
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "ȷ��ִ�����") = 0)
            Control.Enabled = blnEnable And (mintִ��״̬ = 0 Or mintִ��״̬ = 3) And Control.Visible
        Case ID_ȡ�����, conMenu_Manage_Undone
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "ȡ��ִ�����") = 0)
            Control.Enabled = blnEnable And mintִ��״̬ = 1 And Control.Visible
        Case ID_ִ�м�¼, conMenu_Manage_ThingAdd
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "ִ������Ǽ�") = 0)
            Control.Enabled = mblnShowExec And blnEnable And (mintִ��״̬ = 0 Or mintִ��״̬ = 3) And Control.Visible And bln����ִ��
        Case ID_ִ��ɾ��, conMenu_Manage_ThingDel
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "ִ������Ǽ�") = 0)
            Control.Enabled = mblnShowExec And blnEnable And (mintִ��״̬ = 0 Or mintִ��״̬ = 3) And (intѪҺִ��״̬ = 1 Or intѪҺִ��״̬ = 2) And Control.Visible And bln����ִ��
        Case ID_ִ��ǰ�˶�, conMenu_Manage_ThingAudit * 100# + 1 'ִ��ǰ�˲�
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "ִ������Ǽ�") = 0)
            Control.Enabled = mblnShowExec And blnEnable And Control.Visible And intѪҺִ��״̬ = 0 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˲���")) = "" '�Ѿ����յ�Ϊ�˶�ѪҺ���ܺ˶�
            If Control.Enabled And gbln���պ����ִ�� = True Then
                Control.Enabled = (int����״̬ = 1 Or int����״̬ = 3)
            End If
        Case ID_ȡ��ִ��ǰ�˶�, conMenu_Manage_ThingDelAudit * 100# + 1 'ȡ��ִ��ǰ�˲�
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "ִ������Ǽ�") = 0)
            Control.Enabled = mblnShowExec And blnEnable And intѪҺִ��״̬ = 0 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˲���")) <> "" And Control.Visible
        Case ID_ִ���к˶�, conMenu_Manage_ThingAudit, ID_ȡ��ִ���к˶�, conMenu_Manage_ThingDelAudit  'ִ���к˶�'ȡ��ִ���к˶�
             If InStr(GetInsidePrivs(mlngSys, mlngModul), "ִ������Ǽ�") = 0 Or Val(Mid(gstrҽ���˶�, 1, 1)) = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = blnEnable And (mintִ��״̬ = 0 Or mintִ��״̬ = 3) And IIf(mint���� = 2, mblnShowExec, True)
                If (mintִ��״̬ = 0 Or mintִ��״̬ = 3) And IIf(mint���� = 2, mblnShowExec, True) Then
                    If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("�˶���")) = "" Then
                        If Control.id = ID_ִ���к˶� Or Control.id = conMenu_Manage_ThingAudit Then Control.Enabled = True
                        If Control.id = ID_ȡ��ִ���к˶� Or Control.id = conMenu_Manage_ThingDelAudit Then Control.Enabled = False
                    Else
                        If Control.id = ID_ִ���к˶� Or Control.id = conMenu_Manage_ThingAudit Then Control.Enabled = False
                        If Control.id = ID_ȡ��ִ���к˶� Or Control.id = conMenu_Manage_ThingDelAudit Then Control.Enabled = True
                    End If
                End If
            End If
        Case conMenu_Manage_ThingModi '����ִ�����
            Control.Visible = False
            Control.Enabled = Control.Visible
    End Select
End Sub

Private Sub cbsExec_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long
    
    On Error Resume Next
    Call cbsExec.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If mint���� = 1 Then lngTop = 0
    If mblnShowExec = True Then
        With vsExec
            .Left = lngLeft
            .Top = lngTop
            .Width = lngRight - lngLeft
            .Height = lngBottom - lngTop
        End With
    Else
        vsExec.Left = lngLeft
        vsExec.Top = lngTop
    End If
End Sub

Private Sub Form_Load()
    mblnShowExec = False
    mblnExecFresh = False
    mintAdviceExecState = 0
    mblnLoad = False
    Call InitTable
    mblnLoad = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf = Nothing
End Sub

Public Function zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsExec_Update(Control)
End Function

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsExec_Execute(Control)
End Sub

Private Sub vsExec_DblClick()
    Dim lngID As Long
    Dim objButton As CommandBarControl
    Dim blnReadOnly As Boolean
    
    If Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID"))) < 0 Or mblnShowExec = False Then Exit Sub
    If Not (Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ִ��״̬"))) > 0) Then Exit Sub
    If CheckDataMoved Then Exit Sub
    Call frmBloodExecEdit.ShowEdit(Me, mlngModul, mlngҽ��ID, mlng���ͺ�, mlngҽ������ID, Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID"))), mlngִ�в���ID, mstrPrivs, True)
End Sub

Private Sub vsExec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
        
    If Button = 2 And mint���� = 1 Then
        Set objPopup = cbsExec.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, ID_ִ��ǰ�˶�, "�˲�"
            .Add xtpControlButton, ID_ȡ��ִ��ǰ�˶�, "ȡ���˲�"
            .Add xtpControlButton, ID_ִ�м�¼, "��¼ִ�����"
            .Add xtpControlButton, ID_ִ��ɾ��, "ɾ��ִ�����"
            .Add xtpControlButton, ID_ִ���к˶�, "�˶�"
            .Add xtpControlButton, ID_ȡ��ִ���к˶�, "ȡ���˶�"
        End With
        
        vsExec.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Sub SetExecState(ByVal intExecState As Integer)
    If Not mfrmParent Is Nothing Then
        On Error Resume Next
        mfrmParent.timBRefresh.Enabled = True
        If Err <> 0 Then Err.Clear
    End If
    mintAdviceExecState = intExecState
End Sub
