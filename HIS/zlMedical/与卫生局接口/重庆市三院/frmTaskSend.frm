VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~4.OCX"
Begin VB.Form frmTaskSend 
   Caption         =   "���������"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11265
   Icon            =   "frmTaskSend.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4725
      ScaleHeight     =   2055
      ScaleWidth      =   2790
      TabIndex        =   1
      Top             =   855
      Width           =   2790
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1320
         Left            =   270
         TabIndex        =   2
         Top             =   390
         Width           =   3135
         _cx             =   5530
         _cy             =   2328
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
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         Begin VB.Line lnX 
            Index           =   0
            Visible         =   0   'False
            X1              =   -555
            X2              =   1230
            Y1              =   555
            Y2              =   555
         End
         Begin VB.Line lnY 
            Index           =   0
            Visible         =   0   'False
            X1              =   270
            X2              =   270
            Y1              =   420
            Y2              =   1635
         End
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3060
      Left            =   60
      TabIndex        =   0
      Top             =   1035
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   5398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�������"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�������"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����״̬"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1155
      Top             =   5190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":076A
            Key             =   "package"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":6FCC
            Key             =   "package_ok"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   7995
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":D82E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":DA4E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":DC6E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":E3E8
            Key             =   "Send"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4500
      Top             =   4650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6630
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTaskSend.frx":EB62
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14790
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmTaskSend.frx":F3F6
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   600
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTaskSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mlngLoop As Long
Private mstrKey As String
Private mfrmMain As Object
Private mvarParam As Variant
Private mstrSQL As String
Private mstrPrive As String
Private mblnShowAll As Boolean

Private Enum mCol
    ����
    �����
    ����
    ����
    �ܼ�
    ���
    ���
End Enum

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ���˵���������
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As Object
    Dim cbrToolBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "����(&P)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Task_Send, "����(&S)")
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControl.BeginGroup = True
    End With

        
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "������Ա(&A)")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        cbrControl.BeginGroup = True
        
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '�����
    With cbsThis.KeyBindings
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    

    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "����")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Task_Send, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
End Function

Private Function InitClient() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ������
    '------------------------------------------------------------------------------------------------------------------
    Dim panTab As Pane
    
    Set panTab = dkpMan.CreatePane(1, 200, 500, DockLeftOf, Nothing)
    panTab.Title = ""
    panTab.Options = PaneNoCaption
    
    Set panTab = dkpMan.CreatePane(2, 500, 200, DockRightOf, Nothing)
    panTab.Title = ""
    panTab.Options = PaneNoCaption
    
    dkpMan.SetCommandBars cbsThis
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
        
End Function

Private Function zlClearData(Optional ByVal strItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ָ���������ʾ����
    '���أ�True
    '------------------------------------------------------------------------------------------------------------------
    
    lvw.ListItems.Clear
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    vsf.RowData(1) = 0

    Call AppendRows(vsf, lnX, lnY)
    
    zlClearData = True
    
End Function

Private Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
   
    On Error GoTo errHand
    
    Select Case strMenuItem
    Case "��ȡ��쵥"
        
        Call zlClearData
        
        If ReadBill Then
            If Not (lvw.SelectedItem Is Nothing) Then
                Call zlMenuClick("��ȡ�ſ�")
            End If
        End If
       
    Case "��ȡ�ſ�"
        
        frmWait.OpenWait Me, "��ȡ���������"
        frmWait.WaitInfo = "���ڶ�ȡ�������Ӧ�������Ա"
    
        vsf.Rows = 2
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        vsf.RowData(1) = 0
        
        Call ReadBillState
        
        frmWait.CloseWait
        
    Case "���ͽ����"
        
        If ConnectAccess(strParam) Then
        
            DoEvents
            
            If SendPackage Then
                ShowSimpleMsg "��������ѱ��ɹ����ͣ�"
                lvw.SelectedItem.SubItems(2) = "�ѷ���"
            End If
            
        End If
        
        If gcnAccess.State = adStateOpen Then gcnAccess.Close
                        
    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    frmWait.CloseWait
    ShowSimpleMsg Err.Description
        
End Function

Private Function ReadBill() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����
    '����:��ȡ�ɹ�����True�����򷵻�False
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strStart As String
    Dim strEnd As String
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    gstrSQL = "Select Decode(B.����״̬,1,'package_ok','package') As Icon,A.ID,B.�������,B.�������,Decode(B.����״̬,1,'�ѷ���','δ����') As ����״̬ From ���ǼǼ�¼ A,���ǼǼ�¼_�ɱ� B Where A.ID=B.�Ǽ�id And A.���״̬>2"
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "���ʱ��", "��  ��"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "���ʱ��", "��  ��"), 2)
    If strStart = "" Then strStart = GetDateTime("��  ��", 1)
    If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
            
    gstrSQL = gstrSQL & "AND A.���ʱ�� BETWEEN TO_DATE('" & strStart & "','yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"

    Call OpenRecord(rs, gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            Set objItem = lvw.ListItems.Add(, "_" & rs("ID").Value, NVL(rs("�������")), NVL(rs("Icon")), NVL(rs("Icon")))
            objItem.SubItems(1) = NVL(rs("�������"))
            objItem.SubItems(2) = NVL(rs("����״̬"))
            rs.MoveNext
        Loop
    End If
    
    ReadBill = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description

End Function

Private Function ReadBillState() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '
    '
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim lngRow As Long
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim lngCount1 As Long
    Dim lngCount0 As Long
    Dim lngCount2 As Long
    
    On Error GoTo errHand
    
    If lvw.SelectedItem Is Nothing Then Exit Function
    
    lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
    
    gstrSQL = "SELECT A.������� AS ���,A.����id AS ID,A.����,B.�����," & _
                      "A.��챨�� AS ����," & _
                      "DECODE(C.������,NULL,NULL,TRIM(TO_CHAR(C.������,'9990.0'))||'%') AS ����," & _
                      "DECODE(A.��첡��ID, Null, 0, 1) As �ܼ�, " & _
                      "DECODE(A.���״̬, 5, 1, 0) As ��� " & _
                 "FROM �����Ա���� A," & _
                      "������Ϣ B," & _
                      "(SELECT ����id,DECODE(COUNT(*), NULL, NULL, 100 * SUM(�Ƿ��Ѽ�) / COUNT(*)) AS ������ " & _
                         "FROM (SELECT C.����id," & _
                                      "(select DECODE(S.����id, NULL, 0, 1) " & _
                                         "FROM ����ҽ����¼ M, ����ҽ������ S " & _
                                        "Where (M.ID = C.ҽ��ID Or M.���id = C.ҽ��ID) AND M.ID = S.ҽ��ID AND S.����id > 0 AND ROWNUM < 2) AS �Ƿ��Ѽ� " & _
                                 "FROM �����Ŀҽ�� C," & _
                                      "(SELECT A.ID, B.����id " & _
                                         "FROM �����Ŀ�嵥 A, �����Ա���� B " & _
                                        "WHERE A.�Ǽ�ID = B.�Ǽ�id AND A.������� = B.������� AND A.�Ǽ�ID = " & lngKey & " " & _
                                       "Union All " & _
                                         "SELECT A.ID, B.����id " & _
                                           "FROM �����Ŀ�嵥 A, �����Ա���� B " & _
                                          "WHERE A.�Ǽ�ID = B.�Ǽ�id AND A.����id = B.����id AND A.�Ǽ�ID = " & lngKey & " " & _
                                       ") D " & _
                                "WHERE C.�嵥ID = D.ID AND C.����ID = D.����id) " & _
                        "GROUP BY ����id) C " & _
                "WHERE A.����ID = B.����ID(+) AND A.����ID = C.����id(+) AND A.�Ǽ�ID = " & lngKey
                
    If mblnShowAll = False Then
        gstrSQL = gstrSQL & " And A.��챨��=1 "
    End If
    gstrSQL = gstrSQL & " ORDER BY A.�������,B.����� "
    
    Call OpenRecord(rs, gstrSQL, Me.Caption)
    If rs.BOF = False Then

        Call LoadGrid(vsf, rs)
        
        Call AppendRows(vsf, lnX, lnY)
        
        'ͳ������������δ��������δ��������
        For lngLoop = 1 To vsf.Rows - 1
            
            'δ����ͳ��
            If Abs(Val(vsf.TextMatrix(lngLoop, mCol.����))) <> 1 Then
                lngCount0 = lngCount0 + 1
            Else
                '����ͳ��
                If Abs(Val(vsf.TextMatrix(lngLoop, mCol.���))) = 1 Then
                    lngCount1 = lngCount1 + 1
                Else
                    'δ��ͳ��
                    lngCount2 = lngCount2 + 1
                End If
            End If
        Next
        
        stbThis.Panels(2).Text = "Ӧ��:" & lngCount0 + lngCount1 + lngCount2 & "��;ʵ��:" & lngCount1 + lngCount2 & "��(���:" & lngCount1 & "��;δ��:" & lngCount2 & "��);δ��:" & lngCount0 & "��"
                    
    End If
    
    
    ReadBillState = True
    
    Exit Function
    
errHand:
    
    ShowSimpleMsg Err.Description
End Function

Private Function SendPackage() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���ܳɹ�����True�����򷵻�False
    '------------------------------------------------------------------------------------------------------------------
    Dim rsSQL As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lngLoop As Long
    Dim blnTran As Boolean
    
    Dim str���� As String
    Dim strSvr��Ͽ��� As String
    Dim strSvr��ϱ��� As String
    Dim strSvr������� As String
    Dim strSvr���ҽ��  As String
    Dim strSvr������� As String
    Dim str����С�� As String
    Dim str������Ŀ���� As String
    Dim str������ĿС�� As String
    
    On Error GoTo errHand
    
    
    If lvw.SelectedItem.SubItems(2) = "�ѷ���" Then
        If MsgBox("�������������Ѿ����ͣ��Ƿ���Ҫ���·��ͣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
        
    blnTran = True
    gcnAccess.BeginTrans

    frmWait.OpenWait frmMain, "���ͽ����"
    frmWait.WaitInfo = "����ɾ��ԭ������..."
    
    gstrSQL = "Delete From hdatadeptest_�ֿ���Ŀ���"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hcheckmemb_�Ѽ���Ա"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatadep_�ֿ�С��"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatadepunion_�����Ͻ��"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatadepdiag_�ֿ���Ͻ��"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatadiag_������Ͻ��"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatarep_���챨��"
    gcnAccess.Execute gstrSQL
    
    frmWait.ShowProgress = True
    
    For mlngLoop = 1 To vsf.Rows - 1
        
        frmWait.WaitInfo = "���ڷ��������..."
        frmWait.WaitProgress = Format(100 * mlngLoop / (vsf.Rows - 1), "0.00")
            
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.����))) = 1 Then
                        
            mstrSQL = GetPublicSQL(SQL.��Ա��������)
                                                    
            Set rs = OpenSQLRecord(mstrSQL, Me.Caption, lvw.SelectedItem.Text, Val(vsf.RowData(mlngLoop)))
'            Call OpenRecord(rs, mstrSQL, Me.Caption)
            If rs.BOF = False Then
                
                str���� = NVL(rs("�������")) & NVL(rs("��Ա���"))
                
                '1.�ϴ��Ѽ���Ա��hcheckmemb_�Ѽ���Ա------------------------------------------------------------------
                
                gstrSQL = "Delete From hcheckmemb_�Ѽ���Ա Where checkcode='" & str���� & "'"
                gcnAccess.Execute gstrSQL
                
                gstrSQL = "Insert Into hcheckmemb_�Ѽ���Ա(checkcode,ifdel,taskcode,taskseq,seq,ifprinted,checkstatus,iffinished,b0110,b0105,b0160,membcode,membtype,a0101,a0107,age,a6405,a0704,checkdate,asmcode,asmseq,asmname,ifasmdep,asmdepstr,checkfee,ifplus,checkfeeplus,remark,tele,email,accesscode,sendwayid,ifsend,"
                gstrSQL = gstrSQL & "fee01,fee02,fee03,fee04,fee05,fee06,fee07,fee08,fee09,fee10,fee11,fee12,fee13,fee14,fee15,fee16,fee17,fee18,fee19,fee20,"
                gstrSQL = gstrSQL & "feesum,ifcard,workunit,undofee,discountfee,pis_01,bseq,ifad,adclass,tasktype) Values ("
                
                gstrSQL = gstrSQL & "'" & str���� & "','0','" & NVL(rs("�������")) & "','" & NVL(rs("��Ա���")) & "','" & NVL(rs("��Ա���")) & "',"
                gstrSQL = gstrSQL & "'0','0','1',"
                gstrSQL = gstrSQL & "'" & NVL(rs("��λ����")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("��λ����")) & "',"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "'" & NVL(rs("������")) & "',"
                gstrSQL = gstrSQL & "'01',"
                gstrSQL = gstrSQL & "'" & NVL(rs("����")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("�Ա�")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("����")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("��ְ���")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("��ְ����")) & "',"
                gstrSQL = gstrSQL & "'" & Format(NVL(rs("���ʱ��")), "yyyy-MM-dd") & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("�ײͱ���")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("�ײ����")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("�ײ�����")) & "',"
                gstrSQL = gstrSQL & "'0',NULL,'0','0','0',NULL,NULL,NULL,NULL,NULL,'0',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,"
                gstrSQL = gstrSQL & "'0',NULL,NULL,'0','0',NULL,NULL,'0','0',NULL"
                gstrSQL = gstrSQL & ")"
                gcnAccess.Execute gstrSQL
                
                                    
                '2.�ϴ��ֿ���Ŀ�����hdatadeptest_�ֿ���Ŀ���------------------------------------------------------------------
                '��ҽԺ���ӵ���Ŀ���ϴ����ϴ��Ķ����������ָ���˵���Ŀ,���û�ж���Ļ���Ҳ���ϴ�
                
                mstrSQL = GetPublicSQL(SQL.�ֿ���Ŀ���)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("�Ǽ�id"))), Val(vsf.RowData(mlngLoop)), lvw.SelectedItem.Text)
                If rsTmp.BOF = False Then
                    Do While Not rsTmp.EOF
                        
                        gstrSQL = "Delete From hdatadeptest_�ֿ���Ŀ��� " & _
                                    "Where checkcode='" & str���� & "' and deptcode='" & NVL(rsTmp("��Ͽ���")) & "' and testcode='" & NVL(rsTmp("��Ŀ����")) & "'"
                        gcnAccess.Execute gstrSQL
                        
                        gstrSQL = "Insert Into hdatadeptest_�ֿ���Ŀ���(gb2260,checkcode,deptcode,testcode,wayid,taskcode,membcode,unioncode,htestcode,testname,testresult,teststatus,testsign,testunit,testrange,testlower,testhigher,warncode,remark) values ("
                        
                        gstrSQL = gstrSQL & "'5000',"                                                                   'gb2260,Ĭ��5000
                        gstrSQL = gstrSQL & "'" & str���� & "',"                                                      'checkcode,����
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��Ͽ���")) & "',"                                            'deptcode,�����ұ���
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��Ŀ����")) & "',"                                            'testcode,��Ŀ����
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��Ŀ����")) & "',"                                            'wayid,����
                        gstrSQL = gstrSQL & "'" & NVL(rs("�������")) & "',"                                            'taskcode,�������
                        gstrSQL = gstrSQL & "'" & NVL(rs("������")) & "',"                                              'membcode,������
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��ϱ���")) & "',"                                            'unioncode,��ϱ���
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��Ŀ��֧")) & "',"                                            'htestcode,��Ŀ���д���
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��Ŀ����")) & "',"                                            'testname,��Ŀ����
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("���")) & "',"                                             'testresult,��Ŀ���
                        gstrSQL = gstrSQL & "NULL,"                                                                     'teststatus,��Ŀ״̬ null��ƫ�� ƫ��
                        gstrSQL = gstrSQL & "'" & Left(NVL(rsTmp("��־"), "0"), 1) & "',"                                           'testsign,0������ 1��ƫ�� 2��ƫ�� 3������ 4������ 5������ 9���쳣
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��λ")) & "',"                                             'testunit,��Ŀ��λ
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("�ο�")) & "',"                                             'testrange,��Χ
                        gstrSQL = gstrSQL & "NULL,"                                                                     'testlower,Ĭ��null
                        gstrSQL = gstrSQL & "NULL,"                                                                     'testhigher,Ĭ��null
                        gstrSQL = gstrSQL & "NULL,"                                                                     'warncode,Ĭ��null
                        gstrSQL = gstrSQL & "NULL"                                                                     'remark,Ĭ��null
                        gstrSQL = gstrSQL & ")"
                        
                        gcnAccess.Execute gstrSQL
                        
                        rsTmp.MoveNext
                    Loop
                End If
                
                '�ϴ�С����������
                mstrSQL = GetPublicSQL(SQL.�ֿ���Ŀ����)
                'Call OpenRecord(rsTmp, mstrSQL, Me.Caption)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("�Ǽ�id"))), Val(vsf.RowData(mlngLoop)), lvw.SelectedItem.Text)
                If rsTmp.BOF = False Then
                    Do While Not rsTmp.EOF
                        
                        strSvr��Ͽ��� = NVL(rsTmp("��Ͽ���"))
                        strSvr��ϱ��� = NVL(rsTmp("��ϱ���"))
                        strSvr������� = NVL(rsTmp("�������"))
                        strSvr���ҽ�� = NVL(rsTmp("��д��"))
                        strSvr������� = Format(NVL(rsTmp("��������")), "yyyy-MM-dd")
                        
                        str����С�� = str����С�� & NVL(rsTmp("��������")) & vbCrLf
                        str������Ŀ���� = str������Ŀ���� & ";" & NVL(rsTmp("��ϱ���"))

                        str������ĿС�� = str������ĿС�� & NVL(rsTmp("��������")) & vbCrLf
                              
                        rsTmp.MoveNext
                        
                        If rsTmp.EOF Then
                            GoTo ����С��
                        Else
                            If strSvr��Ͽ��� <> NVL(rsTmp("��Ͽ���")) Then
                                GoTo ����С��
                            ElseIf strSvr��ϱ��� <> NVL(rsTmp("��ϱ���")) Then
                                GoTo ������ĿС��
                            End If
                        End If
                        
                        GoTo OverPoint
����С��:
                        If str����С�� <> "" Then
                                
                            If str������Ŀ���� <> "" Then str������Ŀ���� = Mid(str������Ŀ����, 2)
                                
                            '3.�ϴ����ҷֿ�С�ᣬhdatadep_�ֿ�С��------------------------------------------------------------------
                            
                            gstrSQL = "Delete From hdatadep_�ֿ�С�� " & _
                                    "Where checkcode='" & str���� & "' and deptcode='" & strSvr��Ͽ��� & "'"
                            gcnAccess.Execute gstrSQL
                        
                            gstrSQL = "Insert Into hdatadep_�ֿ�С��(gb2260,checkcode,deptcode,unioncode,taskcode,seq,membcode,membtype,initday,checkstatus,sampleno,depresult,checkdate,checkdoc,reviewdoc,iffinished,iflock,ifdata,unionfee,ifplus,checklevel,tag,remark,depsignstr,depdiagstr,depopsstr,oper,ifad,adclass,deptseq) values ("
                            
                            gstrSQL = gstrSQL & "'5000',"                                                                           'gb2260,Ĭ��5000
                            gstrSQL = gstrSQL & "'" & str���� & "',"                                                              'checkcode,����
                            gstrSQL = gstrSQL & "'" & strSvr��Ͽ��� & "',"                                                         'deptcode,�����ұ���
                            gstrSQL = gstrSQL & "'" & str������Ŀ���� & "',"                                                        'unioncode,�����ϱ���
                            gstrSQL = gstrSQL & "'" & NVL(rs("�������")) & "',"                                                    'taskcode,�������
                            gstrSQL = gstrSQL & "'" & NVL(rs("��Ա���")) & "',"                                            'seq,��������
                            gstrSQL = gstrSQL & "'" & NVL(rs("������")) & "',"                                              'membcode,������
                            gstrSQL = gstrSQL & "'01',"                                                                     'membtype,Ĭ��01
                            gstrSQL = gstrSQL & "'" & Format(NVL(rs("���ʱ��")), "yyyy-MM-dd") & "',"                      'initday,�������
                            gstrSQL = gstrSQL & "'5',"                                                                      'checkstatus,Ĭ��5
                            gstrSQL = gstrSQL & "NULL,"                                                                     'sampleno,Ĭ��null
                            gstrSQL = gstrSQL & "'" & str����С�� & "',"                                                  'depresult,�ֿ�С��
                            gstrSQL = gstrSQL & "NULL,"                                                                     'checkdate,����������
                            gstrSQL = gstrSQL & "NULL,"                                                                     'checkdoc,���ҽ��
                            gstrSQL = gstrSQL & "NULL,"                                                                     'reviewdoc,����ҽ��,Ĭ��null
                            gstrSQL = gstrSQL & "'1',"                                                                      'iffinished,Ĭ��1
                            gstrSQL = gstrSQL & "'0',"                                                                      'iflock,Ĭ��0
                            gstrSQL = gstrSQL & "'0',"                                                                      'ifdata,Ĭ��0
                            gstrSQL = gstrSQL & "NULL,"                                                                     'unionfee,Ĭ��null
                            gstrSQL = gstrSQL & "'0',"                                                                      'ifplus,Ĭ��0
                            gstrSQL = gstrSQL & "'0',"                                                                      'checklevel,Ĭ��0
                            gstrSQL = gstrSQL & "NULL,"                                                                     'tag,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'remark,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'depsignstr,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'depdiagstr,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'depopsstr,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'oper,Ĭ��null
                            gstrSQL = gstrSQL & "'0',"                                                                      'ifad,Ĭ��0
                            gstrSQL = gstrSQL & "'DEP',"                                                                    'adclass,Ĭ��DEP
                            gstrSQL = gstrSQL & "'0'"                                                                      'deptseqĬ��0
                            
                            gstrSQL = gstrSQL & ")"
                            gcnAccess.Execute gstrSQL
                        End If
                        
                        str����С�� = ""
                        str������Ŀ���� = ""
                                                    
������ĿС��:
                        If str������ĿС�� <> "" Then
                                
                            '�ϴ������Ͻ��
                            
                            gstrSQL = "Delete From hdatadepunion_�����Ͻ�� " & _
                                    "Where checkcode='" & str���� & "' and deptcode='" & strSvr��Ͽ��� & "' and unioncode='" & strSvr��ϱ��� & "'"
                            gcnAccess.Execute gstrSQL
                            
                            gstrSQL = "Insert Into hdatadepunion_�����Ͻ��(gb2260,checkcode,deptcode,unioncode,taskcode,seq,membcode,membtype,initday,checkstatus,sampleno,depresult,checkdate,checkdoc,reviewdoc,iffinished,iflock,ifdata,testsignstr,unionfee,ifplus,tag,ifsettle,rackno,rackoper,racktime,uname,deptseq,rackbatch,uniondesc,regstatus,checklevel,settlecode) values ("
                                'work
                            
                            gstrSQL = gstrSQL & "'5000',"                                                                  'gb2260,Ĭ��5000
                            gstrSQL = gstrSQL & "'" & str���� & "',"                                                      'checkcode,����
                            gstrSQL = gstrSQL & "'" & strSvr��Ͽ��� & "',"                                             'deptcode,�����ұ���
                            gstrSQL = gstrSQL & "'" & strSvr��ϱ��� & "',"                                     'unioncode,��ϱ���
                            gstrSQL = gstrSQL & "'" & NVL(rs("�������")) & "',"                                        'taskcode,�������
                            gstrSQL = gstrSQL & "'" & NVL(rs("��Ա���")) & "',"                                        'seq,��������
                            gstrSQL = gstrSQL & "'" & NVL(rs("������")) & "',"                                          'membcode,������
                            gstrSQL = gstrSQL & "'01',"                                                                  'membtype,Ĭ��01
                            gstrSQL = gstrSQL & "'" & Format(NVL(rs("���ʱ��")), "yyyy-MM-dd") & "',"                  'initday,�������,����/ʱ��
                            gstrSQL = gstrSQL & "'0',"                                                                  'checkstatus,Ĭ��0
                            gstrSQL = gstrSQL & "NULL,"                                                                 'sampleno,Ĭ��null
                            gstrSQL = gstrSQL & "'" & str������ĿС�� & "',"                                                 'depresult,��Ͻ��
                            gstrSQL = gstrSQL & "'" & strSvr������� & "',"                                         'checkdate,�������,����/ʱ��
                            gstrSQL = gstrSQL & "'" & strSvr���ҽ�� & "',"                                        'checkdoc,���ҽ��
                            gstrSQL = gstrSQL & "NULL,"                                                                 'reviewdoc,����ҽ������ѡ
                            gstrSQL = gstrSQL & "'1',"                                                                  'iffinished,Ĭ��1
                            gstrSQL = gstrSQL & "'0',"                                                                  'iflock,Ĭ��0
                            gstrSQL = gstrSQL & "'0',"                                                                  'ifdata,Ĭ��0
                            gstrSQL = gstrSQL & "'0',"                                                                  'testsignstr,Ĭ��0
                            gstrSQL = gstrSQL & "'0',"                                                                  'unionfee,��Ϸ���
                            gstrSQL = gstrSQL & "0,"                                                                  'ifplus,�Ƿ���� 0���� 1����
                            gstrSQL = gstrSQL & "NULL,"                                                                 'tag,Ĭ��null
                            gstrSQL = gstrSQL & "'0',"                                                                  'ifsettle,Ĭ��0
                            gstrSQL = gstrSQL & "NULL,"                                                                 'rackno,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                 'rackoper,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                 'racktime,Ĭ��null
                            gstrSQL = gstrSQL & "'" & strSvr������� & "',"                                             'uname,�������
                            gstrSQL = gstrSQL & "0,"                                                                  'deptseq,Ĭ��0
                            'gstrSQL = gstrSQL & "NULL,"                                                                 'work,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                 'rackbatch,Ĭ��null
                            gstrSQL = gstrSQL & "NULL,"                                                                 'uniondesc,Ĭ��null
                            gstrSQL = gstrSQL & "0,"                                                                  'regstatus,Ĭ��0
                            gstrSQL = gstrSQL & "'0',"                                                                  'checklevel,Ĭ��0
                            gstrSQL = gstrSQL & "NULL"                                                                 'settlecode,Ĭ�Ͽ��ַ���
                            gstrSQL = gstrSQL & ")"
                            gcnAccess.Execute gstrSQL
                        End If
                        
                        str������ĿС�� = ""
                                                    
OverPoint:
                    Loop
                End If
                
                '�ϴ��������
                mstrSQL = GetPublicSQL(SQL.�ֿ���Ŀ���)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("�Ǽ�id"))), Val(vsf.RowData(mlngLoop)), lvw.SelectedItem.Text)
                If rsTmp.BOF = False Then
                    Do While Not rsTmp.EOF
                        
                        '4.�ϴ��ֿ���Ͻ����hdatadepdiag_�ֿ���Ͻ��------------------------------------------------------------------
                        
                        gstrSQL = "Delete From hdatadepdiag_�ֿ���Ͻ�� " & _
                                "Where checkcode='" & str���� & "' and deptcode='" & NVL(rsTmp("��Ͽ���")) & "' and signcode='" & NVL(rsTmp("��ϱ���")) & "'"
                        gcnAccess.Execute gstrSQL
                        
                        gstrSQL = "Insert Into hdatadepdiag_�ֿ���Ͻ��(gb2260,checkcode,deptcode,unioncode,taskcode,membcode,ifdel,iftag,ifnewsign,ifwhy,signtype,signcode,signclass,signstatus,diagdeptcode,seq,htestcode,diagwhere,diagdegree,diagclass,stdcode,diagcode,mcode,diagname,diagviewform,checkdate,checkdoc,testinfoch,diaginfoch,remark,tag,ifguide,ifimpt,ifdoubt) values ("
                        
                        gstrSQL = gstrSQL & "'5000',"                                                           'gb2260,Ĭ��5000
                        gstrSQL = gstrSQL & "'" & str���� & "',"                                              'checkcode,����
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��Ͽ���")) & "',"                                 'deptcode,�����ұ���
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��ϱ���")) & "',"                                 'unioncode,��ϱ���
                        gstrSQL = gstrSQL & "'" & NVL(rs("�������")) & "',"                                    'taskcode,�������
                        gstrSQL = gstrSQL & "'" & NVL(rs("������")) & "',"                                      'membcode,������
                        gstrSQL = gstrSQL & "'0',"                                                              'ifdel,Ĭ��0
                        gstrSQL = gstrSQL & "'0',"                                                              'iftag,?
                        gstrSQL = gstrSQL & "'0',"                                                              'ifnewsign ,�Ƿ��·������,Ĭ��0
                        gstrSQL = gstrSQL & "'0',"                                                              'ifwhy,�Ƿ�������,Ĭ��0
                        gstrSQL = gstrSQL & "'1',"                                                              'signtype,������� 0:������� 1:�������
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��ϱ���")) & "',"                                 'signcode,��ϱ���
                        gstrSQL = gstrSQL & "NULL,"                                                             'signclass,Ĭ��null
                        gstrSQL = gstrSQL & "NULL,"                                                             'signstatus,Ĭ��null
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagdeptcode ,��Ͽ��ұ���
                        gstrSQL = gstrSQL & "NULL,"                                                             'seq,Ĭ��null
                        gstrSQL = gstrSQL & "'',"                                 'htestcode,��Ŀ�������
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagwhere,��Ϸ�λ��Ĭ��null
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagdegree,��ϳ̶ȣ�Ĭ��null
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagclass,��ϼ���Ĭ��null
                        gstrSQL = gstrSQL & "'ICD-10',"                                                         'stdcode,��׼���뼯����
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��������")) & "',"                                 'diagcode,��׼����
                        gstrSQL = gstrSQL & "NULL,"                                                             'mcode,Ĭ��null
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("�������")) & "',"                                 'diagname,�������
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagviewform,Ĭ��null
                        gstrSQL = gstrSQL & "'" & Format(NVL(rs("���ʱ��")), "yyyy-MM-dd") & "',"              'checkdate,�������
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��д��")) & "',"                                   'checkdoc,���ҽ��
                        gstrSQL = gstrSQL & "NULL,"                                                             'testinfoch,Ĭ��null
                        gstrSQL = gstrSQL & "NULL,"                                                             'diaginfoch,Ĭ��null
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("��Ͻ���")) & "',"                                 'remark,��Ͻ���
                        gstrSQL = gstrSQL & "NULL,"                                                             'Tag,Ĭ��null
                        gstrSQL = gstrSQL & "'0',"                                                              'ifguide,Ĭ��0
                        gstrSQL = gstrSQL & "'0',"                                                              'ifimpt,Ĭ��0
                        gstrSQL = gstrSQL & "'0'"                                                               'ifdoubt,Ĭ��0
                        gstrSQL = gstrSQL & ")"
                        gcnAccess.Execute gstrSQL
                        
                        rsTmp.MoveNext
                    Loop
                End If
                                    
                
                '7.�ϴ����챨�棬hdatarep_���챨��------------------------------------------------------------------
                mstrSQL = GetPublicSQL(SQL.�ܼ챨�潨��)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("�Ǽ�id"))), Val(vsf.RowData(mlngLoop)))
                If rsTmp.BOF = False Then
                    
                    gstrSQL = "Delete From hdatarep_���챨�� Where checkcode='" & str���� & "'"
                    gcnAccess.Execute gstrSQL
                    
                    gstrSQL = "Insert Into hdatarep_���챨��(gb2260,checkcode,taskcode,seq,membcode,membtype,initday,checkstatus,iffinished,iflock,hresult,hresultother,hadvice,checkdoc,reviewdoc,checkdate,remark,workunit,iftrace,ifprint,ifad,adclass) values ("

                    gstrSQL = gstrSQL & "'5000',"                                                              'gb2260,Ĭ��5000
                    gstrSQL = gstrSQL & "'" & str���� & "',"                                                  'checkcode,����
                    gstrSQL = gstrSQL & "'" & NVL(rs("�������")) & "',"                                        'taskcode,�������
                    gstrSQL = gstrSQL & "'" & NVL(rs("��Ա���")) & "',"                                        'seq,��������
                    gstrSQL = gstrSQL & "'" & NVL(rs("������")) & "',"                                          'membcode,������
                    gstrSQL = gstrSQL & "'01',"                                                                 'membtype,Ĭ��01
                    gstrSQL = gstrSQL & "'" & Format(NVL(rs("���ʱ��")), "yyyy-MM-dd") & "',"                  'initday,�������
                    gstrSQL = gstrSQL & "'3',"                                                              'checkstatus,Ĭ��3
                    gstrSQL = gstrSQL & "'1',"                                                              'iffinished,Ĭ��1
                    gstrSQL = gstrSQL & "'0',"                                                              'iflock,Ĭ��0
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("����ͷ")) & "',"                                      'hresult,��챨��ͷ
                    gstrSQL = gstrSQL & "NULL,"                                                             'hresultother,Ĭ��null
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("����ָ��")) & "',"                                      'hadvice,�ۺϽ���ָ��
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("��д��")) & "',"                               'checkdoc,����ҽ��
                    gstrSQL = gstrSQL & "NULL,"                                                         'reviewdoc,����ҽ�� ��ѡ
                    gstrSQL = gstrSQL & "'" & Format(NVL(rsTmp("��д����")), "yyyy-MM-dd") & "',"              'checkdate,��������
                    gstrSQL = gstrSQL & "NULL,"                                                             'remark,Ĭ��null
                    gstrSQL = gstrSQL & "NULL,"                                                             'workunit,Ĭ��null
                    gstrSQL = gstrSQL & "'0',"                                                              'iftrace,Ĭ��0
                    gstrSQL = gstrSQL & "'0',"                                                              'ifprint,Ĭ��0
                    gstrSQL = gstrSQL & "'0',"                                                              'ifad,Ĭ��0
                    gstrSQL = gstrSQL & "'REP'"                                                              'adclass,Ĭ��REP

                    gstrSQL = gstrSQL & ")"
                    gcnAccess.Execute gstrSQL
                    
                End If
                
                
                '8.�ϴ�������Ͻ����hdatadiag_������Ͻ��------------------------------------------------------------------
                mstrSQL = GetPublicSQL(SQL.������Ͻ��)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("�Ǽ�id"))), Val(vsf.RowData(mlngLoop)), lvw.SelectedItem.Text)
                If rsTmp.BOF = False Then
                    
                    'gb2260,checkcode,deptcode,unioncode,taskcode,membcode,ifdel,iftag,ifnewsign,ifwhy,signtype,signcode,signclass,signstatus,diagdeptcode,seq,htestcode,diagwhere,diagdegree,diagclass,stdcode,diagcode,mcode,diagname,diagviewform,checkdate,checkdoc,testinfoch,diaginfoch,remark,tag,ifguide,ifimpt,ifdoubt
                    gstrSQL = "Delete From hdatadiag_������Ͻ�� Where checkcode='" & str���� & "' and signcode='" & NVL(rsTmp("��ϱ���")) & "'"
                    gcnAccess.Execute gstrSQL
                    
                    gstrSQL = "Insert Into hdatadiag_������Ͻ��(gb2260,checkcode,deptcode,unioncode,taskcode,membcode,ifdel,iftag,ifnewsign,ifwhy,signtype,signcode,signclass,signstatus,diagdeptcode,seq,htestcode,diagwhere,diagdegree,diagclass,stdcode,diagcode,mcode,diagname,diagviewform,checkdate,checkdoc,testinfoch,diaginfoch,remark,tag,ifguide,ifimpt,ifdoubt) values ("

                    gstrSQL = gstrSQL & "'5000',"                                                           'gb2260,Ĭ��5000
                    gstrSQL = gstrSQL & "'" & str���� & "',"                                              'checkcode,����
                    gstrSQL = gstrSQL & "'',"                                 'deptcode,�����ұ���
                    gstrSQL = gstrSQL & "'',"                                 'unioncode,��ϱ���
                    gstrSQL = gstrSQL & "'" & NVL(rs("�������")) & "',"                                    'taskcode,�������
                    gstrSQL = gstrSQL & "'" & NVL(rs("������")) & "',"                                      'membcode,������
                    gstrSQL = gstrSQL & "'0',"                                                              'ifdel,Ĭ��0
                    gstrSQL = gstrSQL & "'0',"                                                              'iftag,?
                    gstrSQL = gstrSQL & "'0',"                                                              'ifnewsign ,�Ƿ��·������,Ĭ��0
                    gstrSQL = gstrSQL & "'0',"                                                              'ifwhy,�Ƿ�������,Ĭ��0
                    gstrSQL = gstrSQL & "'1',"                                                              'signtype,������� 0:������� 1:�������
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("��ϱ���")) & "',"                                 'signcode,��ϱ���
                    gstrSQL = gstrSQL & "NULL,"                                                             'signclass,Ĭ��null
                    gstrSQL = gstrSQL & "NULL,"                                                             'signstatus,Ĭ��null
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagdeptcode ,��Ͽ��ұ���
                    gstrSQL = gstrSQL & "NULL,"                                                             'seq,Ĭ��null
                    gstrSQL = gstrSQL & "'',"                                 'htestcode,��Ŀ�������
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagwhere,��Ϸ�λ��Ĭ��null
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagdegree,��ϳ̶ȣ�Ĭ��null
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagclass,��ϼ���Ĭ��null
                    gstrSQL = gstrSQL & "'ICD-10',"                                                         'stdcode,��׼���뼯����
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("��������")) & "',"                                 'diagcode,��׼����
                    gstrSQL = gstrSQL & "NULL,"                                                             'mcode,Ĭ��null
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("�������")) & "',"                                 'diagname,�������
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagviewform,Ĭ��null
                    gstrSQL = gstrSQL & "'" & Format(NVL(rs("���ʱ��")), "yyyy-MM-dd") & "',"              'checkdate,�������
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("��д��")) & "',"                                   'checkdoc,���ҽ��
                    gstrSQL = gstrSQL & "NULL,"                                                             'testinfoch,Ĭ��null
                    gstrSQL = gstrSQL & "NULL,"                                                             'diaginfoch,Ĭ��null
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("��Ͻ���")) & "',"                                 'remark,��Ͻ���
                    gstrSQL = gstrSQL & "NULL,"                                                             'Tag,Ĭ��null
                    gstrSQL = gstrSQL & "'0',"                                                              'ifguide,Ĭ��0
                    gstrSQL = gstrSQL & "'0',"                                                              'ifimpt,Ĭ��0
                    gstrSQL = gstrSQL & "'0'"                                                               'ifdoubt,Ĭ��0
                    gstrSQL = gstrSQL & ")"

                    gcnAccess.Execute gstrSQL
                    
                End If
                
                gstrSQL = "Update ���ǼǼ�¼_�ɱ� Set ����״̬=1 Where �������='" & NVL(rs("�������")) & "'"
                gcnOracle.Execute gstrSQL
            
            End If
        End If
    Next
    
    gcnAccess.CommitTrans
    frmWait.CloseWait
    blnTran = False
        
    SendPackage = True
    
    Exit Function
    
errHand:
    Dim strError As String
    
    strError = Err.Description
    If blnTran Then gcnAccess.RollbackTrans
    
    frmWait.CloseWait
    ShowSimpleMsg strError
    
'    Resume
End Function

Private Function InitData() As Boolean
    
    Dim strVsf As String
    
    strVsf = "����,900,1,1,1,;�����,1080,7,1,1,;����,600,4,1,1,;����,810,7,1,1,;�ܼ�,600,4,1,1,;���,600,4,1,1,;���,2100,1,1,1,"
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(mCol.����) = flexDTBoolean
    vsf.ColDataType(mCol.�ܼ�) = flexDTBoolean
    vsf.ColDataType(mCol.���) = flexDTBoolean
    Call AppendRows(vsf, lnX, lnY)
    
    mblnShowAll = False
    
    InitData = True
    
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As Object

    On Error GoTo errHand
    
    Select Case Control.ID
                
        Case conMenu_File_Parameter
            
            If frmTaskSendFilter.ShowFilter(Me) Then
                Call zlMenuClick("��ȡ��쵥")
                If Not (lvw.SelectedItem Is Nothing) Then Call zlMenuClick("��ȡ�ſ�")
            End If
            
        Case conMenu_Task_Send
            
            If MsgBox("ȷ������Ҫ���ͽ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
   
            dlg.Flags = &H4 Or &H200000 Or &H800 & &H1000
            dlg.Filter = "�������|�����.mdb"
            dlg.FilterIndex = 0
            
            dlg.DialogTitle = "���������"
            dlg.FileName = ""
            dlg.ShowOpen
            If dlg.FileName <> "" Then Call zlMenuClick("���ͽ����", dlg.FileName)
            
            
        Case conMenu_View_ToolBar_Button
        
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_Expend_CurExpend
            
            mblnShowAll = Not mblnShowAll
            If Not (lvw.SelectedItem Is Nothing) Then Call zlMenuClick("��ȡ�ſ�")
            
        Case conMenu_View_Refresh
            
            Call zlMenuClick("��ȡ��쵥")
            If Not (lvw.SelectedItem Is Nothing) Then Call zlMenuClick("��ȡ�ſ�")
                        
        Case conMenu_Help_Help
        
            Call ShowHelp(Me.hWnd, Me.Name)
        
        Case conMenu_Help_About
            
            frmAbout.Show 1, Me
            
        Case conMenu_File_Exit
        
            Unload Me
            Exit Sub
            
    End Select
    
    
    cbsThis.RecalcLayout
    
    Exit Sub
    
errHand:
    
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub


Private Sub cbsThis_Resize()
    
    Call AppendRows(vsf, lnX, lnY)
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case conMenu_Task_Send
            
        Control.Visible = (InStr(mstrPrive, ";���ͽ��;") > 0)
        Control.Enabled = (lvw.ListItems.Count > 0)
        
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Expend_CurExpend
    
        Control.Checked = mblnShowAll
        
    End Select
    
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    
    Select Case Item.ID
    Case 1
        
        
        Item.Handle = lvw.hWnd
        
    Case 2
        
       Item.Handle = picContainer.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
        
    If InitData = False Then
        Unload Me
        Exit Sub
    End If
    
    DoEvents
    mblnStartUp = False
    
    Call zlMenuClick("��ȡ��쵥")
    If Not (lvw.SelectedItem Is Nothing) Then Call zlMenuClick("��ȡ�ſ�")
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call InitMenuBar
    Call InitClient
    
    Call RestoreWinState(Me, App.ProductName)
    mstrPrive = gstrPrive
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey <> Item.Key Then
        
        mstrKey = Item.Key
        
        Call zlMenuClick("��ȡ�ſ�")
        
    End If
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    
    vsf.Left = 0
    vsf.Top = 0
    vsf.Width = picContainer.Width
    vsf.Height = picContainer.Height
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

