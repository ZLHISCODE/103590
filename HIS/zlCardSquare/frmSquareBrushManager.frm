VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareBrushManager 
   Caption         =   "���㿨ˢ������"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frmSquareBrushManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11745
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8025
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSquareBrushManager.frx":74F2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
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
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2715
      Left            =   105
      TabIndex        =   0
      Top             =   825
      Width           =   9885
      _cx             =   17436
      _cy             =   4789
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   350
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSquareBrushManager.frx":7D86
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   120
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
      ExplorerBar     =   7
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmSquareBrushManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mlngModule As Long, mstrPrivs As String, mintSucces As Integer
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private Const mconMenu_Edit_Affirm = 225
Private mrsBrushData As ADODB.Recordset
Private mrsFeeList As ADODB.Recordset
Private mdbl������Ѷ� As Double
Private WithEvents mobjBrushCard As clsBrushSequareCard
Attribute mobjBrushCard.VB_VarHelpID = -1
Private mbytCall As Byte  '�������� 0-  ������õ��� 1-  סԺ���ʵ���,3-����
Private mstrTitle As String '���ڴ�����Ի�����Ĵ�����
Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĹ�����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-12-31 10:45:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    '����Ƿ���������ص�ˢ������
    Set mobjBrushCard = New clsBrushSequareCard
    CheckDepend = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlShowBrushCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal rsFeeList As ADODB.Recordset, dbl������Ѷ� As Double, rsRequare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ���ӿ�
    '���:frmMain-���õ�������
    '     lngModule-���õ�ģ���
    '     strPrivs-���õ�Ȩ�޴�
    '     dbl������Ѷ�-����ˢ�������ˢ����
    '     rsFeeList-������ϸ��Ϣ()
    '����:rsRequare-���ؽ�����Ϣ
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 10:33:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mdbl������Ѷ� = dbl������Ѷ�
    Set mrsFeeList = rsFeeList  '������ϸ
    If CheckDepend = False Then Exit Function
    Select Case mlngModule
    Case 1121 '  1121,'�����շѹ���
        mbytCall = 0
    Case 1137  '���˽��ʴ���
        mbytCall = 1
    Case Else
        mbytCall = 3
    End Select
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    Set rsRequare = mrsBrushData
    zlShowBrushCard = mintSucces > 0
End Function

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ�
    '����:���˺�
    '����:2009-12-23 10:57:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("����")) = "1|0"
        .ColData(.ColIndex("��������")) = "1|1"
        .Clear 1
        .Rows = 2
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(.ColIndex("���㿨����")) = True
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 10:02:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup, rsTemp As ADODB.Recordset
    
      
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
       Set .Font = vsGrid.Font
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    cbsThis.ActiveMenuBar.Visible = False
        
  
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_MoveCard, "�Ƴ�ˢ����¼"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_Affirm, "ȷ��   "): mcbrControl.BeginGroup = True
        mcbrControl.Flags = xtpFlagRightAlign
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�  "): mcbrControl.BeginGroup = True
        mcbrControl.Flags = xtpFlagRightAlign
    End With
    
    
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    Set mcbrToolBar = cbsThis.Add("���㿨", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    mcbrToolBar.ShowTextBelowIcons = True
    With mcbrToolBar.Controls
        Set rsTemp = zlGet���ѿ��ӿ�
        rsTemp.Sort = "���ƿ�,���"
        Do While Not rsTemp.EOF
            Set mcbrControl = .Add(xtpControlButton, conMenu_Square_BrushCard + Val(rsTemp!���), Nvl(rsTemp!����)): mcbrControl.BeginGroup = True
            mcbrControl.IconId = 3816 ' conMenu_Square_BrushCard
            mcbrControl.Parameter = Val(rsTemp!���)
 
            rsTemp.MoveNext
        Loop
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FALT, Asc("O"), mconMenu_Edit_Affirm
        .Add FALT, Asc("X"), conMenu_Edit_CardModify
 
         If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
         Do While Not rsTemp.EOF
            .Add FCONTROL, Asc(Trim(CStr(Chr(Val(rsTemp!���) + 64)))), conMenu_Square_BrushCard + Val(rsTemp!���)
            rsTemp.MoveNext
         Loop
     End With
         
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Err = 0: On Error Resume Next
    cbsThis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    With vsGrid
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - stbThis.Height
    End With
End Sub

Private Sub Form_Load()
    mstrTitle = "���㿨ˢ������"
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlDefCommandBars
    Call InitVsGrid
    Call zlInitBrushCardRec(mrsBrushData)
    Call vsGrid_GotFocus
End Sub
Private Function zlDeleteBrushCard(ByVal lng�ӿڱ�� As Long, Optional strCardNo As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��ˢ������
    '���:lng�ӿڱ��-�ӿڱ��
    '     strCardNo-����
    '����:
    '����:�ɹ�,����ture,���򷵻�False
    '����:���˺�
    '����:2009-12-31 11:10:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If mrsBrushData Is Nothing Then Exit Function
    If mrsBrushData.State <> 1 Then Exit Function
    If strCardNo = "" Then
        mrsBrushData.Filter = "�ӿڱ��=" & lng�ӿڱ��
    Else
        mrsBrushData.Filter = "�ӿڱ��=" & lng�ӿڱ�� & " and ����='" & strCardNo & "'"
    End If
    If mrsBrushData.EOF = False Then
        mrsBrushData.Delete (adAffectGroup)
    End If
    mrsBrushData.Filter = 0
    zlDeleteBrushCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlGet������Ѷ�(ByVal lng�ӿڱ�� As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ӿڵ�������Ѷ�
    '���:lng�ӿڱ��-�ӿڱ��
    '����:
    '����:���˺�
    '����:2009-12-31 11:40:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�������Ѻϼ� As Double
    dbl�������Ѻϼ� = 0
    Err = 0: On Error GoTo Errhand:
    If mrsBrushData Is Nothing Then GoTo ToCalc:
    If mrsBrushData.State <> 1 Then GoTo ToCalc:
    With mrsBrushData
        .Filter = "�ӿڱ��<>" & lng�ӿڱ��
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl�������Ѻϼ� = dbl�������Ѻϼ� + Val(Nvl(!������))
            .MoveNext
        Loop
        .Filter = "�ӿڱ��=" & lng�ӿڱ��
        
    End With
ToCalc:
    zlGet������Ѷ� = mdbl������Ѷ� - dbl�������Ѻϼ�
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
'ִ�о��幦��
Private Function zlExecuteBrushCard(ByVal lng�ӿڱ�� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�о��幦��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-12-23 10:22:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, rsSquare As ADODB.Recordset, dbl�������Ѻϼ� As Double, dbl������Ѷ� As Double
    Err = 0: On Error GoTo Errhand:
    Set rsTemp = zlGet���ѿ��ӿ�
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    rsTemp.Find "���=" & lng�ӿڱ��, , , 1
    If rsTemp.EOF Then Exit Function
    If Val(Nvl(rsTemp!���ƿ�)) <> 1 Then
        'ִ����ص������ӿ�
        If mobjBrushCard.zlInitInterFacel(lng�ӿڱ��) = False Then Exit Function
        
        '��ͨ��Ҫ�õ���ǰѡ��ˢ����Ϣ,��ˣ��������ݴ���,�Ա�鿴(����Ӱ���Ѿ�ˢ�˵�����)
        Set rsSquare = zlDatabase.CopyNewRec(mrsBrushData)
        If mobjBrushCard.zlBrushCardSquare(mbytCall, Me, lng�ӿڱ��, mrsFeeList, zlGet������Ѷ�(lng�ӿڱ��), rsSquare) = False Then Exit Function
        dbl�������Ѻϼ� = 0
        If Not rsSquare Is Nothing Then
            If rsSquare.State = 1 Then
                '��Ҫ����Ƿ񳬹���������Ѷ�
                rsSquare.Filter = "�ӿڱ��=" & lng�ӿڱ��
                If rsSquare.RecordCount <> 0 Then rsSquare.MoveFirst
                Do While Not rsSquare.EOF
                    dbl�������Ѻϼ� = dbl�������Ѻϼ� + Val(Nvl(rsSquare!������))
                    If Val(Nvl(rsSquare!���)) < Val(Nvl(rsSquare!������)) Then
                        ShowMsgbox "ע��:" & _
                                   "    " & rsTemp!���� & " �Ŀ���Ϊ:" & Nvl(rsSquare!����) & "�����(" & Format(Val(Nvl(rsSquare!���)), gVbFmtString.FM_���) & ")������֧��ˢ�����(" & Format(Val(Nvl(rsSquare!������)), gVbFmtString.FM_���) & ")������!"
                        
                        Exit Function
                    End If
                    rsSquare.MoveNext
                Loop
                If dbl�������Ѻϼ� > dbl������Ѷ� Then
                    ShowMsgbox "ע��:" & vbCrLf & "    ����ˢ���������ֻ��ˢ" & Format(dbl������Ѷ�, gVbFmtString.FM_���) & "Ԫ,��������ˢ��" & Format(dbl�������Ѻϼ�, gVbFmtString.FM_���) & "Ԫ,����!"
                    Exit Function
                End If
                '��Ҫ��rsSquare�е����ݣ����µ��Ѿ�ˢ����������
                'ɾ������;
                 Call zlDeleteBrushCard(lng�ӿڱ��, "")
                If rsSquare.RecordCount <> 0 Then rsSquare.MoveFirst
                Do While Not rsSquare.EOF
                    With mrsBrushData
                        .AddNew
                         !�ӿڱ�� = rsSquare!�ӿڱ��
                         !���ѿ�ID = rsSquare!���ѿ�ID
                         !���� = rsSquare!����
                         !���㷽ʽ = rsTemp!���㷽ʽ
                         !������ = rsSquare!������
                         !��� = rsSquare!���
                         !������ = rsSquare!������
                         !����ʱ�� = rsSquare!����ʱ��
                         !��ע = rsSquare!��ע
                         !�����־ = 0
                         .Update
                    End With
                    rsSquare.MoveNext
                Loop
            Else
                Call zlDeleteBrushCard(lng�ӿڱ��, "")
            End If
        Else
            Call zlDeleteBrushCard(lng�ӿڱ��, "")
        End If
        GoTo BrushData:
    End If
    mrsBrushData.Filter = "�ӿڱ��=" & lng�ӿڱ��
     
    '���ƿ�,��Ҫ������ص�ˢ������
    If frmSquareBrushCard.zlShowBrushCard(Me, lng�ӿڱ��, mbytCall, mrsFeeList, zlGet������Ѷ�(lng�ӿڱ��), mrsBrushData) = False Then Exit Function
BrushData:
    Dim strCardNo As String
    lng�ӿڱ�� = 0
    With vsGrid
        If .Row > 0 Then
            lng�ӿڱ�� = Val(.Cell(flexcpData, .Row, .ColIndex("���㿨����")))
            strCardNo = Trim(.Cell(flexcpData, .Rows - 1, .ColIndex("����")))
        End If
    End With
    Call FullDataToGrid(lng�ӿڱ��, strCardNo)
    zlExecuteBrushCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlMoveCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƴ���ǰˢ���Ŀ�Ƭ��Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 11:15:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCurRow  As Long
    Err = 0: On Error GoTo Errhand:
    
    With vsGrid
        If .Row < 0 Then Exit Function
        If .Rows < 2 Then Exit Function
        If Trim(.Cell(flexcpData, .Row, .ColIndex("����"))) <> "" Then
            '���ҿ���
            mrsBrushData.Filter = "�ӿڱ��=" & Val(.Cell(flexcpData, .Row, .ColIndex("���㿨����"))) & " and ����='" & Trim(.Cell(flexcpData, .Row, .ColIndex("����"))) & "'"
            If mrsBrushData.EOF = False Then
                mrsBrushData.Delete adAffectCurrent
                mrsBrushData.MoveNext
            End If
            mrsBrushData.Filter = 0
        End If
        lngCurRow = .Row
        Call FullDataToGrid
        If lngCurRow < .Rows - 1 Then
            lngCurRow = lngCurRow + 1
        Else
            lngCurRow = .Rows - 1
        End If
        If lngCurRow < 1 Then lngCurRow = 1
        If lngCurRow > 1 Then .Row = lngCurRow
    End With
    zlMoveCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub FullDataToGrid(Optional lngDefault�ӿڱ�� As Long = 0, Optional strDefaultCardNo As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:lngDefault�ӿڱ��-ȱʡָ��Ľӿ����
    '     strDefaultCardNo-ȱʡָ��Ŀ���
    '����:���˺�
    '����:2009-12-23 11:42:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�ӿڱ�� As Long, lngRow As Long, dbl����ˢ�� As Double, dbl����ˢ���ܼ� As Double
    With vsGrid
        .Clear 1: .Rows = 2
        .TextMatrix(1, .ColIndex("���㿨����")) = ""
        mrsBrushData.Filter = 0
        mrsBrushData.Sort = "�ӿڱ��,����"
        lngRow = 1: dbl����ˢ�� = 0: dbl����ˢ���ܼ� = 0
        If mrsBrushData.RecordCount <> 0 Then mrsBrushData.MoveFirst
        Do While Not mrsBrushData.EOF
            If lng�ӿڱ�� <> Val(Nvl(mrsBrushData!�ӿڱ��)) Then
                If lng�ӿڱ�� <> 0 Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("���㿨����")) = .TextMatrix(.Rows - 2, .ColIndex("���㿨����"))
                    .Cell(flexcpData, .Rows - 1, .ColIndex("���㿨����")) = .Cell(flexcpData, .Rows - 2, .ColIndex("���㿨����"))
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = "С��"
                    .TextMatrix(.Rows - 1, .ColIndex("�����")) = ""
                    .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = ""
                    .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(dbl����ˢ��, gVbFmtString.FM_���)
                    If lngDefault�ӿڱ�� = lngDefault�ӿڱ�� And "С��" = strDefaultCardNo Then
                        .Row = .Rows - 1
                    End If
                End If
                dbl����ˢ�� = 0
                lng�ӿڱ�� = Val(Nvl(mrsBrushData!�ӿڱ��))
            End If
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("���㿨����"))) <> "" Then
                .Rows = .Rows + 1
            End If
            .TextMatrix(.Rows - 1, .ColIndex("���㿨����")) = Nvl(mrsBrushData!������)
            .Cell(flexcpData, .Rows - 1, .ColIndex("���㿨����")) = Nvl(mrsBrushData!�ӿڱ��)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(zlIsCardNoShowPW(Val(Nvl(mrsBrushData!�ӿڱ��))), "****", Nvl(mrsBrushData!����))
            .Cell(flexcpData, .Rows - 1, .ColIndex("����")) = Nvl(mrsBrushData!����)
            .TextMatrix(.Rows - 1, .ColIndex("�����")) = Format(Val(Nvl(mrsBrushData!���)), gVbFmtString.FM_���)
            .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = Nvl(mrsBrushData!���㷽ʽ)
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(Val(Nvl(mrsBrushData!������)), gVbFmtString.FM_���)
            .TextMatrix(.Rows - 1, .ColIndex("��ע")) = Nvl(mrsBrushData!��ע)
            If lngDefault�ӿڱ�� = lngDefault�ӿڱ�� And Nvl(mrsBrushData!����) = strDefaultCardNo Then
                .Row = .Rows - 1
            End If
            dbl����ˢ�� = dbl����ˢ�� + Val(Nvl(mrsBrushData!������))
            dbl����ˢ���ܼ� = dbl����ˢ���ܼ� + Val(Nvl(mrsBrushData!������))
            mrsBrushData.MoveNext
        Loop
        If mrsBrushData.RecordCount <> 0 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("���㿨����")) = .TextMatrix(.Rows - 2, .ColIndex("���㿨����"))
            .Cell(flexcpData, .Rows - 1, .ColIndex("���㿨����")) = .Cell(flexcpData, .Rows - 2, .ColIndex("���㿨����"))
            .TextMatrix(.Rows - 1, .ColIndex("����")) = "С��"
            .TextMatrix(.Rows - 1, .ColIndex("�����")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(dbl����ˢ��, gVbFmtString.FM_���)
            If lngDefault�ӿڱ�� = Val(.Cell(flexcpData, .Rows - 1, .ColIndex("���㿨����"))) And "С��" = strDefaultCardNo Then
                .Row = .Rows - 1
            End If
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("���㿨����")) = "�ϼ�"
            .Cell(flexcpData, .Rows - 1, .ColIndex("���㿨����")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("����")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("�����")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(dbl����ˢ���ܼ�, gVbFmtString.FM_���)
            If lngDefault�ӿڱ�� = 0 And "�ϼ�" = strDefaultCardNo Then
                .Row = .Rows - 1
            End If
        End If
        If .Row < 0 And .Rows > 1 Then .Row = 1
    End With
End Sub
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.ID
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Edit_MoveCard '�Ƴ�ˢ����¼
            Call zlMoveCard
        Case mconMenu_Edit_Affirm
            mintSucces = mintSucces + 1
            Unload Me
        Case conMenu_File_Exit '
            mintSucces = 0
            Unload Me
        Case Else
            If Val(Control.Parameter) > 0 Then
                'ִ�о��幦��:
                Call zlExecuteBrushCard(Val(Control.Parameter))
            End If
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If Me.Visible = False Then Exit Sub

    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
         
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
End Sub
 
Private Sub vsGrid_GotFocus()
    zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_LostFocus()
    zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub
