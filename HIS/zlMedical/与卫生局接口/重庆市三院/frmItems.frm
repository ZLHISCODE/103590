VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~4.OCX"
Begin VB.Form frmItems 
   Caption         =   "�������Ŀ"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmItems.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   4740
      Left            =   3075
      ScaleHeight     =   4740
      ScaleWidth      =   8160
      TabIndex        =   2
      Top             =   690
      Width           =   8160
      Begin zlPiesFlat.VsfGrid vsf 
         Height          =   2130
         Left            =   195
         TabIndex        =   3
         Top             =   135
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   3757
      End
      Begin VB.Frame fra 
         Height          =   1530
         Left            =   315
         TabIndex        =   4
         Top             =   2145
         Width           =   8085
         Begin VB.Frame fra2 
            Height          =   75
            Left            =   30
            TabIndex        =   11
            Top             =   540
            Width           =   8010
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1065
            TabIndex        =   10
            Top             =   225
            Width           =   2250
         End
         Begin VB.CommandButton cmdMenu 
            Height          =   270
            Left            =   120
            Picture         =   "frmItems.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   285
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1155
            TabIndex        =   8
            Top             =   720
            Width           =   1245
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   3450
            TabIndex        =   7
            Top             =   735
            Width           =   3840
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   1155
            TabIndex        =   6
            Top             =   1080
            Width           =   1245
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   3
            Left            =   3450
            TabIndex        =   5
            Top             =   1080
            Width           =   3840
         End
         Begin VB.Label lblFind 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&2.����"
            Height          =   180
            Left            =   480
            TabIndex        =   16
            Top             =   285
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&N.�ɱ�����"
            Height          =   180
            Index           =   1
            Left            =   180
            TabIndex        =   15
            Top             =   780
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&M.�ɱ�����"
            Height          =   180
            Index           =   0
            Left            =   2475
            TabIndex        =   14
            Top             =   795
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&A.��Ŀ��֧"
            Height          =   180
            Index           =   2
            Left            =   180
            TabIndex        =   13
            Top             =   1140
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&B.��Ŀ����"
            Height          =   180
            Index           =   3
            Left            =   2475
            TabIndex        =   12
            Top             =   1125
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   2520
      Left            =   105
      TabIndex        =   0
      Top             =   870
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   4445
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3705
      Top             =   1605
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
            Picture         =   "frmItems.frx":6AD8
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":7072
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7470
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmItems.frx":D8D4
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
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
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8580
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":E168
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":E388
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":E5A8
            Key             =   "Refresh"
         EndProperty
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
      DesignerControls=   "frmItems.frx":ED22
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
Attribute VB_Name = "frmItems"
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
Private mblnEditMode As Boolean
Private mstrSvrFind As String
Private mlngRow As Long
Private mblnShowAll As Boolean
Private mblnShowOK As Boolean

Private WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1

Private Enum mCol
    �ɱ����� = 6
    �ɱ����� = 7
    ��Ŀ��֧ = 8
    ��Ŀ���� = 9
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
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        
    End With

        
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "��ʾ�����¼�(&A)")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "��ʾ�Ѷ�����(&S)")
        
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

Private Function RefreshStateInfo() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ˢ��״̬������ʾ��Ϣ��
    '���أ� True
    '------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    
    If tvw.SelectedItem Is Nothing Then
        strInfo = ""
    Else
        strInfo = "���࡮" & tvw.SelectedItem.Text & "��"
        If Val(vsf.RowData(1)) > 0 Then
            strInfo = strInfo & "�¹��� " & vsf.Rows & " ����Ŀ��"
        Else
            strInfo = strInfo & "������Ŀ��"
        End If
        
    End If
    
    stbThis.Panels(2).Text = strInfo
    
    RefreshStateInfo = True
    
End Function

Private Function ApplyEditColor() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���ÿɱ༭�е���ɫ����ʾ����
    '���أ� True
    '------------------------------------------------------------------------------------------------------------------
    vsf.Cell(flexcpBackColor, 1, mCol.�ɱ�����, vsf.Rows - 1, mCol.��Ŀ����) = &HFFEBD7
    ApplyEditColor = True
    
End Function

Private Function zlMenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ʵ�ֻ����Ĳ�������
    '������
    '       strMenuItem          ��������
    '���أ� �ɹ�����True;���򷵻�False
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case strMenuItem
    Case "��������"
        
        tvw.Nodes.Clear
        vsf.Rows = 2
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        vsf.RowData(1) = 0
        
        tvw.Nodes.Add , , "R1", "������Ŀ", "Root", "Root"
        tvw.Nodes.Add , , "R2", "��ʷ����", "Root", "Root"
        tvw.Nodes.Add , , "R4", "�������", "Root", "Root"
        tvw.Nodes.Add , , "K5", "������Ŀ", "Root", "Root"
        
        gstrSQL = "Select ID,�ϼ�id,���� ,����  from ������������ where ����=1 Start With �ϼ�id is null connect by prior id =�ϼ�id  Order By ����"
        Call OpenRecordSet(rs)
        
        Do Until rs.EOF
            If IsNull(rs("�ϼ�id")) Then
                tvw.Nodes.Add "R1", tvwChild, "_" & rs("id"), "��" & rs("����") & "��" & rs("����"), "Class", "Class"
            Else
                tvw.Nodes.Add "_" & rs("�ϼ�id"), tvwChild, "_" & rs("id"), "��" & rs("����") & "��" & rs("����"), "Class", "Class"
            End If
            rs.MoveNext
        Loop
    
        gstrSQL = "Select ID,�ϼ�id,���� ,����  from ������������ where ����=2 Start With �ϼ�id is null connect by prior id =�ϼ�id  Order By ���� "
        Call OpenRecordSet(rs)
        Do Until rs.EOF
            If IsNull(rs("�ϼ�id")) Then
                tvw.Nodes.Add "R2", tvwChild, "_" & rs("id"), "��" & rs("����") & "��" & rs("����"), "Class", "Class"
            Else
                tvw.Nodes.Add "_" & rs("�ϼ�id"), tvwChild, "_" & rs("id"), "��" & rs("����") & "��" & rs("����"), "Class", "Class"
            End If
            rs.MoveNext
        Loop
        
        gstrSQL = "Select ID,�ϼ�id,���� ,����  from ������������ where ����=4 Start With �ϼ�id is null connect by prior id =�ϼ�id Order By ���� "
        Call OpenRecordSet(rs)
        Do Until rs.EOF
            If IsNull(rs("�ϼ�id")) Then
                tvw.Nodes.Add "R4", tvwChild, "_" & rs("id"), "��" & rs("����") & "��" & rs("����"), "Class", "Class"
            Else
                tvw.Nodes.Add "_" & rs("�ϼ�id"), tvwChild, "_" & rs("id"), "��" & rs("����") & "��" & rs("����"), "Class", "Class"
            End If
            rs.MoveNext
        Loop
        
        
    Case "��ϸ����"
        
        vsf.Rows = 2
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        vsf.RowData(1) = 0
    
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        gstrSQL = "Select RowNum As ���,A.ID,A.����,A.������,A.Ӣ����,Decode(A.����,0,'����',1,'�ı�') As ����,A.����,A.С��,A.��λ,B.���� as �������� " & _
                    "From "
        
        If Left(tvw.SelectedItem.Key, 1) <> "K" Then
            If mblnShowAll Then
                If Left(tvw.SelectedItem.Key, 1) = "R" Then
                    gstrSQL = gstrSQL & "(Select ID,���� From ������������ Where ����=" & Val(Mid(tvw.SelectedItem.Key, 2)) & ") B,"
                Else
                    gstrSQL = gstrSQL & "(Select ID,���� From ������������ Connect by Prior ID=�ϼ�id Start With ID = " & Val(Mid(tvw.SelectedItem.Key, 2)) & ") B,"
                End If
            Else
                gstrSQL = gstrSQL & "(Select ID,���� From ������������ Where ID=" & Val(Mid(tvw.SelectedItem.Key, 2)) & ") B,"
            End If
            
            gstrSQL = gstrSQL & _
                        "����������Ŀ A " & _
                    "Where B.ID=A.����ID "
        Else
            '������Ŀ
            
            gstrSQL = "Select RowNum As ���,A.ID,A.����,A.������,A.Ӣ����,Decode(A.����,0,'����',1,'�ı�') As ����,A.����,A.С��,A.��λ,'' as �������� " & _
                    "From ����������Ŀ A Where A.����id Is Null "
            
        End If
        
        If mblnShowOK Then
            gstrSQL = "Select A.*,B.�ɱ�����,B.�ɱ�����,B.��Ŀ��֧,B.��Ŀ���� From (" & gstrSQL & ") A,����������Ŀ_�ɱ� B Where A.ID=B.������Ŀid Order By ���"
        Else
            gstrSQL = "Select A.*,B.�ɱ�����,B.�ɱ�����,B.��Ŀ��֧,B.��Ŀ���� From (" & gstrSQL & ") A,����������Ŀ_�ɱ� B Where A.ID=B.������Ŀid(+) Order By ���"
        End If
        
        Call OpenRecordSet(rs, Me.Caption)
        If rs.BOF = False Then
            
            Call FillGrid(vsf, rs)
            
        End If
        
    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:

    ShowSimpleMsg Err.Description
    
End Function

Private Function CheckValid() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� �Ա༭�����ݺϷ��Խ���У��
    '���أ� ��Ч����True;��Ч����False
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strCode As String

    lngKey = Val(vsf.RowData(vsf.Row))
    strCode = Trim(txt(1).Text)

    '���Ψһ��
    gstrSQL = "Select 1 From ����������Ŀ_�ɱ� Where ������Ŀid<>" & lngKey & " And �ɱ�����='" & strCode & "'"
    rs.Open gstrSQL, gcnOracle
    If rs.BOF = False Then

        ShowSimpleMsg "����[" & strCode & "]�Ѿ���Ӧ������һ���Ӧ�����Ŀ��"

        vsf.Row = vsf.Row
        vsf.Col = mCol.�ɱ�����
        vsf.ShowCell vsf.Row, vsf.Col

        DoEvents
        LocationObj txt(1)

        Exit Function

    End If
    
    CheckValid = True
    
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ������ĺ����������
    '���أ� �ɹ�True;����False
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngKey As Long
    Dim strCode As String
    Dim blnTran As Boolean
    
    On Error GoTo errHand
    
    lngKey = Val(vsf.RowData(vsf.Row))
    strCode = Trim(vsf.TextMatrix(vsf.Row, mCol.�ɱ�����))
    
    If lngKey > 0 Then
        blnTran = True
        gcnOracle.BeginTrans
        
        strSQL = "Delete From ����������Ŀ_�ɱ� Where ������Ŀid=" & lngKey
        gcnOracle.Execute strSQL

        If strCode <> "" Then
            
            strSQL = "Insert Into ����������Ŀ_�ɱ�(������Ŀid,�ɱ�����,�ɱ�����,��Ŀ��֧,��Ŀ����) Values (" & lngKey & ",'" & strCode & "','" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɱ�����)) & "','" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ��֧)) & "','" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ����)) & "')"
            gcnOracle.Execute strSQL
  
        End If
        
        gcnOracle.CommitTrans
        blnTran = False
        
    End If
    
    SaveData = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���빦��ʱ�ĳ�ʼ������
    '���أ� True
    '------------------------------------------------------------------------------------------------------------------
    
    With vsf
        
        .Cols = 0
        .NewColumn "", 255, 4
        
        .NewColumn "������", 1800, 1
        .NewColumn "����", 900, 1
        .NewColumn "����", 900, 1
        .NewColumn "��λ", 900, 1
        
        .NewColumn "��������", 1800, 1
        .NewColumn "�ɱ�����", 900, 1, , 1, GetMaxLength("����������Ŀ_�ɱ�", "�ɱ�����")
        .NewColumn "�ɱ�����", 1500, 1, , 1, GetMaxLength("����������Ŀ_�ɱ�", "�ɱ�����")
        .NewColumn "��Ŀ��֧", 900, 1, , 1, GetMaxLength("����������Ŀ_�ɱ�", "��Ŀ��֧")
        .NewColumn "��Ŀ����", 1500, 1, , 1, GetMaxLength("����������Ŀ_�ɱ�", "��Ŀ����")
        
        .NewColumn "", 15, 1
        
        .ExtendLastCol = True
        .FixedCols = 1
        .Body.GridColor = &HC1C1C1
        .Body.GridColorFixed = &HC1C1C1
        .Body.GridLines = flexGridFlat
        .Body.BackColorFixed = .Body.BackColorBkg
        
        .Body.Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
        '.Body.ColHidden(6) = True
        
        If mblnEditMode = False Then
            .EditMode(mCol.�ɱ�����) = 0
            .EditMode(mCol.�ɱ�����) = 0
            .EditMode(mCol.��Ŀ��֧) = 0
            .EditMode(mCol.��Ŀ����) = 0
        End If
        
        .AppendRow = True
        
    End With
    
    txt(0).MaxLength = GetMaxLength("����������Ŀ_�ɱ�", "�ɱ�����")
    txt(1).MaxLength = GetMaxLength("����������Ŀ_�ɱ�", "�ɱ�����")
    
    txt(2).MaxLength = GetMaxLength("����������Ŀ_�ɱ�", "��Ŀ��֧")
    txt(3).MaxLength = GetMaxLength("����������Ŀ_�ɱ�", "��Ŀ����")
    
    txt(0).Enabled = mblnEditMode
    txt(1).Enabled = mblnEditMode
    
    txt(2).Enabled = mblnEditMode
    txt(3).Enabled = mblnEditMode
    
    txt(0).BackColor = IIf(mblnEditMode, &H80000005, &H8000000F)
    txt(1).BackColor = IIf(mblnEditMode, &H80000005, &H8000000F)
    txt(2).BackColor = IIf(mblnEditMode, &H80000005, &H8000000F)
    txt(3).BackColor = IIf(mblnEditMode, &H80000005, &H8000000F)
    
    InitData = True
    
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As Object

    On Error GoTo errHand
    
    Select Case Control.ID
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
                
                
        If Not (tvw.SelectedItem Is Nothing) Then
            mstrKey = ""
            Call tvw_NodeClick(tvw.SelectedItem)
        End If
    
        Case conMenu_View_Expend_CurExpend
'
            mblnShowAll = Not mblnShowAll
            If Not (tvw.SelectedItem Is Nothing) Then
                mstrKey = ""
                Call tvw_NodeClick(tvw.SelectedItem)
            End If
        
        Case conMenu_View_Expend_AllExpend
            
            mblnShowOK = Not mblnShowOK
            
            If Not (tvw.SelectedItem Is Nothing) Then
                mstrKey = ""
                Call tvw_NodeClick(tvw.SelectedItem)
            End If
            
        Case conMenu_View_Refresh
            
            Call RefreshData
                        
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

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Expend_CurExpend
    
        Control.Checked = mblnShowAll
    Case conMenu_View_Expend_AllExpend
        
        Control.Checked = mblnShowOK
        
    End Select
    
    
End Sub


Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenuByCursor
    
    txtFind.Text = ""
    
    LocationObj txtFind
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    
    Select Case Item.ID
    Case 1
        Item.Handle = tvw.hWnd
    Case 2
       Item.Handle = picContainer.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    If InitData = False Then
        Unload Me
        Exit Sub
    End If
    
    DoEvents
    
    Call RefreshData
        
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    Call InitMenuBar
    Call InitClient
    
    mblnShowAll = True
    mblnShowOK = False
    
    Call RestoreWinState(Me, App.ProductName)
    
    mblnEditMode = (InStr(gstrPrive, ";���ݶ���;") > 0)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub RefreshData()
    
    Dim strTvwKey As String
    Dim strVsfKey As String
    
    If Not (tvw.SelectedItem Is Nothing) Then strTvwKey = tvw.SelectedItem.Key
    strVsfKey = Val(vsf.RowData(vsf.Row))
    
    'װ�ط�������
    Call zlMenuClick("��������")
    
    On Error Resume Next
    
    tvw.Nodes(strTvwKey).Selected = True
    tvw.Nodes(strTvwKey).EnsureVisible
    
    On Error GoTo 0
    
    If tvw.SelectedItem Is Nothing Then
        If tvw.Nodes.Count > 0 Then
            tvw.Nodes(1).Selected = True
            tvw.Nodes(1).EnsureVisible
            tvw.Nodes(1).Expanded = True
        End If
    End If
    
    If Not (tvw.SelectedItem Is Nothing) Then
        'װ����ϸ����
        Call zlMenuClick("��ϸ����")
                        
        If Val(strVsfKey) > 0 Then
            For mlngLoop = 1 To vsf.Rows - 1
                If Val(vsf.RowData(mlngLoop)) = Val(strVsfKey) Then
                    vsf.Row = mlngLoop
                    vsf.ShowCell vsf.Row, vsf.Col
                    Exit For
                End If
            Next
        End If
        Call RefreshStateInfo
        Call ApplyEditColor
    End If
End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Dim strChar As String
    Dim intIndex As Integer
    
    strChar = "123456789ABCDEFGHIJKLMNOPQUVRSTWXYZ"
    
    For mlngLoop = 0 To vsf.Cols - 1
        
        If Trim(vsf.TextMatrix(0, mlngLoop)) <> "" Then
            
            intIndex = intIndex + 1
            
            mobjPopMenu.Add intIndex, "&" & Mid(strChar, intIndex, 1) & "." & Trim(vsf.TextMatrix(0, mlngLoop))
            
        End If
        
    Next

End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)

    lblFind.Caption = Caption
    
    txtFind.Left = lblFind.Left + lblFind.Width + 60
    
   
End Sub


Private Sub picContainer_Resize()
    On Error Resume Next
    
    With vsf
        .Left = 0
        .Top = 0
        .Width = picContainer.Width - .Left
        .Height = picContainer.Height - fra.Height + 60 - .Top
    End With
    
    With fra
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height - 60
        .Width = vsf.Width
    End With
    
    fra2.Left = 0
    fra2.Width = fra.Width
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    Call zlMenuClick("��ϸ����")
    Call RefreshStateInfo
    
    vsf.AppendRow = True
    
    Call ApplyEditColor

End Sub

Private Sub txt_GotFocus(Index As Integer)
    TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCol As String
    Dim lngCol As Long
    
    If KeyAscii = vbKeyReturn Then
        
        If CheckValid = False Then
            Exit Sub
        End If
        
        If Index = 1 Then vsf.TextMatrix(vsf.Row, mCol.�ɱ�����) = txt(Index)
        If Index = 0 Then vsf.TextMatrix(vsf.Row, mCol.�ɱ�����) = txt(Index)
        If Index = 2 Then vsf.TextMatrix(vsf.Row, mCol.��Ŀ��֧) = txt(Index)
        If Index = 3 Then vsf.TextMatrix(vsf.Row, mCol.��Ŀ����) = txt(Index)
        
        If SaveData Then
            If Index = 3 Then
                txtFind.SetFocus
            Else
                SendKeys "{TAB}"
            End If
        End If
        
    End If
    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index), txt(Index).MaxLength)
    
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strCol As String
    Dim lngCol As Long
    
    Dim lngLoop As Long
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    Dim lngRow As Long
    
    If KeyAscii = vbKeyReturn Then
        
        If Trim(txtFind.Text) <> "" Then
            
            strCol = Mid(lblFind.Caption, 4)
            lngCol = GetCol(vsf, strCol)
            
            If lngCol < 0 Then Exit Sub
            
            If mstrSvrFind <> txtFind.Text Then
                
                mstrSvrFind = txtFind.Text
                
                For lngLoop = 1 To vsf.Rows - 1
                    If InStr(UCase(vsf.TextMatrix(lngLoop, lngCol)), UCase(mstrSvrFind)) > 0 Then
                        mlngRow = lngLoop
                        Exit For
                    End If
                Next
                If lngLoop = vsf.Rows Then mlngRow = -1
            Else
                
                For lngLoop = mlngRow + 1 To vsf.Rows - 1
                    If InStr(UCase(vsf.TextMatrix(lngLoop, lngCol)), UCase(mstrSvrFind)) > 0 Then
                        mlngRow = lngLoop
                        Exit For
                    End If
                Next
                
                If lngLoop = vsf.Rows Then mlngRow = -1
            End If
            
            If mlngRow = -1 Then
                ShowSimpleMsg "�Ѿ������꣬���ٲ��ҽ���������һ�Σ�"
                mlngRow = 0
                DoEvents
            Else
                vsf.Row = mlngRow
                vsf.ShowCell vsf.Row, vsf.Col
                
                txt(1).Text = vsf.TextMatrix(vsf.Row, mCol.�ɱ�����)
                txt(0).Text = vsf.TextMatrix(vsf.Row, mCol.�ɱ�����)
                
                txt(3).Text = vsf.TextMatrix(vsf.Row, mCol.��Ŀ����)
                txt(2).Text = vsf.TextMatrix(vsf.Row, mCol.��Ŀ��֧)
                
                SendKeys "{TAB}"
            End If
            
        End If
        
        txtFind.SetFocus
        TxtSelAll txtFind
   
    End If
End Sub

Private Sub txtFind_GotFocus()
    TxtSelAll txtFind
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If mblnEditMode Then Call SaveData
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngCol As Long

    If OldRow <> NewRow Then

        lngCol = GetCol(vsf, "�ɱ�����")

        On Error Resume Next

        If OldRow + 1 > vsf.FixedRows Then
            vsf.Cell(flexcpBackColor, OldRow, vsf.FixedCols, OldRow, lngCol - 1) = vsf.Body.BackColor
            vsf.Cell(flexcpBackColor, OldRow, lngCol + 4, OldRow, vsf.Cols - 1) = vsf.Body.BackColor

            vsf.Cell(flexcpForeColor, OldRow, vsf.FixedCols, OldRow, lngCol - 1) = vsf.Body.ForeColor
            vsf.Cell(flexcpForeColor, OldRow, lngCol + 4, OldRow, vsf.Cols - 1) = vsf.Body.ForeColor
        End If

        If NewRow + 1 > vsf.FixedRows Then
            vsf.Cell(flexcpBackColor, NewRow, vsf.FixedCols, NewRow, lngCol - 1) = vsf.Body.BackColorSel
            vsf.Cell(flexcpBackColor, NewRow, lngCol + 4, NewRow, vsf.Cols - 1) = vsf.Body.BackColorSel

            vsf.Cell(flexcpForeColor, NewRow, vsf.FixedCols, NewRow, lngCol - 1) = &H80000005
            vsf.Cell(flexcpForeColor, NewRow, lngCol + 4, NewRow, vsf.Cols - 1) = &H80000005

        End If

    End If
    
    If vsf.Col < mCol.�ɱ����� Then vsf.Col = mCol.�ɱ�����
    If vsf.Col > mCol.��Ŀ���� Then vsf.Col = mCol.��Ŀ����
    
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_GotFocus()
    mlngRow = -1
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strCode As String
    
    If Col = mCol.�ɱ����� Then
        lngKey = Val(vsf.RowData(vsf.Row))
        strCode = Trim(vsf.EditText)
    
        '���Ψһ��
        gstrSQL = "Select 1 From ����������Ŀ_�ɱ� Where ������Ŀid<>" & lngKey & " And �ɱ�����='" & strCode & "'"
        rs.Open gstrSQL, gcnOracle
        If rs.BOF = False Then
    
            ShowSimpleMsg "����[" & strCode & "]�Ѿ���Ӧ������һ���Ӧ�����Ŀ��"
    
            Cancel = True
    
        End If
    End If
    
End Sub


