VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmCombo 
   Caption         =   "������Ŀ����"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmCombo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin zlLisFlat.VsfGrid vsf 
      Height          =   2130
      Left            =   4785
      TabIndex        =   6
      Top             =   1110
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   3757
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3345
      Left            =   645
      TabIndex        =   10
      Top             =   1470
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   5900
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
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
            Picture         =   "frmCombo.frx":6852
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCombo.frx":6DEC
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   1170
      Left            =   675
      TabIndex        =   7
      Top             =   5310
      Width           =   8085
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   4395
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1155
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   120
         Picture         =   "frmCombo.frx":D64E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1065
         TabIndex        =   1
         Top             =   225
         Width           =   2250
      End
      Begin VB.Frame fra2 
         Height          =   75
         Left            =   30
         TabIndex        =   8
         Top             =   540
         Width           =   8010
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&M.��ϱ���"
         Height          =   180
         Index           =   0
         Left            =   3420
         TabIndex        =   4
         Top             =   780
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&N.�������"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   900
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&2.����"
         Height          =   180
         Left            =   480
         TabIndex        =   0
         Top             =   285
         Width           =   540
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   7380
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
            Picture         =   "frmCombo.frx":D8D4
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
            Picture         =   "frmCombo.frx":E168
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCombo.frx":E388
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCombo.frx":E5A8
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   11880
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   1138
         ButtonWidth     =   1455
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  ˢ��  "
               Key             =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "  ˢ��  "
               ImageKey        =   "Refresh"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgY 
      Height          =   4680
      Left            =   2550
      MousePointer    =   9  'Size W E
      Top             =   870
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "��ʾ�����¼�(&A)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShowOK 
         Caption         =   "�����Ѷ�����(&L)"
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmCombo"
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

Private WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1

Private Enum mCol
    LIS���� = 6
    LIS��ϱ��� = 7
End Enum

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
    vsf.Cell(flexcpBackColor, 1, mCol.LIS����, vsf.Rows - 1, mCol.LIS��ϱ���) = &HFFEBD7
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
        
        tvw.Nodes.Add , , "Root", "������Ŀ", "Root", "Root"
        
        gstrSQL = "select * " & _
             "from (Select DISTINCT ID,�ϼ�ID,����,���� " & _
                     "From ���Ʒ���Ŀ¼ " & _
                    "Where ���� = 5 " & _
                    "Start With ID IN (SELECT DISTINCT ����id FROM ������ĿĿ¼ WHERE (����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or ����ʱ�� is NULL) AND ���='C') " & _
                   "Connect by Prior �ϼ�ID = ID " & _
                   ") A " & _
            "ORDER BY A.����"
        
        Call OpenRecordset(rs)
        Do Until rs.EOF
            If IsNull(rs("�ϼ�id")) Then
                tvw.Nodes.Add "Root", tvwChild, "_" & rs("id"), "��" & rs("����") & "��" & rs("����"), "Class", "Class"
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
        
        
        gstrSQL = "Select RowNum As ���,A.ID,A.����,A.����,Decode(A.�����Ŀ,1,'��','') As ���,D.����,B.���� as �������� " & _
                    "From "
    
        If Val(Mid(tvw.SelectedItem.Key, 2)) > 0 Then
            
            If mnuViewShowAll.Checked Then
                gstrSQL = gstrSQL & "(Select ID,���� From ���Ʒ���Ŀ¼ Connect by Prior ID=�ϼ�id Start With ID = " & Val(Mid(tvw.SelectedItem.Key, 2)) & ") B,"
            Else
                gstrSQL = gstrSQL & "(Select ID,���� From ���Ʒ���Ŀ¼ where ID = " & Val(Mid(tvw.SelectedItem.Key, 2)) & ") B,"
            End If
        Else
            gstrSQL = gstrSQL & "(Select ID,���� From ���Ʒ���Ŀ¼) B,"
        End If
    
        gstrSQL = gstrSQL & _
                        "������ĿĿ¼ A," & _
                        "(Select * From ������Ŀ���� Where ����=1 And ����=1) D " & _
                    "Where (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01','YYYY-MM-DD')) " & _
                            "and A.ID=D.������Ŀid(+) " & _
                            "and A.���='C' " & _
                            "and B.ID=A.����ID "
        
        If mnuViewShowOK.Checked Then
            gstrSQL = "Select A.*,B.LIS����,B.LIS��ϱ��� From (" & gstrSQL & ") A,������ĿĿ¼_LIS B Where A.ID=B.������Ŀid Order By A.����"
        Else
            gstrSQL = "Select A.*,B.LIS����,B.LIS��ϱ��� From (" & gstrSQL & ") A,������ĿĿ¼_LIS B Where A.ID=B.������Ŀid(+) Order By A.����"
        End If
        
        Call OpenRecordset(rs, Me.Caption)
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
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strCode As String

    lngKey = Val(vsf.RowData(vsf.Row))
    strCode = Trim(txt(1).Text)

    '���Ψһ��
    gstrSQL = "Select 1 From ������ĿĿ¼_LIS Where ������Ŀid<>" & lngKey & " And LIS����='" & strCode & "'"
    rs.Open gstrSQL, gcnOracle
    If rs.BOF = False Then

        ShowSimpleMsg "����[" & strCode & "]�Ѿ���Ӧ������һ���Ӧ�����Ŀ��"

        vsf.Row = vsf.Row
        vsf.Col = mCol.LIS����
        vsf.ShowCell vsf.Row, vsf.Col

        DoEvents
        LocationObj txt(1)

        Exit Function

    End If
    
    CheckValid = True
    
End Function

Private Function SaveData() As Boolean
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngKey As Long
    Dim strCode As String
    Dim blnTran As Boolean
    
    On Error GoTo errHand
    
    lngKey = Val(vsf.RowData(vsf.Row))
    strCode = Trim(vsf.TextMatrix(vsf.Row, mCol.LIS����))
    
    If lngKey > 0 Then
    
        blnTran = True
        gcnOracle.BeginTrans
        
        strSQL = "Delete From ������ĿĿ¼_LIS Where ������Ŀid=" & lngKey
        gcnOracle.Execute strSQL

        If strCode <> "" Then
            
            strSQL = "Insert Into ������ĿĿ¼_LIS(������Ŀid,LIS����,LIS��ϱ���) Values (" & lngKey & ",'" & strCode & "','" & Trim(vsf.TextMatrix(vsf.Row, mCol.LIS��ϱ���)) & "')"
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
    
    With vsf
        
        .Cols = 0
        .NewColumn "", 255, 4
        
        .NewColumn "����", 1800, 1
        .NewColumn "����", 1080, 1
        .NewColumn "����", 1080, 1
        .NewColumn "���", 600, 1
        .NewColumn "��������", 1500, 1
        .NewColumn "LIS����", 1080, 1, , 1, GetMaxLength("������ĿĿ¼_LIS", "LIS����")
        .NewColumn "LIS��ϱ���", 1080, 1, , 1, GetMaxLength("������ĿĿ¼_LIS", "LIS��ϱ���")
                        
        .NewColumn "", 15, 1
        
        .ExtendLastCol = True
        .FixedCols = 1
        .Body.GridColor = &HC1C1C1
        .Body.GridColorFixed = &HC1C1C1
        .Body.GridLines = flexGridFlat
        .Body.BackColorFixed = .Body.BackColorBkg
        
        .Body.Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
        
        If mblnEditMode = False Then
            .EditMode(mCol.LIS����) = 0
            .EditMode(mCol.LIS��ϱ���) = 0
        End If
        
        .AppendRow = True
        
    End With
    
    txt(1).MaxLength = GetMaxLength("������ĿĿ¼_LIS", "LIS����")
    
    txt(1).Enabled = mblnEditMode
    
    txt(1).BackColor = IIf(mblnEditMode, &H80000005, &H8000000F)
    
    InitData = True
    
End Function

Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenuByCursor
    
    txtFind.Text = ""
    
    LocationObj txtFind
    
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    If InitData = False Then
        Unload Me
        Exit Sub
    End If
    
    DoEvents
    
    Call mnuViewRefresh_Click
        
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    Call RestoreWinState(Me, App.ProductName)
    'mblnEditMode = (InStr(gstrPrive, ";���ݶ���;") > 0)
    mblnEditMode = True
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With tvw
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With imgY
        .Top = tvw.Top
        .Width = 45
        .Height = tvw.Height
    End With
    
    With vsf
        .Left = imgY.Left + imgY.Width
        .Top = tvw.Top
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - fra.Height + 60 - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height - 60
        .Width = vsf.Width
    End With
    
    fra2.Left = 0
    fra2.Width = fra.Width
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY.Left = imgY.Left + X
    
    If imgY.Left < 3000 Then imgY.Left = 3000
    If Me.Width - imgY.Left - imgY.Width < 1000 Then imgY.Left = Me.Width - imgY.Width - 1000

    Form_Resize
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuHelpTopic_Click()
    Call ShowHelp(Me.hWnd, Me.Name)
End Sub

Private Sub mnuViewRefresh_Click()
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

Private Sub mnuViewShowAll_Click()
    
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    
    If Not (tvw.SelectedItem Is Nothing) Then
        mstrKey = ""
        Call tvw_NodeClick(tvw.SelectedItem)
    End If
    
End Sub

Private Sub mnuViewShowOK_Click()
    mnuViewShowOK.Checked = Not mnuViewShowOK.Checked
    
    If Not (tvw.SelectedItem Is Nothing) Then
        mstrKey = ""
        Call tvw_NodeClick(tvw.SelectedItem)
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
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


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ˢ��"
        Call mnuViewRefresh_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
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
        
        If Index = 1 Then vsf.TextMatrix(vsf.Row, mCol.LIS����) = txt(Index)
        If Index = 0 Then vsf.TextMatrix(vsf.Row, mCol.LIS��ϱ���) = txt(Index)
        
        If SaveData Then
            
            If Index = 0 Then
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
                
                txt(1).Text = vsf.TextMatrix(vsf.Row, mCol.LIS����)
                txt(0).Text = vsf.TextMatrix(vsf.Row, mCol.LIS��ϱ���)
                                
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

        lngCol = GetCol(vsf, "LIS����")

        On Error Resume Next
       
        If OldRow + 1 > vsf.FixedRows Then
            vsf.Cell(flexcpBackColor, OldRow, vsf.FixedCols, OldRow, lngCol - 1) = vsf.Body.BackColor
            vsf.Cell(flexcpBackColor, OldRow, lngCol + 2, OldRow, vsf.Cols - 1) = vsf.Body.BackColor

            vsf.Cell(flexcpForeColor, OldRow, vsf.FixedCols, OldRow, lngCol - 1) = vsf.Body.ForeColor
            vsf.Cell(flexcpForeColor, OldRow, lngCol + 2, OldRow, vsf.Cols - 1) = vsf.Body.ForeColor
        End If

        If NewRow + 1 > vsf.FixedRows Then
            vsf.Cell(flexcpBackColor, NewRow, vsf.FixedCols, NewRow, lngCol - 1) = vsf.Body.BackColorSel
            vsf.Cell(flexcpBackColor, NewRow, lngCol + 2, NewRow, vsf.Cols - 1) = vsf.Body.BackColorSel

            vsf.Cell(flexcpForeColor, NewRow, vsf.FixedCols, NewRow, lngCol - 1) = &H80000005
            vsf.Cell(flexcpForeColor, NewRow, lngCol + 2, NewRow, vsf.Cols - 1) = &H80000005

        End If

    End If
    
    If vsf.Col < mCol.LIS���� Then vsf.Col = mCol.LIS����
    If vsf.Col > mCol.LIS��ϱ��� Then vsf.Col = mCol.LIS��ϱ���
            
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
    
    If Col = mCol.LIS���� Then
        lngKey = Val(vsf.RowData(vsf.Row))
        strCode = Trim(vsf.EditText)
    
        '���Ψһ��
        gstrSQL = "Select 1 From ������ĿĿ¼_LIS Where ������Ŀid<>" & lngKey & " And LIS����='" & strCode & "'"
        rs.Open gstrSQL, gcnOracle
        If rs.BOF = False Then
    
            ShowSimpleMsg "����[" & strCode & "]�Ѿ���Ӧ������һ���Ӧ�����Ŀ��"
    
            Cancel = True
    
        End If
    End If
    
End Sub


