VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.51#0"; "Codejock.CommandBars.Unicode.9510.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInsertElement 
   BorderStyle     =   0  'None
   Caption         =   "����̶�����Ҫ��"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   Icon            =   "frmInsertElement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cmbMainType 
      Height          =   300
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   585
      Width           =   2850
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1305
      ScaleHeight     =   330
      ScaleWidth      =   3075
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   45
      Width           =   3075
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   2490
      End
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2520
         Picture         =   "frmInsertElement.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   180
      ScaleHeight     =   1725
      ScaleWidth      =   1545
      TabIndex        =   5
      Top             =   1215
      Width           =   1545
      Begin MSComctlLib.TreeView Tree 
         Height          =   1230
         Left            =   45
         TabIndex        =   10
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2170
         _Version        =   393217
         Style           =   1
         ImageList       =   "imgList"
         Appearance      =   0
      End
      Begin VB.Shape shpTree 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   2160
      ScaleHeight     =   4515
      ScaleWidth      =   3615
      TabIndex        =   4
      Top             =   1170
      Width           =   3615
      Begin VB.PictureBox picExample 
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   90
         ScaleHeight     =   780
         ScaleWidth      =   3030
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3690
         Width           =   3030
         Begin RichTextLib.RichTextBox rtbThis 
            Height          =   420
            Left            =   135
            TabIndex        =   7
            Top             =   180
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   741
            _Version        =   393217
            BackColor       =   -2147483633
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmInsertElement.frx":0714
         End
         Begin VB.Shape shpExample 
            BorderColor     =   &H00808080&
            Height          =   375
            Left            =   90
            Top             =   135
            Width           =   330
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid fgThis 
         Height          =   2595
         Left            =   90
         TabIndex        =   2
         Top             =   990
         Width           =   3000
         _cx             =   5292
         _cy             =   4577
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
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
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   15198183
         GridColorFixed  =   15198183
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   240
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInsertElement.frx":07B1
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
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
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Shape shpFG 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   0
         Top             =   900
         Width           =   330
      End
   End
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4920
      Left            =   1890
      MouseIcon       =   "frmInsertElement.frx":0831
      MousePointer    =   99  'Custom
      ScaleHeight     =   4920
      ScaleMode       =   0  'User
      ScaleWidth      =   75
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1170
      Width           =   75
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   540
      Top             =   45
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
            Picture         =   "frmInsertElement.frx":0983
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsertElement.frx":0A96
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsertElement.frx":0BFD
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsertElement.frx":0D07
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsertElement.frx":12A1
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   45
      Top             =   45
      _Version        =   589875
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmInsertElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Ԫ������
Public Enum ElementTypeEnum
    cprETAllType = 0    '������������Ҫ��
    cprETFixed = 1      '�̶�����Ҫ��
    cprETTemp = 2       '��ʱ����Ҫ��
    cprETReplace = 3    '�滻��Ŀ
    cprETPoint = 4      'ָ����Ŀ
End Enum

Private Const ID_InsAndExit = 3         '���벢�˳�

Public EditStyle As ElementTypeEnum
Private frmParent As frmMain            'ϵͳ������

Public Sub SetParent(Parent As Object)
    Set frmParent = Parent
    '����ͼ��
    CommandBars.Icons = gfrmPublic.ImageManager.Icons
    CommandBars.AddImageList imgList
End Sub

Private Sub cmbMainType_Click()
    EditStyle = cmbMainType.ListIndex
    Select Case cmbMainType.ListIndex
    Case cprETAllType
        FillTree
    Case cprETFixed
        EditStyle = cprETFixed
        FillTree
    Case cprETTemp
        EditStyle = cprETTemp
        Tree.Nodes.Clear
        FillGrid 0
    Case cprETReplace
        EditStyle = cprETReplace
        FillTree
    Case cprETPoint
        EditStyle = cprETPoint
        Tree.Nodes.Clear
        FillGrid 0
    End Select
End Sub

Private Sub fgThis_DblClick()
    '��������Ҫ��
    PlayAction "Alert"
    
    Dim lngId As Long, strTMP As String
    Dim i As Long, j As Long, p As Long, q As Long

    '���ҹؼ��� ID ��
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lID As Long, bInKeys As Boolean, bNeeded As Boolean
    bInKeys = frmParent.IsBetweenAnyKeys(frmParent.Editor1.SelStart + 1, sType, lSS, lSE, lES, lEE, lID, bNeeded)
    If bInKeys = True Then Exit Sub

    '˫�����ĵ��в���Ԫ�أ�������Ŀ��
    With frmParent.Editor1
        If fgThis.Row > 0 And fgThis.Rows > 1 Then
            lngId = Val(fgThis.Cell(flexcpText, fgThis.Row, 0))
            '����Ԫ��
            Dim Rs As New ADODB.Recordset
            Rs.CursorLocation = adUseClient
            If EditStyle = cprETAllType Then
                Rs.Open "select * FROM ����������Ŀ where ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
            ElseIf EditStyle = cprETFixed Then
                Rs.Open "select * FROM ����������Ŀ where �滻�� = 0 or �滻�� is null and ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
            ElseIf EditStyle = cprETReplace Then
                Rs.Open "select * FROM ����������Ŀ where �滻�� = 1 and ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
            End If
            If Not Rs.EOF Then
                Dim Ele As New cCPRElement
                Set Ele = frmParent.Document.Elements.Add
                Ele.ID = Rs("ID")
                Ele.������ = NVL(Rs("������"))
                Ele.Ӣ���� = NVL(Rs("Ӣ����"))
                Ele.��ʼֵ = IIf(Trim(NVL(Rs("��ʼֵ"))) = "", "  ", NVL(Rs("��ʼֵ")))
                Ele.��ֵ�� = NVL(Rs("��ֵ��"))
                Ele.���� = NVL(Rs("����"), 0)
                Ele.��ʾ�� = NVL(Rs("��ʾ��"), 0)
                Ele.��λ = NVL(Rs("��λ"))
                Ele.�滻�� = NVL(Rs("�滻��"), 0)

                .Freeze

                .ForceEdit = True

                i = .SelStart
                If Ele.�滻�� = 0 Then
                    strTMP = Trim(Ele.������) & "��"
                    strTMP = strTMP & "ES(" & Format(Ele.��ˮ��, "00000000") & ",1,0)" & Ele.��ʼֵ & Ele.��λ
                    strTMP = strTMP & "EE(" & Format(Ele.��ˮ��, "00000000") & ",1,0)" & "��"
                    '�ַ������帳ֵ
                    .Range(i, i).Text = strTMP
    
                    '�����������ʾ����
                    i = i + Len(Trim(Ele.������) & "��")
                    .Range(i, i + 32 + Len(Ele.��ʼֵ) + Len(Ele.��λ)).Font.Protected = True
                    .Range(i, i + 16).Font.Hidden = True
                    .Range(i + 16 + Len(Ele.��ʼֵ) + Len(Ele.��λ), i + 16 + Len(Ele.��ʼֵ) + Len(Ele.��λ) + 16).Font.Hidden = True
                    .Range(i + 16, i + 16 + Len(Ele.��ʼֵ)).Font.Underline = tomHair 'tomthick
                    .Range(i + 16, i + 16 + Len(Ele.��ʼֵ)).Font.ForeColor = vbBlue
    
                    .Range(i + 32 + Len(Ele.��ʼֵ) + Len(Ele.��λ) + 1, i + 32 + Len(Ele.��ʼֵ) + Len(Ele.��λ) + 1).Selected
                ElseIf Ele.�滻�� = 1 Then
                    j = Len("{{" & Ele.Ӣ���� & "}}")
                    strTMP = "ES(" & Format(Ele.��ˮ��, "00000000") & ",1,0){{" & Ele.Ӣ���� & "}}EE(" & Format(Ele.��ˮ��, "00000000") & ",1,0)" & "��"
                    .Range(i, i).Text = strTMP
                    .Range(i, i + j + 32).Font.Protected = True
                    .Range(i, i + 16).Font.Hidden = True
                    .Range(i + 16 + j, i + 32 + j).Font.Hidden = True
                    .Range(i + 32 + j, i + 32 + j + 1).Font.Protected = False
                    .Range(i + 16, i + 16 + j).Font.Underline = tomHair
                    .Range(i + 16, i + 16 + j).Font.ForeColor = vbBlue
                    
                    .Range(i + 32 + j + 1, i + 32 + j + 1).Selected
                    
                End If
                .ForceEdit = False
                .UnFreeze
                .SetFocus
            End If
            Rs.Close
            Set Rs = Nothing
        End If
    End With
End Sub

Private Sub fgThis_RowColChange()
    Dim lngId As Long
    If fgThis.Row > 0 And fgThis.Rows > 1 Then
        lngId = Val(fgThis.Cell(flexcpText, fgThis.Row, 0))
        FillExample lngId
    End If
End Sub

Private Sub FillExample(lngId As Long)
    '���ʾ���ı�
    Dim i As Long, j As Long, lngLen As Long
    LockWindowUpdate rtbThis.hWnd
    Dim Rs As New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    If EditStyle = cprETAllType Then
        Rs.Open "select * FROM ����������Ŀ where ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
    ElseIf EditStyle = cprETFixed Then
        Rs.Open "select * FROM ����������Ŀ where �滻�� = 0 or �滻�� is null and ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
    ElseIf EditStyle = cprETReplace Then
        Rs.Open "select * FROM ����������Ŀ where �滻�� = 1 and ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
    Else
        Exit Sub
    End If
    With rtbThis
        If Rs.EOF Then
            .Text = ""
        Else
            If Rs("�滻��") = 1 Then
                .Text = Rs("������") & ": " & "{{" & Rs("Ӣ����") & "}}"
                .SelStart = Len(Rs("������") & ": ")
                .SelLength = Len(.Text)
                .SelColor = vbBlue
                .SelUnderline = True
            Else
                .Text = Rs("������") & "��" & IIf(Trim(NVL(Rs("��ʼֵ"))) = "", "  ", NVL(Rs("��ʼֵ"))) & Rs("��λ")
                lngLen = Len(Rs("������"))
                .SelStart = lngLen + 1
                .SelLength = Len(IIf(Trim(NVL(Rs("��ʼֵ"))) = "", "  ", NVL(Rs("��ʼֵ"))))
                .SelColor = vbBlue
                .SelUnderline = True
            End If
        End If
    End With
    LockWindowUpdate 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Me.Hide
    End If
End Sub

'������������༰����Ŀ��TreeView
Private Sub CreateItemTree()
    Dim rsItem As New ADODB.Recordset
    Dim sCurID As String
    Dim iStackPoint As Integer '��ջָ��
    Dim aStack() As String '��ջ
    Dim TmpNode As Node
    Dim i As Integer, AttrID As String
    
    '������������������ȡ
    
    clsDatabase.OpenRecordset rsItem, "Select * From ������������ Order By ����", ""
    Do While Not rsItem.EOF
        Load cmdTab(cmdTab.Count)
        With cmdTab(cmdTab.Count - 1)
            .Caption = rsItem("����") '+ IIf(rsItem("�̶�") = 1, "��ֻ����", "")
            .Tag = rsItem("�̶�") & "-" & rsItem("����")
            .ZOrder 0
            .Visible = True
        End With
        Load tvwItem(tvwItem.Count)
        tvwItem(tvwItem.Count - 1).Visible = True
        
        rsItem.MoveNext
    Loop
    
    For i = 1 To cmdTab.Count - 1
        AttrID = Mid(cmdTab(i).Tag, InStr(cmdTab(i).Tag, "-") + 1)
    
        ReDim aStack(0)
        aStack(0) = ""
        iStackPoint = 0
        
        Do While iStackPoint > -1
            sCurID = aStack(iStackPoint)
            '����¼����������
            clsDatabase.OpenRecordset rsItem, "Select * From ������������ Where �ϼ�ID" + IIf(sCurID = "", " is null ", "='" + sCurID + "' ") + "And ����='" + AttrID + "'", "��ѯ������Ŀ����"
            
            '�÷�����¼��Ѵ�������Ӷ�ջ�е���
            iStackPoint = iStackPoint - 1
            
            Do While Not rsItem.EOF
                If sCurID = "" Then
                    Set TmpNode = tvwItem(i).Nodes.Add(, , "Key" & rsItem("ID"), rsItem("����"), "Class")
                Else
                    Set TmpNode = tvwItem(i).Nodes.Add("Key" + sCurID, tvwChild, "Key" & rsItem("ID"), rsItem("����"), "Class")
                End If
                TmpNode.Tag = rsItem("����") & "||" & rsItem("����") & "||" & rsItem("����") & "||" & rsItem("����")
                
                '���·���ѹ���ջ
                iStackPoint = iStackPoint + 1
                ReDim Preserve aStack(iStackPoint)
                aStack(iStackPoint) = rsItem("ID")
                
                rsItem.MoveNext
            Loop
        Loop
    Next
End Sub


Private Sub Form_Load()
    '##########################################################################################
    '## ����λ�ûָ�
    Me.Left = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainLeft", (Screen.Width - 10000) / 2)
    Me.Top = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainTop", (Screen.Height - 8000) / 2)
    Me.Width = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainWidth", 5000)
    Me.Height = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainHeight", 8000)

    '##########################################################################
    '��������ť��ʼ��
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    'ͼ���
    Dim ControlAdd As CommandBarControl
    With CommandBars.ActiveMenuBar.Controls
        Set ControlAdd = .Add(xtpControlButton, ID_REFRESH, "ˢ��")
        ControlAdd.BeginGroup = True
        .Add xtpControlButton, ID_ADD, "���"
        .Add xtpControlButton, ID_DELETE, "ɾ��"
        .Add xtpControlButton, ID_MODIFY, "�޸�"
        .Add xtpControlButton, ID_SEARCH, "����"
        
        Set ControlAdd = .Add(xtpControlButton, ID_InsAndExit, "����")
        ControlAdd.BeginGroup = True
        ControlAdd.Style = xtpButtonIconAndCaption
    End With
    
    '��ʾ��չ��ť
    CommandBars.ActiveMenuBar.EnableDocking xtpFlagStretched
    CommandBars.Options.ShowExpandButtonAlways = False
    CommandBars.EnableCustomization (False)
    CommandBars.Options.UseDisabledIcons = True
    '�Ƿ���ʾ���в˵�
    CommandBars.Options.AlwaysShowFullMenus = False
    '##########################################################################
    
    Call FillTree
    Call FillCmb
    
    fgThis.Editable = flexEDNone
    fgThis.ToolTipText = "˫�����������Ҫ��"
    rtbThis.ToolTipText = "ʾ���ı�"
    txtFind.ToolTipText = "�������ѯ����"
    cmdFind.ToolTipText = "�����ѯ"
    cmbMainType.ToolTipText = "����Ҫ�ط���"
    
    Me.KeyPreview = True
End Sub

Private Sub FillTree()
    '�������Ҫ�ط���
    Tree.Nodes.Clear
    Tree.Style = tvwPictureText
    Dim obj As Node
    Dim Rs As New ADODB.Recordset, i As Long
    Rs.CursorLocation = adUseClient
    Select Case EditStyle
    Case cprETAllType
        Rs.Open "select A.*,(select count(*) from ����������Ŀ B where B.����ID=A.ID) as ��Ŀ FROM ������������ A", gcnOracle, adOpenStatic, adLockReadOnly
    Case cprETFixed
        Rs.Open "select A.*,(select count(*) from ����������Ŀ B where B.����ID=A.ID and (�滻�� = 0 or �滻�� is null)) as ��Ŀ FROM ������������ A", gcnOracle, adOpenStatic, adLockReadOnly
    Case cprETTemp
        Exit Sub
    Case cprETReplace
        Rs.Open "select A.*,(select count(*) from ����������Ŀ B where B.����ID=A.ID and �滻�� = 1) as ��Ŀ FROM ������������ A", gcnOracle, adOpenStatic, adLockReadOnly
    Case cprETPoint
        Exit Sub
    End Select
   
    If Not Rs.EOF Then
        FillGrid Rs("ID")
        ReDim EleTypeID(Rs.RecordCount) As Long
        i = 1
        Do While Not Rs.EOF
            Tree.Nodes.Add , , "K" & Rs("ID"), Rs("����"), IIf(Rs("��Ŀ") = 0, 3, 1)
            i = i + 1
            Rs.MoveNext
        Loop
    End If
    Tree.Nodes(1).Selected = True
End Sub

Private Sub FillCmb()
    '������η���
    With cmbMainType
        .Clear
        .AddItem "���з���", 0
        .AddItem "�̶�����Ҫ��", 1
        .AddItem "��ʱ����Ҫ��", 2
        .AddItem "�滻��Ŀ", 3
        .AddItem "ָ����Ŀ", 4
    End With
    cmbMainType.ListIndex = 0
End Sub

Private Sub FillGrid(ByVal lngId As Long)
    '�������Ҫ��
    Dim Rs As New ADODB.Recordset, i As Long
    Rs.CursorLocation = adUseClient
    fgThis.Clear
    fgThis.Rows = 1
    fgThis.Cell(flexcpText, 0, 2) = "����"
    fgThis.Cell(flexcpText, 0, 3) = "����"
'    fgThis.Select 0, 1
'    fgThis.CellAlignment = flexAlignCenterCenter
'    fgThis.Select 0, 2
'    fgThis.CellAlignment = flexAlignCenterCenter
    fgThis.ColAlignment(2) = flexAlignLeftCenter
    
    If lngId = 0 Then
        If EditStyle = cprETAllType Then
            Rs.Open "select * from ����������Ŀ ", gcnOracle, adOpenStatic, adLockReadOnly
        ElseIf EditStyle = cprETFixed Then
            Rs.Open "select * from ����������Ŀ where �滻�� = 0 or �滻�� is null ", gcnOracle, adOpenStatic, adLockReadOnly
        ElseIf EditStyle = cprETReplace Then
            Rs.Open "select * from ����������Ŀ where �滻�� = 1", gcnOracle, adOpenStatic, adLockReadOnly
        Else
            Exit Sub
        End If
    Else
        If EditStyle = cprETAllType Then
            Rs.Open "select * from ����������Ŀ where ����ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
        ElseIf EditStyle = cprETFixed Then
            Rs.Open "select * from ����������Ŀ where �滻�� = 0 or �滻�� is null and ����ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
        ElseIf EditStyle = cprETReplace Then
            Rs.Open "select * from ����������Ŀ where �滻�� = 1 and ����ID=" & lngId, gcnOracle, adOpenStatic, adLockReadOnly
        Else
            Exit Sub
        End If
    End If
    
    i = 1
    If Not Rs.EOF Then
        fgThis.Cols = 4
        fgThis.Rows = Rs.RecordCount + 1
        Do While Not Rs.EOF
            fgThis.Cell(flexcpText, i, 0) = Rs("ID")
            fgThis.Cell(flexcpPicture, i, 1) = imgList.ListImages(2).Picture
            fgThis.Cell(flexcpText, i, 2) = Rs("����")
            fgThis.Cell(flexcpText, i, 3) = Rs("������")
            i = i + 1
            Rs.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '���洰��λ��
    If Me.WindowState <> vbMinimized Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainLeft", Me.Left
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainTop", Me.Top
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainWidth", Me.Width
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainHeight", Me.Height
    End If
End Sub

Private Sub picFind_Resize()
On Error Resume Next
    txtFind.Width = picFind.ScaleWidth - Screen.TwipsPerPixelX - cmdFind.Width
    cmdFind.Left = txtFind.Left + txtFind.Width + Screen.TwipsPerPixelX
End Sub

Private Sub picLeft_Resize()
    Dim lngX As Long, lngY As Long
    lngX = Screen.TwipsPerPixelX
    lngY = Screen.TwipsPerPixelY
    
    Tree.Move lngX, lngY, picLeft.ScaleWidth - lngX * 2, picLeft.ScaleHeight - lngY * 2
    shpTree.Move 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    Dim lngX As Long, lngY As Long
    lngX = Screen.TwipsPerPixelX
    lngY = Screen.TwipsPerPixelY
    
    With picRight
        picExample.Move lngX, .ScaleHeight - picExample.Height, .ScaleWidth - lngX * 2
        fgThis.Move lngX * 2, lngY * 2, .ScaleWidth - lngX * 4, .ScaleHeight - picExample.Height - lngY * 5
        shpFG.Move lngX, fgThis.Top - lngY, fgThis.Width + lngX * 2, fgThis.Height + lngY * 2
        
        rtbThis.Move lngX * 2, lngY * 2, picExample.ScaleWidth - lngX * 4, picExample.ScaleHeight - 350
        shpExample.Move 0, 0, picExample.Width, picExample.Height
    End With
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '��ֱ�ָ����¼�
    If Button = 1 Then
        If picLeft.Width + X / Screen.TwipsPerPixelX > 30 And picRight.Width - X / Screen.TwipsPerPixelX > 100 Then
            picVBar.Left = picVBar.Left + X / Screen.TwipsPerPixelX
            picLeft.Width = picLeft.Width + X / Screen.TwipsPerPixelX
            
            picRight.Width = picRight.Width - X / Screen.TwipsPerPixelX
            picRight.Left = picRight.Left + X / Screen.TwipsPerPixelX
        End If
    End If
End Sub

Private Sub CommandBars_Resize()
    'λ�õ���
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.GetClientRect Left, Top, Right, Bottom
    picFind.Move Left + 1, Top + 1, Right - Left - 2
    cmbMainType.Move Left + 1, picFind.Top + picFind.Height + 1, Right - Left - 2
    
    picLeft.Move Left + 1, cmbMainType.Top + cmbMainType.Height + 1, picLeft.Width, Bottom - Top - picFind.Height - cmbMainType.Height - 4
    picVBar.Move picLeft.Left + picLeft.Width, picLeft.Top, 2, picLeft.Height
    picRight.Move picVBar.Left + picVBar.Width, picLeft.Top, Right - Left - picVBar.Left - picVBar.Width, picVBar.Height
End Sub

Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
    If EditStyle = cprETFixed Or EditStyle = cprETReplace Or EditStyle = cprETAllType Then
        FillGrid Mid(Tree.SelectedItem.Key, 2)
    Else
        fgThis.Clear
        fgThis.Rows = 1
        fgThis.Cell(flexcpText, 0, 1) = "����"
        fgThis.Cell(flexcpText, 0, 2) = "����"
        fgThis.ColAlignment(1) = flexAlignLeftCenter
    End If
End Sub
