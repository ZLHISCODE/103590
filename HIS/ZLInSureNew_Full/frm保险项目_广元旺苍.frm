VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm������Ŀ_��Ԫ���� 
   BackColor       =   &H8000000A&
   Caption         =   "ҽ����Ŀ����"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   10110
   Icon            =   "frm������Ŀ_��Ԫ����.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ZL9BillEdit.BillEdit mshSum_S 
      Height          =   2745
      Left            =   3090
      TabIndex        =   4
      Top             =   1020
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   4842
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2670
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":0E42
            Key             =   "R"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":115C
            Key             =   "C"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":12B6
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3525
      Left            =   90
      TabIndex        =   7
      Top             =   960
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6218
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
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   3450
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":1708
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":1924
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":1B40
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":1D5A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":1F76
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   2760
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":2192
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":23AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":25CA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":27E4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀ_��Ԫ����.frx":2A00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1270
      BandCount       =   2
      _CBWidth        =   10110
      _CBHeight       =   720
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   5370
      Key1            =   "only"
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "ҽ������"
      Child2          =   "cmb����"
      MinHeight2      =   300
      Width2          =   2325
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   6345
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   3675
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   635
      SimpleText      =   $"frm������Ŀ_��Ԫ����.frx":2C1C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm������Ŀ_��Ԫ����.frx":2C63
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
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
   Begin VB.CommandButton cmdRestore 
      Caption         =   "��ԭ(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6750
      TabIndex        =   6
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5340
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuCX 
         Caption         =   "��ѯ������Ŀ"
      End
      Begin VB.Menu mnuSB 
         Caption         =   "�걨������Ŀ"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGet 
         Caption         =   "������ȡ��Ŀ�����Ϣ(&G)"
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
         Begin VB.Menu mnuViewToolSplit 
            Caption         =   "-"
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
      Begin VB.Menu mnuViewSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "���༭��Ŀ����(&I)"
      End
      Begin VB.Menu mnuViewClass 
         Caption         =   "���༭ҽ������(&C)"
      End
      Begin VB.Menu mnuViewSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R) "
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpWebL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)��"
      End
   End
End
Attribute VB_Name = "frm������Ŀ_��Ԫ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private int��˱�־ As Integer
Private classInsure As New clsInsure

Private Enum ColumnEnum
    cOL���� = 0
    cOL���� = 1
    col���� = 2
    COL��� = 3
    COL���� = 4
    COL��λ = 5
    col�۸� = 6
    col�ı䷽ʽ = 7
    col����ID = 8
    COLҽ������ = 9
    colҽ������ = 10
    colҽ������ = 11
    colҽ����ע = 12
    colԭ���� = 13
    col�������� = 14
    col��ҽ�� = 15
    'Modified By ���� ��������ɳ ԭ��û����ֻ�м���
    colƥ�����к� = 16
    col��˱�־ = 17
End Enum
Private Const mlng���볤�� As Long = 20

Dim mlngListIndex As Long   '�����ϴ��������ѡ������
Dim mblnLoad As Boolean
Dim msngStartX As Single    '�ƶ�ǰ����λ��
Dim mstrȨ�� As String

Dim mstrKey As String       'ǰһ�����ڵ�Ĺؼ�ֵ
Dim mint���� As Integer     '��ǰ��ʾ������
Dim mint���õ��� As Integer '����ר�ã�0��ʾ����������1��ʾ����������ɾ������˵���Ŀ��

Dim mlngCol As Long, mblnDesc As Boolean
Private mcnYB As New ADODB.Connection   'ҽ��ǰ�÷���������
Private mint���� As Integer         '��ǰ��ʾ������


Private Sub cbrThis_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm(NewHeight)
End Sub

Private Sub cmdRestore_Click()
    'Modified By ���� ��������ɳ
    If MsgBox("��ȷ��Ҫ�����޸���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Call FillSum(True)
    mshSum_S.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim lngRow As Long
    
    If mint���� = TYPE_�ɶ����� Then
        gcnOracle_�ɶ�����.BeginTrans
    ElseIf mint���� = TYPE_�ϳ����� Then
        gcnOracle_�ϳ�����.BeginTrans
    Else
        gcnOracle_��Ԫ����.BeginTrans
    End If
    
    On Error GoTo errHandle
    
    With mshSum_S
        '��������
        For lngRow = 1 To .Rows - 1
            If mint���� = TYPE_��Ԫ���� And InitInfor_��Ԫ����.���õ��� = 0 Then Exit Sub
            Select Case .TextMatrix(lngRow, col�ı䷽ʽ)
                Case "����", "�޸�"
                    '���������޸ķ���һ�������д���
                    '���̲���:
                    '    �շ�ϸĿID_IN IN ҽ��֧����Ŀ.�շ�ϸĿID%TYPE,
                    '    ����_IN       IN ҽ��֧����Ŀ.����%TYPE,
                    '    ����_IN       IN ҽ��֧����Ŀ.����%TYPE,
                    '    ����ID_IN     IN ҽ��֧����Ŀ.����ID%TYPE,
                    '    ��Ŀ����_IN   IN ҽ��֧����Ŀ.��Ŀ����%TYPE,
                    '    ��Ŀ����_IN   IN ҽ��֧����Ŀ.��Ŀ����%TYPE,
                    '    ��ע_IN       IN ҽ��֧����Ŀ.��ע%TYPE,
                    '    �Ƿ�ҽ��_IN   IN ҽ��֧����Ŀ.�Ƿ�ҽ��%TYPE
    
                    gstrSQL = "ZL_ҽ��֧����Ŀ_Modify(" & .RowData(lngRow) & "," & mint���� & "," & mint���� & "," & _
                               IIf(Val(.TextMatrix(lngRow, col����ID)) = 0, "null", .TextMatrix(lngRow, col����ID)) & ",'" & _
                               .TextMatrix(lngRow, COLҽ������) & "','" & .TextMatrix(lngRow, colҽ������) & "','" & .TextMatrix(lngRow, colҽ����ע) & _
                               IIf(mint���� = TYPE_������, "^^" & .TextMatrix(lngRow, colƥ�����к�) & "||" & _
                               IIf(Trim(.TextMatrix(lngRow, col��˱�־)) = "��", 1, IIf(Trim(.TextMatrix(lngRow, col��˱�־)) = "��", 2, 0)), "") & _
                               "'," & IIf(Trim(.TextMatrix(lngRow, col��ҽ��)) = "��", 0, 1) & ")"
                    If mint���� = TYPE_�ɶ����� Then
                        ExecuteProcedure_�ɶ����� Me.Caption
                    ElseIf mint���� = TYPE_�ϳ����� Then
                        ExecuteProcedure_�ϳ����� Me.Caption
                    Else
                        ExecuteProcedure_��Ԫ���� Me.Caption
                    End If
                    .TextMatrix(lngRow, colԭ����) = .TextMatrix(lngRow, COLҽ������)
                Case "ɾ��"
                    '���̲���:
                    '    �շ�ϸĿID_IN IN ҽ��֧����Ŀ.�շ�ϸĿID%TYPE,
                    '    ����_IN       IN ҽ��֧����Ŀ.����%TYPE,
                    '    ����_IN       IN ҽ��֧����Ŀ.����%TYPE
    
                    gstrSQL = "ZL_ҽ��֧����Ŀ_Delete(" & .RowData(lngRow) & "," & mint���� & "," & mint���� & ")"
                    If mint���� = TYPE_�ɶ����� Then
                        ExecuteProcedure_�ɶ����� Me.Caption
                    ElseIf mint���� = TYPE_�ϳ����� Then
                        ExecuteProcedure_�ϳ����� Me.Caption
                    Else
                        ExecuteProcedure_��Ԫ���� Me.Caption
                    End If
                    .TextMatrix(lngRow, colԭ����) = .TextMatrix(lngRow, COLҽ������)
            End Select
        Next
        
        '�����ݴ���������������������״̬
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, col�ı䷽ʽ) = ""
        Next
    End With
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
    If mint���� = TYPE_�ɶ����� Then
        gcnOracle_�ɶ�����.CommitTrans
    ElseIf mint���� = TYPE_�ϳ����� Then
        gcnOracle_�ϳ�����.CommitTrans
    Else
        gcnOracle_��Ԫ����.CommitTrans
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If mint���� = TYPE_�ɶ����� Then
        gcnOracle_�ɶ�����.RollbackTrans
    ElseIf mint���� = TYPE_�ϳ����� Then
        gcnOracle_�ϳ�����.RollbackTrans
    Else
        gcnOracle_��Ԫ����.RollbackTrans
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call FillTree
    End If
    
    Call mshSum_S_EnterCell(1, cOL����)
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mstrKey = ""
    mlngCol = 0
    mblnDesc = False
    mblnLoad = True
    
    
    If mint���� = TYPE_��Ԫ���� Then Call ҽ����ʼ��_��Ԫ����
    If mint���� = TYPE_�ɶ����� Then Call ҽ����ʼ��_�ɶ�����
    If mint���� = TYPE_�ϳ����� Then Call ҽ����ʼ��_�ϳ�����
    
    gstrSQL = "select ���,����,���� from ��������Ŀ¼ where ���<>0 and ����=[1] order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    
    With cmb����
        .Clear
        Do Until rsTemp.EOF
            .AddItem Nvl(rsTemp!����) & "--" & rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("���")
            If rsTemp("���") = mint���� Then
                '��ǰҽ����
                'ʹ��API�����Բ�����Click�¼�
                zlControl.CboSetIndex .hwnd, .NewIndex
                Call Fill����
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then
            'ʹ��API�����Բ�����Click�¼�
            zlControl.CboSetIndex .hwnd, 0
            Call Fill����
        End If
    End With
    mint���� = cmb����.ItemData(cmb����.ListIndex)
    
    Call InitSum
    RestoreWinState Me, App.ProductName
    
    mnuViewItem.Checked = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewItem", "False") <> "False"
    If mnuViewItem.Checked = False Then
        '�����жϴ�����
        mnuViewClass.Checked = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewClass", "False") <> "False"
    End If
    Call SetSkip
    zlControl.CboSetHeight cmb����, 3600
    

    If Nvl(InitInfor_��Ԫ����.���õ���, "0") = 0 And mint���� = TYPE_��Ԫ���� Then
       mnuEdit.Visible = True
       mnuCX.Visible = True
       mnuSB.Visible = True
       mnuEditGet.Visible = False
    Else
       mnuEdit.Visible = False
    End If
End Sub

Private Sub InitSum()
    '��ʼ�����ܱ����ʽ
    Dim lngCol As Long
    
    With mshSum_S
        ClearGrid mshSum_S
        
        'Modified By ���� ��������ɳ ԭ�������С���ƥ�����к�
        .Cols = 18
        
        .TextMatrix(0, cOL����) = "����"
        .TextMatrix(0, cOL����) = "�շ�ϸĿ"
        .TextMatrix(0, COL���) = "���"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, COL��λ) = "��λ"
        .TextMatrix(0, col�۸�) = "�۸�"
        .TextMatrix(0, col�ı䷽ʽ) = "�Ƿ��޸�"
        .TextMatrix(0, COLҽ������) = "ҽ����Ŀ����"
        .TextMatrix(0, colҽ������) = "ҽ����Ŀ����"
        .TextMatrix(0, COL����) = "����"
        .TextMatrix(0, colҽ������) = "����"
        .TextMatrix(0, col��˱�־) = "���"
        .TextMatrix(0, colҽ����ע) = "ҽ����Ŀ��ע"
        .TextMatrix(0, colԭ����) = "ԭҽ����Ŀ����"
        .TextMatrix(0, col����ID) = "����ID"
        .TextMatrix(0, col��������) = "ҽ����������"
        If Nvl(InitInfor_��Ԫ����.���õ���, "0") = 0 And mint���� = TYPE_��Ԫ���� Then
           .TextMatrix(0, col��ҽ��) = "δ����"
        Else
           .TextMatrix(0, col��ҽ��) = "�Ƿ�ҽ��"
        End If
        .TextMatrix(0, colƥ�����к�) = "ƥ�����к�"
        
        .ColWidth(cOL����) = 1000
        .ColWidth(cOL����) = 2000
        .ColWidth(COL���) = 1000
        .ColWidth(col����) = 600
        .ColWidth(COL��λ) = 600
        .ColWidth(col�۸�) = 800
        .ColWidth(col�ı䷽ʽ) = 0
        .ColWidth(COLҽ������) = 1200
        .ColWidth(colҽ������) = 1200
        .ColWidth(colҽ����ע) = 0
        .ColWidth(colԭ����) = 0
        .ColWidth(col����ID) = 0
        .ColWidth(col��������) = 1200
        .ColWidth(col��ҽ��) = 800
        .ColWidth(colƥ�����к�) = 0
        
        .ColWidth(COL����) = 0
        .ColWidth(colҽ������) = 0
        .ColWidth(col��˱�־) = 0
        
        
        For lngCol = 0 To .Cols - 1
            .ColAlignment(lngCol) = 1
        Next
        .ColAlignment(col�۸�) = 7
        .ColAlignment(col��ҽ��) = 4
        
        '���ø��еı༭����
        .ColData(COL����) = 5
        .ColData(colҽ������) = 5
        .ColData(col��˱�־) = 5
        .ColData(cOL����) = 5 '����ѡ��
        .ColData(cOL����) = 5
        .ColData(COL���) = 5
        .ColData(col����) = 5
        .ColData(COL��λ) = 5
        .ColData(col�۸�) = 5
        .ColData(col�ı䷽ʽ) = 5
        .ColData(COLҽ������) = 1
        .ColData(colҽ������) = 5
        
        .ColData(colҽ����ע) = 5
        .ColData(colԭ����) = 5
        .ColData(col����ID) = 5
        .ColData(col��������) = 3 'ѡ����
        .ColData(col��ҽ��) = -1 'ѡ����
        .ColData(colƥ�����к�) = 5
        
        .PrimaryCol = cOL����
        Call SetSkip
        .AllowAddRow = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdSave.Enabled = True Then
        MsgBox "ҽ����Ŀ�б������ڱ༭״̬�������˳�����", vbInformation, gstrSysName
        Cancel = 1
        Exit Sub
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewItem", mnuViewItem.Checked
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewClass", mnuViewClass.Checked
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Call ResizeForm(cbrThis.Height)
End Sub

Private Sub ResizeForm(ByVal cbrHeight As Single)
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrHeight, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    '�ұ�
    'tvwMain_S��λ��
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIf(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    'picV��λ��
    picV.Top = sngTop
    picV.Height = tvwMain_S.Height
    picV.Left = tvwMain_S.Left + tvwMain_S.Width
    
    cmdRestore.Top = sngBottom - cmdRestore.Height - 100
    cmdRestore.Left = ScaleWidth - cmdRestore.Width - 300
    cmdSave.Top = cmdRestore.Top
    cmdSave.Left = cmdRestore.Left - cmdSave.Width - 300
    
    If InStr(mstrȨ��, "��ɾ��") > 0 Then
        '���Ա༭
        sngBottom = cmdRestore.Top - 100
    End If
    
    mshSum_S.Left = picV.Left + picV.Width
    If ScaleWidth - mshSum_S.Left > 0 Then mshSum_S.Width = ScaleWidth - mshSum_S.Left
    mshSum_S.Top = sngTop
    mshSum_S.Height = IIf(sngBottom - mshSum_S.Top > 0, sngBottom - mshSum_S.Top, 0)
    
    Refresh
End Sub

Private Sub mnuCX_Click()
    Dim i As Integer
    Dim strSQL As String, intID As Integer
    Dim StrInput As String, strOutput As String
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim int����id As Integer
    Dim strTmpArr As Variant, strArr As Variant

    '��ȡ��������
    StrInput = vbTab & g�������_��Ԫ����.��������
    StrInput = StrInput & vbTab & "0"
    Me.Caption = "ҽ����Ŀ����      ���ڴ�������ȡ������Ŀ����....."
    If ҵ������_��Ԫ����(��ȡ��������_����, StrInput, strOutput) = False Then Exit Sub
    
    gstrSQL = "select A.ID,E.��Ŀ���� as ����,F.���� as ���," & _
              "E.��Ŀ���� As ��������,'' as Ӣ������, " & _
              "zlspellcode(A.����) as ����,Substr(A.����,1,20) as ����,A.���㵥λ, " & _
              "substr(A.���,1,instr(A.���,'��')-1) as ���, " & _
              "A.�������� as �������,nvl(E.�Ƿ�ҽ��,0) as ���ñ�־ " & _
              "from �շ�ϸĿ A,ҽ��֧����Ŀ E,����֧������ F " & _
              "where " & _
              "nvl(A.����ʱ��,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
              "A.ID=E.�շ�ϸĿID And E.����=[1] And E.����=[2]" & _
              " And E.����ID=F.ID And E.����=F.���� and nvl(E.�Ƿ�ҽ��,0)=0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", mint����, mint����)
    
    i = 0
    intID = 0
    
    Do While Not rsTmp.EOF
       If rsTmp!���ñ�־ = 0 And intID <> rsTmp!ID Then
            i = i + 1
            
            StrInput = vbTab & g�������_��Ԫ����.��������
            StrInput = StrInput & vbTab & rsTmp!����
            
            If ҵ������_��Ԫ����(��ȡ��Ŀ_����, StrInput, strOutput) = False Then Exit Sub
            
            strArr = Split(strOutput, "@$")
            strTmpArr = Split(strArr(0), "||")
        
            If rsTmp!������� <> strTmpArr(4) Or rsTmp!��� <> strTmpArr(2) Then
                   
                   '���·������
                  '$IF HIS9.19
                  #If gverControl = 0 Then
                        gstrSQL = "ZL_�շ�ϸĿ_UPDATE_����(" & rsTmp!ID & ",'" & rsTmp!������� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "���·������")
                  #Else
                  '$ELSE  HIS+
                        gstrSQL = "ZL_�շ���ĿĿ¼_UPDATE_����(" & rsTmp!ID & ",'" & rsTmp!������� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "���·������")
                  #End If
            End If
                      
            gstrSQL = "select nvl(ID,0) as ID from ����֧������ where ����=[1] And ����=[2]"
            Set rsTmp1 = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ֧������", mint����, CStr(strTmpArr(2)))
            int����id = rsTmp1!ID
        
            gstrSQL = "ZL_ҽ��֧����Ŀ_Modify(" & rsTmp!ID & "," & mint���� & "," & mint���� & "," & _
                      int����id & ",'" & rsTmp!���� & "','" & rsTmp!�������� & "','" & Format(zlDatabase.Currentdate, "YYYY-MM-DD") & "'," & IIf(strTmpArr(1) = "����", 1, 0) & ")"
            ExecuteProcedure_��Ԫ���� "����ҽ��֧����Ŀ"
            
            Me.Caption = "ҽ����Ŀ����      ���ڲ�ѯ��" & i & "��δ������Ŀ�����Ժ�....."
            If strTmpArr(1) <> "����" Then
               MsgBox "��Ŀ��" & rsTmp!���� & "��" & rsTmp!�������� & "��������δ���á�"
               Exit Sub
            End If
       End If
       intID = rsTmp!ID
       rsTmp.MoveNext
    Loop
    
    MsgBox "����δ�걨��Ŀ�Ĳ�ѯ�Ѿ�ȫ����ɡ�"
End Sub

Private Sub mnuSB_Click()
    Dim i As Integer
    Dim strSQL As String, intID As Integer
    Dim StrInput As String, strOutput As String
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim int����id As Integer
    
    gstrSQL = "select A.ID,A.����,decode(A.���,'J','����','1','����','5','ҩƷ','6','ҩƷ','7','ҩƷ','����') as ���," & _
              "A.���� As ��������,'' as Ӣ������, " & _
              "zlspellcode(A.����) as ����,Substrb(A.����,1,40) as ����,substrb(A.���㵥λ,1,20) as ���㵥λ, " & _
              "B.�ּ�,substrb(substr(A.���,1,instr(A.���,'��')-1),1,20) as ���, " & _
              "D.���� as ������Ŀ,A.�������� as �������,E.��Ŀ���� " & _
              "from �շ�ϸĿ A,�շѼ�Ŀ B,������Ŀ D,ҽ��֧����Ŀ E " & _
              "where A.ID=B.�շ�ϸĿID and B.������ĿID=D.ID And " & _
              "nvl(B.��ֹ����,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
              "A.ID=E.�շ�ϸĿID(+) And E.����(+)=[1]And E.����(+)=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", mint����, mint����)
    
    i = 0
    intID = 0
    
    Do While Not rsTmp.EOF
       If IsNull(rsTmp!��Ŀ����) And intID <> rsTmp!ID Then
            i = i + 1
                        
            StrInput = vbTab & g�������_��Ԫ����.��������
            StrInput = StrInput & vbTab & rsTmp!���� & "||"
            StrInput = StrInput & rsTmp!��� & "||"
            StrInput = StrInput & rsTmp!�������� & "||"
            StrInput = StrInput & rsTmp!Ӣ������ & "||"
            StrInput = StrInput & rsTmp!���� & "||"
            StrInput = StrInput & rsTmp!���� & "||"
            StrInput = StrInput & rsTmp!���㵥λ & "||"
            StrInput = StrInput & rsTmp!�ּ� & "||"
            StrInput = StrInput & rsTmp!��� & "||"
            StrInput = StrInput & rsTmp!������Ŀ & "||"
            StrInput = StrInput & rsTmp!�������
            
            StrInput = StrInput & vbTab & gstrUserName
            StrInput = StrInput & vbTab & Format(zlDatabase.Currentdate, "YYYY-M-DD")
            
            If ҵ������_��Ԫ����(�걨��Ŀ_����, StrInput, strOutput) = False Then Exit Sub
            
            gstrSQL = "select nvl(ID,0) as ID from ����֧������ where ����=[1] And ����=[2]"
            Set rsTmp1 = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ֧������", mint����, CStr(rsTmp!���))
            int����id = rsTmp1!ID
            
            gstrSQL = "ZL_ҽ��֧����Ŀ_Modify(" & rsTmp!ID & "," & mint���� & "," & mint���� & "," & _
                       int����id & ",'" & rsTmp!���� & "','" & rsTmp!���� & "','" & Format(zlDatabase.Currentdate, "YYYY-MM-DD") & "',0)"
            ExecuteProcedure_��Ԫ���� "����ҽ��֧����Ŀ"
            Me.Caption = "ҽ����Ŀ����      �����ϴ���" & i & "���걨��Ŀ�����Ժ�....."
       End If
       intID = rsTmp!ID
       rsTmp.MoveNext
    Loop
    
    MsgBox "����δ�걨��Ŀ�Ѿ�ȫ���ϴ���ɡ�"
    
End Sub

Private Sub mnuViewFind_Click()
    If cmdSave.Enabled = True Then
        MsgBox "ҽ����Ŀ�б������ڱ༭״̬������ʹ�ò��ҹ��ܡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    frm������Ŀ���ҹ�Ԫ����.Show vbModal, Me
End Sub

Private Sub cmb����_Click()
    Call Fill����
    Call FillSum(False)
End Sub

Private Sub mnuViewClass_Click()
    mnuViewItem.Checked = False
    mnuViewClass.Checked = Not mnuViewClass.Checked
    Call SetSkip
End Sub

Private Sub mnuViewItem_Click()
    mnuViewClass.Checked = False
    mnuViewItem.Checked = Not mnuViewItem.Checked
    Call SetSkip
End Sub

Private Sub SetSkip()
'���ñ�����Ծ����
    With mshSum_S
        If mnuViewItem.Checked = False Then
        
            .ColData(COLҽ������) = 1
            .LocateCol = COLҽ������
            .ColData(col��������) = IIf(mnuViewClass.Checked = True, 5, 3)
        Else
            .ColData(col��������) = 3 'ѡ����
            .LocateCol = col��������
            .ColData(COLҽ������) = 5
        End If
        If .ColData(.COL) = 5 Then
            '��ǰ���Ѿ�����ѡ�������¶�λ
            .COL = .LocateCol
        End If
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    'ֻˢ���б�����
    Call FillSum
End Sub

Private Sub mshSum_S_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    'ʼ���ǲ�����ɾ����
    Cancel = True
    
    With mshSum_S
        If .TextMatrix(Row, col�ı䷽ʽ) = "����" Then
            .TextMatrix(Row, col�ı䷽ʽ) = "" '�൱��ʲô��û����
        Else
            .TextMatrix(Row, col�ı䷽ʽ) = "ɾ��" '���
        End If
        
        .TextMatrix(Row, COLҽ������) = ""
        .TextMatrix(Row, colҽ������) = ""
        .TextMatrix(Row, colҽ������) = ""
        .TextMatrix(Row, colҽ����ע) = ""
        .TextMatrix(Row, col����ID) = ""
        .TextMatrix(Row, col��������) = ""
        .TextMatrix(Row, col��ҽ��) = ""
        .TextMatrix(Row, col��˱�־) = ""
    End With
    cmdSave.Enabled = True
    cmdRestore.Enabled = True
End Sub

Private Sub mshSum_S_cboClick(ListIndex As Long)
    With mshSum_S
        If .COL <> col�������� Then Exit Sub
        
        If .TextMatrix(.Row, col��������) <> .CboText Then
            '��ֹ�޸ı��մ���,ֻ����ͨ��ѡ����ϸ��ȷ������
            If mint���� = TYPE_������ Then
                .ListIndex = mlngListIndex
                Exit Sub
            End If
            mlngListIndex = ListIndex
            .TextMatrix(.Row, col��������) = .CboText
            Call ��Ǹı�
        Else
            mlngListIndex = ListIndex
        End If
        
        If .CboText = "" Then
            '����Ϊ��
            .TextMatrix(.Row, col����ID) = ""
            .TextMatrix(.Row, col��������) = ""
        Else
            .TextMatrix(.Row, col����ID) = .ItemData(.ListIndex)
            .TextMatrix(.Row, col��������) = .CboText
        End If
        
    End With
End Sub

Private Sub mshSum_S_cboKeyDown(KeyCode As Integer, Shift As Integer)
    With mshSum_S
        If KeyCode = vbKeyReturn Then
            If .TextMatrix(.Row, col��������) <> .CboText Then
                .TextMatrix(.Row, col��������) = .CboText
                Call ��Ǹı�
            End If
            
            If .CboText = "" Then
                '����Ϊ��
                .TextMatrix(.Row, col����ID) = ""
                .TextMatrix(.Row, col��������) = ""
                .COL = col��ҽ��
            Else
                .TextMatrix(.Row, col����ID) = .ItemData(.ListIndex)
                .TextMatrix(.Row, col��������) = .CboText
            End If
        End If
    End With
End Sub

Private Sub mshSum_S_CommandClick()
'���ܣ���ȡҽ����Ŀ��ѡ��
'��������
'���أ�ҽ����Ŀ����
    Dim strCode As String
    Dim strSelected As String
    Dim STRNAME As String
    Dim strlastCode As String
    Dim strMemo As String
    
    With mshSum_S
        strCode = .TextMatrix(.Row, COLҽ������)
        If InitInfor_��Ԫ����.���õ��� = 0 And mint���� = TYPE_��Ԫ���� Then
            strCode = .TextMatrix(.Row, cOL����)
            If Frmҽ������_����.GetCode(strCode, mint����, mint����) = True Then
               strSelected = strCode
            End If
        Else
            If frm������Ŀѡ���Ԫ����.GetCode(strCode, mint����, mint����) = True Then
               strSelected = strCode
            End If
        End If
        
        If strSelected <> "" Then
            .TextMatrix(.Row, COLҽ������) = strSelected
            If STRNAME = "" Then
                Call Get��������
            Else
                '�Ѿ��������ƣ��Ͳ����ٵ���
                .TextMatrix(.Row, colҽ������) = STRNAME
                .TextMatrix(.Row, colҽ����ע) = ""
                .TextMatrix(.Row, col��ҽ��) = ""
            End If
            Call ��Ǹı�
        End If
    End With
End Sub

Private Sub mshSum_S_DblClick(Cancel As Boolean)
    With mshSum_S
        If .Active = False Then Exit Sub
        Call ��Ǹı�
    End With
End Sub

Private Sub mshSum_S_EnterCell(Row As Long, COL As Long)
    Static lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    
    If COL = col�������� And Trim(mshSum_S.TextMatrix(Row, COL)) = "" Then
        mshSum_S.ListIndex = -1
    End If
End Sub

Private Sub mshSum_S_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '������Ŀ����
    Dim strǰ As String, strText As String, str���� As String
    Dim rsTemp As New ADODB.Recordset, blnReturn As Boolean
    Dim strLeft As String
    Dim strTemp As String

    strǰ = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0") = "0", "%", "") '˫��ƥ��
    
    On Error GoTo errHandle
    
    With mshSum_S
        If .COL <> COLҽ������ Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If mint���� = TYPE_��Ԫ���� And InitInfor_��Ԫ����.���õ��� = 0 Then
                If strText = "" Then strText = .TextMatrix(.Row, cOL����)
                If Frmҽ������_����.GetCode(strText, mint����, mint����) = True Then blnReturn = strText
            Else
                If .TxtVisible = True Then
                    strText = Replace(Trim(.Text), "`", "")
                    .Text = strText
                    If zlCommFun.StrIsValid(strText, mlng���볤��) = False Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Trim(strText) = "" Then
                        '����Ҫ��ȥ����Ƿ���ƥ��ı��룬�൱��ɾ���ñ���
                        .TextMatrix(.Row, COLҽ������) = Trim(strText)
                    Else
                        '����SQL���
                        If mint���� = TYPE_�ɶ����� Then
                            gstrSQL = "Select ����  ҽ������,����,����,��ע " & _
                                      " FROM ҽ���շ���Ŀ_���� WHERE " & _
                                      " ����=[1] and ����=[2] and (���� like [3] || '%' or Upper(����) like [3] || '%' Or Upper(����) like [3] || '%')"
                        Else
                            gstrSQL = "Select ����  ҽ������,����,����,��ע " & _
                                         "   FROM ҽ���շ���Ŀ WHERE " & _
                                      " ����=[1] and ����=[2] and (���� like [3] || '%' or Upper(����) like [3] || '%' Or Upper(����) like [3] || '%')"
                        End If
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, mint����, strText)
    
                        If rsTemp.RecordCount > 0 Then
                            '����ѡ����
                            If rsTemp.RecordCount >= 1 Or rsTemp.Fields.Count > 3 Then
                                '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
                                blnReturn = frmListSel.ShowSelect(mint����, rsTemp, "ҽ������", "ҽ����Ŀѡ��", "��ѡ���Ӧ��ҽ����Ŀ��")
                            End If
                        End If
                        
                        If blnReturn = False Then
                            '��¼����û�п�ѡ�������
                            If rsTemp.RecordCount > 0 Then
                                '��¼�������ݣ���ȡ����ѡ��
                                Cancel = True
                                .TxtVisible = True
                                .TxtSetFocus
                                Exit Sub
                            Else
                                .Text = strText
                                .TextMatrix(.Row, COLҽ������) = strText
                            End If
                        Else
                            .Text = rsTemp("ҽ������")
                            .TextMatrix(.Row, COLҽ������) = rsTemp("ҽ������")
                        End If
                    End If
                    Call Get��������
                    Call ��Ǹı�
                End If
            End If
        Else
            If .TextMatrix(.Row, COLҽ������) = "" Then
                .TextMatrix(.Row, COLҽ������) = " "
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Cancel = True
End Sub

Private Sub ��Ǹı�()
    '��ǰ�����Ѿ���Ч���������ܷ�õ���������
    If mint���� = TYPE_��Ԫ���� And InitInfor_��Ԫ����.���õ��� = 0 Then
        cmdRestore.Enabled = False
        cmdSave.Enabled = False
    Else
        cmdRestore.Enabled = True
        cmdSave.Enabled = True
    End If
    
    With mshSum_S
        If Trim(.TextMatrix(.Row, COLҽ������)) = "" And Trim(.TextMatrix(.Row, col��������)) = "" Then
            .TextMatrix(.Row, col�ı䷽ʽ) = "ɾ��"
        Else
            If Trim(.TextMatrix(.Row, col�ı䷽ʽ)) <> "�޸�" Then
                'Ϊ�գ����Ѿ��ǡ�������
                .TextMatrix(.Row, col�ı䷽ʽ) = "����"
            End If
        End If
    End With
End Sub

Private Sub Get��������()
'���ܣ����ݵ�ǰ�еı�����Ŀ���룬�õ�������Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long, lngPos As Long
    Dim str������� As String, strTemp As String, varPart As Variant
    
    On Error GoTo errHandle
    With mshSum_S
        If mint���� = TYPE_�ɶ����� Then
            gstrSQL = "select ����,�������,��ע from ҽ���շ���Ŀ_���� where ����=[1] and ����=[2] and ����=[3]"
        Else
            gstrSQL = "select ����,�������,��ע from ҽ���շ���Ŀ where ����=[1] and ����=[2] and ����=[3]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(.TextMatrix(.Row, COLҽ������)), mint����, CLng(cmb����.ItemData(cmb����.ListIndex)))
        
        If rsTemp.RecordCount = 0 Then
            'û�ж�Ӧ�ı�����Ŀ��ֻ�����øñ���
            .TextMatrix(.Row, colҽ������) = ""
            .TextMatrix(.Row, colҽ����ע) = ""
            .TextMatrix(.Row, col��ҽ��) = ""
        Else
            .TextMatrix(.Row, colҽ������) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(.Row, colҽ����ע) = IIf(IsNull(rsTemp("��ע")), "", rsTemp("��ע"))
            str������� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
        End If
        For lngIndex = 0 To .ListCount - 1
            lngPos = InStr(.List(lngIndex), ".")
            If lngPos = 0 Then
                strTemp = .List(lngIndex)
            Else
                strTemp = Mid(.List(lngIndex), 1, lngPos - 1)
            End If
            If strTemp = str������� Then
                '�ҵ���ƥ��Ĵ������
                .TextMatrix(.Row, col����ID) = .ItemData(lngIndex)
                .TextMatrix(.Row, col��������) = .List(lngIndex)
                Exit For
            End If
        Next
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mshSum_S_KeyPress(KeyAscii As Integer)
    With mshSum_S
        If Not .Active Then Exit Sub
        If .ColData(.COL) = -1 Then Call ��Ǹı�
    End With
End Sub

Private Sub mshSum_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mshSum_S.ToolTipText = mshSum_S.TextMatrix(mshSum_S.MouseRow, mshSum_S.MouseCol)
End Sub

Private Sub mshSum_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rsTemp As New ADODB.Recordset, lngID As Long
    Dim lngRow As Long, lngPos As Long, blnActive As Boolean
    Dim blnEnable As Boolean
    
    If mshSum_S.Active = False Then Exit Sub
    If mshSum_S.MouseRow = 0 Then
        If mlngCol = mshSum_S.MouseCol Then
            mblnDesc = Not mblnDesc
        Else
            mlngCol = mshSum_S.MouseCol
            mblnDesc = False
        End If
        
        blnEnable = cmdRestore.Enabled
        blnActive = mshSum_S.Active
        mshSum_S.Active = False
        mshSum_S.msfObj.MousePointer = vbHourglass
        
        '���ɼ�¼����Ȼ��ˢ�±��
        rsTemp.CursorLocation = adUseClient
        rsTemp.CursorType = adOpenDynamic
        rsTemp.LockType = adLockOptimistic
        With rsTemp.Fields
            .Append "ID", adDouble, adFldIsNullable
            .Append "����", adVarChar, 20, adFldIsNullable
            .Append "����", adVarChar, 50, adFldIsNullable
            .Append "���", adVarChar, 80, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "��λ", adVarChar, 20, adFldIsNullable
            .Append "�Ƿ���", adInteger, adFldIsNullable
            .Append "�۸�", adVarNumeric, 20, adFldIsNullable
            .Append "�ı䷽ʽ", adVarChar, 4, adFldIsNullable
            'Modified By ���� 2003-12-09 ��������ɽ
            .Append "��Ŀ����", adVarChar, 50, adFldIsNullable
            .Append "��Ŀ����", adVarChar, 50, adFldIsNullable
            .Append "��ע", adVarChar, 50, adFldIsNullable
            .Append "ԭ����", adVarChar, 20, adFldIsNullable
            .Append "�Ƿ�ҽ��", adInteger
            .Append "����ID", adDouble
            .Append "�������", adVarChar, 10, adFldIsNullable
            .Append "��������", adVarChar, 50, adFldIsNullable
        End With
        
        rsTemp.Open
        With mshSum_S
            For lngRow = 1 To .Rows - 1
                rsTemp.AddNew
                
                rsTemp("ID") = .RowData(lngRow)
                rsTemp("����") = .TextMatrix(lngRow, cOL����)
                rsTemp("����") = .TextMatrix(lngRow, cOL����)
                rsTemp("���") = .TextMatrix(lngRow, COL���)
                rsTemp("����") = .TextMatrix(lngRow, COL����)
                rsTemp("����") = Substr(.TextMatrix(lngRow, col����), 1, 100)
                rsTemp("��λ") = .TextMatrix(lngRow, COL��λ)
                If .TextMatrix(lngRow, col�۸�) = "" Then
                    rsTemp("�Ƿ���") = 1
                    rsTemp("�۸�") = 0
                Else
                    rsTemp("�Ƿ���") = 0
                    rsTemp("�۸�") = Val(.TextMatrix(lngRow, col�۸�))
                End If
                rsTemp("�ı䷽ʽ") = .TextMatrix(lngRow, col�ı䷽ʽ)
                rsTemp("��Ŀ����") = .TextMatrix(lngRow, COLҽ������)
                rsTemp("��Ŀ����") = .TextMatrix(lngRow, colҽ������)
                rsTemp("��ע") = .TextMatrix(lngRow, colҽ����ע)
                rsTemp("ԭ����") = .TextMatrix(lngRow, colԭ����)
                rsTemp("����ID") = Val(.TextMatrix(lngRow, col����ID))
                rsTemp("�Ƿ�ҽ��") = IIf(.TextMatrix(lngRow, col��ҽ��) = "��", 0, 1)
                
                
                lngPos = InStr(.TextMatrix(lngRow, col��������), ".")
                If lngPos = 0 Then
                    rsTemp("�������") = Null
                    rsTemp("��������") = Null
                Else
                    rsTemp("�������") = Mid(.TextMatrix(lngRow, col��������), 1, lngPos - 1)
                    rsTemp("��������") = Mid(.TextMatrix(lngRow, col��������), lngPos + 1)
                End If
                
                rsTemp.Update
            Next
            lngID = .RowData(.Row)
        End With
        Call FillGrid(rsTemp, lngID)
    
        mshSum_S.Active = blnActive '�ָ�
        mshSum_S.msfObj.MousePointer = vbDefault
        MousePointer = vbDefault
        cmdRestore.Enabled = blnEnable
        cmdSave.Enabled = blnEnable
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    'ֻˢ���б�����
    FillSum
End Sub

Private Sub mshSum_S_GotFocus()
    Call MenuSet
End Sub

Private Sub mshSum_S_LostFocus()
    mshSum_S.CmdVisible = False
    mshSum_S.CboVisible = False
    
    Call MenuSet
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picV.Left + x - msngStartX
        If sngTemp > 1500 And ScaleWidth - (sngTemp + picV.Width) > 1600 Then
            picV.Left = sngTemp
            tvwMain_S.Width = picV.Left - tvwMain_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Find"
            mnuViewFind_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwMain_S.SelectedItem
    Do Until nod.Parent Is Nothing
        Set nod = nod.Parent
    Loop
    
    Set objPrint.Body = mshSum_S.msfObj
    objPrint.Title.Text = nod.Text & "���շ�ϸĿҽ����Ŀ��Ӧ��"
    'objRow.Add "ҽԺ���ƣ�" & gstr��λ����
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & gstrUserName
    objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
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
End Sub
    
Private Sub Fill����()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    'ֻˢ���б�����
    
    '���Ȼ��ҽ������
    mshSum_S.Active = True
    If mint���� = TYPE_�ɶ��ϳ� Then
        If mcnYB.State = 1 Then mcnYB.Close
        mcnYB.Open GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LCConnectionString"), "dsn=lcyb;uID=hisuser;pwd=hiscdgk")
        Exit Sub
    End If
    
    gstrSQL = "select ID,����,���� from ����֧������ " & _
              "where ����=[1] order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    
    mshSum_S.Clear
    Do Until rsTemp.EOF
        mshSum_S.AddItem rsTemp("����") & "." & rsTemp("����")
        mshSum_S.ItemData(mshSum_S.NewIndex) = rsTemp("ID")
        rsTemp.MoveNext
    Loop
    If mint���� = TYPE_�ɶ����� Then
        If Not gcnOracle_�ɶ����� Is Nothing Then Exit Sub
    ElseIf mint���� = TYPE_�ϳ����� Then
        If Not gcnOracle_�ϳ����� Is Nothing Then Exit Sub
    Else
        If Not gcnOracle_��Ԫ���� Is Nothing Then Exit Sub
    End If
    '�����´�ҽ��
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & mint����
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ��������"
                strServer = strTemp
            Case "ҽ���û���"
                strUser = strTemp
            Case "ҽ���û�����"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If mint���� = TYPE_�ɶ����� Then
        Set gcnOracle_�ɶ����� = New ADODB.Connection
        If OraDataOpen(gcnOracle_�ɶ�����, strServer, strUser, strPass) = False Then
            Exit Sub
        End If
    ElseIf mint���� = TYPE_�ϳ����� Then
        Set gcnOracle_�ϳ����� = New ADODB.Connection
        If OraDataOpen(gcnOracle_�ϳ�����, strServer, strUser, strPass) = False Then
            Exit Sub
        End If
    Else
        Set gcnOracle_��Ԫ���� = New ADODB.Connection
        If OraDataOpen(gcnOracle_��Ԫ����, strServer, strUser, strPass) = False Then
            Exit Sub
        End If
    End If
End Sub

Private Function FillTree() As Boolean
'����:װ���շ������շ�ϸĿ�����з��ൽtvwMain_S
    '�����������ڵ�����������KEYֵ��һ���ַ������ڶ�λ��������

    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim nod As Node
    
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    
    MousePointer = vbHourglass
    
    mstrKey = ""     'ȫ��ˢ��ʱ���൱���û�û����κνڵ�
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    gstrSQL = "select ����,��� from �շ���� where ����<>'5' and ����<>'6' and ����<>'7' order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    LockWindowUpdate tvwMain_S.hwnd
    'ɾ�����нڵ�
    With tvwMain_S.Nodes
        .Clear
        '�������
        Do Until rsTemp.EOF
            .Add , , "R" & rsTemp("����"), "��" & rsTemp("����") & "��" & rsTemp("���"), "R", "R"
            tvwMain_S.Nodes("R" & rsTemp("����")).Sorted = True
            rsTemp.MoveNext
        Loop
        .Add , , "D5", "��5������ҩ", "R", "R"
        tvwMain_S.Nodes("D5").Sorted = True
        .Add , , "E6", "��6���г�ҩ", "R", "R"
        tvwMain_S.Nodes("E6").Sorted = True
        .Add , , "F7", "��7���в�ҩ", "R", "R"
        tvwMain_S.Nodes("F7").Sorted = True
        
        '������ͨ�շ���Ŀ����ڵ�
        gstrSQL = "select id,�ϼ�id,���,����,���� from �շ�ϸĿ  where ���<>'5' and ���<>'6' and ���<>'7' and ĩ�� <> 1 " & _
             " start with �ϼ�ID is null  connect by prior id=�ϼ�ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
        Do Until rsTemp.EOF
            '��ӽڵ�
            If IsNull(rsTemp("�ϼ�id")) Then
                .Add "R" & rsTemp("���"), tvwChild, "C" & rsTemp("���") & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "C", "C"
            Else
                .Add "C" & rsTemp("���") & rsTemp("�ϼ�id"), tvwChild, "C" & rsTemp("���") & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "C", "C"
            End If
            tvwMain_S.Nodes("C" & rsTemp("���") & rsTemp("ID")).Sorted = True
            rsTemp.MoveNext
        Loop
    
        '��װ��ҩƷ��;���������
        gstrSQL = "select id,�ϼ�id,����,����,���� from ҩƷ��;����  " & _
             " start with �ϼ�ID is null  connect by prior id=�ϼ�ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
        Do Until rsTemp.EOF
            '��ӽڵ�
            Select Case rsTemp("����")
                Case "�г�ҩ"
                    If IsNull(rsTemp("�ϼ�id")) Then
                        Set nod = .Add("E6", tvwChild, "E6" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    Else
                        Set nod = .Add("E6" & rsTemp("�ϼ�id"), tvwChild, "E6" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    End If
                Case "�в�ҩ"
                    If IsNull(rsTemp("�ϼ�id")) Then
                        Set nod = .Add("F7", tvwChild, "F7" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    Else
                        Set nod = .Add("F7" & rsTemp("�ϼ�id"), tvwChild, "F7" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    End If
                Case Else '����ҩ
                    If IsNull(rsTemp("�ϼ�id")) Then
                        Set nod = .Add("D5", tvwChild, "D5" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    Else
                        Set nod = .Add("D5" & rsTemp("�ϼ�id"), tvwChild, "D5" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
                    End If
                End Select
            nod.Sorted = True
            rsTemp.MoveNext
        Loop
    End With
    
    LockWindowUpdate 0
    MousePointer = 0
    
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes(1)
        nod.Selected = True
    Else
        Err.Clear
        nod.Selected = True
        nod.EnsureVisible
    End If
    Call FillSum
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockWindowUpdate 0
    MousePointer = 0
End Function

Public Sub FillSum(Optional ByVal blnForce As Boolean = False)
'����:װ�����ͳ������
    Dim rsTemp As New ADODB.Recordset
    Dim nod As Node
    Dim str���ʷ��� As String
    Dim lngID As Long

    If tvwMain_S.SelectedItem Is Nothing Then
        ClearGrid mshSum_S
        Call MenuSet
        Exit Sub
    End If
    
    If blnForce = False Then
        If mstrKey = tvwMain_S.SelectedItem.Key And mint���� = cmb����.ItemData(cmb����.ListIndex) Then
            '��ȫû�иı䣬������ˢ��
            Exit Sub
        End If
        
        If cmdSave.Enabled = True Then
            '�Ѿ��޸ģ���ʾ�Ƿ���Ҫ���浱ǰ������
            If MsgBox("������Ŀ�Ѿ��޸ģ��Ƿ���Ҫ���棿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                Call cmdSave_Click
            End If
        End If
    End If
    
    cmdSave = False
    cmdRestore = False
    '��ȡ������Ŀ���������ؼ���
    mstrKey = tvwMain_S.SelectedItem.Key
    mint���� = cmb����.ItemData(cmb����.ListIndex)
    
    Set nod = tvwMain_S.SelectedItem
    
    '���ݲ�ͬ�Ľڵ㣬������ͬ����ʾ
    '�������Ҫ����ʾһ��
    If Mid(nod.Key, 2, 1) = "5" Or Mid(nod.Key, 2, 1) = "6" Or Mid(nod.Key, 2, 1) = "7" Then
        'ҩƷ�Ĵ���Ҫ�鷳һЩ
        mshSum_S.TextMatrix(0, col����) = "����"
        
        Select Case Left(nod.Key, 1)
            Case "D"
                str���ʷ��� = "����ҩ"
            Case "E"
                str���ʷ��� = "�г�ҩ"
            Case "F"
                str���ʷ��� = "�в�ҩ"
        End Select
        
        If nod.Image = "R" Then
            gstrSQL = "select A.ҩƷID as ID,A.����,B.ͨ������||decode(M.����,null,'',b.ͨ������,'',' ��'||M.����||'��') as ����,A.���,A.����,A.�ۼ۵�λ as ��λ,D.�Ƿ���,E.���� ���� " & _
                        "from ҩƷĿ¼ A,ҩƷ��Ϣ B,�շ�ϸĿ D,ҩƷ���� E,(Select distinct ҩƷid,���� from ҩƷ���� ) M " & _
                        "where A.ҩ��ID=B.ҩ��ID and d.id=M.ҩƷID(+) and B.����=E.����(+) and B.���ʷ���='" & str���ʷ��� & "'" & _
                        "      and A.ҩƷID=D.ID and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "select A.ҩƷID as ID,A.����,B.ͨ������||decode(M.����,null,'',b.ͨ������,'',' ��'||M.����||'��') as ����,A.���,A.����,A.�ۼ۵�λ as ��λ,D.�Ƿ���,E.���� ���� " & _
                      "from ҩƷĿ¼ A,ҩƷ��Ϣ B,�շ�ϸĿ D,ҩƷ���� E,(Select distinct ҩƷid,���� from ҩƷ����) M ,(select ID from ҩƷ��;���� start with ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID) C " & _
                      "where A.ҩ��ID=B.ҩ��ID and B.����=E.����(+) and d.id=M.ҩƷID(+) and B.���ʷ���='" & str���ʷ��� & "' and B.��;����ID=C.ID" & _
                      "       and A.ҩƷID=D.ID and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))"
        End If
        
    Else
        '��ҩƷ�����׵ö���
        mshSum_S.TextMatrix(0, col����) = "˵��"
        
        If nod.Image = "R" Then
            gstrSQL = "select id,����,����,���,˵�� as ����,���㵥λ as ��λ,�Ƿ���,'' ���� from �շ�ϸĿ where ĩ��=1 and ���='" & Mid(nod.Key, 2, 1) & "' " & _
                        " and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "select id,����,����,���,˵�� as ����,���㵥λ as ��λ,�Ƿ���,'' ���� from �շ�ϸĿ where ĩ��=1 and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))" & _
                        " start with �ϼ�ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID "
        End If
    End If
    Dim strTable As String
    If mint���� = TYPE_�ɶ����� Then
        strTable = "ҽ��֧����Ŀ_����"
    Else
        strTable = "ҽ��֧����Ŀ"
    End If
    gstrSQL = "select A.ID,A.����,A.����,A.���,A.����,A.����,A.��λ,A.�Ƿ���,D.�۸�,'' as �ı䷽ʽ" & _
               " ,B.��Ŀ����,B.��Ŀ����,B.��ע,B.��Ŀ���� as ԭ����,B.�Ƿ�ҽ��,B.����ID,C.���� as �������,C.���� as �������� " & _
               " from (" & gstrSQL & ") A," & strTable & " B,����֧������ C," & _
               "      (select sum(�ּ�) as �۸�,�շ�ϸĿID from �շѼ�Ŀ where ִ������<=sysdate and (��ֹ����>=sysdate or ��ֹ���� is null) group by �շ�ϸĿID) D " & _
               " Where A.ID=B.�շ�ϸĿID(+) and B.����ID=c.id(+)   and b.����(+)=[1] and  B.����(+)= [2]" & _
               "       and A.ID=D.�շ�ϸĿID(+)  "
    
    MousePointer = 11
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, mint����)
    
    lngID = mshSum_S.RowData(mshSum_S.Row)
    Call FillGrid(rsTemp, lngID)
    
    stbThis.Panels(2).Text = "�����շ���Ŀ" & rsTemp.RecordCount & "��"
    
    MousePointer = 0
    Call MenuSet
End Sub

Private Sub FillGrid(rsTemp As ADODB.Recordset, ByVal lngID As Long)
    Dim strSort As String
    Dim strDemo As String
    Dim intMatch As Integer
    Dim lngRow As Long, lngRowSelect As Long
    
    Select Case mlngCol
        Case cOL����
            strSort = "����"
        Case cOL����
            strSort = "����"
        Case COL���
            strSort = "���"
        Case col����
            strSort = "����"
        Case COL��λ
            strSort = "��λ"
        Case col�۸�
            strSort = "�۸�"
        Case COLҽ������
            strSort = "��Ŀ����"
        Case colҽ������
            strSort = "��Ŀ����"
        Case col��������
            strSort = "��������"
        Case col��ҽ��
            strSort = "�Ƿ�ҽ��"
        Case Else
            rsTemp.Sort = "����"
    End Select
    rsTemp.Sort = strSort & IIf(mblnDesc, " DESC", "")
    
    mshSum_S.TxtVisible = False
    mshSum_S.CboVisible = False
    mshSum_S.Redraw = False
    ClearGrid mshSum_S
    If rsTemp.RecordCount <> 0 Then
        mshSum_S.Rows = rsTemp.RecordCount + 1
    End If
    lngRow = 1
    With mshSum_S
        Do Until rsTemp.EOF
            If rsTemp("ID") = lngID Then
                lngRowSelect = lngRow
            End If
            
            .RowData(lngRow) = rsTemp("ID")
            .TextMatrix(lngRow, cOL����) = rsTemp("����")
            .TextMatrix(lngRow, cOL����) = rsTemp("����")
            .TextMatrix(lngRow, COL���) = IIf(IsNull(rsTemp("���")), "", rsTemp("���"))
            .TextMatrix(lngRow, col����) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, COL����) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, COL��λ) = IIf(IsNull(rsTemp("��λ")), "", rsTemp("��λ"))
            .TextMatrix(lngRow, col�۸�) = IIf(rsTemp("�Ƿ���") = 0, Format(rsTemp("�۸�"), "0.000"), "")
            .TextMatrix(lngRow, col�ı䷽ʽ) = IIf(IsNull(rsTemp("�ı䷽ʽ")), "", rsTemp("�ı䷽ʽ"))
            .TextMatrix(lngRow, COLҽ������) = IIf(IsNull(rsTemp("��Ŀ����")), "", rsTemp("��Ŀ����"))
            .TextMatrix(lngRow, colҽ������) = IIf(IsNull(rsTemp("��Ŀ����")), "", rsTemp("��Ŀ����"))
            .TextMatrix(lngRow, colԭ����) = IIf(IsNull(rsTemp("ԭ����")), "", rsTemp("ԭ����"))
            .TextMatrix(lngRow, col����ID) = IIf(IsNull(rsTemp("����ID")), "", rsTemp("����ID"))
            .TextMatrix(lngRow, col��ҽ��) = IIf(rsTemp("�Ƿ�ҽ��") = "0", "��", "")
            .TextMatrix(lngRow, colҽ����ע) = IIf(IsNull(rsTemp("��ע")), "", rsTemp("��ע"))
            .TextMatrix(lngRow, colƥ�����к�) = ""
            
            If IsNull(rsTemp("�������")) Or IsNull(rsTemp("��������")) Then
                .TextMatrix(lngRow, col��������) = ""
            Else
                .TextMatrix(lngRow, col��������) = rsTemp("�������") & "." & rsTemp("��������")
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    If lngRowSelect > 0 And lngRowSelect < mshSum_S.Rows - 1 Then
        mshSum_S.msfObj.TopRow = lngRowSelect
        mshSum_S.Row = lngRowSelect
    End If
    mshSum_S.Redraw = True
End Sub

Private Sub ClearGrid(objGrid As Object)
'���ܣ�������,����ɲ��ֳ�ʼ��
    Dim i As Long
    
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
    With objGrid.msfObj
        .Rows = 2
        .RowData(1) = 0
        For i = 0 To objGrid.Cols - 1
            objGrid.TextMatrix(1, i) = ""
        Next
    
    End With
End Sub

Private Sub MenuSet()
'����:��ʾ�˵��͹�������״̬(��ӡ)
    Dim blnPrint As Boolean
    
    blnPrint = Not (mshSum_S.Rows = 2 And mshSum_S.TextMatrix(1, 0) = "")

    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    tbrThis.Buttons("Preview").Enabled = blnPrint
    tbrThis.Buttons("Print").Enabled = blnPrint
    
    If InStr(mstrȨ��, "��ɾ��") > 0 Then
        mshSum_S.Active = blnPrint
        If mint���� = TYPE_������ Then
            'ǿ�Ʋ���ʹ��
            If gcn����.State = adStateClosed Then mshSum_S.Active = False
        End If
    Else
        mshSum_S.Active = False
    End If
End Sub

Public Sub ShowForm(frmParent As Form, ByVal int���� As Integer)
    
    Dim rsTemp As New ADODB.Recordset
    mint���� = int����
    
    gstrSQL = "select ���,���� from ��������Ŀ¼ where ���<>0 and ����=[1] order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", mint����)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п����籣���������ڲ��������أ�����ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frm������Ŀ_��Ԫ����.Visible = True Then
        frm������Ŀ_��Ԫ����.Show
        Exit Sub
    End If
    
    mstrȨ�� = gstrPrivs
    frm������Ŀ_��Ԫ����.Show , frmParent
End Sub


Public Function CheckForm(ByVal int���� As Integer) As Boolean
    
    Dim rsTemp As New ADODB.Recordset
    mint���� = int����
    
    gstrSQL = "select ���,���� from ��������Ŀ¼ where ���<>0 and ����=[1] order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", mint����)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п����籣���������ڲ��������أ�����ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frm������Ŀ_��Ԫ����.Visible = True Then
        CheckForm = True
        Exit Function
    End If
    
    mstrȨ�� = gstrPrivs
    CheckForm = True
End Function
