VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmҽ���ʻ���֧���� 
   Caption         =   "�ʻ���֧����"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7470
   Icon            =   "frmҽ���ʻ���֧����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ImgColor 
      Left            =   240
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":06EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":0904
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":0B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":0E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":108A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":1810
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":1A2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBlack 
      Left            =   780
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":1C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":1E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":2078
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":2292
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":24AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":26C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":28E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":2AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ���֧����.frx":2D14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1270
      BandCount       =   2
      _CBWidth        =   7470
      _CBHeight       =   720
      _Version        =   "6.7.9782"
      Child1          =   "tbrTool"
      MinWidth1       =   3000
      MinHeight1      =   660
      Width1          =   2820
      NewRow1         =   0   'False
      BandForeColor2  =   -2147483646
      BandBackColor2  =   -2147483638
      Caption2        =   "�������"
      Child2          =   "cbo����"
      MinWidth2       =   1800
      MinHeight2      =   300
      Width2          =   1335
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   660
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgBlack"
         HotImageList    =   "ImgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Printview"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Adjust"
               Object.Tag             =   "����"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Single"
                     Text            =   "��������"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Batch"
                     Text            =   "��������"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.Tag             =   "�鿴"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.Tag             =   "�˳�"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   5580
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   1800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�ʻ��䶯��¼ 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7223
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   4860
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmҽ���ʻ���֧����.frx":2F2E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8096
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
            AutoSize        =   2
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
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdjust 
         Caption         =   "�ʻ�������(&A)"
         Begin VB.Menu mnuEditAdjust_Single 
            Caption         =   "��������(&S)"
         End
         Begin VB.Menu mnuEditAdjust_Batch 
            Caption         =   "��������(&B)"
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸ı䶯��¼(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ���䶯��¼(&D)"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "�鿴(&V)"
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
         Begin VB.Menu mnuViewTool_1 
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
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
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmҽ���ʻ���֧����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String                   '���Ҵ���ȱʡΪ���ҵ�һ��ҽ�����ĵĲ��˵��ʻ��䶯��¼������������ģ����������
Private lngCardRow As Long
Private mint���� As Integer
Private mblnLoad As Boolean
Private mstrPrivs As String
Private Const glng��ɫ As Long = &H80000005
Private Const glng��ɫ As Long = &H80000008
Private Const glng���ɫ As Long = &HC0C0C0
Private Const glng��ɫ As Long = &H8000000F
Private Const glng��ɫ As Long = &HE0E0E0
Private Const glng��ɫ As Long = &HC0
Private Const glng����ɫ As Long = &H8000000D

Private Enum ��Enum
    colID = 0
    col���� = 1
    col���� = 2
    colҽ���� = 3
    col����ID = 4
    col���� = 5
    col��� = 6
    col������ = 7
    colʱ�� = 8
    col���� = 9
    col˵�� = 10
    col���� = 11
End Enum

Private Sub cbo����_Click()
    With cbo����
        If mint���� = .ItemData(.ListIndex) Then Exit Sub
        mint���� = .ItemData(.ListIndex)
    End With
    
    '�������
    Call FillList
End Sub

Private Sub cbrTool_Resize()
    Call Form_Resize
End Sub

Private Sub InitBill(Optional ByVal bln���� As Boolean = True)
    '��ʼ�����
    Dim lngCol As Integer
    
    '���ø�ʽ
    With Msf�ʻ��䶯��¼
        .Clear
        .Rows = 2
        .Cols = col����
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
        Next
        
        .TextMatrix(0, colID) = "ID"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, colҽ����) = "ҽ����"
        .TextMatrix(0, col����ID) = "����ID"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col���) = "���"
        .TextMatrix(0, col������) = "������"
        .TextMatrix(0, colʱ��) = "ʱ��"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col˵��) = "˵��"
        If Not mblnLoad Then
            .ColWidth(colID) = 0
            .ColWidth(col����) = IIf(bln����, 1500, 0)
            .ColWidth(col����) = 900
            .ColWidth(colҽ����) = 900
            .ColWidth(col����ID) = 1000
            .ColWidth(col����) = 800
            .ColWidth(col���) = 1200
            .ColWidth(col������) = 800
            .ColWidth(colʱ��) = 1800
            .ColWidth(col����) = 0
            .ColWidth(col˵��) = 2000
            'Call RestoreFlexState(Msf�ʻ��䶯��¼, Me.Caption)
            .ColWidth(col����) = IIf(bln����, 1500, 0)
        End If
        For lngCol = 0 To .Cols - 1
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .COL = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub FillList()
    Dim str��ʼʱ�� As String, str����ʱ�� As String
    Dim bln���� As Boolean
    Dim rsAccount As New ADODB.Recordset
    '�������
    
    bln���� = ��������(mint����)
    
    '��ȡ�ʻ��䶯��¼�嵥������䵽�����
    If mstrFind = "" Then
        str��ʼʱ�� = Format(DateAdd("d", -1, zlDatabase.Currentdate), "yyyy-MM-dd HH:mm:ss")
        str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        mstrFind = " And B.����=1 And Trunc(B.ʱ��) " & _
                   " Between to_date('" & str��ʼʱ�� & "','yyyy-MM-dd hh24:mi:ss') " & _
                   " And to_date('" & str����ʱ�� & "','yyyy-MM-dd hh24:mi:ss') "
    End If
    If bln���� Then
        gstrSQL = "Select B.ID,D.���� ����,A.����,A.ҽ����,C.����ID,C.����,ltrim(to_char(B.���,'900090000.00')) ���,B.������, " & _
                 " To_char(B.ʱ��,'yyyy-MM-dd hh24:mi:ss') ʱ��,����,˵�� " & _
                 " From �����ʻ� A,�ʻ��䶯��¼ B,������Ϣ C,��������Ŀ¼ D " & _
                 " Where A.����=B.���� And A.����ID=B.����ID And A.����ID=C.����ID  " & _
                 " And A.����=D.���� And A.����=D.��� And A.����=" & mint���� & mstrFind
    Else
        gstrSQL = "Select B.ID,'' ����,A.����,A.ҽ����,C.����ID,C.����,ltrim(to_char(B.���,'900090000.00')) ���,B.������, " & _
                 " To_char(B.ʱ��,'yyyy-MM-dd hh24:mi:ss') ʱ��,����,˵�� " & _
                 " From �����ʻ� A,�ʻ��䶯��¼ B,������Ϣ C " & _
                 " Where A.����=B.���� And A.����ID=B.����ID And A.����ID=C.����ID  " & _
                 " And Nvl(A.����,0)=0 And A.����=" & mint���� & mstrFind
    End If
    gstrSQL = gstrSQL & " Order by ����,����,ʱ��"
    Call OpenRecordset(rsAccount, "��ȡ�ʻ��䶯��¼�嵥")
    Call InitBill(bln����)
    If Not rsAccount.EOF Then
        Set Msf�ʻ��䶯��¼.DataSource = rsAccount
        Msf�ʻ��䶯��¼.ColWidth(col����) = 0
        '�������н�����ɫ����
        Call SetItemColor
    End If
    
    Dim lngCol As Long
    For lngCol = 0 To Msf�ʻ��䶯��¼.Cols - 1
        Msf�ʻ��䶯��¼.ColAlignmentFixed(lngCol) = 4
    Next
    
    '�Ƚ��˵��빤��������Ϊ�ң���ͨ������EnterCell�����ݵ�ǰ�е�״̬���ò˵���������
    Call SetMenu
    If rsAccount.RecordCount <> 0 Then
        Msf�ʻ��䶯��¼.Row = 1
        Call Msf�ʻ��䶯��¼_EnterCell
    End If
End Sub

Private Sub SetItemColor()
    Dim lngRow As Long, lngCol As Long, lngColor As Long
    Dim lngSaveRow As Long, lngSaveCol As Long
    On Error Resume Next
    
    With Msf�ʻ��䶯��¼
        .Redraw = False
        lngSaveRow = .Row: lngSaveCol = .COL
        For lngRow = 1 To .Rows - 1
            .Row = lngRow
            Select Case .TextMatrix(.Row, col����)
            Case 1
                lngColor = glng��ɫ
            Case 2
                lngColor = glng����ɫ
            Case Else
                lngColor = glng��ɫ
            End Select
            
            For lngCol = 0 To .Cols - 1
                .COL = lngCol
                .CellForeColor = lngColor
            Next
        Next
        .Row = lngSaveRow: .COL = lngSaveCol
        .Redraw = True
    End With
End Sub

Private Sub SetMenu(Optional ByVal blnState As Boolean = False)
    '���ò˵�״̬
    mnuEditAdjust.Enabled = True
    mnuEditModify.Enabled = blnState
    mnuEditDelete.Enabled = blnState
    mnuEditView.Enabled = True
    tbrTool.Buttons("Modify").Enabled = blnState
    tbrTool.Buttons("Delete").Enabled = blnState
    tbrTool.Buttons("View").Enabled = True
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim blnCanUse As Boolean
    
    gstrSQL = "select ���,����,nvl(��������,0) as �������� from ������� where nvl(�Ƿ��ֹ,0)<>1 order by ���"
    Call OpenRecordset(rsTemp, "�����ʻ�")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п��ñ�����𣬲���ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mint���� = 0
    Call InitBill
    
    With cbo����
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("���")
            If rsTemp("���") = gintInsure Then
                '��ǰҽ����
                'ʹ��API�����Բ�����Click�¼�
                .ListIndex = .ListCount - 1
            End If
            
            rsTemp.MoveNext
        Loop
    End With
    
    Call RestoreWinState(Me, App.ProductName)
    Call Ȩ������
    mblnLoad = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With Msf�ʻ��䶯��¼
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuEditAdjust_Batch_Click()
    Dim blnRefresh As Boolean
    With frm�����ʻ�����
        blnRefresh = .ShowME(Me, 2, mint����, 0)
    End With
    If blnRefresh Then Call FillList
End Sub

Private Sub mnuEditAdjust_Single_Click()
    Dim blnRefresh As Boolean
    With frm�����ʻ�����
        blnRefresh = .ShowME(Me, 1, mint����, 0)
    End With
    If blnRefresh Then Call FillList
End Sub

Private Sub mnuEditDelete_Click()
    Dim lngID As Long
    Dim blnTrans As Boolean
    On Error GoTo errHand
    
    lngID = Val(Msf�ʻ��䶯��¼.TextMatrix(Msf�ʻ��䶯��¼.Row, colID))
    If lngID = 0 Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ�������ʻ�������¼��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    blnTrans = True
    gcnOracle.BeginTrans
    gstrSQL = "ZL_�ʻ��䶯��¼_DELETE(" & lngID & ",'" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call ����ʻ���Ϣ_����(Msf�ʻ��䶯��¼.TextMatrix(Msf�ʻ��䶯��¼.Row, col����), True)
    gcnOracle.CommitTrans
    blnTrans = False
    
    Call FillList
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Private Sub mnuEditModify_Click()
    Dim blnRefresh As Boolean
    With frm�����ʻ�����
        blnRefresh = .ShowME(Me, 3, mint����, Val(Msf�ʻ��䶯��¼.TextMatrix(Msf�ʻ��䶯��¼.Row, colID)))
    End With
    If blnRefresh Then Call FillList
End Sub

Private Sub mnuEditView_Click()
    With frm�����ʻ�����
        Call .ShowME(Me, 4, mint����, Val(Msf�ʻ��䶯��¼.TextMatrix(Msf�ʻ��䶯��¼.Row, colID)))
    End With
End Sub

Private Sub mnuFileExcel_Click()
    Call subPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call subPrint(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call subPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub subPrint(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    If gstrUserName = "" Then GetUserInfo
    intRow = Msf�ʻ��䶯��¼.Row
    
    '��ͷ
    objOut.Title.Text = "�ʻ��䶯��¼�嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    objRow.Add "ҽ�����" & cbo����.Text
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate, "yyyy��MM��DD��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = Msf�ʻ��䶯��¼
    
    '���
    Msf�ʻ��䶯��¼.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    Msf�ʻ��䶯��¼.Redraw = True
    
    Msf�ʻ��䶯��¼.Row = intRow
    Msf�ʻ��䶯��¼.COL = 0: Msf�ʻ��䶯��¼.ColSel = Msf�ʻ��䶯��¼.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewFind_Click()
    Dim strTmp As String
    With frm�ʻ���֧����_����
         strTmp = .ShowME(Me, mint����)
    End With
    
    If Trim(strTmp) = "" Then Exit Sub
    mstrFind = strTmp
    Call FillList
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbrTool.Visible = Not cbrTool.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrTool.Buttons.Count
        tbrTool.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrTool.Buttons(i).Tag, "")
    Next
    cbrTool.Bands(1).MinHeight = tbrTool.ButtonHeight
    Form_Resize
End Sub

Private Sub Msf�ʻ��䶯��¼_DblClick()
    Call Msf�ʻ��䶯��¼_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub Msf�ʻ��䶯��¼_EnterCell()
    Dim intCol As Integer, lngColor As Long
    Dim lngSelectRow As Long
    On Error Resume Next
    
    With Msf�ʻ��䶯��¼
        '-----���ϴ�ѡ���м���ǰѡ���н�����ɫ����-----
        .Redraw = False
        lngSelectRow = .Row     '���浱ǰѡ����
        If lngCardRow <> 0 Then
            .Row = lngCardRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                .COL = intCol
                .CellBackColor = glng��ɫ
                Select Case .TextMatrix(.Row, col����)
                Case 1
                    lngColor = glng��ɫ
                Case 2
                    lngColor = glng����ɫ
                Case Else
                    lngColor = glng��ɫ
                End Select
                .CellForeColor = lngColor
            Next
            .COL = 0
        End If
        
        lngCardRow = lngSelectRow
        .Row = lngCardRow       '���õ�ǰѡ����
        If Not ActiveControl Is Nothing Then
            For intCol = 0 To .Cols - 1
                .COL = intCol
                .CellBackColor = glng���ɫ
                Select Case .TextMatrix(.Row, col����)
                Case 1
                    lngColor = glng��ɫ
                Case 2
                    lngColor = glng����ɫ
                Case Else
                    lngColor = glng��ɫ
                End Select
                .CellForeColor = lngColor
            Next
        End If
        .COL = 0
        .Redraw = True
        
        '-----���ݵ�ǰ��¼��״̬�����ò˵���������-----
        Call SetMenu(Val(.TextMatrix(.Row, col����)) = 1)
    End With
End Sub

Private Sub Msf�ʻ��䶯��¼_GotFocus()
    Dim intCol As Integer
    Dim lngColor As Long
    
    With Msf�ʻ��䶯��¼
        .GridColorFixed = glng��ɫ
        .GridColor = glng��ɫ
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .COL = intCol
            .CellBackColor = glng���ɫ
            Select Case .TextMatrix(.Row, col����)
            Case 1
                lngColor = glng��ɫ
            Case 2
                lngColor = glng����ɫ
            Case Else
                lngColor = glng��ɫ
            End Select
            .CellForeColor = lngColor
            .Redraw = True
        Next
        .COL = 0
    End With
End Sub

Private Sub Msf�ʻ��䶯��¼_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With Msf�ʻ��䶯��¼
            If Val(.TextMatrix(.Row, colID)) = 0 Then Exit Sub
            Call mnuEditView_Click
        End With
    End If
End Sub

Private Sub Msf�ʻ��䶯��¼_LostFocus()
    Dim intCol As Integer
    Dim lngColor As Long
    
    With Msf�ʻ��䶯��¼
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .COL = intCol
            .CellBackColor = glng��ɫ
            Select Case .TextMatrix(.Row, col����)
            Case 1
                lngColor = glng��ɫ
            Case 2
                lngColor = glng����ɫ
            Case Else
                lngColor = glng��ɫ
            End Select
            .CellForeColor = lngColor
            .Redraw = True
        Next
        .COL = 0
    End With
End Sub

Private Sub Msf�ʻ��䶯��¼_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu mnuEdit, 2
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Print"
        Call mnuFilePrint_Click
    Case "Printview"
        Call mnuFilePreview_Click
    Case "Adjust"
        Call mnuEditAdjust_Single_Click
    Case "Modify"
        Call mnuEditModify_Click
    Case "Delete"
        Call mnuEditDelete_Click
    Case "View"
        Call mnuEditView_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpTitle_Click
    Case "Quit"
        Call mnuFileQuit_Click
    End Select
End Sub

Private Sub tbrTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Single"
        Call mnuEditAdjust_Single_Click
    Case "Batch"
        Call mnuEditAdjust_Batch_Click
    End Select
End Sub

Private Sub tbrTool_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu mnuViewTool, 2
End Sub

Private Sub Ȩ������()
    mstrPrivs = gstrPrivs
    If InStr(1, mstrPrivs, "�༭") = 0 Then
        mnuEditAdjust.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
        mnuEditSplit1.Visible = False
        Me.tbrTool.Buttons("Adjust").Visible = False
        Me.tbrTool.Buttons("Modify").Visible = False
        Me.tbrTool.Buttons("Delete").Visible = False
    End If
End Sub
