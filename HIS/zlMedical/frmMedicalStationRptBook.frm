VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedicalStationRptBook 
   Caption         =   "#"
   ClientHeight    =   7800
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   11820
   Icon            =   "frmMedicalStationRptBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11820
   Begin zl9Medical.VsfGrid vsf 
      Height          =   2685
      Index           =   1
      Left            =   3120
      TabIndex        =   18
      Top             =   3000
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   4736
   End
   Begin zl9Medical.VsfGrid vsf 
      Height          =   1470
      Index           =   0
      Left            =   2850
      TabIndex        =   17
      Top             =   1140
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   2593
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   8370
      Top             =   4665
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":020A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":05A4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":083A
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":0BD4
            Key             =   "סԺ"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":0F6E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":1308
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":16A2
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":1938
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":1BCE
            Key             =   "GChecked"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":1E64
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":20FA
            Key             =   "Checked"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7440
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationRptBook.frx":2390
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15769
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7950
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":2C24
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":339E
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":3B18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":3D32
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":3F4C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":416C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":438C
            Key             =   "mail"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":4B06
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":5280
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":59FA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":5C14
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":5E2E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":604E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":626E
            Key             =   "mail"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   11820
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   1138
         ButtonWidth     =   1296
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&V.Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��(Alt+V)"
               Object.Tag             =   "&V.Ԥ��"
               ImageKey        =   "PrintView"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&P.��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ(Alt+P)"
               Object.Tag             =   "&P.��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Alt+A)"
               Object.Tag             =   "&A.ȫѡ"
               ImageKey        =   "SelectAll"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Alt+C)"
               Object.Tag             =   "&C.ȫ��"
               ImageKey        =   "ClearAll"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fra1 
      Height          =   1770
      Left            =   90
      TabIndex        =   8
      Top             =   765
      Width           =   2580
      Begin VB.CheckBox chk 
         Caption         =   "&4.��ӡ����"
         Height          =   240
         Index           =   3
         Left            =   255
         TabIndex        =   15
         Top             =   1395
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chk 
         Caption         =   "&2.��ӡ�ܼ�"
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   10
         Top             =   570
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chk 
         Caption         =   "&3.��ӡ��Ŀ"
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   11
         Top             =   900
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chk 
         Caption         =   "&1.��ӡ����"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   9
         Top             =   255
         Value           =   1  'Checked
         Width           =   1410
      End
   End
   Begin VB.Frame fra2 
      Height          =   3885
      Left            =   75
      TabIndex        =   12
      Top             =   2475
      Width           =   2595
      Begin VB.OptionButton opt 
         Caption         =   "&6.��Ŀ�����ʽ"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   2295
         Width           =   1590
      End
      Begin VB.OptionButton opt 
         Caption         =   "&5.��Ŀͳһ��ʽ"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   345
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.ListBox lstStyle 
         Height          =   1530
         Left            =   480
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   675
         Width           =   2025
      End
   End
   Begin VB.Frame fra3 
      Height          =   630
      Left            =   4110
      TabIndex        =   3
      Top             =   5985
      Width           =   3315
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1980
         TabIndex        =   1
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   675
         Picture         =   "frmMedicalStationRptBook.frx":69E8
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Tag             =   "����"
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&7.����"
         Height          =   180
         Index           =   1
         Left            =   1020
         TabIndex        =   0
         Tag             =   "����"
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "Ԥ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRptGroup 
         Caption         =   "���屨�浥"
         Begin VB.Menu mnuFileRptGroupPrintView 
            Caption         =   "Ԥ��(&V)"
         End
         Begin VB.Menu mnuFileRptGroupPrint 
            Caption         =   "��ӡ(&P)"
         End
         Begin VB.Menu mnuFileRptGroupExcel 
            Caption         =   "�����&Excel"
         End
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "ȫѡ(&A)"
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "ȫ��(&C)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
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
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
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
Attribute VB_Name = "frmMedicalStationRptBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnStarted As Boolean
Private mlng����id As Long
Private mblnDataMoved As Boolean

Private WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
        
    If vData = False Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    End If
    
    
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePrintView.Enabled
'
    
    
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next



    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng����id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    mlng����id = lng����id
    mlngKey = lngKey
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadPersonData(mlngKey, lng����id) = False Then Exit Function
    
    If Val(vsf(0).RowData(vsf(0).Row)) > 0 Then
        If ReadData(mlngKey, Val(vsf(0).RowData(vsf(0).Row))) = False Then Exit Function
    End If
    
    If opt(0).Value Then
        Call opt_Click(0)
    Else
        Call opt_Click(1)
    End If
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����������ʽ", "a")
    If strTmp <> "a" Then
        strTmp = ";" & strTmp & ";"
        For lngLoop = 0 To lstStyle.ListCount - 1
            If InStr(strTmp, ";" & lstStyle.List(lngLoop) & ";") > 0 Then
                lstStyle.Selected(lngLoop) = True
            Else
                lstStyle.Selected(lngLoop) = False
            End If
        Next
    End If
    
    
'    EditChanged = (Val(vsf(0).RowData(1)) > 0)

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadPersonData(ByVal lng�Ǽ�id As Long, Optional ByVal lng����id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
            
          
    gstrSQL = _
            "Select 0 As ѡ��, B.����, B.�����, B.������,b.���￨��, b.���֤��,a.�����,B.����id As ID,b.�Ա�,'' As �嵥 " & vbNewLine & _
            "From �����Ա���� A, ������Ϣ B" & vbNewLine & _
            "Where ��챨�� = 1 And A.����id = B.����id And A.�Ǽ�id = [1] "
    
    If lng����id > 0 Then
        gstrSQL = gstrSQL & " And a.����id=[2] "
    End If
    
    Call ClearGrid(vsf(0))
    
    mblnDataMoved = DataMove(lng�Ǽ�id)
    If mblnDataMoved Then
        gstrSQL = Replace(gstrSQL, "�����Ա����", "H�����Ա����")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�Ǽ�id, lng����id)
    If rs.BOF = False Then
        Call FillGrid(vsf(0), rs)
'        Call LoadGrid(vsf(0), rs, , , ils13)
        
    End If
    vsf(0).AppendRow = True
    ReadPersonData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function ShowItemSelect(ByVal lngRow As Long)
    Dim lngLoop As Long
    
    If vsf(0).TextMatrix(lngRow, 5) = "-1" Then
        'ȫ��
        vsf(1).Cell(flexcpText, 1, 0, vsf(1).Rows - 1, 0) = 1
    ElseIf vsf(0).TextMatrix(lngRow, 5) = "" Then
        vsf(1).Cell(flexcpText, 1, 0, vsf(1).Rows - 1, 0) = 0
    Else
        For lngLoop = 1 To vsf(1).Rows - 1
            If InStr(vsf(0).TextMatrix(lngRow, 5), "," & vsf(1).RowData(lngLoop) & ",") > 0 Then
                vsf(1).TextMatrix(lngLoop, 0) = 1
            End If
        Next
    End If
        
End Function

Private Function ReadData(ByVal lngKey As Long, ByVal lng����id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
            
    Call ClearGrid(vsf(1))
    
    gstrSQL = " Select x.*,0 As ѡ��," & _
                      "y.���� As ִ�п���, " & _
                      "z.���� As ��Ŀ, " & _
                      "DECODE(x.����id,NULL,DECODE(d.�����ļ�id, NULL, '', '����'),Decode(h.��д��, NULL, '����', '����')) AS ״̬, " & _
                      "d.�����ļ�id as ����id, " & _
                      "h.��д�� AS ������, " & _
                      "TO_CHAR(h.��д����, 'yyyy-mm-dd hh24:mi') AS ʱ�� " & _
                 "From (Select e.id,c.����id, " & _
                              "a.ִ�п���id, " & _
                              "a.������Ŀid, " & _
                              "a.����;��, " & _
                              "DECODE(g.ִ��״̬,1,'��ȫִ��',2,'ȡ��ִ��',3,'����ִ��','') As ִ��״̬, " & _
                              "g.����id, " & _
                              "g.NO, " & _
                              "Decode(a.����id, Null, '', '����') As ���� " & _
                         "From �����Ŀҽ�� b, " & _
                              "�����Ŀ�嵥 a, " & _
                              "�����Ա���� c, " & _
                              "����ҽ����¼ e, " & _
                              "����ҽ������ g " & _
                        "Where a.ID = b.�嵥id " & _
                              "and b.����id = c.����id " & _
                              "and c.�Ǽ�id = a.�Ǽ�id " & _
                              "and e.id = b.ҽ��id " & _
                              "and e.������� In ('C', 'D') "
    gstrSQL = gstrSQL & _
                              "and g.ҽ��id = e.id " & _
                               "and c.�Ǽ�ID =[1]  And c.����id=[2] " & _
                       " Union All " & _
                         "Select f.id,c.����id, " & _
                                "a.ִ�п���id, " & _
                                "a.������Ŀid, " & _
                                "a.����;��, " & _
                                "DECODE(g.ִ��״̬,1,'��ȫִ��',2,'ȡ��ִ��',3,'����ִ��','') As ִ��״̬, " & _
                                "g.����id, " & _
                                "g.NO, " & _
                                "Decode(a.����id, Null, '', '����') As ���� " & _
                           "From �����Ŀҽ�� b, " & _
                                "�����Ŀ�嵥 a, " & _
                                "�����Ա���� c, " & _
                                "����ҽ����¼ e, " & _
                                "����ҽ����¼ f, " & _
                                "����ҽ������ g " & _
                          "Where a.ID = b.�嵥id " & _
                                "and b.����id = c.����id " & _
                                "and c.�Ǽ�id = a.�Ǽ�id " & _
                                "and e.id = b.ҽ��id " & _
                                "and e.������� = 'E' " & _
                                "and e.id = f.���id " & _
                                "and g.ҽ��id = f.id "
    gstrSQL = gstrSQL & _
                                "and c.�Ǽ�ID =[1] And c.����id=[2] " & _
                       ") x, " & _
                      "���ű� y, " & _
                      "������ĿĿ¼ z, " & _
                      "���Ƶ���Ӧ�� d, " & _
                      "���˲�����¼ h " & _
                "Where x.ִ�п���id = y.ID " & _
                      "and z.id = x.������Ŀid " & _
                      "and x.����id = h.id(+) " & _
                      "and d.Ӧ�ó���(+)=4 " & _
                      "and x.������Ŀid = d.������Ŀid(+) Order By y.����"
                      
    If mblnDataMoved Then
        gstrSQL = Replace(gstrSQL, "�����Ա����", "H�����Ա����")
        gstrSQL = Replace(gstrSQL, "�����Ŀ�嵥", "H�����Ŀ�嵥")
        gstrSQL = Replace(gstrSQL, "�����Ŀҽ��", "H�����Ŀҽ��")
        gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
        gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        gstrSQL = Replace(gstrSQL, "���˲�����¼", "H���˲�����¼")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, lng����id)
    If rs.BOF = False Then
        
        Call FillGrid(vsf(1), rs)
'        Call LoadGrid(vsf(1), rs, , , ils13)
        
    End If
    
    vsf(1).AppendRow = True
    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Me.Caption = "��챨����"
    
    mnuFileRptGroup.Visible = False
    mnuFile_2.Visible = False
    
    With vsf(0)
        .Cols = 0
'        .NewColumn "", 240
        .NewColumn "ѡ��", 450, 4, , 1, , flexDTBoolean
        .NewColumn "����", 900
        .NewColumn "�Ա�", 600, 1
        .NewColumn "�����", 900, 7
        .NewColumn "������", 900, 7
        .NewColumn "���￨��", 0, 1
        .NewColumn "���֤��", 0, 1
        .NewColumn "�����", 990, 1
                
        .NewColumn "�嵥", 0
'        .FixedCols = 1
        .NewColumn "", 15
        .ExtendLastCol = True
        .SelectMode = True
        .Body.GridColor = COLOR.ǳ��ɫ
        .Body.GridColorFixed = COLOR.ǳ��ɫ
        .AppendRow = True
    End With
    
    With vsf(1)
        .Cols = 0
'        .NewColumn "", 240
        .NewColumn "ѡ��", 450, 4, , 1, , flexDTBoolean
        .NewColumn "��Ŀ", 2400
        .NewColumn "ִ�п���", 1080
        .NewColumn "ִ��״̬", 900
        .NewColumn "����id", 0
        .NewColumn "����id", 0
        .NewColumn "No", 0
        .NewColumn "", 15
'        .FixedCols = 1
        .ExtendLastCol = True
        .Body.GridColor = COLOR.ǳ��ɫ
        .Body.GridColorFixed = COLOR.ǳ��ɫ
        .SelectMode = True
        .AppendRow = True
    End With

    mnuFileRptGroup.Visible = (mlng����id = 0)
    mnuFile_2.Visible = (mlng����id = 0)

    
    lstStyle.Clear
    gstrSQL = "select a.˵�� As ��ʽ from zlRPTFMTs a,zlreports b where a.����id=b.id and  b.���=[1] Order By a.���"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "ZL1_BILL_1861_2")
    If rs.BOF = False Then
        Do While Not rs.EOF
            lstStyle.AddItem zlCommFun.NVL(rs("��ʽ"))
            lstStyle.Selected(lstStyle.NewIndex) = True
            rs.MoveNext
        Loop
    End If
    
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function

Private Function GetReportCode(ByVal lngKey As Long, ByRef strCode As String, ByRef strNo As String, ByRef bytMode As Byte) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����;
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If lngKey = 0 Then Exit Function
    
    
    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' AS ������," & _
                       "D.NO," & _
                       "D.��¼���� " & _
                "FROM �����ļ�Ŀ¼ C,(SELECT A.NO,A.��¼����,E.�����ļ�id FROM ����ҽ������ A,����ҽ����¼ B,���Ƶ���Ӧ�� E WHERE E.Ӧ�ó���=4 AND E.������Ŀid=B.������Ŀid AND B.������� IN ('C','D') AND A.ҽ��id=B.ID AND (B.���id=[1] OR B.ID=[1]) AND ROWNUM<2) D " & _
                "Where C.ID=D.�����ļ�id"

    If mblnDataMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        strCode = zlCommFun.NVL(rs("������"))
        strNo = zlCommFun.NVL(rs("NO"))
        bytMode = zlCommFun.NVL(rs("��¼����"), 1)
    End If
    
    GetReportCode = True
    
End Function

Private Function PrintData(ByVal bytMode As Byte, Optional ByVal blnGroup As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strReportCode As String
    Dim lngLoop As Long
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim intѡ�� As Integer
    Dim strSQL As String
    Dim int����id As Integer
    Dim int����� As Integer
    Dim int����id As Integer
    Dim rs As New ADODB.Recordset
    Dim strSvr����� As String
    Dim lng����id As Long
    Dim intCount As Integer
    Dim lngRow As Long
    
    On Error GoTo errHand
    
    
    intѡ�� = 0
    
    If blnGroup Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_3", Me, "�Ǽ�id=" & mlngKey, bytMode)
    Else
        
        For lngLoop = 1 To vsf(0).Rows - 1
        
            If Val(vsf(0).RowData(lngLoop)) > 0 And Abs(Val(vsf(0).TextMatrix(lngLoop, intѡ��))) = 1 Then
                
                 '1.����"�������"
                If chk(0).Value = 1 Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_2_1", Me, "�Ǽ�id=" & mlngKey, "����id=" & Val(vsf(0).RowData(lngLoop)), bytMode)
                End If
                
                '2.����"����ܼ�"
                If chk(2).Value = 1 Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_2_2", Me, "�Ǽ�id=" & mlngKey, "����id=" & Val(vsf(0).RowData(lngLoop)), bytMode)
                End If
                
                '3.ͳһ������
                If opt(0).Value And chk(1).Value = 1 Then

                    gstrSQL = "zl_�����Ŀҽ��_Update(" & mlngKey & "," & Val(vsf(0).RowData(lngLoop)) & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    
                    For intCount = 0 To lstStyle.ListCount - 1
                        If lstStyle.Selected(intCount) Then
                            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_2", Me, "�Ǽ�id=" & mlngKey, "����id=" & Val(vsf(0).RowData(lngLoop)), "����=" & chk(3).Value, "REPORTFORMAT=" & (intCount + 1), bytMode)
                        End If
                    Next
                End If
                
                '4.����"��Ŀ����",ȱʡ����
                If opt(1).Value And chk(1).Value = 1 Then
                    
                    vsf(0).Row = lngLoop
                    Call vsf_AfterRowColChange(0, -1, -1, vsf(0).Row, vsf(0).Col)
                    
                    For lngRow = 1 To vsf(1).Rows - 1
                        
                        If Val(vsf(1).RowData(lngRow)) > 0 And Abs(Val(vsf(1).TextMatrix(lngRow, intѡ��))) = 1 Then
                        
                            If GetReportCode(Val(vsf(1).RowData(lngRow)), strReportCode, strReportParaNo, bytReportParaMode) Then
                                Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, bytMode)
                            End If
                            
                        End If
                    Next
                    
                End If
                
                '�����Ԥ����ֻһ��Ԥ��
                If bytMode = 1 Then Exit For
                
            End If
        Next
        
    End If
      
    PrintData = True

    Exit Function

errHand:

    If ErrCenter = 1 Then
        Resume
    End If

End Function


Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 255 * 8)
    
    txt(1).Text = ""
    LocationObj txt(1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("ȫѡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫѡ"))
        Case vbKeyC
            If tbrThis.Buttons("ȫ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫ��"))
        Case vbKeyM
            If tbrThis.Buttons("�ʼ�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�ʼ�"))
        Case vbKeyV
            If tbrThis.Buttons("Ԥ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("Ԥ��"))
        Case vbKeyP
            If tbrThis.Buttons("��ӡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("��ӡ"))
        Case vbKeyH
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyX
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End If
    End If
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Load()

    Call RestoreWinState(Me, App.ProductName)
    
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "0")) = 1 Then
        'ʹ�ø��Ի�����
      
        lbl(1).Caption = "&6." & (GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", "����"))
        lbl(1).Tag = Mid(lbl(1).Caption, 4)

    End If
    
    On Error Resume Next
    opt(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�������ʽ", 0))).Value = True

        
    chk(0).Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ����", 1))
    chk(1).Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ��Ŀ", 1))
    chk(2).Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ�ܼ�", 1))
    chk(3).Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ����", 1))
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fra1
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
    End With
    
    With fra2
        .Left = fra1.Left
        .Top = fra1.Top + fra1.Height - 90
        .Width = fra1.Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With

    With vsf(0)
        .Left = fra1.Left + fra1.Width
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fra3.Height + 90 - IIf(vsf(1).Visible, vsf(1).Height + 30, 0)
    End With
    
    vsf(1).Move vsf(0).Left, vsf(0).Top + vsf(0).Height + 30, vsf(0).Width
    
    If vsf(1).Visible Then
        fra3.Move vsf(1).Left, vsf(1).Top + vsf(1).Height - 90, vsf(1).Width
    Else
        fra3.Move vsf(0).Left, vsf(0).Top + vsf(0).Height - 90, vsf(0).Width
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngLoop As Long
    Dim strTmp As String
    
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", lbl(1).Tag)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ����", chk(0).Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ��Ŀ", chk(1).Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ�ܼ�", chk(2).Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ����", chk(3).Value)
    
    If opt(0).Value Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�������ʽ", 0)
        
        For lngLoop = 0 To lstStyle.ListCount - 1
            If lstStyle.Selected(lngLoop) Then strTmp = strTmp & ";" & lstStyle.List(lngLoop)
        Next
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����������ʽ", strTmp)
        
    Else
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�������ʽ", 1)
    End If
    
End Sub


Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    Dim intѡ�� As Integer
    
    intѡ�� = 0
    If intѡ�� >= 0 Then
    
        For lngLoop = 1 To vsf(0).Rows - 1
            If Val(vsf(0).RowData(lngLoop)) > 0 Then
                vsf(0).TextMatrix(lngLoop, intѡ��) = 0
                vsf(0).TextMatrix(lngLoop, 5) = ""
            End If
        Next
        
        EditChanged = False
        
    End If
    Call ShowItemSelect(vsf(0).Row)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    
    Call PrintData(2)

End Sub


'Private Sub mnuFilePrintSetRpt_Click(Index As Integer)
'
'    Select Case Index
'    Case 0
'        ReportPrintSet gcnOracle, glngSys, "ZL1_BILL_1861_2_1"
'    Case 1
'        ReportPrintSet gcnOracle, glngSys, "ZL1_BILL_1861_2"
'    Case 2
'        ReportPrintSet gcnOracle, glngSys, "ZL1_BILL_1861_2_2"
'    End Select
'
'End Sub

Private Sub mnuFilePrintSet_Click()

    Call frmMedicalStationRptPrintSet.ShowEdit(Me, mlngKey, mlng����id)
    
End Sub

Private Sub mnuFilePrintView_Click()
    
    Call PrintData(1)
    
End Sub

Private Sub mnuFileRptGroupExcel_Click()
    Call PrintData(3, True)
End Sub

Private Sub mnuFileRptGroupPrint_Click()
    Call PrintData(2, True)
End Sub

Private Sub mnuFileRptGroupPrintView_Click()
    Call PrintData(1, True)
End Sub

Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long
    Dim intѡ�� As Integer
    
    intѡ�� = 0
    If intѡ�� >= 0 Then
        For lngLoop = 1 To vsf(0).Rows - 1
            If Val(vsf(0).RowData(lngLoop)) > 0 Then
                vsf(0).TextMatrix(lngLoop, intѡ��) = 1
                vsf(0).TextMatrix(lngLoop, 5) = "-1"
                EditChanged = True
            End If
        Next
    End If
    Call ShowItemSelect(vsf(0).Row)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
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
    
    Select Case mbytPopMenu

    Case 3
        
        mobjPopMenu.Add 1, "&1.����", , , True, , (lbl(1).Tag = "����")
        mobjPopMenu.Add 2, "&2.�����", , , True, , (lbl(1).Tag = "�����")
        mobjPopMenu.Add 3, "&3.������", , , True, , (lbl(1).Tag = "������")
        mobjPopMenu.Add 4, "&4.���￨��", , , True, , (lbl(1).Tag = "���￨��")
        mobjPopMenu.Add 5, "&5.����ƴ��", , , True, , (lbl(1).Tag = "����ƴ��")
        mobjPopMenu.Add 6, "&6.�������", , , True, , (lbl(1).Tag = "�������")
        mobjPopMenu.Add 7, "&7.���֤��", , , True, , (lbl(1).Tag = "���֤��")
        mobjPopMenu.Add 8, "&8.�����", , , True, , (lbl(1).Tag = "�����")
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu

    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(1).Caption = "&7." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(1).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    End Select
End Sub

Private Sub opt_Click(Index As Integer)
    If Index = 0 Then
        vsf(1).Visible = False
    Else
        vsf(1).Visible = True
    End If
    
    Call Form_Resize
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ȫѡ"
        Call mnuFileSelectAll_Click
    Case "ȫ��"
        Call mnuFileClearAll_Click
    Case "Ԥ��"
        Call mnuFilePrintView_Click
    Case "��ӡ"
        Call mnuFilePrint_Click
    Case "�ʼ�"
        
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngLoop As Long
    Dim strCol As String
    Dim lngCol As Long
    Dim lngRow As Long
    Dim blnCard As Boolean
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    strCol = Mid(lbl(1).Caption, 4)
    lngCol = GetCol(vsf(0), strCol)
            
    If strCol = "���￨��" And KeyAscii <> vbKeyReturn Then
        '���￨�ţ��Զ�ʶ��

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.���￨���볤�� - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = vbKeyReturn

        End If

    End If
    
    If KeyAscii = vbKeyReturn Then
        
        If Index = 1 And Trim(txt(Index).Text) <> "" Then
            lngCol = -1
            Select Case Mid(lbl(1).Caption, 4)
            Case "����", "����ƴ��", "�������"
                lngCol = 1
            Case "�����"
                lngCol = 3
            Case "������"
                lngCol = 4
            Case "���￨��"
                lngCol = 5
            Case "���֤��"
                lngCol = 6
            Case "�����"
                lngCol = 7
            End Select
            If lngCol < 0 Then Exit Sub
            
            lngRow = 0
            If vsf(0).Row + 1 <= vsf(0).Rows - 1 Then
                For lngLoop = vsf(0).Row + 1 To vsf(0).Rows - 1
                
                    lngRow = 0
                    Select Case strCol
                    Case "�����"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "������"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���￨��"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���֤��"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����ƴ��"
                        If zlGetSymbol(UCase(vsf(0).TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "�������"
                        If zlGetSymbol(UCase(vsf(0).TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For

                Next
            End If
            
            If lngRow = 0 Then
                For lngLoop = 1 To vsf(0).Row
                    lngRow = 0
                    Select Case strCol
                    Case "�����"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "������"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���￨��"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���֤��"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����ƴ��"
                        If zlGetSymbol(UCase(vsf(0).TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "�������"
                        If zlGetSymbol(UCase(vsf(0).TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For
                Next
            End If
            
            If lngRow <= 0 Then
                ShowSimpleMsg "û���ҵ�����Ҫ�����Ϣ��"
                txt(Index).Text = ""
            Else
                vsf(0).ShowCell lngRow, vsf(0).Col
                vsf(0).Row = lngRow
            End If
        End If
        
        txt(Index).SetFocus
        zlControl.TxtSelAll txt(Index)
    End If
End Sub


Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    
    If Index = 1 Then
        
        vsf(0).TextMatrix(vsf(0).Row, 5) = ""
        For lngLoop = 1 To vsf(1).Rows - 1
            If Abs(Val(vsf(1).TextMatrix(lngLoop, 0))) = 1 Then
                vsf(0).TextMatrix(vsf(0).Row, 5) = vsf(0).TextMatrix(vsf(0).Row, 5) & "," & Val(vsf(1).RowData(lngLoop))
            End If
        Next
        If vsf(0).TextMatrix(vsf(0).Row, 5) <> "" Then vsf(0).TextMatrix(vsf(0).Row, 5) = vsf(0).TextMatrix(vsf(0).Row, 5) & ","
    Else
        If Abs(Val(vsf(0).TextMatrix(Row, 0))) = 1 Then
            vsf(0).TextMatrix(Row, 5) = "-1"
        Else
            vsf(0).TextMatrix(Row, 5) = ""
        End If
        Call ShowItemSelect(Row)
    End If
    
    EditChanged = True
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngLoop As Long
    
    On Error Resume Next
    
    If Index = 0 Then
        Call ReadData(mlngKey, Val(vsf(0).RowData(vsf(0).Row)))
        Call ShowItemSelect(NewRow)
                
    End If
End Sub

Private Sub vsf_BeforeDeleteCell(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeDeleteRow(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(Index As Integer, ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Abs(Val(vsf(0).TextMatrix(vsf(0).Row, 0))) = 0 And Index = 1 Then
        Cancel = True
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

