VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLaterVisitEdit 
   Caption         =   "�����ü�¼"
   ClientHeight    =   6645
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   10500
   Icon            =   "frmLaterVisitEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10500
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1410
      Left            =   0
      TabIndex        =   5
      Top             =   1275
      Width           =   6645
      _cx             =   11721
      _cy             =   2487
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
      BackColorSel    =   8388608
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   1
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
      Begin VB.Line lnY2 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
      Begin VB.Line lnX2 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin VB.Frame fra1 
      Height          =   525
      Left            =   0
      TabIndex        =   6
      Top             =   2625
      Width           =   8385
      Begin VB.OptionButton opt 
         Caption         =   "&1.����"
         Height          =   210
         Index           =   0
         Left            =   945
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton opt 
         Caption         =   "&2.�۲�"
         Height          =   210
         Index           =   1
         Left            =   1905
         TabIndex        =   9
         Top             =   210
         Width           =   840
      End
      Begin VB.OptionButton opt 
         Caption         =   "&3.����"
         Height          =   210
         Index           =   2
         Left            =   2805
         TabIndex        =   10
         Top             =   210
         Width           =   840
      End
      Begin VB.OptionButton opt 
         Caption         =   "&4.����"
         Height          =   210
         Index           =   3
         Left            =   3690
         TabIndex        =   11
         Top             =   210
         Width           =   840
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "���(&L)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   225
         Width           =   705
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   6285
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLaterVisitEdit.frx":076A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13467
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   9225
      Top             =   1545
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
            Picture         =   "frmLaterVisitEdit.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1218
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1438
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1652
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1872
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1545
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
            Picture         =   "frmLaterVisitEdit.frx":1A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":2218
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":2438
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10500
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
         TabIndex        =   17
         Top             =   30
         Width           =   10380
         _ExtentX        =   18309
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
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+S)"
               Object.Tag             =   "&S.����"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&R.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+R)"
               Object.Tag             =   "&R.����"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fra2 
      Height          =   2700
      Left            =   0
      TabIndex        =   12
      Top             =   3075
      Width           =   8385
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2445
         Left            =   915
         TabIndex        =   14
         Top             =   165
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   4313
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         MaxLength       =   4000
         TextRTF         =   $"frmLaterVisitEdit.frx":2658
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   13
         Top             =   210
         Width           =   705
      End
   End
   Begin VB.Frame fra0 
      Height          =   570
      Left            =   405
      TabIndex        =   0
      Top             =   720
      Width           =   9555
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   3705
         MaxLength       =   20
         TabIndex        =   4
         Top             =   165
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   1215
         TabIndex        =   2
         Top             =   165
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75431939
         CurrentDate     =   38406
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         Caption         =   "12345678"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5520
         TabIndex        =   19
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NO:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5190
         TabIndex        =   18
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�����(&M)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2775
         TabIndex        =   3
         Top             =   225
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�������(&D)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "����(&R)"
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
Attribute VB_Name = "frmLaterVisitEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mstrNo As String
Private mlng����id As Long
Private mvarParam As Variant

Private Enum mCol
    ��� = 1
    ��Ŀ
    ����
    ��λ
End Enum

'�������Զ�����̻���************************************************************************************************
Private Function ShowOpenList(Optional strText As String) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '����:���б�ṹ�����Ƽ���걾����
    '����:������2;�ɹ�����1;ȡ������0
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strLvw = "����,900,0,1;����,1800,0,0;�Ƿ񼲲�,900,0,0"

    ShowOpenList = 2
    
    strSQL = _
                "SELECT ��� AS ID, " & _
                        "A.����, " & _
                        "A.����, " & _
                        "Decode(�Ƿ񼲲�,1,'��','') As �Ƿ񼲲� " & _
                "FROM �����Ͻ��� A " & _
                "WHERE NVL(ĩ��,0)=1 "
    strSQL = strSQL & " AND (A.���� Like [1] OR A.���� Like [1] OR A.���� Like [1])"
            
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "%" & UCase(strText) & "%")
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
        
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "����±���ѡ��һ�����", sglX + 60, sglY, 9000, 5100, 300, , Me.Name & "\�����۹���ѡ��", , False) Then GoTo Over
    
    Exit Function
    
Over:
    vsf.RowData(vsf.Row) = 1
    vsf.EditText = zlCommFun.NVL(rs("����").Value)
    vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
    vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub GoNextCell()
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngCol As Long
    Dim blnCancel As Boolean
    
    If GetAllowCol(vsf.Col + 1) > vsf.Cols - 1 Then
        '����֮ǰ���ȼ���Ƿ������У����Ƿ��б������Ŀû������
                
        If vsf.Row = vsf.Rows - 1 Then
            blnCancel = False
            
            lngCol = 1
            
            If blnCancel Then
                vsf.Col = lngCol
                vsf.ShowCell vsf.Row, vsf.Col
                Exit Sub
            End If
            
            Call InsertNewRow
        Else
            vsf.Row = vsf.Row + 1
        End If
        
        '�ҵ�һ�����Ա༭����
        vsf.Col = GetAllowCol(1)
    Else
        '����һ�����Ա༭����
        vsf.Col = GetAllowCol(vsf.Col + 1)
    End If
    
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Private Sub InsertNewRow()
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    If vsf.Editable <> flexEDNone Then
        vsf.AddItem "", vsf.Rows
        vsf.Row = vsf.Rows - 1
    Else
        vsf.Row = vsf.Rows - 1
    End If
    
    Call AdjustRowFlag(vsf)
    
End Sub

Private Function GetAllowCol(ByVal lngFromCol As Long) As Long
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lngLoop As Long
    
    lngRow = vsf.Row
    
    For lngLoop = lngFromCol To vsf.Cols - 1
        If lngLoop = 3 Then Exit For
    Next
    
    GetAllowCol = lngLoop
End Function

Private Sub AdjustRowFlag(ByRef objVsf As Object, Optional ByVal intRow As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    Dim lngLoop As Long
    
    For lngLoop = 0 To vsf.Rows - 1
        vsf.TextMatrix(lngLoop, 0) = lngLoop + 1 & "��"
    Next
End Sub

Private Sub ShowSelectRow(ByRef objVsf As Object, Optional ByVal intRow As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Cell(flexcpFontBold, 0, 0, vsf.Rows - 1, vsf.Cols - 1) = False
    vsf.Cell(flexcpFontBold, intRow, 0, intRow, vsf.Cols - 1) = True
    
End Sub

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    mnuFileSave.Enabled = True
    mnuFileRestore.Enabled = True

    If vData = False Then
        mnuFileSave.Enabled = False
        mnuFileRestore.Enabled = False

    End If
    
    tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    tbrThis.Buttons("����").Enabled = mnuFileRestore.Enabled
        
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error Resume Next
    
    vsf.Rows = 1
    vsf.Cell(flexcpText, 0, 0, vsf.Rows - 1, vsf.Cols - 1) = ""
    vsf.RowData(0) = 0
    vsf.TextMatrix(0, 0) = "1��"
    
    Call ReadRow(vsf.Row)
    
    On Error GoTo 0
    
    EditChanged = True
    
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    Dim varGroup As Variant
    Dim lngLoop As Long
    
    mblnStartUp = True
    mblnOK = False
    
    '����id,��쵥��,��õ���
    mvarParam = Split(strParam, "'")
    
    mstrNo = mvarParam(2)
    mlng����id = Val(mvarParam(0))
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    EditChanged = False
    
    If ReadData = False Then Exit Function
    If mstrNo <> "" Then EditChanged = False
    Call ShowSelectRow(vsf, 0)
    
    'stbThis.Panels(2).Text = "��д/ѡ�������Ա���ϡ�"
                
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If mstrNo = "" Then
                        
        lblNo.Caption = GetNextNo(79)
        
        '���ȴ��ϴ�����ռ�
        gstrSQL = "select A.���ʱ��," & _
                        "'" & UserInfo.���� & "' as �����," & _
                        "'" & lblNo.Caption & "' as no," & _
                        "A.�������," & _
                        "0 AS ��ý��," & _
                        "'' AS ������ " & _
                    "from �����ü�¼ A,�����Ա���� B,���ǼǼ�¼ C " & _
                    "Where A.����ID = B.����ID " & _
                        "AND A.��쵥��=C.���� " & _
                        "AND A.���ʱ��=B.���ʱ�� " & _
                        "AND A.��ý��<>1 " & _
                        "AND B.�Ǽ�id=C.ID " & _
                        "AND B.����id=[1] " & _
                        "AND C.����=[2] " & _
                    "order by A.���"
                        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����id, CStr(mvarParam(1)))
        If rs.BOF = False Then
                
        Else
            
            '�ٴ�������ռ�
                        
            gstrSQL = _
                "select TRUNC(sysdate) as ���ʱ��," & _
                       "'" & UserInfo.���� & "' as �����," & _
                       "'" & lblNo.Caption & "' as no," & _
                       "A.�������� as �������," & _
                       "0 as ��ý��," & _
                       "'' as ������ " & _
                "from �����Ա���� A," & _
                     "(SELECT ��첡��ID FROM �����Ա���� U,���ǼǼ�¼ T WHERE ROWNUM<2 AND U.�Ǽ�id=T.ID AND U.���״̬=5 AND U.����id=" & mlng����id & " AND T.����='" & mvarParam(1) & "' AND U.���ʱ��=(SELECT MAX(X.���ʱ��) FROM  �����Ա���� X,���ǼǼ�¼ Y WHERE X.�Ǽ�id=Y.ID AND X.����id=" & mlng����id & " AND Y.����='" & mvarParam(1) & "')) B," & _
                     "���˲������� C " & _
                "Where B.��첡��ID = C.������¼ID " & _
                      "AND C.ID=A.����id " & _
                      "AND A.��¼����=0 " & _
                "ORDER BY A.��¼���"
        End If
    Else
        '�ӱ������ռ�
        
        gstrSQL = "SELECT * FROM �����ü�¼ WHERE NO=[1] Order by ���"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo)
    If rs.BOF = False Then
        
        dtp.Value = Format(zlCommFun.NVL(rs("���ʱ��")), dtp.CustomFormat)
        txt.Text = zlCommFun.NVL(rs("�����"))
        lblNo.Caption = zlCommFun.NVL(rs("NO"))
        
        Do While Not rs.EOF
            
            If Val(vsf.RowData(vsf.Rows - 1)) = 1 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = 1
            vsf.TextMatrix(vsf.Rows - 1, 0) = vsf.Rows & "��"
            vsf.TextMatrix(vsf.Rows - 1, 1) = zlCommFun.NVL(rs("��ý��"), 0)
            vsf.TextMatrix(vsf.Rows - 1, 2) = zlCommFun.NVL(rs("������"))
            vsf.TextMatrix(vsf.Rows - 1, 3) = zlCommFun.NVL(rs("�������"))
            
            rs.MoveNext
        Loop
    End If
    
    Call ReadRow(vsf.Row)

            
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
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    vsf.FixedRows = 0
    vsf.FixedCols = 1
    vsf.Cols = 4
    vsf.Rows = 1
    vsf.ColWidth(0) = 300
    
    vsf.ComboList = "..."
    
    vsf.ColHidden(1) = True         '������
    vsf.ColHidden(2) = True         '��������
    
    vsf.TextMatrix(0, 0) = "1��"
    vsf.BackColorFixed = vsf.BackColor
    vsf.GridLines = flexGridNone
    vsf.GridLinesFixed = flexGridNone
                        
    vsf.Editable = flexEDKbdMouse
    
    dtp.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    txt.Text = UserInfo.����
    
    lblNo.Caption = ""
    
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
    Dim rs As New ADODB.Recordset
    
    For lngLoop = 0 To vsf.Rows - 1
        
        If StrIsValid(vsf.TextMatrix(lngLoop, 3), 100) = False Then
            vsf.Row = lngLoop
            vsf.Col = 3
            Call vsf.ShowCell(vsf.Row, vsf.Col)
            Exit Function
        End If
                
    Next
    
    If StrIsValid(rtb.Text, 4000) = False Then
        rtb.SetFocus
        Exit Function
    End If
                                                                
    ValidEdit = True
    
End Function

Private Function SaveEdit(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strNow As String
    Dim rsPati As New ADODB.Recordset
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    strSQL(ReDimArray(strSQL)) = "ZL_�����ü�¼_DELETE('" & lblNo.Caption & "')"
    For lngLoop = 0 To vsf.Rows - 1
        If Trim(vsf.TextMatrix(lngLoop, 3)) <> "" Then
            strSQL(ReDimArray(strSQL)) = "ZL_�����ü�¼_INSERT(" & mlng����id & ",'" & mvarParam(1) & "','" & lblNo.Caption & "'," & lngLoop + 1 & ",'" & Trim(vsf.TextMatrix(lngLoop, 3)) & "'," & Val(vsf.TextMatrix(lngLoop, 1)) & ",'" & vsf.TextMatrix(lngLoop, 2) & "','" & txt.Text & "',to_date('" & Format(dtp.Value, "yyyy-MM-dd") & "','yyyy-mm-dd hh24:mi:ss'),to_date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function


Private Sub dtp_Change()
    EditChanged = True
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyS
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyR
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
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

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    
    With fra0
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 75
        .Width = Me.ScaleWidth - .Left
    End With
    
    With vsf
        .Left = 0
        .Top = fra0.Top + fra0.Height + 15
        .Width = Me.ScaleWidth - .Left
    End With
    
    With fra1
        .Left = 0
        .Top = vsf.Top + vsf.Height - 60
        .Width = vsf.Width
    End With
    
    With fra2
        .Left = 0
        .Top = fra1.Top + fra1.Height - 90
        .Width = vsf.Width
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
    
    With rtb
        .Top = 150
        .Width = fra2.Width - .Left - 75
        .Height = fra2.Height - .Top - 75
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mnuFileSave.Enabled Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRestore_Click()
        
    If MsgBox("ȷʵҪ�ָ�����ǰ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Call ClearData
    Call ReadData
    
    EditChanged = False
    
End Sub

Private Sub mnuFileSave_Click()
    Dim lngKey As Long
    
    Call WriteRow(vsf.Row)
        
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit(lngKey) Then
        EditChanged = False
        mblnOK = True
        
        mstrNo = lblNo.Caption
        
        On Error Resume Next
        Call mfrmMain.EditRefresh("��ü�¼", lblNo.Caption)
        
    End If
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

Private Sub opt_Click(Index As Integer)
    EditChanged = True
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub rtb_Change()
    EditChanged = True
End Sub

Private Sub rtb_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtb_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub rtb_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(rtb.Text, rtb.MaxLength)
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "����"
        Call mnuFileSave_Click
    Case "����"
        Call mnuFileRestore_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub txt_Change()
    EditChanged = True
End Sub

Private Sub WriteRow(ByVal lngRow As Long)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim blnSvr As Boolean
    
    blnSvr = mnuFileSave.Enabled
    
    If opt(0).Value Then
        vsf.TextMatrix(lngRow, 1) = "1"
    ElseIf opt(1).Value Then
        vsf.TextMatrix(lngRow, 1) = "2"
    ElseIf opt(2).Value Then
        vsf.TextMatrix(lngRow, 1) = "3"
    ElseIf opt(3).Value Then
        vsf.TextMatrix(lngRow, 1) = "4"
    Else
        vsf.TextMatrix(lngRow, 1) = ""
    End If
    
    vsf.TextMatrix(lngRow, 2) = rtb.Text
    
    EditChanged = blnSvr
End Sub

Private Sub ReadRow(ByVal lngRow As Long)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim blnSvr As Boolean
    
    blnSvr = mnuFileSave.Enabled
    
    If Val(vsf.TextMatrix(lngRow, 1)) >= 1 And Val(vsf.TextMatrix(lngRow, 1)) <= 4 Then
        opt(Val(vsf.TextMatrix(lngRow, 1)) - 1).Value = True
    Else
        opt(0).Value = True
    End If
    
    rtb.Text = vsf.TextMatrix(lngRow, 2)
    
    EditChanged = blnSvr
End Sub

Private Sub txt_GotFocus()
    
    zlCommFun.OpenIme True
    
    zlControl.TxtSelAll txt
    
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        vsf.Col = 3
        vsf.SetFocus
    End If
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    EditChanged = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    
    If OldRow <> NewRow Then
        Call ShowSelectRow(vsf, NewRow)
    End If
    vsf.ComboList = "..."
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
     
    On Error Resume Next
    
    If OldRow <> NewRow Then
        Call WriteRow(OldRow)
        Call ReadRow(NewRow)
    End If
    
    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "SELECT -1 AS ID," & _
                        "0 AS �ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "'' AS ����," & _
                        "'���з���' AS ����, " & _
                        "'' AS ���� " & _
                "FROM dual "
                
    gstrSQL = gstrSQL & _
            " UNION ALL " & _
            "SELECT ��� AS ID," & _
                        "DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "����," & _
                        "����, " & _
                        "'' AS ���� " & _
                "FROM �����Ͻ��� " & _
                "WHERE NVL(ĩ��,0)=0 " & _
                "START WITH �ϼ���� is NULL CONNECT BY PRIOR ��� = �ϼ���� "
    
    gstrSQL = gstrSQL & _
                "UNION ALL " & _
                "SELECT ��� AS ID, " & _
                        "DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.����, " & _
                        "DECODE(�Ƿ񼲲�,1,'��','��') AS ���� " & _
                "FROM �����Ͻ��� A " & _
                "WHERE NVL(ĩ��,0)=1 "
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If ShowGrdSelect(Me, vsf, "����,900,0,1;����,1800,0,0;����,900,0,0", Me.Name & "\������ѡ��", "����б���ѡ��һ�����ۡ�", rsData, rs, 9000, 5100) Then

        
        vsf.RowData(vsf.Row) = 1
        vsf.EditText = zlCommFun.NVL(rs("����").Value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
        
        EditChanged = True
        
    End If
        
End Sub

Private Sub vsf_DblClick()

    Call vsf_KeyPress(32)
    
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngLoop As Long
    Dim blnCancel As Boolean
    
    On Error GoTo errHand
    
    Select Case KeyCode
    Case vbKeyDelete
        
        If Shift = 0 And vsf.Editable <> flexEDNone Then
            'ɾ�����м�����
            
            If vsf.Rows > 0 Then
                If vsf.Rows = 1 And vsf.Row = 0 Then
                    For lngLoop = 0 To vsf.Cols - 1
                        vsf.TextMatrix(0, lngLoop) = ""
                    Next
                    vsf.RowData(0) = ""
                Else
                    vsf.RemoveItem vsf.Row
                    
                End If
                Call AdjustRowFlag(vsf, vsf.Row)
                
            End If
            
        End If
        
        If Shift = 1 And vsf.Editable <> flexEDNone And vsf.Col = 3 Then
            'ɾ����ǰ��Ԫ�������
            
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
            
        End If
    End Select
    
    Exit Sub
    
errHand:
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strSvrText As String
    
    If KeyCode = vbKeyReturn Then
        '����2-�����͵����
        
        If InStr(vsf.EditText, "'") > 0 Then
            KeyCode = 0
            Exit Sub
        End If

        strSvrText = vsf.EditText
        Select Case ShowOpenList(vsf.EditText)
        Case 2
            'ȡ���˱���ѡ��
            KeyCode = 0
            
            vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
            vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)

        End Select
    End If
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Trim(vsf.TextMatrix(vsf.Row, vsf.Col)) = "" Then
            zlCommFun.PressKey vbKeyTab
        Else
            Call GoNextCell
        End If
    Else
        If vsf.ComboList = "..." Then vsf.ComboList = ""
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call GoNextCell
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

