VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmFeeVerifyManger 
   Caption         =   "������˹���"
   ClientHeight    =   10830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15075
   Icon            =   "frmFeeVerifyManger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15075
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   10470
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFeeVerifyManger.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16907
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4584
            MinWidth        =   4584
            Picture         =   "frmFeeVerifyManger.frx":0E1E
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox picMzToZy 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   135
      ScaleHeight     =   4935
      ScaleWidth      =   14475
      TabIndex        =   3
      Top             =   780
      Width           =   14475
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   630
         TabIndex        =   27
         Top             =   150
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmFeeVerifyManger.frx":189F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12;F5"
         MustSelectItems =   "����"
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdBrush 
         Caption         =   "ˢ��(&N)"
         Height          =   375
         Left            =   11595
         TabIndex        =   25
         Top             =   555
         Width           =   1245
      End
      Begin VB.CheckBox chk��ת������ 
         Caption         =   "��ʾ��ת������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2385
         TabIndex        =   20
         Top             =   645
         Width           =   2004
      End
      Begin VB.Frame fra���� 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   8595
         TabIndex        =   17
         Top             =   90
         Width           =   2715
         Begin VB.ComboBox cbo�շѵ� 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   675
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   60
            Width           =   2040
         End
         Begin VB.Label lblBill 
            AutoSize        =   -1  'True
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   165
            TabIndex        =   18
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.CheckBox chk��� 
         Caption         =   "��ʾ����˷���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   12
         Top             =   615
         Width           =   2064
      End
      Begin VB.CommandButton cmdAllCls 
         Caption         =   "ȫ��(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1335
         TabIndex        =   15
         Top             =   4500
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllSel 
         Caption         =   "ȫѡ(&A)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   14
         Top             =   4500
         Width           =   1200
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6690
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   150
         Width           =   1815
      End
      Begin VB.TextBox txtOld 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   150
         Width           =   585
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   600
      End
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1305
         MaxLength       =   100
         TabIndex        =   4
         Top             =   150
         Width           =   2040
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFee 
         Height          =   3120
         Left            =   108
         TabIndex        =   13
         Top             =   948
         Width           =   9816
         _cx             =   17314
         _cy             =   5503
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeVerifyManger.frx":1935
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
         ExplorerBar     =   2
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
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8610
         TabIndex        =   16
         Top             =   4395
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   5700
         TabIndex        =   21
         Top             =   585
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   8670
         TabIndex        =   22
         Top             =   585
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "������˺ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   4170
         Width           =   1560
      End
      Begin VB.Label lbl���� 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4710
         TabIndex        =   24
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8385
         TabIndex        =   23
         Top             =   675
         Width           =   120
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5970
         TabIndex        =   11
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4740
         TabIndex        =   10
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   9
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   150
         TabIndex        =   5
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   495
      ScaleHeight     =   5010
      ScaleWidth      =   9510
      TabIndex        =   0
      Top             =   2280
      Width           =   9510
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   4995
         Left            =   195
         TabIndex        =   1
         Top             =   465
         Width           =   9510
         _Version        =   589884
         _ExtentX        =   16775
         _ExtentY        =   8811
         _StockProps     =   64
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   1170
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmFeeVerifyManger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlngModule As Long, mstrPrivs As String
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private Enum mPgIndex
    pg_����תסԺ = 1
End Enum
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnValid As Boolean
Private mblnMultiBalance As Boolean
Private Enum ҽԺҵ��
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    support�൥���շѱ���ȫ�� = 39  '�൥���շѱ���ȫ��
End Enum
Private mrsFeeList As ADODB.Recordset
Private mblnNotClick As Boolean
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------

 Private Sub cbo�շѵ�_Click()
    If mblnNotClick Then Exit Sub
    If mrsFeeList Is Nothing Then Exit Sub
    ReadListData True
End Sub

Private Sub chk���_Click()
    If mrsFeeList Is Nothing Then Exit Sub
    ReadListData True
End Sub

Private Sub chk��ת������_Click()
    If mrsFeeList Is Nothing Then Exit Sub
    ReadListData True
End Sub

Private Sub cmdALLCls_Click()
   Dim i As Long
    With vsFee
        '40526
        If .Rows <= 1 Or .Cols <= 0 Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" Then
               If Val(.TextMatrix(i, .ColIndex("ת����־"))) = 0 Then
                    .TextMatrix(i, .ColIndex("���")) = 0
                End If
            End If
        Next
        Call SetSumMoney(True)
    End With
End Sub
Private Sub cmdAllSel_Click()
    Dim i As Long
    With vsFee
        '40526
        If .Rows <= 1 Or .Cols <= 0 Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" Then
                If Val(.TextMatrix(i, .ColIndex("ת����־"))) = 0 Then
                    If CheckIsInput(i) = True Then
                        .TextMatrix(i, .ColIndex("���")) = -1
                        SetRowSelected (i)
                    End If
                End If
            End If
        Next
        Call SetSumMoney
    End With
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub cmdBrush_Click()
    If mrsInfo Is Nothing Then
        MsgBox "����ѡ����,����!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    ElseIf mrsInfo.State <> 1 Then
        MsgBox "����ѡ����,����!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    Call ReadListData
End Sub

Private Sub cmdOK_Click()
    If SaveData = False Then
        stbThis.Panels(2).Text = "����ʧ��!"
        Exit Sub
    End If
    Call ReadListData
    mblnChange = False
    stbThis.Panels(2).Text = "����ɹ�!"
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
      If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ز���
    '����:���˺�
    '����:2011-02-09 11:46:35
    '---------------------------------------------------------------------------------------------------------------------------------------------


End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Parameter
            If frmFeeVrerfyParaSet.ShowMe(Me, mlngModule, mstrPrivs) = False Then Exit Sub
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
    Case conMenu_View_Refresh
        Call cmdBrush_Click
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call zlCallCustomReprot(Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Err = 0: On Error Resume Next
    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With picList
        .Left = lngLeft + 50: .Top = lngTop
        .Width = lngRight - 100
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_����תסԺ Then
            Control.Enabled = Trim(vsFee.TextMatrix(1, vsFee.ColIndex("���ݺ�"))) <> ""
        Else
            Control.Enabled = False
        End If
    Case conMenu_View_Refresh
    End Select
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2011-01-25 15:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(mPgIndex.pg_����תסԺ, "����תסԺ����", picMzToZy.hWnd, 0)
    objItem.Tag = mPgIndex.pg_����תסԺ
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
    stbThis.Top = Me.ScaleHeight - Me.stbThis.Height
End Sub
Private Sub Form_Activate()
    Dim strKey As String
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnChange = False
    
End Sub

Private Sub Form_Load()
    Dim i As Long
    RestoreWinState Me, App.ProductName
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    dtpEnd.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtpEnd.Value = dtpEnd.MaxDate
    dtpBegin.MaxDate = dtpEnd.MaxDate
    dtpBegin.Value = dtpEnd.Value - 7
    
    Call initCardSquareData
    i = Val(zlDatabase.GetPara("������˵���", glngSys, mlngModule, 2, Array(cbo�շѵ�, lblBill), InStr(1, mstrPrivs, ";��������;") > 0))
    mblnFirst = True
    With cbo�շѵ�
        mblnNotClick = True
        .AddItem "�շѵ�"
        .ItemData(.NewIndex) = 0
        If i = 0 Then .ListIndex = .NewIndex
        .AddItem "���ʵ�"
        .ItemData(.NewIndex) = 1
        If i = 1 Then .ListIndex = .NewIndex
        .AddItem "�շѵ��ͼ��ʵ�"
        .ItemData(.NewIndex) = 2
        If i = 2 Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = .NewIndex
        mblnNotClick = False
    End With
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call zlDefCommandBars
    Call InitPage
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    chk���.Value = IIf(Val(zlDatabase.GetPara("��ʾ����˷���", glngSys, mlngModule, 0, Array(chk���), InStr(1, mstrPrivs, ";��������;") > 0)) = 0, 0, 1)
    chk��ת������.Value = IIf(Val(zlDatabase.GetPara("��ʾ��ת������", glngSys, mlngModule, 0, Array(chk��ת������), InStr(1, mstrPrivs, ";��������;") > 0)) = 0, 0, 1)

    Set mrsInfo = New ADODB.Recordset
    vsFee.OwnerDraw = flexODContent
    '���ŵ���ʹ�ö��ֽ��㷽ʽģʽ
    mblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
    Call zlCreateObject
 End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If mblnChange Then
        If MsgBox("ע��:" & vbCrLf & "    ���޸�������,���㻹δ����,�Ƿ����Ҫ�˳�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    zlDatabase.SetPara "��ʾ����˷���", chk���.Value, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "������˵���", cbo�շѵ�.ListIndex, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "��ʾ��ת������", chk��ת������.Value, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, "��ϸ�б�", True
    SaveWinState Me, App.ProductName
    
    Call zlCloseObject
    Set mrsFeeList = Nothing
    Set mrsInfo = Nothing
End Sub
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picMzToZy_Resize()
    Err = 0: On Error Resume Next
    With picMzToZy
        cmdAllCls.Top = .ScaleHeight - cmdAllCls.Height - 50
        cmdAllSel.Top = cmdAllCls.Top
        cmdOk.Top = cmdAllCls.Top
        cmdOk.Left = .ScaleWidth - cmdOk.Width - vsFee.Left * 2
        lblSum.Top = cmdAllCls.Top - lblSum.Height - 30
        
        vsFee.Width = .ScaleWidth - vsFee.Left * 2
        vsFee.Height = lblSum.Top - vsFee.Top - 50
         'chk���.Left = .ScaleWidth - chk���.Width
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   If Val(tbPage.Selected.Tag) = mPgIndex.pg_����תסԺ Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Else
        Exit Sub
    End If
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
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
    End With
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2011-01-25 15:14:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid, rsTemp As New ADODB.Recordset, strSQL As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "����������"
    objRow.Add "���ˣ�" & txtPatient.Text
    objRow.Add "�Ա�" & txtSex.Text
    objRow.Add "���䣺" & txtOld.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set vsGrid = vsFee
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("ѡ��") Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = vsGrid
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub
Private Sub zlCallCustomReprot(ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ص��Զ��屨��
    '����:���˺�
    '����:2011-01-25 15:16:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As Variant, lng����ID As Long
    With vsFee
        If .Row > 0 Then
            strNO = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�")))
            lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("���ݺ�")))
        End If
        If strNO <> "" Then
            Call ReportOpen(gcnOracle, lngSys, strReprotName, Me, _
                "NO=" & strNO, "����ID=" & lng����ID)
        Else
            Call ReportOpen(gcnOracle, lngSys, strReprotName, Me)
        End If
    End With
End Sub
Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
    stbThis.Panels(2).Text = ""
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
'        With frmPatiSelect
'            If (mbytUseType = 0 Or mbytUseType = 1) Then
'                .mlngUnitID = mlngUnitID
'            Else
'                .mlngUnitID = mlngDeptID
'            End If
'            .mbytUseType = mbytUseType
'            .mstrPrivs = mstrPrivs
'            Set .mfrmParent = Me
'            .Show 1, Me
'        End With
    Else
        If IDKind.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.IDKind = IDKind.GetKindIndex("�����") Or IDKind.IDKind = IDKind.GetKindIndex("סԺ��") Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
         Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
            KeyAscii = 0
            'ˢ�²�����Ϣ:"-����ID"
            Call GetPatient(IDKind.GetCurCard, txtPatient.Tag, False)
            If mrsInfo.State = 0 Then   '
                txtPatient.Text = "": txtOld.Text = ""
                txtSex.Text = "": txtסԺ��.Text = ""
                Exit Sub
            End If
            Call ReadListData
            Exit Sub
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnICCard As Boolean, blnMsg As Boolean, blnIDCard As Boolean
    
   '54899
    If objCard.���� Like "IC��*" And objCard.ϵͳ = True And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ = True And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        If blnCard Then
            If Not blnMsg Then MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����", vbInformation, gstrSysName
            txtPatient.Text = "": txtOld.Text = ""
            txtסԺ��.Text = ""
            vsFee.Clear 1
            vsFee.Rows = 2
            Exit Sub
        End If
        If Not blnMsg Then MsgBox "���ܶ�ȡ������Ϣ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtPatient
        txtOld.Text = "": txtSex.Text = "": txtסԺ��.Text = ""
        vsFee.Clear 1
        vsFee.Rows = 2
        Exit Sub
    End If
    
    '��ȡ�ɹ�
    '���￨������
     If (objCard.���� Like "IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ = True And blnCard Then blnCard = False
     If Mid(gstrCardPass, 6, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
            vsFee.Clear 1
            vsFee.Rows = 2
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    Call ReadListData
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub
Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����: blnOutMsg-�Ѿ���ʾ,�������ⲿ����ʾ
    '����:
    '����:���˺�
    '����:2011-01-25 16:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, dtDate As Date, vRect As RECT
    Dim strPati As String, blnCancel As Boolean
    
    On Error GoTo errH
    '��˱�־:50459
    strSQL = _
        "Select A.����ID,Nvl(B.��ҳID,0) as ��ҳID,A.����� as �����,A.��ǰ����,B.��Ժ����," & _
        "       Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(B.����,A.����) as ����,A.IC����,A.���￨��,A.����֤��," & _
        "       Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,C.���� as ��ǰ����,A.��ǰ����ID,D.���� as ��Ժ����,B.��Ժ����ID, A.���� as ����,E.����,E.ҽ����,E.����," & _
        "       A.�Ǽ�ʱ��,Nvl(B.״̬,0) as ״̬,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Nvl(B.��˱�־,0) as ��˱�־,B.��Ժ����,B.��Ժ����,B.��������,B.��������" & _
        " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,ҽ�����˵��� E,ҽ�����˹����� F" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+) And A.��ҳID=B.��ҳID(+) " & _
        "           And A.����ID=F.����ID(+) And F.��־(+)=1 And F.ҽ����=E.ҽ����(+) And F.����=E.����(+) And F.���� = E.����(+)" & _
        "           And A.��ǰ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+)" & _
        "           And A.ͣ��ʱ�� is NULL "
    
    If blnCard = True And objCard.���� Like "����*" Then  'ˢ��
        lng�����ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strSQL = strSQL & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
    Else '��������
        Select Case objCard.����
            Case "����", "��������￨"
                If mrsInfo.State = 1 Then
                    If mrsInfo!���� = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                End If
                '53816
                'ͨ����������
                strPati = "Select A.����ID as ID,A.����ID,A.סԺ��, A.�����, Nvl(b.�Ա�, a.�Ա�) as �Ա�, A.����, A.סԺ����, A.��ͥ��ַ, A.������λ," & vbNewLine & _
                        "To_Char(A.��������,'YYYY-MM-DD') as ��������,  To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����, To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����" & vbNewLine & _
                        "From ������Ϣ A, ������ҳ B" & vbNewLine & _
                        "Where A.����id = B.����id(+) And A.��ҳID = B.��ҳid(+) And A.ͣ��ʱ�� Is Null And A.���� = [1] " & vbNewLine & strPati & vbNewLine & _
                        "Order By Decode(סԺ��, Null, 1, 0), ��Ժ���� Desc"
                        
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If mrsInfo Is Nothing Or blnCancel Then
                    txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
                    txtסԺ��.Text = ""
                    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                    Set mrsInfo = New ADODB.Recordset
                    Exit Function
                End If
                strInput = "-" & Val(mrsInfo!����ID)
                strSQL = strSQL & " And A.����ID=[1]"
                    
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    If Not mrsInfo.EOF Then
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        txtPatient.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
        txtPatient.Text = Nvl(mrsInfo!����): txtOld.Text = Nvl(mrsInfo!����): txtSex.Text = Nvl(mrsInfo!�Ա�)
        txtסԺ��.Text = Nvl(mrsInfo!�����)
        If IsNull(mrsInfo!��Ժ����) Then
            
            dtDate = zlDatabase.Currentdate '53816
            dtpEnd.MaxDate = Format(dtDate, "yyyy-mm-dd 23:59:59")
            dtpEnd.Value = dtDate
        Else
            dtDate = CDate(Format(mrsInfo!��Ժ����, "yyyy-mm-dd HH:MM:SS"))
            If dtDate > dtpEnd.MaxDate Then dtpEnd.MaxDate = dtDate
            
            dtpEnd.Value = Format(mrsInfo!��Ժ����, "yyyy-mm-dd HH:MM:SS")
            dtpEnd.MaxDate = dtpEnd.Value + 1
            dtpBegin.MaxDate = dtpEnd.Value
            '   ����:36609 ����Ժʱ��Ҫ��һ��,��Ϊ���ܴ��ڲ�����û���������ʱ,����Ժ,��ȥ�������,�Ӷ�����������ת���˵����.
        End If
    
        If dtpBegin.Value > dtpEnd.Value Then
            dtpBegin.Value = dtpEnd.Value - 7   '��ȥ7��
        End If
        GetPatient = True
        Exit Function
    Else
        txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
        txtסԺ��.Text = ""
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
NotFoundPati:
    txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
    txtסԺ��.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Set mrsInfo = New ADODB.Recordset
End Function

Private Function ReadListData(Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ��˵���ϸ����
    '����:��ȡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, dtEndDate As Date, dtStartDate As Date
    Dim strWhere As String, strInsure As String
    
    dtEndDate = dtpEnd.Value: dtStartDate = dtpBegin.Value
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    If dtEndDate - dtStartDate > 30 Then    '����30��,�򰴲���IDΪ��������
        strWhere = " And A.����ID=[1] And (A.����ʱ��+0 between [2] and [3] )"
        strInsure = " And ����ID=[1] And (����ʱ��+0 between [2] and [3] )"
    Else
        strWhere = " And A.����ID+0=[1] And (A.����ʱ�� between [2] and [3] )"
        strInsure = " And ����ID+0=[1] And (����ʱ�� between [2] and [3] )"
    End If
    
    On Error GoTo errHandle
    If blnFilter = False Then zlCommFun.ShowFlash "���ڶ�ȡ�շѵ���,���Ժ� ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    
    strTable = "" & _
    "       Select Nvl(Max(��¼״̬), 0) As ת����־, Decode(Max(����), 0, '', '��') As ҽ��," & vbNewLine & _
    "              Max(Decode(�۸񸸺�, Null, Decode(��¼״̬, 2, To_Number(Null), ID), To_Number(Null))) As ID, '�շѵ�' As ����, NO, ʵ��Ʊ��," & vbNewLine & _
    "              ���, �շ����, ��������, �շ�ϸĿid, ִ�в���id, Avg(Nvl(����, 1)) As ����, Sum(����) ����, ��׼���� As ����, Sum(Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
    "              Sum(ʵ�ս��) As ʵ�ս��, ������, To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, Max(����) As ����, Max(�����) As �����," & vbNewLine & _
    "              To_Char(Max(�������), 'YYYY-MM-DD HH24:MI:SS') As �������, Min(����id) As ����id" & vbNewLine & _
    "       From (Select a.Id, m.��¼״̬, Nvl(b.����, 0) As ����, a.�۸񸸺�, a.No, a.ʵ��Ʊ��, a.���, a.�շ����, a.��������, a.�շ�ϸĿid, a.ִ�в���id, a.����," & vbNewLine & _
    "                     a.����, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��, a.������, a.����ʱ��, m.�����, m.�������, a.����id" & vbNewLine & _
    "              From ������ü�¼ A," & vbNewLine & _
    "                   (Select Distinct ��¼id, ����" & vbNewLine & _
    "                     From ���ս����¼" & vbNewLine & _
    "                     Where ���� = 1" & strInsure & ") B, ������˼�¼ M" & vbNewLine & _
    "              Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 " & strWhere & vbNewLine & _
    "                    And a.����id = b.��¼id(+) And a.Id = m.����id(+) And" & vbNewLine & _
    "                    m.����(+) = 1 And Nvl(a.���ӱ�־, 0) <> 9)" & vbNewLine & _
    "       Group By NO, ʵ��Ʊ��, ���, ��׼����, �շ����, �շ�ϸĿid, ��������, ִ�в���id, ������, ����ʱ��" & vbNewLine & _
    "       Having Sum(����) <> 0"

    
    strTable = strTable & "Union ALL " & _
    "       Select Nvl(Max(��¼״̬), 0) As ת����־, Decode(Max(����), 0, '', '��') As ҽ��," & vbNewLine & _
    "              Max(Decode(�۸񸸺�, Null, Decode(��¼״̬, 2, To_Number(Null), ID), To_Number(Null))) As ID, '�շѵ�' As ����, NO, ʵ��Ʊ��," & vbNewLine & _
    "              ���, �շ����, ��������, �շ�ϸĿid, ִ�в���id, Avg(Nvl(����, 1)) As ����, Sum(����) ����, ��׼���� As ����, Max(Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
    "              Max(ʵ�ս��) As ʵ�ս��, ������, To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, Max(����) As ����, Max(�����) As �����," & vbNewLine & _
    "              To_Char(Max(�������), 'YYYY-MM-DD HH24:MI:SS') As �������, Min(����id) As ����id" & vbNewLine & _
    "       From (Select a.Id, m.��¼״̬, Nvl(b.����, 0) As ����, a.�۸񸸺�, a.No, a.ʵ��Ʊ��, a.���, a.�շ����, a.��������, a.�շ�ϸĿid, a.ִ�в���id, a.����," & vbNewLine & _
    "                     a.����, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��, a.������, a.����ʱ��, m.�����, m.�������, a.����id, a.���ʽ��" & vbNewLine & _
    "              From ������ü�¼ A," & vbNewLine & _
    "                   (Select Distinct ��¼id, ����" & vbNewLine & _
    "                     From ���ս����¼" & vbNewLine & _
    "                     Where ���� = 1" & strInsure & ") B, ������˼�¼ M" & vbNewLine & _
    "              Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 " & strWhere & vbNewLine & _
    "                    And a.����id = b.��¼id(+) And Exists" & vbNewLine & _
    "               (Select 1" & vbNewLine & _
    "                     From ������ü�¼ J, ������˼�¼ K" & vbNewLine & _
    "                     Where j.Id = k.����id And j.����id = [1] And k.���� = 1 And j.No = a.No And j.��� = a.��� And" & vbNewLine & _
    "                           Mod(j.��¼����, 10) = 1) And a.Id = m.����id(+) And m.����(+) = 1 And Nvl(a.���ӱ�־, 0) <> 9)" & vbNewLine & _
    "       Group By NO, ʵ��Ʊ��, ���, ��׼����, �շ����, �շ�ϸĿid, ��������, ִ�в���id, ������, ����ʱ��" & vbNewLine & _
    "       Having Sum(����) = 0"

    strTable = strTable & "Union ALL " & _
    " Select  nvl(max(M.��¼״̬),0) as ת����־,Decode(NULL,Null,'','��') as ҽ��,Max(decode(A.�۸񸸺�,NULL,ID,0))  as ID, " & _
    "       '���ʵ�' as ����,A.No,A.ʵ��Ʊ��, A.��� as ���,A.�շ����,A.��������,A.�շ�ϸĿID,A.ִ�в���ID, " & _
    "       Avg(Nvl(A.����,1)) as ����, Sum(A.����) ����, A.��׼���� as ����,Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "       A.������,To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, NULL as ����," & vbNewLine & _
    "       Max(M.�����) as �����,To_Char(Max(M.�������), 'YYYY-MM-DD HH24:MI:SS') As �������,Null As ����ID " & vbNewLine & _
    "From ������ü�¼ A,������˼�¼ M" & vbNewLine & _
    "Where A.��¼���� = 2 And A.��¼״̬ <> 0 " & strWhere & _
    "           And A.ID = M.����ID(+) And M.����(+)=1 " & vbNewLine & _
    "Group By A.NO, A.ʵ��Ʊ��,A.���,A.�շ����,A.�շ�ϸĿID,A.��������,A.��׼����,A.ִ�в���id," & _
    "       A.������, A.����ʱ�� Having Sum(A.����) <> 0" & vbNewLine
    
    strTable = strTable & "Union ALL " & _
    " Select  nvl(max(M.��¼״̬),0) as ת����־,Decode(NULL,Null,'','��') as ҽ��,Max(decode(A.�۸񸸺�,NULL,ID,0))  as ID, " & _
    "       '���ʵ�' as ����,A.No,A.ʵ��Ʊ��, A.��� as ���,A.�շ����,A.��������,A.�շ�ϸĿID,A.ִ�в���ID, " & _
    "       Avg(Nvl(A.����,1)) as ����, Sum(A.����) ����, A.��׼���� as ����,Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "       A.������,To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, NULL as ����," & vbNewLine & _
    "       Max(M.�����) as �����,To_Char(Max(M.�������), 'YYYY-MM-DD HH24:MI:SS') As �������,Null As ����ID " & vbNewLine & _
    "From ������ü�¼ A,������˼�¼ M" & vbNewLine & _
    "Where A.��¼���� = 2 And A.��¼״̬ <> 0 " & strWhere & _
    "           And A.ID = M.����ID(+) And M.����(+)=1 And Exists" & vbNewLine & _
    "               (Select 1" & vbNewLine & _
    "                     From ������ü�¼ J, ������˼�¼ K" & vbNewLine & _
    "                     Where j.Id = k.����id And j.����id = [1] And k.���� = 1 And j.No = a.No And j.��� = a.��� And" & vbNewLine & _
    "                           j.��¼���� = 2) And a.Id = m.����id(+) And m.����(+) = 1" & vbNewLine & _
    "Group By A.NO, A.ʵ��Ʊ��,A.���,A.�շ����,A.�շ�ϸĿID,A.��������,A.��׼����,A.ִ�в���id," & _
    "       A.������, A.����ʱ�� Having Sum(A.����) = 0" & vbNewLine
    
    
    strSQL = "" & _
    " Select  A.ID,A.ת����־,decode(A.�����,NULL,0,-1) as ���,A.����,A.No as ���ݺ�,A.ʵ��Ʊ�� As Ʊ�ݺ�, " & _
    "       A.���,A.��������,A.�շ�ϸĿID,A.ִ�в���ID,A.�շ����,P.���, " & _
    "       C.���� as ����,C.����||'-'||Nvl(B.����,C.����) as ����,E1.���� as ��Ʒ��,C.���," & _
    "       A.����, A.����,C.���㵥λ," & _
    "       ltrim(to_char(A.����,'9999990.00000')) as ����," & _
    "       ltrim(to_char(A.Ӧ�ս��,'9999990.00')) as Ӧ�ս��," & _
    "       ltrim(to_char(A.ʵ�ս��,'9999990.00')) as ʵ�ս��," & _
    "       A.������,A.����ʱ��,A.ҽ��, A.����,A.�����,A.�������,A.����ID" & vbNewLine & _
    "From (" & strTable & ") A,�շ���ĿĿ¼ C,�շ���Ŀ���� B,�շ���Ŀ���� E1,�շ���� P" & _
    " Where A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=B.�շ�ϸĿID(+)  And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       and A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
    "       And A.�շ����=P.����(+)" & _
    " Order by A.����,A.NO,A.���"

    If mrsFeeList Is Nothing Or blnFilter = False Then
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, dtStartDate, dtEndDate)
    Else
        mrsFeeList.Filter = 0
    End If
    strFilter = IIf(cbo�շѵ�.ItemData(cbo�շѵ�.ListIndex) = 2, "", " And ����='" & cbo�շѵ�.Text & "'")
    strFilter = strFilter & IIf(chk���.Value = 1, "", " And  ���=0")
    strFilter = strFilter & IIf(chk��ת������.Value = 1, "", " And  ת����־=0")
    mrsFeeList.Filter = Mid(strFilter, 5)
    vsFee.Redraw = flexRDNone
    mblnNotClick = True
    vsFee.Clear: vsFee.Cols = 1: vsFee.Rows = 2: vsFee.FixedRows = 1
    mblnNotClick = False
    Set vsFee.DataSource = mrsFeeList
    If vsFee.Rows <= 1 Then vsFee.Rows = 2
    With vsFee
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",����,����,���,��������,ת����־,�շ����,����ID,", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .ColDataType(.ColIndex("���")) = flexDTBoolean
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsFee, Me.Caption, "��ϸ�б�", True
        If gTy_System_Para.bytҩƷ������ʾ <> 2 Then    '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
            .ColHidden(.ColIndex("��Ʒ��")) = True
        End If
        '����
        Dim strNO As String, str���� As String
        
        For lngRow = 1 To .Rows - 1
            If strNO <> Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) _
                And str���� = Trim(.TextMatrix(lngRow, .ColIndex("����"))) And strNO <> "" Then
                '�����ָ���
                .CellBorderRange lngRow, .FixedCols, lngRow, .Cols - 1, vbBlue, 0, 1, 0, 0, 0, 0
            End If
            If str���� <> Trim(.TextMatrix(lngRow, .ColIndex("����"))) And str���� <> "" Then
                .CellBorderRange lngRow, .FixedCols, lngRow, .Cols - 1, vbRed, 0, 1, 0, 0, 0, 0
            End If
            .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
            .Cell(flexcpData, lngRow, .ColIndex("���")) = Val(.TextMatrix(lngRow, .ColIndex("���")))
            If Val(.TextMatrix(lngRow, .ColIndex("���"))) <> 0 Then
                Select Case Val(.TextMatrix(lngRow, .ColIndex("ת����־")))
                Case 0
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &HFF0000       '��ɫ
                Case 1, 2
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000015
'                Case 2
'                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000010
                End Select
            End If
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
            str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
        Next
        .Editable = flexEDKbdMouse
    End With
    If blnFilter = False Then zlCommFun.StopFlash
    Call SetSumMoney
    Call StatusShowBillSum
    vsFee.Redraw = flexRDDirect
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsFee.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   If blnFilter = False Then zlCommFun.StopFlash
End Function
Private Sub vsFee_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsFee
        Select Case Col
        Case .ColIndex("���")
                SetNOBill .TextMatrix(Row, .ColIndex("����")), .TextMatrix(Row, .ColIndex("���ݺ�")), Val(.TextMatrix(Row, .Col)) <> 0
                Call SetRowSelected(Row)
                mblnChange = True
                Call SetSumMoney
        Case Else
        End Select
    End With
End Sub
Private Sub vsFee_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, "��ϸ�б�", True
End Sub

Private Sub vsFee_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNotClick = True Then Exit Sub
    If OldRow <> NewRow Then
        Call StatusShowBillSum
    End If
End Sub

Private Sub vsFee_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, "��ϸ�б�", True
End Sub

Private Sub vsFee_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsFee
        Select Case Col
        Case .ColIndex("���")
            If Val(.TextMatrix(Row, .ColIndex("ת����־"))) <> 0 Then
                stbThis.Panels(2).Text = "��������Ѿ�ת��,���ܸ������״̬"
                Cancel = True: Exit Sub
            End If
            
            If GetVsGridBoolColVal(vsFee, Row, Col) Then
                If InStr(1, mstrPrivs, ";ȡ���������;") = 0 And .TextMatrix(Row, .ColIndex("�����")) <> UserInfo.���� And .TextMatrix(Row, .ColIndex("�����")) <> "" Then
                    stbThis.Panels(2).Text = "��û��Ȩ��ȡ��������˵ķ���"
                    Cancel = True: Exit Sub
                End If
            End If
            If CheckIsInput(Row) = False Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsFee_DblClick()
     '   vsFee.TextMatrix(vsFee.Row, vsFee.ColIndex("���")) = "��"
End Sub

Private Sub vsFee_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݻ��ߺ������
    '����:���˺�
    '����:2011-01-26 09:57:32
    '˵��:
    '       1.OwnerDrawҪ����ΪOver(������Ԫ��������)
    '       2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
    '       3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    Dim strText As String
    strText = " "
    With vsFee
        '����������еı��߼�����
        lngLeft = .ColIndex("���"): lngRight = .ColIndex("���")
        
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        Call GetBillNOStartAndEndRow(Row, lngBegin, lngEnd)
        If lngBegin = lngEnd Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, strText, 1, 0
        Done = True
    End With
End Sub
Private Sub GetBillNOStartAndEndRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������
    '����:���˺�
    '����:2011-01-26 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    lngBegin = lngRow: lngEnd = lngRow
    With vsFee
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�")) Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�")) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub
Private Function SetNOBill(ByVal str���� As String, ByVal strNO As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ȫѡ��ȫ�嵥��
    '���:str����-��������(�շѵ�,���ʵ�)
    '       strNO-ָ����NO
    '        blnSel:true��ʾȫѡ,����ȫ��
    '����:
    '����:
    '����:���˺�
    '����:2011-01-24 10:47:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsFee
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" _
                And .TextMatrix(i, .ColIndex("���ݺ�")) = strNO _
                And .TextMatrix(i, .ColIndex("����")) = str���� Then
                .TextMatrix(i, .ColIndex("���")) = IIf(blnSel, -1, 0)
            End If
        Next
    End With
    SetNOBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:��˳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-26 13:31:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim cllProc As Collection, cllTemp As Collection, i As Long
    Dim blnSel As Boolean, lngRow As Long, strDate As String
    If mrsInfo Is Nothing Or mrsInfo.State = 0 Then Exit Function
    Set cllProc = New Collection: Set cllTemp = New Collection
    If mrsInfo Is Nothing Or mrsInfo.State = 0 Then Exit Function
    If Val(Nvl(mrsInfo!��ҳID)) <> 0 Then
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = False Then
            Exit Function
        End If
    End If
    
    '�ȴ���ȡ����˲���
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    With vsFee
        If .Rows <= 1 Then Exit Function
        If .Cols <= 1 Then Exit Function
        
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("ID"))) <> 0 Then
                blnSel = GetVsGridBoolColVal(vsFee, lngRow, .ColIndex("���"))
                If Val(.Cell(flexcpData, lngRow, .ColIndex("���"))) <> 0 And Not blnSel Then
                    ' Zl_������˼�¼_Delete
                    strSQL = "Zl_������˼�¼_Delete("
                    '  ����id_In In ������˼�¼.����id%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(lngRow, .ColIndex("ID"))) & ","
                    '  ����_In   In ������˼�¼.����%Type
                    strSQL = strSQL & "1)"
                    zlAddArray cllProc, strSQL
                ElseIf Val(.Cell(flexcpData, lngRow, .ColIndex("���"))) = 0 And blnSel Then
                    '����
                    'Zl_������˼�¼_Insert
                    strSQL = "Zl_������˼�¼_Insert("
                    '  ����_In     In ������˼�¼.����%Type,
                    strSQL = strSQL & "" & 1 & ","
                    '  ����id_In   In ������˼�¼.����id%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(lngRow, .ColIndex("ID"))) & ","
                    '  ����id_In   In ������˼�¼.����id%Type,
                    strSQL = strSQL & "" & Val(Nvl(mrsInfo!����ID)) & ","
                    '  ��ҳid_In   In ������˼�¼.��ҳid%Type,
                    strSQL = strSQL & "" & IIf(Val(Nvl(mrsInfo!��ҳID)) = 0, "Null", Val(Nvl(mrsInfo!��ҳID))) & ","
                    '  �����_In   In ������˼�¼.�����%Type,
                    strSQL = strSQL & "'" & UserInfo.���� & "',"
                    '  �������_In In ������˼�¼.�������%Type
                    strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'))"
                    zlAddArray cllTemp, strSQL
                End If
            End If
        Next
    End With
    If cllTemp.Count = 0 And cllProc.Count = 0 Then
        MsgBox "δѡ����صĵ�����Ŀ,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    For i = 1 To cllTemp.Count
        zlAddArray cllProc, cllTemp(i)
    Next
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllProc, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetLocaleNO(ByVal str���� As String, ByVal strNO As String, ByVal blnSelect As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����NO
    '����:���˺�
    '����:2011-02-09 14:56:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    With vsFee
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) = strNO _
                And Trim(.TextMatrix(lngRow, .ColIndex("����"))) = str���� Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(blnSelect, -1, 0)
            End If
        Next
    End With
End Sub
Private Function CheckIsInput(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ������������
    '���:lngRow-ָ������
    '����:
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-09 15:04:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim lng����ID As Long, str���� As String
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    lng����ID = Val(Nvl(mrsInfo!����ID))
    With vsFee
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
            strNO = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
            str���� = .TextMatrix(lngRow, .ColIndex("����"))
            If intInsure > 0 And str���� = "�շѵ�" Then
                If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure) Then
                    stbThis.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧�������������,���в�����ѡ��ת��!"
                    Exit Function
                Else
                    '���жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                    strTmp = GetBalanceType(strNO)
                    If strTmp <> "" Then
                        arrBalanceType = Split(strTmp, ",")
                        For i = 0 To UBound(arrBalanceType)
                            strBalanceType = arrBalanceType(i)
                            If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure, strBalanceType) Then
                                stbThis.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                Exit Function
                            End If
                        Next
                    End If
                End If
            End If
    End With
    CheckIsInput = True
End Function
Private Function SetRowSelected(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ�е�ѡ��״̬
    '       ����Ƕ��ŵ����е�һ��,����ͬʱ���ö����е���������
    '����:���˺�
    '����:2011-02-09 14:50:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim blnSelect As Boolean, lng����ID As Long, str���� As String
    lng����ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
    End If
    With vsFee
        intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
        blnSelect = GetVsGridBoolColVal(vsFee, lngRow, .ColIndex("���"))
        str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
        If intInsure > 0 And str���� = "�շѵ�" Then 'ȫ��ѡ���ȡ��
            If gclsInsure.GetCapability(support�൥���շѱ���ȫ��, lng����ID, intInsure) Or Not IsYBSingle(.TextMatrix(lngRow, .ColIndex("���ݺ�")), intInsure) Then
                If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
            End If
        Else '�ֽ�����Ҫ����൥���շ����
            If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
        End If
    End With
    SetRowSelected = True
End Function

Private Function IsYBSingle(ByVal strNO As String, ByVal intInsure As Integer) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset, blnInsureSingle As Boolean
    
    blnInsureSingle = gclsInsure.GetCapability(83, , intInsure)
    If blnInsureSingle = False Then
        IsYBSingle = False
        Exit Function
    Else
        strSQL = "Select 1 From ҽ��������ϸ Where NO = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.EOF Then
            IsYBSingle = False
        Else
            If CheckAllTurn(strNO) Then
                IsYBSingle = False
            Else
                IsYBSingle = True
            End If
        End If
    End If
    
End Function

Private Function GetBalanceType(ByVal strNO As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ�ŵ����е�ҽ�����㷽ʽ��
    '����:ҽ�����㷽ʽ��
    '����:���˺�
    '����:2011-02-09 15:01:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
    On Error GoTo errH
    strSQL = "Select A.���㷽ʽ From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
            "Where A.���㷽ʽ = B.���� And B.���� In (3, 4) And A.NO = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    For i = 1 To rsTmp.RecordCount
        GetBalanceType = GetBalanceType & "," & rsTmp!���㷽ʽ
        rsTmp.MoveNext
    Next
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAllTurn(ByVal strNO As String) As Boolean
    Dim strSQL As String, rsData As ADODB.Recordset
    strSQL = "Select 1" & vbNewLine & _
            " From ����Ԥ����¼ A," & vbNewLine & _
            "     (Select Distinct ����id" & vbNewLine & _
            "       From ������ü�¼" & vbNewLine & _
            "       Where NO In (Select Distinct NO" & vbNewLine & _
            "                    From ������ü�¼" & vbNewLine & _
            "                    Where ����id In" & vbNewLine & _
            "                          (Select ����id" & vbNewLine & _
            "                           From ����Ԥ����¼" & vbNewLine & _
            "                           Where ������� In (Select b.�������" & vbNewLine & _
            "                                          From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
            "                                          Where a.No = [1] And a.��¼���� = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))) And" & vbNewLine & _
            "             ��¼���� = 1 And ��¼״̬ <> 0) B" & vbNewLine & _
            " Where a.����id = b.����id And a.��¼���� = 3 And (Exists (Select 1 From ҽ�ƿ���� Where ID = a.�����id And �Ƿ�ȫ�� = 1) Or Exists" & vbNewLine & _
            "       (Select 1 From ���ѿ����Ŀ¼ Where ��� = a.���㿨��� And �Ƿ�ȫ�� = 1))" & vbNewLine & _
            " Group By ���㷽ʽ" & vbNewLine & _
            " Having Sum(��Ԥ��) <> 0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsData.EOF Then
        CheckAllTurn = False
    Else
        CheckAllTurn = True
    End If
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ŵ�������ѡ���ȡ��
    '       ���ҽ�����ŵ���Ҫ�������˷�,ѡ������һ��ʱ,ȫѡ����,ȡ��ʱȫȡ��
    '���:lngRow-��ǰ��
    '        blnSelect-�Ƿ�ѡ��
    '        intInsure-����
    '����:
    '����:���˺�
    '����:2011-02-09 15:41:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, k As Long, strNO As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim lng����ID As Long, str���� As String, blnAllTurn As Boolean
    lng����ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
    End If
    With vsFee
        str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
        If intInsure = 0 Then
            If CheckAllTurn(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) = True Then
                blnAllTurn = True
            Else
                blnAllTurn = False
            End If
            If mblnMultiBalance Or blnAllTurn Then     '   �൥��,���ֽ��㷽ʽ
                '33635:ԭ���Ƕ൥���Ҷ��ֽ��㷽ʽ,���ܲ�����
                strNO = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                        And Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) <> "" _
                        And .TextMatrix(k, .ColIndex("����")) = str���� Then
                          If InStr(1, "," & strNO & ",", "," & .TextMatrix(k, .ColIndex("���ݺ�")) & ",") = 0 Then
                                strNO = strNO & "," & .TextMatrix(k, .ColIndex("���ݺ�"))
                          End If
                      End If
                Next
                If strNO <> "" Then strNO = Mid(strNO, 2)
                If InStr(1, strNO, ",") > 0 Then    '֤��Ϊ�൥��
                    'һԺҪ��,ֻҪ�Ƕ൥�ݽ����,��תʱ,������ȫת
                    'If CheckSingleBalance(strNo) = False Then    '�Ƕ��ֽ��㷽ʽ,�������˷�,'ȫѡ
                        For k = 1 To .Rows - 1
                              If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                                  And Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) <> "" _
                                   And .TextMatrix(k, .ColIndex("����")) = str���� Then
                                    .TextMatrix(k, .ColIndex("���")) = IIf(blnSelect, -1, 0)
                              End If
                        Next
                    'End If
                End If
            End If
            '����Ƿ�������ѿ��Ľ���,�������,�ֲ�֧���ⲿ�����ݵĴ���
            If strNO = "" Then strNO = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
'            If str���� = "�շѵ�" Then
'                If zlIsExistsSquareCard(strNO) Then
'                    stbThis.Panels(2).Text = "�ݲ�֧�ֶ����ѿ����ݵ��������תסԺ����!"
'                    For k = 1 To .Rows - 1
'                          If .TextMatrix(k, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�")) And Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) <> "" Then
'                                .TextMatrix(k, .ColIndex("���")) = 0
'                          End If
'                    Next
'                End If
'            End If
            '����Ƿ�������ѿ�,����൥���д������ѿ�,Ҳ����ȫѡ
            SetMultiOther = True
            Exit Function
        End If
        If IsYBSingle(vsFee.TextMatrix(lngRow, .ColIndex("���ݺ�")), intInsure) Then SetMultiOther = True: Exit Function
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                And i <> lngRow And .TextMatrix(i, .ColIndex("����")) = str���� Then
                If GetVsGridBoolColVal(vsFee, i, .ColIndex("���")) <> GetVsGridBoolColVal(vsFee, lngRow, .ColIndex("���")) Then
                   If intInsure <> 0 And str���� = "�շѵ�" And blnSelect Then
                        strNO = .TextMatrix(i, .ColIndex("���ݺ�"))
                        '�жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                         strTmp = GetBalanceType(strNO)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure, strBalanceType) Then
                                     stbThis.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                     For k = 1 To .Rows - 1
                                        If .TextMatrix(k, .ColIndex("���ݺ�")) = .TextMatrix(i, .ColIndex("���ݺ�")) _
                                            And .TextMatrix(k, .ColIndex("����")) = .TextMatrix(i, .ColIndex("����")) Then
                                            .TextMatrix(k, .ColIndex("���")) = 0
                                        End If
                                     Next
                                     Exit Function
                                 End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, .ColIndex("���")) = IIf(blnSelect, -1, 0)
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function

Private Function CheckSingleBalance(ByVal strNO As String) As Boolean
'���ܣ��ж�ָ���������Ƿ�ֻ��һ�ַ�ҽ�����㷽ʽ(��Ԥ������)
'       :strNO(��ʽΪ"E01,E02"):����:34035
'������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strNO = Replace(strNO, "'", "")
    CheckSingleBalance = True
    
    strSQL = "" & _
    " Select /*+ rule */ Count(Distinct A.���㷽ʽ) num" & vbNewLine & _
    " From ����Ԥ����¼ A, ���㷽ʽ B,Table( f_Str2list([1])) J" & vbNewLine & _
    " Where   A.��¼���� = 3 And A.��¼״̬ In (1, 3) " & _
    "           And A.���㷽ʽ = B.���� And B.���� In (1, 2)  And A.NO = J.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNO)
    If rsTmp!Num > 1 Then CheckSingleBalance = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function zlIsExistsSquareCard(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ�Ϊ�����㵥��
    '���:strNos-���ݺ�(����Ϊ����,�ö��ŷ���)
    '����:
    '����:����,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    
    On Error GoTo errHandle
    
    strNoIns = Replace(strNos, "'", "")
    strSQL = "" & _
    "   Select /*+ rule */ A.ID As ������id " & _
    "   From ���˿������¼ A, ����Ԥ����¼ B,Table( f_Str2list([1])) J " & _
    "   Where A.����id = B.ID and B.��¼����=3 And B.NO = J.Column_Value And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����շѵ��Ƿ����ˢ����¼", strNoIns)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetSumMoney(Optional blnCls As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ú���ʾ�ϼ�
    '����:���˺�
    '����:2011-03-04 14:17:20
    '����:36285
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblSumMoney As Double
    With vsFee
        If blnCls = False Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("���"))) <> 0 And _
                  Val(.Cell(flexcpData, i, .ColIndex("���"))) = 0 Then
                    dblSumMoney = dblSumMoney + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                End If
            Next
        Else
            dblSumMoney = 0
        End If
    End With
    lblSum.Caption = "������˺ϼ�:" & Format(dblSumMoney, "###0.00;-###0.00;0.00;0.00")
End Sub

Public Sub StatusShowBillSum()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '����:���˺�
    '����:2011-03-11 18:09:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cur��� As Currency, dbl��Ʊ��� As Double, strNO As String, str��Ʊ�� As String
    Dim strTemp As String
    
    With vsFee
        strTemp = "": dbl��Ʊ��� = 0: cur��� = 0
        If Not (.Row > .Rows - 1 Or .Row < 1) Then
            strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
            str��Ʊ�� = .TextMatrix(.Row, .ColIndex("Ʊ�ݺ�"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ݺ�")) = strNO Then
                        cur��� = cur��� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                End If
                If .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = str��Ʊ�� Then
                        dbl��Ʊ��� = dbl��Ʊ��� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                End If
            Next
            strTemp = "����(" & strNO & ")�ϼ�:" & Format(cur���, "###0.00;-###0.00;0.00;0.00")
            strTemp = strTemp & "  ��Ʊ(" & str��Ʊ�� & ")�ϼ�:" & Format(dbl��Ʊ���, "###0.00;-###0.00;0.00;0.00")
        End If
        stbThis.Panels(2).Text = strTemp
    End With
End Sub
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
End Sub

Private Sub zlCreateObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������¼�����
    '����: �����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-28 16:16:00
    '˵��:
    '����:54896
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '������������
    Err = 0: On Error Resume Next
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
         Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    
End Sub
Private Sub zlCloseObject()
    '�ر���ض���
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub

