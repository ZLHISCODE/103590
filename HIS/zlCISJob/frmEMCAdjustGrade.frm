VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmEMCAdjustGrade 
   BackColor       =   &H00EFFEFE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���鼶�����"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5985
   Icon            =   "frmEMCAdjustGrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      BackColor       =   &H00EFFEFE&
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1560
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txtGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5FEFE&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   650
         Width           =   1455
      End
      Begin VSFlex8Ctl.VSFlexGrid vsflist 
         Height          =   1815
         Left            =   1560
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
         _cx             =   4471
         _cy             =   3201
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEMCAdjustGrade.frx":6852
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
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00EFFEFE&
         Caption         =   "�޶�˵��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   8
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblPatient 
         BackColor       =   &H00EFFEFE&
         Caption         =   "��ǰ���ﲡ�ˣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00EFFEFE&
         Caption         =   "�޶����鼶��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00EFFEFE&
         Caption         =   "��ǰ���鼶��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   700
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4680
      TabIndex        =   1
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4680
      TabIndex        =   0
      Top             =   3480
      Width           =   1100
   End
End
Attribute VB_Name = "frmEMCAdjustGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mstrPatient As String, mstrGrade As String, mlng�Һ�ID As Long

Public Function ShowMe(frmParent As Object, ByVal lng�Һ�id As Long, ByVal strPatient As String, ByVal strGrade As String)
    
    mstrPatient = strPatient
    mstrGrade = strGrade
    mlng�Һ�ID = lng�Һ�id
    
    Me.Show vbModal, frmParent
    
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, lngGrade As Long, strGrade As String
    Dim lngRemark As Long, strRemark As String
    
    On Error GoTo errH
    
    strRemark = Trim(txtRemark.Text)
    lngRemark = zlCommFun.ActualLen(strRemark)
    If lngRemark > 100 Then
        MsgBox "�޶�˵�����ֻ��������100���ַ�����ǰ������" & lngRemark & "���ַ��������µ�����", vbInformation
        Exit Sub
    End If
    
    strGrade = vsflist.TextMatrix(vsflist.Row, vsflist.ColIndex("��������"))
    If strGrade = mstrGrade Then
        MsgBox "�޶����鼶���뵱ǰ���鼶����ͬ��������ѡ��", vbInformation
        Exit Sub
    End If
    
    lngGrade = vsflist.TextMatrix(vsflist.Row, vsflist.ColIndex("���"))
    
    strSQL = "Zl_���ﲡ�鼶��_Edit(" & mlng�Һ�ID & "," & lngGrade & ",'" & strRemark & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "�޶����鼶��")
    mblnOk = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim lngRow As Long
    
    Call InitVSFList
    Call LoadList
    
    lngRow = vsflist.FindRow(mstrGrade, , vsflist.ColIndex("��������"), False, True)
    If lngRow > 0 Then vsflist.Row = lngRow
    
    lblPatient.Caption = lblPatient.Caption & mstrPatient
    txtGrade.Text = mstrGrade
        
End Sub

Private Sub LoadList()
'���ܣ���ϵͳ��ż������ݱ��嵥
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select ���, ���� as ��������, ���س̶�, ���߱�ʶ��ɫ From ���ﲡ�鼶��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ﲡ�鼶��")
    
    'Set vsflist.DataSource = rstmp
    '�󶨷�ʽ�ᵼ��Colkey�����Զ�ʧ
    With vsflist
        .Redraw = flexRDNone
        .BackColorFixed = &HC0DFE0
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
        i = .FixedRows
                
        While Not rsTmp.EOF
            .TextMatrix(i, .ColIndex("���")) = rsTmp!���
            .TextMatrix(i, .ColIndex("��������")) = rsTmp!��������
            .TextMatrix(i, .ColIndex("���س̶�")) = rsTmp!���س̶�
            '.TextMatrix(i, .ColIndex("���߱�ʶ��ɫ")) = rsTmp!���߱�ʶ��ɫ
            
            .Cell(flexcpFloodColor, i, 0, .Cols - 1) = "&H" & rsTmp!���߱�ʶ��ɫ
            i = i + 1
            rsTmp.MoveNext
        Wend
        .Redraw = flexRDDirect
    End With
    
    
    strSQL = "select �޶�˵�� from ��������¼ where �Һ�ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�޶�˵��", mlng�Һ�ID)
    If Not rsTmp.EOF Then
        txtRemark.Text = "" & rsTmp!�޶�˵��
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub InitVSFList()
    Dim strHead As String
    
    strHead = "���,0,4;��������,1400,1;���س̶�,550,1;���߱�ʶ��ɫ,0,1"
    Call zl9ComLib.Grid.Init(vsflist, strHead)
    
    With vsflist
        '.Editable = flexEDKbdMouse
        .ExtendLastCol = True
       
        .SelectionMode = flexSelectionByRow
        .AllowSelection = True
        .RowHeightMin = 280
        '.AllowUserResizing = flexResizeColumns
        '.ExplorerBar = flexExSortShow
    End With
End Sub

