VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicWorkTimeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ϰ�ʱ������"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicWorkTimeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4468.677
   ScaleMode       =   0  'User
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   378
      Left            =   4500
      TabIndex        =   31
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "�����˳�(&C)"
      Height          =   378
      Left            =   4265
      TabIndex        =   30
      Top             =   4110
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   378
      Left            =   390
      TabIndex        =   27
      Top             =   4110
      Width           =   1245
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   3915
      Left            =   30
      ScaleHeight     =   3915
      ScaleWidth      =   7350
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Width           =   7350
      Begin VB.TextBox txtԤ��ʱ�� 
         Height          =   330
         Left            =   3630
         TabIndex        =   16
         Text            =   "0"
         Top             =   1590
         Width           =   975
      End
      Begin MSComCtl2.UpDown updԤ��ʱ�� 
         Height          =   300
         Left            =   4590
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtԤ��ʱ��"
         BuddyDispid     =   196613
         OrigLeft        =   6300
         OrigTop         =   1560
         OrigRight       =   6555
         OrigBottom      =   1860
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   330
         Left            =   6360
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2055
         Width           =   885
      End
      Begin VB.ComboBox cboʱ��� 
         Height          =   330
         ItemData        =   "frmClinicWorkTimeEdit.frx":000C
         Left            =   930
         List            =   "frmClinicWorkTimeEdit.frx":000E
         TabIndex        =   5
         Top             =   585
         Width           =   2505
      End
      Begin VB.CommandButton cmdAddRestTime 
         Caption         =   "����(&A)"
         Height          =   330
         Left            =   5460
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2055
         Width           =   885
      End
      Begin MSComCtl2.DTPicker dtpRestEndTime 
         Height          =   330
         Left            =   2550
         TabIndex        =   21
         Top             =   2055
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin VB.Frame fraLineBetween2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   25
         Left            =   -60
         TabIndex        =   29
         Top             =   3840
         Width           =   8145
      End
      Begin VB.Frame fraLineBetween1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   25
         Left            =   -30
         TabIndex        =   6
         Top             =   1020
         Width           =   8025
      End
      Begin VB.ComboBox cbo���� 
         Height          =   330
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   150
         Width           =   2595
      End
      Begin MSComCtl2.DTPicker dtpDefaultTime 
         Height          =   330
         Left            =   6000
         TabIndex        =   12
         Top             =   1155
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker dtpPriorTime 
         Height          =   330
         Left            =   930
         TabIndex        =   14
         Top             =   1605
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf��Ϣʱ�� 
         Height          =   1335
         Left            =   930
         TabIndex        =   24
         Top             =   2400
         Width           =   6315
         _cx             =   11139
         _cy             =   2355
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.ComboBox cboNodeNo 
         Height          =   330
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   330
         Left            =   3630
         TabIndex        =   10
         Top             =   1155
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   330
         Left            =   930
         TabIndex        =   8
         Top             =   1155
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker dtpRestStartTime 
         Height          =   330
         Left            =   930
         TabIndex        =   19
         Top             =   2055
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin VB.Label lblԤ��ʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ��ʱ��(��)"
         Height          =   210
         Left            =   2370
         TabIndex        =   15
         Top             =   1650
         Width           =   1260
      End
      Begin VB.Label lblPriorTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰʱ��"
         Height          =   210
         Left            =   60
         TabIndex        =   13
         Top             =   1650
         Width           =   840
      End
      Begin VB.Label lblDefaultTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡʱ��"
         Height          =   210
         Left            =   5130
         TabIndex        =   11
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lblRestTimeAnd 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   210
         Left            =   2250
         TabIndex        =   20
         Top             =   2130
         Width           =   210
      End
      Begin VB.Label lblRestTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ϣʱ��"
         Height          =   210
         Left            =   60
         TabIndex        =   18
         Top             =   2100
         Width           =   840
      End
      Begin VB.Label lblEndTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ֹʱ��"
         Height          =   210
         Left            =   2730
         TabIndex        =   9
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   210
         Left            =   60
         TabIndex        =   7
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lblʱ��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ���"
         Height          =   210
         Left            =   270
         TabIndex        =   4
         Top             =   645
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   210
         Left            =   4185
         TabIndex        =   2
         Top             =   210
         Width           =   420
      End
      Begin VB.Label lblNodeNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "վ��"
         Height          =   210
         Left            =   480
         TabIndex        =   0
         Top             =   210
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   378
      Left            =   5790
      TabIndex        =   26
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdSaveAdd 
      Caption         =   "��������(&O)"
      Height          =   378
      Left            =   2730
      TabIndex        =   25
      Top             =   4110
      Width           =   1335
   End
End
Attribute VB_Name = "frmClinicWorkTimeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const M_CurDate As String = "2016/01/01 "
Private mbytFun As G_Enum_Fun '0-�鿴,1-���,2-����
Private mstrվ�� As String
Private mstr���� As String
Private mstrʱ��� As String

Private Enum mGridHeadCol
    COL_��� = 0
    COL_��ʼʱ��
    COL_����ʱ��
End Enum
Private mrsʱ��� As ADODB.Recordset
Private mblnOK As Boolean

Public Function ShowMe(frmParent As Form, ByVal bytFun As G_Enum_Fun, _
    Optional ByVal strվ�� As String, Optional ByVal str���� As String, _
    Optional ByVal strʱ��� As String) As Boolean
    '�������
    '��Σ�
    '   frmParent - ������
    '   bytFun - ��������, 0-�鿴��1-������2-�޸�
    mbytFun = bytFun
    mstrվ�� = strվ��: mstr���� = str����: mstrʱ��� = strʱ���
    
    Err = 0: On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cboNodeNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboʱ���_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "-" Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboʱ���_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cboʱ���.Text) > 6 Then
        MsgBox "ʱ�������ֻ��������6���ַ���3�����֣�", vbInformation, gstrSysName
        zlControl.TxtSelAll cboʱ���
        Cancel = True
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdSaveExit_Click()
    On Error GoTo ErrHandler
    If IsValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    mblnOK = True
    
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSaveAdd_Click()
    On Error GoTo ErrHandler
    If IsValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    mblnOK = True
    
    '��������
    Call ClearFaceInfor
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDefaultTime_Change()
    dtpDefaultTime.Tag = Format(dtpDefaultTime.Value, "hh:mm:ss")
End Sub

Private Sub dtpDefaultTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpPriorTime_Change()
    dtpPriorTime.Tag = Format(dtpPriorTime.Value, "hh:mm:ss")
End Sub

Private Sub dtpPriorTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpRestEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpRestStartTime_Change()
    dtpRestEndTime.Value = dtpRestStartTime.Value
End Sub

Private Sub dtpRestStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Long, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
     Me.Caption = Choose(mbytFun + 1, "�鿴", "����", "�޸�", "ɾ��") & "�ϰ�ʱ��"
    If InitGridHead() = False Then Unload Me: Exit Sub
    
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If InitData() = False Then Unload Me: Exit Sub
    End If
    If mbytFun = Fun_Add Then
        cmdSaveAdd.Visible = True
        cmdSaveExit.Visible = True
        cmdOK.Visible = False
        Exit Sub
    Else
        cmdSaveAdd.Visible = False
        cmdSaveExit.Visible = False
        cmdOK.Visible = True
    End If
    
    If mbytFun = Fun_View Then '������༭�޸�
        cmdCancel.Visible = False
        cmdOK.Left = cmdCancel.Left
        Call SetEnabled(Me.Controls, False)
    Else
        MsgBox "���ѣ�" & vbCrLf & _
               "    �벻Ҫ�����޸��ϰ�ʱ��Σ�һ���޸���Ҫ��ʱ������ʹ���˵�ǰ�ϰ�ʱ����������˷�ʱ�εİ��Ž������»���ʱ�Σ����򣬿��ܻᵼ��ԤԼ�Һų���", vbInformation, gstrSysName
    End If
    
    '��������
    If LoadData(mstrվ��, mstr����, mstrʱ���) = False Then Unload Me: Exit Sub
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Unload Me
End Sub

Private Function InitData() As Boolean
    Dim i As Long, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    '����վ������
    strSQL = "Select ���, ���� From Zlnodelist"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboNodeNo.Clear
    cboNodeNo.AddItem ""
    Do While Not rsTemp.EOF
        cboNodeNo.AddItem Nvl(rsTemp!���) & "-" & Nvl(rsTemp!����)
        rsTemp.MoveNext
    Loop
    
    '���غ�������
    strSQL = "Select ����, ����, ����, Nvl(ȱʡ��־, 0) As ȱʡ��־ From ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo����.Clear
    cbo����.AddItem ""
    Do While Not rsTemp.EOF
        cbo����.AddItem Nvl(rsTemp!����)
        'If Nvl(rsTemp!ȱʡ��־) = 1 Then cbo����.ListIndex = cbo����.NewIndex
        rsTemp.MoveNext
    Loop
    
    '���������ϰ�ʱ��Σ��Ա�ѡ��
    strSQL = "Select Distinct ʱ��� From ʱ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboʱ���.Clear
    Do While Not rsTemp.EOF
        cboʱ���.AddItem Nvl(rsTemp!ʱ���)
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadData(ByVal strվ�� As String, ByVal str���� As String, _
    ByVal strʱ��� As String, Optional ByVal blnDefault As Boolean) As Boolean
    '����ʱ�������
    '��Σ�blnDefault True-ѡ��ʱ���ȱʡ��������
    Dim i As Long
    Dim strSQL As String, strWhere As String, rs�ϰ�ʱ�� As ADODB.Recordset
    Dim varTims As Variant, varRow As Variant
    
    Err = 0: On Error GoTo ErrHandler
    If Not blnDefault Then
        strWhere = " And Nvl(վ��, '-') = Nvl([1], '-') And Nvl(����, '-') = Nvl([2], '-')"
    End If
    strWhere = strWhere & " And a.ʱ���=[3]"
    
    strSQL = "Select a.ʱ���, a.����, a.��ʼʱ��, a.��ֹʱ��, a.��Ϣʱ��," & vbNewLine & _
            "        a.ȱʡʱ��, a.��ǰʱ��, a.����Ԥ��ʱ��, " & vbNewLine & _
            "        b.���, b.���� As վ��" & vbNewLine & _
            " From ʱ��� A, Zlnodelist B" & vbNewLine & _
            " Where a.վ�� = b.���(+)" & strWhere & vbNewLine & _
            " Order By Nvl(b.���, -1), Nvl(a.����, -1)"
    Set rs�ϰ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strվ��, str����, strʱ���)
    If rs�ϰ�ʱ��.EOF Then Exit Function
    
    If Not blnDefault Then
        zlControl.CboSetText cboNodeNo, Nvl(rs�ϰ�ʱ��!վ��)
        If cboNodeNo.ListIndex = -1 Then cboNodeNo.AddItem Nvl(rs�ϰ�ʱ��!վ��): cboNodeNo.ListIndex = cboNodeNo.NewIndex
        zlControl.CboSetText cbo����, Nvl(rs�ϰ�ʱ��!����)
        If cbo����.ListIndex = -1 Then cbo����.AddItem Nvl(rs�ϰ�ʱ��!����): cbo����.ListIndex = cbo����.NewIndex
        zlControl.CboSetText cboʱ���, Nvl(rs�ϰ�ʱ��!ʱ���)
        If cboʱ���.ListIndex = -1 Then cboʱ���.AddItem Nvl(rs�ϰ�ʱ��!ʱ���): cboʱ���.ListIndex = cboʱ���.NewIndex
    End If
    
    dtpStartTime.Value = Nvl(rs�ϰ�ʱ��!��ʼʱ��)
    dtpEndTime.Value = Nvl(rs�ϰ�ʱ��!��ֹʱ��)
    dtpDefaultTime.Value = Nvl(rs�ϰ�ʱ��!ȱʡʱ��, Nvl(rs�ϰ�ʱ��!��ʼʱ��))
    dtpDefaultTime.Tag = dtpDefaultTime.Value
    dtpPriorTime.Value = Nvl(rs�ϰ�ʱ��!��ǰʱ��, Nvl(rs�ϰ�ʱ��!��ʼʱ��))
    dtpPriorTime.Tag = dtpPriorTime.Value
    txtԤ��ʱ��.Text = Val(Nvl(rs�ϰ�ʱ��!����Ԥ��ʱ��, 0))
    
    vsf��Ϣʱ��.Clear 1
    vsf��Ϣʱ��.Rows = 1
    If Nvl(rs�ϰ�ʱ��!��Ϣʱ��) <> "" Then
        varTims = Split(Nvl(rs�ϰ�ʱ��!��Ϣʱ��), ";")
        For i = 0 To UBound(varTims)
            If varTims(i) <> "" Then
                varRow = Split(varTims(i), "-")
                vsf��Ϣʱ��.AddItem CStr(i + 1) & vbTab & Format(varRow(0), "hh:mm:ss") & vbTab & Format(varRow(1), "hh:mm:ss")
            End If
        Next
        vsf��Ϣʱ��.RowHeight(-1) = vsf��Ϣʱ��.RowHeight(0)
    End If
    LoadData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cboʱ���_Click()
    Dim varStr As Variant, strTims As Variant
    Dim i As Long, Row As Long
    
    Err = 0: On Error GoTo ErrHandler
    Call LoadData("", "", cboʱ���.Text, True)
    dtpRestStartTime.Value = dtpStartTime.Value
    dtpRestEndTime.Value = dtpStartTime.Value
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then Unload Me: Exit Sub
    
    cmdOK.Enabled = False
    If IsValied() = False Then cmdOK.Enabled = True: Exit Sub
    If SaveData() = False Then cmdOK.Enabled = True: Exit Sub
    mblnOK = True
    
    '��������
    If mbytFun = Fun_Add Then
        cmdOK.Enabled = True
        Exit Sub
    Else
        MsgBox "ע�⣺" & vbCrLf & _
               "    �ϰ�ʱ����޸ĳɹ����뼰ʱ������ʹ���˵�ǰ�ϰ�ʱ����������˷�ʱ�εİ��Ž������»���ʱ�Σ����򣬿��ܻᵼ��ԤԼ�Һų���", vbExclamation, gstrSysName
    End If
    
    Unload Me
    Exit Sub
ErrHandler:
    cmdOK.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ClearFaceInfor()
    '����:���������Ϣ���Ա�������������
    On Error GoTo errHandle
    cboNodeNo.ListIndex = -1
    cbo����.ListIndex = -1
    cboʱ���.Text = "": cboʱ���.ListIndex = -1
    
    dtpStartTime.Value = "00:00:00": dtpEndTime.Value = "00:00:00"
    dtpDefaultTime.Value = "00:00:00": dtpDefaultTime.Tag = ""
    dtpPriorTime.Value = "00:00:00": dtpPriorTime.Tag = ""
    txtԤ��ʱ��.Text = 0
    
    vsf��Ϣʱ��.Clear 1
    vsf��Ϣʱ��.Rows = 1
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim str��Ϣʱ�� As String, i  As Integer
    Dim strSQL As String, strColor As String
    Dim dtStartTime As Date, dtEndTime As Date
    Dim dtRestStartTime As Date, dtRestEndTime As Date
    Dim dtPriorTime As Date, dtDefaultTime As Date
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun <> Fun_Delete Then
        For i = 1 To vsf��Ϣʱ��.Rows - 1
            str��Ϣʱ�� = str��Ϣʱ�� & ";" & vsf��Ϣʱ��.TextMatrix(i, COL_��ʼʱ��) & "-" & vsf��Ϣʱ��.TextMatrix(i, COL_����ʱ��)
        Next
        If str��Ϣʱ�� <> "" Then str��Ϣʱ�� = Mid(str��Ϣʱ��, 2)
        
        Call FormatTime(0, dtStartTime, dtEndTime, dtRestStartTime, dtRestEndTime, dtPriorTime, dtDefaultTime)
    End If
    
    'CREATE OR REPLACE Procedure Zl_�ϰ�ʱ��_Modify
    '(
    '  ��������_In     Number,
    '  վ��_In         ʱ���.վ��%Type,
    '  ����_In         ʱ���.����%Type,
    '  ʱ���_In       ʱ���.ʱ���%Type,
    '  ��ʼʱ��_In     ʱ���.��ʼʱ��%Type := Null,
    '  ��ֹʱ��_In     ʱ���.��ֹʱ��%Type := Null,
    '  ��Ϣʱ��_In     ʱ���.��Ϣʱ��%Type := Null,
    '  ȱʡʱ��_In     ʱ���.ȱʡʱ��%Type := Null,
    '  ��ǰʱ��_In     ʱ���.��ǰʱ��%Type := Null,
    '  ����Ԥ��ʱ��_In ʱ���.����Ԥ��ʱ��%Type := 0,
    '  ԭվ��_In       ʱ���.վ��%Type := Null,
    '  ԭ����_In       ʱ���.����%Type := Null,
    '  ԭʱ���_In     ʱ���.ʱ���%Type := Null
    ') As
    '  --��������_In 0-������1-�޸ģ�2-ɾ��
    Select Case mbytFun
    Case Fun_Add
        strSQL = "Zl_�ϰ�ʱ��_Modify("
        strSQL = strSQL & "" & 0 & ","
        strSQL = strSQL & "'" & NeedCode(cboNodeNo.Text) & "',"
        strSQL = strSQL & "'" & cbo����.Text & "',"
        strSQL = strSQL & "'" & cboʱ���.Text & "',"
        strSQL = strSQL & "To_Date('" & Format(dtStartTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "To_Date('" & Format(dtEndTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "'" & str��Ϣʱ�� & "',"
        strSQL = strSQL & "To_Date('" & Format(dtDefaultTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "To_Date('" & Format(dtPriorTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "" & Val(txtԤ��ʱ��.Text) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Case Fun_Update
        strSQL = "Zl_�ϰ�ʱ��_Modify("
        strSQL = strSQL & "" & 1 & ","
        strSQL = strSQL & "'" & NeedCode(cboNodeNo.Text) & "',"
        strSQL = strSQL & "'" & cbo����.Text & "',"
        strSQL = strSQL & "'" & cboʱ���.Text & "',"
        strSQL = strSQL & "To_Date('" & Format(dtStartTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "To_Date('" & Format(dtEndTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "'" & str��Ϣʱ�� & "',"
        strSQL = strSQL & "To_Date('" & Format(dtDefaultTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "To_Date('" & Format(dtPriorTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "" & Val(txtԤ��ʱ��.Text) & ","
        strSQL = strSQL & "'" & mstrվ�� & "',"
        strSQL = strSQL & "'" & mstr���� & "',"
        strSQL = strSQL & "'" & mstrʱ��� & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End Select
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InitGridHead() As Boolean
    Dim strHead As String
    Dim i As Long, varData As Variant

    Err = 0: On Error GoTo ErrHandler
    strHead = "���,4,500|��ʼʱ��,4,1300|����ʱ��,4,1300"
    With vsf��Ϣʱ��
        .Redraw = False
        .FixedCols = 1: .FixedRows = 1
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .RowHeight(-1) = 280
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = True
    End With
    InitGridHead = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdAddRestTime_Click()
    Dim i As Long
    Dim dtRestStartTime As Date, dtRestEndTime As Date
    Dim dtStartTime As Date, dtEndTime As Date
    Dim dtTempStart As Date, dtTempEnd As Date
    
    Err = 0: On Error GoTo ErrHandler
    Call FormatTime(1, dtStartTime, dtEndTime, dtRestStartTime, dtRestEndTime)
    If dtRestStartTime >= dtRestEndTime Then
        MsgBox "��Ϣʱ��Ľ���ʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
        If dtpRestEndTime.Visible And dtpRestEndTime.Enabled Then dtpRestEndTime.SetFocus
        Exit Sub
    End If
    If Not ((dtRestStartTime >= dtStartTime And dtRestStartTime <= dtEndTime) _
            And (dtRestEndTime >= dtStartTime And dtRestEndTime <= dtEndTime)) Then
        MsgBox "��Ϣʱ��������ϰ�ʱ��(" & Format(dtStartTime, "hh:mm:ss") & "-" & Format(dtEndTime, "hh:mm:ss") & ")��Χ�ڣ�", vbInformation, gstrSysName
        If dtpRestStartTime.Visible And dtpRestStartTime.Enabled Then dtpRestStartTime.SetFocus
        Exit Sub
    End If
    
    For i = 1 To vsf��Ϣʱ��.Rows - 1
        dtTempStart = M_CurDate & vsf��Ϣʱ��.TextMatrix(i, COL_��ʼʱ��)
        dtTempEnd = M_CurDate & vsf��Ϣʱ��.TextMatrix(i, COL_����ʱ��)
        If dtTempEnd <= dtTempStart Then dtTempEnd = DateAdd("d", 1, dtTempEnd)
        
        If Not ((dtRestStartTime < dtTempStart And dtRestEndTime < dtTempStart) _
                Or (dtRestStartTime > dtTempEnd And dtRestEndTime > dtTempEnd)) Then
            MsgBox "��Ϣʱ�䲻�ܰ�������������Ϣʱ�䷶Χ�ڣ�", vbInformation, gstrSysName
            If dtpRestStartTime.Visible And dtpRestStartTime.Enabled Then dtpRestStartTime.SetFocus
            Exit Sub
        End If
    Next
    vsf��Ϣʱ��.AddItem CStr(vsf��Ϣʱ��.Rows) & vbTab & Format(dtRestStartTime, "hh:mm:ss") & vbTab & Format(dtRestEndTime, "hh:mm:ss")
    vsf��Ϣʱ��.RowHeight(-1) = vsf��Ϣʱ��.RowHeight(0)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub FormatTime(ByVal bytType As Byte, ByRef dtStartTime As Date, ByRef dtEndTime As Date, _
    Optional ByRef dtRestStartTime As Date, Optional ByRef dtRestEndTime As Date, _
    Optional ByRef dtPriorTime As Date, Optional ByRef dtDefaultTime As Date)
    Dim blnChanged As Boolean
    
    '��ʽ��ʱ��
    dtStartTime = M_CurDate & Format(dtpStartTime.Value, "hh:mm:ss")
    dtEndTime = M_CurDate & Format(dtpEndTime.Value, "hh:mm:ss")
    dtRestStartTime = M_CurDate & Format(dtpRestStartTime.Value, "hh:mm:ss")
    dtRestEndTime = M_CurDate & Format(dtpRestEndTime.Value, "hh:mm:ss")
    
    dtPriorTime = M_CurDate & Format(dtpPriorTime.Value, "hh:mm:ss")
    dtDefaultTime = M_CurDate & Format(dtpDefaultTime.Value, "hh:mm:ss")
    
    If bytType = 1 Then
        blnChanged = False
        If dtEndTime <= dtStartTime Then
            blnChanged = True
            dtEndTime = DateAdd("d", 1, dtEndTime) '��ʼʱ����ڽ���ʱ�䣬�����ʱ���һ��
        End If
        If dtRestEndTime <= dtRestStartTime And blnChanged Then dtRestEndTime = DateAdd("d", 1, dtRestEndTime) '��Ϣ��ʼʱ�������Ϣ����ʱ�䣬����Ϣ����ʱ���һ��
        If dtRestStartTime < dtStartTime Then dtRestStartTime = DateAdd("d", 1, dtRestStartTime) '��ʼʱ�������Ϣ��ʼʱ�䣬����Ϣ��ʼʱ���һ��
        If dtRestEndTime < dtStartTime Then dtRestEndTime = DateAdd("d", 1, dtRestEndTime) '��ʼʱ�������Ϣ����ʱ�䣬����Ϣ����ʱ���һ��
        If dtDefaultTime < dtStartTime Then dtDefaultTime = DateAdd("d", 1, dtDefaultTime) '��ʼʱ�����ȱʡԤԼʱ�䣬��ȱʡԤԼʱ���һ��
        If dtPriorTime > dtStartTime Then dtPriorTime = DateAdd("d", -1, dtPriorTime) '��ʼʱ��С����ǰ�Һ�ʱ�䣬����ǰ�Һ�ʱ���һ��
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsʱ��� Is Nothing Then Set mrsʱ��� = Nothing
End Sub

Private Sub txtԤ��ʱ��_GotFocus()
    zlControl.TxtSelAll txtԤ��ʱ��
End Sub

Private Sub txtԤ��ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub vsf��Ϣʱ��_EnterCell()
    cmdDelete.Enabled = vsf��Ϣʱ��.Row > 0
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If vsf��Ϣʱ��.Row > 0 Then
        If MsgBox("��ȷ��Ҫɾ���� " & vsf��Ϣʱ��.Row & " �У�", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            vsf��Ϣʱ��.RemoveItem vsf��Ϣʱ��.Row
            For i = 1 To vsf��Ϣʱ��.Rows - 1 '���±��
                vsf��Ϣʱ��.TextMatrix(i, 0) = i
            Next
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function IsValied() As Boolean
    Dim dtRestStartTime As Date, dtRestEndTime As Date
    Dim dtStartTime As Date, dtEndTime As Date
    Dim dtPriorTime As Date, dtDefaultTime As Date
    Dim dtTempStart As Date, dtTempEnd As Date
    Dim i As Integer, lngMinute As Long
    
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    If cboʱ���.Text = "" Then
        MsgBox "ʱ��β���Ϊ�գ�", vbInformation, gstrSysName
        If cboʱ���.Visible And cboʱ���.Enabled Then cboʱ���.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(cboʱ���.Text) > 6 Then
        MsgBox "ʱ�������ֻ��������6���ַ���3�����֣�", vbInformation, gstrSysName
        If cboʱ���.Visible And cboʱ���.Enabled Then cboʱ���.SetFocus
        zlControl.TxtSelAll cboʱ���
        Exit Function
    End If
    If IsNumeric(Val(txtԤ��ʱ��.Text)) = False Then
        MsgBox "Ԥ��ʱ��ֻ��Ϊ���֣�", vbInformation, gstrSysName
        If txtԤ��ʱ��.Visible And txtԤ��ʱ��.Enabled Then txtԤ��ʱ��.SetFocus
        zlControl.TxtSelAll txtԤ��ʱ��
        Exit Function
    End If
    If mbytFun = Fun_Add Then
        If CheckExist(NeedCode(cboNodeNo.Text), cbo����.Text, cboʱ���.Text) Then
            MsgBox NeedName(cboNodeNo.Text) & "�Ѵ���" & IIf(cbo����.Text = "", "���ֺ���", "����Ϊ��" & cbo����.Text & "��") & "�ġ�" & cboʱ���.Text & "��ʱ��Σ�", vbInformation, gstrSysName
            If cboʱ���.Visible And cboʱ���.Enabled Then cboʱ���.SetFocus
            zlControl.TxtSelAll cboʱ���
            Exit Function
        End If
    ElseIf mbytFun = Fun_Update Then
        If mstrվ�� <> NeedCode(cboNodeNo.Text) Or mstr���� <> cbo����.Text Or mstrʱ��� <> cboʱ���.Text Then
            If CheckHaveUsed(mstrվ��, mstr����, mstrʱ���) Then
                MsgBox "��ǰ�ϰ�ʱ����ѱ�ʹ�ã������޸�վ�㡢���༰ʱ������ƣ�", vbInformation, gstrSysName
                If cboʱ���.Visible And cboʱ���.Enabled Then cboʱ���.SetFocus
                zlControl.TxtSelAll cboʱ���
                Exit Function
            End If
            If CheckExist(NeedCode(cboNodeNo.Text), cbo����.Text, cboʱ���.Text) Then
                MsgBox NeedName(cboNodeNo.Text) & "�Ѵ���" & IIf(cbo����.Text = "", "���ֺ���", "����Ϊ��" & cbo����.Text & "��") & "�ġ�" & cboʱ���.Text & "��ʱ��Σ�", vbInformation, gstrSysName
                If cboʱ���.Visible And cboʱ���.Enabled Then cboʱ���.SetFocus
                zlControl.TxtSelAll cboʱ���
                Exit Function
            End If
        End If
    End If
    
    Call FormatTime(1, dtStartTime, dtEndTime, dtRestStartTime, dtRestEndTime, dtPriorTime, dtDefaultTime)
    If dtPriorTime > dtStartTime Then
        MsgBox "��ǰ�Һ�ʱ�����С�ڵ��ڿ�ʼʱ�䣡", vbInformation, gstrSysName
        If dtpPriorTime.Visible And dtpPriorTime.Enabled Then dtpPriorTime.SetFocus
        Exit Function
    End If
    
    If dtDefaultTime < dtStartTime Or dtDefaultTime > dtEndTime Then
        MsgBox "ȱʡԤԼʱ��������ϰ�ʱ��(" & Format(dtStartTime, "hh:mm:ss") & "-" & Format(dtEndTime, "hh:mm:ss") & ")��Χ�ڣ�", vbInformation, gstrSysName
        If dtpDefaultTime.Visible And dtpDefaultTime.Enabled Then dtpDefaultTime.SetFocus
        Exit Function
    End If
    lngMinute = DateDiff("n", dtStartTime, dtEndTime)
    
    For i = 1 To vsf��Ϣʱ��.Rows - 1
        dtTempStart = M_CurDate & vsf��Ϣʱ��.TextMatrix(i, COL_��ʼʱ��)
        dtTempEnd = M_CurDate & vsf��Ϣʱ��.TextMatrix(i, COL_����ʱ��)
        If dtTempEnd <= dtTempStart Then dtTempEnd = DateAdd("d", 1, dtTempEnd) '��Ϣ��ʼʱ�������Ϣ����ʱ�䣬����Ϣ����ʱ���һ��
        If dtTempStart < dtStartTime Then dtTempStart = DateAdd("d", 1, dtTempStart) '��ʼʱ�������Ϣ��ʼʱ�䣬����Ϣ��ʼʱ���һ��
        If dtTempEnd < dtStartTime Then dtTempEnd = DateAdd("d", 1, dtTempEnd) '��ʼʱ�������Ϣ����ʱ�䣬����Ϣ����ʱ���һ��

        If Not ((dtTempStart >= dtStartTime And dtTempStart <= dtEndTime) _
            And (dtTempEnd >= dtStartTime And dtTempEnd <= dtEndTime)) Then
            MsgBox "��" & i & "����Ϣʱ�䲻���ϰ�ʱ��(" & Format(dtStartTime, "hh:mm:ss") & "-" & Format(dtEndTime, "hh:mm:ss") & ")��Χ�ڣ�", vbInformation, gstrSysName
            vsf��Ϣʱ��.Row = i
            Exit Function
        End If
        lngMinute = lngMinute - DateDiff("n", dtTempStart, dtTempEnd)
    Next
    
    'Ԥ��ʱ�䲻�ܴ����ܵķ�����
    If Val(txtԤ��ʱ��.Text) > lngMinute Then
        MsgBox "Ԥ��ʱ�䲻�ܴ����ϰ�ʱ�ε���ʱ�䣡", vbInformation, gstrSysName
        If txtԤ��ʱ��.Visible And txtԤ��ʱ��.Enabled Then txtԤ��ʱ��.SetFocus
        zlControl.TxtSelAll txtԤ��ʱ��
        Exit Function
    End If
    
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckExist(ByVal strվ�� As String, ByVal str���� As String, ByVal strʱ��� As String) As Boolean
    '����¼�Ƿ��Ѵ���
    Dim strSQL As String, rs�ϰ�ʱ�� As ADODB.Recordset
    Dim varTims As Variant, varRow As Variant
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select 1 From ʱ��� A, Zlnodelist B" & vbNewLine & _
            " Where a.վ�� = b.���(+)" & vbNewLine & _
            "       And Nvl(վ��, '-') = Nvl([1], '-') And Nvl(����, '-') = Nvl([2], '-') And ʱ��� = [3]"
    Set rs�ϰ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strվ��, str����, strʱ���)
    CheckExist = Not rs�ϰ�ʱ��.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckHaveUsed(ByVal strվ�� As String, ByVal str���� As String, ByVal strʱ��� As String) As Boolean
    '��鵱ǰ�ϰ�ʱ����Ƿ��ѱ�ʹ��
    Dim strSQL As String, rs�ϰ�ʱ�� As ADODB.Recordset
    Dim varTims As Variant, varRow As Variant
    
    Err = 0: On Error GoTo ErrHandler
    '���ԭ�ϰ�ʱ���Ƿ�ʹ�ã���ʹ�õĲ����޸�վ�㡢���ࡢʱ���
    '����ɾ����ʹ�õķ�Χ������һ��,��ʹ�õ�ʱ��ֻҪ��һ�����ɣ���ͬվ�㣬��ͬ������ܻ��ж��ͬ����ʱ��Σ�
    '�ٴ������Դ����
    strSQL = "Select 1" & vbNewLine & _
            " From (Select b.�ϰ�ʱ��, c.վ��, a.����," & vbNewLine & _
            "              Row_Number() Over(Partition By b.�ϰ�ʱ�� Order By b.�ϰ�ʱ��, c.վ�� Desc, a.���� Desc) As ���" & vbNewLine & _
            "        From �ٴ������Դ A, �ٴ������Դ���� B, ���ű� C" & vbNewLine & _
            "        Where a.Id = b.��Դid And a.����id = c.Id)" & vbNewLine & _
            " Where ��� = 1 And Nvl(վ��, '-') = Nvl([1], '-') And Nvl(����, '-') = Nvl([2], '-') And �ϰ�ʱ�� = [3] And Rownum < 2"
    Set rs�ϰ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strվ��, str����, strʱ���)
    If Not rs�ϰ�ʱ�� Is Nothing Then
        If Not rs�ϰ�ʱ��.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '�ٴ���������(�̶�����ģ��)
    strSQL = "Select 1" & vbNewLine & _
            " From (Select a.�ϰ�ʱ��, c.վ��, b.����," & vbNewLine & _
            "              Row_Number() Over(Partition By a.�ϰ�ʱ�� Order By a.�ϰ�ʱ��, c.վ�� Desc, b.���� Desc) As ���" & vbNewLine & _
            "        From �ٴ��������� A, �ٴ����ﰲ�� D, �ٴ������Դ B, ���ű� C" & vbNewLine & _
            "        Where a.����id = d.Id And d.��Դid = b.Id And b.����id = c.Id)" & vbNewLine & _
            " Where ��� = 1 And Nvl(վ��, '-') = Nvl([1], '-') And Nvl(����, '-') = Nvl([2], '-') And �ϰ�ʱ�� = [3] And Rownum < 2"
    Set rs�ϰ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strվ��, str����, strʱ���)
    If Not rs�ϰ�ʱ�� Is Nothing Then
        If Not rs�ϰ�ʱ��.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '�ٴ������¼
    '����飬��Ϊ�ñ�̫������ϰ�ʱ�ε���Ϣ����������������У�û���ҵ��ϰ�ʱ��ʱ������������������ȡ
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsf��Ϣʱ��_GotFocus()
    If vsf��Ϣʱ��.Rows > 1 Then
        vsf��Ϣʱ��.Row = 1
    Else
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vsf��Ϣʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
