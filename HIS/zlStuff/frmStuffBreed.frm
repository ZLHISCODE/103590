VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStuffBreed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������Ʒ�ֱ༭"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmStuffBreed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "������������(&B)"
      Height          =   350
      Left            =   3360
      TabIndex        =   29
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "���������Ʒ��(&A)"
      Height          =   350
      Left            =   1560
      TabIndex        =   28
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   7440
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "����"
      ToolTipText     =   "��*��ѡ����"
      Top             =   728
      Width           =   285
   End
   Begin VB.ComboBox cbo�����Ա� 
      Height          =   300
      Left            =   5475
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2280
      Width           =   2220
   End
   Begin VB.ComboBox cmbStationNo 
      Height          =   300
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid vsEditBill 
      Height          =   1410
      Left            =   1515
      TabIndex        =   19
      Top             =   2715
      Width           =   6165
      _cx             =   10874
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmStuffBreed.frx":030A
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
   Begin VB.ComboBox cbo��λ 
      Height          =   300
      Left            =   1515
      TabIndex        =   11
      Top             =   2295
      Width           =   2205
   End
   Begin VB.Frame fra 
      Height          =   60
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   8115
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6555
      TabIndex        =   24
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   255
      Picture         =   "frmStuffBreed.frx":03A0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�����˳�(&O)"
      Height          =   350
      Left            =   5160
      TabIndex        =   23
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtӢ�� 
      Height          =   300
      Left            =   5520
      MaxLength       =   40
      TabIndex        =   13
      Top             =   1125
      Width           =   2175
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Left            =   4800
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1920
      Width           =   2340
   End
   Begin VB.TextBox txtƴ�� 
      Height          =   300
      Left            =   1515
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1920
      Width           =   2160
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1515
      Width           =   6175
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1515
      MaxLength       =   13
      TabIndex        =   4
      Top             =   1125
      Width           =   2175
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -15
      TabIndex        =   16
      Top             =   4620
      Width           =   8490
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   450
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   6375
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4395
      Top             =   6375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffBreed.frx":04EA
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffBreed.frx":0A84
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffBreed.frx":101E
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffBreed.frx":15B8
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1515
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   2
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label lbl�����Ա� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ա�(&X)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4440
      TabIndex        =   26
      Top             =   2355
      Width           =   990
   End
   Begin VB.Label lblStationNo 
      AutoSize        =   -1  'True
      Caption         =   "Ժ��"
      Height          =   180
      Left            =   795
      TabIndex        =   20
      Top             =   4275
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&Q)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   22
      Top             =   2745
      Width           =   990
   End
   Begin VB.Label Lbl��λ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ɢװ��λ(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   10
      Top             =   2355
      Width           =   990
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "ע����Ʒ�ֽ�����2003-09-01"
      Height          =   180
      Left            =   3915
      TabIndex        =   14
      Top             =   4305
      Width           =   2580
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�����������ϵ����Ʒ��."
      Height          =   180
      Left            =   825
      TabIndex        =   0
      Top             =   240
      Width           =   2070
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   165
      Picture         =   "frmStuffBreed.frx":1B52
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblӢ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ������(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4440
      TabIndex        =   12
      Top             =   1185
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���Ϸ���(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   1
      Top             =   825
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���Ƽ���(&S)                         (ƴ��)                                (���)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   1980
      Width           =   7200
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ͨ������(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   5
      Top             =   1575
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   795
      TabIndex        =   3
      Top             =   1185
      Width           =   630
   End
End
Attribute VB_Name = "frmStuffBreed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr����ID As String         '��ǰ�༭�Ĳ���ID
Dim mlng����id As Long

Dim mintSuccess As Integer
Dim mintEditType As gEditType    '�༭����
Dim mblnChange As Boolean
Dim mstrPrivs As String         'Ȩ�޴�
Dim mblnFrist As Boolean        '��һ������ϵͳʱ
Dim mintCount As Integer
Dim mintCodeLength As Integer   '����ĳ���,�����ݿ��ж�ȡ�����ĳ���
Private mlngƷ��id As Long      '��¼Ʒ��id
Private Const mlngModule = 1711


Private Sub GetDefineSize()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    '����:���˺�
    '����:2007/05/24
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSQL = "Select ����,���� From ������ĿĿ¼ Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    mintCodeLength = rsTmp.Fields("����").DefinedSize
    txt����.MaxLength = rsTmp.Fields("����").DefinedSize
    txt����.MaxLength = rsTmp.Fields("����").DefinedSize
    
    gstrSQL = "Select ����,���� From ������Ŀ���� Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    txtƴ��.MaxLength = rsTmp.Fields("����").DefinedSize
    txt���.MaxLength = txtƴ��.MaxLength
    txtӢ��.MaxLength = rsTmp.Fields("����").DefinedSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowEditCard(ByVal frmMain As Object, _
    intEditType As gEditType, Optional ByVal str����id As String = "", Optional ByVal lng����id As Long, Optional strPrivs As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�༭��������
    '--�����:frmMain-���õ�������
    '--       intEditType -�༭����
    '--       str����ID-�༭�����ĵ�ǰ����ID
    '         strPrivs-Ȩ�޴�
    '--������:
    '--��  ��:�༭�ɹ�,����ture,����false
    '����:���˺�
    '����;2007/05/24
    '-----------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim intTemp As Byte
    Dim strTemp As String
    
    mlng����id = lng����id
    mstr����ID = str����id
    mstrPrivs = strPrivs
    mintEditType = intEditType
    mintSuccess = 0
    
    frmStuffBreed.Show 1, frmMain
    ShowEditCard = mintSuccess > 0
End Function

Private Sub cbo��λ_Change()
        mblnChange = True
End Sub

Private Sub cbo��λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub cmdSaveAddItem_Click()
    Call cmdOK_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    Call cmdOK_Click
End Sub

Private Sub cmd����_Click()
    If Me.tvwClass.Nodes.Count = 0 Then
        Call Load���Ʒ�����Ϣ
    End If
   With Me.tvwClass
        .Left = Me.txt����.Left
        .Top = Me.txt����.Top + Me.txt����.Height
        .Width = txt����.Width
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub
Private Sub cbo��λ_LostFocus()
    Dim strTmp As String
    Dim i As Long
    Dim blnAdd As Boolean
    ImeLanguage False
    
    strTmp = cbo��λ.Text
    blnAdd = True
    For i = 0 To cbo��λ.ListCount - 1
        If cbo��λ.List(i) = Trim(strTmp) Then
            blnAdd = False
            Exit For
        End If
    Next
    If blnAdd And strTmp <> "" Then
        cbo��λ.AddItem strTmp
    End If
    
End Sub

Private Sub cbo��λ_GotFocus()
    Me.cbo��λ.SelStart = 0: Me.cbo��λ.SelLength = 100
    ImeLanguage True
End Sub

Private Sub cbo��λ_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
       Exit Sub
    Case Else
        zlControl.TxtCheckKeyPress cbo��λ, KeyAscii, m�ı�ʽ
    End Select
End Sub


Private Sub InitCardData(ByVal lng����ID As Long)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ����������Ʒ�ֵĿ�Ƭ����
    '����:lng����-ָ��������ID
    '����:���˺�
    '����:2007/05/24
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHandle
    Me.lblNote.Caption = ""
    If mintEditType <> g�鿴 Then
        gstrSQL = "select distinct ���㵥λ from ������ĿĿ¼ where ��� ='4' and ���㵥λ is not null"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���㵥λ"
        With rsTemp
            cbo��λ.Clear
            Do While Not .EOF
                Me.cbo��λ.AddItem .Fields(0).Value
                .MoveNext
            Loop
        End With
    End If
    
    Me.cbo�����Ա�.Clear
    Me.cbo�����Ա�.AddItem "0-���Ա�����"
    Me.cbo�����Ա�.AddItem "1-����"
    Me.cbo�����Ա�.AddItem "2-Ů��"
    Me.cbo�����Ա�.ListIndex = 0
    
    Me.vsEditBill.Clear 1
    Me.vsEditBill.Rows = 2
    If mintEditType = g���� Then
        Me.tvwClass.Nodes("_" & mlng����id).Selected = True
        Me.txt����.Text = Me.tvwClass.SelectedItem.Text
        Me.txt����.Tag = mlng����id
        Me.txt����.Text = GetMaxCode()
        Me.txt����.Text = ""
        Me.txtƴ��.Text = ""
        Me.txt���.Text = ""
        Me.txtӢ��.Text = ""
        Me.cbo��λ.Text = ""
        Exit Sub
    End If

    '������Ϣ��Ŀ
    gstrSQL = "select I.����ID,I.����,I.����,I.���㵥λ," & _
            "        I.����ʱ��,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ��,Nvl(I.�����Ա�,0) As �����Ա� " & _
            " from ������ĿĿ¼ I" & _
            " where  I.ID=[1]   "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
        
    With rsTemp
        If Not .EOF Then
            With cmbStationNo
                For i = 1 To .ListCount - 1
                    If Mid(cmbStationNo.List(i), 1, InStr(1, cmbStationNo.List(i), "-") - 1) = zlStr.nvl(rsTemp!վ��) Then
                        .ListIndex = i: Exit For
                    End If
                Next
            End With
            Me.lblNote.Caption = "ע���ò��Ͻ�����" & Format(!����ʱ��, "YYYY-MM-DD")
            If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                Me.lblNote.Caption = Me.lblNote.Caption & "����" & Format(!����ʱ��, "YYYY-MM-DD") & "ͣ�á�"
            End If
            
            Me.tvwClass.Nodes("_" & !����id).Selected = True
            Me.txt����.Text = Me.tvwClass.SelectedItem.Text
            Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            mlng����id = Val(Me.txt����.Tag)
            Me.txt����.Text = !����
            Me.txt����.Text = !����
            Me.cbo��λ.Text = zlStr.nvl(!���㵥λ)
            Me.cbo�����Ա�.ListIndex = !�����Ա�
        End If
    End With
       
    '����������Ӣ����
    gstrSQL = "select ����,����,����,���� from ������Ŀ���� where ���� in (1,2) and ������ĿID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    With rsTemp
        Do While Not .EOF
            If !���� = 1 And !���� = 1 Then Me.txtƴ��.Text = zlStr.nvl(!����)
            If !���� = 1 And !���� = 2 Then Me.txt���.Text = zlStr.nvl(!����)
            If !���� = 2 Then Me.txtӢ��.Text = zlStr.nvl(!����)
            .MoveNext
        Loop
    End With
    
    '��������
    gstrSQL = "select N.����,P.���� as ƴ��,W.���� as ���" & _
            " from (select distinct ���� from ������Ŀ���� where ������ĿID=[1] and ����=9) N," & _
            "      (select ����,���� from ������Ŀ���� where ������ĿID=[1] and ����=9 and ����=1) P," & _
            "      (select ����,���� from ������Ŀ���� where ������ĿID=[1] and ����=9 and ����=2) W" & _
            " where N.����=P.����(+) and N.����=W.����(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    
    With rsTemp
        Do While Not .EOF
            If Me.vsEditBill.Rows - 1 < .AbsolutePosition Then Me.vsEditBill.Rows = Me.vsEditBill.Rows + 1
            Me.vsEditBill.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.vsEditBill.TextMatrix(.AbsolutePosition, 1) = zlStr.nvl(!����)
            Me.vsEditBill.TextMatrix(.AbsolutePosition, 2) = zlStr.nvl(!ƴ��)
            Me.vsEditBill.TextMatrix(.AbsolutePosition, 3) = zlStr.nvl(!���)
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEditCtrEnable()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ñ༭�ؼ���Enable����
    '����:���˺�
    '����:2007/05/24
    '------------------------------------------------------------------------------------------------------------------
    
    Dim blnStuffModify As Boolean
    
    If mintEditType = g���� Or mintEditType = g�޸� Then
        blnStuffModify = True
    Else
        cmdOK.Visible = False
    End If
    Me.txt����.Enabled = blnStuffModify
    Me.txt����.Enabled = blnStuffModify
    Me.txt����.Enabled = blnStuffModify
    Me.cmd����.Enabled = blnStuffModify
    Me.txtƴ��.Enabled = blnStuffModify
    Me.txt���.Enabled = blnStuffModify
    Me.txtӢ��.Enabled = blnStuffModify
    Me.cbo��λ.Enabled = blnStuffModify
    Me.cmbStationNo.Enabled = blnStuffModify
    Me.cbo�����Ա�.Enabled = blnStuffModify
    If blnStuffModify Then
        vsEditBill.Editable = flexEDKbdMouse
    Else
        vsEditBill.Editable = flexEDNone
    End If
    
    If blnStuffModify = False Then
        SetCtlBackColor txt����
        SetCtlBackColor txt����
        SetCtlBackColor txt����
        SetCtlBackColor txtƴ��
        SetCtlBackColor txtӢ��
        SetCtlBackColor txt���
    End If
End Sub


Private Sub Form_Activate()

    If mblnFrist = False Then Exit Sub
    mblnFrist = False
    
    '��ʼվ��
    cmbStationNo.Visible = gSystem_Para.bln����վ��
    lblStationNo.Visible = cmbStationNo.Visible
    
    
    '----------������ص����볤��-------------------------------------
    Call GetDefineSize
     
    'ȡ���Ʒ���Ŀ¼����
    Call Load���Ʒ�����Ϣ
     
    '----------��ʼ��Ƭ����-------------------------------------
    Call InitCardData(Val(mstr����ID))
    
    '���ñ༭�ؼ�
    Call SetEditCtrEnable
    If txt����.Enabled Then txt����.SetFocus
End Sub

Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:�Ϸ�,����true,���򷵻�False
    '--����:���˺�
    '--����:2007/05/24
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTmp As String, strTemp As String
    Dim strName As String
    
    ISValied = False
  '�༭���ݼ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������ϱ��룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > mintCodeLength Then
        MsgBox "����ĳ��ȳ��������" & mintCodeLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "������������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > txt����.MaxLength Then
        MsgBox "�������Ƴ��ȳ��������" & txt����.MaxLength & "���ַ���" & txt����.MaxLength \ 2 & "�����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txtƴ��.Text), vbFromUnicode)) > txtƴ��.MaxLength Then
        MsgBox "����ƴ�����볤�ȳ��������" & txtƴ��.MaxLength & "���ַ���" & txtƴ��.MaxLength \ 2 & "�����֣���", vbInformation, gstrSysName
        Me.txtƴ��.SetFocus: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt���.Text), vbFromUnicode)) > txt���.MaxLength Then
        MsgBox "������ʼ��볤�ȳ��������" & txt���.MaxLength & "���ַ���" & txt���.MaxLength \ 2 & "�����֣���", vbInformation, gstrSysName
        Me.txt���.SetFocus: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txtӢ��.Text), vbFromUnicode)) > txtӢ��.MaxLength Then
        MsgBox "Ӣ�����Ƴ��ȳ��������" & txtӢ��.MaxLength & "���ַ���" & txtӢ��.MaxLength \ 2 & "�����֣���", vbInformation, gstrSysName
        Me.txtӢ��.SetFocus: Exit Function
    End If
    If Trim(Me.cbo��λ.Text) = "" Then
        MsgBox "������ɢװ��λ��", vbInformation, gstrSysName
        Me.cbo��λ.SetFocus: Exit Function
    End If
    If zlClinicCodeRepeat(txt����.Text, Val(mstr����ID)) = True Then
        Me.txt����.SetFocus: Exit Function
    End If

    '�������
    strTemp = ";" & Trim(Me.txt����.Text) & ";" & Trim(Me.txtӢ��.Text)
    With Me.vsEditBill
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("��������"))) <> "" Then
                If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(i, .ColIndex("��������"))) & ";") > 0 Then
                    MsgBox "���������ظ�������ͨ�����ƺ�Ӣ�����ƣ���", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("��������")
                    .SetFocus: Exit Function
                Else
                    strTemp = strTemp & ";" & Trim(.TextMatrix(i, .ColIndex("��������")))
                End If
            End If
            If zlCommFun.ActualLen(.TextMatrix(i, .ColIndex("��������"))) > txtӢ��.MaxLength Then
                MsgBox "���������������" & txtӢ��.MaxLength & "���ַ��� " & txtӢ��.MaxLength \ 2 & "���ֺ���,���飡", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("��������")
                .SetFocus: Exit Function
            End If
            If InStr(1, .TextMatrix(i, .ColIndex("��������")), "|") > 0 Then
                MsgBox "�����в��ܰ����ַ���|����", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("��������")
                .SetFocus: Exit Function
            End If
            If zlCommFun.ActualLen(.TextMatrix(i, .ColIndex("�����"))) > txt���.MaxLength Then
                MsgBox "����������������" & txt���.MaxLength & "���ַ��� " & txt���.MaxLength \ 2 & "���ֺ���,���飡", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("�����")
                .SetFocus: Exit Function
            End If
            If zlCommFun.ActualLen(.TextMatrix(i, .ColIndex("ƴ����"))) > txt���.MaxLength Then
                MsgBox "ƴ���������������" & txt���.MaxLength & "���ַ��� " & txt���.MaxLength \ 2 & "���ֺ���,���飡", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("ƴ����")
                .SetFocus: Exit Function
            End If
            
            If InStr(1, .TextMatrix(i, .ColIndex("�����")), "|") > 0 Then
                MsgBox "������в��ܰ����ַ���|����", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("�����")
                .SetFocus: Exit Function
            End If
            If InStr(1, .TextMatrix(i, .ColIndex("ƴ����")), "|") > 0 Then
                MsgBox "ƴ�����в��ܰ����ַ���|����", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("ƴ����")
                .SetFocus: Exit Function
            End If
        Next
    End With
    ISValied = True
End Function

Public Function zlClinicCodeRepeat(str���� As String, Optional lngSelfID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ����������Ŀ������Ƿ������б����ظ����ظ��������ʾ
    '��Σ�strInputCode-����ı��룻lngSelfID-�Լ���ID�ţ����޸�ʱ����Ҫ��������������ж�
    '���Σ��ظ�����True��������Flase
    '����:���˺�
    '����:2007/05/24
    '------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select K.����||' ['||I.����||']'||I.���� as ����" & _
            " from ������ĿĿ¼ I,������Ŀ��� K" & _
            " where I.���=K.���� and I.����=[1] " & _
            "       and I.ID<>[2]"
    err = 0: On Error GoTo ErrHand
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ�����ظ��ı���", str����, lngSelfID)
        
    With rsTmp
        If .RecordCount <> 0 Then
            MsgBox "����Ŀ�롰" & !���� & "�������ظ���", vbExclamation, gstrSysName
            zlClinicCodeRepeat = True
        Else
            zlClinicCodeRepeat = False
        End If
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlClinicCodeRepeat = True
End Function

Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������������Ʒ������
    '--�����:
    '--������:
    '--��  ��:����ɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, intTemp As Integer, i As Long
    Dim strTemp As String
    Dim strվ�� As String
    
    If mintEditType = g���� Then
        lng����ID = sys.NextId("������ĿĿ¼")
        gstrSQL = "zl_����Ʒ��_INSERT("
        Me.cmdOK.Tag = lng����ID
        mlngƷ��id = lng����ID
    Else
        lng����ID = Val(mstr����ID)
        gstrSQL = "zl_����Ʒ��_UPDATE("
        Me.cmdOK.Tag = lng����ID
    End If
    
    strTemp = ""
    With Me.vsEditBill
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("��������"))) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(i, .ColIndex("��������")))
                strTemp = strTemp & "^" & Trim(.TextMatrix(i, .ColIndex("ƴ����")))
                strTemp = strTemp & "^" & Trim(.TextMatrix(i, .ColIndex("�����")))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    '������������
    If LenB(strTemp) > 4000 Then
        vsEditBill.SetFocus
        MsgBox "�����ַ���̫��������ٱ����������߱������ȡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If cmbStationNo.Text = "" Then
        strվ�� = "Null"
    Else
        strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    'Zl_����Ʒ��_Update Or zl_����Ʒ��_INSERT
    '  ����id_In In ������ĿĿ¼.����id%Type := Null,
    '  Id_In     In ������ĿĿ¼.ID%Type,
    '  ����_In   In ������ĿĿ¼.����%Type,
    '  ����_In   In ������ĿĿ¼.����%Type,
    '  ��λ_In   In ������ĿĿ¼.���㵥λ%Type := Null,
    '  ƴ��_In   In ������Ŀ����.����%Type := Null,
    '  ���_In   In ������Ŀ����.����%Type := Null,
    '  Ӣ��_In   In ������Ŀ����.����%Type := Null,
    '  վ��_In   In ������ĿĿ¼.վ��%Type := Null,
    '  ����_In   In Varchar2 := Null --��"|"�ָ��ı�����¼��ÿ����¼��"����^ƴ��^���"��֯
    gstrSQL = gstrSQL & "" & mlng����id & ","
    gstrSQL = gstrSQL & "" & lng����ID & ","
    gstrSQL = gstrSQL & "'" & Trim(Me.txt����.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.txt����.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.cbo��λ.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.txtƴ��.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.txt���.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.txtӢ��.Text) & "',"
    gstrSQL = gstrSQL & IIf(cmbStationNo.Visible = True And Trim(cmbStationNo.Text) <> "", "'" & strվ�� & "'", "NULL") & ","
    gstrSQL = gstrSQL & "" & Left(Me.cbo�����Ա�.Text, 1) & ","
    gstrSQL = gstrSQL & "'" & strTemp & "')"
    
    err = 0: On Error GoTo ErrHand
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmdOK_Click()
    Dim intTemp As Integer
    '�����ҳ����������Ƿ���ȷ
    If ISValied = False Then Exit Sub
    If mintEditType <> g���� And mintEditType <> g�޸� Then
        Unload Me
        Exit Sub
    End If
    
    If SaveData = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    
    If mintEditType = g���� Then
'        intTemp = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��������ģʽ\", "Ʒ��->���", "0"))
'        intTemp = Val(zlDatabase.GetPara("Ʒ�ֹ��ģʽ", glngSys, mlngModule, "0"))
'        If intTemp = 1 Then
'            '��Ҫ���ӹ��
'            Call frmStuffSpec.ShowEditCard(Me, g����, Val(Me.cmdOK.Tag), "", mstrPrivs)
'        End If
''        intTemp = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��������ģʽ\", "Ʒ��", "0"))
'        intTemp = Val(zlDatabase.GetPara("Ʒ������ģʽ", glngSys, mlngModule, "0"))
'        If intTemp = 1 Then
'            Call InitCardData(0)
'            If txt����.Enabled Then Me.txt����.SetFocus
'        Else
'            Unload Me
'            Exit Sub
'        End If
        Select Case ActiveControl
            Case cmdSaveAddItem '��������Ʒ��
                Call InitCardData(0)
                If txt����.Enabled Then Me.txt����.SetFocus
            Case cmdSaveAddSpec '�������ӹ��
                mlngƷ��id = Val(Me.cmdOK.Tag)
                Unload Me
                Call frmStuffSpec.ShowEditCard(frmStuffMgr, g����, mlngƷ��id, mlng����id, "", mstrPrivs)
            Case Else   'ֱ�ӱ����˳�
                Unload Me
        End Select
    Else
        Unload Me
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub
Private Sub cmd����_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Function GetMaxCode() As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�����
    '--�����:
    '--������:
    '--��  ��:�����
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsCode As ADODB.Recordset
    Dim strTemp As String
    Dim intCodeType As Integer
    Dim str���� As String
    
    On Error GoTo ErrHandle
    intCodeType = Val(zlDatabase.GetPara("�������ģʽ", glngSys, mlngModule))
    strTemp = Mid(Me.txt����.Text, 2, InStr(1, Me.txt����.Text, "]") - 2)
    
    If intCodeType = 0 Or Len(strTemp) >= 16 Then
    '0000000001��0000000002
        gstrSQL = "Select Nvl(����, '000000000') As ����" & vbNewLine & _
                        "From (Select ���� From ������ĿĿ¼ Where ��� = '4' Order By Length(����) Desc, ���� Desc)" & vbNewLine & _
                        "Where Rownum = 1"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsTemp
            str���� = zlCommFun.IncStr(!����)
            GetMaxCode = str����
        End With
      
    Else
    
        gstrSQL = "Select a.Id, a.����id, a.����, a.���� From ������ĿĿ¼ A Where ����id =[1] Order By ����id, ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����id)

        If Len(strTemp) >= 7 Then
            str���� = "01"
            str���� = IIf(intCodeType = 1, "4", "") & strTemp & str����
        Else
            str���� = Mid("000000000", 1, 9 - Len(strTemp) - IIf(intCodeType = 1, 1, 0))
            str���� = IIf(intCodeType = 1, "4", "") & strTemp & str����
            str���� = zlCommFun.IncStr(str����)
        End If
        
        GetMaxCode = str����
    
        Do While True
            rsTemp.Filter = ""
            rsTemp.Filter = "����='" & GetMaxCode & "'"
            If rsTemp.RecordCount = 0 Then
                Exit Do
            End If
            GetMaxCode = zlCommFun.IncStr(GetMaxCode)
    
            rsTemp.MoveNext
        Loop
    End If
    
    gstrSQL = "Select ���� From ������ĿĿ¼ "
    Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
    Do While True
        rsCode.Filter = ""
        rsCode.Filter = "����='" & GetMaxCode & "'"
        If rsCode.RecordCount = 0 Then
            Exit Do
        End If
        GetMaxCode = zlCommFun.IncStr(GetMaxCode)
    Loop
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Load���Ʒ�����Ϣ()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:����ѡ��
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As Node
    
    On Error GoTo ErrHandle
    '����ѡ����װ��
    gstrSQL = "select ID,�ϼ�ID,����,����,����" & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� =7 " & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !Id, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !Id, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!����), "", !����)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        err = 0: On Error Resume Next
        Me.tvwClass.Nodes("_" & mlng����id).Selected = True
        Me.txt����.Text = Me.tvwClass.SelectedItem.Text
        mlng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim rsrecord As ADODB.Recordset
    
    On Error GoTo ErrHandle
    mblnFrist = True
'    With cmbStationNo
'        .Clear
'        .AddItem ""
'        .AddItem "0"
'        .AddItem "1"
'        .AddItem "2"
'        .AddItem "3"
'        .AddItem "4"
'        .AddItem "5"
'        .AddItem "6"
'        .AddItem "7"
'        .AddItem "8"
'        .AddItem "9"
'        .ListIndex = 0
'    End With
    strSql = "select ���,���� from zlnodelist"
    Set rsrecord = zlDatabase.OpenSQLRecord(strSql, "վ���ѯ")
    With cmbStationNo
        .AddItem ""
        Do While Not rsrecord.EOF
            .AddItem rsrecord!��� & "-" & rsrecord!����
            rsrecord.MoveNext
        Loop
    End With
    If mintEditType <> g���� Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If tvwClass.Visible Then
        tvwClass.Visible = False
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub txt����_Change()
    mblnChange = True
End Sub

Private Sub txt����_GotFocus()
    ImeLanguage False
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        End Select
        KeyAscii = 0
End Sub

Private Sub txt����_LostFocus()
    ImeLanguage False
End Sub

Private Sub txt����_Change()
    mlng����id = 0
    txt����.Tag = ""
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt����.Text = Me.tvwClass.SelectedItem.Text
    mlng����id = Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
    txt����.Tag = mlng����id
    txt����.Text = GetMaxCode
    If txt����.Enabled Then Me.txt����.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd���� Is ActiveControl Then
        Exit Sub
    End If
    Me.tvwClass.Visible = False
End Sub

Private Sub txt����_Change()
    mblnChange = True
    'ƴ�������
    Me.txtƴ��.Text = zlStr.GetCodeByORCL(Me.txt����.Text, 0, Me.txtƴ��.MaxLength)
    Me.txt���.Text = zlStr.GetCodeByORCL(Me.txt����.Text, 1, Me.txt���.MaxLength)
End Sub

Private Sub txt����_GotFocus()
    ImeLanguage True
    zlControl.TxtSelAll txt����
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����_LostFocus()
    ImeLanguage False
End Sub

Private Sub txtƴ��_Change()
    mblnChange = True
End Sub

Private Sub txtƴ��_GotFocus()
    ImeLanguage False
End Sub

Private Sub txtƴ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub txt���_Change()
    mblnChange = True
    
End Sub

Private Sub txt���_GotFocus()
    ImeLanguage False
End Sub

Private Sub txt���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub txtӢ��_Change()
    mblnChange = True
End Sub

Private Sub txtӢ��_GotFocus()
    ImeLanguage False
    zlControl.TxtSelAll txtӢ��
End Sub

Private Sub txtӢ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

 
Private Sub vsEditBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strKey As String
    With vsEditBill
        Select Case Col
        Case .ColIndex("��������")
            strKey = Trim(.TextMatrix(.Row, .Col))
            If strKey = "" Then Exit Sub
            .TextMatrix(Row, .ColIndex("ƴ����")) = zlStr.GetCodeByORCL(strKey, 0, Me.txtƴ��.MaxLength)
            .TextMatrix(Row, .ColIndex("�����")) = zlStr.GetCodeByORCL(strKey, 1, Me.txt���.MaxLength)
            If .Row = .Rows - 1 Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        Case .ColIndex("ƴ����")
            If Trim(.TextMatrix(.Row, .ColIndex("��������"))) = "" Then Exit Sub
            If .Row = .Rows - 1 Then
                .Rows = .Rows + 1: .Row = .Rows - 1
            End If
        Case Else
        End Select
    End With
    '�����к�
    Call RedoRowNo
End Sub

Private Sub vsEditBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsEditBill
        Select Case Col
        Case .ColIndex("��������")
        Case .ColIndex("�����"), .ColIndex("ƴ����")
            If Trim(.Cell(flexcpData, Row, .ColIndex("��������"))) = "" Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsEditBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsEditBill_EnterCell()
    If mintEditType = g�鿴 Then Exit Sub
    If vsEditBill.Col = vsEditBill.ColIndex("��������") Then
        OS.OpenIme True
        vsEditBill.EditMaxLength = Me.txt����.MaxLength
    Else
        OS.OpenIme False
    End If
End Sub

Private Sub vsEditBill_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    If mintEditType <> g�鿴 Then
        With vsEditBill
            If KeyCode = vbKeyDelete Then
                If MsgBox("���Ƿ����Ҫɾ�����еĲ��ϱ�����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
                If .Row = .Rows - 1 And .Row = 1 Then
                    For lngCol = 0 To .Cols - 1
                        .TextMatrix(.Row, lngCol) = ""
                        .Cell(flexcpData, .Row, lngCol) = ""
                    Next
                Else
                    .RemoveItem .Row
                End If
            End If
            Call RedoRowNo
        End With
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsEditBill
        If Trim(.TextMatrix(.Row, .ColIndex("��������"))) = "" Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Select Case .Col
        Case .ColIndex("�����")
            .Col = .ColIndex("��������")
            If .Row >= .Rows - 1 Then
                If mintEditType = g�鿴 Then
                Else
                    .Rows = .Rows + 1
                End If
                .Row = .Rows - 1
            Else
                .Row = .Row + 1
            End If
            .SetFocus
        Case Else
            OS.PressKey vbKeyRight
        End Select
    End With
End Sub

Private Sub vsEditBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsEditBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = Asc("^") Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Col < vsEditBill.ColIndex("�����") Then
            If Col = vsEditBill.ColIndex("��������") Then
                OS.PressKey vbKeyDown
            Else
                OS.PressKey vbKeyRight
            End If
        End If
        Exit Sub
    End If
    
    With vsEditBill
        Select Case Col
        Case .ColIndex("��������"), .ColIndex("ƴ����"), .ColIndex("�����")
            Call VsFlxGridCheckKeyPress(vsEditBill, Row, Col, KeyAscii, m�ı�ʽ)
        Case Else
        End Select
    End With
End Sub

Private Sub vsEditBill_LeaveCell()
    If mintEditType = g�鿴 Then Exit Sub
    OS.OpenIme False
End Sub

Private Sub vsEditBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    
    If mintEditType = g�鿴 Then Cancel = True: Exit Sub
    
    strKey = Trim(vsEditBill.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")
    With vsEditBill
        Select Case Col
        Case .ColIndex("��������")
            If zlCommFun.ActualLen(strKey) > txtӢ��.MaxLength Then
                ShowMsgBox "�������Ʊ���ΪС�ڵ���" & txtӢ��.MaxLength & "���ַ���" & Int(txtӢ��.MaxLength / 2) & "������,���������룡"
                Cancel = True
                Exit Sub
            End If
        Case .ColIndex("�����"), .ColIndex("ƴ����") '
            If zlCommFun.ActualLen(strKey) > txt���.MaxLength Then
                ShowMsgBox vsEditBill.TextMatrix(0, Col) & "����ΪС�ڵ���" & txt���.MaxLength & "���ַ���" & txt���.MaxLength \ 2 & "������,���������룡"
                Cancel = True
                Exit Sub
            End If
        End Select
    End With
End Sub
Private Sub RedoRowNo()
    '------------------------------------------------------------------------------
    '����:�����к�
    '����:
    '����:���˺�
    '����:2007/08/14
    '------------------------------------------------------------------------------
    Dim i As Long
    With vsEditBill
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("�к�")) = i
        Next
    End With

End Sub

Private Sub cmbStationNo_Change()
    mblnChange = True
End Sub

Private Sub cmbStationNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey vbKeyTab
    
End Sub
