VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInsCompend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7410
   Icon            =   "frmInsCompend.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picVBar_S 
      BackColor       =   &H8000000C&
      Height          =   4815
      Left            =   2565
      MouseIcon       =   "frmInsCompend.frx":058A
      MousePointer    =   99  'Custom
      ScaleHeight     =   4815
      ScaleWidth      =   30
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   30
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   2655
      ScaleHeight     =   4260
      ScaleWidth      =   4575
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   15
      Width           =   4575
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1185
         Width           =   3555
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����(&K),����дʱ����ٶβ�����ɾ��"
         Height          =   210
         Left            =   855
         TabIndex        =   8
         Top             =   2985
         Width           =   3585
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   2250
         TabIndex        =   9
         Top             =   3765
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   3360
         TabIndex        =   10
         Top             =   3765
         Width           =   1100
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   3600
         Width           =   4410
      End
      Begin VB.OptionButton optԤ�� 
         Caption         =   "�����Զ������(&2)"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Top             =   705
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton optԤ�� 
         Caption         =   "����Ԥ�����(&1)"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Top             =   435
         Width           =   2775
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   855
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1590
         Width           =   3555
      End
      Begin VB.TextBox txt˵�� 
         Height          =   840
         Left            =   855
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2010
         Width           =   3555
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   1005
         Width           =   4410
      End
      Begin VB.Label lblԤ�� 
         AutoSize        =   -1  'True
         Caption         =   "ע��:   ��ӦԤ����١���###"
         Height          =   180
         Left            =   150
         TabIndex        =   18
         Top             =   3300
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   150
         TabIndex        =   4
         Top             =   1620
         Width           =   630
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   60
         Picture         =   "frmInsCompend.frx":06DC
         Top             =   75
         Width           =   480
      End
      Begin VB.Label lblҪ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������û���ѽ�����Ԥ�������ѡ����롣"
         Height          =   180
         Left            =   690
         TabIndex        =   15
         Top             =   135
         Width           =   3870
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         Caption         =   "˵��(&S)"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   2010
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Left            =   150
         TabIndex        =   2
         Top             =   1245
         Width           =   630
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   4500
      Left            =   0
      TabIndex        =   12
      Top             =   240
      Width           =   2370
      _cx             =   4180
      _cy             =   7937
      Appearance      =   2
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInsCompend.frx":0FA6
      ScrollTrack     =   -1  'True
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
      WordWrap        =   -1  'True
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
   Begin VB.Label lblTitle 
      BackColor       =   &H80000003&
      Caption         =   " ��ѡԤ�����(&P)"
      Height          =   225
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2370
   End
End
Attribute VB_Name = "frmInsCompend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'################################################################################################################
'## �ֲ�����
'################################################################################################################

Private EditMode As EditModeEnum        '�༭��ʽ���������޸ģ�
Private frmParent As frmMain            '������
Private Compends As New cEPRCompends    '����ټ���
Private Compend As New cEPRCompend      '������ٶ���
Private edtThis As Object               '�༭��

Private mblnOK As Boolean               '��ʱ��������ʾ�Ƿ񱣴��޸Ľ����
Private mlngOldKey As Long              '��ʱ����������ԭʼKeyֵ��
Private mblnWithName As Boolean         '��ʱ�������Ƿ�����������ַ��������������ʱ�Ƿ�ѡ����һ���ı���Ϊ���ƣ�

'################################################################################################################
'## ���ܣ�  ��ʾ��ٱ༭����
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub ShowMe(ByRef oParentForm As frmMain, ByRef oedtThis As Object, _
    ByRef oCompends As cEPRCompends, _
    Optional ByVal oCompend As cEPRCompend)
    
    Dim lngCurCompKey As Long, i As Long, j As Long
    Dim ArrayKeys() As Long
    
    Set frmParent = oParentForm
    Set Compends = oCompends
    Set edtThis = oedtThis
    Call FillPreDefinedComps(frmParent.Document.EPRFileInfo.����)
    
    mblnWithName = False
    If oCompend Is Nothing Then
        EditMode = cprEM_����
        Me.Caption = "�������"
        If edtThis.Selection.Text <> "" Then
            Me.txt����.Text = MidB(edtThis.Selection.Text, 1, 40)      '�������Ƴ���Ϊ40��
            mblnWithName = True
        End If
        Compends.UpdateOrdersFromText edtThis   '��������ڲ����
        
        lngCurCompKey = frmParent.Document.GetCurCompendKey(frmParent.Editor1)
    Else
        EditMode = cprEM_�޸�
        Me.Caption = "�޸����"
        Set Compend = oCompend.Clone(True)
        mlngOldKey = Compend.Key
        Me.txt����.Text = Compend.����
        txt˵��.Text = Compend.˵��
        Me.cbo����.Tag = Compend.��Key
        If Compend.�������� Then chk����.Value = vbChecked
        Me.lblԤ��.Tag = Compend.Ԥ�����ID
        
        Compends.UpdateOrdersFromText edtThis   '��������ڲ����
        
        Dim lS As Long, lE As Long
        Compend.GetPosition frmParent.Editor1, lS, lE
        lngCurCompKey = frmParent.Document.GetCurCompendKey(frmParent.Editor1, lS)
    
    End If
    
    '����ϼ��б���������Ϊ�ϼ����ǵ�ǰλ��ǰ����٣��Ҵ����һ��һ����ٿ�ʼ�������һ�������ֹ��
    With Me.cbo����
        .Clear
        .AddItem "1��": .ItemData(.NewIndex) = 0
        If lngCurCompKey > 0 Then
            ' 'ѭ���ҳ�����ٵĸ������
            i = lngCurCompKey
            ReDim ArrayKeys(1 To 1) As Long
            ArrayKeys(1) = i
            Do While i > 0
                i = Compends.GetParentNodeKey(i)
                If i > 0 Then
                    ReDim Preserve ArrayKeys(1 To UBound(ArrayKeys) + 1) As Long
                    ArrayKeys(UBound(ArrayKeys)) = i
                End If
            Loop
            'װ���Ѿ��������ټ���
            For j = UBound(ArrayKeys) To 1 Step -1
                .AddItem .ListCount + 1 & "��(��" & Compends("K" & ArrayKeys(j)).���� & "�����¼�)"
                .ItemData(.NewIndex) = ArrayKeys(j)
                If .ItemData(.NewIndex) = Val(.Tag) Then .ListIndex = .NewIndex
                If .ListCount >= 8 Then Exit For
            Next
        End If
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    
    If Val(Me.lblԤ��.Tag) <> 0 Then
        With Me.vfgThis
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) = Me.lblԤ��.Tag Then
                    .Row = i
                    Call vfgThis_DblClick
                Else
                    If i = .Rows - 1 Then Me.lblԤ��.Tag = 0: Me.lblԤ��.ToolTipText = ""
                End If
            Next
        End With
    End If
    Me.Show vbModal, frmParent
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk����_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngKey As Long, lngStart As Long, lngEnd As Long, i As Long
    
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > 40 Then MsgBox "���Ƴ��������40���ַ���20�����֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt˵��.Text), vbFromUnicode)) > 500 Then MsgBox "˵�����������500���ַ���250�����֣���", vbInformation, gstrSysName: Me.txt˵��.SetFocus: Exit Sub
'    '��֤Ԥ����ٲ��ظ�
'    If optԤ��(0).Value Then
'        With frmParent.Document
'            For i = 1 To .Compends.Count
'                If .Compends(i).Ԥ�����ID <> 0 Then
'                    If Val(Me.lblԤ��.Tag) = .Compends(i).Ԥ�����ID Then
'                        MsgBox "Ԥ����ٲ������ظ���", vbOKOnly + vbInformation, Me.Caption
'                        Exit Sub
'                    End If
'                End If
'            Next
'        End With
'    End If
    If Len(Trim(txt����)) = 0 Then
        MsgBox "��������������ƣ����������룡", vbOKOnly + vbInformation, Me.Caption
        Me.txt����.SetFocus: Exit Sub
    End If
    If Me.cbo����.ListIndex = 0 Then
        Compend.��ID = 0
        Compend.��Key = 0
        Compend.Level = 1
    Else
        Compend.��ID = 0
        Compend.��Key = Compends("K" & Me.cbo����.ItemData(Me.cbo����.ListIndex)).Key
        Compend.Level = Compends("K" & Me.cbo����.ItemData(Me.cbo����.ListIndex)).Level + 1
    End If
    Compend.���� = Trim(txt����)
    Compend.˵�� = Trim(txt˵��)
    If optԤ��(0).Value Then
        Compend.Ԥ�����ID = Val(Me.lblԤ��.Tag)
    Else
        Compend.Ԥ�����ID = 0
    End If
    Compend.�������� = (Me.chk����.Value = vbChecked)
    
    If EditMode = cprEM_�޸� Then
        Compends.Remove "K" & mlngOldKey
        Compend.Key = mlngOldKey
        lngKey = Compends.AddExistNode(Compend, True)
        mlngOldKey = lngKey
    Else
        lngKey = Compends.AddExistNode(Compend, False)
        mlngOldKey = lngKey
    End If
    
    '�����޸������������ٵĸ�Key�Ĺ�ϵ�仯
    Call UpdateParentKeys
    
    If EditMode = cprEM_���� Then
        frmParent.InsertCompend frmParent.Editor1.Selection.StartPos, frmParent.Editor1.Selection.EndPos, Compends("K" & lngKey), True
        If mblnWithName Then
            lngEnd = lngEnd + 32
            edtThis.Range(lngEnd, lngEnd).Selected
        End If
    Else
        frmParent.ModifyCompend Compends("K" & lngKey)
    End If
    
    Compends.UpdateOrdersFromText edtThis
    Compends.FillTree frmParent.mfrmCompends.Tree, lngKey
    
    mblnOK = True
    Unload Me
End Sub

Private Sub UpdateParentKeys()
    '##################################################################################################################
    '����������� Comp0 ��Ӱ�쵽��������٣�ǰһ��ͬ������ Comp1 �����һ��ͬ������ Comp2 ֮���������٣�
    'ԭ�� 1������Comp1��Comp0֮�����٣����Comp0��Comp2֮���������Է�Χ1�����Ϊ����ٵģ������Key��
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lngLevel As Long, lngCurPos As Long, lngOldPos As Long
    Dim i As Long, j As Long, blnFinded As Boolean, blnFirst As Boolean
    Dim ArrayPrevKeys() As Long, ArrayNextKeys() As Long, lngPrevCount As Long, lngNextCount As Long
    
    Compends.UpdateOrdersFromText edtThis       '����˳��
    lngLevel = Compend.Level
    
    '-------------------------------
    
    lngCurPos = edtThis.Selection.StartPos
    lngOldPos = lngCurPos
    
    ReDim ArrayPrevKeys(1 To 1) As Long
    blnFinded = False
    blnFirst = True
    lngPrevCount = 0
    blnFinded = FindPrevKey(edtThis, lngCurPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    Do While blnFinded
        If lKey <> Compend.Key Then
            If Compends("K" & lKey).Level < lngLevel Then
                Exit Do
            ElseIf Compends("K" & lKey).Level = lngLevel Then
                '�ҵ�ǰһ��ͬ��ε����Comp1
                If blnFirst Then
                    ArrayPrevKeys(1) = lKey
                    blnFirst = False
                Else
                    i = UBound(ArrayPrevKeys) + 1
                    ReDim Preserve ArrayPrevKeys(1 To i) As Long
                    ArrayPrevKeys(i) = lKey
                End If
                lngPrevCount = lngPrevCount + 1
                Exit Do
            Else
                If blnFirst Then
                    ArrayPrevKeys(1) = lKey
                    blnFirst = False
                Else
                    i = UBound(ArrayPrevKeys) + 1
                    ReDim Preserve ArrayPrevKeys(1 To i) As Long
                    ArrayPrevKeys(i) = lKey
                End If
                lngPrevCount = lngPrevCount + 1
            End If
        End If
        lngCurPos = lSS
        blnFinded = FindPrevKey(edtThis, lngCurPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    Loop
    
    '-------------------------------
    
    ReDim ArrayNextKeys(1 To 1) As Long
    lngCurPos = lngOldPos
    
    blnFinded = False
    blnFirst = True
    lngNextCount = 0
    blnFinded = FindNextKey(edtThis, lngCurPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    Do While blnFinded
        If lKey <> Compend.Key Then
            If Compends("K" & lKey).Level < lngLevel Then
                Exit Do
            ElseIf Compends("K" & lKey).Level = lngLevel Then
                '�ҵ�ǰһ��ͬ��ε����Comp1
                If blnFirst Then
                    ArrayNextKeys(1) = lKey
                    blnFirst = False
                Else
                    i = UBound(ArrayNextKeys) + 1
                    ReDim Preserve ArrayNextKeys(1 To i) As Long
                    ArrayNextKeys(i) = lKey
                End If
                lngNextCount = lngNextCount + 1
                Exit Do
            Else
                If blnFirst Then
                    ArrayNextKeys(1) = lKey
                    blnFirst = False
                Else
                    i = UBound(ArrayNextKeys) + 1
                    ReDim Preserve ArrayNextKeys(1 To i) As Long
                    ArrayNextKeys(i) = lKey
                End If
                lngNextCount = lngNextCount + 1
            End If
        End If
        lngCurPos = lEE
        blnFinded = FindNextKey(edtThis, lngCurPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    Loop
    
    '-------------------------------

    If lngPrevCount > 0 And lngNextCount > 0 Then
        '���������ϵ
        For i = 1 To UBound(ArrayPrevKeys)
            For j = 1 To UBound(ArrayNextKeys)
                If Compends("K" & ArrayNextKeys(j)).��Key = ArrayPrevKeys(i) Then
                    Compends("K" & ArrayNextKeys(j)).��Key = 0
                    Compends("K" & ArrayNextKeys(j)).��ID = 0
                End If
            Next
        Next
    End If
    '##################################################################################################################
End Sub

Private Sub FillPreDefinedComps(ByVal lngKind As Long)
    '---------------------------------------------
    '���ܣ���д�����ļ�Ŀ¼
    '---------------------------------------------
    Dim RS As New ADODB.Recordset
    Dim i As Long
    gstrSQL = "Select Id, To_Char(�������, '000') As ���, �����ı� As ����, �������� As ˵��" & _
            " From �����ļ��ṹ" & _
            " Where �ļ�id Is Null And Substr(ʹ��ʱ��, [1], 1) = '1'" & _
            " Order By �������"
    Err = 0: On Error GoTo errHand
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKind)
    With Me.vfgThis
        .Clear
        Set .DataSource = RS
        .ColWidth(0) = 0
        For i = .FixedCols To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    Me.picVBar_S.BackColor = Me.BackColor
End Sub

Private Sub Form_Resize()
    Dim lngHWarp As Long, lngWWarp As Long
    lngHWarp = Me.Height - Me.ScaleHeight
    lngWWarp = Me.Width - Me.ScaleWidth
    With Me.picVBar_S
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
        If .Left < 0 Then .Left = 0
        If .Left > 5000 Then .Left = 5000
    End With
    With Me.lblTitle
        .Top = Me.ScaleTop
        .Left = Me.ScaleLeft: .Width = Me.picBack.Left - .Left - 30
    End With
    With Me.vfgThis
        .Top = Me.lblTitle.Top + Me.lblTitle.Height + 15: .Height = Me.ScaleHeight - .Top
        .Left = Me.ScaleLeft: .Width = Me.picBack.Left - .Left - 30
    End With
    With Me.picBack
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width
        .Top = Me.ScaleTop
    End With
    Me.Width = Me.picBack.Left + Me.picBack.Width + lngWWarp
    Me.Height = Me.picBack.Height + lngHWarp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set Compends = Nothing
    Set Compend = Nothing
    Set edtThis = Nothing
     Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optԤ��_Click(Index As Integer)
    If Me.optԤ��(0).Value And Val(Me.lblԤ��.Tag) = 0 Then
        MsgBox "����ͨ��˫���б��е�ĳһ����ѡ������Ԥ����٣�", vbExclamation, gstrSysName
        Me.optԤ��(0).Value = False: Me.optԤ��(1).Value = True
        Exit Sub
    End If
    If Me.optԤ��(0).Value Then
        Me.lblԤ��.Visible = True
        If Val(Me.lblԤ��.ToolTipText) = 1 Then
            Me.cbo����.ListIndex = 0: Me.cbo����.Enabled = False
            Me.chk����.Value = vbChecked: Me.chk����.Enabled = False
        Else
            Me.cbo����.Enabled = True: Me.chk����.Enabled = True
        End If
    Else
        Me.cbo����.Enabled = True: Me.chk����.Enabled = True: Me.lblԤ��.Visible = False
    End If
End Sub

Private Sub optԤ��_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub optԤ��_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picVBar_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picVBar_S.Left = Me.picVBar_S.Left + x: Me.picVBar_S.BackColor = RGB(192, 192, 192)
End Sub

Private Sub picVBar_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.picVBar_S.BackColor = Me.BackColor
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_Change()
    ValidControlText txt˵��
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgThis_DblClick()
    If vfgThis.Row = 0 Then Exit Sub
    Me.lblԤ��.Tag = Me.vfgThis.TextMatrix(Me.vfgThis.Row, 0)
    Me.lblԤ��.Caption = "ע��:   ��ӦԤ����١���" & Me.vfgThis.TextMatrix(Me.vfgThis.Row, 2)
    Me.lblԤ��.ToolTipText = Val(Me.vfgThis.TextMatrix(Me.vfgThis.Row, 1))
    Me.optԤ��(0).Value = True
    If EditMode <> cprEM_�޸� Then
        Me.txt���� = Me.vfgThis.TextMatrix(Me.vfgThis.Row, 2)
        Me.txt˵�� = Me.vfgThis.TextMatrix(Me.vfgThis.Row, 3)
    End If
    Call optԤ��_Click(0)
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vfgThis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vfgThis_DblClick
        KeyAscii = 0
    End If
End Sub


