VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmLabBloodSugar 
   Caption         =   "�������ϲ�"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11760
   Icon            =   "frmLabBloodSugar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11760
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicList1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   90
      ScaleHeight     =   1155
      ScaleWidth      =   2595
      TabIndex        =   12
      Top             =   3750
      Width           =   2625
      Begin VSFlex8Ctl.VSFlexGrid vfgList1 
         Height          =   2595
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   8685
         _cx             =   15319
         _cy             =   4577
         Appearance      =   0
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
   End
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   2880
      ScaleHeight     =   2715
      ScaleWidth      =   5355
      TabIndex        =   10
      Top             =   1440
      Width           =   5385
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   2595
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   8685
         _cx             =   15319
         _cy             =   4577
         Appearance      =   0
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
   End
   Begin VB.Frame frmList 
      Caption         =   "�ϲ��걾"
      Height          =   1035
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   11265
      Begin VB.TextBox txtNumber 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   2235
      End
      Begin VB.ComboBox cboMachine 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   2235
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&X)"
         Height          =   350
         Left            =   7530
         TabIndex        =   2
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�ϲ�(&U)"
         Height          =   350
         Left            =   6090
         TabIndex        =   1
         Top             =   180
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Left            =   4170
         TabIndex        =   5
         Top             =   210
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         Format          =   100073473
         CurrentDate     =   39682
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�� �� ��"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lbl�ϲ��걾 
         AutoSize        =   -1  'True
         Caption         =   "������Ҫ�ϲ��ı걾��!"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   3360
         TabIndex        =   8
         Top             =   660
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�걾ʱ��"
         Height          =   180
         Left            =   3360
         TabIndex        =   6
         Top             =   270
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList Imglist 
      Left            =   480
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":6852
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":6DEC
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":7386
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":7920
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":7EBA
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":8454
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":87EE
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":8B88
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBloodSugar.frx":8F22
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabBloodSugar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlngMachine As Long                 '����ID
Private mlngSample As Long                  '�걾ID

Private Enum mCol
    ID = 0
    �걾��
    ������Ŀ
    ������
    ��־
    ������Ŀid
    ���ϲ���ĿID
End Enum

Private Enum mUCol
    ID = 0
    �걾��
    ������Ŀ
    ������Ŀid
    ������
    ��־
    ������
    ѡ��
End Enum

Private Sub cboMachine_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord
    Dim intColCount As Integer
    Dim strItems As String
    Dim strStartDate As String, strEndDate As String
    Dim intLoop As Integer

    On Error GoTo errH
    strStartDate = GetDateTime("��  ��", 1, Me.dtpDate)
    strEndDate = GetDateTime("��  ��", 2, Me.dtpDate)
    '����ϲ��걾
    gstrSql = "select /*+ rule */ id,�걾���,����ʱ��,decode(������Դ,2,סԺ��,�����) as ��ʶ��,������, " & vbNewLine & _
                "������Ŀ,����,�Ա�,���� from ����걾��¼ where id = [1]  and ����ʱ�� between [2] and [3] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSample, CDate(strStartDate), CDate(strEndDate))

    If rsTmp.EOF = False Then
        lbl�ϲ��걾 = "�걾��:" & Nvl(rsTmp("�걾���")) & "   ����:" & Nvl(rsTmp("����")) & "   �Ա�:" & Nvl(rsTmp("�Ա�")) & _
                      "   ����:" & Nvl(rsTmp("����")) & "    ������Ŀ:" & Nvl(rsTmp("������Ŀ"))
        Me.txtNumber.Text = Nvl(rsTmp("�걾���")): Me.txtNumber.Tag = Nvl(rsTmp("ID"))
    Else
        lbl�ϲ��걾 = "������걾��ѡ��һ���ϲ��ı걾��"
        Me.txtNumber.Text = "": Me.txtNumber.Tag = ""
    End If


    gstrSql = " Select B.ID, B.����, B.������, B.Ӣ����" & vbNewLine & _
            " From ����������Ŀ A, ����������Ŀ B" & vbNewLine & _
            " Where A.����id = [1] And Nvl(A.��������Ŀ, 0) = -1 And A.��Ŀid = B.ID order by b.����"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.cboMachine.ItemData(Me.cboMachine.ListIndex)))
    Me.vfgList.Rows = 1
    Do Until rsTmp.EOF
        With Me.vfgList
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mCol.������Ŀ) = Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")"
            .TextMatrix(.Rows - 1, mCol.������Ŀid) = Nvl(rsTmp("ID"))
            strItems = strItems & "," & Nvl(rsTmp("ID"))
        End With

        rsTmp.MoveNext
    Loop
    If Me.vfgList.Rows > 1 Then Me.vfgList.Row = 1
    
    If strItems <> "" Then strItems = Mid(strItems, 2)

    gstrSql = "Select /*+ rule */ A.ID,a.�걾���,c.������,c.Ӣ����,b.������,b.�����־,b.������ĿID, " & vbNewLine & _
            "Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As �����־ " & vbNewLine & _
            " From ����걾��¼ A, ������ͨ��� B,����������Ŀ C" & vbNewLine & _
            " Where A.ID = B.����걾id And A.����id = [1] And A.����ʱ�� Between  [2] And [3] And" & vbNewLine & _
            "      B.������Ŀid In (Select * From Table(Cast(f_Num2list([4]) As ZLTOOLS.t_Numlist)))" & vbNewLine & _
            " And a.ҽ��id is null And B.������Ŀid = C.ID and a.ID <> [5] order by �걾��� "
    strStartDate = GetDateTime("��  ��", 1, Me.dtpDate)
    strEndDate = GetDateTime("��  ��", 2, Me.dtpDate)
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.cboMachine.ItemData(Me.cboMachine.ListIndex)), _
                    CDate(strStartDate), CDate(strEndDate), strItems, mlngSample)
    Me.vfgList1.Rows = 1
    Do Until rsTmp.EOF
        With Me.vfgList1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mUCol.ID) = Nvl(rsTmp("ID"))
            .TextMatrix(.Rows - 1, mUCol.�걾��) = Nvl(rsTmp("�걾���"))
            .TextMatrix(.Rows - 1, mUCol.������Ŀ) = Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")"
            .TextMatrix(.Rows - 1, mUCol.������) = Nvl(rsTmp("������"))
            .TextMatrix(.Rows - 1, mUCol.��־) = Nvl(rsTmp("�����־"))
            .TextMatrix(.Rows - 1, mUCol.������Ŀid) = Nvl(rsTmp("������ĿID"))
        End With

        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim lngSourceID As Long             '�ϲ�ID
    Dim intLoop As Integer
    
    On Error GoTo errH
    If txtNumber.Tag = "" Then
        MsgBox "��ѡ��һ���걾�����кϲ�!"
        Me.txtNumber.SetFocus
        Exit Sub
    End If
    lngSourceID = txtNumber.Tag
    gstrSql = ""
    gcnOracle.BeginTrans
    
    With Me.vfgList
        For intLoop = 1 To .Rows - 1
            If .TextMatrix(intLoop, mCol.ID) <> "" Then
                '��ʼ�ϲ�
                gstrSql = "Zl_����������_Union(" & lngSourceID & "," & .TextMatrix(intLoop, mCol.ID) & _
                          "," & .TextMatrix(intLoop, mCol.������Ŀid) & ")"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            End If
        Next
    End With
    gcnOracle.CommitTrans
    If gstrSql = "" Then
        MsgBox "û���ҵ����Ժϲ�����Ŀ!", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.txtNumber.Text = ""
    Me.txtNumber.SetFocus
    cboMachine_Click
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = frmList.hWnd
        Case 2
            Item.Handle = vfgList.hWnd
        Case 3
            Item.Handle = vfgList1.hWnd
    End Select
End Sub

Private Sub dtpDate_Change()
    Call cboMachine_Click
End Sub

Private Sub Form_Load()
    Dim Column As ReportColumn
    Dim rsTmp As New ADODB.Recordset
    
    Call CreateDockPane

    Call InitList
    
    '========================================��������=======================================
    On Error GoTo errH
    Me.dtpDate = Now
    gstrSql = "select  ID,���� , ���� from �������� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        Me.cboMachine.AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
        Me.cboMachine.ItemData(Me.cboMachine.NewIndex) = rsTmp("ID")
        If rsTmp("ID") = mlngMachine Then
            Me.cboMachine.ListIndex = Me.cboMachine.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If Me.cboMachine.ListIndex = -1 Then Me.cboMachine.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane
    Dim intLoop As Integer
    On Error Resume Next

    If Me.WindowState = 1 Then Exit Sub

    Set Pane1 = Me.dkpMain.FindPane(1)
    Pane1.MinTrackSize.SetSize 6954 / Screen.TwipsPerPixelX, 1000 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize Pane1.MaxTrackSize.Width, 1000 / Screen.TwipsPerPixelY
    
  
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
    
End Sub








Public Sub ShowMe(Objfrm As Object, lngMachine As Long, lngSampleID As Long)
    mlngMachine = lngMachine
    mlngSample = lngSampleID
    Me.Show vbModal, Objfrm
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub picList_Resize()
    With Me.vfgList
        .Top = 0
        .Left = 0
        .Width = Me.PicList.Width
        .Height = Me.PicList.Height
    End With
End Sub

Private Sub PicList1_Resize()
    With Me.vfgList1
        .Top = 0
        .Left = 0
        .Width = Me.PicList1.Width
        .Height = Me.PicList1.Height
    End With
End Sub

Private Sub txtNumber_GotFocus()
    Me.txtNumber.SelStart = 0
    Me.txtNumber.SelLength = Len(Me.txtNumber)
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strStartDate As String, strEndDate As String
    
    On Error GoTo errH
    
    If KeyAscii = 13 Then
        strStartDate = GetDateTime("��  ��", 1, Me.dtpDate)
        strEndDate = GetDateTime("��  ��", 2, Me.dtpDate)
        gstrSql = "select id from ����걾��¼ where �걾��� = [1] and ����ʱ�� between [2] and [3] and ����ID =[4] " & vbNewLine & _
                  " and ����״̬ <> 2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.txtNumber.Text, CDate(strStartDate), CDate(strEndDate), mlngMachine)
        If rsTmp.EOF = True Then
            MsgBox "û���ҵ��걾�������Ǳ걾�����ڻ�걾����ˣ�", vbInformation, Me.Caption
            Me.txtNumber.Tag = ""
            lbl�ϲ��걾.Caption = "������Ҫ�ϲ��ı걾��!"
            Me.txtNumber.SelStart = 0
            Me.txtNumber.SelLength = Len(Me.txtNumber)
            Me.txtNumber.SetFocus
            Exit Sub
        End If
        Me.txtNumber.SelStart = 0
        Me.txtNumber.SelLength = Len(Me.txtNumber)
        Me.txtNumber.Tag = Nvl(rsTmp("ID"))
        mlngSample = rsTmp("ID")
        Call cboMachine_Click
    End If
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub CreateDockPane()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Dim lngPane5Width As Long, lngPane2Height As Long, lngPane2Width As Long, lngPane3Height As Long
    
    
    dkpMain.Options.HideClient = True
    
    Set Pane1 = dkpMain.CreatePane(1, 200, 300, DockTopOf, Nothing)
    Pane1.Title = "�ϲ��걾��Ϣ"
    Pane1.Handle = Me.frmList.hWnd
    Pane1.Options = PaneNoCaption

    Set Pane2 = dkpMain.CreatePane(2, 200, 600, DockBottomOf, Nothing)
    Pane2.Title = "�ϲ��걾    (����ֱ���ڱ걾������걾��)"
    Pane2.Handle = PicList.hWnd
'    Pane2.Options = PaneNoCaption
    
    Set Pane3 = dkpMain.CreatePane(3, 400, 600, DockBottomOf, Pane2)
    Pane3.Title = "���ϲ��걾   (˫����ǰ�����걾���ϲ��걾��)"
    Pane3.Handle = PicList1.hWnd

    
    Pane1.Select
    
End Sub

Private Sub InitList()
'���ܣ���ʼ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
   
    strHead = "ID,0,4;" & _
        "�걾��,1500,1;������Ŀ,5000,1;������,3000,7;��־,800,1;������ĿId,0,7;���ϲ���ĿID,0,7"
    arrHead = Split(strHead, ";")
    
    With vfgList
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
    
    strHead = "ID,0,4;" & _
        "�걾��,1500,1;������Ŀ,5000,1;������Ŀid,0,7;������,3000,7;��־,800,1;������ĿId,0,7;���ϲ���ĿID,0,7"
    arrHead = Split(strHead, ";")
    With vfgList1
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub vfgList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = mCol.�걾�� Then
        Me.vfgList.Editable = flexEDKbdMouse
    Else
        Me.vfgList.Editable = flexEDNone
    End If
End Sub

Private Sub vfgList_DblClick()
    'Ҫ�ѵ�ǰ��ȡ���ı걾��ʾ����
    
    With Me.vfgList
        If .Row = 0 Or .Rows <= 1 Then Exit Sub
        If .TextMatrix(.Row, mCol.ID) = "" Then Exit Sub
        For intLoop = 1 To Me.vfgList1.Rows - 1
            If .TextMatrix(.Row, mCol.ID) = Me.vfgList1.TextMatrix(intLoop, mUCol.ID) Then
                Me.vfgList1.RowHidden(intLoop) = False
                .TextMatrix(.Row, mCol.ID) = ""
                .TextMatrix(.Row, mCol.�걾��) = ""
                .TextMatrix(.Row, mCol.������) = ""
                .TextMatrix(.Row, mCol.��־) = ""
                .TextMatrix(.Row, mCol.���ϲ���ĿID) = ""
                Exit For
            End If
        Next
    End With
End Sub

Private Sub vfgList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim blnFind As Boolean
    
    If Me.vfgList.Rows <= 0 Then Exit Sub
    If Me.vfgList1.Rows <= 0 Then Exit Sub
    Me.vfgList.TextMatrix(Row, mCol.�걾��) = Me.vfgList.EditText
    If Me.vfgList.TextMatrix(Row, mCol.�걾��) = "" Then
        With Me.vfgList
            For intRow = 1 To Me.vfgList1.Rows - 1
                If .TextMatrix(.Row, mCol.ID) = Me.vfgList1.TextMatrix(intRow, mCol.ID) Then
                    Me.vfgList1.RowHidden(intRow) = False
                End If
            Next
        End With
        Exit Sub
    End If
    
    With Me.vfgList
        For intRow = 1 To Me.vfgList.Rows - 1
            If .TextMatrix(.Row, mCol.�걾��) = .TextMatrix(intRow, mCol.�걾��) And .Row <> intRow Then
                .TextMatrix(intRow, mCol.ID) = ""
                .TextMatrix(intRow, mCol.�걾��) = ""
                .TextMatrix(intRow, mCol.������) = ""
                .TextMatrix(intRow, mCol.��־) = ""
                .TextMatrix(intRow, mCol.���ϲ���ĿID) = ""
            End If
        Next
    End With
    
    If Col = mCol.�걾�� Then
        For intLoop = 1 To Me.vfgList1.Rows - 1
            If Me.vfgList.TextMatrix(Row, mCol.�걾��) = Me.vfgList1.TextMatrix(intLoop, mUCol.�걾��) Then
                Me.vfgList1.RowHidden(intLoop) = True
                With Me.vfgList
                    If .TextMatrix(.Row, mCol.ID) <> "" And .TextMatrix(.Row, mCol.ID) <> Me.vfgList1.TextMatrix(intLoop, mUCol.ID) Then
                        For intRow = 1 To Me.vfgList1.Rows - 1
                            If .TextMatrix(.Row, mCol.ID) = Me.vfgList1.TextMatrix(intRow, mCol.ID) Then
                                Me.vfgList1.RowHidden(intRow) = False
                            End If
                        Next
                    End If
                    .TextMatrix(.Row, mCol.ID) = Me.vfgList1.TextMatrix(intLoop, mUCol.ID)
                    .TextMatrix(.Row, mCol.�걾��) = Me.vfgList1.TextMatrix(intLoop, mUCol.�걾��)
                    .TextMatrix(.Row, mCol.������) = Me.vfgList1.TextMatrix(intLoop, mUCol.������)
                    .TextMatrix(.Row, mCol.��־) = Me.vfgList1.TextMatrix(intLoop, mUCol.��־)
                    .TextMatrix(.Row, mCol.���ϲ���ĿID) = Me.vfgList1.TextMatrix(intLoop, mUCol.������Ŀid)
                    blnFind = True
                    Exit For
                End With
            End If
        Next
        If blnFind = flase Then
            With Me.vfgList
                For intRow = 1 To Me.vfgList1.Rows - 1
                    If .TextMatrix(.Row, mCol.ID) = Me.vfgList1.TextMatrix(intRow, mCol.ID) Then
                        Me.vfgList1.RowHidden(intRow) = False
                    End If
                Next
                .TextMatrix(.Row, mCol.ID) = ""
                .TextMatrix(.Row, mCol.�걾��) = ""
                .TextMatrix(.Row, mCol.������) = ""
                .TextMatrix(.Row, mCol.��־) = ""
                .TextMatrix(.Row, mCol.���ϲ���ĿID) = ""
                .EditText = ""
            End With
        End If
    End If
End Sub

Private Sub vfgList1_DblClick()
    Dim lngKey As Long
    Dim intLoop As Integer
    
    If Me.vfgList1.Row = 0 Then Exit Sub
    If Me.vfgList.Row = 0 Or Me.vfgList.Rows <= 1 Then
        MsgBox "��ѡ��һ����������Ŀ!"
        Exit Sub
    End If
    
    Me.vfgList1.RowHidden(Me.vfgList1.Row) = True
    
    With Me.vfgList
        If .TextMatrix(.Row, mCol.ID) <> Me.vfgList1.TextMatrix(Me.vfgList1.Row, mUCol.ID) And .TextMatrix(.Row, mCol.ID) <> "" Then
            'Ҫ�ѵ�ǰ��ȡ���ı걾��ʾ����
            For intLoop = 1 To Me.vfgList1.Rows - 1
                If .TextMatrix(.Row, mCol.ID) = Me.vfgList1.TextMatrix(intLoop, mUCol.ID) Then
                    Me.vfgList1.RowHidden(intLoop) = False
                    Exit For
                End If
            Next
        End If
        .TextMatrix(.Row, mCol.ID) = Me.vfgList1.TextMatrix(Me.vfgList1.Row, mUCol.ID)
        .TextMatrix(.Row, mCol.�걾��) = Me.vfgList1.TextMatrix(Me.vfgList1.Row, mUCol.�걾��)
        .TextMatrix(.Row, mCol.������) = Me.vfgList1.TextMatrix(Me.vfgList1.Row, mUCol.������)
        .TextMatrix(.Row, mCol.��־) = Me.vfgList1.TextMatrix(Me.vfgList1.Row, mUCol.��־)
        .TextMatrix(.Row, mCol.���ϲ���ĿID) = Me.vfgList1.TextMatrix(Me.vfgList1.Row, mUCol.������Ŀid)
    End With
    
'    If Me.vfgList.Row < Me.vfgList.Rows - 1 Then
'        Call Me.vfgList.Select(Me.vfgList.Row + 1, 0, Me.vfgList.Row + 1, Me.vfgList.Cols - 1)
'    Else
'        Call Me.vfgList.Select(1, 0, 1, Me.vfgList.Cols - 1)
'    End If
End Sub


