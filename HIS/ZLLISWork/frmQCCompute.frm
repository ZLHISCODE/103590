VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQCCompute 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ʧ�ؼ���"
   ClientHeight    =   8415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7710
   Icon            =   "frmQCCompute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo��Ŀ 
      Height          =   300
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   795
      Width           =   5000
   End
   Begin VB.Frame fraRule 
      Height          =   5190
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   7300
      Begin VB.CheckBox chk����� 
         Caption         =   "�����"
         Height          =   225
         Left            =   375
         TabIndex        =   18
         Top             =   225
         Visible         =   0   'False
         Width           =   6600
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "���ƹ���"
         Height          =   350
         Left            =   5940
         TabIndex        =   14
         Top             =   4650
         Width           =   1100
      End
      Begin VB.Frame fra2 
         Caption         =   "������ƽ��޹���"
         Height          =   1545
         Left            =   210
         TabIndex        =   13
         Top             =   2130
         Width           =   6900
         Begin VB.CheckBox chk�ؽ� 
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   20
            Top             =   255
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lbl��ʾ 
            Caption         =   "������ÿ��ˮƽ��>1ʱ����ѡ�������ƽ��޹���"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   1065
            TabIndex        =   22
            Top             =   630
            Visible         =   0   'False
            Width           =   4050
         End
      End
      Begin VB.Frame fra1 
         Caption         =   "�����ʿع���"
         Height          =   1545
         Left            =   225
         TabIndex        =   12
         Top             =   525
         Width           =   6900
         Begin VB.CheckBox chk���� 
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   19
            Top             =   255
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.Frame fra3 
         Caption         =   "�ۻ��͹���"
         Height          =   795
         Left            =   210
         TabIndex        =   11
         Top             =   3765
         Width           =   6900
         Begin VB.CheckBox chk�ۻ� 
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   300
            Visible         =   0   'False
            Width           =   1600
         End
      End
   End
   Begin VB.CheckBox chkALL 
      Caption         =   "���㱾����������Ŀ"
      Height          =   195
      Left            =   4290
      TabIndex        =   9
      Top             =   1485
      Width           =   1995
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -45
      TabIndex        =   8
      Top             =   345
      Width           =   11865
   End
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   350
      Left            =   6210
      TabIndex        =   1
      Top             =   495
      Width           =   1100
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "����(&E)"
      Height          =   350
      Left            =   6210
      TabIndex        =   0
      Top             =   930
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1275
      Left            =   480
      TabIndex        =   6
      Top             =   1725
      Width           =   6900
      _cx             =   12171
      _cy             =   2249
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
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
      Rows            =   6
      Cols            =   3
      FixedRows       =   0
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
      AutoResize      =   0   'False
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComCtl2.DTPicker dtp���� 
      Height          =   300
      Index           =   0
      Left            =   990
      TabIndex        =   15
      Top             =   1140
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   118161411
      CurrentDate     =   39110
   End
   Begin MSComCtl2.DTPicker dtp���� 
      Height          =   300
      Index           =   1
      Left            =   4295
      TabIndex        =   16
      Top             =   1140
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   118161411
      CurrentDate     =   39110
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmQCCompute.frx":058A
      Top             =   75
      Width           =   240
   End
   Begin VB.Label lbl�ʿ�Ʒ 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ��2����ͬŨ��ˮƽ���ʿ�Ʒ:"
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   1485
      Width           =   2700
   End
   Begin VB.Label lbl��Ŀ 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ: ####"
      Height          =   180
      Left            =   450
      TabIndex        =   5
      Top             =   825
      Width           =   900
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����������õ��ʿع����Զ�����ʧ�ؼ��㣬���ʧ��״̬��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   4
      Top             =   90
      Width           =   5040
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����: ####"
      Height          =   180
      Left            =   450
      TabIndex        =   3
      Top             =   1155
      Width           =   900
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����: ####"
      Height          =   180
      Left            =   450
      TabIndex        =   2
      Top             =   495
      Width           =   900
   End
End
Attribute VB_Name = "frmQCCompute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum mCol
    ID = 0: ѡ��: �ʿ�Ʒ: ˮƽ: ��ѡ
End Enum

Private mlngDevID As Long       '����id
Private mlngItemID As Long      '��Ŀid
Private mdtBeging As Date        '����
Private mdtEnd As Date          '��������
Private mintLevel As Integer    '�������������ʿ�ˮƽ��
Private mblnModify As Boolean   '�Ƿ�ִ��

Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Function ShowMe(frmParent As Form, lngDevId As Long, lngItemID As Long, dtBegin As Date, Optional lngResId As Long, Optional dtEnd As Date) As Boolean
    '���ܣ�����ָ����������Ŀ�����ڣ�����ʾ����Ի���
    Dim rsTemp As New adodb.Recordset
    
    mlngDevID = lngDevId
    mlngItemID = lngItemID
    mdtBeging = dtBegin
    If dtEnd = CDate(0) Then
        mdtEnd = dtBegin
    Else
        mdtEnd = dtEnd
    End If
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ���� || ', ' || ���� As ������, �ʿ�ˮƽ�� From �������� Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe", mlngDevID)
    With rsTemp
        If .RecordCount <= 0 Then MsgBox "ָ�����������ڣ�", vbInformation, gstrSysName: Exit Function
        mintLevel = !�ʿ�ˮƽ��
        Me.lbl����.Caption = "����: " & !������
    End With
    

    
    gstrSql = "Select ���� || ', ' || ������ || ', ' || Ӣ���� As ��Ŀ�� From ����������Ŀ Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
    With rsTemp
        If .RecordCount <= 0 Then MsgBox "ָ����Ŀ�����ڣ�", vbInformation, gstrSysName: Exit Function
        Me.lbl��Ŀ.Caption = "��Ŀ: " & !��Ŀ��
    End With
    Me.lbl����.Caption = "����: " & Format(dtBegin, "yyyy��mm��dd��")
    Me.lbl�ʿ�Ʒ.Caption = "��ѡ��" & mintLevel & "����ͬŨ��ˮƽ���ʿ�Ʒ:"
    
    Me.dtp����(0).Value = mdtBeging: Me.dtp����(0).MinDate = mdtBeging: Me.dtp����(0).MaxDate = mdtEnd
    Me.dtp����(1).Value = mdtEnd: Me.dtp����(1).MinDate = mdtBeging: Me.dtp����(1).MaxDate = mdtEnd
    
    gstrSql = "Select M.ID, '' As ѡ��, M.���� || ', ' || M.���� || ', ˮƽ:' || M.ˮƽ As �ʿ�Ʒ, M.ˮƽ, 0 As ��ѡ" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I" & vbNewLine & _
            "Where M.ID = I.�ʿ�Ʒid And M.����id=[1] And I.��Ŀid = [2] And " & vbNewLine & _
            "      To_Date([3],'yyyy-MM-dd') Between M.��ʼ���� And M.��������" & vbNewLine & _
            "Order By M.��ʼ����, M.ˮƽ"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDevId, lngItemID, Format(mdtBeging, "yyyy-MM-dd"))
    With Me.vfgList
        Set .DataSource = rsTemp
        .ColWidth(mCol.ѡ��) = 280
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.ˮƽ) = 0: .ColWidth(mCol.��ѡ) = 0
        .ColHidden(mCol.ID) = True: .ColHidden(mCol.ˮƽ) = True: .ColHidden(mCol.��ѡ) = True
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.ID) = lngResId Or lngCount < mintLevel Then
                .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked
            Else
                .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexUnchecked
            End If
        Next
    End With
    
    mblnModify = False
    Me.Show vbModal, frmParent
    ShowMe = mblnModify
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initRuleCtr()
    '��ʼ���ʿع���ѡ��ؼ�
    Dim rsTmp As adodb.Recordset
    Dim strSQL As String
    Dim lngLeft As Long, lngTop As Long
    
    On Error GoTo ErrHandle
    '��ˮƽ����ʹ�ÿ��ƽ��޹���
    lbl��ʾ.Visible = False
    If mintLevel > 1 Then
        strSQL = "Select B.ID, B.����, B.����, B.˵��, B.��ʽ, B.��ˮƽ From �����ʿع��� B Order By ����,��ʽ, B.���� "
    Else
        strSQL = "Select B.ID, B.����, B.����, B.˵��, B.��ʽ, B.��ˮƽ From �����ʿع��� B Where ���� In (1, 3)  Order By ����,��ʽ, B.���� "
        lbl��ʾ.Visible = True
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do Until rsTmp.EOF
        Select Case "" & rsTmp!����
        Case 1 '�����ʿع���
            
            If Trim(chk����(chk����.UBound).Tag) <> "" Then
                Load chk����(chk����.UBound + 1)
                lngLeft = chk����(chk����.UBound - 1).Left + chk����(chk����.UBound - 1).Width + 45
                lngTop = chk����(chk����.UBound - 1).Top
            Else
                lngLeft = chk����(chk����.UBound).Left
                lngTop = chk����(chk����.UBound).Top
            End If
            
            chk����(chk����.UBound).Caption = rsTmp!����
            chk����(chk����.UBound).Value = 0
            chk����(chk����.UBound).Tag = rsTmp!ID
            chk����(chk����.UBound).Visible = False
            If lngLeft + chk����(chk����.UBound).Width > Me.fra1.Left + Me.fra1.Width - 155 Then
                lngLeft = chk����(0).Left
                lngTop = chk����(chk����.UBound - 1).Top + chk����(chk����.UBound - 1).Height + 45
            End If
            chk����(chk����.UBound).Left = lngLeft
            chk����(chk����.UBound).Top = lngTop
            If Trim(chk����(chk����.UBound).Caption) <> "" Then chk����(chk����.UBound).Visible = True

        Case 2 '���ƽ��޹���
            
            If Trim(chk�ؽ�(chk�ؽ�.UBound).Tag) <> "" Then
                Load chk�ؽ�(chk�ؽ�.UBound + 1)
                lngLeft = chk�ؽ�(chk�ؽ�.UBound - 1).Left + chk�ؽ�(chk�ؽ�.UBound - 1).Width + 45
                lngTop = chk�ؽ�(chk�ؽ�.UBound - 1).Top
            Else
                lngLeft = chk�ؽ�(chk�ؽ�.UBound).Left
                lngTop = chk�ؽ�(chk�ؽ�.UBound).Top
            End If
            
            chk�ؽ�(chk�ؽ�.UBound).Caption = rsTmp!����
            chk�ؽ�(chk�ؽ�.UBound).Value = 0
            chk�ؽ�(chk�ؽ�.UBound).Tag = rsTmp!ID
            chk�ؽ�(chk�ؽ�.UBound).Visible = False
            If lngLeft + chk�ؽ�(chk�ؽ�.UBound).Width > Me.fra1.Left + Me.fra1.Width - 155 Then
                lngLeft = chk�ؽ�(0).Left
                lngTop = chk�ؽ�(chk�ؽ�.UBound - 1).Top + chk�ؽ�(chk�ؽ�.UBound - 1).Height + 45
            End If
            chk�ؽ�(chk�ؽ�.UBound).Left = lngLeft
            chk�ؽ�(chk�ؽ�.UBound).Top = lngTop
            If Trim(chk�ؽ�(chk�ؽ�.UBound).Caption) <> "" Then chk�ؽ�(chk�ؽ�.UBound).Visible = True
        
        Case Else   '�ۻ��͹���
            
            If Trim(chk�ۻ�(chk�ۻ�.UBound).Tag) <> "" Then
                Load chk�ۻ�(chk�ۻ�.UBound + 1)
                lngLeft = chk�ۻ�(chk�ۻ�.UBound - 1).Left + chk�ۻ�(chk�ۻ�.UBound - 1).Width + 45
                lngTop = chk�ۻ�(chk�ۻ�.UBound - 1).Top
            Else
                lngLeft = chk�ۻ�(chk�ۻ�.UBound).Left
                lngTop = chk�ۻ�(chk�ۻ�.UBound).Top
            End If
            
            chk�ۻ�(chk�ۻ�.UBound).Caption = rsTmp!����
            chk�ۻ�(chk�ۻ�.UBound).Value = 0
            chk�ۻ�(chk�ۻ�.UBound).Tag = rsTmp!ID
            chk�ۻ�(chk�ۻ�.UBound).Visible = False
            If lngLeft + chk�ۻ�(chk�ۻ�.UBound).Width > Me.fra1.Left + Me.fra1.Width - 155 Then
                lngLeft = chk�ۻ�(0).Left
                lngTop = chk�ۻ�(chk�ۻ�.UBound - 1).Top + chk�ۻ�(chk�ۻ�.UBound - 1).Height + 45
            End If
            chk�ۻ�(chk�ۻ�.UBound).Left = lngLeft
            chk�ۻ�(chk�ۻ�.UBound).Top = lngTop
            If Trim(chk�ۻ�(chk�ۻ�.UBound).Caption) <> "" Then chk�ۻ�(chk�ۻ�.UBound).Visible = True
        End Select
        
        rsTmp.MoveNext
    Loop
    
    '�����ؼ�
    strSQL = "Select Distinct I.ID, I.����, I.������, I.Ӣ����" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ Q, ����������Ŀ I" & vbNewLine & _
            "Where M.ID = Q.�ʿ�Ʒid And Q.��Ŀid = I.ID And M.����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDevID)
    With rsTmp
        Me.cbo��Ŀ.Clear
        Do Until .EOF
            Me.cbo��Ŀ.AddItem !���� & ", " & !������ & "/" & !Ӣ����
            Me.cbo��Ŀ.ItemData(Me.cbo��Ŀ.NewIndex) = !ID
            If !ID = mlngItemID Then
                Me.cbo��Ŀ.ListIndex = Me.cbo��Ŀ.NewIndex
            End If
            .MoveNext
        Loop
        If Me.cbo��Ŀ.ListCount = 0 Then MsgBox "��δ��������ʿ�Ʒ���ã�", vbInformation, gstrSysName: Unload Me: Exit Sub
        If cbo��Ŀ.ListIndex = -1 Then
            Me.cbo��Ŀ.ListIndex = 0
        End If

    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub refRuleStat()
    '��ʾ��ǰ��Ŀ�Ĺ���״̬
    Dim rsTmp As adodb.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    
    On Error GoTo ErrHandle
    
    '---- ��Ϊ��ʼ״̬
    chk�����.Value = 0
    For intCount = chk����.LBound To chk����.UBound
        chk����(intCount).Value = 0
    Next
    For intCount = chk�ؽ�.LBound To chk�ؽ�.UBound
        chk�ؽ�(intCount).Value = 0
    Next
    For intCount = chk�ۻ�.LBound To chk�ۻ�.UBound
        chk�ۻ�(intCount).Value = 0
    Next
    '-----
    mlngItemID = Me.cbo��Ŀ.ItemData(Me.cbo��Ŀ.ListIndex)
    
    strSQL = "Select A.ID, A.����id, A.����id, A.����, B.����, B.����, B.˵��, B.��ʽ, B.��ˮƽ, A.�Ƿ�ʹ��" & vbNewLine & _
            "From ������������ A, �����ʿع��� B" & vbNewLine & _
            "Where A.����id = B.ID And A.�ϼ�id Is Null And A.����id = [1] And A.��Ŀid = [2] " & vbNewLine & _
            "Order By A.����, B.����, A.����id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDevID, mlngItemID)
    chk�����.Visible = False
    Do Until rsTmp.EOF
        If Val("" & rsTmp!����) = 0 Then
            
            chk�����.Value = Val("" & rsTmp!�Ƿ�ʹ��)
            chk�����.Visible = True
        ElseIf Val("" & rsTmp!����) = 1 Then
            If Val("" & rsTmp!�Ƿ�ʹ��) = 1 Then
                Select Case Val("" & rsTmp!����)
                Case 1
                    For intCount = chk����.LBound To chk����.UBound
                        If Val(chk����(intCount).Tag) = Val("" & rsTmp!����id) Then
                            chk����(intCount).Value = 1
                            Exit For
                        End If
                    Next
                Case 2
                    For intCount = chk�ؽ�.LBound To chk�ؽ�.UBound
                        If Val(chk�ؽ�(intCount).Tag) = Val("" & rsTmp!����id) Then
                            chk�ؽ�(intCount).Value = 1
                            Exit For
                        End If
                    Next
                Case Else
                    For intCount = chk�ۻ�.LBound To chk�ۻ�.UBound
                        If Val(chk�ۻ�(intCount).Tag) = Val("" & rsTmp!����id) Then
                            chk�ۻ�(intCount).Value = 1
                            Exit For
                        End If
                    Next
                End Select
            End If
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CheckRule(ByVal lng����ID As Long, ByVal int�Ƿ�ʹ�� As Integer)
    '�˶Լ�����������ļ�¼�������������ӣ����޸����޸ġ�
    'ע��ֻ�ܶԸ��ӹ�����д���
    Dim strSQL As String
    Dim rsTmp As adodb.Recordset
    Dim blnHave  As Boolean
    On Error GoTo ErrHandle

    strSQL = "ZL_������������_SetUsed(" & mlngDevID & "," & mlngItemID & "," & lng����ID & "," & int�Ƿ�ʹ�� & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub cbo��Ŀ_Click()
    Call refRuleStat
End Sub


Private Sub chk����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lng����ID As Long
    If Button = 1 Then
        lng����ID = Val(Me.chk����(Index).Tag)
        If Me.chk����(Index).Value = 0 Then
            Call CheckRule(lng����ID, 1)
        Else
            Call CheckRule(lng����ID, 0)
        End If
    End If
End Sub

Private Sub chk�ؽ�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lng����ID As Long
    If Button = 1 Then
        lng����ID = Val(Me.chk�ؽ�(Index).Tag)
        If Me.chk�ؽ�(Index).Value = 0 Then
            Call CheckRule(lng����ID, 1)
        Else
            Call CheckRule(lng����ID, 0)
        End If
    End If
End Sub

Private Sub chk�ۻ�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lng����ID As Long
    If Button = 1 Then
        lng����ID = Val(Me.chk�ۻ�(Index).Tag)
        If Me.chk�ۻ�(Index).Value = 0 Then
            Call CheckRule(lng����ID, 1)
        Else
            Call CheckRule(lng����ID, 0)
        End If
    End If
End Sub

Private Sub cmdApply_Click()
    Call frmAppRuleCopy.ShowMe(mlngDevID, mlngItemID, Me)
End Sub

Private Sub cmdExecute_Click()
    Dim strResList As String, strLevels As String
    Dim rsTemp As New adodb.Recordset, rsTmp As New adodb.Recordset
    Dim strReturn As String
    Dim lngLoop As Long, lngDate As Date, lngCount As Long, strInfo As String
    
    Dim dtBeging As Date, dtEnd As Date
    
    strResList = "": strLevels = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked Then
                If InStr(1, strLevels, Trim(.TextMatrix(lngCount, mCol.ˮƽ))) > 0 Then
                    MsgBox "ֻ������һ��ˮƽ" & Trim(.TextMatrix(lngCount, mCol.ˮƽ)) & "���ʿ�Ʒ��", vbInformation, gstrSysName
                    Exit Sub
                End If
                strLevels = strLevels & "," & Trim(.TextMatrix(lngCount, mCol.ˮƽ))
                strResList = strResList & "," & Trim(.TextMatrix(lngCount, mCol.ID))
            End If
        Next
    End With
    If strResList <> "" Then strResList = Mid(strResList, 2)
'    2009-06-03 ���ã���ˮƽ����ʱ������Ҫÿ��ˮƽ�Ĳ��Ը���һ�£�ֻ��Ҫ�ܸ�������Ϳ��Լ���
'    If UBound(Split(strResList, ",")) <> mintLevel - 1 Then
'        MsgBox "�밴�����ʿ�Ҫ��ѡ��" & mintLevel & "����ͬˮƽ���ʿ�Ʒ��", vbInformation, gstrSysName
'        Exit Sub
'    End If
    
    Err = 0: On Error GoTo ErrHand
    dtBeging = dtp����(0).Value: dtEnd = dtp����(1).Value
    
    If Me.chkALL.Value = 1 Then
        
        gstrSql = "Select Distinct B.��Ŀid, C.����, C.������, C.Ӣ����" & vbNewLine & _
                    " From �����ʿ�Ʒ A, �����ʿ�Ʒ��Ŀ B, ����������Ŀ C" & vbNewLine & _
                    " Where A.ID = B.�ʿ�Ʒid And B.��Ŀid = C.ID And A.����id = [1] "
            
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevID)
        Do Until rsTmp.EOF
            '����һ��ʱ��
            lngCount = DateDiff("d", dtBeging, dtEnd)
            For lngLoop = 0 To lngCount
                gstrSql = "Select Zl_�����ʿؼ�¼_Compute(" & mlngDevID & ", " & rsTmp("��ĿID") & ", To_Date('" & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & strResList & "') From Dual"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)

                If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")  ������̵��ô���" & vbCrLf
                If InStr(rsTemp.Fields(0).Value, "����ʧ�أ�") > 0 Then
                    strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                    
                    ' 2009-06-03 ���ã���ǰ�����ʧ��ʱ�������㲻�ټ���,
                    Exit For
                ElseIf InStr(rsTemp.Fields(0).Value, "������ɣ�") <= 0 Then
                    If InStr(rsTemp.Fields(0).Value, "������δ���־����ʧ�أ�") <= 0 Then
                    strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                    End If
                End If
            Next
            rsTmp.MoveNext
        Loop
    Else
        lngCount = DateDiff("d", dtBeging, dtEnd)
        
        For lngLoop = 0 To lngCount
            gstrSql = "Select Zl_�����ʿؼ�¼_Compute(" & mlngDevID & ", " & mlngItemID & ", To_Date('" & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & strResList & "') From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(DateAdd("d", lngCount, dtBeging), "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")  ������̵��ô���" & vbCrLf
            If InStr(rsTemp.Fields(0).Value, "����ʧ�أ�") > 0 Then
                strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & rsTemp.Fields(0).Value & vbCrLf
                ' 2009-06-03 ���ã���ǰ�����ʧ��ʱ�������㲻�ټ���,
                Exit For
            ElseIf InStr(rsTemp.Fields(0).Value, "������ɣ�") <= 0 Then
                If InStr(rsTemp.Fields(0).Value, "������δ���־����ʧ�أ�") <= 0 Then
                strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & rsTemp.Fields(0).Value & vbCrLf
                End If
            End If
       Next
    End If
    If Trim(strReturn) = "" Then
        strReturn = "������ɣ�������δ���־����ʧ�أ�"
        MsgBox strReturn, vbInformation, gstrSysName
    Else
        Call frmQCShowInfo.ShowMe(Me.Caption, strReturn, Me)
    End If
    mblnModify = True
    If Left(strReturn, 4) = "�������" Then Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
    
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    chkALL.Value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����������Ŀ", 1)
    Call initRuleCtr
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����������Ŀ", chkALL.Value)

    For lngCount = 0 To chk����.Count - 1
        If lngCount > 0 Then Unload chk����(lngCount)
    Next
    For lngCount = 0 To chk�ؽ�.Count - 1
        If lngCount > 0 Then Unload chk�ؽ�(lngCount)
    Next
    
    For lngCount = 0 To chk�ۻ�.Count - 1
        If lngCount > 0 Then Unload chk�ۻ�(lngCount)
    Next
End Sub

Private Sub vfgList_DblClick()
    With Me.vfgList
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexUnchecked Or Val(.TextMatrix(.Row, mCol.��ѡ)) = 1 Then
            .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexChecked
        Else
            .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexUnchecked
        End If
    End With
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfgList_DblClick
End Sub
