VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmPathSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "·��ѡ��"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8295
   Icon            =   "frmPathSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPath 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   0
      ScaleHeight     =   1920
      ScaleWidth      =   8295
      TabIndex        =   10
      Top             =   5700
      Width           =   8295
      Begin VB.Frame fraRule 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         TabIndex        =   28
         Top             =   1560
         Width           =   2415
         Begin VB.OptionButton optPrtRule 
            Caption         =   "�����ӡ"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optPrtRule 
            Caption         =   "���׶δ�ӡ"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   29
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkIsForwardSend 
         Caption         =   "������ǰ���������·����Ŀ"
         Height          =   180
         Left            =   4080
         TabIndex        =   27
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CheckBox chkIsEvaluate 
         Caption         =   "����ǰһ�첻���������ɽ����·����Ŀ"
         Height          =   180
         Left            =   4080
         TabIndex        =   26
         Top             =   840
         Width           =   3615
      End
      Begin VB.OptionButton optPrintDay 
         Caption         =   "3��"
         Height          =   180
         Index           =   1
         Left            =   2760
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton optPrintDay 
         Caption         =   "2��"
         Height          =   180
         Index           =   0
         Left            =   2160
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txt��ҩζ�� 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "30"
         Top             =   840
         Width           =   300
      End
      Begin VB.CheckBox chkδ����ʱ�������ҽ�������� 
         Caption         =   "δ����ʱ�������ҽ��������"
         Height          =   180
         Left            =   4080
         TabIndex        =   18
         Top             =   480
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   3000
         TabIndex        =   17
         Top             =   1740
         Width           =   300
      End
      Begin VB.TextBox txtSendAdvice 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   16
         Text            =   "1"
         Top             =   1560
         Width           =   300
      End
      Begin VB.CheckBox chkOutTable 
         Caption         =   "ҽ�������´�ҽ������·�����ϼ�¼"
         Height          =   180
         Left            =   4080
         TabIndex        =   15
         Top             =   120
         Width           =   3375
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2805
         TabIndex        =   14
         Top             =   1030
         Width           =   420
      End
      Begin VB.Frame fraPathExe 
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3135
         Begin VB.CheckBox chkInvocation 
            Caption         =   "����·��ִ�л���"
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   0
            Width           =   1815
         End
         Begin VB.CheckBox chkDoctor 
            Caption         =   "ҽ��"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox chkNurse 
            Caption         =   "��ʿ"
            Height          =   255
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Label lblPrtRule 
         Caption         =   "·������ӡ����"
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblPrintDays 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·����ÿҳ��ӡ������"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1980
      End
      Begin VB.Label lblSendAdvice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����·��ʱ��ҽ����������ǰʱ��    ��"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   3420
      End
      Begin VB.Label lbl��ҩζ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ�䷽�����޸ĵ���ҩζ������    %"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   3150
      End
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   9
      Top             =   5520
      Width           =   6975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7080
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8295
      TabIndex        =   6
      Top             =   0
      Width           =   8295
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����·����Ŀ����˳��"
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
         Left            =   1095
         TabIndex        =   8
         Top             =   120
         Width           =   2145
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    ·����Ŀ����ҽ��ʱȱʡ��·�����иý׶ζ���ķ��༰��Ŀ˳���г��������Ȱ��±�������˳�����У�ÿ������ҽ��ʱҲ���Ե���˳��"
         Height          =   360
         Left            =   1095
         TabIndex        =   7
         Top             =   360
         Width           =   6165
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathSetup.frx":038A
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   1
      Left            =   7680
      Picture         =   "frmPathSetup.frx":6504
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   0
      Left            =   7080
      Picture         =   "frmPathSetup.frx":69B5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7080
      TabIndex        =   2
      Top             =   1440
      Width           =   1100
   End
   Begin VB.PictureBox picAddRow 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   220
      Left            =   6000
      Picture         =   "frmPathSetup.frx":6E6E
      ScaleHeight     =   225
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   360
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6900
      _cx             =   12171
      _cy             =   8017
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSetup.frx":71F8
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Editable        =   2
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmPathSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mbytFun  As Byte     '0=�ٴ�·��ģ�����,1=ҽ��վ����
Private Enum CNAME
    c˳�� = 0
    c��Ч = 1
    c������� = 2
    c�������� = 3
    c��ҩ���� = 4
    c���� = 5
End Enum

Private Sub chkDoctor_Click()
    If chkNurse.Value = Unchecked And chkDoctor.Value = Unchecked Then
        chkInvocation.Value = Unchecked
    End If
End Sub

Private Sub chkInvocation_Click()
    
    If chkInvocation.Value = Checked Then
        chkDoctor.Enabled = True
        chkNurse.Enabled = True
        chkDoctor.Value = Checked
        chkNurse.Value = Checked
    Else
        chkDoctor.Enabled = False
        chkNurse.Enabled = False
        chkDoctor.Value = Unchecked
        chkNurse.Value = Unchecked
    End If
End Sub

Private Sub chkIsEvaluate_Click()
    If chkIsEvaluate.Value = Unchecked Then
        chkIsForwardSend.Value = Unchecked
        chkIsForwardSend.Enabled = False
    Else
        chkIsForwardSend.Enabled = True
    End If
End Sub

Private Sub chkNurse_Click()
    If chkNurse.Value = Unchecked And chkDoctor.Value = Unchecked Then
        chkInvocation.Value = Unchecked
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click(Index As Integer)
    With vsItem
            If Index = 0 And .Row > .FixedRows Then
                .RowPosition(.Row) = .Row - 1
                .TextMatrix(.Row, c˳��) = .TextMatrix(.Row, c˳��) + 1
                .TextMatrix(.Row - 1, c˳��) = .TextMatrix(.Row - 1, c˳��) - 1
                .Row = .Row - 1
            ElseIf Index = 1 And .Row < .Rows - 1 Then
                .RowPosition(.Row) = .Row + 1
                .TextMatrix(.Row, c˳��) = .TextMatrix(.Row, c˳��) - 1
                .TextMatrix(.Row + 1, c˳��) = .TextMatrix(.Row + 1, c˳��) + 1
                .Row = .Row + 1
            End If
    End With
End Sub

Private Sub cmdOK_Click()
    If Not (vsItem.Rows = 2 And vsItem.TextMatrix(1, CNAME.c��Ч) = "" And vsItem.TextMatrix(1, CNAME.c�������) = "" _
            And vsItem.TextMatrix(1, CNAME.c��������) = "" And vsItem.TextMatrix(1, CNAME.c��ҩ����) = "") Then
        If CheckData = False Then Exit Sub
    End If
    
    Call SaveData
    Unload Me
End Sub

Private Function CheckData() As Boolean
    Dim i As Long, str�������� As String, str��ҩ���� As String
    Dim rsSQL As ADODB.Recordset, strKey As String
    
    
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "�к�", adBigInt
    rsSQL.Fields.Append "ֵ", adVarChar, 200, adFldIsNullable
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    With vsItem
        .Cell(flexcpBackColor, .FixedRows, c˳��, .Rows - 1, .Cols - 1) = vbWhite
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, c��Ч) = "" Then
                MsgBox "��ѡ��ҽ����Ч��", vbInformation, gstrSysName
                .Select i, c��Ч
                Exit Function
            ElseIf .TextMatrix(i, c�������) = "" Then
                MsgBox "��ѡ��������Ŀ���", vbInformation, gstrSysName
                .Select i, c�������
                Exit Function
            ElseIf .TextMatrix(i, c��������) = "" Then
                If .Cell(flexcpData, i, c�������) = "H" Or .Cell(flexcpData, i, c�������) = "E" Then
                    MsgBox "��ѡ��������͡�", vbInformation, gstrSysName
                    .Select i, c��������
                    Exit Function
                End If
            ElseIf .TextMatrix(i, c��ҩ����) = "" Then
                If .TextMatrix(i, c�������) = "��ҩ�г�ҩ" Then
                    MsgBox "��ѡ���ҩ���ࡣ", vbInformation, gstrSysName
                    .Select i, c��ҩ����
                    Exit Function
                End If
            End If
        
    
            '����ظ�ֵ
            If .TextMatrix(i, c��������) = "" Then
                str�������� = "Null"
            Else
                str�������� = .Cell(flexcpData, i, c��������)
            End If
            
            If .TextMatrix(i, c��ҩ����) = "" Then
                str��ҩ���� = "Null"
            Else
                str��ҩ���� = .Cell(flexcpData, i, c��ҩ����)
            End If
            strKey = .Cell(flexcpData, i, c��Ч) & "," & .Cell(flexcpData, i, c�������) & "," & str�������� & "," & str��ҩ����
            
            rsSQL.Filter = "ֵ='" & strKey & "'"
            If rsSQL.RecordCount > 0 Then
                MsgBox "��" & i & "�����" & rsSQL!�к� & "�е������ظ���", vbInformation, gstrSysName
                .Cell(flexcpBackColor, Val(rsSQL!�к�), c˳��, Val(rsSQL!�к�), .Cols - 1) = &H80C0FF
                .Select i, c��Ч
                Exit Function
            Else
                rsSQL.AddNew
                rsSQL!�к� = i
                rsSQL!ֵ = strKey
                rsSQL.Update
            End If
        Next
    End With
    CheckData = True
End Function

Private Sub SaveData()
    Dim strSql As String
    Dim i As Long, str�������� As String, str��ҩ���� As String
    Dim colSQL As New Collection, blnTrans As Boolean, blnSetup As Boolean
    Dim intOnlyDel As Integer
    Dim strTmp As String
    
    On Error GoTo errH
    With vsItem
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, c��������) = "" Then
                str�������� = "Null"
            Else
                str�������� = "'" & .Cell(flexcpData, i, c��������) & "'"
            End If
            
            If .TextMatrix(i, c��ҩ����) = "" Then
                str��ҩ���� = "Null"
            Else
                str��ҩ���� = .Cell(flexcpData, i, c��ҩ����)
            End If

            If vsItem.Rows = 2 And vsItem.TextMatrix(1, CNAME.c��Ч) = "" And vsItem.TextMatrix(1, CNAME.c�������) = "" _
                    And vsItem.TextMatrix(1, CNAME.c��������) = "" And vsItem.TextMatrix(1, CNAME.c��ҩ����) = "" Then
                intOnlyDel = 1
            Else
                intOnlyDel = 0
            End If
            strSql = "Zl_·����Ŀ˳��_Insert(" & .TextMatrix(i, c˳��) & "," & _
                IIf(.Cell(flexcpData, i, c��Ч) = "", "null", .Cell(flexcpData, i, c��Ч)) & _
                ",'" & .Cell(flexcpData, i, c�������) & "'," & str�������� & "," & str��ҩ���� & "," & _
                intOnlyDel & ")"
            colSQL.Add strSql, "C" & colSQL.count + 1
        Next
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    '�����������
    If mbytFun = 0 Then
        blnSetup = InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";��������;") > 0
        Call zlDatabase.SetPara("�Ƿ�����·��ִ�л���", chkInvocation.Value, glngSys, p�ٴ�·��Ӧ��, blnSetup)
        Call zlDatabase.SetPara("·��ִ�л������ó���", chkDoctor.Value & chkNurse.Value, glngSys, p�ٴ�·��Ӧ��, blnSetup)
      
        Call zlDatabase.SetPara("ҽ��ҽ����·������", chkOutTable.Value, glngSys, p�ٴ�·��Ӧ��, blnSetup)

        Call zlDatabase.SetPara("·��ҽ�����ɳ�ǰ����", txtSendAdvice.Text, glngSys, p�ٴ�·��Ӧ��, blnSetup)
        Call zlDatabase.SetPara("��ҩ�䷽�����޸ĵ���ҩζ������", txt��ҩζ��.Text, glngSys, p�ٴ�·��Ӧ��, blnSetup)
        i = IIf(optPrintDay(0).Value, Val(optPrintDay(0).Caption), Val(optPrintDay(1).Caption))
        Call zlDatabase.SetPara("·����ÿҳ��ӡ������", i & "", glngSys, p�ٴ�·��Ӧ��, blnSetup)
        Call zlDatabase.SetPara("δ����ʱ�������ҽ��������", chkδ����ʱ�������ҽ��������.Value, glngSys, p�ٴ�·��Ӧ��, blnSetup)
        Call zlDatabase.SetPara("����ǰһ�첻���������ɽ����·����Ŀ", chkIsEvaluate.Value, glngSys, p�ٴ�·��Ӧ��, blnSetup)
        Call zlDatabase.SetPara("������ǰ���������·����Ŀ", chkIsForwardSend.Value, glngSys, p�ٴ�·��Ӧ��, blnSetup)
        i = IIf(optPrtRule(0).Value, 0, 1)
        Call zlDatabase.SetPara("·������ӡ����", i & "", glngSys, p�ٴ�·��Ӧ��, blnSetup)
    End If
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim blnParSet As Boolean
    Dim lngDays As Long
    Dim strTmp As String
    
    If mbytFun = 0 Then
        Me.Caption = "·��ѡ��"
        lblInfo.Caption = "ҽ����·����Ŀ����˳��"
        lblNote.Caption = "    ·����Ŀ����ҽ��ʱȱʡ��·�����иý׶ζ���ķ��༰��Ŀ˳���г��������Ȱ��±�������˳�����У�ÿ������ҽ��ʱҲ���Ե���˳��"
    Else
        Me.Caption = "ҽ����������"
        lblInfo.Caption = "ҽ���´���Զ�����"
        lblNote.Caption = "    ҽ������ǰ���Ա����¿���ҽ�����Զ����±�������˳�����У������Ҳ����ʹ��ҽ��˳�����������������˳��"
        Frame1.Visible = False
        picPath.Visible = False
        vsItem.Height = vsItem.Height + 1200
    End If
    
    picAddRow.Visible = False
    Call InitData
    Call LoadData
    
    If mbytFun = 0 Then
        blnParSet = InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";��������;")

        chkInvocation.Value = zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, p�ٴ�·��Ӧ��, "1", Array(chkInvocation), blnParSet)
        If chkInvocation.Value = Checked Then
            strTmp = zlDatabase.GetPara("·��ִ�л������ó���", glngSys, p�ٴ�·��Ӧ��, "11", Array(chkDoctor, chkNurse), blnParSet)
            chkDoctor.Value = Val(Mid(strTmp, 1, 1))
            chkNurse.Value = Val(Mid(strTmp, 2, 1))
        Else
            chkDoctor.Enabled = False
            chkNurse.Enabled = False
        End If
        chkOutTable.Value = zlDatabase.GetPara("ҽ��ҽ����·������", glngSys, p�ٴ�·��Ӧ��, "0", Array(chkOutTable), blnParSet)
        lngDays = Val(zlDatabase.GetPara("·����ÿҳ��ӡ������", glngSys, p�ٴ�·��Ӧ��, "2", Array(lblPrintDays, optPrintDay(0), optPrintDay(1)), blnParSet))
        If lngDays = 2 Then
            optPrintDay(0).Value = True
        Else
            optPrintDay(1).Value = True
        End If
        strTmp = zlDatabase.GetPara("·������ӡ����", glngSys, p�ٴ�·��Ӧ��, "0", Array(lblPrtRule, optPrtRule(0), optPrtRule(1)), blnParSet)
        optPrtRule(Val(strTmp)).Value = True
        txtSendAdvice.Text = zlDatabase.GetPara("·��ҽ�����ɳ�ǰ����", glngSys, p�ٴ�·��Ӧ��, "1", Array(lblSendAdvice, txtSendAdvice), blnParSet)
        chkIsEvaluate.Value = zlDatabase.GetPara("����ǰһ�첻���������ɽ����·����Ŀ", glngSys, p�ٴ�·��Ӧ��, "1", Array(chkIsEvaluate), blnParSet)
        chkIsForwardSend.Value = zlDatabase.GetPara("������ǰ���������·����Ŀ", glngSys, p�ٴ�·��Ӧ��, "0", Array(chkIsForwardSend), blnParSet)
        chkδ����ʱ�������ҽ��������.Value = zlDatabase.GetPara("δ����ʱ�������ҽ��������", glngSys, p�ٴ�·��Ӧ��, "1", Array(chkδ����ʱ�������ҽ��������), blnParSet)
        txt��ҩζ��.Text = zlDatabase.GetPara("��ҩ�䷽�����޸ĵ���ҩζ������", glngSys, p�ٴ�·��Ӧ��, "30", Array(lbl��ҩζ��, txt��ҩζ��), blnParSet)
        txt��ҩζ��.Tag = txt��ҩζ��.Text
        If chkIsEvaluate.Value = Unchecked Then
            chkIsForwardSend.Value = Unchecked
            chkIsForwardSend.Enabled = False
        End If
    End If
    '�������Ƶ���һ��
    If vsItem.Rows > 0 Then vsItem.Row = 1
End Sub

Private Sub LoadData()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    
    On Error GoTo errH
    strSql = "Select a.˳��,a.ҽ����Ч,a.������� as ������,a.ִ�з���,a.��������,b.���� as ������� From ·����Ŀ˳�� a,������Ŀ��� b " & _
        "Where a.������� = b.���� Order by a.˳��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "·����Ŀ˳��")
    
    If rsTmp.RecordCount = 0 Then vsItem.TextMatrix(1, c˳��) = 1: Exit Sub
    
    With vsItem
        .Redraw = False
        .Rows = .FixedRows + rsTmp.RecordCount
        i = .FixedRows
        While Not rsTmp.EOF
            .TextMatrix(i, c˳��) = i
            .TextMatrix(i, c��Ч) = IIf(rsTmp!ҽ����Ч = 0, "����", "����")
            .Cell(flexcpData, i, c��Ч) = Val(rsTmp!ҽ����Ч)
            
            .Cell(flexcpData, i, c�������) = CStr(rsTmp!������)  'ҩƷ���Ǵ�Ϊ������ı���E
            
            If rsTmp!������ = "E" And Val("" & rsTmp!��������) = 2 Then
                .TextMatrix(i, c�������) = "��ҩ�г�ҩ"
                
            ElseIf rsTmp!������ = "E" And Val("" & rsTmp!��������) = 4 Then
                .TextMatrix(i, c�������) = "�в�ҩ"
                
            Else
                .TextMatrix(i, c�������) = rsTmp!�������
            End If
            
            '��ֻ֧�֣������ࣺ0-��ͨ;1-��������;2-��ҩ����(��ҩ);3-��ҩ�巨;4-��ҩ��(��)��;5-��������;6-�ɼ�����(����);7-��Ѫ����(Ѫ��);8-��Ѫ;����
            '            �����ࣺ0-�����棻1-����ȼ���
            If Not IsNull(rsTmp!��������) And (rsTmp!������ = "H" Or rsTmp!������ = "E") Then
                If rsTmp!������ = "H" Then
                    .TextMatrix(i, c��������) = IIf(rsTmp!�������� = 0, "������", "����ȼ�")
                Else
                     .TextMatrix(i, c��������) = Choose(Val(rsTmp!��������) + 1, "��ͨ", "��������", "��ҩ����", "��ҩ�巨", "��ҩ�÷�", "��������", "�ɼ�����", "��Ѫ����", "��Ѫ;��")
                End If
                .Cell(flexcpData, i, c��������) = Val(rsTmp!��������)
            End If
            
            If Not IsNull(rsTmp!ִ�з���) Then
                .TextMatrix(i, c��ҩ����) = Choose(rsTmp!ִ�з��� + 1, "����", "��Һ", "ע��", "Ƥ��", "�ڷ�")
                .Cell(flexcpData, i, c��ҩ����) = Val("" & rsTmp!ִ�з���)
            End If
            
            i = i + 1
            rsTmp.MoveNext
        Wend
        .Redraw = True
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitData()
    Dim rsTmp As ADODB.Recordset, strTmp As String
    
    Set rsTmp = Get�������
    strTmp = "#E;��ҩ�г�ҩ|#E;�в�ҩ"  '�̶����ƣ���������洢
    While Not rsTmp.EOF
        strTmp = strTmp & "|#" & rsTmp!���� & ";" & rsTmp!����
        rsTmp.MoveNext
    Wend
    
    With vsItem
        
        .ColComboList(c��Ч) = "#1;����|#0;����"
        .ColComboList(c�������) = strTmp
        .ColComboList(c��ҩ����) = "#0;����|#1;��Һ|#2;ע��|#3;Ƥ��|#4;�ڷ�"
        .Rows = .FixedRows
        .Rows = .FixedRows + 1 '��ʼһ����
    End With
End Sub

Private Function Get�������() As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ����,���� From ������Ŀ��� Where ���� Not In('5','6','7')"
    Set Get������� = zlDatabase.OpenSQLRecord(strSql, "�������")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub picAddRow_Click()
    Dim i As Long
    
    If vsItem.Row = vsItem.Rows - 1 Then
        vsItem.Rows = vsItem.Rows + 1
        vsItem.TextMatrix(vsItem.Rows - 1, c˳��) = vsItem.Rows - 1
        vsItem.Select vsItem.Rows - 1, c��Ч
    Else
        i = vsItem.Row
        vsItem.AddItem "", i
        Call Reset���
        vsItem.Select i, c��Ч
    End If
    
End Sub

Private Sub Reset���()
    Dim i As Long
    
    For i = vsItem.FixedRows To vsItem.Rows - 1
        vsItem.TextMatrix(i, c˳��) = i
    Next
End Sub

Private Sub txtSendAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txt��ҩζ��_GotFocus()
    zlControl.TxtSelAll txt��ҩζ��
End Sub

Private Sub txt��ҩζ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txt��ҩζ��_Validate(Cancel As Boolean)
    If Val(txt��ҩζ��.Text) > 100 Then
        MsgBox "��ҩ�䷽�����޸ĵ���ҩζ�����ޱ���ֻ����0-100֮������֡�", vbInformation, Me.Caption
        Cancel = True
        txt��ҩζ��.Text = txt��ҩζ��.Tag
        zlControl.TxtSelAll txt��ҩζ��
    Else
        txt��ҩζ��.Tag = txt��ҩζ��.Text
    End If
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsItem.ComboData = "" Then   'δѡ��ʱ�뿪����
        vsItem.TextMatrix(Row, Col) = CStr(vsItem.Tag)
        Exit Sub
    End If
    
    With vsItem
        If .Tag <> "" Then
            If .Tag = CStr(.ComboItem) Then Exit Sub
        End If
        .TextMatrix(Row, Col) = .ComboItem
        .Cell(flexcpData, Row, Col) = .ComboData
        
        If Col = c������� Then
            If .TextMatrix(Row, c�������) = "��ҩ�г�ҩ" Then
                .TextMatrix(Row, c��������) = "��ҩ����"
                .Cell(flexcpData, Row, c��������) = 2
            
            ElseIf .TextMatrix(Row, c�������) = "�в�ҩ" Then
                .TextMatrix(Row, c��������) = "��ҩ�÷�"
                .Cell(flexcpData, Row, c��������) = 4
                
            Else
                .TextMatrix(Row, c��������) = ""
                .Cell(flexcpData, Row, c��������) = ""
            End If
            
            .TextMatrix(Row, c��ҩ����) = ""
            .Cell(flexcpData, Row, c��ҩ����) = ""
            
        ElseIf Col = c�������� Then
            .TextMatrix(Row, c��ҩ����) = ""
            .Cell(flexcpData, Row, c��ҩ����) = ""
        End If
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If (OldRow <> NewRow Or OldRow = NewRow And OldRow = 1) And NewRow > vsItem.FixedRows - 1 Then
        If Me.Visible Then
            If picAddRow.Visible = False Then picAddRow.Visible = True
        End If
        picAddRow.Top = vsItem.Top + vsItem.Cell(flexcpTop, NewRow, c����) + 30
        picAddRow.Left = vsItem.Left + vsItem.Cell(flexcpLeft, NewRow, c����) + 120
    End If
End Sub

Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '��ֻ֧�֣������ࣺ0-��ͨ;1-��������;2-��ҩ����(��ҩ);3-��ҩ�巨;4-��ҩ��(��)��;5-��������;6-�ɼ�����(����);7-��Ѫ����(Ѫ��);8-��Ѫ;����
            '            �����ࣺ0-�����棻1-����ȼ���
    With vsItem
        .Tag = .TextMatrix(Row, Col)  '����AfterEdit���ж��Ƿ�ı���ֵ
        If Col = c���� Then
            Cancel = True
        ElseIf Col = c�������� Then '���ƺͻ��������
            If .Cell(flexcpData, Row, c�������) = "H" Then
                .ComboList = "#0;������|#1;����ȼ�"
                
            ElseIf .TextMatrix(Row, c�������) = "��ҩ�г�ҩ" Or .TextMatrix(Row, c�������) = "�в�ҩ" Then
                .ComboList = ""
                Cancel = True
            
            ElseIf .Cell(flexcpData, Row, c�������) = "E" Then
                .ComboList = "#0;��ͨ|#1;��������|#5;��������"
                
            Else
                .ComboList = ""
                Cancel = True
            End If
        ElseIf Col = c��ҩ���� Then 'ҩƷ������
            If Not (.TextMatrix(Row, c�������) = "��ҩ�г�ҩ") Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub vsItem_ChangeEdit()
    Call vsItem_AfterEdit(vsItem.Row, vsItem.Col)
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        If vsItem.Row = 0 Then Exit Sub
        If MsgBox("Ҫɾ����ǰ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
        With vsItem
        
            If .Rows > 2 Then
                vsItem.RemoveItem vsItem.Row
                Call Reset���
            ElseIf .Rows = 2 Then
                If .TextMatrix(1, CNAME.c��Ч) = "" And .TextMatrix(1, CNAME.c�������) = "" _
                        And .TextMatrix(1, CNAME.c��������) = "" And .TextMatrix(1, CNAME.c��ҩ����) = "" Then
                    MsgBox "û�п�ɾ�������ˡ�", vbInformation, gstrSysName
                Else
                    .TextMatrix(1, CNAME.c��Ч) = ""
                    .TextMatrix(1, CNAME.c�������) = ""
                    .TextMatrix(1, CNAME.c��������) = ""
                    .TextMatrix(1, CNAME.c��ҩ����) = ""
                End If
            End If
        
        End With
       
    End If
End Sub

Private Sub EnterNextCell()
   
    With vsItem
        If .Col = .Cols - 1 And .Row < .Rows - 1 Then
            .Select .Row + 1, c��Ч
        ElseIf .Col < .Cols - 1 Then
            .Col = .Col + 1
        End If
    End With
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call EnterNextCell
    End If
End Sub

Private Sub vsItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsItem.ComboIndex <> -1 Then
            Call vsItem_KeyPress(13)
        End If
    End If
End Sub
