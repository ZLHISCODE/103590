VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������ƻ�"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12030
   Icon            =   "frmBatch.frx":0000
   LinkTopic       =   "��������ƻ�"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkȫѡ 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "ȫѡ"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   7800
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6480
      Width           =   675
   End
   Begin VB.CommandButton cmdȫ������ 
      Caption         =   "ȫ������(&A)"
      Height          =   300
      Left            =   10680
      TabIndex        =   29
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frmline2 
      Height          =   120
      Left            =   120
      TabIndex        =   24
      Top             =   6120
      Width           =   11895
   End
   Begin VB.CommandButton cmd�������� 
      Caption         =   "��������(&B)"
      Height          =   300
      Left            =   9240
      TabIndex        =   22
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   1100
   End
   Begin VB.Frame Frmline1 
      Height          =   120
      Left            =   120
      TabIndex        =   14
      Top             =   620
      Width           =   11895
   End
   Begin VB.TextBox txt�ƻ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7440
      TabIndex        =   7
      Top             =   960
      Width           =   1710
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   9600
      TabIndex        =   10
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10800
      TabIndex        =   9
      Top             =   6360
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp����ʱ�� 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   4440
      TabIndex        =   11
      Top             =   960
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   124125187
      CurrentDate     =   36846.5833333333
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   4725
      Left            =   3480
      TabIndex        =   23
      Top             =   1350
      Width           =   8460
      _cx             =   14922
      _cy             =   8334
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBatch.frx":6852
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
      ExplorerBar     =   5
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
   Begin VB.Frame fraTemp 
      Caption         =   "��ȡ��������"
      Height          =   5235
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   3345
      Begin VB.TextBox txt��Ӧ�� 
         Height          =   300
         Left            =   1140
         TabIndex        =   1
         Top             =   480
         Width           =   1770
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   2085
      End
      Begin VB.OptionButton optClass 
         Caption         =   "ҩƷ(&1)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Value           =   -1  'True
         Width           =   1000
      End
      Begin VB.OptionButton optClass 
         Caption         =   "����(&2)"
         Height          =   180
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   2400
         Width           =   1000
      End
      Begin VB.CommandButton cmd��ȡ���� 
         Caption         =   "��ȡ����"
         Height          =   350
         Left            =   2125
         TabIndex        =   19
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CommandButton cmd��Ӧ�� 
         Caption         =   "��"
         Height          =   300
         Left            =   2880
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   465
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   1410
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   124125187
         CurrentDate     =   40848
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   1905
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   148504579
         CurrentDate     =   40848
      End
      Begin VB.Label lbl��Ӧ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "�� Ӧ ��"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label lbl��ʼ���� 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Label lblδ������� 
      AutoSize        =   -1  'True
      Caption         =   "30000000"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6120
      TabIndex        =   21
      Top             =   6480
      Width           =   840
   End
   Begin VB.Label lbl�ܽ���� 
      AutoSize        =   -1  'True
      Caption         =   "50000000"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      TabIndex        =   20
      Top             =   6480
      Width           =   840
   End
   Begin VB.Label lblInfor 
      Caption         =   "�����⹺��ⵥ����˵ĵ��ݣ������ƶ�����ƻ���"
      Height          =   285
      Left            =   810
      TabIndex        =   15
      Top             =   360
      Width           =   8535
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   255
      Picture         =   "frmBatch.frx":695D
      Top             =   100
      Width           =   480
   End
   Begin VB.Label lblδ����� 
      AutoSize        =   -1  'True
      Caption         =   "δ����"
      Height          =   180
      Left            =   5280
      TabIndex        =   13
      Top             =   6495
      Width           =   900
   End
   Begin VB.Label lbl�ƻ���� 
      AutoSize        =   -1  'True
      Caption         =   "�ƻ����"
      Height          =   180
      Left            =   6600
      TabIndex        =   8
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lbl�ܽ�� 
      AutoSize        =   -1  'True
      Caption         =   "�ܽ�"
      Height          =   180
      Left            =   3360
      TabIndex        =   0
      Top             =   6495
      Width           =   720
   End
   Begin VB.Label lbl����ʱ�� 
      BackColor       =   &H80000004&
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   3480
      TabIndex        =   12
      Top             =   1005
      Width           =   855
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlng��Ӧ��ID As Long
Private mint״̬ As Integer  '0--���ӣ�1--ɾ��
Private mstr��Ӧ�� As String
Private mblnOK As Boolean
Private Const mconlngColor As Long = &HFFFFFF        '�����޸�����ɫΪ��ɫ
Private Const mconlngCanColColor As Long = &HFFE3C8        '���޸�����ɫΪ����ɫ
Private Const mlngBorderColor As Long = &H0&    'ѡ���б߿���ɫ
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' ûѡ���б߿���ɫ

Public Sub ShowCard(ByVal frmMain As Object, ByVal strPrivs As String, Optional lng��Ӧ��ID As Long, Optional str��Ӧ�� As String, Optional int״̬ As Integer)

    mstrPrivs = strPrivs
    mstr��Ӧ�� = str��Ӧ��
    mlng��Ӧ��ID = lng��Ӧ��ID
    mint״̬ = int״̬
    mblnOK = False
    Me.Show vbModal, frmMain
End Sub

Private Function ValidData() As Boolean
    Dim i As Integer
    Dim dbl�ƻ���� As Double
    Dim blnɾ�� As Boolean
    
    If vsfList.Rows < 2 Then Exit Function
    With vsfList
        If mint״̬ = 0 Then
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("�ƻ����"))) > Val(.TextMatrix(i, .ColIndex("Ӧ�����"))) Then
                    MsgBox "��" & i & "�мƻ���������Ӧ�������飡", vbOKOnly + vbInformation, gstrSysName
                    .Row = i
                    .Col = .ColIndex("�ƻ����")
                    .TopRow = i
                    Exit Function
                End If
                
                dbl�ƻ���� = dbl�ƻ���� + Val(.TextMatrix(i, .ColIndex("�ƻ����")))
            Next
            
            If dbl�ƻ���� = 0 Then
                MsgBox "���мƻ���Ϊ0����ˣ����飡", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If mint״̬ = 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("�к�")) = "��" Then
                    blnɾ�� = True
                    Exit For
                End If
            Next
            
            If blnɾ�� = False Then
                MsgBox "���ȹ�ѡһ�����ݺ��ٱ��棡", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    ValidData = True
End Function

Private Sub chkȫѡ_Click()
    Dim i As Integer
    
    If vsfList.Rows < 2 Then chkȫѡ.Value = 0: Exit Sub
    With vsfList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�к�")) = "" And chkȫѡ.Value = 1 Then
                .TextMatrix(i, .ColIndex("�к�")) = "��"
            ElseIf .TextMatrix(i, .ColIndex("�к�")) = "��" And chkȫѡ.Value = 0 Then
                .TextMatrix(i, .ColIndex("�к�")) = ""
            End If
        Next
    End With
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim arrSql() As Variant     '��¼�洢���̵�����
    Dim blnTrans As Boolean
    Dim i As Integer
    Dim dbl�ƻ���� As Double
    
    On Error GoTo ErrHand:
    
    If ValidData = False Then Exit Sub

    If Format(dtp����ʱ��.Value, "yyyy-MM-dd") < Format(Sys.Currentdate, "yyyy-MM-dd") Then
        If MsgBox("�ƻ���������С���ƶ��ƻ����ڣ��Ƿ�ȷ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox(IIf(mint״̬ = 0, "�Ƿ�ȷ�����棿", "�Ƿ�ȷ��ɾ����"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
        
    arrSql = Array()
    
    For i = 1 To vsfList.Rows - 1
        If mint״̬ = 0 Then
            If Val(vsfList.TextMatrix(i, vsfList.ColIndex("�ƻ����"))) > 0 Then
                dbl�ƻ���� = Val(vsfList.TextMatrix(i, vsfList.ColIndex("�ƻ����")))
                
                gstrSQL = "Select ID, �ƻ����, ʣ��Ӧ�����" & vbNewLine & _
                                "From (Select a.Id, Max(Nvl(a.�ƻ����, 0)) As �ƻ����, Sum(Nvl(a.��Ʊ���, 0)) / Count(1) - Sum(Nvl(a.�ƻ����, 0)) As ʣ��Ӧ�����" & vbNewLine & _
                                "       From Ӧ����¼ A, ҩƷ�շ���¼ B" & vbNewLine & _
                                "       Where a.�շ�id = b.Id And b.���� = [2]  And b.No =[1] And b.������� Is Not Null" & vbNewLine & _
                                "       Group By a.Id" & vbNewLine & _
                                "       Order By a.Id)" & vbNewLine & _
                                "Where ʣ��Ӧ����� > 0"
    
                Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "������ȡҩƷ��ϸ", vsfList.TextMatrix(i, vsfList.ColIndex("���ݺ�")), IIf(optClass(0).Value = True, 1, 15))
        
                With rsTmp
                    Do While Not .EOF
                    
                        If dbl�ƻ���� > 0 Then
                            gstrSQL = "ZL_����ƻ�_INSERT("
                            
                            'ID_IN        IN Ӧ����¼.ID%Type,
                            gstrSQL = gstrSQL & !ID
                            
                            '�ƻ����_IN    IN Ӧ����¼.�ƻ����%Type,
                            gstrSQL = gstrSQL & "," & !�ƻ���� + 1
                            
                            '�ƻ����_IN    IN Ӧ����¼.�ƻ����%Type,
                            gstrSQL = gstrSQL & "," & IIf(dbl�ƻ���� > !ʣ��Ӧ�����, !ʣ��Ӧ�����, dbl�ƻ����)
                            
                            '�ƻ�����_IN    IN Ӧ����¼.�ƻ�����%Type,
                            gstrSQL = gstrSQL & "," & "TO_DATE('" & Format(dtp����ʱ��.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                            
                            '�ƻ���_IN    IN Ӧ����¼.�ƻ���%Type,
                            gstrSQL = gstrSQL & ",'" & UserInfo.���� & "'"
                            
                            '�ƶ�����_IN    IN Ӧ����¼.�ƶ�����%Type
                            gstrSQL = gstrSQL & "," & "TO_DATE('" & Format(Sys.Currentdate, "yyyy-MM-dd") & "','yyyy-MM-dd')"
                            
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = gstrSQL
                            
                            dbl�ƻ���� = dbl�ƻ���� - !ʣ��Ӧ�����
                        End If
                        
                        .MoveNext
                    Loop
                End With
            End If
        Else
            If vsfList.TextMatrix(i, vsfList.ColIndex("�к�")) = "��" Then
                gstrSQL = "Select a.id" & vbNewLine & _
                                "From Ӧ����¼ A, ҩƷ�շ���¼ B" & vbNewLine & _
                                "Where a.�շ�id = b.Id And b.���� = [3] And b.������� Is Not Null And a.��¼���� = -1 And a.������� Is Null and b.no=[1]" & vbNewLine & _
                                "And a.�ƻ����� =[2]"

                Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "������ȡҩƷ��ϸ", vsfList.TextMatrix(i, vsfList.ColIndex("���ݺ�")), _
                CDate(Format(vsfList.TextMatrix(i, vsfList.ColIndex("��������")), "yyyy-mm-dd")), IIf(optClass(0).Value = True, 1, 15))
        
                With rsTmp
                    Do While Not .EOF
                        gstrSQL = "ZL_����ƻ�_DELETE("
                        
                        'ID_IN        IN Ӧ����¼.ID%Type,
                        gstrSQL = gstrSQL & !ID
                        
                        '�ƻ�����_In    IN Ӧ����¼.��ʼ����_In,
                        gstrSQL = gstrSQL & "," & "TO_DATE('" & Format(vsfList.TextMatrix(i, vsfList.ColIndex("��������")), "yyyy-MM-dd") & "','yyyy-MM-dd')"
                        
                        gstrSQL = gstrSQL & ")"
                        
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = gstrSQL
                        
                        .MoveNext
                    Loop
                End With
            End If
        End If
    Next
                
    gcnOracle.BeginTrans: blnTrans = True          '��������
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '�ύ����
    
    mblnOK = True
    MsgBox IIf(mint״̬ = 0, "����ɹ���", "ɾ���ɹ���"), vbOKOnly + vbInformation, gstrSysName
    cmd��ȡ����_Click
    mblnOK = False
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd��������_Click()
    Dim i As Integer
    Dim dbl�ƻ���� As Double
    
    If vsfList.Rows < 2 Then
        Exit Sub
    End If
    
    If Val(txt�ƻ����.Text) > Val(lbl�ܽ����.Caption) Then
        MsgBox "�ƻ����������ܽ�� [" & Format(Val(lbl�ܽ����.Caption), "0.00") & "]�����������룡", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If

    If Val(txt�ƻ����.Text) = 0 Then
        MsgBox "�ƻ�����Ϊ�ջ�0�����飡", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    dbl�ƻ���� = Val(txt�ƻ����.Text)
    With vsfList
        For i = 1 To .Rows - 1
        
            If dbl�ƻ���� > Val(.TextMatrix(i, .ColIndex("Ӧ�����"))) Then
                .TextMatrix(i, .ColIndex("�ƻ����")) = Format(.TextMatrix(i, .ColIndex("Ӧ�����")), "0.00")
            Else
                .TextMatrix(i, .ColIndex("�ƻ����")) = Format(dbl�ƻ����, "0.00")
            End If

            dbl�ƻ���� = dbl�ƻ���� - Val(.TextMatrix(i, .ColIndex("�ƻ����")))
        Next
    End With
    
    
    lblδ�������.Caption = Format(Val(lbl�ܽ����.Caption) - Val(txt�ƻ����.Text), "0.00")
    
End Sub

Private Sub cmdȫ������_Click()
    txt�ƻ����.Text = lbl�ܽ����.Caption
    cmd��������_Click
End Sub

Private Sub cmd��ȡ����_Click()
    Dim rsTemp As ADODB.Recordset
    Dim dbl�ܽ�� As Double
    
    On Error GoTo ErrHand:
    
    vsfList.Rows = 1
    
    If mint״̬ = 0 Then
            gstrSQL = "Select NO, �������, ժҪ, Sum(Nvl(������, 0)) - Sum(Nvl(�ƻ����, 0)) As Ӧ�����" & vbNewLine & _
                            "From (Select b.��¼����, a.No, a.�������, a.ժҪ, Decode(��¼����, 0, b.��Ʊ���, 0) As ������, Decode(��¼����, -1, b.�ƻ����, 0) As �ƻ����" & vbNewLine & _
                            "       From ҩƷ�շ���¼ A, Ӧ����¼ B, " & IIf(optClass(0).Value = True, "ҩƷ��� D", "�������� D") & vbNewLine & _
                            "       Where a.Id = b.�շ�id And a.��ҩ��λid + 0 = [1] And" & vbNewLine & _
                            "             a.������� Between [2] And [3] And a.���� =[4] And a.ҩƷid = " & IIf(optClass(0).Value = True, "d.ҩƷid", "d.����id") & ")" & vbNewLine & _
                            "Group By NO, �������, ժҪ" & vbNewLine & _
                            "Having Sum(������) - Sum(�ƻ����) > 0" & vbNewLine & _
                            "Order By NO, �������"

            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "������ȡ����ƻ�", Val(txt��Ӧ��.Tag), _
            CDate(Format(dtp��ʼʱ��.Value, "yyyy-mm-dd")), CDate(Format(dtp����ʱ��.Value, "yyyy-mm-dd") & " 23:59:59"), _
             IIf(optClass(0).Value = True, 1, 15))
            
            If rsTemp.RecordCount = 0 And mblnOK = False Then
                MsgBox "δ��ѯ�����������ĵ��ݣ����飡", vbOKOnly + vbInformation, gstrSysName
            End If
            
            With rsTemp
                Do While Not .EOF
                    vsfList.Rows = vsfList.Rows + 1
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("�к�")) = vsfList.Rows - 1
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("���ݺ�")) = !No
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("��������")) = Format(!�������, "yyyy-mm-dd hh:mm:ss")
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("ժҪ")) = Nvl(!ժҪ)
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("�ƻ����")) = Format(0, "0.00")
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("Ӧ�����")) = Format(!Ӧ�����, "0.00")
                    
                    dbl�ܽ�� = dbl�ܽ�� + !Ӧ�����
                rsTemp.MoveNext
                Loop
            End With
            
            lbl�ܽ����.Caption = Format(dbl�ܽ��, "0.00")
            lblδ�������.Caption = Format(dbl�ܽ��, "0.00")
            txt�ƻ����.Text = Format(0, "0.00")
            
            Call setColEdit
        Else
            gstrSQL = "Select b.No, a.�ƻ�����, b.ժҪ, Sum(Nvl(a.�ƻ����, 0)) As �ƻ����, Sum(Nvl(a.��Ʊ���, 0)) As Ӧ�����" & vbNewLine & _
                            "From Ӧ����¼ A, ҩƷ�շ���¼ B, " & IIf(optClass(0).Value = True, "ҩƷ��� D", "�������� D") & vbNewLine & _
                            "Where a.�շ�id = b.Id And b.ҩƷid = " & IIf(optClass(0).Value = True, "d.ҩƷid", "d.����id") & " And b.��ҩ��λid + 0 = [1] And b.���� = [4] And a.��¼���� = -1" & vbNewLine & _
                            "      And a.������� Is Null And a.�ƻ����� Between [2] And [3] And b.������� Is Not Null" & vbNewLine & _
                            "Group By a.�ƻ�����, b.No, b.ժҪ" & vbNewLine & _
                            "Order By b.No, a.�ƻ�����"

            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "������ȡ����ƻ�", Val(txt��Ӧ��.Tag), _
            CDate(Format(dtp��ʼʱ��.Value, "yyyy-mm-dd")), CDate(Format(dtp����ʱ��.Value, "yyyy-mm-dd") & " 23:59:59"), _
            IIf(optClass(0).Value = True, 1, 15))
            
            If rsTemp.RecordCount = 0 And mblnOK = False Then
                MsgBox "δ��ѯ�����������ĵ��ݣ����飡", vbOKOnly + vbInformation, gstrSysName
            End If
            
            With rsTemp
                Do While Not .EOF
                    vsfList.Rows = vsfList.Rows + 1
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("���ݺ�")) = !No
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("��������")) = Format(!�ƻ�����, "yyyy-mm-dd")
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("ժҪ")) = Nvl(!ժҪ)
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("�ƻ����")) = Format(!�ƻ����, "0.00")
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.ColIndex("Ӧ�����")) = Format(!Ӧ�����, "0.00")
                    
                rsTemp.MoveNext
                Loop
            End With
            
            chkȫѡ.Value = 0
        End If
        
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call initComboBox
    Call Int��ʼ������
End Sub


Private Sub txt��Ӧ��_Change()
    With txt��Ӧ��
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub txt��Ӧ��_GotFocus()
    txt��Ӧ��.SelStart = 0
    txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strȨ�� As String
    
    vRect = zlControl.GetControlRect(txt��Ӧ��.hwnd)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errHandle
    With txt��Ӧ��
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        
        strȨ�� = " and " & Get����Ȩ��(mstrPrivs)

        gstrSQL = "" & _
            "  Select   ID,����,����,����,����" & _
            "  From  ��Ӧ�� " & _
            "  Where (����ʱ�� is null or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
            "       " & zl_��ȡվ������ & "  and ĩ��=1  " & _
            "       And ( ���� Like [1] or ���� like [1] or ����  like upper([1])) " & strȨ�� & _
            "      order by  ����  "
            
        Set adoProvider = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "��ҩ��λ", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", strProviderText & "%", gstrNodeNo)

        If blnCancel = True Then .SetFocus: Exit Sub  '��ѡ����ʱ����Esc�������´���
        
        If adoProvider.State = 0 Then
            MsgBox "û��������Ĺ�ҩ��λ�������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If

        .Text = "[" & adoProvider!���� & "]" & adoProvider!����
        .Tag = adoProvider!ID
        
        
        adoProvider.Close
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    strTemp = frm��Ӧ��ѡ��.SelDept(mstrPrivs)
    If strTemp = "" Then
        Unload frm��Ӧ��ѡ��
        If txt��Ӧ��.Enabled Then txt��Ӧ��.SetFocus
        Exit Sub
    End If
    txt��Ӧ��.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    txt��Ӧ��.Tag = Val(Left(strTemp, InStr(strTemp, ",") - 1))
    On Error GoTo errHandle
    Set rsTemp = zldatabase.OpenSQLRecord("select ���� from ��Ӧ�� where id=[1] ", Caption & "-��ȡ��Ӧ������", txt��Ӧ��.Tag)

    rsTemp.Close
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub initComboBox()
    With cbo�������
        .Clear
        .AddItem "����"
        If mint״̬ = 0 Then
            .AddItem "һ������"
            .AddItem "һ������"
            .AddItem "��������"
        Else
            .AddItem "δ��һ����"
            .AddItem "δ��һ����"
            .AddItem "δ��������"
        End If
        .AddItem "�Զ�������"
        .ListIndex = 0
    End With
End Sub

Private Sub cbo�������_Click()
    Dim dateCurrentDate As Date
    
    If cbo�������.Text = "�Զ�������" Then
        dtp��ʼʱ��.Enabled = True
        dtp����ʱ��.Enabled = True
        
    Else
        dtp��ʼʱ��.Enabled = False
        dtp����ʱ��.Enabled = False
    End If
    
    '����ѡ��ı�ʱ��
    dateCurrentDate = Sys.Currentdate
    If mint״̬ = 0 Then
        Select Case cbo�������.ListIndex
            Case 0
                dtp��ʼʱ��.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
                dtp����ʱ��.Value = dateCurrentDate
            Case 1
                dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
                dtp����ʱ��.Value = dateCurrentDate
            Case 2
                dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
                dtp����ʱ��.Value = dateCurrentDate
            Case 3
                dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
                dtp����ʱ��.Value = dateCurrentDate
        End Select
    Else
        Select Case cbo�������.ListIndex
            Case 0
                dtp��ʼʱ��.Value = dateCurrentDate
                dtp����ʱ��.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
            Case 1
                dtp��ʼʱ��.Value = dateCurrentDate
                dtp����ʱ��.Value = CDate(Format(DateAdd("d", 7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            Case 2
                dtp��ʼʱ��.Value = dateCurrentDate
                dtp����ʱ��.Value = CDate(Format(DateAdd("d", 30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            Case 3
                dtp��ʼʱ��.Value = dateCurrentDate
                dtp����ʱ��.Value = CDate(Format(DateAdd("d", 90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
        End Select
    End If
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��Ӧ��, KeyAscii, m�ı�ʽ
End Sub

Private Sub txt�ƻ����_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt�ƻ����, KeyAscii, m���ʽ
End Sub

Private Sub txt�ƻ����_LostFocus()
    txt�ƻ����.Text = Format(txt�ƻ����.Text, "0.00")
End Sub

Private Sub vsfList_DblClick()
    If mint״̬ = 1 Then
        With vsfList
            If .Rows < 2 Then Exit Sub
            .TextMatrix(.Row, .ColIndex("�к�")) = IIf(.TextMatrix(.Row, .ColIndex("�к�")) = "", "��", "")
        End With
    End If
End Sub

Private Sub vsfList_EnterCell()
    Dim i As Integer
    With vsfList
        If .Row = 0 Then Exit Sub
        .Editable = flexEDNone
        .FocusRect = flexFocusLight
        
        If mint״̬ = 0 Then
            If .Row > 0 And .Col = .ColIndex("�ƻ����") Then
                .Editable = flexEDKbdMouse
                .FocusRect = flexFocusSolid
            End If
    
            If .Rows > 1 Then
                For i = 1 To .Rows - 1
                    .CellBorderRange i, 0, i, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
                Next
                .CellBorderRange .Row, 0, .Row, .Cols - 1, mlngBorderColor, 0, 2, 0, 2, 0, 2
            End If
            
            vsfList.Col = vsfList.ColIndex("�ƻ����")
        End If
    End With
    
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then Exit Sub
    If vsfList.Row = 0 Then Exit Sub
    With vsfList
        If .Rows > 1 Then
            If MsgBox("�Ƿ�ȷ��ɾ����" & .Row & "�е��ݺ�Ϊ��" & .TextMatrix(.Row, .ColIndex("���ݺ�")) & "���ļ�¼��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            Else
                .RemoveItem .Row
            End If
        End If
    End With
    
    Call ������
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    With vsfList
        If KeyAscii = vbKeyBack Then Exit Sub
        Select Case .Col
            Case .ColIndex("�ƻ����")

                VsFlxGridCheckKeyPress vsfList, Row, Col, KeyAscii, m���ʽ

                strKey = .EditText
                If strKey = "" Then
                    strKey = .TextMatrix(.Row, .Col)
                End If
                
                If LenB(StrConv(.EditText, vbFromUnicode)) >= 16 Then
                    KeyAscii = 0
                End If
                                 
        End Select
    End With
End Sub


Private Sub ������()
    Dim dbl�ܽ�� As Double
    Dim dbl�ƻ���� As Double
    Dim i As Integer
    
    If vsfList.Rows < 2 Then Exit Sub
    With vsfList
        For i = 1 To .Rows - 1
            dbl�ܽ�� = dbl�ܽ�� + Val(.TextMatrix(i, .ColIndex("Ӧ�����")))
            dbl�ƻ���� = dbl�ƻ���� + Val(.TextMatrix(i, .ColIndex("�ƻ����")))
        Next
    End With
    
    lbl�ܽ����.Caption = Format(dbl�ܽ��, "0.00")
    lblδ�������.Caption = Format(dbl�ܽ�� - dbl�ƻ����, "0.00")
    txt�ƻ����.Text = Format(dbl�ƻ����, "0.00")
    
End Sub


Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    If vsfList.Rows < 2 Then Exit Sub
    strKey = vsfList.EditText
    With vsfList
         .EditText = Format(Val(strKey), "0.00")
         .TextMatrix(.Row, .ColIndex("�ƻ����")) = Format(Val(strKey), "0.00")
        If Val(.TextMatrix(.Row, .ColIndex("�ƻ����"))) > Val(.TextMatrix(.Row, .ColIndex("Ӧ�����"))) Then
            MsgBox "��" & .Row & "�мƻ���������Ӧ�������������룡", vbOKOnly + vbInformation, gstrSysName
        End If
    End With
    
    Call ������
End Sub

Private Sub setColEdit()
    Dim intRow As Integer

    With vsfList
        If .Rows < 2 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = mconlngColor
        .Cell(flexcpBackColor, 1, .ColIndex("�ƻ����"), .Rows - 1, .ColIndex("�ƻ����")) = mconlngCanColColor
    End With

End Sub

Private Sub Int��ʼ������()
    vsfList.Rows = 1
    txt��Ӧ��.Text = mstr��Ӧ��
    txt��Ӧ��.Tag = mlng��Ӧ��ID
    vsfList.AllowSelection = False '���ܶ�ѡ
    dtp����ʱ��.Value = Sys.Currentdate
    If mint״̬ = 0 Then
        lbl�ܽ����.Caption = Format(0, "0.00")
        lblδ�������.Caption = Format(0, "0.00")
        txt�ƻ����.Text = Format(0, "0.00")
        chkȫѡ.Visible = False
        frmBatch.Caption = "������������ƻ�"
    Else
        lbl�ܽ��.Visible = False
        lbl�ܽ����.Visible = False
        lblδ�����.Visible = False
        lblδ�������.Visible = False
        lbl����ʱ��.Visible = False
        dtp����ʱ��.Visible = False
        lbl�ƻ����.Visible = False
        txt�ƻ����.Visible = False
        cmd��������.Visible = False
        cmdȫ������.Visible = False
        lbl�������.Caption = "��������"
        lblInfor.Caption = "�����⹺��ⵥ����˵ĵ��ݣ�����ɾ������б���ѹ�ѡ�ĸ���ƻ���"
        chkȫѡ.Left = vsfList.Left
        Frmline1.Top = 580
        vsfList.Top = 950
        vsfList.Height = 5110
        chkȫѡ.Top = vsfList.Top - chkȫѡ.Height - 10
        vsfList.TextMatrix(0, vsfList.ColIndex("�к�")) = "���"
        vsfList.TextMatrix(0, vsfList.ColIndex("��������")) = "��������"
        
        vsfList.ColWidth(vsfList.ColIndex("���ݺ�")) = 1800
        vsfList.ColWidth(vsfList.ColIndex("��������")) = 2400
        
        vsfList.ColHidden(vsfList.ColIndex("Ӧ�����")) = True
        frmBatch.Caption = "����ɾ������ƻ�"
    End If
End Sub
