VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeTurnNew 
   AutoRedraw      =   -1  'True
   Caption         =   "��(��)�����תסԺ"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   12465
   Icon            =   "frmChargeTurnNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   12465
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBill 
      BorderStyle     =   0  'None
      Height          =   4710
      Left            =   60
      ScaleHeight     =   4710
      ScaleWidth      =   11355
      TabIndex        =   7
      Top             =   660
      Width           =   11355
      Begin VSFlex8Ctl.VSFlexGrid vsfBill 
         Height          =   4440
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   11040
         _cx             =   19473
         _cy             =   7832
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12465
      TabIndex        =   5
      Top             =   0
      Width           =   12465
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ�ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   4
         Left            =   4620
         TabIndex        =   6
         Top             =   180
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2145"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   14
         Left            =   5430
         TabIndex        =   16
         Top             =   180
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   13
         Left            =   3750
         TabIndex        =   14
         Top             =   180
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   3
         Left            =   3150
         TabIndex        =   13
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   12
         Left            =   2340
         TabIndex        =   12
         Top             =   180
         Width           =   210
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   2
         Left            =   1740
         TabIndex        =   11
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����С"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   11
         Left            =   840
         TabIndex        =   10
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ˣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   630
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12465
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7515
      Width           =   12465
      Begin VB.TextBox txtSum 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2580
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   138
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9660
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   8490
         TabIndex        =   0
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   210
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "ת���ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1500
         TabIndex        =   15
         Top             =   190
         Width           =   1020
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8070
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeTurnNew.frx":058A
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
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
End
Attribute VB_Name = "frmChargeTurnNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlng�Һ�ID As Long
Private mbln����ִ�� As Boolean
Private mcllExecSql As Collection
Private mstr�˵�ID As String 'JSON��
Private mblnOk As Boolean

Private Enum idx_Lable
    lblName = 1
    txtName = 11
    lblSex = 2
    txtSex = 12
    lblAge = 3
    txtAge = 13
    lblInNumber = 4
    txtInNumber = 14
End Enum

Private mrsPerson As ADODB.Recordset '��Ա
Private mrsDepartment As ADODB.Recordset '����
Private mrsChargeitem As ADODB.Recordset '�շ���Ŀ

Private mrsPatient As ADODB.Recordset '������Ϣ
Private mrsFeeBill  As ADODB.Recordset '������Ϣ

Public Function ShowMe(objParent As Object, ByVal lng�Һ�ID As Long, _
    Optional ByVal bln����ִ�� As Boolean = True, _
    Optional ByRef cllExecSql As Collection, Optional ByRef str�˵�ID As String) As Boolean
    '����:�����ﲡ���������תסԺ����
    '���:
    '   lng�Һ�ID - �Һ�ID
    '   bln����ִ�� - �Ƿ����ִ�У�����Ƕ���ִ������ύ���ݵ����ݿ⣬������cllSql����ִ��SQL
    '����:
    '   cllExecSql bln����ִ��ΪFalseʱ������ִ��SQL
    '   str�˵�ID �ɹ������������תסԺ���������˵�ID,JSON��ʽ
    '����:�ɹ�����True,���򷵻�False
    mlng�Һ�ID = lng�Һ�ID
    mbln����ִ�� = bln����ִ��
    Set cllExecSql = Nothing: mstr�˵�ID = ""
    mblnOk = False
    
    On Error Resume Next
    Me.Show vbModal, objParent
    ShowMe = mblnOk
    If Not bln����ִ�� Then
        Set cllExecSql = mcllExecSql
        str�˵�ID = mstr�˵�ID
    End If
End Function

Private Sub Form_Load()
    Dim strData As String
    
    zlCommFun.ShowFlash "���ڻ�ȡ��ת�����������б����Ժ�...", Me
    If GetBillData(mlng�Һ�ID, strData) = False Then GoTo ErrExit:
    If InitData(mlng�Һ�ID) = False Then GoTo ErrExit:
    If AnalyzeData(strData, mrsFeeBill) = False Then GoTo ErrExit:
    If InitFace() = False Then GoTo ErrExit:
    If ShowBills(mrsFeeBill) = False Then GoTo ErrExit:
    zlCommFun.StopFlash
    Exit Sub
ErrExit:
    zlCommFun.StopFlash
    Unload Me: Exit Sub
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picTop.Move 0, 0, Me.ScaleWidth, picTop.Height
    
    sta.Move 0, Me.ScaleHeight - sta.Height, Me.ScaleWidth, sta.Height
    picBottom.Move 0, sta.Top - picBottom.Height, Me.ScaleWidth, picBottom.Height
    
    picBill.Move 0, picTop.Height, Me.ScaleWidth, picBottom.Top - picTop.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    zl_vsGrid_Para_Save 1131, vsfBill, Me.Caption, "����תסԺ�б�_New", True
    
    Set mrsPerson = Nothing
    Set mrsDepartment = Nothing
    Set mrsChargeitem = Nothing
    
    Set mrsPatient = Nothing
    Set mrsFeeBill = Nothing
End Sub

Private Sub cmdOk_Click()
    Dim strSql As String
    Dim blnTrans As Boolean, i As Integer
    Dim strData As String
    Dim strPreNo As String, strNO As String, int��� As Integer
    Dim str�Ǽ�ʱ�� As String
    
    On Error GoTo ErrHander
    mstr�˵�ID = ""
    Set mcllExecSql = New Collection
    
    str�Ǽ�ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:MM:ss")
    With mrsFeeBill
        .Sort = "���ݺ� Asc"
        Do While Not .EOF
            If InStr(mstr�˵�ID, "{""outp_bill_id"":" & NVL(mrsFeeBill!�˵�ID) & "}") = 0 Then
                mstr�˵�ID = mstr�˵�ID & ",{""outp_bill_id"":" & NVL(mrsFeeBill!�˵�ID) & "}"
            End If
            
            If NVL(!���ݺ�) <> strPreNo Then
                strNO = zlDatabase.NextNo(14)
                int��� = 1
                strPreNo = NVL(!���ݺ�)
            End If
            
            'Zl_�������תסԺ_����ת��(
            strSql = "Zl_�������תסԺ_����ת��("
            '  No_In         סԺ���ü�¼.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  ���_In       סԺ���ü�¼.���%Type,
            strSql = strSql & "" & int��� & ","
            '  ����id_In     סԺ���ü�¼.����id%Type,
            strSql = strSql & "" & NVL(mrsPatient!����ID) & ","
            '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type,
            strSql = strSql & "" & "{����ҳID��}" & ","
            '  ��ʶ��_In     סԺ���ü�¼.��ʶ��%Type,
            strSql = strSql & "" & "{��סԺ�š�}" & ","
            '  ����_In       סԺ���ü�¼.����%Type,
            strSql = strSql & "'" & NVL(mrsPatient!����) & "',"
            '  �Ա�_In       סԺ���ü�¼.�Ա�%Type,
            strSql = strSql & "'" & NVL(mrsPatient!�Ա�) & "',"
            '  ����_In       סԺ���ü�¼.����%Type,
            strSql = strSql & "'" & NVL(mrsPatient!����) & "',"
            '  ����_In       סԺ���ü�¼.����%Type,
            strSql = strSql & "'" & NVL(mrsPatient!����) & "',"
            '  �ѱ�_In       סԺ���ü�¼.�ѱ�%Type,
            strSql = strSql & "'" & NVL(!�ѱ�) & "',"
            '  ����id_In     סԺ���ü�¼.���˲���id%Type,
            strSql = strSql & "" & "{�����˲���ID��}" & ","
            '  ����id_In     סԺ���ü�¼.���˿���id%Type,
            strSql = strSql & "" & "{�����˿���ID��}" & ","
            '  ��������id_In סԺ���ü�¼.��������id%Type,
            strSql = strSql & "" & NVL(!��������ID) & ","
            '  ������_In     סԺ���ü�¼.������%Type,
            strSql = strSql & "'" & NVL(!������) & "',"
            '  ��������_In   סԺ���ü�¼.��������%Type,
            strSql = strSql & "" & "NULL" & ","
            '  �շ�ϸĿid_In סԺ���ü�¼.�շ�ϸĿid%Type,
            strSql = strSql & "" & NVL(!�շ�ϸĿID) & ","
            '  �շ����_In   סԺ���ü�¼.�շ����%Type,
            strSql = strSql & "'" & NVL(!���) & "',"
            '  ���㵥λ_In   סԺ���ü�¼.���㵥λ%Type,
            strSql = strSql & "'" & NVL(!��λ) & "',"
            '  ����_In       סԺ���ü�¼.����%Type,
            strSql = strSql & "" & NVL(!����) & ","
            '  ����_In       סԺ���ü�¼.����%Type,
            strSql = strSql & "" & NVL(!����) & ","
            '  ִ�в���id_In סԺ���ü�¼.ִ�в���id%Type,
            strSql = strSql & "" & NVL(!ִ�п���ID) & ","
            '  �۸񸸺�_In   סԺ���ü�¼.�۸񸸺�%Type,
            strSql = strSql & "" & "NULL" & ","
            '  ������Ŀid_In סԺ���ü�¼.������Ŀid%Type,
            strSql = strSql & "" & NVL(!������ĿID) & ","
            '  �վݷ�Ŀ_In   סԺ���ü�¼.�վݷ�Ŀ%Type,
            strSql = strSql & "'" & NVL(!�վݷ�Ŀ) & "',"
            '  ��׼����_In   סԺ���ü�¼.��׼����%Type,
            strSql = strSql & "" & NVL(!����) & ","
            '  Ӧ�ս��_In   סԺ���ü�¼.Ӧ�ս��%Type,
            strSql = strSql & "" & NVL(!Ӧ�ս��) & ","
            '  ʵ�ս��_In   סԺ���ü�¼.ʵ�ս��%Type,
            strSql = strSql & "" & NVL(!ʵ�ս��) & ","
            '  ����ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
            strSql = strSql & "To_Date('{����Ժʱ�䡻}','yyyy-mm-dd hh24:mi:ss'),"
            '  �Ǽ�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
            strSql = strSql & "To_Date('" & str�Ǽ�ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),"
            '  ����_In       Number,
            strSql = strSql & "" & IIf(NVL(!����״̬) = 1, 0, 1) & ","
            '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
            strSql = strSql & "'" & NVL(!����Ա���) & "',"
            '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
            strSql = strSql & "'" & NVL(!����Ա����) & "',"
            '  ִ����_In     סԺ���ü�¼.ִ����%Type,
            strSql = strSql & "'" & NVL(!ִ����) & "',"
            '  ִ��ʱ��_In   סԺ���ü�¼.ִ��ʱ��%Type,
            strSql = strSql & "To_Date('" & NVL(!ִ��ʱ��) & "','yyyy-mm-dd hh24:mi:ss'),"
            '  ҽ�����_In   סԺ���ü�¼.ҽ�����%Type:=Null
            strSql = strSql & "" & NVL(!ҽ��ID, "NULL") & ")"
            mcllExecSql.Add strSql
            
            int��� = int��� + 1
            mrsFeeBill.MoveNext
        Loop
    End With
    
    If mstr�˵�ID <> "" Then
        mstr�˵�ID = Mid(mstr�˵�ID, 2)
        mstr�˵�ID = "[" & mstr�˵�ID & "]"
    End If
    mstr�˵�ID = "{""input"":{""head"":{""bizno"":""RJ003"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}," & _
        """bill_list"":" & mstr�˵�ID & "}}"
    
    If mbln����ִ�� Then
        zlCommFun.ShowFlash "���ڽ����������תסԺ�������Ժ�...", Me
        cmdOK.Enabled = False
        gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To mcllExecSql.Count
            strSql = Replace(mcllExecSql(i), "{��סԺ�š�}", Val(NVL(mrsPatient!סԺ��)))
            strSql = Replace(strSql, "{����ҳID��}", Val(NVL(mrsPatient!��ҳID)))
            strSql = Replace(strSql, "{����Ժʱ�䡻}", Format(mrsPatient!��Ժ����, "yyyy-mm-dd hh:MM:ss"))
            strSql = Replace(strSql, "{�����˿���ID��}", Val(NVL(mrsPatient!���˿���ID)))
            strSql = Replace(strSql, "{�����˲���ID��}", Val(NVL(mrsPatient!���˲���ID)))
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        Next
        
        '����������������תסԺȷ�ϡ�����
        '����   ����            ����        ˵��    ��������        ��ע
        '       outp_bill_id    �����˵�ID          Number(18)      �ǿ�
        '���   ����        ����        ˵��                ��������        ��ע
        '       result      ִ�н��    1-�ɹ���-1-ʧ��     Number(1)       �ǿ�
        '       errmsg      ������Ϣ    ʧ��ʱ���ش�����Ϣ  Varchar2(200)
        Call Sys.NewSystemSvr("������ϵͳ", "�������תסԺ����ȷ��", mstr�˵�ID, strData)
        If strData = "" Then strData = "{}"
        If Val(zlStr.JSONParse("result", strData)) <> 1 Then
            gcnOracle.RollbackTrans
            zlCommFun.StopFlash: cmdOK.Enabled = True
            MsgBox zlStr.JSONParse("errmsg", strData), vbInformation, gstrSysName
            Exit Sub
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        zlCommFun.StopFlash: cmdOK.Enabled = True
    End If
    
    mblnOk = True
    Unload Me
    Exit Sub
ErrHander:
    zlCommFun.StopFlash
    cmdOK.Enabled = True
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub picBill_Resize()
    On Error Resume Next
    vsfBill.Move 0, 0, picBill.ScaleWidth, picBill.ScaleHeight
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    With picBottom
        cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.width - 1000
        cmdOK.Left = cmdCancel.Left - cmdOK.width - 100
    End With
End Sub

Private Function InitFace() As Boolean
    '��ʼ������
    Dim strHead As String
    Dim varHead As Variant, varItem As Variant
    Dim i As Long
    
    On Error GoTo ErrHandler
    With vsfBill
        .Redraw = flexRDNone
        .RowHeightMin = 300
        .Clear
        .Rows = 2
        .FixedRows = 1: .FixedCols = 0
        
        strHead = "���ݺ�,1,0|�˵�ID,1,0|��������,1,0|������,1,0|�ѱ�,1,0|����,1,0|" & _
                "���,1,800|����,1,2100|���,1,1400|��λ,1,600|����,7,800|����,7,1000|" & _
                "Ӧ�ս��,7,1000|ʵ�ս��,7,1000|ִ�п���,1,1000|˵��,1,850|����ʱ��,1,1800"
        varHead = Split(strHead, "|")
        .Cols = UBound(varHead) + 1
        For i = 0 To UBound(varHead)
            varItem = Split(varHead(i), ",")
            .TextMatrix(0, i) = varItem(0)
            .ColKey(i) = varItem(0)
            .ColAlignment(i) = varItem(1)
            .ColWidth(i) = varItem(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        zl_vsGrid_Para_Restore 1131, vsfBill, Me.Caption, "����תסԺ�б�_New", True
        .Redraw = flexRDBuffered
    End With
    
    lbl(txtName).Caption = NVL(mrsPatient!����)
    lbl(txtSex).Caption = NVL(mrsPatient!�Ա�)
    lbl(txtAge).Caption = NVL(mrsPatient!����)
    lbl(txtInNumber).Caption = NVL(mrsPatient!סԺ��)
    Call SetPatiControl
    InitFace = True
    Exit Function
ErrHandler:
    vsfBill.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData(ByVal lng�Һ�ID As Long) As Boolean
    '��ʼ������
    Dim strSql As String
    
    On Error GoTo ErrHandler
    '��Ա
    strSql = _
        "Select ID, ���, ����" & vbNewLine & _
        "From ��Ա��" & vbNewLine & _
        "Where (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01','yyyy-mm-dd'))"
    Set mrsPerson = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '����
    strSql = _
        "Select ID, ����, ����" & vbNewLine & _
        "From ���ű�" & vbNewLine & _
        "Where (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01','yyyy-mm-dd'))"
    Set mrsDepartment = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '�շ���Ŀ
    strSql = _
        "Select Distinct a.Id, a.����, a.���, a.���㵥λ, a.���, d.���� As �������, b.������ĿID, c.�վݷ�Ŀ" & vbNewLine & _
        "From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շ���Ŀ��� D" & vbNewLine & _
        "Where a.Id = b.�շ�ϸĿid And b.������Ŀid = c.Id And a.��� = d.����" & vbNewLine & _
        "      And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01','yyyy-mm-dd'))"
    Set mrsChargeitem = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '������Ϣ
   strSql = _
        "Select a.����id, a.��ҳid, a.סԺ��," & vbNewLine & _
        "       Nvl(a.����, b.����) As ����, Nvl(a.�Ա�, b.�Ա�) As �Ա�, Nvl(a.����, b.����) As ����," & vbNewLine & _
        "       Nvl(a.�ѱ�, b.�ѱ�) As �ѱ�,Nvl(a.��Ժ����, b.��ǰ����) As ����, a.��Ժ����," & vbNewLine & _
        "       a.��Ժ����ID As ���˲���id, a.��Ժ����id As ���˿���id" & vbNewLine & _
        "From ������ҳ A, ������Ϣ B" & vbNewLine & _
        "Where a.����id = b.����id And a.�Һ�id = [1]"
    Set mrsPatient = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng�Һ�ID)
    If mrsPatient.RecordCount = 0 Then
        MsgBox "δ�ҵ�������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBills(ByVal rsBill As ADODB.Recordset) As Boolean
    '��ʾ���˿�����תסԺ�ķ���
    Dim i As Integer
    Dim str���� As String
    
    On Error GoTo ErrHandler
    With vsfBill
        .Redraw = flexRDNone
        .Clear 1
        .Rows = rsBill.RecordCount + 1
        i = 1
        If rsBill.RecordCount > 0 Then rsBill.MoveFirst
        Do While Not rsBill.EOF
            .TextMatrix(i, .ColIndex("�˵�ID")) = NVL(rsBill!�˵�ID)
            .TextMatrix(i, .ColIndex("���ݺ�")) = NVL(rsBill!���ݺ�)
            .TextMatrix(i, .ColIndex("��������")) = NVL(rsBill!������������)
            .TextMatrix(i, .ColIndex("������")) = NVL(rsBill!������)
            .TextMatrix(i, .ColIndex("�ѱ�")) = NVL(rsBill!�ѱ�)
            
            If Val(NVL(rsBill!��������)) = 2 Then
                str���� = IIf(Val(NVL(rsBill!����״̬)) = 0, "���˻��۵�", "���˵�")
            Else
                str���� = IIf(Val(NVL(rsBill!����״̬)) = 0, "�շѻ��۵�", "�շѵ�")
            End If
            .TextMatrix(i, .ColIndex("����")) = str����
            .TextMatrix(i, .ColIndex("���")) = NVL(rsBill!�������) & IIf(i Mod 2 = 1, "", " ")
            .TextMatrix(i, .ColIndex("����")) = NVL(rsBill!��Ŀ����)
            .TextMatrix(i, .ColIndex("���")) = NVL(rsBill!���)
            .TextMatrix(i, .ColIndex("��λ")) = NVL(rsBill!��λ)
            .TextMatrix(i, .ColIndex("����")) = FormatEx(NVL(rsBill!����) * NVL(rsBill!����), 6, , , 2)
            .TextMatrix(i, .ColIndex("����")) = FormatEx(NVL(rsBill!����), 6, , , 2)
            
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = FormatEx(NVL(rsBill!Ӧ�ս��), 6, , , 2)
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = FormatEx(NVL(rsBill!ʵ�ս��), 6, , , 2)
            .TextMatrix(i, .ColIndex("ִ�п���")) = NVL(rsBill!ִ�п�������)
            .TextMatrix(i, .ColIndex("˵��")) = IIf(NVL(rsBill!ִ����) = "", "δִ��", "��ȫִ��")
            .TextMatrix(i, .ColIndex("����ʱ��")) = Format(NVL(rsBill!����ʱ��), "yyyy-mm-dd hh:MM:ss")
            
            i = i + 1
            rsBill.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
   
    Call SetSumMoney 'ת���ϼ�
    Call SplitGroupShow '������ʾ
    ShowBills = True
    Exit Function
ErrHandler:
    vsfBill.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SplitGroupShow()
    '�����б���Ϣ���з�����ʾ
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo ErrHandler
    With vsfBill
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("Ӧ�ս��"), gstrDec, &H8000000F, , False, "%s", , True
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("ʵ�ս��"), gstrDec, &H8000000F, , False, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("���")
        .OutlineCol = .ColIndex("���")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                .MergeRow(i) = True
                .RowHeight(i) = 350

                strTemp = .TextMatrix(i + 1, .ColIndex("���ݺ�")) & "(" & .TextMatrix(i + 1, .ColIndex("����")) & ")"
                strTemp = strTemp & Space(2) & "�ѱ�:" & .TextMatrix(i + 1, .ColIndex("�ѱ�"))
                strTemp = strTemp & Space(2) & "��������:" & .TextMatrix(i + 1, .ColIndex("��������"))
                strTemp = strTemp & Space(2) & "������:" & .TextMatrix(i + 1, .ColIndex("������"))
                
                For j = 0 To .Cols - 1
                   If j >= .ColIndex("���") And j < .ColIndex("Ӧ�ս��") Then
                       .Cell(flexcpText, i, j) = strTemp
                   ElseIf .ColIndex("Ӧ�ս��") = j Then
                       .TextMatrix(i, j) = FormatEx(Val(.TextMatrix(i, j)), 6, , , 2)
                   ElseIf .ColIndex("ʵ�ս��") = j Then
                       .TextMatrix(i, j) = " " & FormatEx(Val(.TextMatrix(i, j)), 6, , , 2)
                   End If
                Next
            End If
        Next
        
        .MergeCells = flexMergeRestrictRows
        For i = 0 To .Cols - 1
            If i < .ColIndex("Ӧ�ս��") Then
                .MergeCol(i) = True
            Else
                .MergeCol(i) = False
            End If
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetSumMoney()
    '���ú���ʾ���úϼ�
    Dim i As Long, dblSum As Double
    
    On Error GoTo ErrHander
    With vsfBill
        For i = .FixedRows To .Rows - 1
            dblSum = dblSum + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
        Next
    End With
    txtSum.Text = Format(dblSum, "###0.00;-###0.00;0.00;0.00")
    Exit Sub
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetPatiControl()
    '���ò�����Ϣ�ؼ�λ��
    Dim sngSplit As Single
    
    sngSplit = 600
    On Error Resume Next
    lbl(txtName).Left = lbl(lblName).Left + lbl(lblName).width
    
    lbl(lblSex).Left = lbl(txtName).Left + lbl(txtName).width + sngSplit
    lbl(txtSex).Left = lbl(lblSex).Left + lbl(lblSex).width
    
    lbl(lblAge).Left = lbl(txtSex).Left + lbl(txtSex).width + sngSplit
    lbl(txtAge).Left = lbl(lblAge).Left + lbl(lblAge).width
    
    lbl(lblInNumber).Left = lbl(txtAge).Left + lbl(txtAge).width + sngSplit
    lbl(txtInNumber).Left = lbl(lblInNumber).Left + lbl(lblInNumber).width
End Sub

Private Function GetBillData(ByVal lng�Һ�ID As Long, ByRef strData As String) As Boolean
    'ͨ�������ȡ����
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strJsonIn As String, strNO As String
    
    On Error GoTo ErrHandler
    strSql = "Select NO From ���˹Һż�¼ Where ID =  [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng�Һ�ID)
    If rsTemp.EOF Then
        MsgBox "δ�ҵ����˹Һż�¼���޷�ȷ���Һŵ��ݺţ�", vbInformation, gstrSysName
        Exit Function
    End If
    strNO = NVL(rsTemp!NO)
    
    '����������������תסԺ������
    '����    ����       ����        ˵��    ��������        ��ע
    '        rgst_no    �Һŵ���            Varchar2(8)      �ǿ�
    '���    ����                   ����            ˵��                    ��������    ��ע
    '        result                 ִ�н��        1-�ɹ���-1-ʧ��         Number(1)   �ǿ�
    '        errmsg                 ������Ϣ        ʧ��ʱ���ش�����Ϣ      Varchar2(200)
    '        outp_bill_id           �����˵�ID      ZLHIS����ȷ����         Number(18)  �ǿ�
    '        outp_bill_no           ���ݺ�          �˵��ţ����ֵ���        VARCHAR2(20)    �ǿ�
    '        outp_kacnt_sign        ��������        1-�շѵ�;2-���˵�       Number(1)   �ǿ�
    '        pricing_sign           ����״̬        0-�����շѵ�/���ۼ��˵�;1-�����շѵ�/������˵� Number(1)   �ǿ�
    '        plcdept_id             ��������ID      ���ű�.ID               Number(18)  �ǿ�
    '        placer                 ����ҽ��        ��Ա��.����             VARCHAR2(70)    �ǿ�
    '        outp_bill_time         ����ʱ��        �����˵�ʱ��            Date    �ǿ�
    '        order_id               ҽ��ID          ����ҽ����¼.ID         Number(18)
    '        outp_bill_creator_id   ����ԱID        �����˵�������ID         Number(18)  �ǿ�
    '        category_id            �ѱ�                                    VARCHAR2(20)
    '        fee_id                 �շ���ĿID      �շ���ĿĿ¼.ID         Number(18)  �ǿ�
    '        acntsubj_id            ������ĿID      ������Ŀ.ID             Number(18)  �ǿ�
    '        crx_qunt               ����            �в�ҩ�ܼ���            NUMBER(4)   �ǿ�
    '        outp_bill_detail_qunt  ����                                    NUMBER(18,5)    �ǿ�
    '        fee_now_disct_price    ����                                    NUMBER(18,4)    �ǿ�
    '        outp_bill_detail_chrg  Ӧ�ս��        ����*����*����          NUMBER(18,3)    �ǿ�
    '        outp_bill_detail_disct_chrg ʵ�ս��   Ӧ�ս��-�ۿ۽��       NUMBER(18,3)    �ǿ�
    '        exedept_id             ִ�п���ID      ���ű�.ID               Number(18)  �ǿ�
    '        exetr                  ִ����          ��Ա��.������Ϊ�ձ�ʾδִ�У���Ϊ�ձ�ʾ��ȫִ�� VARCHAR2(70)
    '        exetime                ִ��ʱ��                                Date
    strJsonIn = "{""head"":{""bizno"":""RJ002"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}"
    strJsonIn = "{""input"":" & strJsonIn & ",""rgst_no"":""" & strNO & """}}"
    Call Sys.NewSystemSvr("������ϵͳ", "�������תסԺ����", strJsonIn, strData)
    If strData = "" Then strData = "{}"
    If Val(zlStr.JSONParse("result", strData)) <> 1 Then
        MsgBox "��ȡ���������תסԺ�ķ�����Ϣʱ����" & vbCrLf & _
            zlStr.JSONParse("errmsg", strData), vbInformation, gstrSysName
        Exit Function
    End If
    GetBillData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AnalyzeData(ByVal strData As String, ByRef rsRecord As ADODB.Recordset) As Boolean
    '��JSON�ַ����н�������
    '��Σ�
    '   strData JSON�ַ���
    '���Σ�
    '   rsRecord ���ü�¼
    '���أ������ɹ�������True,���򷵻�False
    Dim i As Integer
    Dim objScript As Object '���ڽ���JSON
    
    On Error GoTo ErrHandler
    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "JScript"
    objScript.AddCode "var obj=" & strData & ";"
    
    Set rsRecord = CreateBillRecord()
    
    With rsRecord
        For i = 0 To objScript.Eval("obj.bill_list.length") - 1
            .AddNew
            !�˵�ID = objScript.Eval("obj.bill_list[" & i & "].outp_bill_id")
            !���ݺ� = objScript.Eval("obj.bill_list[" & i & "].outp_bill_no")
            !�������� = objScript.Eval("obj.bill_list[" & i & "].outp_kacnt_sign")
            !����״̬ = objScript.Eval("obj.bill_list[" & i & "].pricing_sign")
            !��������ID = objScript.Eval("obj.bill_list[" & i & "].plcdept_id")
            mrsDepartment.Filter = "ID=" & !��������ID
            If mrsDepartment.EOF Then
                MsgBox "δ�ҵ����ݡ�" & NVL(!���ݺ�) & "���Ŀ���������Ϣ��", vbInformation, gstrSysName
                Exit Function
            Else
                !������������ = NVL(mrsDepartment!����)
            End If
            !������ = objScript.Eval("obj.bill_list[" & i & "].placer")
            !�ѱ� = objScript.Eval("obj.bill_list[" & i & "].category_id")
            If NVL(!�ѱ�) = "" Then !�ѱ� = NVL(mrsPatient!�ѱ�)
            !����ʱ�� = objScript.Eval("obj.bill_list[" & i & "].outp_bill_time")
            mrsPerson.Filter = "ID=" & Val(objScript.Eval("obj.bill_list[" & i & "].outp_bill_creator_id"))
            If mrsPerson.EOF Then
                MsgBox "δ�ҵ����ݡ�" & NVL(!���ݺ�) & "���Ĳ���Ա��Ϣ��", vbInformation, gstrSysName
                Exit Function
            Else
                !����Ա���� = NVL(mrsPerson!����)
                !����Ա��� = NVL(mrsPerson!���)
            End If

            !�շ�ϸĿID = objScript.Eval("obj.bill_list[" & i & "].fee_id")
            !������ĿID = objScript.Eval("obj.bill_list[" & i & "].acntsubj_id")
            mrsChargeitem.Filter = "ID=" & !�շ�ϸĿID & " And ������ĿID=" & !������ĿID
            If mrsChargeitem.EOF Then
                MsgBox "δ�ҵ����ݡ�" & NVL(!���ݺ�) & "�����շ���Ŀ��Ϣ��", vbInformation, gstrSysName
                Exit Function
            Else
                !��� = NVL(mrsChargeitem!���)
                !������� = NVL(mrsChargeitem!�������)
                !��Ŀ���� = NVL(mrsChargeitem!����)
                !��� = NVL(mrsChargeitem!���)
                !��λ = NVL(mrsChargeitem!���㵥λ)
                !�վݷ�Ŀ = NVL(mrsChargeitem!�վݷ�Ŀ)
            End If
            !���� = objScript.Eval("obj.bill_list[" & i & "].crx_qunt")
            If Val(NVL(!����)) = 0 Then !���� = 1
            !���� = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_qunt")
            !���� = objScript.Eval("obj.bill_list[" & i & "].fee_now_disct_price")
            !Ӧ�ս�� = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_chrg")
            !ʵ�ս�� = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_disct_chrg")

            !ִ�п���ID = objScript.Eval("obj.bill_list[" & i & "].exedept_id")
            mrsDepartment.Filter = "ID=" & !ִ�п���ID
            If mrsDepartment.EOF Then
                MsgBox "δ�ҵ����ݡ�" & NVL(!���ݺ�) & "����ִ�п�����Ϣ��", vbInformation, gstrSysName
                Exit Function
            Else
                !ִ�п������� = NVL(mrsDepartment!����)
            End If
            !ִ���� = objScript.Eval("obj.bill_list[" & i & "].exetr")
            !ִ��ʱ�� = objScript.Eval("obj.bill_list[" & i & "].exetime")
            !ҽ��ID = objScript.Eval("obj.bill_list[" & i & "].order_id")
        Next
        .UpdateBatch '��������
    End With
    Set objScript = Nothing
    AnalyzeData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateBillRecord() As ADODB.Recordset
    '������¼������
    Dim rsRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Set rsRecord = New ADODB.Recordset
    rsRecord.Fields.Append "�˵�ID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "���ݺ�", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "��������", adInteger, , adFldIsNullable
    rsRecord.Fields.Append "����״̬", adInteger, , adFldIsNullable
    rsRecord.Fields.Append "��������ID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "������������", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "�ѱ�", adVarChar, 50, adFldIsNullable
    rsRecord.Fields.Append "����ʱ��", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "����Ա���", adVarChar, 10, adFldIsNullable
    rsRecord.Fields.Append "����Ա����", adVarChar, 100, adFldIsNullable
    
    rsRecord.Fields.Append "���", adVarChar, 10, adFldIsNullable
    rsRecord.Fields.Append "�������", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "��Ŀ����", adVarChar, 200, adFldIsNullable
    rsRecord.Fields.Append "���", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "��λ", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "�վݷ�Ŀ", adVarChar, 50, adFldIsNullable
    rsRecord.Fields.Append "����", adDouble, , adFldIsNullable
    rsRecord.Fields.Append "����", adDouble, , adFldIsNullable
    rsRecord.Fields.Append "����", adDouble, , adFldIsNullable
    rsRecord.Fields.Append "Ӧ�ս��", adDouble, , adFldIsNullable
    rsRecord.Fields.Append "ʵ�ս��", adDouble, , adFldIsNullable
    
    rsRecord.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "ִ�п�������", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "ִ����", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "ִ��ʱ��", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "ҽ��ID", adBigInt, , adFldIsNullable
    
    rsRecord.CursorLocation = adUseClient
    rsRecord.LockType = adLockOptimistic
    rsRecord.CursorType = adOpenStatic
    rsRecord.Open
    
    Set CreateBillRecord = rsRecord
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
