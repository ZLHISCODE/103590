VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistPlanDatSet 
   AutoRedraw      =   -1  'True
   Caption         =   "�ҺŰ���ʱ������"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   Icon            =   "frmregistplandatset.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10410
   StartUpPosition =   1  '����������
   Begin VB.Frame fraӦ���� 
      Caption         =   "Ӧ����(&B)"
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   6650
      Width           =   7095
      Begin VB.OptionButton opt���� 
         Caption         =   "Ӧ��������"
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "Ӧ���뱾����"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt��ҽ�� 
         Caption         =   "Ӧ���ڱ�ҽ��"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8640
      TabIndex        =   25
      Top             =   6840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   24
      Top             =   6840
      Width           =   1100
   End
   Begin VB.Frame fraInfo 
      Caption         =   "������Ϣ"
      Height          =   1380
      Left            =   120
      TabIndex        =   6
      Top             =   75
      Width           =   10095
      Begin VB.ComboBox cbo���� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3360
         TabIndex        =   26
         Text            =   "cbo����"
         Top             =   307
         Width           =   1155
      End
      Begin VB.TextBox txt��Լ 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6720
         MaxLength       =   5
         TabIndex        =   14
         Top             =   307
         Width           =   1215
      End
      Begin VB.TextBox txt�޺� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4965
         MaxLength       =   5
         TabIndex        =   13
         Top             =   307
         Width           =   975
      End
      Begin VB.CheckBox chk��ſ��� 
         Caption         =   "��ſ���"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   330
         Width           =   1095
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�Һ�ʱ���뽨����"
         Enabled         =   0   'False
         Height          =   195
         Left            =   8040
         TabIndex        =   5
         Top             =   360
         Width           =   1845
      End
      Begin VB.ComboBox cbo���� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         TabIndex        =   2
         Text            =   "cbo����"
         Top             =   705
         Width           =   2115
      End
      Begin VB.ComboBox cboDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6720
         TabIndex        =   4
         Top             =   705
         Width           =   2115
      End
      Begin VB.ComboBox cboItem 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3360
         TabIndex        =   3
         Text            =   "cboItem"
         Top             =   705
         Width           =   2235
      End
      Begin VB.TextBox txt�ű� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   5
         TabIndex        =   0
         Top             =   307
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��Լ"
         Height          =   180
         Left            =   6240
         TabIndex        =   16
         Top             =   367
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�޺�"
         Height          =   180
         Left            =   4560
         TabIndex        =   15
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   3000
         TabIndex        =   12
         Top             =   367
         Width           =   360
      End
      Begin VB.Label lblҽ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ժ��ҽ��"
         Height          =   180
         Left            =   5940
         TabIndex        =   10
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ"
         Height          =   180
         Left            =   3000
         TabIndex        =   9
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ű�"
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
         Left            =   120
         TabIndex        =   7
         Top             =   367
         Width           =   390
      End
   End
   Begin VB.Frame fraDate 
      Caption         =   "ʱ������"
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   10215
      Begin VB.PictureBox picTime 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   120
         ScaleHeight     =   4665
         ScaleWidth      =   9945
         TabIndex        =   17
         Top             =   240
         Width           =   9945
         Begin VB.CommandButton cmdOtherCalc 
            Caption         =   "���������ƻ�(&R)"
            Height          =   345
            Left            =   3735
            TabIndex        =   32
            Top             =   30
            Width           =   1860
         End
         Begin VB.CommandButton cmd����ʱ�� 
            Caption         =   "��������(&F)"
            Height          =   350
            Left            =   2520
            TabIndex        =   31
            ToolTipText     =   "������¼���ʱ��"
            Top             =   35
            Width           =   1150
         End
         Begin VB.TextBox txtTimeOut 
            Height          =   300
            Left            =   1560
            TabIndex        =   29
            Text            =   "10"
            Top             =   60
            Width           =   500
         End
         Begin MSComCtl2.UpDown udTime 
            Height          =   345
            Left            =   2160
            TabIndex        =   28
            Top             =   38
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComctlLib.TabStrip tbWeekTime 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTime 
            Height          =   3825
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   9765
            _cx             =   17224
            _cy             =   6747
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
            GridColor       =   12632256
            GridColorFixed  =   0
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmregistplandatset.frx":000C
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
            Editable        =   2
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
            Begin VB.CommandButton cmdɾ�� 
               Caption         =   "ɾ"
               Height          =   255
               Left            =   7305
               TabIndex        =   33
               Top             =   2025
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton cmdԤԼ 
               Caption         =   "Ԥ"
               Height          =   255
               Left            =   7320
               TabIndex        =   27
               Top             =   1560
               Visible         =   0   'False
               Width           =   375
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ʱ����(��)"
            Height          =   180
            Left            =   360
            TabIndex        =   30
            Top             =   120
            Width           =   1080
         End
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "Ժ��ҽ��"
         Index           =   0
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "����Ԯҽ��"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRegistPlanDatSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit 'Ҫ���������
'
Public Enum ViewMode
     NewItem
     ViewItem '�鿴
     Edit '�༭
End Enum
Private mViewMode         As ViewMode    'ҳ����ʾģʽ
Private mlng����ID        As Long        '����ID
Private mlngPre����ID     As Long
Private mrsTime          As ADODB.Recordset
Private mrs�޺�          As ADODB.Recordset
Private mrs�ϰ�ʱ���     As ADODB.Recordset
Private mrs����          As ADODB.Recordset
Private mblnCellChange   As Boolean
Private mstrKey         As String
Private mblnChange      As Boolean
Private mblnReload      As Boolean '�ڹҺŰ��Ź���ҳ����� ShowMe�Ժ� �Ƿ���Ҫˢ��
Private mstr�����޸� As String '��ĳһ����߶���İ������Ƹ���
'�����ϰ�ʱ��
Private Type t_�ϰ�ʱ��
  dat_�����ϰ� As Date
  dat_�����°� As Date
  dat_�����ϰ� As Date
  dat_�����°� As Date
End Type
Private t_ʱ�� As t_�ϰ�ʱ��
Private Const strMaskKey As String = "09:00-09:00"
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther
Attribute mfrmOtherCalc.VB_VarHelpID = -1

Private Sub chk��ſ���_Click()
    cmdOtherCalc.Visible = chk��ſ���.Value = 1
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
    cmdOK.Enabled = False
    zlCommFun.ShowFlash "���ڱ���ҺŰ���ʱ������,���Ժ򡭡�"
    If SaveDate() = True Then
        '************************
        '�������ɹ���Ҫ���¶�
        '�ҺŰ���ʱ�ν�����ȡ
        '************************
        Call InitData
        mblnChange = False
        mblnReload = True
        'If tbWeekTime.Tabs.Count > 0 Then tbWeekTime.Tabs(1).Selected = True
    End If
    zlCommFun.StopFlash
   cmdOK.Enabled = True
End Sub


Private Sub cmdOtherCalc_Click()
    Dim str���� As String
    
    If chk��ſ���.Value <> 1 Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    
    Set mfrmOtherCalc = New frmRegistPlanTimeOther
    str���� = Replace(Split(tbWeekTime.SelectedItem.Caption & "(", "(")(1), ")", "")
    Call mfrmOtherCalc.zlShowMe(Me, str����, Val(txtTimeOut.Text))
    If Not mfrmOtherCalc Is Nothing Then Unload mfrmOtherCalc
    Set mfrmOtherCalc = Nothing
End Sub

Private Sub cmdɾ��_Click()
    Call DeleteSelectPain
End Sub

Private Sub cmd����ʱ��_Click()
'�ԹҺŰ���ʱ�ν�������
    Dim str����         As String
    
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrsTime.Filter = "����='" & str���� & "'"
    If mrsTime.RecordCount > 0 Then
      '****************************************************************
      '�����йҺŰ���ʱ�ε������
      '��ʾ����Ա �Ƿ���Ҫ���¼���ʱ��
      '****************************************************************
        If MsgBox("�˰�����" & str���� & "�Ѿ�����ʱ�� " & vbCrLf & "�Ƿ����¼���ʱ��?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            mrsTime.Filter = 0
            Exit Sub
        End If
    End If
    Select Case chk��ſ���.Value = 1
    Case True:
        Setר�Һ�ʱ��
        setVsFlexBgColor (True)
    Case False:
        Set��ͨ��ʱ��
        setVsFlexBgColor
    End Select
    
    mblnChange = True
End Sub
Private Sub Set��ͨ��ʱ��()
    Dim strSQL      As String
    Dim str����     As String
    Dim strʱ��     As String
    Dim lng�޺�     As Long
    Dim lng��Լ     As Long
    Dim lng���     As Long
    Dim dblDatCount As Long '��ʱ����
    Dim datʱ��     As Date 'ÿ��ʱ��ε�
    Dim blnȫ��     As Boolean  '�Ƿ���ȫ�춼����Һ� �����ȫ�����Ϊ���������
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim i           As Long
    Dim j           As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim strData     As String
    Dim strTime     As String
    Dim strList()   As String
    Dim blnExit     As Boolean
    Dim lngIndex    As Long
    Dim lngStart    As Long
    On Error GoTo Hd
    If mrs�ϰ�ʱ��� Is Nothing Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrs�޺�.Filter = "����='" & str���� & "'"
    If mrs�޺�.RecordCount = 0 Then
        MsgBox "��ǰ�ű���" & str���� & ",û�ж�Ӧ�ĹҺŰ�������" & vbCrLf & "�뵽�ҺŰ���������!", vbOKOnly, Me.Caption
        Exit Sub '����ҺŰ�����û�����ô������Ϣ �Ͳ���������
    End If
    lng�޺� = Nvl(mrs�޺�!�޺���, 0): lng��Լ = Nvl(mrs�޺�!��Լ��, 0)
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str���� & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Sub
    End If
    Me.txt�޺�.Text = lng�޺�
    Me.txt��Լ.Text = lng��Լ
    If lng��Լ = 0 Then lng��Լ = lng�޺� '�����ԤԼû����������Ϊ�����Լ�����޺�����ͬ
    strʱ�� = Nvl(mrs����(str����).Value)
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
     
    '*********************************
    '��ʱ�ξ��崦�� ��Ϊȫ��ͷ�ȫ��
    'ȫ���Ϊ���������
    '*********************************
  
    lng��� = Val(txtTimeOut.Text)
   
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
    End With
    '*************************************
    '��ͨ��
    '*************************************
    With vsTime
        .Cols = 8: .FixedCols = 0
        .Rows = 1: .FixedRows = 1
        For i = 0 To .Cols - 1 Step 2
           .TextMatrix(0, i) = "ʱ���"
        Next
        For i = 1 To .Cols - 1 Step 2
           .TextMatrix(0, i) = "ԤԼ����"
        Next
        lngRow = 1: lngCol = -1
        j = 1: lngStart = 1
        Do While Not mrs�ϰ�ʱ���.EOF
            If blnExit Then Exit Do
            datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00"))
            For i = j To lng�޺�
                If lngStart > lng�޺� Then
                    blnExit = True
                    Exit For
                End If
              
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                
                lngCol = lngCol + 1
                If lngCol * 2 > .Cols - 2 Then lngRow = lngRow + 1: lngCol = 0
                strData = IIf(lng��Լ >= i, 1, 0)
                strTime = Format(datʱ��, "HH:mm") & "-" & _
                      IIf(Format(DateAdd("n", lng���, datʱ��), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
                      Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
               
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, lngCol * 2) = strTime
                .TextMatrix(lngRow, lngCol * 2 + 1) = strData
                lngStart = lngStart + 1
                datʱ�� = DateAdd("n", lng���, datʱ��)
            Next
            mrs�ϰ�ʱ���.MoveNext
        Loop
       
 
         For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
         Next
         .Redraw = flexRDBuffered
    End With
     
Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
Private Sub Setר�Һ�ʱ��()
    Dim strSQL      As String
    Dim str����     As String
    Dim strʱ��     As String
    Dim lng�޺�     As Long
    Dim lng��Լ     As Long
    Dim lng���     As Long
    Dim dblDatCount As Long '��ʱ����
    Dim datʱ��     As Date 'ÿ��ʱ��ε�
    Dim strʱ��     As String
    Dim blnȫ��     As Boolean  '�Ƿ���ȫ�춼����Һ� �����ȫ�����Ϊ���������
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim i           As Long
    Dim j           As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim strData     As String
    Dim strTime     As String
    Dim strList()   As String
    Dim blnExit     As Boolean
    Dim lngIndex    As Long
    Dim lngStart    As Long
    On Error GoTo Hd
    If mrs�ϰ�ʱ��� Is Nothing Then Exit Sub
    If mrs�޺� Is Nothing Then
        strSQL = _
        "Select ����id, ������Ŀ as ���� , �޺���, ��Լ�� From �ҺŰ������� Where ����id = [1]"
        Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(txt�ű�.Tag))
        If mrsTime.RecordCount = 0 Then
            MsgBox "��ǰ�ű�û�ж�Ӧ�ĹҺŰ�������" & vbCrLf & "�뵽�ҺŰ���������!", vbOKOnly, Me.Caption
            Set mrs�޺� = Nothing
            Exit Sub '����ҺŰ�����û�����ô������Ϣ �Ͳ���������
        End If
    End If
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrs�޺�.Filter = "����='" & str���� & "'"
    If mrs�޺�.RecordCount = 0 Then
        MsgBox "��ǰ�ű���" & str���� & ",û�ж�Ӧ�ĹҺŰ�������" & vbCrLf & "�뵽�ҺŰ���������!", vbOKOnly, Me.Caption
        Exit Sub '����ҺŰ�����û�����ô������Ϣ �Ͳ���������
    End If
    lng�޺� = Nvl(mrs�޺�!�޺���, 0): lng��Լ = Nvl(mrs�޺�!��Լ��, 0)
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str���� & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Sub
    End If
    Me.txt�޺�.Text = lng�޺�
    Me.txt��Լ.Text = lng��Լ
    lng��Լ = lng�޺�
    strʱ�� = Nvl(mrs����(str����).Value)
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    
'*************************************************************
'ʱ�������� ���õļ��
'*************************************************************
      lng��� = Val(Me.txtTimeOut.Text)
     ' datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!��ʼʱ��, "00:00:00"))
        
      With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
      End With
    '*************************************
    'ר�Һ�
    '���������
    '���� ʱ��α��е� ���°�ʱ�����ж�
    '���� ȫ���������  ��Ϊ���������
    '*************************************
    
    With vsTime
         .Cols = 2
         lngRow = -1: lngCol = 0
         j = 1
         lngStart = 1
         Do While Not mrs�ϰ�ʱ���.EOF
            If blnExit Then Exit Do
             
            datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00"))
             For i = j To lng��Լ
                If lngStart > lng��Լ Then
                    blnExit = True
                    Exit For
                End If
              
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                 End If
                lngCol = lngCol + 1
                If strʱ�� <> Format(datʱ��, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
                If lngCol = 1 Then
                     If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
                     strʱ�� = Format(datʱ��, "HH") & ":00"
                     vsTime.TextMatrix(lngRow - 1, 0) = strʱ��
                     vsTime.TextMatrix(lngRow, 0) = strʱ��
                
                End If
                strData = lngStart
                lngStart = lngStart + 1
                strTime = Format(datʱ��, "HH:mm") & "-" & _
                           IIf(Format(DateAdd("n", lng���, datʱ��), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
                           Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
    
                If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
                vsTime.TextMatrix(lngRow - 1, lngCol) = strData
                vsTime.TextMatrix(lngRow, lngCol) = strTime
                '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
                
                datʱ�� = DateAdd("n", lng���, datʱ��)
             Next
             mrs�ϰ�ʱ���.MoveNext
         Loop
'         '***********************
'         '������
'         '**********************
'         For i = 1 To lng��Լ
'            If Format(datʱ��, "dd:mm:ss") >= Format(CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), "dd:mm:ss") Then Exit For
'            lngCol = lngCol + 1
'            If strʱ�� <> Format(datʱ��, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
'            If lngCol = 1 Then
'                 If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'                 strʱ�� = Format(datʱ��, "HH") & ":00"
'                 vsTime.TextMatrix(lngRow - 1, 0) = strʱ��
'                 vsTime.TextMatrix(lngRow, 0) = strʱ��
'
'            End If
'            strData = i
'            strTime = Format(datʱ��, "HH:mm") & "-" & _
'                       IIf(DateAdd("n", lng���, datʱ��) > CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), _
'                       Format(CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
'
'            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
'            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
'            vsTime.TextMatrix(lngRow, lngCol) = strTime
'            '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
'
'            datʱ�� = DateAdd("n", lng���, datʱ��)
'         Next
'         If blnȫ�� Then
'             mrs�ϰ�ʱ���.Filter = "ʱ���='����'"
'            datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!��ʼʱ��, "00:00:00"))
'         End If
'         j = i
'         For i = j To lng��Լ
'            If Format(datʱ��, "dd:mm:ss") >= CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")) Then Exit For
'            lngCol = lngCol + 1
'            If lngCol > vsTime.Cols - 1 Then lngRow = lngRow + 2: lngCol = 1
'            strData = i
'            strTime = Format(datʱ��, "HH:mm") & "-" & _
'                       IIf(DateAdd("n", lng���, datʱ��) > CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), _
'                       Format(CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
'            If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'            If lngRow < 0 Then vsTime.Rows = vsTime.Rows + 2: lngRow = lngRow + 2
'            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
'            vsTime.TextMatrix(lngRow, lngCol) = strTime
'
'            '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
'            If lngCol = 1 Then
'                 vsTime.TextMatrix(lngRow - 1, 0) = Format(datʱ��, "HH:mm")
'                 vsTime.TextMatrix(lngRow, 0) = Format(datʱ��, "HH:mm")
'            End If
'            datʱ�� = DateAdd("n", lng���, datʱ��)
'         Next
         For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
         Next
         .ColWidth(0) = 1200
         .FixedAlignment(0) = flexAlignRightTop
         .ColAlignment(0) = flexAlignRightTop
         If .Rows > 0 Then
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
         End If
         .Redraw = flexRDBuffered
    End With
     
Exit Sub
Hd:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Sub

Private Sub cmdԤԼ_Click()
    Dim intFirstRow As Integer '�����:57488
    Dim intSecondRow As Integer '�����:57488
    '�����:57488
    With vsTime
        intFirstRow = .Row: intSecondRow = intFirstRow + 1
        If .Row Mod 2 = 1 Then
            intFirstRow = .Row - 1
            intSecondRow = intFirstRow + 1
        End If
    End With
    
    '��ʱ����ܷ�ԤԼ��������
    If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Then Exit Sub
    If mViewMode = ViewMode.ViewItem Or vsTime.TextMatrix(vsTime.MouseRow, vsTime.MouseCol) = "" Then Exit Sub
    
    With vsTime
        If .CellForeColor = vbBlue Then
            .Cell(flexcpForeColor, intFirstRow, .Col, intSecondRow, .Col) = &H80000008
            .Cell(flexcpFontBold, intFirstRow, .Col, intSecondRow, .Col) = False
         Else
            .Cell(flexcpForeColor, intFirstRow, .Col, intSecondRow, .Col) = vbBlue
            .Cell(flexcpFontBold, intFirstRow, .Col, intSecondRow, .Col) = True
        End If
    End With
    mblnChange = True
End Sub

Private Sub Form_Activate()
    Me.Icon = frmRegistPlan.Icon
End Sub

Private Sub Form_Load()
    Initʱ���
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  '********************************************
  '�������� �������С��Ⱥ���С�߶�
  '********************************************
  If Me.Width < 701 * Screen.TwipsPerPixelX Then Me.Width = 701 * Screen.TwipsPerPixelX
  If Me.Height < 511 * Screen.TwipsPerPixelY Then Me.Height = 511 * Screen.TwipsPerPixelY
  '********************************************
  '�ҺŰ��Ż�����Ϣ λ�ò��ƶ��ƶ�
  '���ƶ� ʱ������
  '********************************************
  With fraDate
     .Width = Me.ScaleWidth - 2 * .Left
     .Height = Me.ScaleHeight - Me.fraInfo.Top - Me.fraInfo.Height - 65 * Screen.TwipsPerPixelY
  End With
  
  With picTime
     .Width = fraDate.Width - 2 * .Left
     .Height = fraDate.Height - .Top * 2
  End With
  With Me.tbWeekTime
    .Width = picTime.ScaleWidth - 2 * .Left
  End With
  With Me.vsTime
    .Width = picTime.ScaleWidth - 2 * .Left
    .Height = picTime.ScaleHeight - .Top - cmd����ʱ��.Top
  End With
  '-------------------------------------------
  'Ӧ���� λ�õĵ���
  '-------------------------------------------
  With Me.fraӦ����
       .Left = .Left
       .Top = Me.fraDate.Top + Me.fraDate.Height + 5 * Screen.TwipsPerPixelY
   
  End With
  
  '********************************************
  'ȷ����ť��ȡ����ť���ƶ�
  '********************************************
  
  With Me.cmdCancel
       .Left = Me.ScaleWidth - 40 * Screen.TwipsPerPixelX - .Width
       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
  End With
  With Me.cmdOK
       .Left = cmdCancel.Left - 20 * Screen.TwipsPerPixelX - .Width
       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
     mlngPre����ID = -1
     mblnChange = False
     Set mrsTime = Nothing
     mstr�����޸� = ""
     Set mrs�޺� = Nothing
     Set mrs�ϰ�ʱ��� = Nothing
     Set mrs���� = Nothing
End Sub
Private Sub mfrmOtherCalc_zlRefreshCon(ByVal VarTimes As Variant)
    Dim strTemp  As String, varData As Variant, varTemp As Variant
    Dim i As Long, int���� As Integer, dtStart As Date, dtEnd As Date
    Dim lngRow As Long, lng��� As Long, dtTemp As Date, j As Long
    Dim lng�޺��� As Long, lng��Լ�� As Long, str���� As String
    Dim lng�ѹ������� As Long '�����:51427
    Dim lngCol As Long '�����:54127
    Dim lng����ID As Long '�����:54127
    Dim K As Long '�����:54127
    
    If chk��ſ���.Value <> 1 Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    lng����ID = Val(txt�ű�.Tag) '�����:51427
    
    If Get�޺���(str����, lng�޺���, lng��Լ��) = False Then Exit Sub
    
    'VarTiems
    '       "ʱ����"
    '       "�ֶμ��":ʱ��(��:8:00��9:00),2;ʱ��2,���;....
    If VarTimes("ʱ����") <> "" Then
        txtTimeOut.Text = Val(VarTimes("ʱ����"))
        Call cmd����ʱ��_Click
        Exit Sub
    End If
    strTemp = VarTimes("�ֶμ��")
    If strTemp = "" Then Exit Sub
    
    '�����:51427
    lng�ѹ������� = ExistsBooking(lng����ID, str����)
    If lng�ѹ������� <> -1 Then
         If MsgBox("�ð��������б��ҳ�ȥ�ĺ�,ֻ���޸ĺ�ɫ������ʾ��ʱ��" & vbCrLf & "��ȷ��Ҫ�����޸���?", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
         End If
    End If
    
    varData = Split(strTemp, ";")
    lngRow = -2: lng��� = 1: lngCol = 1
    '�����:51427
    For i = 0 To vsTime.Rows - 1
        For j = 0 To vsTime.Cols - 1
            If IsNumeric(vsTime.TextMatrix(i, j)) = True Then
                If CLng(vsTime.TextMatrix(i, j)) = lng�ѹ������� Then
                    lngRow = i: lngCol = j
                End If
            End If
        Next
    Next
    
    '��ʼ��vsTime�ؼ�
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = IIf(lngRow = -2, 2, lngRow + 2): lngRow = IIf(lngRow = 0 And lngCol = 1, -2, lngRow): i = 0: .FixedCols = 1
        .FixedRows = 0
        If lngRow = -2 Then
            .Rows = 0
            .Rows = 2
        End If
    lng��� = IIf(lng�ѹ������� = -1, 1, lng�ѹ������� + 1)
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ",")
        int���� = Val(varTemp(1))
        varTemp = Split(varTemp(0), "��")
        dtStart = CDate(varTemp(0))
        dtEnd = CDate(varTemp(1))
        
        'ͬһʱ�����û�йҳ��ĺ���
        If dtStart = IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) Then
            j = IIf(lngCol = 1 And lngRow = -2, 0, lngCol) + 1
            '���û�ҳ���ѡ��
            For K = j To .Cols - 1
                .TextMatrix(IIf(lngRow = -2, 0, lngRow), K) = ""
                .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, K) = ""
            Next
            .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = Format(dtStart, "HH:00")
            .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, 0) = Format(dtStart, "HH:00")
            If lngCol = 1 Then
                dtStart = .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, 0)
            Else
                dtStart = Split(.TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, lngCol), "-")(1)
            End If
            Do While True
                If j > .Cols - 1 Then .Cols = .Cols + 1
                dtTemp = Format(dtStart + int���� * 1 / 24 / 60, "HH:MM")
                If dtTemp > dtEnd Or lng��� > lng�޺��� Then Exit Do
                .TextMatrix(IIf(lngRow = -2, 0, lngRow), j) = lng���
                .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, j) = Format(dtStart, "HH:MM") & "-" & Format(dtTemp, "HH:MM")
                dtStart = dtTemp: lng��� = lng��� + 1
                j = j + 1
            Loop
        dtStart = "00:00:00"
        End If
        '��ͬʱ���û�б��ҳ��ĺ���
        If dtStart > IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) Then
            If IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) <> Format(dtStart, "HH:00") Then
                 If lng��� > 1 Then
                     lngRow = IIf(lngRow = -2, 0, lngRow)
                 End If
                 lngRow = lngRow + 2
                .Rows = .Rows + 2
                .TextMatrix(lngRow, 0) = Format(dtStart, "HH:00")
                .TextMatrix(lngRow + 1, 0) = Format(dtStart, "HH:00")
            End If
            j = 1
            Do While True
                If j > .Cols - 1 Then .Cols = .Cols + 1
                dtTemp = Format(dtStart + int���� * 1 / 24 / 60, "HH:MM")
                If dtTemp > dtEnd Or lng��� > lng�޺��� Then Exit Do
                .TextMatrix(lngRow, j) = lng���
                .TextMatrix(lngRow + 1, j) = Format(dtStart, "HH:MM") & "-" & Format(dtTemp, "HH:MM")
                dtStart = dtTemp: lng��� = lng��� + 1
                j = j + 1
            Loop
        End If
    Next
    For i = 1 To .Cols - 1
        .ColAlignment(i) = flexAlignCenterCenter
        .ColWidth(i) = 1200
    Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
    If .Rows > 0 Then
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
    End If
    .Redraw = flexRDBuffered
    End With
    Call setVsFlexBgColor(True)
End Sub

Private Sub tbWeekTime_Click()
    Dim i       As Integer
    Dim j As Long '�����:51427
    
    If mstrKey = Mid(tbWeekTime.SelectedItem.Key, 2) Then Exit Sub
    If mblnChange Then
        mblnChange = False
        If MsgBox("��ǰ�ҺŰ�����" & mstrKey & "��ʱ���Ѹı�!�Ƿ񱣴�?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption) = vbYes Then
            cmdOK_Click
'         For i = 1 To tbWeekTime.Tabs.Count
'            If tbWeekTime.Tabs(i).Key = "K" & mstrKey Then
'                tbWeekTime.Tabs(i).Selected = True
'                Exit For
'            End If
'         Next
        End If
    End If
    mstrKey = Mid(tbWeekTime.SelectedItem.Key, 2)
     If mstr�����޸� <> "" Then
        vsTime.Editable = flexEDKbdMouse: cmd����ʱ��.Enabled = True
        If InStr(mstr�����޸�, ";" & mstrKey & ";") > 0 Then vsTime.Editable = flexEDNone: cmd����ʱ��.Enabled = False
    End If
    Select Case mViewMode
        Case ViewMode.ViewItem:
             Call LoadTimePlan(mlng����ID, Me.chk��ſ���.Value = 1)
        Case ViewMode.Edit:
            cmdԤԼ.Visible = False
            cmdɾ��.Visible = False
            Call LoadEditTimePlan(mlng����ID, Me.chk��ſ���.Value = 1)
    End Select
     setVsFlexBgColor (Me.chk��ſ���.Value = 1)
End Sub


 

Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
   
    '���Ʒ���������
    If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    If txtTimeOut.Text = "" And KeyAscii = Asc(0) Then KeyAscii = 0
End Sub

Private Sub txtTimeOut_Validate(Cancel As Boolean)
    If Val(txtTimeOut.Text) < 1 Then Cancel = True
End Sub

 

Private Sub udTime_DownClick()
    If Val(txtTimeOut.Text) < 2 Then Exit Sub
    txtTimeOut.Text = Val(txtTimeOut.Text) - 1
End Sub

Private Sub udTime_UpClick()
  txtTimeOut.Text = Val(txtTimeOut.Text) + 1
End Sub


 
 
'Private Sub vsTime_Click()
'  Select Case mViewMode
'    Case ViewMode.Edit, ViewMode.NewItem:
'       If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Or (chk��ſ���.Value = 0 And vsTime.MouseRow < 1) Then Exit Sub
'       Select Case chk��ſ���.Value = 1
'            Case True:
'            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'            Case False:
'            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'       End Select
'        If vsTime.MouseRow < 0 Or vsTime.MouseCol < 1 Then Exit Sub
'
'        If chk��ſ���.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
'            cmdԤԼ.Left = vsTime.MouseCol * 1200 + 20
'            cmdԤԼ.Top = vsTime.MouseRow * 400 + 20
'            cmdԤԼ.Visible = True
'        End If
'
'    Case ViewMode.ViewItem:
'         vsTime.Editable = flexEDNone
'  End Select
'End Sub

Public Function ShowMe(lng����ID As Long, mode As ViewMode) As Boolean
    mViewMode = mode: mlng����ID = lng����ID
    If InitData() = False Then
        '���عҺŰ��Ż�����Ϣ
         Exit Function
    End If
    Select Case mViewMode
         Case ViewMode.ViewItem:
                vsTime.Editable = flexEDNone
                Me.txtTimeOut.Enabled = False
                Me.cmd����ʱ��.Enabled = False
               '�鿴
              Call LoadTimePlan(mlng����ID, chk��ſ���.Value = 1, False)
         Case ViewMode.Edit
              If LoadEditTimePlan(mlng����ID, chk��ſ���.Value = 1, False) = False Then
               Exit Function
              End If
    End Select
    setVsFlexBgColor (chk��ſ���.Value = 1)
    Me.Show 1
    ShowMe = mblnReload
End Function
'------------------------------------------------------------------------
'ҳ����ù����뷽��
'------------------------------------------------------------------------
Public Function InitData() As Boolean
    Dim strSQL          As String
    Dim lng����ID       As Long
    If mlng����ID = -1 Then Exit Function
     lng����ID = mlng����ID
     On Error GoTo Hd
     strSQL = " " & _
        "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id," & _
        "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,nvl(A.Ĭ��ʱ�μ��,5) As Ĭ��ʱ�μ��, " & _
        "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & _
        "   From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� D " & _
        "   Where A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
        "         And A.Id=[1]"
         Set mrs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
         
         If mrs����.EOF Then
              ShowMsgbox "δ�ҵ�ָ���ĺű�,����!"
             Exit Function
        End If
        strSQL = "Select ������Ŀ,�޺���,  ��Լ��,������Ŀ as ���� From  �ҺŰ������� where ����ID=[1]  Order BY ������Ŀ      "
        Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        cbo����.Text = Nvl(mrs����!����)
        txt�ű�.Tag = Nvl(mrs����!����ID)
        txtTimeOut.Tag = Val(Nvl(mrs����!Ĭ��ʱ�μ��, 0))
        txtTimeOut.Text = txtTimeOut.Tag
        txt�ű�.Text = Nvl(mrs����!����)
        cbo����.Text = Nvl(mrs����!����)
        cboItem.Text = Nvl(mrs����!��Ŀ)
        cboDoctor.Text = Nvl(mrs����!ҽ������)
        chk����.Value = IIf(Val(Nvl(mrs����!��������)) = 1, 1, 0)
        chk��ſ���.Value = IIf(Val(Nvl(mrs����!��ſ���)) = 1, 1, -0):  chk��ſ���.Tag = chk��ſ���.Value
        '�����:51429
        Call chk��ſ���_Click
        strSQL = "" & _
        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
        "               ��������,�Ƿ�ԤԼ" & _
        "   From  �ҺŰ���ʱ�� " & _
        "   Where ����ID=[1]" & _
        "   Order by ����,ʱ��,���"
        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        mstr�����޸� = Get��Լ����(mlng����ID)
       InitData = True
Exit Function
Hd:
     If ErrCenter() = 1 Then Resume
     SaveErrLog
End Function

 
Private Function LoadEditTimePlan(ByVal lng����ID As Long, ByVal bln��ſ��� As Boolean, _
    Optional bln�ƻ� As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str����          As String
    Dim i                As Long
    Dim j                As Long
    Dim r                As Integer
    Dim lngRow           As Long
    Dim lngCol           As Integer
    Dim strʱ��          As String
    Dim strTime          As String
    Dim strData          As String
    Dim strKey           As String
    Dim lng�ѹ������� As Long '�����:51427
     
    On Error GoTo errHandle
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    If mrsTime Is Nothing Then
        mlngPre����ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre����ID = -1
    End If
    If mlngPre����ID <> lng����ID Then
        mlngPre����ID = lng����ID
        tbWeekTime.Tabs.Clear
        With tbWeekTime
            If Not mrs�޺�.EOF Then
                mrs�޺�.Filter = "����='��һ'"
                If mrs�޺�.RecordCount > 0 Then
                '�޺���,  ��Լ��,������Ŀ
                    If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K��һ", "��һ" & IIf(Nvl(mrs����!��һ) = "", "", "(" & Nvl(mrs����!��һ) & ")")
                    End If
                End If
                mrs�޺�.Filter = "����='�ܶ�'"
                If mrs�޺�.RecordCount > 0 Then
                   If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                        "K�ܶ�", "�ܶ�" & IIf(Nvl(mrs����!�ܶ�) = "", "", "(" & Nvl(mrs����!�ܶ�) & ")")
                    End If
                End If
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                     If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                        "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
                    End If
                 End If
                 
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                  If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                      "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
                  End If
                End If
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                     If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
                     End If
                End If
                
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                   If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                          "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
                   End If
                End If
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                    If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
                    End If
                End If
                mrs�޺�.Filter = 0
            End If
            .Visible = tbWeekTime.Tabs.Count <> 0
            If .Tabs.Count > 0 Then
                .Tabs(1).Selected = True
            Else
                MsgBox "�ð���û�����ö�Ӧ���޺�������Լ��,����!", vbOKOnly, Me.Caption
                Exit Function
            End If
           
        End With
    End If
    str���� = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "����='" & str���� & "'"
    mrs�޺�.Filter = "����='" & str���� & "'"
    txt�޺�.Text = ""
    txt��Լ.Text = ""
    If mrs�޺�.RecordCount <> 0 Then
        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
    End If
     strʱ�� = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln��ſ��� Then
             .Cols = 8: .FixedCols = 0
             .Rows = 1: .FixedRows = 1
             For i = 0 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "ʱ���"
             Next
             For i = 1 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "ԤԼ����"
             Next
             
             r = 1: i = -1
            Do While Not mrsTime.EOF
                i = i + 1
                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
                strData = Val(Nvl(mrsTime!��������))
                strTime = mrsTime!ʱ�䷶Χ
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
            LoadEditTimePlan = True
            Exit Function
        End If
        .Cols = 7: .FixedCols = 1
        .Rows = 0: .FixedRows = 0
        i = 1: r = -1
        lngRow = -1: lngCol = 1
        '******************************************
        With vsTime
         .Cols = 2
         lngRow = -1: lngCol = 0
         '***********************
         '������
         '**********************
         r = mrsTime.RecordCount
         For i = 1 To r
            If mrsTime.EOF Then Exit For
            lngCol = lngCol + 1
            If strʱ�� <> Nvl(mrsTime!ʱ��) Then lngRow = lngRow + 2: lngCol = 1
             If lngCol = 1 Then
                strʱ�� = Nvl(mrsTime!ʱ��)
                If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
                vsTime.TextMatrix(lngRow - 1, 0) = strʱ��
                vsTime.TextMatrix(lngRow, 0) = strʱ��
             End If
            strData = mrsTime!���
            strTime = mrsTime!ʱ�䷶Χ
            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
            'If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
            vsTime.TextMatrix(lngRow, lngCol) = strTime
            '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
            If lngCol = 1 Then
            End If
            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
            End If
            mrsTime.MoveNext
         Next
         
         End With
        '******************************************
'        Do While Not mrsTime.EOF
'            If i = 1 Then
'                r = r + 2
'                strʱ�� = Nvl(mrsTime!ʱ��)
'                If r > .Rows - 1 Then .Rows = .Rows + 2
'                .TextMatrix(r, 0) = strʱ��
'                .TextMatrix(r - 1, 0) = strʱ��
'            End If
'            i = i + 1
'            strData = mrsTime!���
'            strTime = mrsTime!ʱ�䷶Χ
'            If i >= .Cols - 1 Then i = 1
'            If r > .Rows - 1 Then .Rows = .Rows + 2
'            .TextMatrix(r, i) = strTime
'            .TextMatrix(r - 1, i) = strData
'
'        Loop
        
        
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        If .Rows > 0 Then
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        End If
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = flexRDBuffered
    End With
    
    '���ò���ɾ������ɫ
    '�����:51427
    lng�ѹ������� = ExistsBooking(mlng����ID, Mid(tbWeekTime.SelectedItem.Key, 2))
    If chk��ſ���.Value = 1 Then
        '�����:51427
        For i = 0 To vsTime.Rows - 1
            For j = 0 To vsTime.Cols - 1
                If IsNumeric(vsTime.TextMatrix(i, j)) = True Then
                    If CLng(vsTime.TextMatrix(i, j)) <= lng�ѹ������� Then
                        vsTime.Cell(flexcpForeColor, i, j) = &HC0C0C0
                        vsTime.Cell(flexcpForeColor, i + 1, j) = &HC0C0C0
                    End If
                End If
            Next
        Next
    End If
    
    LoadEditTimePlan = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
 
 
Private Sub LoadEditTimePlantext(ByVal lng����ID As Long, ByVal bln��ſ��� As Boolean, _
    Optional bln�ƻ� As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str����          As String
    Dim i                As Long
    Dim r                As Integer
    Dim strʱ��          As String
    Dim strTime          As String
    Dim strData          As String
    Dim strKey           As String
     
    On Error GoTo errHandle
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    If mrsTime Is Nothing Then
        mlngPre����ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre����ID = -1
    End If
    If mlngPre����ID <> lng����ID Then
        mlngPre����ID = lng����ID
        tbWeekTime.Tabs.Clear
        With mrsTime
            strTime = ""
            Do While Not .EOF
                If strTime <> Nvl(mrsTime!����) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!����), Nvl(mrsTime!����)
                    strTime = Nvl(mrsTime!����)
                End If
                .MoveNext
            Loop
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
        End With
    End If
    str���� = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "����='" & str���� & "'"
    mrs�޺�.Filter = "����='" & str���� & "'"
    txt�޺�.Text = ""
    txt��Լ.Text = ""
    If mrs�޺�.RecordCount <> 0 Then
        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
    End If
     strʱ�� = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln��ſ��� Then
             .Cols = 8: .FixedCols = 0
             .Rows = 1: .FixedRows = 1
             For i = 0 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "ʱ���"
             Next
             For i = 1 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "ԤԼ����"
             Next
             
             r = 1: i = -1
            Do While Not mrsTime.EOF
                If i * 2 > .Cols - 2 Then r = r + 1: i = -1
                i = i + 1
                strData = Val(Nvl(mrsTime!��������))
                strTime = mrsTime!ʱ�䷶Χ
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
             Exit Sub
        End If
        Do While Not mrsTime.EOF
            If strʱ�� <> Nvl(mrsTime!ʱ��) Then
                r = r + 2
                strʱ�� = Nvl(mrsTime!ʱ��)
                If r > .Rows - 1 Then .Rows = .Rows + 2
                .TextMatrix(r, 0) = strʱ��
                .TextMatrix(r - 1, 0) = strʱ��
                i = 0
            End If
            i = i + 1
            strData = mrsTime!���
            strTime = mrsTime!ʱ�䷶Χ
            If i > .Cols - 1 Then .Cols = .Cols + 1
            If r > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(r, i) = strTime
            .TextMatrix(r - 1, i) = strData
            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
                 
                .Cell(flexcpForeColor, r - 1, i, r, i) = vbBlue
                .Cell(flexcpFontBold, r - 1, i, r, i) = True
            End If
            mrsTime.MoveNext
        Loop
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
 
Private Sub LoadTimePlan(ByVal lng����ID As Long, ByVal bln��ſ��� As Boolean, _
    Optional bln�ƻ� As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str����          As String
    Dim i                As Long
    Dim r                As Integer
    Dim strʱ��          As String
    Dim strTime          As String
    Dim strKey           As String
    On Error GoTo errHandle
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    If mrsTime Is Nothing Then
         mlngPre����ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre����ID = -1
    End If
    If mlngPre����ID <> lng����ID Then
        mlngPre����ID = lng����ID
'        strSQL = "" & _
'        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
'        "               ��������,�Ƿ�ԤԼ" & _
'        "   From  �ҺŰ���ʱ�� " & _
'        "   Where ����ID=[1]" & _
'        "   Order by ����,ʱ��,���"
'        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        tbWeekTime.Tabs.Clear
        With mrsTime
            strTime = ""
            Do While Not .EOF
                If strTime <> Nvl(mrsTime!����) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!����), Nvl(mrsTime!����)
                    strTime = Nvl(mrsTime!����)
                End If
                .MoveNext
            Loop
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
        End With
        If tbWeekTime.Tabs.Count = 0 Then
            MsgBox "�ð���û�����ö�Ӧ��ʱ��,����!"
        End If
    End If
    str���� = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "����='" & str���� & "'"
    mrs�޺�.Filter = "����='" & str���� & "'"
    txt�޺�.Text = ""
    txt��Լ.Text = ""
    If mrs�޺�.RecordCount <> 0 Then
        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
    End If
     strʱ�� = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 800: .RowHeightMin = 800
        .Rows = 1: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln��ſ��� Then
             .Cols = 8: .FixedCols = 0
             r = 0: i = 0
            Do While Not mrsTime.EOF
                i = i + 1
                If i > .Cols - 1 Then r = r + 1: i = 0
                strTime = "ԤԼ" & Val(Nvl(mrsTime!��������)) & "��" & vbCrLf & vbCrLf
                strTime = strTime & mrsTime!ʱ�䷶Χ
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i) = strTime
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
             Exit Sub
        End If
        Do While Not mrsTime.EOF
            If strʱ�� <> Nvl(mrsTime!ʱ��) Then
                r = r + 1
                strʱ�� = Nvl(mrsTime!ʱ��)
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, 0) = strʱ��
                i = 0
            End If
            i = i + 1
            strTime = mrsTime!��� & vbCrLf & vbCrLf
            strTime = strTime & mrsTime!ʱ�䷶Χ
            If i > .Cols - 1 Then .Cols = .Cols + 1
            If r > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(r, i) = strTime
            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
                .Cell(flexcpForeColor, r, i, r, i) = vbBlue
                .Cell(flexcpFontBold, r, i, r, i) = True
            End If
            mrsTime.MoveNext
        Loop
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
    
Private Sub vsTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
 If vsTime.Row < 0 Or vsTime.Col < 0 Or (chk��ſ���.Value = 0 And vsTime.Row < 1) Then cmdԤԼ.Visible = False: mblnCellChange = False: Exit Sub
 
    Call SetCtrlMove
 
    Select Case mViewMode
    Case ViewMode.Edit, ViewMode.NewItem:
       Select Case chk��ſ���.Value = 1
            Case True:
            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
            '******************************************
            '�������������ʽ
            '******************************************
            If vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
            Case False:
            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
            '******************************************
            '�������������ʽ
            '******************************************
            If NewCol Mod 2 = 0 And vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
       End Select
        If vsTime.Row < 0 Or vsTime.Col < 1 Then Exit Sub
        
        If chk��ſ���.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
            mblnCellChange = True
        Else
           mblnCellChange = False
        End If
        
    Case ViewMode.ViewItem:
         mblnCellChange = False
         vsTime.Editable = flexEDNone
  End Select
   If mstr�����޸� <> "" Then
        vsTime.Editable = flexEDKbdMouse
        If InStr(mstr�����޸�, ";" & mstrKey & ";") > 0 Then vsTime.Editable = flexEDNone
        
   End If
End Sub



Private Sub vsTime_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����:51429
        If cmdɾ��.Visible = False Then Exit Sub
        If KeyCode = 46 Then '��ݼ�Delete
            Call DeleteSelectPain
        End If
End Sub
Private Sub SetCtrlMove()
    Dim blnDel As Boolean
    With vsTime
        If chk��ſ���.Value = 1 Then
            cmdɾ��.Left = .CellLeft + .CellWidth - cmdɾ��.Width
            If .Row Mod 2 <> 0 Then
                cmdɾ��.Top = .CellTop - .CellHeight - 15
            Else
                cmdɾ��.Top = .CellTop + 15
            End If
            cmdԤԼ.Left = .CellLeft + 15
            cmdԤԼ.Top = cmdɾ��.Top
            If .Col < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(.Row, .Col + 1)) = ""
            Else
                blnDel = True
            End If
            blnDel = blnDel And Trim(.TextMatrix(.Row, .Col)) <> ""
            cmdɾ��.Visible = blnDel And chk��ſ���.Value = 1
            cmdԤԼ.Visible = Val(txt��Լ.Text) <> 0 And InStr(mstr�����޸�, mstrKey) = 0
        Else
            cmdԤԼ.Left = .CellLeft + 15
            cmdԤԼ.Top = .CellTop + 15
            cmdԤԼ.Visible = False ' Val(txt��Լ.Text) <> 0 And InStr(mstr�����޸�, mstrKey) = 0
        End If
    End With
End Sub


Private Sub vsTime_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    '**************************************************************
    '������Ա �϶�������ʱ �� ԤԼ��ť ����
    '**************************************************************
    Me.cmdԤԼ.Visible = False
     '�����:51429
    Me.cmdɾ��.Visible = False
End Sub
Private Sub vsTime_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If mViewMode = ViewItem Then Exit Sub
    Select Case chk��ſ���.Value = 1
        Case True:
            '******************************************
            'ר�Һ�ʱ ��������
            '******************************************
            If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
        Case False:
            '******************************************
            '��ͨ��ʱ ��������
            '******************************************
            If Col Mod 2 = 0 Then
                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
            Else
                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13) Then KeyAscii = 0: Exit Sub
            End If
            
    End Select
   
 
End Sub
 
Private Function validateVsFlex() As Boolean
    '***************************************
    '��֤�û��ԹҺŰ���ʱ�ε��޸�
    '***************************************
     Dim i          As Long
     Dim j          As Long
     Dim lngԤԼ    As Long
     Dim lng��Լ    As Long
     Dim lng�޺�    As Long
     Dim str����    As String
     If tbWeekTime.SelectedItem Is Nothing Then Exit Function
      str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
     lng�޺� = Val(txt�޺�.Text)
     lng��Լ = Val(txt��Լ.Text)
     If lng��Լ = 0 Then lng��Լ = lng�޺�
     Select Case chk��ſ���.Value = 1
     Case True:
     '*************************************
     'ר�Һż����Լ���Ƿ�����޺���
     '*************************************
        With vsTime
            For i = 0 To .Rows - 1 Step 2
                For j = 1 To .Cols - 1
                    If .Cell(flexcpForeColor, i, j, i, j) = vbBlue And .TextMatrix(i, j) <> "" Then
                        lngԤԼ = lngԤԼ + 1
                    End If
                Next
            Next
        End With
     Case False:
     '*************************************
     '��ͨ�ż����Լ���Ƿ�����޺���
     '*************************************
        With vsTime
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) <> "" Then
                        lngԤԼ = lngԤԼ + Val(.TextMatrix(i, j))
                    End If
                Next
            Next
        End With
     End Select
     If lngԤԼ > lng��Լ Then
        MsgBox "��" & str���� & "���õ�ԤԼ��" & lngԤԼ & "������" & IIf(lng�޺� = lng��Լ, "�޺���" & lng��Լ, "��Լ��" & lng��Լ) & ",����!", vbOKOnly, Me.Caption
        Exit Function
     End If
    validateVsFlex = True
    Exit Function
End Function

Private Function SaveDate() As Boolean
    '*********************************
    '�ԹҺŰ���ʱ�ν��б���
    '*********************************
    Dim strSQL      As String
    Dim cllSQL      As Collection
    Dim i           As Long
    Dim j           As Long
    Dim blnTrans    As Boolean
    Dim lng����ID   As Long
    Dim str����     As String
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim blnԤԼ     As Boolean
    Dim lng����     As Long '�ҺŰ���ʱ�ε���������
    Dim blnר�Һ�   As Boolean
    Dim lng���     As Long
    Dim lngType     As Long
    Dim lng��Ч�ƻ�ID As Long '�����:51427
    Dim cll�ƻ�SQL As Collection '�����:51427
    
    If validateVsFlex() = False Then Exit Function '�������ݵ���֤
    
    
    lng����ID = Val(txt�ű�.Tag)
    str���� = mstrKey
    blnר�Һ� = chk��ſ���.Value = 1
    
    Set cllSQL = New Collection
    '****************************************************
    'CREATE OR REPLACE Procedure Zl_�ҺŰ���ʱ��_Delete(
    '����id_In �ҺŰ���ʱ��.����id%Type,
    '����_In   �ҺŰ���ʱ��.���� %Type)
    '**********ɾ����ǰ�Դ����ڰ��ŵ�ʱ��*****************

    strSQL = "Zl_�ҺŰ���ʱ��_Delete(" & lng����ID & ",'" & str���� & "')"
    zlAddArray cllSQL, strSQL
    
   
    Select Case blnר�Һ�
    Case True:
       lng��� = 0
       For i = 1 To vsTime.Rows - 1 Step 2
            For j = 1 To vsTime.Cols - 1
               If vsTime.TextMatrix(i, j) = "" Then Exit For
               str��ʼʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
               str����ʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
               lng���� = 1
               lng��� = lng��� + 1
               blnԤԼ = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
               strSQL = GetInsertSql(lng����ID, lng���, str��ʼʱ��, str����ʱ��, 1, blnԤԼ, str����)
               zlAddArray cllSQL, strSQL
            Next
       Next
    Case False:
        lng��� = 0
        For i = 1 To vsTime.Rows - 1
            For j = 0 To vsTime.Cols - 1 Step 2
               If vsTime.TextMatrix(i, j) <> "" Then
                str��ʼʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
                str����ʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
                lng���� = Val(vsTime.TextMatrix(i, j + 1))
                lng��� = lng��� + 1
                blnԤԼ = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
                strSQL = GetInsertSql(lng����ID, lng���, str��ʼʱ��, str����ʱ��, lng����, blnԤԼ, str����)
                zlAddArray cllSQL, strSQL
               End If
            Next
        Next
    End Select
     
    If opt��ҽ��.Value Then
        lngType = 1
    ElseIf opt����.Value Then
        lngType = 2
    ElseIf opt����.Value Then
        lngType = 3
    End If
    If lngType <> 0 Then
        '--type_in
        '--1-Ӧ���뱾��
        '--2-Ӧ���뱾����
        '--3 or others -Ӧ��������
       'CREATE OR REPLACE Procedure zl_�ҺŰ���ʱ��_����Ӧ��
       strSQL = "zl_�ҺŰ���ʱ��_����Ӧ��("
       '����Id_in �ҺŰ���ʱ��.����Id%Type,
       strSQL = strSQL & lng����ID & ","
       'Type_In Number:=1
       strSQL = strSQL & lngType & ")"
       zlAddArray cllSQL, strSQL
    End If
     
  On Error GoTo Errhand
    gcnOracle.BeginTrans
    
    For i = 1 To cllSQL.Count
        strSQL = cllSQL(i)
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
 '�����:51427
' '���¹Һżƻ�ʱ����Ϣ
'    lng��Ч�ƻ�ID = �Ƿ�����Ч�ƻ�(lng����ID, zlDatabase.Currentdate)
'
'    If lng��Ч�ƻ�ID > 0 Then '������Ч�ļƻ�������Ҫͬ�����¹Һżƻ�ʱ�ε���Ϣ
'        Set cll�ƻ�SQL = New Collection
'        strSQL = "Zl_�Һżƻ�ʱ��_Delete(" & lng��Ч�ƻ�ID & ",'" & str���� & "')"
'        zlAddArray cll�ƻ�SQL, strSQL
'
'        Select Case blnר�Һ�
'        Case True:
'           lng��� = 0
'           For i = 1 To vsTime.Rows - 1 Step 2
'                For j = 1 To vsTime.Cols - 1
'                   If vsTime.TextMatrix(i, j) = "" Then Exit For
'                   str��ʼʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
'                   str����ʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
'                   lng���� = 1
'                   lng��� = lng��� + 1
'                   blnԤԼ = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
'                   strSQL = GetInsert�ƻ�Sql(lng��Ч�ƻ�ID, lng���, str��ʼʱ��, str����ʱ��, 1, blnԤԼ, str����)
'                   zlAddArray cll�ƻ�SQL, strSQL
'                Next
'           Next
'        Case False:
'            lng��� = 0
'            For i = 1 To vsTime.Rows - 1
'                For j = 0 To vsTime.Cols - 1 Step 2
'                   If vsTime.TextMatrix(i, j) <> "" Then
'                    str��ʼʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
'                    str����ʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
'                    lng���� = Val(vsTime.TextMatrix(i, j + 1))
'                    lng��� = lng��� + 1
'                    blnԤԼ = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
'                    strSQL = GetInsert�ƻ�Sql(lng��Ч�ƻ�ID, lng���, str��ʼʱ��, str����ʱ��, lng����, blnԤԼ, str����)
'                    zlAddArray cll�ƻ�SQL, strSQL
'                   End If
'                Next
'            Next
'        End Select
'        For i = 1 To cll�ƻ�SQL.Count
'            strSQL = cll�ƻ�SQL(i)
'            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
'        Next
'    End If
    gcnOracle.CommitTrans
    SaveDate = True
 Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    Call ErrCenter
    SaveErrLog
End Function

Private Function GetInsertSql(ByVal lngID As Long, ByVal lng��� As Long, ByVal str��ʼʱ�� As String, _
        ByVal str����ʱ�� As String, ByVal lng�������� As Long, ByVal bln�Ƿ�ԤԼ As Boolean, ByVal str���� As String)
    '�����ṩ����Ϣ����sql���
    Dim strSQL      As String
   '********************************************************
    '    'CREATE OR REPLACE Procedure Zl_�ҺŰ���ʱ��_Insert
    '    (
    '    ����id_In   �ҺŰ���ʱ��.����id%Type,
    '    ���_In     �ҺŰ���ʱ��.���%Type,
    '    ��ʼʱ��_In �ҺŰ���ʱ��.��ʼʱ��%Type,
    '    ����ʱ��_In �ҺŰ���ʱ��.����ʱ��%Type,
    '    ��������_In �ҺŰ���ʱ��.��������%Type,
    '    �Ƿ�ԤԼ_In �ҺŰ���ʱ��.�Ƿ�ԤԼ%Type,
    '    ����_In     �ҺŰ���ʱ��.����%Type
    '    )
    '********************************************************
    strSQL = "  Zl_�ҺŰ���ʱ��_Insert("
     '����id_In   �ҺŰ���ʱ��.����id%Type,
    strSQL = strSQL & lngID & ","
     '���_In     �ҺŰ���ʱ��.���%Type,
    strSQL = strSQL & lng��� & ","
     '��ʼʱ��_In �ҺŰ���ʱ��.��ʼʱ��%Type,
     strSQL = strSQL & str��ʼʱ�� & ","
      '����ʱ��_In �ҺŰ���ʱ��.����ʱ��%Type,
    strSQL = strSQL & str����ʱ�� & ","
      '��������_In �ҺŰ���ʱ��.��������%Type,
    strSQL = strSQL & lng�������� & ","
     '�Ƿ�ԤԼ_In �ҺŰ���ʱ��.�Ƿ�ԤԼ%Type,
    strSQL = strSQL & IIf(bln�Ƿ�ԤԼ, 1, 0) & ","
     '����_In     �ҺŰ���ʱ��.����%Type
    strSQL = strSQL & "'" & str���� & "')"
    GetInsertSql = strSQL
End Function

                             

Private Function ConvertToDate(ByVal strDate As String, Optional ByVal haveYear = False) As String
    '**********************************************************
    '���ַ���ת����oracle���ݿ��ܹ�ʶ�������
    '**********************************************************
    Select Case haveYear
    Case True:
        ConvertToDate = "To_Date('" & strDate & "', 'YYYY-MM-DD HH24:MI:SS')"
    Case False:
        ConvertToDate = "To_Date('" & strDate & "', 'HH24:MI:SS')"
    End Select
End Function



Private Sub vsTime_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim i         As Long
  Dim j         As Long
  Dim lng�޺�   As Long
  Dim lng��Լ   As Long
  Dim lngԤԼ�� As Long
  If mViewMode = ViewItem Then Exit Sub

  '*************************************
  'ʱ�������֤ ������ʱ�䷶Χ
  '**************************************
  If vsTime.Editable = flexEDKbdMouse And vsTime.ColEditMask(vsTime.Col) = strMaskKey Then
    Validateʱ�� Row, Col, Cancel
    If Not Cancel Then mblnChange = True
    Exit Sub
  End If
  '****************************************
  '����ͨ�� ��ʱ�� �����������ԤԼ����������
  '****************************************
   If chk��ſ���.Value = 0 And vsTime.ColEditMask(vsTime.Col) <> strMaskKey And vsTime.Editable = flexEDKbdMouse Then
        If vsTime.EditText = "" Then vsTime.EditText = "0"
        mblnChange = True
   End If
End Sub

Private Sub Validateʱ��(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim i         As Long
  Dim j         As Long
  Dim lng�޺�   As Long
  Dim lng��Լ   As Long
  Dim lngԤԼ�� As Long
   
  Dim strʱ��()  As String
  If mViewMode = ViewItem Then Exit Sub
  
  '*************************************
  '��֤ʱ��
  '**************************************
  strʱ�� = Split(vsTime.EditText, "-")
  If UBound(strʱ��) <> 1 Then Cancel = True: Exit Sub
   If Not IsDate(strʱ��(0)) Then Cancel = True: Exit Sub
   If Not IsDate(strʱ��(1)) Then Cancel = True: Exit Sub
   If CDate(strʱ��(0)) >= CDate(strʱ��(1)) Then
        MsgBox "��ʼʱ�����С�ڽ���ʱ��!����!", vbOKOnly, Me.Caption
        Cancel = True
   End If
   
End Sub

Private Sub setVsFlexBgColor(Optional ByVal bln��ſ��� As Boolean = False)
    '**************************************************************
    '��ʱ������ü������
    '**************************************************************
     Dim i           As Long
     If (bln��ſ��� And vsTime.Rows = 0) Or (bln��ſ��� = False And vsTime.Rows = 1) Then Exit Sub
     For i = IIf(bln��ſ���, 0, 1) To vsTime.Rows - 1 Step 2
            vsTime.Cell(flexcpBackColor, i, IIf(bln��ſ���, 1, 0), i, vsTime.Cols - 1) = &HE0E0D3
     Next
End Sub
 



Private Sub Initʱ���()
  '--------------------------------
  '����:��ȡ���°�ʱ���
  '--------------------------------
    Dim strTmp      As String
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strDat      As String
    On Error GoTo Hd
    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "07:00:00 AND 12:00:00")
    strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����ϰ� = CDate("08:00:00")
    End If
   
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����°� = CDate("1900-01-01 12:00:00")
    End If
    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "14:00:00 AND 18:00:00")
    
     strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 14:00:00")
    End If
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����°� = CDate("1900-01-01 18:00:00")
    End If
    With t_ʱ��
         If .dat_�����ϰ� > .dat_�����°� Then
            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
         End If
         If .dat_�����ϰ� > .dat_�����°� Then
            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
         End If
    End With
    strSQL = _
    "       Select ʱ���, �ϰ�, �°� " & vbNewLine & _
    "       From (" & vbNewLine & _
    "           With Tb As (Select ʱ���,To_Date('1900-01-01 ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ʼʱ��," & vbNewLine & _
    "                               To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ֹʱ��," & _
    "                               Sign(��ʼʱ�� - ��ֹʱ��) As ����, " & vbNewLine & _
    "                                To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��, " & vbNewLine & _
    "                                To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��, " & vbNewLine & _
    "                                 To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��," & vbNewLine & _
    "                                 To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��"
    strSQL = strSQL & vbNewLine & _
    "                       From ʱ��� )" & vbNewLine & _
    "           Select ʱ���, '��' As ��ǩ, 0 As ��־, ��ʼʱ�� As �ϰ�, ��ֹʱ�� As �°�, ��ʼʱ��, ��ֹʱ��," & _
    "                  �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ��" & vbNewLine & _
    "            From Tb  Where (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) And " & _
    "                      (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
    "           Union All" & vbNewLine & _
    "           Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & vbNewLine & _
    "                        Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) As �°�, ��ʼʱ��, ��ֹʱ��, " & _
    "                        �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
    "           From Tb a Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
    "           Union All " & vbNewLine & _
    "            Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & _
    "                   Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) As �°�, ��ʼʱ��, ��ֹʱ��, �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
    "         From Tb a   Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��)" & vbNewLine & _
    "            ) b" & vbNewLine & _
    "         Order By ʱ���,�ϰ�"
     Set mrs�ϰ�ʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
Private Function Get��Լ����(ByVal lng����ID As Long) As String
    '��ȡ�����޸ĵİ�������
    Dim strSQL As String
    Dim rsTmp   As ADODB.Recordset
    Dim strTmp  As String
    strSQL = "Select Decode(To_Char(A.ԤԼʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7'," & _
    "                             '����') As ���� " & vbCrLf & _
    "          From ���˹Һż�¼ A,�ҺŰ��š�B " & vbCrLf & _
    "        Where  A.�ű�=B.���� And A.��¼״̬ = 1 And b.ID = [1] And A.����ʱ�� > A.�Ǽ�ʱ�� And A.ԤԼʱ�� Is Not Null"
    
    If gintԤԼ���� = 0 Then
        strSQL = strSQL & " And A.ԤԼʱ�� > Sysdate "
    Else
        strSQL = strSQL & " And A.ԤԼʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTmp.EOF Then Exit Function
    
    Do While Not rsTmp.EOF
        If InStr(strTmp, Nvl(rsTmp!����)) < 0 Or strTmp = "" Then
            strTmp = strTmp & ";" & Nvl(rsTmp!����)
        End If
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then strTmp = strTmp & ";"
    Get��Լ���� = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function Get�޺���(ByVal str���� As String, ByRef lng�޺��� As Long, ByRef lng��Լ�� As Long) As Boolean
    Dim strSQL As String
    If mrs�޺� Is Nothing Then
        strSQL = _
        "Select ����id, ������Ŀ as ���� , �޺���, ��Լ�� From �ҺŰ������� Where ����id = [1]"
        Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(txt�ű�.Tag))
        If mrs�޺�.RecordCount = 0 Then
            MsgBox "��ǰ�ű�û�ж�Ӧ�ĹҺŰ�������" & vbCrLf & "�뵽�ҺŰ���������!", vbOKOnly, Me.Caption
            Set mrs�޺� = Nothing
            Exit Function
        End If
    End If
    mrs�޺�.Filter = "����='" & str���� & "'"
    If mrs�޺�.RecordCount <> 0 Then
        lng�޺��� = Val(Nvl(mrs�޺�!�޺���))
        lng��Լ�� = Val(Nvl(mrs�޺�!��Լ��))
        Get�޺��� = True
    End If
End Function
'
Private Function �Ƿ�����Ч�ƻ�(lng����ID As Long, dat���� As Date) As Long
    '**************************************************************
    '���ð������Ƿ�����Ч�ļƻ�
    '������lng����ID-����ID��dat����-��ǰ����
    '����ֵ:�з�����Ч�ƻ�ID,�޷���-1
    '**************************************************************
    Dim strSQL As String
    Dim rs��Ч�ƻ� As Recordset
    On Error GoTo errH
     strSQL = "" & _
     "      Select A.ID From �ҺŰ��żƻ� A" & _
     "      Where A.����ID=[1] And [2] between Nvl(A.��Чʱ��, [2]) And A.ʧЧʱ�� And A.���ʱ�� is Not Null" & _
     "      order By A.��Чʱ�� Desc"
    Set rs��Ч�ƻ� = zlDatabase.OpenSQLRecord(strSQL, "���ָ���Ű����Ƿ�����Ч�ƻ�", lng����ID, dat����)
    
    If rs��Ч�ƻ� Is Nothing Then �Ƿ�����Ч�ƻ� = -1: Exit Function
    If rs��Ч�ƻ�.RecordCount = 0 Then �Ƿ�����Ч�ƻ� = -1: Exit Function
    rs��Ч�ƻ�.MoveFirst
    �Ƿ�����Ч�ƻ� = rs��Ч�ƻ�!ID
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function GetInsert�ƻ�Sql(ByVal lngID As Long, ByVal lng��� As Long, ByVal str��ʼʱ�� As String, _
        ByVal str����ʱ�� As String, ByVal lng�������� As Long, ByVal bln�Ƿ�ԤԼ As Boolean, ByVal str���� As String)
    '�����ṩ����Ϣ���ɼƻ�sql���
    '�����:51427
    Dim strSQL      As String
   '********************************************************
    '    'CREATE OR REPLACE Procedure Zl_�Һżƻ�ʱ��_Insert
    '    (
    '    �ƻ�ID_In   �Һżƻ�ʱ��.�ƻ�ID%Type,
    '    ���_In     �Һżƻ�ʱ��.���%Type,
    '    ��ʼʱ��_In �Һżƻ�ʱ��.��ʼʱ��%Type,
    '    ����ʱ��_In �Һżƻ�ʱ��.����ʱ��%Type,
    '    ��������_In �Һżƻ�ʱ��.��������%Type,
    '    �Ƿ�ԤԼ_In �Һżƻ�ʱ��.�Ƿ�ԤԼ%Type,
    '    ����_In     �Һżƻ�ʱ��.����%Type
    '    )
    '********************************************************
    strSQL = "  Zl_�Һżƻ�ʱ��_Insert("
     '�ƻ�ID_In   �Һżƻ�ʱ��.�ƻ�ID%Type,
    strSQL = strSQL & lngID & ","
     '���_In     �Һżƻ�ʱ��.���%Type,
    strSQL = strSQL & lng��� & ","
     '��ʼʱ��_In �Һżƻ�ʱ��.��ʼʱ��%Type,
     strSQL = strSQL & str��ʼʱ�� & ","
      '����ʱ��_In �Һżƻ�ʱ��.����ʱ��%Type,
    strSQL = strSQL & str����ʱ�� & ","
      '��������_In �Һżƻ�ʱ��.��������%Type,
    strSQL = strSQL & lng�������� & ","
     '�Ƿ�ԤԼ_In �Һżƻ�ʱ��.�Ƿ�ԤԼ%Type,
    strSQL = strSQL & IIf(bln�Ƿ�ԤԼ, 1, 0) & ","
     '����_In     �Һżƻ�ʱ��.����%Type
    strSQL = strSQL & "'" & str���� & "')"
    GetInsert�ƻ�Sql = strSQL
End Function

Private Function ExistsBooking(ByVal lng����ID As String, str���� As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ű��Ƿ����ԤԼ�Һŵ�
    '���:str�ű�-�ű�;str����-���ڼ��İ���
    '����:����,�������Һ����,�����ڷ���-1
    '����:
    '����:2012-04-26 10:32:02
    '�����:51657
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "" & _
    "   Select max(����) as ����  From ���˹Һż�¼ A, �ҺŰ��� B " & _
    "   Where A.�ű� = B.���� " & _
    "       And ��¼״̬ = 1 and b.id=[1] " & _
    "       And Decode(To_Char(A.����ʱ��, 'D'), '1', '����', '2','��һ', '3', '�ܶ�', '4', '����', '5', '����', '6','����', '7', '����', Null) =[2]" & _
    "       And A.����ʱ�� >= Trunc(Sysdate)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, str����)
    ExistsBooking = CLng(Nvl(rsTmp!����, "-1"))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub DeleteSelectPain()
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��ѡ�е�ʱ�����
    '����:����
    '����:2012-07-12 10:32:02
    '�����:51429
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String
    Dim lng����ID As Long
    Dim lng������ As Long
    Dim lng��ǰ����к� As Long
    Dim lng��ǰ���� As Long
    Dim blnDel As Boolean
    Dim i As Long
    Dim j As Long
    
    If chk��ſ���.Value <> 1 Then Exit Sub
    If vsTime.TextMatrix(vsTime.Row, vsTime.Col) = "" Then Exit Sub
    cmdɾ��.Visible = False
    cmdԤԼ.Visible = False
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    lng����ID = Val(txt�ű�.Tag)
    lng������ = ExistsBooking(lng����ID, str����)
    
    '����Ƿ��Ǵ����ʼɾ��
    With vsTime
'         For i = 0 To vsTime.Rows - 1
'            For j = 0 To vsTime.Cols - 1
'                If IsNumeric(.TextMatrix(i, j)) = True Then
'                    If lng������ < IIf(.TextMatrix(i, j) = "", "0", .TextMatrix(i, j)) Then
'                        lng������ = .TextMatrix(i, j)
'                    End If
'                End If
'            Next
'         Next

'         If lng������ <> CLng(IIf(.TextMatrix(lng��ǰ����к�, .Col) = "", "0", .TextMatrix(lng��ǰ����к�, .Col))) Then
'                MsgBox "ֻ�ܴ����ĺ���ʼɾ����", vbInformation, Me.Caption
'                Exit Sub
'         End If
     If .Row Mod 2 = 0 Then
            lng��ǰ����к� = .Row
         Else
            lng��ǰ����к� = .Row - 1
     End If
     lng��ǰ���� = Val(.TextMatrix(lng��ǰ����к�, .Col))
   
    '����Ƿ�úű��Ѿ����ҳ�
     If lng������ >= lng��ǰ���� Then
                MsgBox lng������ & "���Ѿ����ҳ�,ֻ��ɾ���ú��Ժ����ţ�", vbInformation, Me.Caption
                Exit Sub
     End If
     
     SetVsTime lng��ǰ����к�, .Col
     '��ո������Ϣ
     
'     .TextMatrix(lng��ǰ����к�, .Col) = ""
'     .TextMatrix(lng��ǰ����к� + 1, .Col) = ""
    End With
End Sub


Public Sub SetVsTime(lngRow As Long, lngCol As Long)
    Dim i As Long
    Dim j As Long
    Dim lng��ǰ��� As Long
    
    With vsTime
         lng��ǰ��� = Val(.TextMatrix(lngRow, .Col))
         .TextMatrix(lngRow, .Col) = ""
         .TextMatrix(lngRow + 1, .Col) = ""
         For i = lngRow + 2 To .Rows - 1 Step 2
            For j = 1 To .Cols - 1
                    If .TextMatrix(i, j) <> "" Then
                        .TextMatrix(i, j) = lng��ǰ���
                         lng��ǰ��� = lng��ǰ��� + 1
                    End If
            Next
         Next
    End With
End Sub
