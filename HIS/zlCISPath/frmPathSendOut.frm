VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmPathSendOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����·����Ŀ"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11790
   Icon            =   "frmPathSendOut.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11790
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPati 
      Height          =   240
      Left            =   7560
      Picture         =   "frmPathSendOut.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��Ӥ��"
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lstPati 
      Appearance      =   0  'Flat
      Height          =   1080
      ItemData        =   "frmPathSendOut.frx":6948
      Left            =   5160
      List            =   "frmPathSendOut.frx":6955
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11790
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5175
      Width           =   11790
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9360
         TabIndex        =   11
         Top             =   240
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpAdviceTime 
         Height          =   300
         Left            =   7320
         TabIndex        =   10
         Top             =   270
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   111673347
         CurrentDate     =   41129.5916666667
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10560
         TabIndex        =   8
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lblAdviceTime 
         Caption         =   "ҽ��ȱʡ��ʼʱ��"
         Height          =   180
         Left            =   5760
         TabIndex        =   9
         Top             =   330
         Width           =   1575
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11790
      TabIndex        =   3
      Top             =   0
      Width           =   11790
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   3960
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰʱ��׶Σ�"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰʱ�䣺"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label lblPhase 
         BackStyle       =   0  'Transparent
         Caption         =   "�����ǿ��Խ����ʱ��׶Σ���������ѡ�񼴽������ʱ��׶Ρ�"
         Height          =   615
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   9015
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathSendOut.frx":6971
         Top             =   45
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   3405
      Left            =   0
      TabIndex        =   2
      Top             =   1720
      Width           =   11655
      _cx             =   20558
      _cy             =   6006
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   8000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSendOut.frx":71F9
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsPhase 
      Height          =   705
      Left            =   30
      TabIndex        =   0
      Top             =   960
      Width           =   11655
      _cx             =   20558
      _cy             =   1244
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
      BackColor       =   15597549
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   15724768
      BackColorSel    =   45056
      ForeColorSel    =   16777215
      BackColorBkg    =   15597549
      BackColorAlternate=   15597549
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   32768
      FloodColor      =   192
      SheetBorder     =   15724768
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   2
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   450
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSendOut.frx":7395
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
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
Attribute VB_Name = "frmPathSendOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPP                 As TYPE_PATH_Pati
Private mPati               As TYPE_Pati
Private mfrmParent          As Object

Private mlngʱ�����        As Integer                  'mlngFun=0ʱ���룬1=��һ�׶���ǰ������,2=��һ�׶��Ӻ�������,-1=��һ�׶��Ӻ󣨼�����ǰ�׶Σ�,0=����
Private mstrBaby            As String                   'Ӥ��������,��������,���˼,...
Private mblnOK              As Boolean

Private mlngFun             As Long                     '0-����·����1-��������(����ѡ��׶�),2-�鿴·���׶ζ���,3-��������ҽ��
Private mlng��ĿID          As Long                     '�������ɵ���ĿID
Private mlngִ��ID          As Long                     '�������ɵ�·��ִ��ID
Private mlng���˽׶�ID      As Long                     '��ǰѡ��Ľ׶�(�鿴ʱ)�����˵�ǰ�׶Σ�����ʱ��
Private mlng��ǰ����        As Long                     '����ʱӦ�����ɵ�����(������һ�׶���ǰʱ)
Private mlng·��ҽ������    As Long                     '·��ҽ�����ɳ�ǰ����
Private mlng����            As Long                     '��ǰӦ�����ɵ�����(ʵ������)

Private mdatDur             As Date                     '·������ʱ��
Private mcol                As Collection
Private mEditType           As Collection

Private mrsPhase            As ADODB.Recordset          '���ý׶�
Private mclsMipModule       As zl9ComLib.clsMipModule   '��Ϣƽ̨����

Private Enum ִ�з�ʽ
    T0����ִ�� = 0
    T1�������� = 1
    T3��Ҫʱ = 3
End Enum

Private Enum TYPE_Func
    Func����·�� = 0
    Func�������� = 1
    Func�鿴·�� = 2
    Func�������� = 3
End Enum

Public Function ShowMe(frmParent As Object, ByVal lngFun As Long, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
                        ByVal lng���˽׶�ID As Long, ByVal lng���� As Long, Optional ByVal lng��ĿID As Long, _
                        Optional ByVal lngִ��ID As Long, Optional ByVal lngʱ����� As Long, Optional ByVal bln��ǰ As Boolean = False) As Boolean
'������lng��ĿID,lngִ��ID=��������ҽ��ʱ���贫��
'      lngʱ�����= mlngFun=0ʱ���룬1=��һ�׶���ǰ,2-��һ�׶���ǰ������,-1=��һ�׶��Ӻ󣨼�����ǰ�׶Σ�,0=����
'      bln��ǰ=true :��ǰ����·��,False-����ǰ����
    Set mfrmParent = frmParent
    mlngFun = lngFun
    mlng��ĿID = lng��ĿID
    mlngִ��ID = lngִ��ID
    
    mPati = t_pati
    mPP = t_pp
    mlng���˽׶�ID = lng���˽׶�ID      'ȱʡѡ�е�ǰ�׶�
    
    mlng���� = lng����
    
    mlng��ǰ���� = lng����
    mlngʱ����� = lngʱ�����
    If bln��ǰ Then                     '��ǰ����
        mdatDur = DateAdd("d", 1, CDate(Format(mPP.��ǰ����, "YYYY-MM-DD 00:00:00")))
    Else
        mdatDur = zlDatabase.Currentdate
    End If
    
    Set mrsPhase = GetPhase(mPP.·��ID, mPP.�汾��, mlng���˽׶�ID, mlng����)
    
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Function GetPhase(ByVal lng·��ID As Long, ByVal lng�汾�� As Long, ByVal lng��ǰ�׶�ID As Long, ByVal lng���� As Long) As ADODB.Recordset
'���ܣ���ȡ��ǰʱ����õĽ׶�
'1��2-7��8-12��13-19��20-30
'���ý׶Σ���ǰ����������Ӧ�Ľ׶Σ������ǰʱ��Ϊ��һ�죬��ֻ��ʾ��һ�׶Σ����
    Dim strSql As String, strIF As String, str�׶η��� As String
    Dim rsTmp As ADODB.Recordset, datPathIn As Date, lngʱ����� As Long
    Dim lng��� As Long
    Dim strMainIF As String
    
    If mlngFun = 2 Then         '�鿴�׶ζ������Ŀ
        strSql = " Select a.Id,Nvl(a.��id,0) as ��id,a.���,a.����,a.˵��,a.��ʼ����,a.��������,a.����" & vbNewLine & _
                 " From ����·���׶� A" & vbNewLine & _
                 " Where a.·��id = [1] And a.�汾�� = [2] And a.id = [4]" & vbNewLine & _
                 " Order by ���"
    Else
        datPathIn = GetPatiInPathOut(mPP.����·��ID)                                            '��ȡ���˵Ľ���·���Ŀ�ʼʱ��
        
        If mlngFun = 0 Then
            If mlngʱ����� = -1 Then                     '�Ӻ�ʱ������ǰ�׶�
                strIF = " And a.id = [4]"
            Else
                If mPP.��ǰ�׶�ID <> 0 Then
                    lng��� = GetPhaseNOOut(mPP.��ǰ�׶�ID)
                End If
                
                If mlngʱ����� = 1 Or mlngʱ����� = 2 Then
                    strIF = " And NVL(d.���,a.���)>[6] "
                Else
                    If mPP.��ǰ�׶�ID <> 0 Then
                        '֮ǰ��������ǰִ�й��Ľ׶ε�ʱ�䷶Χ�ڵ�ǰ�����ڣ�Ҫ�ų���Щ�׶Σ�·����תʱ����顣
                        strIF = " And NVL(d.���,a.���)>=[6] "
                    End If
                    
                     'ͬһ���ж���׶�ʱ����ǰ�׶μ���֧��������,����ǽ�����һ���ˣ���˵��û����ͬ�����Ľ׶�
                    If lng���� = mPP.��ǰ���� Then
                        strIF = strIF & " And Nvl(a.��id,a.id) <> " & IIf(mPP.�׶θ�ID <> 0, "[7]", "[4]")
                    End If
                End If
                
                str�׶η��� = Get�׶η���Out(mPP.����·��ID)
                If str�׶η��� <> "" Then
                    strIF = strIF & " And (a.��id is Null Or a.��id is Not Null And a.���� = [5])"
                End If

                strMainIF = strIF
                
                'strIF = strIF & " And (a.��ʼ���� Is Null Or [3] Between a.��ʼ���� And Nvl(a.��������,a.��ʼ����) " & ")"
            End If
        Else
            strIF = " And a.id = [4]"
        End If
      
        strSql = " Select a.Id, Nvl(a.��id, 0) As ��id, a.���, a.����, a.˵��, a.��ʼ����, a.��������, a.����" & vbNewLine & _
                 " From ����·���׶� A, ����·���׶� D" & vbNewLine & _
                 " Where a.��id = d.Id(+) And a.·��id = [1] And a.�汾�� = [2]" & vbNewLine & _
                   strIF & " Order By Nvl(d.���, a.���)"
    End If
    On Error GoTo errH
    Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ý׶�", lng·��ID, lng�汾��, lng����, lng��ǰ�׶�ID, str�׶η���, lng���, mPP.�׶θ�ID)
    
    If (mlngʱ����� = 1 Or mlngʱ����� = 2) And GetPhase.RecordCount = 0 Then
        '�׶���ǰʱ�������ǰ�׶��ж��죬�򰴵�ǰ����ȡ������һ�׶Σ�ֱ��ȡ��Ŵ��ڵ�ǰ�׶ε���һ�׶�
        strSql = " Select * From (Select a.Id, Nvl(a.��id,0) as ��id, a.���, a.����, a.˵��,a.��ʼ����, a.��������, a.����" & vbNewLine & _
                 " From ����·���׶� A,����·���׶� D " & vbNewLine & _
                 " Where a.��ID=d.id(+) and a.·��id = [1] And a.�汾�� = [2]" & _
                  strMainIF & vbNewLine & " Order by NVL(d.���,a.���)) Where Rownum=1"
        Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ý׶�", lng·��ID, lng�汾��, lng����, lng��ǰ�׶�ID, str�׶η���, lng���)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdPati_Click()
    Dim i As Long
    Dim arrtmp As Variant
    Dim lngW As Long
    Dim lngH As Long
    Dim strTmp As String
    Dim strSelect As String
            
    lstPati.Visible = True
    lblFont.FontSize = lblPhase.FontSize
    With lstPati
        .Clear
        strTmp = "���˱���|" & mstrBaby
        arrtmp = Split(strTmp, "|")
        For i = LBound(arrtmp) To UBound(arrtmp)
            lblFont.Caption = arrtmp(i)
            If lngW < lblFont.Width Then
                lngW = lblFont.Width
            End If
           .AddItem arrtmp(i)
        Next
        lngH = (i - 1) * 210 + 240
        If lngH > 1080 Then lngH = 1080
        lngW = lngW + 700
        If lngW > 2500 Then lngW = 2500
    End With

    With vsItem
        strSelect = .TextMatrix(.Row, .Col)
        For i = 0 To lstPati.ListCount - 1
            If InStr("|" & strSelect & "|", "|" & lstPati.List(i) & "|") > 0 Then
                lstPati.Selected(i) = True
            End If
        Next
        If lngW < .ColWidth(mcol("Ӥ��")) Then
            lngW = .ColWidth(mcol("Ӥ��"))
        End If
        lstPati.Move .Left + .ColPos(.Col), .Top + .RowPos(.Row) + .RowHeight(.Row) + 30, lngW, lngH
    End With
    Call lstPati.SetFocus
End Sub

Private Sub Form_Load()
    If mlngFun <> 2 Then
        vsItem.Editable = flexEDKbdMouse
    End If
    
    vsItem.Top = vsPhase.Top + vsPhase.Height + 45
    
    Call LoadPhase                          '���ؿ�ѡ��Ľ׶�
    
    mlng·��ҽ������ = Val(zlDatabase.GetPara("·��ҽ�����ɳ�ǰ����", glngSys, P����·��Ӧ��, "1"))
    
    If vsPhase.Cols = 1 Then
        vsPhase.Visible = False
        lblPhase.Caption = vsPhase.TextMatrix(0, 0) & vbCrLf & vsPhase.Cell(flexcpData, 0, 0)
                
        vsItem.Top = vsPhase.Top
        vsItem.Height = picBottom.Top - vsItem.Top
    Else
        lblNote.Visible = False
        lblPhase.Left = lblNote.Left
    
        If Grid.HScrollVisible(vsPhase) Then
            '���������
            vsPhase.Height = 1000
            vsItem.Height = vsItem.Height - (vsPhase.Top + vsPhase.Height - vsItem.Top + 120)
            vsItem.Top = vsPhase.Top + vsPhase.Height + 60
        Else
            If vsPhase.Rows = 1 Then
                vsPhase.RowHeightMax = vsPhase.Height
                vsPhase.RowHeight(0) = vsPhase.Height
            End If
        End If
    End If
        
    If mlngFun = 2 Then
        Me.Caption = "�鿴�׶ζ������Ŀ"
        lblDate.Visible = False
        
        cmdOK.Visible = False
        cmdCancel.Caption = "�˳�(&X)"
        mstrBaby = ""
        dtpAdviceTime.Visible = False
        lblAdviceTime.Visible = False
    Else
        'ҽ��ȱʡʱ��Ĭ��ȡ��ǰʱ��
        dtpAdviceTime.Value = mdatDur
        lblDate.Caption = "����·����Ŀ���ڣ�" & Format(mdatDur, "yyyy-MM-dd") & ",·����" & mlng���� & "��"
        mstrBaby = GetBabyRegList
    End If
    
    Call InitItem
    
    If mlngFun = 2 Then                                         '�鿴ʱֻ��ʾ���� , ��Ŀ����
        Me.Width = vsItem.Width + 360
        cmdCancel.Left = vsItem.Left + vsItem.Width - 1200
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 150
    End If
    
    Set mEditType = New Collection
    Call LoadItem(Val(vsPhase.ColData(vsPhase.Col)), vsItem, mPP.·��ID, mPP.�汾��)
    
    If vsItem.Rows = 1 Then
        vsItem.Rows = 2
        vsItem.TextMatrix(1, mcol("��Ŀ����")) = "û�б���ִ�л��ѡ�Ե�·����Ŀ"
        cmdOK.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsPhase = Nothing
    Set mcol = Nothing
    Set mEditType = Nothing
    Set mclsMipModule = Nothing
End Sub

Private Sub lstPati_ItemCheck(Item As Integer)
    Dim i As Long
    Dim strList As String
    strList = ""
    For i = 0 To lstPati.ListCount - 1
        If lstPati.Selected(i) Then
            strList = strList & "|" & lstPati.List(i)
        End If
    Next
    If strList <> "" Then
        strList = Mid(strList, 2)
    Else
        strList = lstPati.List(0)                           'ȱʡѡ�в��˱���,������Ϊ��
    End If
    
    vsItem.TextMatrix(vsItem.Row, vsItem.Col) = strList
End Sub

Private Sub lstPati_KeyPress(KeyAscii As Integer)
    If lstPati.Visible Then
        If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
            lstPati.Visible = False
        End If
    End If
End Sub

Private Sub lstPati_LostFocus()
    lstPati.Visible = False
End Sub

Private Sub vsItem_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    With vsItem
        If cmdPati.Visible Then
            cmdPati.Move .Left + .ColPos(.Col) + .ColWidth(.Col) - 255, .Top + .RowPos(.Row) + 15, 255, 240
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    With vsItem
        If .Col = mcol("Ӥ��") And cmdPati.Visible Then
            cmdPati.Move .Left + .ColPos(.Col) + .ColWidth(.Col) - 255, .Top + .RowPos(.Row) + 15, 255, 240
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_Click()
    With vsItem
        If lstPati.Visible Then
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_DblClick()
    Dim lng��ĿID As Long
    
    If vsItem.Col = mcol("��Ŀ����") Then
        lng��ĿID = Val(vsItem.TextMatrix(vsItem.Row, mcol("ID")))
        If lng��ĿID <> 0 Then
            Call frmPathItemEditOut.ShowView(mfrmParent, lng��ĿID)
        End If
    End If
End Sub

Private Sub vsItem_GotFocus()
    vsItem.ForeColorSel = vbWhite
    vsItem.BackColorSel = &H8000000D
End Sub

Private Sub vsItem_LostFocus()
    vsItem.ForeColorSel = vbBlack
    vsItem.BackColorSel = vbWhite
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ResultEnterNextCell(vsItem)
    End If
End Sub

Private Sub ResultEnterNextCell(vsthis As VSFlexGrid)
    With vsthis
        If .Col <= .Cols - 1 Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsItem
        If Col = mcol("ѡ��") Then
            If .Cell(flexcpChecked, Row, mcol("ѡ��")) = 2 Then
                'δѡ��ʱ��������ԭ��ѡ��
                If mlngFun = 0 Then
                    If .RowData(Row) = ִ�з�ʽ.T1�������� Then
                        Call vsItem_CellButtonClick(Row, mcol("����ԭ��"))
                    End If
                End If
            ElseIf .Cell(flexcpChecked, Row, mcol("ѡ��")) = 1 Then
                If .RowData(Row) = ִ�з�ʽ.T1�������� Then
                    'ѡ��ʱ���������ԭ��
                    .TextMatrix(Row, mcol("����ԭ��")) = ""
                    .Cell(flexcpData, Row, mcol("����ԭ��")) = ""
                End If
            End If
        ElseIf Col = mcol("ȫѡ") Then
            If .Cell(flexcpChecked, Row, mcol("ȫѡ")) = 2 Then
                For i = Row To .Rows - 1
                    'ÿ�����ɵ�ȡ����Ҫ��дԭ�����Բ�ȡ����ѡ��Ҫȡ��ֻ��һ��һ��ȡ��
                    If .TextMatrix(i, mcol("����")) <> .TextMatrix(Row, mcol("����")) Then Exit For
                    If Not (.RowData(i) = ִ�з�ʽ.T0����ִ�� Or .RowData(i) = ִ�з�ʽ.T1��������) Then
                        If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                            .Cell(flexcpChecked, i, mcol("ѡ��")) = 2
                        End If
                    End If
                Next
                For i = Row - 1 To .FixedRows Step -1
                    If .TextMatrix(i, mcol("����")) <> .TextMatrix(Row, mcol("����")) Then Exit For
                    If Not (.RowData(i) = ִ�з�ʽ.T0����ִ�� Or .RowData(i) = ִ�з�ʽ.T1��������) Then
                        If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                            .Cell(flexcpChecked, i, mcol("ѡ��")) = 2
                        End If
                    End If
                Next
                
            ElseIf .Cell(flexcpChecked, Row, mcol("ȫѡ")) = 1 Then
                For i = Row To .Rows - 1
                    If .TextMatrix(i, mcol("����")) <> .TextMatrix(Row, mcol("����")) Then Exit For
                    If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                        .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        If .RowData(i) = ִ�з�ʽ.T1�������� Then
                        'ѡ��ʱ���������ԭ��
                            .TextMatrix(i, mcol("����ԭ��")) = ""
                            .Cell(flexcpData, i, mcol("����ԭ��")) = ""
                        End If
                    End If
                Next
                For i = Row - 1 To .FixedRows Step -1
                    If .TextMatrix(i, mcol("����")) <> .TextMatrix(Row, mcol("����")) Then Exit For
                    If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                        .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        If .RowData(i) = ִ�з�ʽ.T1�������� Then
                        'ѡ��ʱ���������ԭ��
                            .TextMatrix(i, mcol("����ԭ��")) = ""
                            .Cell(flexcpData, i, mcol("����ԭ��")) = ""
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsItem
        If NewRow >= .FixedRows And Me.Visible Then
            If mlngFun = 0 Then
                If NewCol = mcol("����ԭ��") Then
                    'δѡ��ʱ�������û�ѡ�����ԭ��
                    If .RowData(NewRow) = ִ�з�ʽ.T1�������� And .Cell(flexcpChecked, NewRow, mcol("ѡ��")) = 2 Then
                        .ColComboList(mcol("����ԭ��")) = "..."
                    Else
                        .ColComboList(mcol("����ԭ��")) = ""
                    End If
                End If
            End If
            cmdPati.Visible = False: lstPati.Visible = False
            If NewCol = mcol("Ӥ��") And mlngFun <> 2 Then
                If mstrBaby <> "" Then
                    cmdPati.Visible = True
                    cmdPati.Enabled = True
                    cmdPati.Move .Left + .ColPos(NewCol) + .ColWidth(NewCol) - 255, .Top + .RowPos(NewRow) + 15, 255, 240
                    If .RowData(NewRow) = ִ�з�ʽ.T0����ִ�� Then
                        cmdPati.Enabled = False
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsItem_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsItem
        If Col = mcol("ѡ��") Then
            '����·��ʱ��ÿ�����ɵģ����Բ�ѡ����Ҫ�������ԭ��
            If .RowData(Row) = ִ�з�ʽ.T0����ִ�� Or mlngFun <> 0 And .RowData(Row) = ִ�з�ʽ.T1�������� Then
                Cancel = True
            End If
        ElseIf Col = mcol("ȫѡ") Then
            If Val(.Cell(flexcpChecked, Row, Col)) = 0 Then Cancel = True
        ElseIf Col = mcol("Ӥ��") Then
            Cancel = True
        ElseIf Col = mcol("����ԭ��") Then
            If .ColComboList(mcol("����ԭ��")) = "" Then Cancel = True
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSql As String, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
            
    With vsItem
        If Col = mcol("����ԭ��") Then
            strSql = "Select b.���� as ����,a.���� as ID,a.����,a.����,a.���� From ������쳣��ԭ�� a,������쳣��ԭ�� b" & _
                    " Where a.����=1 And a.ĩ��=1 And a.�ϼ�=b.���� And b.ĩ��=0 " & _
                    " Order by ����,a.����"
            
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "������쳣��ԭ��", True, , , True, True, True, _
                     Me.Left + .Left + .ColPos(Col), Me.Top + .Top + .RowPos(Row) + .RowHeight(Row) * 2, .RowHeight(Row), blnCancel, False, True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "ϵͳû�г�ʼ������쳣��ԭ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                .TextMatrix(Row, Col) = rsTmp!����
                .Cell(flexcpData, Row, Col) = CStr(rsTmp!����)
            End If
        End If
    End With
End Sub

Private Sub vsPhase_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng�׶�ID As Long
    Dim str���� As String

    If OldCol <> NewCol And Me.Visible = True And NewCol >= 0 And NewRow >= 0 And vsPhase.Redraw = flexRDDirect Then
        lng�׶�ID = vsPhase.ColData(NewCol)
        mrsPhase.Filter = "ID=" & lng�׶�ID
        mlng���� = (mrsPhase!��ʼ���� & "")
        Call LoadItem(lng�׶�ID, vsItem, mPP.·��ID, mPP.�汾��)
    End If
End Sub

Private Sub LoadPhase()
'���ܣ����ؿ�ѡ��Ľ׶�,������˵ĵ�ǰʱ��׶���Ȼ���ã���ѡ�У�����ȱʡΪ��һ��
    Dim i As Long, j As Long, str�׶η��� As String
    Dim rsNode As ADODB.Recordset

    With vsPhase
        .Clear
        .Redraw = flexRDNone
        .Col = -1
        mrsPhase.Filter = ""
        .Cols = mrsPhase.RecordCount
        str�׶η��� = Get�׶η���Out(0, mPP.��ǰ�׶�ID)
        If mlngFun = 0 And mlngʱ����� <> -1 Then '�������ɡ���������ʱ����һ�׶��Ӻ󣨼�����ǰ�׶Σ���ֻ�е�ǰ�׶εļ�¼
            mrsPhase.Filter = "��ID<>0 "
            If mrsPhase.RecordCount > 0 Then    '�б��÷�֧
                Set rsNode = mrsPhase.Clone
                .Rows = 2
                .MergeRow(0) = True
            Else
                .Rows = 1
            End If
            mrsPhase.Filter = "��ID=0"
        End If
    
        For i = 0 To .Cols - 1
            .ColWidth(i) = 2000
            .ColAlignment(i) = flexAlignCenterCenter
            .TextMatrix(0, i) = mrsPhase!����
            .Cell(flexcpData, 0, i) = CStr(IIf(IsNull(mrsPhase!����), "", "���ࣺ" & mrsPhase!���� & " ") & mrsPhase!˵��)
            .ColData(i) = Val(mrsPhase!ID)
            
            If .ColData(i) = mlng���˽׶�ID Then .Col = i
            If Not rsNode Is Nothing Then
                rsNode.Filter = "��ID=" & mrsPhase!ID
                If rsNode.RecordCount = 0 Then
                     .MergeCol(i) = True
                     .TextMatrix(1, i) = mrsPhase!����
                Else
                     .TextMatrix(1, i) = "ȱʡ"
                     .ColWidth(i) = 1000
                    For j = 1 To rsNode.RecordCount
                        i = i + 1
                         .ColWidth(i) = 1000
                         .ColAlignment(i) = flexAlignCenterCenter
                        .TextMatrix(0, i) = mrsPhase!����           '��һ��������ͬ�������ںϲ�
                        .TextMatrix(1, i) = IIf(IsNull(rsNode!˵��), "��֧" & j, "" & rsNode!˵��)
                        .Cell(flexcpData, 1, i) = CStr(IIf(IsNull(rsNode!����), "", "���ࣺ" & rsNode!���� & " ") & rsNode!˵��)
                        
                        .ColData(i) = Val(rsNode!ID)
                        If .ColData(i) = mlng���˽׶�ID Then
                            .Col = i
                        ElseIf .Col = 0 And str�׶η��� <> "" Then
                            If str�׶η��� = "" & rsNode!���� Then
                                .Col = i
                            End If
                        End If
                        rsNode.MoveNext
                    Next
                End If
            End If
            mrsPhase.MoveNext
        Next
        If .Col < 0 Then .Col = 0
        mrsPhase.Filter = "ID=" & Val(.ColData(.Col))
        .Redraw = True
        vsPhase_AfterRowColChange -1, -1, .Row, .Col
    End With
End Sub

Private Sub LoadItem(lng�׶�ID As Long, objVsg As VSFlexGrid, ByVal lng·��ID As Long, ByVal lng�汾�� As Long)
'���ܣ����ص�ǰ�׶ε�·����Ŀ
'������objVsg����Ҫ���صı��
    Dim i As Long, j As Long, blnFocus As Boolean
    Dim rsTmp As ADODB.Recordset, strSql As String, strIDs As String, strTmp As String
    Dim rsAdvice As ADODB.Recordset, rsFile As ADODB.Recordset
    Dim lngRow As Long
    Dim str������ĿIDs As String
    Dim lng��Ҫ·���׶�ID As Long
    Dim strNewTmp As String
     
    If mlngFun = 1 Then '�������ɣ�����ִ�еĲ���ʾ������ִ�й��Ĳ����ظ�����,ִֻ��һ�εĵ�ǰ�׶���ִ������ʾ
        strSql = " And a.ִ�з�ʽ<>0 And Not Exists(Select 1 From ��������·��ִ�� c " & _
                 " Where c.·����¼id = [4] And c.�׶�id = [7] And c.��Ŀid = a.id And (c.���� = [5] and a.ִ�з�ʽ<>4 or a.ִ�з�ʽ=4))"
        lng��Ҫ·���׶�ID = mPP.��ǰ�׶�ID
    ElseIf mlngFun = 3 Then '��������
        strSql = " And a.ID = [6]"
    Else
        strSql = " And (a.ִ�з�ʽ<>4 or a.ִ�з�ʽ=4 And Not Exists(Select 1 From ��������·��ִ�� c " & _
                 " Where c.·����¼id = [4] And c.�׶�id = [7] And c.��Ŀid = a.id))"
        lng��Ҫ·���׶�ID = lng�׶�ID
    End If
    '���ӡ�����·�����ࡱ��ֻ��Ϊ�˰���������'����ʱ�ټ�飬�Ƿ�Ϊ����׶ε����һ�죬����ִ��һ�ε���Ŀ�Ƿ�ѡ��
    strSql = " Select a.����, a.ID, a.��Ŀ����, a.ִ�з�ʽ, a.ͼ��id, a.����Ҫ��" & vbNewLine & _
             " From ����·����Ŀ A, ����·������ B" & vbNewLine & _
             " Where a.���� = b.���� And a.·��id = b.·��id And a.�汾�� = b.�汾�� And a.·��id = [1] And a.�汾�� = [2] And a.�׶�id = [3] " & vbNewLine & _
               strSql & IIf(mlngFun = 3, "", "Order By b.���, a.��Ŀ���")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ID, lng�汾��, lng�׶�ID, mPP.����·��ID, mlng����, mlng��ĿID, lng��Ҫ·���׶�ID)
    
    With objVsg
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
        lngRow = 1
        '���ڹ̶��кϲ�����Ӱ�������У�����������һ���жϰ�����ϲ�ȫѡ��
        .MergeCells = flexMergeRestrictAll
        .MergeCol(mcol("����")) = True
        .MergeCol(mcol("����ֵ")) = True
        .MergeCol(mcol("ȫѡ")) = True

        For i = lngRow To rsTmp.RecordCount + lngRow - 1
            .TextMatrix(i, mcol("ID")) = rsTmp!ID
            strIDs = strIDs & "," & rsTmp!ID
            .TextMatrix(i, mcol("����")) = rsTmp!����
            .TextMatrix(i, mcol("����ֵ")) = rsTmp!����
            .TextMatrix(i, mcol("��Ŀ����")) = rsTmp!��Ŀ����
            
            If mlngFun <> 2 Then
                .TextMatrix(i, mcol("����Ҫ��")) = Val("" & rsTmp!����Ҫ��)
            End If
            
            .TextMatrix(i, mcol("ִ�з�ʽ")) = Decode(rsTmp!ִ�з�ʽ, 0, "��", 1, "����", 3, "��Ҫʱ")
            .RowData(i) = Val(rsTmp!ִ�з�ʽ)
            
            If mlngFun <> 2 Then
                Select Case rsTmp!ִ�з�ʽ
                    Case ִ�з�ʽ.T0����ִ��
                        .TextMatrix(i, mcol("ѡ��")) = " "
                        .Cell(flexcpBackColor, i, mcol("ѡ��")) = &H8000000F
                    Case ִ�з�ʽ.T1��������
                        .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        .Cell(flexcpPictureAlignment, i, mcol("ѡ��")) = flexPicAlignCenterCenter
                        If mlngFun <> 0 Then
                            .Cell(flexcpBackColor, i, mcol("ѡ��")) = &H8000000F
                        End If
                        '����ʱ��������Ϊ��ɫ��������Ϊ���Բ�ѡ�����ɣ���¼�����ԭ��
                   Case Else
                        If mlngFun = 3 Then '��ѡʱ��ѡ����δ��ʾ���Զ�����
                            .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        Else
                            .Cell(flexcpChecked, i, mcol("ѡ��")) = IIf(rsTmp.RecordCount = 1, 1, 2)
                            .Cell(flexcpPictureAlignment, i, mcol("ѡ��")) = flexPicAlignCenterCenter
                        End If
                End Select
            End If
            
            If Not IsNull(rsTmp!ͼ��ID) Then
                Call .Select(i, mcol("��Ŀ����"))
                .CellPictureAlignment = flexPicAlignRightCenter 'flexPicAlignLeftCenter
                .CellPicture = GetPathIcon(rsTmp!ͼ��ID)
            End If
            
            If mstrBaby <> "" Then
                .TextMatrix(i, mcol("Ӥ��")) = "���˱���"
            End If
            
            If (rsTmp!ִ�з�ʽ = ִ�з�ʽ.T3��Ҫʱ) And blnFocus = False Then
                Call .Select(i, mcol("ѡ��"))
                blnFocus = True
            End If
            rsTmp.MoveNext
        Next
        
        strIDs = Mid(strIDs, 2)
        '������Ŀ��Ӧ��ҽ��
        Set rsAdvice = GetAdviceOut(strIDs)
        If rsAdvice.RecordCount > 0 Then
            For i = .FixedRows To .Rows - 1
                rsAdvice.Filter = "·����ĿID=" & ZVal(Val(.TextMatrix(i, mcol("ID"))))
                strTmp = ""
                str������ĿIDs = ""
                
                For j = 1 To rsAdvice.RecordCount
                    strTmp = strTmp & "," & rsAdvice!ҽ������ID
                    str������ĿIDs = str������ĿIDs & "," & rsAdvice!������ĿID
                    rsAdvice.MoveNext
                Next
                If strTmp <> "" Then
                    .TextMatrix(i, mcol("ҽ������ID")) = Mid(strTmp, 2)
                    If mlngFun <> 2 Then
                        .TextMatrix(i, mcol("������ĿID")) = Mid(str������ĿIDs, 2)
                    End If
                    .TextMatrix(i, mcol("��Ŀ����")) = .TextMatrix(i, mcol("��Ŀ����")) & " ����"
                End If
            Next
        End If
        
        '������Ŀ��Ӧ�Ĳ����ļ�
        If mlngFun <> 3 Then
            Set rsFile = GetFile(strIDs, 1)
            If rsFile.RecordCount > 0 Then
                strIDs = ""
                For i = .FixedRows To .Rows - .FixedRows
                    rsFile.Filter = "·����ĿID=" & ZVal(Val(.TextMatrix(i, mcol("ID"))))
                    strTmp = "": strNewTmp = "" '��¼�°���Ӳ���ID
                    For j = 1 To rsFile.RecordCount
                        If rsFile!�ļ�ID & "" <> "" Then
                            strTmp = strTmp & "," & rsFile!�ļ�ID
                            If InStr(strIDs & ",", "," & rsFile!�ļ�ID & ",") = 0 Then
                                On Error Resume Next
                                mEditType.Add Val(rsFile!����), "C" & rsFile!�ļ�ID
                                On Error GoTo errH
                            End If
                        Else
                            strNewTmp = strNewTmp & "," & rsFile!ԭ��ID
                        End If
                        rsFile.MoveNext
                    Next
                    If strTmp <> "" Or strNewTmp <> "" Then
                        strIDs = strIDs & strTmp    'ͬһ��·����Ŀ���ļ�ID�����أ����Էŵ��ڶ���ѭ����
                        .TextMatrix(i, mcol("�ļ�ID")) = IIf(strTmp <> "", Mid(strTmp, 2), "") & "|" & IIf(strNewTmp <> "", Mid(strNewTmp, 2), "")
                        .TextMatrix(i, mcol("��Ŀ����")) = .TextMatrix(i, mcol("��Ŀ����")) & " ��"
                    End If
                Next
            End If
        End If
        If .Rows = .FixedRows Then
            .Rows = .Rows + 1
        End If
        .Redraw = True
    End With
               
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitItem()
'����: ��ʼ��·����Ŀ��ͷ
    Dim strcol As String, arrHead As Variant
    Dim i As Long
    
    If mlngFun = 2 Then
        strcol = "����,1200,4;����ֵ;ȫѡ,450,4;��Ŀ����,5950,1;ִ�з�ʽ;ѡ��;Ӥ��;ID;ҽ������ID;�ļ�ID"
    Else
        strcol = "����,1200,4;����ֵ;ȫѡ" & IIf(mlngFun <> Func��������, ",450,4", "") & ";��Ŀ����," & IIf(mstrBaby = "", 5950, 4950) & ",1;ִ�з�ʽ,900,1" & _
                ";ѡ��" & IIf(mlngFun <> Func��������, ",500,4", "") & _
                ";Ӥ��" & IIf(mstrBaby = "", "", ",1100,1") & _
                ";ID;ҽ������ID;�ļ�ID;����Ҫ��;����ԭ��" & IIf(mlngFun = 0, ",1800,4", "") & ";�Ƿ����һ��;�׶�ID;������ĿID;�ظ���Ŀ"
    End If
    arrHead = Split(strcol, ";")
    Set mcol = New Collection
   
    With vsItem
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        
        For i = 0 To UBound(arrHead)
            mcol.Add i, Split(arrHead(i), ",")(0)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Sub vsPhase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsPhase.MouseCol >= 0 And vsPhase.MouseRow >= 0 Then
        Dim strInfo As String
        strInfo = Trim(vsPhase.Cell(flexcpData, vsPhase.MouseRow, vsPhase.MouseCol))
        Call zlCommFun.ShowTipInfo(vsPhase.Hwnd, strInfo)
    End If
End Sub

Private Function GetBabyRegList() As String
'���ܣ���ȡ���˵�Ӥ�������б�
'���أ�"����1,����2,����3��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ���,Ӥ������ From ������������¼ Where ����ID=[1] And ��ҳID=[2] Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetBabyRegList", mPati.����ID, mPati.�Һ�ID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = IIf(strSql = "", "", strSql & "|") & "Ӥ��:" & NVL(Replace(rsTmp!Ӥ������, "|", "_"))
        rsTmp.MoveNext
    Loop
    GetBabyRegList = strSql
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetBabyIndex(strtxt As String) As String
'���ܣ����ݵ�ǰ�е����ݷ���Ӥ�����
    Dim i As Long, j As Long
    Dim arrtmp As Variant
    Dim arrBaby As Variant
    Dim strBaby As String
    
    If mstrBaby <> "" Then
        arrtmp = Split("���˱���|" & mstrBaby, "|")
        arrBaby = Split(strtxt, "|")
        For j = 0 To UBound(arrBaby)
            For i = 0 To UBound(arrtmp)
                If arrtmp(i) = arrBaby(j) Then
                    strBaby = strBaby & "|" & i
                    Exit For
                End If
            Next
        Next
        GetBabyIndex = Mid(strBaby, 2)
    Else
        GetBabyIndex = "0"                  'û��Ӥ��ȱʡȡ���˱���
    End If
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long, blnEnd As Boolean, strIDs As String, strAdviceOfItem As String
    Dim arrSQL As Variant, arrBaby As Variant, DatCurr As Date
    Dim strTmp As String, strBaby As String, strBB As String, lng���� As Long
    Dim rsTmp As ADODB.Recordset, rsLastAdvice As ADODB.Recordset
    Dim rsUsed As ADODB.Recordset                   '��������ʱУ�Ե�δ���ϵ�ҽ��
    Dim blnHaveDoc As Boolean
    Dim str��ĿIDs As String
    Dim k As Long, n As Long, strAgain As String
    Dim colItem As New Collection
    Dim strAgaignTmp As String
    Dim str·����ĿIDs As String                    '·������ʱ��ҽ�޸��˵��䷽�ģ��ҳ����������޸��䷽�ı�������Ŀ����Ӧ�ı���ԭ����ĿID1|�������1,��Ŀ2|�������2��������
    Dim colPathItems As New Collection
    
    arrSQL = Array()
    '1.������ִ��һ�ε���Ŀ
    With vsItem
        If mlngFun = 0 Then
            If k = 0 Then
                With mrsPhase
                    .Filter = "ID=" & vsPhase.ColData(vsPhase.Col)
                    If Not IsNull(!��ʼ����) Then
                        If IsNull(!��������) Then
                            blnEnd = (Val(!��ʼ����) = mlng����)
                        Else
                            blnEnd = (Val(!��������) = mlng����)
                        End If
                    End If
                End With
            End If

            '�������ɵ���Ŀ�����û��ѡ��������������ԭ��
            For i = 1 To .Rows - 1
                If .RowData(i) = ִ�з�ʽ.T1�������� Then
                    If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                        If .TextMatrix(i, mcol("����ԭ��")) = "" Then
                            MsgBox "�������ɵ���Ŀ�����ѡ�����ɣ���Ҫ�����ѡ�����ԭ��", vbInformation, gstrSysName
                            If .Visible And .Enabled Then .SetFocus
                            .Select i, mcol("����ԭ��")
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If
    End With
    blnEnd = False
    
    '��������ҽ��������������Ŀ��Ӧ��ҽ��Ϊ�¿�״̬��ɾ������Ŀ��Ӧ������ҽ��,���²���ҽ����¼����;���ڷ���δ���ϵ�ҽ������Ҫ�����ϡ�
    If mlngFun = Func�������� Then
        Set rsLastAdvice = GetUsedAdvice(mlngִ��ID, mlng��ĿID)
    End If
    
    strIDs = ""
    With vsItem
        '�������ɵģ����ѡ������ҽ����ѡ���˱���ԭ�򣩣���Ҫ����·����Ŀ
        If mlngFun = 0 Then
            For i = 1 To .Rows - 1
                If .RowData(i) = ִ�з�ʽ.T1�������� Then
                    If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                        If InStr("," & str��ĿIDs & ",", "," & .TextMatrix(i, mcol("ID")) & ",") = 0 Then
                            str��ĿIDs = str��ĿIDs & "," & .TextMatrix(i, mcol("ID"))
                        End If
                    End If
                End If
            Next
        End If
        
        '����Ҫ����ҽ������ĿID����ǰ�������ɳ�������Ŀ����������,�������ɵĶ�ѡ���˲����ɵĲ�������
        For i = 1 To .Rows - 1
            .TextMatrix(i, mcol("�ظ���Ŀ")) = ""
            If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                If InStr(str��ĿIDs, "," & .TextMatrix(i, mcol("ID")) & ",") = 0 Then
                    strTmp = Trim(.TextMatrix(i, mcol("ҽ������ID")))
                    If strTmp <> "" Then
                        '��Ŀ������ͬ�Ҷ�Ӧ��ҽ��Ҳ��ͬ�����ظ�����
                        strAgaignTmp = Trim(.TextMatrix(i, mcol("������ĿID")))
                        '��¼ҽ�����ж��ظ�
                        If InStr(strAgain & vbCrLf, vbCrLf & strAgaignTmp & vbCrLf) = 0 Or strAgaignTmp = "" Then
                            strAgain = strAgain & vbCrLf & strAgaignTmp
                            If strAgaignTmp <> "" Then
                                colItem.Add .TextMatrix(i, mcol("ID")) & vbCrLf & .TextMatrix(i, mcol("��Ŀ����")), strAgaignTmp
                            End If
                            arrBaby = Split(GetBabyIndex(.TextMatrix(i, mcol("Ӥ��"))), "|")
                            For n = LBound(arrBaby) To UBound(arrBaby)
                                strBB = arrBaby(n) & ":" & .TextMatrix(i, mcol("ID"))
                                If InStr(strTmp, ",") = 0 Then
                                    strBaby = strTmp & ":" & strBB
                                Else
                                    strBaby = Replace(strTmp, ",", ":" & strBB & ",") & ":" & strBB
                                End If
                                
                                strIDs = strIDs & "," & strBaby
                            Next
                        Else
                            .TextMatrix(i, mcol("�ظ���Ŀ")) = colItem(strAgaignTmp)
                        End If
                    End If
                End If
                If blnHaveDoc = False Then
                    If Trim(.TextMatrix(i, mcol("�ļ�ID"))) <> "" Then
                        blnHaveDoc = True
                    End If
                End If
            End If
        Next
    End With
    strIDs = Mid(strIDs, 2)             'ҽ������ID:Ӥ�����:·����ĿID,...������227:0:38,335:1:69
    
    If blnHaveDoc Then
        If InStr(GetInsidePrivs(pסԺ��������), ";������д;") = 0 Then
            MsgBox "��û�в�����д��Ȩ�ޣ��������ɰ���������·����Ŀ��", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
    End If
    
    '����ҽ����ȱʡ��ʼִ��ʱ��
    DatCurr = mdatDur
    
    If strIDs <> "" Then    'ȫ������ִ�е���Ŀʱ������ҽ������Ҫ����·��ִ����Ŀ
        If InStr(GetInsidePrivs(p����ҽ���´�), ";ҽ���´�;") = 0 Then
            MsgBox "��û��ҽ���´��Ȩ�ޣ��������ɰ���ҽ����·����Ŀ��", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        '���ʱ��
        If Format(DatCurr, "YYYY-MM-DD") > Format(dtpAdviceTime.Value, "YYYY-MM-DD") Or Format(dtpAdviceTime.Value, "YYYY-MM-DD") > Format(DatCurr + mlng·��ҽ������, "YYYY-MM-DD") Then
            MsgBox "�����ٴ�·����ҽ�������ڵ�ǰ���ں���ǰ������֮�䣬��ǰ������ǰ" & mlng·��ҽ������ & "�졣", vbInformation, gstrSysName
            If dtpAdviceTime.Enabled And dtpAdviceTime.Visible Then
                dtpAdviceTime.SetFocus
            End If
            Exit Sub
        End If
        
        Me.Hide
        If gobjKernel.ShowOutAdviceEdit(mfrmParent, 0, 1, mPati.����ID, mPati.�Һ�NO, strIDs, CDate(dtpAdviceTime.Value), arrSQL, strAdviceOfItem, rsLastAdvice, DatCurr, str·����ĿIDs, mclsMipModule, , mPati.����ID) = False Then
            Unload Me
            Exit Sub
        End If
        '�����ҽ�䷽�޸ĵ�ζ���������õı�׼�����������˱���ԭ�����ϱ���ԭ��
        If str·����ĿIDs <> "" Then
            '���һ����Ŀ����������ԭ����ȡ��һ��
            On Error Resume Next
            For i = 0 To UBound(Split(str·����ĿIDs, ","))
                colPathItems.Add Split(Split(str·����ĿIDs, ",")(i), "|")(1), "_" & Split(Split(str·����ĿIDs, ",")(i), "|")(0)
            Next
            For i = 1 To vsItem.Rows - 1
                strTmp = ""
                If vsItem.TextMatrix(i, mcol("ID")) & "" <> "" Then
                    strTmp = colPathItems("_" & vsItem.TextMatrix(i, mcol("ID")))
                    If strTmp <> "" Then
                        vsItem.Cell(flexcpData, i, mcol("����ԭ��")) = strTmp
                    End If
                End If
            Next
            On Error GoTo 0
        End If
    End If
    
    Call SaveData(arrSQL, strAdviceOfItem, lng����)
    '����ҽ�����
    Call ModifyAdviceSerialNum
    mblnOK = True
    Unload Me
End Sub

Private Sub ModifyAdviceSerialNum()
'���ܣ���������ҽ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String

    On Error GoTo errH
    Screen.MousePointer = 11
    strSql = "Select Count(*) as Num From (Select ���,Count(ID) From ����ҽ����¼ Where ����ID=[1] And �Һŵ�=[2] Having Count(ID)>1 Group by ���)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯ����ҽ������", mPati.����ID, mPati.�Һ�NO)

    If rsTmp.EOF Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    If NVL(rsTmp!Num, 0) = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    strSql = "ZL_����ҽ����¼_�������(NULL,NULL," & mPati.����ID & ",'" & mPati.�Һ�NO & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "����ҽ�����")

    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetUsedAdvice(ByVal lngִ��ID As Long, ByVal lng��ĿID As Long) As ADODB.Recordset
'����:��������ʱ,���ص�ǰ��Ŀ��Ӧ����Чҽ����¼
    Dim strSql As String
    
    strSql = " Select [1] as ��ĿID, a.����ҽ��id, Nvl(b.���id, b.Id) As ��id, b.������Ŀid " & vbNewLine & _
             " From ��������·��ҽ�� A, ����ҽ����¼ B" & vbNewLine & _
             " Where a.����ҽ��id = b.Id And a.·��ִ��id = [2] " & vbNewLine & _
             " Order By b.���"
    On Error GoTo errH
    
    Set GetUsedAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng��ĿID, lngִ��ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetLastEvaluate(strLastVariation As String, str����� As String, str������ As String)
'���ܣ�������һ����������Ϣ
    Dim strSql As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "Select ����ԭ��,���������,������ From ��������·������ Where ·����¼ID=[1] And ����=[2] And �׶�ID=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ����, mPP.��ǰ�׶�ID)
    If rsTmp.RecordCount > 0 Then
        strLastVariation = rsTmp!����ԭ�� & ""
        str����� = rsTmp!��������� & ""
        str������ = rsTmp!������ & ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetFirstType()
'���ܣ����һ����Ŀ��û�У�������ݿ���ȡ��һ������
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select ���� from ����·������ Where ·��ID=[1] and �汾��=[2] and ���=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ȡ��һ������", mPP.·��ID, mPP.�汾��)
    If rsTmp.RecordCount > 0 Then
        GetFirstType = rsTmp!���� & ""
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveData(ByVal arrSQL As Variant, ByVal strAdviceOfItem As String, ByRef lng���� As Long)
'���ܣ�����·����Ŀ
'������strAdviceOfItem=·����Ŀ��ҽ��ID�Ķ�Ӧ,����38:1983,69:1978
    Dim colSQL As New Collection, colDoc As New Collection, blnTrans As Boolean, colNewDoc As New Collection
    Dim strSql As String, i As Long, j As Long, l As Long, k As Long, lngBaby As Long
    Dim strDate As String, strAddDate As String, strAdviceIDs As String, strFileIDs As String, strFileID As String
    Dim str���˲���IDs As String, strEMRID As String, lng����ID As Long, strVariation As String, strBaby As String
    Dim strFileIDsTmp As String, strFiles As String
    Dim arrItem As Variant, lng��� As Long
    Dim blnIsSend As Boolean   '�ж��û��Ƿ�ѡ����Ŀ
    Dim strLastVariation As String
    Dim str����� As String
    Dim str������ As String
    Dim varFilter As Variant
    Dim AddDate As Date
    Dim strFirstType As String
    Dim strEPR As String
    Dim blnAgain As Boolean
    Dim strAgaignTmp As String
    Dim strAgain As String
    Dim colItemName As New Collection
    Dim blnDef As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim arrtmp As Variant
    
    AddDate = zlDatabase.Currentdate
    strAddDate = "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrItem = Split(strAdviceOfItem, ",")
    
    '���һ����Ŀ��û�У�������ݿ���ȡ��һ������
    If vsItem.TextMatrix(1, mcol("����")) = "" Then
        strFirstType = GetFirstType
    Else
        strFirstType = vsItem.TextMatrix(1, mcol("����"))
    End If

    lng���� = mlng����
    
    strDate = "To_Date('" & Format(mdatDur, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    '��ǰ����
    If mlngʱ����� = 2 Then
        k = mlng��ǰ����
    Else
        k = mlng��ǰ���� + 1
    End If
     '�жϵ�ǰѡ��Ľ׶ο�ʼ�����Ƿ���ڴ��������������м��������δ�����κ���Ŀ������
    If lng���� > k Then
        varFilter = mrsPhase.Filter
        Call GetLastEvaluate(strLastVariation, str�����, str������)
        For i = k To lng���� - 1
            '����
            mrsPhase.Filter = "��ʼ����=" & i & " And ��ID = 0"
            If Not mrsPhase.EOF Then
                strSql = "Zl_��������·������_Insert(1," & mPati.����ID & "," & mPati.�Һ�ID & ",NULL," & mPati.����ID & "," & _
                        mPP.����·��ID & "," & mrsPhase!ID & _
                        "," & strDate & "," & mlng��ǰ���� & _
                        ",'" & strFirstType & "',Null" & _
                        ",Null,Null,Null,'" & UserInfo.���� & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & _
                        "','YYYY-MM-DD HH24:MI:SS'),'δ�����κ���Ŀ','�Ѿ�ִ��|1" & vbTab & "�Ѿ�ִ��')"
                colSQL.Add strSql, "C" & colSQL.count + 1
                
                '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
                AddDate = AddDate + 1 / 24 / 60 / 60
                '����
                strSql = "Zl_��������·������_Insert(1," & mPP.����·��ID & "," & mrsPhase!ID & _
                        "," & strDate & "," & mlng��ǰ���� & ",'" & _
                        str������ & "',1,'','" & UserInfo.���� & "','" & str����� & "','" & strLastVariation & "',1,Null,1)"
                        
                colSQL.Add strSql, "C" & colSQL.count + 1
            End If
        Next
        mrsPhase.Filter = varFilter
    End If
        
    With vsItem
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Or .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 And mlngFun = 0 And .RowData(i) = ִ�з�ʽ.T1�������� Then
                strBaby = GetBabyIndex(.TextMatrix(i, mcol("Ӥ��")))
                
                strAdviceIDs = ""
                str���˲���IDs = ""
                strFileIDs = ""
                strVariation = ""
                blnAgain = False
                strFileIDsTmp = ""
                blnDef = False
                
                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                    If Val(.TextMatrix(i, mcol("ҽ������ID"))) <> 0 Then
                        strAgaignTmp = Trim(.TextMatrix(i, mcol("������ĿID")))
                        If InStr(strAgain & vbCrLf, vbCrLf & strAgaignTmp & vbCrLf) = 0 Then
                            For j = 0 To UBound(arrItem)
                                If Split(arrItem(j), ":")(0) = .TextMatrix(i, mcol("ID")) Then  '·����ĿID
                                    strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  'ҽ��ID
                                ElseIf .TextMatrix(i, mcol("�ظ���Ŀ")) <> "" Then
                                    '������ظ���Ŀ�������Ŀ������ͬ�����ظ�������ͬ��Ŀ��������Ʋ�ͬ����������Ŀָ����ͬҽ��
                                    If Split(arrItem(j), ":")(0) = Split(.TextMatrix(i, mcol("�ظ���Ŀ")), vbCrLf)(0) Then
                                        If .TextMatrix(i, mcol("��Ŀ����")) = Split(.TextMatrix(i, mcol("�ظ���Ŀ")), vbCrLf)(1) Then
                                            blnAgain = True
                                        Else
                                            strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  'ҽ��ID
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            '����̳е���Ŀ��������ظ�����ֻ����һ����Ŀ
                            If .TextMatrix(i, mcol("��Ŀ����")) = colItemName("C" & strAgaignTmp) Then
                                blnAgain = True
                            Else
                                For j = 0 To UBound(arrItem)
                                    If Split(arrItem(j), ":")(0) = .TextMatrix(i, mcol("ID")) Then  '·����ĿID
                                        strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  'ҽ��ID
                                    ElseIf .TextMatrix(i, mcol("�ظ���Ŀ")) <> "" Then
                                        '������ظ���Ŀ�������Ŀ������ͬ�����ظ�������ͬ��Ŀ��������Ʋ�ͬ����������Ŀָ����ͬҽ��
                                        If Split(arrItem(j), ":")(0) = Split(.TextMatrix(i, mcol("�ظ���Ŀ")), vbCrLf)(0) Then
                                            If .TextMatrix(i, mcol("��Ŀ����")) = Split(.TextMatrix(i, mcol("�ظ���Ŀ")), vbCrLf)(1) Then
                                                blnAgain = True
                                            Else
                                                strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  'ҽ��ID
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            blnDef = True
                        End If
                        If Not blnAgain And Not blnDef Then
                            strAgain = strAgain & vbCrLf & strAgaignTmp
                            colItemName.Add .TextMatrix(i, mcol("��Ŀ����")), "C" & strAgaignTmp
                        End If
                        strAdviceIDs = Mid(strAdviceIDs, 2)
                        
                        '�����ҩ�䷽�޸Ĺ�����д�˱���ԭ���򱣴浽������Ŀ��
                        If .Cell(flexcpData, i, mcol("����ԭ��")) <> "" Then
                            strVariation = .Cell(flexcpData, i, mcol("����ԭ��"))
                        End If
                    End If
                    
                    strEPR = Trim(.TextMatrix(i, mcol("�ļ�ID")))     '�����ж��
                    If strEPR <> "" Then
                       strFiles = Split(strEPR, "|")(0)  '�ɰ�
                       strEPR = ""
                       If strFiles <> "" Then
                            arrtmp = Split(strBaby, "|")
                            For l = LBound(arrtmp) To UBound(arrtmp)
                                strEMRID = "": strFileIDsTmp = ""
                                For j = 0 To UBound(Split(strFiles, ","))
                                    strFileID = Split(strFiles, ",")(j)
                                     '����ʼ�ղ������ظ��ģ�һ�������ļ�ֻ����һ��
                                    lngBaby = CLng(arrtmp(l) & "")
                                    If InStr(strEPR & ",", "," & lngBaby & "_" & strFileID & ",") = 0 Then
                                        lng����ID = zlDatabase.GetNextId("���Ӳ�����¼")
                                        strEMRID = strEMRID & "," & lng����ID
                                        colDoc.Add lng����ID & ":" & lngBaby & ":" & mEditType("C" & strFileID), "C" & (colDoc.count + 1)
                                        strFileIDsTmp = strFileIDsTmp & "," & strFileID
                                        strEPR = strEPR & "," & lngBaby & "_" & strFileID
                                    End If
                                Next
                                str���˲���IDs = str���˲���IDs & "|" & Mid(strEMRID, 2)
                                strFileIDs = strFileIDs & "|" & Mid(strFileIDsTmp, 2)
                            Next
                            str���˲���IDs = Mid(str���˲���IDs, 2)
                            strFileIDs = Mid(strFileIDs, 2)
                        End If

                        If str���˲���IDs = "" Then
                            blnAgain = True
                        End If
                    End If
                Else
                    strVariation = .Cell(flexcpData, i, mcol("����ԭ��"))
                End If
                
                If Not blnAgain Then
                    lng��� = lng��� + 1
                    strSql = "Zl_��������·������_Insert(" & lng��� & "," & mPati.����ID & "," & mPati.�Һ�ID & ",'" & strBaby & "'," & mPati.����ID & "," & _
                        mPP.����·��ID & "," & mrsPhase!ID & "," & strDate & "," & mlng��ǰ���� & ",'" & .TextMatrix(i, mcol("����")) & "'," & .TextMatrix(i, mcol("ID")) & _
                        ",'" & strAdviceIDs & "','" & strFileIDs & "','" & str���˲���IDs & "'" & _
                        ",'" & UserInfo.���� & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null,Null,Null,Null,'" & strVariation & "')"
                    colSQL.Add strSql, "C" & colSQL.count + 1
                    blnIsSend = True
                End If
            End If
        Next
    End With
    '���û�й�ѡ�κ���Ŀ��������һ���������Ŀ��δ�����κ���Ŀ
    If Not blnIsSend Then
        If mlngFun = 0 Then
            lng��� = lng��� + 1
            strSql = "Zl_��������·������_Insert(" & lng��� & "," & mPati.����ID & "," & mPati.�Һ�ID & ",NULL," & mPati.����ID & "," & _
                    mPP.����·��ID & "," & mrsPhase!ID & "," & strDate & "," & mlng��ǰ���� & ",'" & strFirstType & "',Null" & _
                    ",Null,Null,Null,'" & UserInfo.���� & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'δ�����κ���Ŀ','�Ѿ�ִ��|1" & vbTab & "�Ѿ�ִ��')"
            colSQL.Add strSql, "C" & colSQL.count + 1
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        If mlngFun = 3 Then
            strSql = "Zl_��������·������_Delete(" & mlngִ��ID & ",1)"
            Debug.Print strSql & vbCrLf
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        End If
        
        '1.�Ȳ���ҽ��,��Ϊ����·��ҽ�������
        For i = 0 To UBound(arrSQL)
            Debug.Print CStr(arrSQL(i)) & vbCrLf
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        '2.��������·�����ݣ��Լ������ļ�����
        For i = 1 To colSQL.count
            Debug.Print colSQL("C" & i) & vbCrLf
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
        '3.���������ļ�RTF����
        For i = 1 To colDoc.count
            arrItem = Split(colDoc("C" & i), ":")
            If arrItem(2) = 0 Or arrItem(2) = 1 Then     'ȫ�ı༭��ʽ�Ĳ���
                lng����ID = (arrItem(0))
                Call ReadRTFData(lng����ID, edtEditor)
                Call SaveRTFData(lng����ID, mPati.����ID, mPati.�Һ�ID, Val(arrItem(1)), edtEditor, 1)
            End If
        Next
    gcnOracle.CommitTrans: blnTrans = False

    Exit Sub
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
