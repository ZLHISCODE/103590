VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDiagEdit 
   BackColor       =   &H00EFF0E0&
   Caption         =   "���ѡ�񼰱༭"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10755
   Icon            =   "frmDiagEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleMode       =   0  'User
   ScaleWidth      =   10974.49
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picZY 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   360
      ScaleHeight     =   3855
      ScaleWidth      =   9615
      TabIndex        =   5
      Top             =   480
      Width           =   9615
      Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
         Height          =   3675
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9495
         _cx             =   16748
         _cy             =   6482
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
         BackColorFixed  =   14811105
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDiagEdit.frx":6852
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
      End
   End
   Begin VB.PictureBox picXY 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   9615
      TabIndex        =   3
      Top             =   600
      Width           =   9615
      Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
         Height          =   3465
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9495
         _cx             =   16748
         _cy             =   6112
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
         BackColorFixed  =   14811105
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDiagEdit.frx":6A47
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
      End
   End
   Begin VB.Frame fraInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   7920
      Begin VB.OptionButton optInput 
         BackColor       =   &H00EFF0E0&
         Caption         =   "������ϱ�׼����(&1)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   0
         Left            =   3840
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   37
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton optInput 
         BackColor       =   &H00EFF0E0&
         Caption         =   "���ݼ�����������(&2)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   1
         Left            =   5880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   37
         Width           =   2010
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   10755
      TabIndex        =   8
      Top             =   5580
      Width           =   10755
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8520
         TabIndex        =   10
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   7200
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.Image imgButtonNew 
         Height          =   240
         Left            =   720
         Picture         =   "frmDiagEdit.frx":6C4E
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgButtonDel 
         Height          =   240
         Left            =   0
         Picture         =   "frmDiagEdit.frx":71D8
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin XtremeSuiteControls.TabControl tbcMain 
      Height          =   4095
      Left            =   60
      TabIndex        =   7
      Top             =   360
      Width           =   9735
      _Version        =   589884
      _ExtentX        =   17171
      _ExtentY        =   7223
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmDiagEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ���
Private mlng����ID As Long
Private mlng����ID As Long
Private mint������Դ As Integer
Private mlngCur��ʶ As Long
Private mlng����ID As Long
Private mstr������ As String
Private mstr���IDs As String
Private mstr���s As String
Private mlng��ҽ��ID As String
Private mblnOK As Boolean
'��������
Private mint��ҽ������� As Integer
Private mint��ҽ������� As Integer
Private mstrPrivs As String
Private mblnChange As Boolean
Private mlngPathState As Long
Private mlngDiagnosisType As Long
Private mstrPathDiag As String
Private mblnIsPathOutTime As Boolean
Private mstr�Ա� As String
Private mblnReturn As Boolean
Private mint���� As Integer
Private mstr������� As String
Private mstrLike As String
Private mint���� As Integer
Private mstr�Һŵ� As String
Private mbln���� As Boolean '�����Ƿ���й�����
Private mlng�����ж� As Long
Private mlng������� As Long
Private mbln��ҽ As Boolean
Private mbytSize As Byte

Private Const M_LNG_PסԺҽ��վ = 1261
Private Const M_LNG_P����ҽ��վ = 1260
Private Const M_LNG_SYS = 100
Private Const ColorUnEditCell = &H8000000B  '����ɫ
Private mrsAdvice As ADODB.Recordset

Private Enum COL������
    col������� = 0
    col���� = 1
    col��ϱ��� = 2
    Col������� = 3
    col��ҽ֤�� = 4
    col����ʱ�� = 5
    col��ע = 6
    col��Ժ���� = 7
    col��Ժ��� = 8
    col�Ƿ�δ�� = 9
    col�Ƿ����� = 10
    col���� = 11
    colDel = 12
    col���ID = 13
    col����ID = 14
    col���� = 15 '1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
    
    colZY���� = 9
    colZY���� = 10
    colZYDel = 11
    colzy���ID = 12
    colzy����ID = 13
    colzy֤��ID = 14
    colzy���� = 15
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng��ʶID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int������Դ As Integer, ByVal lng��������ID As Long, ByVal str������ As String, _
                    ByRef str���IDs As String, ByRef str���S As String, ByVal bytSize As Byte, Optional ByVal lngҽ����ID As Long) As Boolean
'������lng����ID=����ID
'      lng����ID=סԺ:��ҳID,����Һŵ�ID
'      int������Դ=1-���2-סԺ
'      lng��������ID=�������ڿ��ң����ʹ��
'      lng��ʶID =�������ָ������뵥�ı�ʶ�����ڱ�����Ӧ�����
'      str������=����Ա��������ϵǼ���
'      str���IDs=�����뵥��ص����ID,������ʱ���ID�Զ��ŷָ�
'      bytSize=0-9�����壬1-12������
'���أ� ShowDiagEdit= ��ȷ������ȡ��
'       str���S=������������ַ����������뵥ʹ��
'       str���IDs=�����뵥ѡ�����ص����ID,������ʱ���ID�Զ��ŷָ�
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mint������Դ = int������Դ
    mlngCur��ʶ = lng��ʶID
    mlng����ID = lng��������ID
    mstr������ = str������
    mstr���IDs = str���IDs
    mstr���s = str���S
    mlng��ҽ��ID = lngҽ����ID
    mbytSize = bytSize
    mstrPrivs = gobjComLib.GetPrivFunc(M_LNG_SYS, IIf(mint������Դ = 2, M_LNG_PסԺҽ��վ, M_LNG_P����ҽ��վ))
    Show 1, frmParent

    str���IDs = mstr���IDs
    str���S = mstr���s
    ShowMe = mblnOK

End Function

Private Sub cmdCancel_Click()
    If vsDiagXY.Tag = "" Or vsDiagZY.Tag = "" And vsDiagZY.Visible Then
        If MsgBox("�˳������������޸Ľ�������Ч���Ƿ��˳���", vbYesNo + vbDefaultButton2 + vbInformation, Me.Caption) = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If CheckData() Then
        Call SaveData
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'סԺ��ҳ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngColWidth As Long
    
    On Error GoTo errH
    
    mblnOK = False
    mlngPathState = -1
    mstrLike = IIf(Val(gobjComLib.zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    
    strSQL = "Select A.����,Nvl(A.·��״̬,-1) ·��״̬" & _
        " From ������ҳ A" & _
        " Where A.����ID=[1] And A.��ҳID=[2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
    If Not rsTmp.EOF Then
        mint���� = NVL(rsTmp!����, 0)
        'mlngPathState=-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
        mlngPathState = Val(rsTmp!·��״̬ & "")
    End If
    
    strSQL = "Select 1 From ���������¼  A Where  A.����ID=[1] And A.��ҳID=[2] "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
    mbln���� = Not rsTmp.EOF
    '������Ϣ����
    '---------------------------------------------------------------
    
    strSQL = "Select �Ա� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    mstr�Ա� = NVL(rsTmp!�Ա�)
    If mint������Դ <> 2 Then
        strSQL = "Select NO From ���˹Һż�¼ Where ����id = [1] And ID = [2]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
        If Not rsTmp.EOF Then
            mstr�Һŵ� = rsTmp!NO & ""
        End If
    End If
    '������뷽ʽ
    mstr������� = gobjComLib.zlDatabase.GetPara(65, M_LNG_SYS, , "11")
    mint���� = Val(gobjComLib.zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    mlng�����ж� = Val(gobjComLib.zlDatabase.GetPara("�����ж����", M_LNG_SYS, M_LNG_PסԺҽ��վ, 2) & "")
    mlng������� = Val(gobjComLib.zlDatabase.GetPara("������ϼ��", M_LNG_SYS, M_LNG_PסԺҽ��վ, 2) & "")
    
    If mlngPathState <> -1 Then
        'ֻ������ҳ���������ϣ���ǰû��ģ�ȱʡ���������ڡ���ҽ��Ժ��ϡ�
        strSQL = "Select Nvl(�������,2) as �������,NVL(����ID,0) As ����ID,NVL(���ID,0) as ���ID,״̬ From �����ٴ�·�� Where ����ID=[1] And ��ҳID=[2] And (�����Դ = 3 or �����Դ is null) Order By ����ʱ��"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
        If rsTmp.RecordCount > 0 Then
            mlngDiagnosisType = rsTmp!�������
            '����ж���·������ȡ��һ����״̬
            If rsTmp.RecordCount >= 2 Then mlngPathState = Val(rsTmp!״̬ & "")
            rsTmp.MoveNext
            Do While Not rsTmp.EOF
                mstrPathDiag = mstrPathDiag & "," & rsTmp!������� & "|" & rsTmp!����ID & "|" & rsTmp!���ID
                rsTmp.MoveNext
            Loop
            mstrPathDiag = Mid(mstrPathDiag, 2)
        Else
            mlngDiagnosisType = 0
        End If
        '���·����ʱ���Ƿ�ȳ�Ժ��ϼ�¼ʱ���()ȡ��һ��·��
        If mlngPathState = 2 Then
            strSQL = "Select Sign(Nvl(a.����ʱ��, Null)-Nvl(b.��¼����, Sysdate)) As �ж�" & vbNewLine & _
                    "From �����ٴ�·�� A, (Select ����id, ��ҳid, ��¼���� From ������ϼ�¼ Where ��¼��Դ = 3 And ��ϴ��� = 1 And ������� = [3]) B" & vbNewLine & _
                    " Where a.����id = b.����id(+) And a.��ҳid = b.��ҳid(+) And a.����ID=[1] And A.��ҳID=[2]" & _
                    " and a.����ʱ��=(Select Min(����ʱ��) From �����ٴ�·�� Where ����ID=[1] and ��ҳID=[2])"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID, IIf(mlngDiagnosisType > 10, 13, 3))
            If rsTmp.RecordCount > 0 Then
                mblnIsPathOutTime = Val(rsTmp!�ж� & "") = 1
            Else
                mblnIsPathOutTime = False
            End If
        End If
    End If
    
    strSQL = "Select 1 From ��������˵�� Where ��������='��ҽ��' And ����ID=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    mbln��ҽ = Not rsTmp.EOF


    If mbln��ҽ Then
        tbcMain.PaintManager.Color = xtpTabColorOffice2003
        tbcMain.PaintManager.ColorSet.ControlFace = &HEFF0E0
        Call tbcMain.InsertItem(0, "��ҽ���", picXY.hwnd, 0)
        Call tbcMain.InsertItem(1, "��ҽ���", picZY.hwnd, 0)
        tbcMain(0).Selected = True
    Else
        tbcMain.Enabled = False
        tbcMain.Visible = False
        vsDiagZY.Visible = False
        picZY.Visible = False
        vsDiagZY.Enabled = False
        If mint������Դ = 2 Then
            If Val(gobjComLib.zlDatabase.GetPara("��ҽ�������", M_LNG_SYS, M_LNG_PסԺҽ��վ, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "��������") > 0)) = 0 Then
                optInput(0).value = True
            Else
                optInput(1).value = True
            End If
        ElseIf mint������Դ = 1 Then
            optInput(Val(gobjComLib.zlDatabase.GetPara("�����������", M_LNG_SYS, M_LNG_P����ҽ��վ, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "��������") > 0))).value = True
        End If
    End If
    
    If mstr���IDs = "" Then
        With grsDiagConn
            .Filter = "��ʶID=" & mlngCur��ʶ
            .Sort = "���ID"
            Do While Not .EOF
                mstr���IDs = mstr���IDs & IIf(mstr���IDs = "", "", ",") & !���ID
                .MoveNext
            Loop
        End With
    End If
    
    Call LoadData
    
    Call SetVSColHidden
    
    If mint������Դ = 2 Then
        vsDiagXY.ColWidth(Col�������) = vsDiagXY.ColWidth(Col�������) - 1000
        Me.Width = Me.Width + 1500
        Me.Height = Me.Height + 927
    Else
        vsDiagXY.ColWidth(col����ʱ��) = vsDiagXY.ColWidth(col����ʱ��) + 400
        vsDiagZY.ColWidth(col����ʱ��) = vsDiagZY.ColWidth(col����ʱ��) + 400
        vsDiagXY.ColWidth(Col�������) = vsDiagXY.ColWidth(Col�������) - 200
        vsDiagZY.ColWidth(Col�������) = vsDiagZY.ColWidth(Col�������) - 200
    End If
    
    If mbytSize = 0 Then
        Me.Width = Me.Width - 2000
        Me.Height = Me.Height - 1236
        vsDiagXY.ColWidth(Col�������) = vsDiagXY.ColWidth(Col�������) + 600
        vsDiagZY.ColWidth(Col�������) = vsDiagZY.ColWidth(Col�������) + 1000
    End If
    
    If Not mbln��ҽ Then
        Me.Width = Me.Width + 400
    End If
    
    Call SetPublicFontSize(mbytSize)
    Call gobjComLib.zlControl.VSFSetFontSize(vsDiagXY, IIf(mbytSize = 0, 9, 12))
    Call gobjComLib.zlControl.VSFSetFontSize(vsDiagZY, IIf(mbytSize = 0, 9, 12))
    lngColWidth = 270 '��ֹ�п����ʱ����ɾ����ť���ֺ�ɫ��Ӱ
    vsDiagXY.ColWidth(colDel) = lngColWidth
    vsDiagXY.ColWidth(col����) = lngColWidth
    vsDiagZY.ColWidth(colZYDel) = lngColWidth
    vsDiagZY.ColWidth(colZY����) = lngColWidth
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Resize()
    If Me.Width > 20000 Then
        Me.Width = 20000
    End If
    If Me.Height > 12000 Then
        Me.Height = 12000
    End If
    
    If Me.Width < 6000 Then
        Me.Width = 6000
    End If
    
    If Me.Height < 5000 Then
        Me.Height = 5000
    End If
    fraInput.Width = Me.Width
    fraInput.Left = 0

    tbcMain.Top = fraInput.Top + fraInput.Height - IIf(mbln��ҽ, 120, 180)
    tbcMain.Height = picBottom.Top - tbcMain.Top
    tbcMain.Width = Me.Width - tbcMain.Left - 100
    If mbln��ҽ Then
        tbcMain.Top = tbcMain.Top - 200
        fraInput.Left = tbcMain.Left + IIf(mbytSize = 0, 1840, 2320)
    End If
    optInput(1).Left = Me.Width - fraInput.Left - optInput(1).Width - 240
    optInput(0).Left = optInput(1).Left - optInput(0).Width - 100
    
    picZY.Top = tbcMain.Top + 210
    picZY.Height = picBottom.Top - picZY.Top
    picZY.Width = tbcMain.Width - 180
    
    picXY.Top = picZY.Top
    picXY.Height = picZY.Height
    picXY.Width = picZY.Width
    
    'ȷ��ȡ����ťλ������
    If mbytSize = 1 Then
        cmdCancel.Top = 90
        cmdOK.Top = 90
    End If
    cmdCancel.Left = Me.Width - cmdCancel.Width - 360
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 180
End Sub

Private Sub picXY_Resize()
    vsDiagXY.Height = picXY.Height - 300
    vsDiagXY.Width = picXY.Width
End Sub


Private Sub picZY_Resize()
    vsDiagZY.Height = picZY.Height - 300
    vsDiagZY.Width = picZY.Width
End Sub

Private Sub SetVSColHidden()
'��������VS���еĿɼ���
    With vsDiagXY
        .ColHidden(col�������) = mint������Դ = 1
        .ColHidden(col��ҽ֤��) = True
        .ColHidden(col��Ժ����) = mint������Դ = 1
        .ColHidden(col��Ժ���) = mint������Դ = 1
        .ColHidden(col�Ƿ�δ��) = mint������Դ = 1
        .ColHidden(col����ʱ��) = mint������Դ <> 1
    End With
    
    With vsDiagZY
        .ColHidden(col�������) = mint������Դ = 1
        .ColHidden(col��Ժ����) = mint������Դ = 1
        .ColHidden(col��Ժ���) = mint������Դ = 1
        .ColHidden(col����ʱ��) = mint������Դ <> 1
        .ColHidden(colZY����) = mint������Դ <> 1
    End With
End Sub



Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mint������Դ = 2 Then
        If Item.Index = 0 Then
            If Val(gobjComLib.zlDatabase.GetPara("��ҽ�������", M_LNG_SYS, M_LNG_PסԺҽ��վ, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "��������") > 0)) = 0 Then
                optInput(0).value = True
            Else
                optInput(1).value = True
            End If

        Else
            If Val(gobjComLib.zlDatabase.GetPara("��ҽ�������", M_LNG_SYS, M_LNG_PסԺҽ��վ, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "��������") > 0)) = 0 Then
                optInput(0).value = True
            Else
                optInput(1).value = True
            End If
        End If
    ElseIf mint������Դ = 1 Then
        optInput(Val(gobjComLib.zlDatabase.GetPara("�����������", M_LNG_SYS, M_LNG_P����ҽ��վ, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "��������") > 0))).value = True
    End If
    Call Form_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mint������Դ = 2 Then
        Call gobjComLib.zlDatabase.SetPara("��ҽ�������", IIf(optInput(0).value, 0, 1), M_LNG_SYS, M_LNG_PסԺҽ��վ, InStr(mstrPrivs, "��������") > 0)
        Call gobjComLib.zlDatabase.SetPara("��ҽ�������", IIf(optInput(1).value, 0, 1), M_LNG_SYS, M_LNG_PסԺҽ��վ, InStr(mstrPrivs, "��������") > 0)
    ElseIf mint������Դ = 1 Then
        Call gobjComLib.zlDatabase.SetPara("�����������", IIf(optInput(1).value, 0, 1), M_LNG_SYS, M_LNG_P����ҽ��վ, InStr(mstrPrivs, "��������") > 0)
    End If
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagXY
        If Col = col��Ժ��� Then
            '��Ҫ����ǻس��뿪:����ComboIndex,ȡ���༭ʱ����
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            If Not XYCellEditable(Row, col�Ƿ�δ��) Then
                .TextMatrix(Row, col�Ƿ�δ��) = ""
            End If
            .Tag = ""
        End If
        If Col = Col������� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '�ڵ���vsDiagXY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagXY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        Call vsDiagXY_AfterRowColChange(-1, -1, .Row, .Col)
        '�ж��Ƿ������޸�
        If vsDiagXY.Tag = "δ�޸�" And Col <> col���� Then
            vsDiagXY.Tag = ""
        End If
    End With
    
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long

    With vsDiagXY
        '���ͼƬ
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, col����) Is Nothing Then
                Set .Cell(flexcpPicture, i, col����) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, colDel) Is Nothing Then
               Set .Cell(flexcpPicture, i, colDel) = Nothing
            End If
        Next

        If Not XYCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing

            If NewCol = Col������� Then
                .ComboList = "..."
            ElseIf NewCol = col��Ժ��� Then
                .ComboList = .ColData(NewCol)
            ElseIf NewCol = col��Ժ���� Then
                If .TextMatrix(NewRow, 0) = "��Ժ���" Or .TextMatrix(NewRow, 0) = "�������" Or .TextMatrix(NewRow, 0) = "" Then
                    .ComboList = "��|�ٴ�δȷ��|�������|��"
                Else
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                End If
            ElseIf NewCol = col���� Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = colDel Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '��ʾͼƬ
            If NewCol <> col���� And .TextMatrix(NewRow, Col�������) <> "" And .TextMatrix(NewRow, 0) <> "��Ժ���" Then
                Set .Cell(flexcpPicture, NewRow, col����) = imgButtonNew.Picture
            End If
            '��ʾͼƬ
            If NewCol <> colDel And .RowData(NewRow) & "" = "" Then
                Set .Cell(flexcpPicture, NewRow, colDel) = imgButtonDel.Picture
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col���� Then Cancel = True
End Sub

Private Sub vsDiagXY_Click()
    With vsDiagXY
        If (.MouseCol = col���� Or .MouseCol = colDel) And .MouseRow >= .FixedRows Then
            If .MouseCol = col���� Then
                If .TextMatrix(.MouseRow, Col�������) = "" Or .TextMatrix(.MouseRow, 0) = "��Ժ���" Then Exit Sub
            End If

            .Select .MouseRow, .MouseCol
            Call vsDiagXY_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub
Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsDiagXY
        If Col = col��Ժ��� Then
            '��λ��ƥ����
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsDiagXY_DblClick()
    Call vsDiagXY_KeyPress(32)
    '����Ϊ���޸�
    If vsDiagXY.Col = col�Ƿ�δ�� Or vsDiagXY.Col = col�Ƿ����� Then
        If vsDiagXY.Tag = "δ�޸�" Then vsDiagXY.Tag = ""
    End If
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long

    With vsDiagXY
        If KeyCode = vbKeyF4 Then
            If .Col = Col������� Then
                Call gobjComLib.zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, Col�������) <> "" Then
                If .RowData(.Row) & "" <> "" Then Exit Sub
                If GetAdviceIDByDiag(Val(.Cell(flexcpData, .Row, col�Ƿ�����) & "")) <> "" Then Exit Sub
                
                If mlngPathState = 1 Or mlngPathState = 2 Then
                    If .TextMatrix(.Row, col�������) = "��Ժ���" And mlngDiagnosisType = 2 Or .TextMatrix(.Row, col�������) = "�������" And mlngDiagnosisType = 1 Then
                        If .TextMatrix(.Row, col�������) <> .TextMatrix(.Row - 1, col�������) Then
                            '��Ҫ��ϲ������
                            Exit Sub
                        End If
                    End If
                End If
                '�ϲ�·��
                If Not CheckMergePath(mlng����ID, mlng����ID, Val(.TextMatrix(.Row, col����)), Val(.TextMatrix(.Row, col����ID))) Then Exit Sub
                '����·������
                If mstrPathDiag <> "" And mlngPathState > 0 Then
                    If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, col����) & "|" & Val(.TextMatrix(.Row, col����ID)) & "|" & Val(.TextMatrix(.Row, col���ID)) & ",") > 0 Then
                        '������ϲ������
                        Exit Sub
                    End If
                End If
                If mlngPathState = 2 And mblnIsPathOutTime Then
                    If .TextMatrix(.Row, col�������) = "��Ժ���" And mlngDiagnosisType <= 2 Then
                        '������ɵĳ�Ժ��ϲ������
                        Exit Sub
                    End If
                End If
                If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    i = Val(.TextMatrix(.Row, col����))
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedCols, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, col����) = i

                    '�����ͬ�������������
                    If .TextMatrix(.Row, col�������) = "" Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, col�������) = "" Then
                                '��һ��Ϊ�ޱ����������ʱ�����ݲ����ƣ�����ǰ��Ϊ�б���ʱֻ�����
                                For i = .Row + 1 To .Rows - 1
                                    If Val(.TextMatrix(i, col����)) = Val(.TextMatrix(.Row, col����)) Then
                                        For j = .FixedCols To .Cols - 1
                                            .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                            .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                        Next
                                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                        .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                        .TextMatrix(i, col����) = Val(.TextMatrix(.Row, col����))
                                        .RowData(i - 1) = .RowData(i)
                                        .RowData(i) = Empty
                                        If i = .Rows - 1 Then
                                            If .TextMatrix(i, col�������) = "" Then .RemoveItem i
                                            Exit For
                                        ElseIf Val(.TextMatrix(i + 1, col����)) <> Val(.TextMatrix(i, col����)) Then
                                            If .TextMatrix(i, col�������) = "" Then .RemoveItem i
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    .Tag = ""
                End If
            ElseIf .TextMatrix(.Row, col�������) = "" Then
                .RemoveItem .Row
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiagXY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    With vsDiagXY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call XYEnterNextCell
        ElseIf KeyAscii = 32 And (.Col = col�Ƿ�δ�� Or .Col = col�Ƿ�����) Then
            If XYCellEditable(.Row, .Col) Then
                KeyAscii = 0
                If .Col = col�Ƿ����� Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "��", "")
                ElseIf .Col = col�Ƿ�δ�� Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "��", "")
                End If
                .Tag = ""
            End If
        Else
            If .Col = Col������� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagXY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagXY.EditSelStart = 0
    vsDiagXY.EditSelLength = gobjComLib.zlCommFun.ActualLen(vsDiagXY.EditText)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not XYCellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = col�Ƿ�δ�� Or Col = col�Ƿ����� Then
        Cancel = True '��ֱ�ӱ༭
    End If
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str�Ա� As String, lngRow As Long

    With vsDiagXY
        If Col = Col������� Then
            If optInput(0).value Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "1", mlng����ID, , True, False)
            Else
                '7-�����ж���Y-�����ж����ⲿԭ��6-������ϣ�M-������̬ѧ���룻������ϣ�D-ICD-10��������
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), mlng����ID, mstr�Ա�, True)
            End If
            If Not rsTmp Is Nothing Then
                .Tag = ""
                Call XYSetDiagInput(Row, rsTmp)
                Call XYEnterNextCell
            End If
        ElseIf Col = col���� Then
            lngRow = Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, col����) = .TextMatrix(Row, col����)
            .Cell(flexcpBackColor, lngRow, col��ϱ���) = ColorUnEditCell      '����ɫ
            
            .Row = lngRow: .Col = Col�������
            .ShowCell .Row, .Col
        ElseIf Col = colDel Then
            Call vsDiagXY_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True

        With vsDiagXY
            If Col = col��Ժ��� Then
                KeyAscii = 0
                If .ComboIndex <> -1 Then
                    '��ʱ.TextMatrix��δ����,����ȡComboItem
                    .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                    If Not XYCellEditable(Row, col�Ƿ�δ��) Then
                        .TextMatrix(Row, col�Ƿ�δ��) = ""
                    End If
                    Call XYEnterNextCell
                    .Tag = ""
                End If
            End If
        End With
    Else
        mblnReturn = False
    End If
End Sub

Private Sub XYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim bln�ֻ��̶� As Boolean

    With vsDiagXY
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    '�����ж�ѡ�����ʱ�Ĵ���
                    If lngRow = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, col����) = .TextMatrix(lngRow, col����)
                    End If
                    'ȷ����ǰ��ʾ��
                    If Val(.TextMatrix(lngRow + 1, col����)) = Val(.TextMatrix(lngRow, col����)) Then
                        For j = lngRow + 1 To .Rows - 1
                            If Val(.TextMatrix(j, col����)) = Val(.TextMatrix(lngRow, col����)) Then
                                lngRow = j
                                If .TextMatrix(j, Col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, Col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col����) = .TextMatrix(lngRow - 1, col����)
                        End If
                    Else
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, col����) = .TextMatrix(lngRow - 1, col����)
                    End If
                End If
                .TextMatrix(lngRow, col����) = 1
                .TextMatrix(lngRow, col��ϱ���) = "" & rsInput!����
                .TextMatrix(lngRow, Col�������) = "" & rsInput!����
                
                .Cell(flexcpData, lngRow, Col�������) = .TextMatrix(lngRow, Col�������)

                '�������ȷ������,����ݼ���ȷ�����
                If optInput(0).value Then
                    .TextMatrix(lngRow, col���ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col����ID) = ""
                    strSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
                Else
                    .TextMatrix(lngRow, col����ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col���ID) = ""
                    strSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
                End If
                On Error GoTo errH
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!��ĿID))
                If Not rsTmp.EOF Then
                    If optInput(0).value Then
                        .TextMatrix(lngRow, col����ID) = NVL(rsTmp!id)
                    Else
                        .TextMatrix(lngRow, col���ID) = NVL(rsTmp!id)
                    End If
                End If

                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col��ϱ���) = ""
            .TextMatrix(lngRow, Col�������) = .EditText
            .Cell(flexcpData, lngRow, Col�������) = .TextMatrix(lngRow, Col�������)
            .TextMatrix(lngRow, col���ID) = ""
            .TextMatrix(lngRow, col����ID) = ""
        End If

        .Cell(flexcpForeColor, 1, col�Ƿ�����, .Rows - 1, col�Ƿ�����) = vbRed
    End With
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str�Ա� As String, int������� As Integer
    Dim strInput As String, vPoint As POINTAPI

    With vsDiagXY
        If Col = Col������� Then
            '.Cell(flexcpData, Row, Col) <> ""�ų����лس�
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                .EditText = ""
                .Tag = ""
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call XYEnterNextCell
            ElseIf .TextMatrix(Row, col��ϱ���) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                '�жϼ���ǰ׺��������Ƿ������������ϱ���
                strInput = UCase(.EditText)
                strSQL = GetSQL(0, strInput, str�Ա�)
                On Error GoTo errH
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, _
                        Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), str�Ա�, mint���� + 1)
                If rsTmp.RecordCount <> 1 Then
                    '�����ڱ�׼������ǰ�����븽����Ϣ
                    .TextMatrix(Row, Col�������) = .EditText
                    .Tag = ""
                Else
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    .Tag = ""
                End If
                '������.Cell(flexcpData, Row, Col)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
            Else
                If Val(.TextMatrix(Row, col����)) = 1 Then
                    int������� = Val(Mid(mstr�������, 1, 1))
                Else
                    int������� = Val(Mid(mstr�������, 2, 1))
                End If
                If int������� = 0 Then int������� = 1

                strInput = UCase(.EditText)
                strSQL = GetSQL(0, strInput, str�Ա�)
                If int������� = 1 And gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    '�����ж��룺Y-�����ж����ⲿԭ�򣻲����������M-������̬ѧ���룻������ϣ�D-ICD-10��������
                    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", _
                        Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), str�Ա�, mint���� + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, _
                        Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), str�Ա�, mint���� + 1)
                        If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                    .Tag = ""
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn And rsTmp Is Nothing Then Call XYEnterNextCell '��������¼��ʱ���ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInput(0).value, "�������", "��������"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), str�Ա�, mint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And ((int������� = 2 Or int������� = 3 And mint���� <> 0)) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .Tag = ""
                            Call XYSetDiagInput(Row, rsTmp): .EditText = .Text
                            'If mblnReturn Then Call XYEnterNextCell    '�ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col����ʱ�� Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    .Tag = ""
                Else
                    MsgBox "��������ȷ�ķ���ʱ�䣬���磺""2012-12-21 00:00""��"
                    Cancel = True
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagZY
        If Col = col��Ժ��� Then
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            .Tag = ""
        End If
        If Col = Col������� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '�ڵ���vsDiagZY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagZY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        Call vsDiagZY_AfterRowColChange(-1, -1, .Row, .Col)
        If Col <> col���� Then .Tag = ""
    End With
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long

    With vsDiagZY

        '���ͼƬ
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, colZY����) Is Nothing Then
                Set .Cell(flexcpPicture, i, colZY����) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, colZYDel) Is Nothing Then
               Set .Cell(flexcpPicture, i, colZYDel) = Nothing
            End If
        Next

        If Not ZYCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing

            If NewCol = Col������� Then
                .ComboList = "..."
            ElseIf NewCol = col��ҽ֤�� Then
                If .TextMatrix(NewRow, Col�������) = "" Then
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                Else
                    .ComboList = "..."
                End If
            ElseIf NewCol = col��Ժ��� Then
                .ComboList = .ColData(NewCol)
            ElseIf NewCol = col��Ժ���� Then
                If .TextMatrix(NewRow, colzy����) = "13" Then
                    .ComboList = "��|�ٴ�δȷ��|�������|��"
                Else
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                End If
            ElseIf NewCol = colZY���� Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = colZYDel Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '��ʾͼƬ
            If NewCol <> colZY���� And .TextMatrix(NewRow, Col�������) <> "" And .TextMatrix(NewRow, 0) <> "��Ҫ���" Then
                Set .Cell(flexcpPicture, NewRow, colZY����) = imgButtonNew.Picture
            End If
            '��ʾͼƬ
            If NewCol <> colZYDel And .RowData(NewRow) & "" = "" Then
                Set .Cell(flexcpPicture, NewRow, colZYDel) = imgButtonDel.Picture
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colZY���� Then Cancel = True
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colZY���� Then Cancel = True
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str�Ա� As String, lngRow As Long
    Dim blnCancle As Boolean

    With vsDiagZY
        If Col = Col������� Then
            If optInput(1).value Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "2", mlng����ID, , True, False)
            Else
                'B-��ҽ��������
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "B", mlng����ID, mstr�Ա�, True)
            End If
            If Not rsTmp Is Nothing Then
                .Tag = ""
                Call ZYSetDiagInput(Row, rsTmp)
                Call ZYEnterNextCell
            End If
        ElseIf Col = col��ҽ֤�� Then
            If optInput(1).value Then
                '���������:�Ȳ��Ƿ��ж�Ӧ
                If Not Set��ҽ֤��(Row, Val(.TextMatrix(Row, colzy���ID))) Then
                    Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "Z", mlng����ID, mstr�Ա�, True)
                Else
                    Exit Sub
                End If
            Else
                'Z-��ҽ��������
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "Z", mlng����ID, mstr�Ա�, True)
            End If
            If Not rsTmp Is Nothing Then
                .Tag = ""
                Call Set��ҽ֤��(Row, 0, rsTmp)
                Call ZYEnterNextCell
            End If
        ElseIf Col = colZY���� Then
            lngRow = Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, colzy����) = .TextMatrix(Row, colzy����)
            .Cell(flexcpBackColor, lngRow, col��ϱ���) = ColorUnEditCell      '����ɫ
            .Row = lngRow: .Col = Col�������
            .ShowCell .Row, .Col
        ElseIf Col = colZYDel Then
            Call vsDiagZY_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiagZY_Click()
    With vsDiagZY
        If (.MouseCol = colZY���� Or .MouseCol = colZYDel) And .MouseRow >= .FixedRows Then
            If .MouseCol = colZY���� Then
                If .TextMatrix(.MouseRow, Col�������) = "" Or .TextMatrix(.MouseRow, 0) = "��Ҫ���" Then Exit Sub
            End If

            .Select .MouseRow, .MouseCol
            Call vsDiagZY_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsDiagZY
        If Col = col��Ժ��� Then
            '��λ��ƥ����
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsDiagZY_DblClick()
    Call vsDiagZY_KeyPress(32)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long

    With vsDiagZY
        If KeyCode = vbKeyF4 Then
            If .Col = Col������� Then
                Call gobjComLib.zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, Col�������) <> "" Then
                If .RowData(.Row) & "" <> "" Then Exit Sub
                If GetAdviceIDByDiag(Val(.Cell(flexcpData, .Row, col�Ƿ�����) & "")) <> "" Then Exit Sub
                If mlngPathState = 1 Or mlngPathState = 2 Then
                    If .TextMatrix(.Row, col�������) = "��Ժ���" And mlngDiagnosisType = 12 Or .TextMatrix(.Row, col�������) = "�������" And mlngDiagnosisType = 11 Then
                        If .TextMatrix(.Row, col�������) <> .TextMatrix(.Row - 1, col�������) Then
                            '��Ҫ��ϲ������
                            Exit Sub
                        End If
                    End If
                End If
                '�ϲ�·��
                If Not CheckMergePath(mlng����ID, mlng����ID, Val(.TextMatrix(.Row, colzy����)), Val(.TextMatrix(.Row, colzy����ID))) Then Exit Sub
                '����·������
                If mstrPathDiag <> "" And mlngPathState > 0 Then
                    If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, colzy����) & "|" & Val(.TextMatrix(.Row, col����ID)) & "|" & Val(.TextMatrix(.Row, col���ID)) & ",") > 0 Then
                        '������ϲ������
                        Exit Sub
                    End If
                End If
                If mlngPathState = 2 And mblnIsPathOutTime Then
                    If .TextMatrix(.Row, col�������) = "��Ҫ���" And mlngDiagnosisType > 10 Then
                        '������ɵĳ�Ժ��ϲ������
                        Exit Sub
                    End If
                End If
                If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    i = Val(.TextMatrix(.Row, colzy����))
                    .Cell(flexcpText, .Row, .FixedRows, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedRows, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, colzy����) = i

                    '�����ͬ�������������
                    If .TextMatrix(.Row, col�������) = "" Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, col�������) = "" Then
                                '��һ��Ϊ�ޱ����������ʱ�����ݲ����ƣ�����ǰ��Ϊ�б���ʱֻ�����
                                For i = .Row + 1 To .Rows - 1
                                    If Val(.TextMatrix(i, colzy����)) = Val(.TextMatrix(.Row, colzy����)) Then
                                        For j = .FixedCols To .Cols - 1
                                            .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                            .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                        Next
                                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                        .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                        .TextMatrix(i, colzy����) = Val(.TextMatrix(.Row, colzy����))
                                        .RowData(i - 1) = .RowData(i)
                                        .RowData(i) = Empty
                                        
                                        If i = .Rows - 1 Then
                                            If .TextMatrix(i, col�������) = "" Then .RemoveItem i
                                            Exit For
                                        ElseIf Val(.TextMatrix(i + 1, colzy����)) <> Val(.TextMatrix(i, colzy����)) Then
                                            If .TextMatrix(i, col�������) = "" Then .RemoveItem i
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    .Tag = ""
                End If
            ElseIf .TextMatrix(.Row, col�������) = "" Then
                .RemoveItem .Row
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiagZY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    With vsDiagZY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call ZYEnterNextCell
        ElseIf KeyAscii = 32 And (.Col = colZY����) Then
            If ZYCellEditable(.Row, .Col) Then
                KeyAscii = 0
                If .Col = colZY���� Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "��", "")
                End If
            End If
        Else
            If .Col = Col������� Or .Col = col��ҽ֤�� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagZY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True

        With vsDiagZY
            If Col = col��Ժ��� Then
                KeyAscii = 0

                '��ʱ.TextMatrix��δ����,����ȡComboItem
                .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                .Tag = ""
                Call ZYEnterNextCell
            End If
        End With
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagZY.EditSelStart = 0
    vsDiagZY.EditSelLength = gobjComLib.zlCommFun.ActualLen(vsDiagZY.EditText)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ZYCellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = colZY���� Then
        Cancel = True '��ֱ�ӱ༭
    End If
End Sub

Private Function GetSQL(ByVal intType As Integer, ByVal strInput As String, ByRef str�Ա� As String, Optional ByVal strOtherInfo As String) As String
'���ܣ���ò�ѯ��ҽ��ϵ�SQL
'������intType:��ȡ��SQL����,0-��ҽ��ϣ�1-��ҽ��ϣ�2-��������
'    strInput-��ѯ������str�Ա�--���˵��Ա�
'   strOtherInfo:��ҽ���-������������
'���أ�strsql--��ѯ��ϵ�SQL
    Dim strSQL As String

    If mstr�Ա� Like "*��*" Then
        str�Ա� = "��"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "Ů"
    End If

    Select Case intType
        Case 0 '��ҽ���
            If optInput(0).value Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If
                strSQL = _
                    " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                    " From �������Ŀ¼ A,������ϱ��� B" & _
                    " Where A.ID=B.���ID And A.���=1" & _
                    " And B.����=[5] And (" & strSQL & ")" & _
                    " Order by A.����"
            Else
                If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(mint���� = 0, "����", "�����") & " Like [2]"
                End If
                strSQL = _
                    " Select ID,ID as ��ĿID,����,����,����," & IIf(mint���� = 0, "����", "����� as ����") & ",˵��" & _
                    " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by ����"
            End If

        Case 1 '��ҽ���
            If optInput(0).value And strOtherInfo <> "Z" Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If
                strSQL = _
                    " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                    " From �������Ŀ¼ A,������ϱ��� B" & _
                    " Where A.ID=B.���ID And A.���=2" & _
                    " And B.����=[4] And (" & strSQL & ")" & _
                    " Order by A.����"
            Else
                'B-��ҽ��������
                If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(mint���� = 0, "����", "�����") & " Like [2]"
                End If
                strSQL = _
                    " Select ID,ID as ��ĿID,����,����,����," & IIf(mint���� = 0, "����", "����� as ����") & ",˵��" & _
                    " From ��������Ŀ¼" & _
                    " Where ���='" & IIf(strOtherInfo = "", "B", strOtherInfo) & "' And (" & strSQL & ")" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by ����"
            End If
    End Select
    GetSQL = strSQL
End Function

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim str�Ա� As String, int������� As Integer

    With vsDiagZY
        If Col = Col������� Or Col = col��ҽ֤�� Then
            '.Cell(flexcpData, Row, Col) <> ""�ų����лس�
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                .EditText = ""
                '��ҽ֢���������������
                If Col = col��ҽ֤�� Then
                    .Cell(flexcpData, Row, Col) = ""
                End If
                .Tag = ""
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ZYEnterNextCell
            ElseIf Col = Col������� And .TextMatrix(Row, col��ϱ���) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                strInput = UCase(.EditText)
                strSQL = GetSQL(1, strInput, str�Ա�)
                On Error GoTo errH
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, str�Ա�, mint���� + 1)
                If rsTmp.RecordCount = 1 Then
                    Call ZYSetDiagInput(Row, rsTmp):
                    .EditText = .Text
                Else
                    '�����ڱ�׼������ǰ�����븽����Ϣ
                    .TextMatrix(Row, Col�������) = .EditText
                End If
                .Tag = ""
                '������.Cell(flexcpData, Row, Col)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
            Else
                If Val(.TextMatrix(Row, colzy����)) = 11 Then
                    int������� = Val(Mid(mstr�������, 1, 1))
                Else
                    int������� = Val(Mid(mstr�������, 2, 1))
                End If
                If int������� = 0 Then int������� = 1

                strInput = UCase(.EditText)
                strSQL = GetSQL(1, strInput, str�Ա�, IIf(Col = Col�������, "B", "Z"))
                If Col = Col������� Then
                    If int������� = 1 And gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1)
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, str�Ա�, mint���� + 1)
                            If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                        End If
                        .Tag = ""
                        Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                        If mblnReturn And rsTmp Is Nothing Then Call ZYEnterNextCell '��������¼��ʱ���ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                    Else
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInput(0).value, "�������", "��������"), False, "", "", False, False, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1)
                        If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                            Cancel = True
                        Else
                            '���������뷽ʽ
                            If rsTmp Is Nothing And ((int������� = 2 Or int������� = 3 And mint���� <> 0)) Then
                                MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                                Cancel = True
                            Else
                                .Tag = ""
                                Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                                'If mblnReturn Then Call ZYEnterNextCell '�ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                            End If
                        End If
                    End If
                ElseIf Col = col��ҽ֤�� Then
                    If optInput(0).value Then
                        '���������:�Ȳ��Ƿ��ж�Ӧ
                        If Set��ҽ֤��(Row, Val(.TextMatrix(Row, colzy���ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .Tag = ""
                            Call Set��ҽ֤��(Row, 0, rsTmp)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col����ʱ�� Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    .Tag = ""
                Else
                    MsgBox "��������ȷ�ķ���ʱ�䣬���磺""2012-12-21 00:00""��"
                    Cancel = True
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub LoadData()
    Dim bln��ҳ��� As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long, lngRow As Long, j As Long
    Dim str���ƽ�� As String
    Dim str���Id As String
    
    On Error GoTo errH
    '��ҽ���
    '--------------------------------------------------------------
    '�ж���ҳ�Ƿ�������
    strSQL = "Select 1 From ������ϼ�¼ Where ����ID=[1] And ��ҳID=[2] And ��¼��Դ=3  And RowNum<2"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
    bln��ҳ��� = rsTmp.RecordCount > 0
    If Not bln��ҳ��� And mint������Դ = 2 Then
        strTmp = " And a.��¼��Դ IN(1,2,3,4) "
    Else
        strTmp = " and a.��¼��Դ=3 "
    End If

    'ȱʡ����ʼ��
    With vsDiagXY
        .ColData(col��Ժ���) = Get���ƽ��
        '1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
        .TextMatrix(1, col����) = 1
        If mint������Դ = 2 Then
            .TextMatrix(2, col����) = 2
            .TextMatrix(3, col����) = 3
            .TextMatrix(4, col����) = 3
            .TextMatrix(5, col����) = 5
            .TextMatrix(6, col����) = 10
            .TextMatrix(7, col����) = 6
            .TextMatrix(8, col����) = 7
        Else
            .Rows = .FixedRows + 1
        End If
    End With

    '��ȡ������Դ�����
    strSQL = "Select a.��ע,a.ID,a.����ID,a.��ҳID,a.ҽ��ID,a.��¼��Դ,a.��ϴ���,a.�������,a.����ID,a.�������,a.����ID,a.��Ժ����," & _
        " a.���ID,a.֤��ID,a.�������,a.��Ժ���,a.�Ƿ�δ��,a.�Ƿ�����,a.��¼����,a.��¼��,a.ȡ��ʱ��,a.ȡ����,a.����ID, b.���� As ��������, c.���� As ��ϱ���,A.����ʱ�� " & _
        " From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & _
        " Where a.����id = b.Id(+) And a.���id = c.Id(+)  And a.������� IN(1,2,3,5,6,7,10,21)" & _
        strTmp & _
        " And a.ȡ��ʱ�� is Null And a.����ID=[1] And a.��ҳID=[2]" & _
        " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.ID"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            If mint������Դ = 2 Then
                strSQL = "1,2,3,5,6,7,10"
            Else
                 strSQL = "1"
            End If
            For i = 0 To UBound(Split(strSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(strSQL, ",")(i)
                If mint������Դ = 2 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(strSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(strSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(strSQL, ",")(i)
                    End If
                End If
                Do While Not rsTmp.EOF
                    'ȷ����ǰ��ʾ��
                    lngRow = .FindRow(CStr(Split(strSQL, ",")(i)), , col����)
                    For j = lngRow To .Rows - 1
                        If Val(.TextMatrix(j, col����)) = Val(Split(strSQL, ",")(i)) Then
                            lngRow = j
                            If .TextMatrix(j, Col�������) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    
                    If .TextMatrix(lngRow, Col�������) <> "" Then
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, col����) = Split(strSQL, ",")(i)
                    End If
                    
                    If InStr("," & mstr���IDs & ",", "," & rsTmp!id & ",") > 0 Then
                        .TextMatrix(lngRow, col����) = 1
                    End If
                    
                    str���Id = str���Id & "," & rsTmp!id
                    
                    If IsNull(rsTmp!�������) Then
                        .TextMatrix(lngRow, col��ϱ���) = ""
                        .TextMatrix(lngRow, Col�������) = ""
                    Else
                        If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���ID & "") = 0 And Val(rsTmp!����ID & "") = 0) Then '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                            '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                            If Val(rsTmp!����ID & "") <> 0 Then
                                .TextMatrix(lngRow, col��ϱ���) = NVL(rsTmp!��������)
                            ElseIf Val(rsTmp!���ID & "") <> 0 Then
                                .TextMatrix(lngRow, col��ϱ���) = NVL(rsTmp!��ϱ���)
                            Else
                                .TextMatrix(lngRow, col��ϱ���) = ""
                            End If
                            .TextMatrix(lngRow, Col�������) = rsTmp!�������
                        Else
                            .TextMatrix(lngRow, col��ϱ���) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                            .TextMatrix(lngRow, Col�������) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                        End If
                    End If
                    If Not IsNull(rsTmp!����ID) Or Not IsNull(rsTmp!���ID) Then
                        .Cell(flexcpData, lngRow, Col�������) = Get�������(Val("" & rsTmp!���ID), Val("" & rsTmp!����ID))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                    Else
                        .Cell(flexcpData, lngRow, Col�������) = .TextMatrix(lngRow, Col�������)
                    End If
                    If mint������Դ = 1 Then
                        .TextMatrix(lngRow, col����ʱ��) = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                    Else
                        .TextMatrix(lngRow, col��Ժ���) = NVL(rsTmp!��Ժ���)
                        .TextMatrix(lngRow, col��Ժ����) = NVL(rsTmp!��Ժ����)
                        .TextMatrix(lngRow, col�Ƿ�δ��) = IIf(NVL(rsTmp!�Ƿ�δ��, 0) = 1, "��", "")
                    End If
                    
                    .TextMatrix(lngRow, col��ע) = NVL(rsTmp!��ע)
                    .Cell(flexcpData, lngRow, col�Ƿ�����) = Val(rsTmp!id & "")
                    .TextMatrix(lngRow, col�Ƿ�����) = IIf(NVL(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                    .TextMatrix(lngRow, col���ID) = NVL(rsTmp!���ID, 0)
                    .TextMatrix(lngRow, col����ID) = NVL(rsTmp!����ID, 0)
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If

    vsDiagXY.Cell(flexcpForeColor, 1, col�Ƿ�����, vsDiagXY.Rows - 1, col�Ƿ�����) = vbRed
    lngRow = GetRow(3)
    If lngRow <> -1 Then
        vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    End If
    vsDiagXY.Cell(flexcpBackColor, 1, col��ϱ���, vsDiagXY.Rows - 1, col��ϱ���) = ColorUnEditCell      '����ɫ
    vsDiagXY.Row = 1: vsDiagXY.Col = Col�������
    Call vsDiagXY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
    vsDiagXY.Tag = "δ�޸�"
    '��ҽ���
    '---------------------------------------------------------------
    If mbln��ҽ Then
        'ȱʡ����ʼ��
        With vsDiagZY
            '11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���(��Ҫ��ϡ��������)
            .ColData(col��Ժ���) = str���ƽ��
            .TextMatrix(1, colzy����) = 11
            If mint������Դ = 2 Then
                .TextMatrix(2, colzy����) = 12
                .TextMatrix(3, colzy����) = 13
                .TextMatrix(4, colzy����) = 13
            Else
                .Rows = .FixedRows + 1
            End If
        End With
        If Not bln��ҳ��� And mint������Դ = 2 Then
            strTmp = " And a.��¼��Դ IN(1,2,3,4) "
        Else
            strTmp = " and a.��¼��Դ=3 "
        End If
    
        '��ȡ������Դ�����
        strSQL = "Select a.��ע, a.Id, a.����id, a.��ҳid, a.ҽ��id, a.��¼��Դ, a.��ϴ���, a.�������, a.����id, a.�������,a.��Ժ����," & _
            " a.����id, a.���id, a.֤��id, a.�������,a.��Ժ���, a.�Ƿ�δ��, a.�Ƿ�����, a.��¼����, a.��¼��, a.ȡ��ʱ��," & _
            " a.ȡ����, a.����id, b.���� As ��������, c.���� As ��ϱ���,d.���� as ֤����� ,A.����ʱ�� From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C,��������Ŀ¼ D" & _
            " Where a.����id = b.Id(+) And a.���id = c.Id(+) And a.֤��ID=d.ID(+) And a.������� IN(11,12,13)" & _
            strTmp & _
            " And ȡ��ʱ�� Is Null And ����ID=[1] And ��ҳID=[2]" & _
            " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.�������,a.ID"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
        strTmp = ""
        If Not rsTmp.EOF Then
            With vsDiagZY
                If mint������Դ = 2 Then
                    strSQL = "11,12,13"
                Else
                     strSQL = "11"
                End If
                
                For i = 0 To UBound(Split(strSQL, ","))
                    
                    rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(strSQL, ",")(i)
                    If mint������Դ = 2 Then
                        If rsTmp.EOF Then
                            rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(strSQL, ",")(i)
                        End If
                        If rsTmp.EOF Then
                            rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(strSQL, ",")(i)
                        End If
                        If rsTmp.EOF Then
                            rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(strSQL, ",")(i)
                        End If
                    End If
                    Do While Not rsTmp.EOF
                        'ȷ����ǰ��ʾ��
                        lngRow = .FindRow(CStr(Split(strSQL, ",")(i)), , colzy����)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, colzy����)) = Val(Split(strSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, Col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        
                        If .TextMatrix(lngRow, Col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, colzy����) = Split(strSQL, ",")(i)
                        End If
                        
                        If InStr("," & mstr���IDs & ",", "," & rsTmp!id & ",") > 0 Then
                            .TextMatrix(lngRow, col����) = 1
                        End If
                        
                        str���Id = str���Id & "," & rsTmp!id
                        
                        If IsNull(rsTmp!�������) Then
                            .TextMatrix(lngRow, col��ϱ���) = ""
                            .TextMatrix(lngRow, Col�������) = ""
                        Else
                            If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���ID & "") = 0 And Val(rsTmp!����ID & "") = 0) Then     '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                                '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                                If Val(rsTmp!����ID & "") <> 0 Then
                                    .TextMatrix(lngRow, col��ϱ���) = NVL(rsTmp!��������)
                                ElseIf Val(rsTmp!���ID & "") <> 0 Then
                                    .TextMatrix(lngRow, col��ϱ���) = NVL(rsTmp!��ϱ���)
                                Else
                                    .TextMatrix(lngRow, col��ϱ���) = ""
                                End If
                                .TextMatrix(lngRow, Col�������) = rsTmp!�������
                            Else
                                .TextMatrix(lngRow, col��ϱ���) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                                .TextMatrix(lngRow, Col�������) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                            End If
                        End If

                        .TextMatrix(lngRow, col��ע) = NVL(rsTmp!��ע)
                       .Cell(flexcpData, lngRow, colZY����) = Val(rsTmp!id & "")
                       .Cell(flexcpData, lngRow, col��ϱ���) = .TextMatrix(lngRow, col��ϱ���)
                        .TextMatrix(lngRow, colzy���ID) = NVL(rsTmp!���ID, 0)
                        .TextMatrix(lngRow, colzy����ID) = NVL(rsTmp!����ID, 0)
                        .TextMatrix(lngRow, colzy֤��ID) = NVL(rsTmp!֤��id, 0)
                        If mint������Դ = 1 Then
                            .TextMatrix(lngRow, colZY����) = IIf(NVL(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                            .TextMatrix(lngRow, col����ʱ��) = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                        Else
                            .TextMatrix(lngRow, col��Ժ���) = NVL(rsTmp!��Ժ���)
                            .TextMatrix(lngRow, col��Ժ����) = NVL(rsTmp!��Ժ����)
                        End If
                        'ȡ֤������
                        If InStr(.TextMatrix(lngRow, Col�������), "(") > 0 And InStr(.TextMatrix(lngRow, Col�������), ")") > 0 Then
                            strTmp = Mid(.TextMatrix(lngRow, Col�������), InStrRev(.TextMatrix(lngRow, Col�������), "(") + 1)
                            strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                            '��ȡ֤��
                            .TextMatrix(lngRow, col��ҽ֤��) = strTmp
                            'ȥ�����������֤��
                            .TextMatrix(lngRow, Col�������) = Mid(.TextMatrix(lngRow, Col�������), 1, InStrRev(.TextMatrix(lngRow, Col�������), "(") - 1)
                        Else
                           .TextMatrix(lngRow, col��ҽ֤��) = ""
                        End If
                        '����¼����ϵ������������Ҫȥ��֤����˴˾�������
                        If Not IsNull(rsTmp!����ID) Or Not IsNull(rsTmp!���ID) Then
                            .Cell(flexcpData, lngRow, Col�������) = Get�������(Val("" & rsTmp!���ID), Val("" & rsTmp!����ID))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                        Else
                            .Cell(flexcpData, lngRow, Col�������) = .TextMatrix(lngRow, Col�������)
                        End If
                        rsTmp.MoveNext
                    Loop
                Next
            End With
        End If
        vsDiagZY.Cell(flexcpForeColor, vsDiagZY.FixedRows, colZY����, vsDiagZY.Rows - 1, colZY����) = vbRed
        lngRow = GetRow(13)
        If lngRow <> -1 Then
            vsDiagZY.Cell(flexcpBackColor, GetRow(13), vsDiagZY.FixedRows, GetRow(13), vsDiagZY.Cols - 1) = &HC0FFC0
        End If
        vsDiagZY.Cell(flexcpBackColor, 1, col��ϱ���, vsDiagZY.Rows - 1, col��ϱ���) = ColorUnEditCell      '����ɫ
        vsDiagZY.Row = 1: vsDiagZY.Col = Col�������
        Call vsDiagZY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
        vsDiagZY.Tag = "δ�޸�"
    End If
    '�������ҽ����ϵ
    If str���Id <> "" Then
        str���Id = Mid(str���Id, 2)
        
        strSQL = "Select /*+ RULE*/" & vbNewLine & _
                " F_List2str(Cast(Collect(A.ҽ��id || '') As T_Strlist)) As ҽ��ids, A.���id" & vbNewLine & _
                "From �������ҽ�� A, ����ҽ����¼ B" & vbNewLine & _
                "Where A.���id In (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) And A.ҽ��id = B.Id And" & vbNewLine & _
                "      B.ҽ��״̬ <> -1 And B.ҽ��״̬ <> 4" & vbNewLine & _
                "Group By A.���id"
                
        Set mrsAdvice = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���Id)
        
        With vsDiagXY
            For i = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, i, col�Ƿ�����) & "") > 0 Then
                    .RowData(i) = GetAdviceIDByDiag(Val(.Cell(flexcpData, i, col�Ƿ�����) & ""))
                End If
            Next
        End With
        
        If mbln��ҽ Then
            With vsDiagZY
                For i = .FixedRows To .Rows - 1
                    If Val(.Cell(flexcpData, i, colZY����) & "") > 0 Then
                        .RowData(i) = GetAdviceIDByDiag(Val(.Cell(flexcpData, i, colZY����) & ""))
                    End If
                Next
            End With
        End If
    End If
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub SaveData()
   Dim arrSQL As Variant
   Dim intIdx As Integer
   Dim i As Long
   Dim str�������  As String
   Dim datCurDate As Date
   Dim blnTrans As Boolean
   Dim lngID As Long
   Dim blnChange As Boolean
   Dim str����ҽ��ID As String
   
    mstr���IDs = ""
    mstr���s = ""
    arrSQL = Array()
    datCurDate = gobjComLib.zlDatabase.Currentdate
    blnChange = vsDiagXY.Tag = ""
    '��ҽ���
    If blnChange Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_DELETE(" & mlng����ID & "," & mlng����ID & ",3,NULL,'1,2,3,5,6,7,10')"
    End If
    With vsDiagXY
        intIdx = 0
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, Col�������)) <> "" Then
                If Trim(.TextMatrix(i, col��ϱ���)) = "" Then
                    str������� = .TextMatrix(i, Col�������)
                Else
                    str������� = "(" & .TextMatrix(i, col��ϱ���) & ")" & .TextMatrix(i, Col�������)
                End If
                lngID = Val(.Cell(flexcpData, i, col�Ƿ�����))
                str����ҽ��ID = ""
                If Not mrsAdvice Is Nothing Then
                    mrsAdvice.Filter = "���ID=" & lngID
                    If Not mrsAdvice.EOF Then
                        mrsAdvice.MoveFirst
                        str����ҽ��ID = mrsAdvice!ҽ��IDs
                    End If
                End If
                
                If Val(.TextMatrix(i, col����)) <> 0 Then
                    If lngID = 0 Then lngID = gobjComLib.zlDatabase.GetNextId("������ϼ�¼")
                    mstr���IDs = mstr���IDs & "," & lngID
                    mstr���s = mstr���s & "," & str�������
                End If
                If blnChange Then
                    If Val(.TextMatrix(i, col����)) <> Val(.TextMatrix(i - 1, col����)) Then intIdx = 0
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                    arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng����ID & ",3,NULL," & _
                        Val(.TextMatrix(i, col����)) & "," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & "," & _
                        "NULL,'" & str������� & "','" & NeedName(.TextMatrix(i, col��Ժ���)) & "'," & _
                        IIf(.TextMatrix(i, col�Ƿ�δ��) = "", 0, 1) & "," & IIf(.TextMatrix(i, col�Ƿ�����) = "", 0, 1) & "," & _
                        "To_Date('" & Format(datCurDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        IIf(str����ҽ��ID = "", "Null,", "'" & str����ҽ��ID & "',") & intIdx & ",'" & .TextMatrix(i, col��ע) & "','" & .TextMatrix(i, col��Ժ����) & "',Null,'" & mstr������ & "'," & IIf(lngID = 0, "Null", lngID) & ")"
                End If
            End If
        Next
    End With
 
    '��ҽ���
    If vsDiagZY.Visible Then
        blnChange = vsDiagZY.Tag = ""
        If blnChange Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_DELETE(" & mlng����ID & "," & mlng����ID & ",3,NULL,'11,12,13')"
        End If
        With vsDiagZY
            intIdx = 0
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, Col�������)) <> "" Then
                    If Trim(.TextMatrix(i, col��ϱ���)) = "" Then
                        str������� = .TextMatrix(i, Col�������) & IIf(.TextMatrix(i, col��ҽ֤��) <> "", "(" & .TextMatrix(i, col��ҽ֤��) & ")", "")
                    Else
                        str������� = "(" & .TextMatrix(i, col��ϱ���) & ")" & .TextMatrix(i, Col�������) & IIf(.TextMatrix(i, col��ҽ֤��) <> "", "(" & .TextMatrix(i, col��ҽ֤��) & ")", "")
                    End If
                    lngID = Val(.Cell(flexcpData, i, colZY����))
                    str����ҽ��ID = ""
                    If Not mrsAdvice Is Nothing Then
                        mrsAdvice.Filter = "���ID=" & lngID
                        If Not mrsAdvice.EOF Then
                            mrsAdvice.MoveFirst
                            str����ҽ��ID = mrsAdvice!ҽ��IDs
                        End If
                    End If
                    If Val(.TextMatrix(i, col����)) <> 0 Then
                        If lngID = 0 Then lngID = gobjComLib.zlDatabase.GetNextId("������ϼ�¼")
                        mstr���IDs = mstr���IDs & "," & lngID
                        mstr���s = mstr���s & "," & str�������
                    End If
                    If blnChange Then
                        If Val(.TextMatrix(i, colzy����)) <> Val(.TextMatrix(i - 1, colzy����)) Then intIdx = 0
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng����ID & ",3,NULL," & _
                            Val(.TextMatrix(i, colzy����)) & "," & ZVal(.TextMatrix(i, colzy����ID)) & "," & ZVal(.TextMatrix(i, colzy���ID)) & "," & _
                            ZVal(.TextMatrix(i, colzy֤��ID)) & ",'" & str������� & "','" & NeedName(.TextMatrix(i, col��Ժ���)) & "'," & _
                            "NULL,NULL,To_Date('" & Format(datCurDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            IIf(str����ҽ��ID = "", "Null,", "'" & str����ҽ��ID & "',") & intIdx & ",'" & .TextMatrix(i, col��ע) & "','" & .TextMatrix(i, col��Ժ����) & "',Null,'" & mstr������ & "'," & IIf(lngID = 0, "Null", lngID) & ")"
                    End If
                End If
            Next
        End With
    End If
    
    If mstr���IDs <> "" Then mstr���IDs = Mid(mstr���IDs, 2)
    If mstr���s <> "" Then mstr���s = Mid(mstr���s, 2)
    
    If vsDiagXY.Tag = "" Or vsDiagZY.Tag = "" And vsDiagZY.Visible Then
        Screen.MousePointer = 11
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call gobjComLib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        
        On Error GoTo 0
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Function Get���ƽ��() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select ����,����,���� From ���ƽ�� Order by ����"
    Call gobjComLib.zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)

    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "|" & rsTmp!���� & "-" & rsTmp!����
        rsTmp.MoveNext
    Loop
    If strSQL = "" Then
        Get���ƽ�� = "1-����|2-��ת|3-δ��|4-����|5-����"
    Else
        Get���ƽ�� = Mid(strSQL, 2)
    End If
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function NeedName(strList As String) As String
'˵��:1-strList��()��[]�ָ����������ʱ��������[����]��(����)��ͷ,�������Ϊ���ֻ���ĸ
'     2-�ָ��������ȼ����س���(Chr(13)��> - > [] > ()

    '�����ж��Իس����ָ�
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
        Exit Function
    End If
    '��[]�ָ�
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
        If gobjComLib.zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
            Exit Function
        End If
    End If
    '��()�ָ�
    If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
        If gobjComLib.zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
            Exit Function
        End If
    End If
    '��-�ָ�
    NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    
End Function

Private Function XYCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsDiagXY
        '�����в��ɱ༭
        If .ColHidden(lngCol) Then Exit Function
        
        If lngCol = col���� Then
            If Trim(.TextMatrix(lngRow, Col�������)) = "" Then
                Exit Function
            End If
        Else
            If .RowData(lngRow) & "" <> "" Then Exit Function
        End If
        
        If lngCol = Col������� And (mlngPathState = 1 Or mlngPathState = 2) Then
            If .TextMatrix(lngRow, col�������) = "��Ժ���" And mlngDiagnosisType = 2 Or .TextMatrix(lngRow, col�������) = "�������" And mlngDiagnosisType = 1 Then
                If .TextMatrix(lngRow, Col�������) <> "" And .TextMatrix(lngRow, col�������) <> .TextMatrix(lngRow - 1, col�������) Then
                    '��Ҫ��ϲ������
                    Exit Function
                End If
            End If
            '�ϲ�·��
            If Not CheckMergePath(mlng����ID, mlng����ID, Val(.TextMatrix(lngRow, col����)), Val(.TextMatrix(lngRow, col����ID))) Then Exit Function
        End If
        If lngCol = Col������� Then
            '����·������
            If mstrPathDiag <> "" And mlngPathState > 0 Then
                If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, col����) & "|" & Val(.TextMatrix(.Row, col����ID)) & "|" & Val(.TextMatrix(.Row, col���ID)) & ",") > 0 Then
                    '������ϲ������
                    Exit Function
                End If
            End If
        End If
        If lngCol = Col������� And mlngPathState = 2 And mblnIsPathOutTime Then
            If .TextMatrix(.Row, col�������) = "��Ժ���" And mlngDiagnosisType <= 2 Then
                '������ɵĳ�Ժ��ϲ������
                Exit Function
            End If
        End If
        '�������������
        If .TextMatrix(lngRow, Col�������) = "" Then
            If lngCol = col��Ժ��� Or lngCol = col��ע Or lngCol = col�Ƿ�δ�� Or lngCol = col�Ƿ����� Or lngCol = col���� Or lngCol = col����ʱ�� Then
                Exit Function
            End If
        End If
        If lngCol = col��ϱ��� Then Exit Function
        
        If lngCol = col���� Then
            If Val(.TextMatrix(lngRow, col����)) = 3 Then
                If .TextMatrix(lngRow, col�������) = "��Ժ���" Then Exit Function
            End If
        End If
        
        '��Ժ��Ϻ�Ժ�ڸ�Ⱦ���������Ժ���(��Ϊ����Ժ�ڸ�Ⱦ�ڳ�Ժʱ�Ѿ���ת��������)
        If Val(.TextMatrix(lngRow, col����)) = 3 Or Val(.TextMatrix(lngRow, col����)) = 5 Or Val(.TextMatrix(lngRow, col����)) = 10 Then
            '��Ժ��ϱ�����������(��δ����ʱ)
            If .TextMatrix(lngRow, Col�������) = "" And Val(.TextMatrix(lngRow, col����)) = 3 Then
                If Val(.TextMatrix(lngRow - 1, col����)) = 3 And .TextMatrix(lngRow - 1, Col�������) = "" Then
                    Exit Function
                End If
            End If

            '��Ժ���Ϊ"����"ʱ�ſ��������Ƿ�δ��
            If .TextMatrix(lngRow, col��Ժ���) <> "����" And lngCol = col�Ƿ�δ�� Then
                Exit Function
            End If
        ElseIf lngCol = col��Ժ��� Or lngCol = col�Ƿ�δ�� Then
            Exit Function
        End If
        
        '��Ժ����ֻ���ڳ�Ժ��Ϻ������������д
        If lngCol = col��Ժ���� Then
            If .TextMatrix(lngRow, col����) <> "3" Then
                Exit Function
            End If
        End If
    End With
    XYCellEditable = True
End Function

Private Function CheckMergePath(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngDiagType As Long, ByVal lngDiag As Long) As Boolean
'���ܣ����ϲ�·����Ӧ����ϲ����޸�
'������lngDiagType���������,lngDiag=����ID
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    If lngDiag = 0 Or lngDiagType = 0 Then CheckMergePath = True: Exit Function
    strSQL = "Select �������,����ID From ���˺ϲ�·�� Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng����ID, lng��ҳID)
    Do While Not rsTmp.EOF
        If lngDiagType = Val(rsTmp!������� & "") And lngDiag = Val(rsTmp!����ID & "") Then
            Exit Function
        End If
        rsTmp.MoveNext
    Loop
    CheckMergePath = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Sub XYEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiagXY
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, Col�������) To col����
                If XYCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= col���� Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Private Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Private Function ZYCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsDiagZY
        '�����в��ɱ༭
        If .ColHidden(lngCol) Then Exit Function

        If lngCol = col���� Then
            If Trim(.TextMatrix(lngRow, Col�������)) = "" Then
                Exit Function
            End If
        Else
            If .RowData(lngRow) & "" <> "" Then Exit Function
        End If
        
        If lngCol = Col������� And (mlngPathState = 1 Or mlngPathState = 2) Then
            If .TextMatrix(lngRow, col�������) = "��Ժ���" And mlngDiagnosisType = 12 Or .TextMatrix(lngRow, col�������) = "�������" And mlngDiagnosisType = 11 Then
                If .TextMatrix(lngRow, Col�������) <> "" And .TextMatrix(lngRow, col�������) <> .TextMatrix(lngRow - 1, col�������) Then
                    '��Ҫ��ϲ������
                    Exit Function
                End If
            End If
            '�ϲ�·��
            If Not CheckMergePath(mlng����ID, mlng����ID, Val(.TextMatrix(lngRow, colzy����)), Val(.TextMatrix(lngRow, colzy����ID))) Then Exit Function
        End If
        
        If lngCol = Col������� Then
            '����·������
            If mstrPathDiag <> "" And mlngPathState > 0 Then
                If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, colzy����) & "|" & Val(.TextMatrix(.Row, col����ID)) & "|" & Val(.TextMatrix(.Row, col���ID)) & ",") > 0 Then
                    '������ϲ������
                    Exit Function
                End If
            End If
        End If
        If lngCol = Col������� And mlngPathState = 2 And mblnIsPathOutTime Then
            If .TextMatrix(.Row, col�������) = "��Ҫ���" And mlngDiagnosisType > 10 Then
                '������ɵĳ�Ժ��ϲ������
                Exit Function
            End If
        End If
        '�������������
        If .TextMatrix(lngRow, Col�������) = "" Then
            If lngCol = col��Ժ��� Or lngCol = col��ע Or lngCol = colZY���� Or lngCol = col����ʱ�� Or lngCol = colZY���� Then Exit Function
        End If
        If lngCol = col��ϱ��� Then Exit Function
        
        If lngCol = colZY���� Then
            If Val(.TextMatrix(lngRow, colzy����)) = 13 Then
                If .TextMatrix(lngRow, col�������) = "��Ҫ���" Then Exit Function
            End If
        End If
        
        If Val(.TextMatrix(lngRow, colzy����)) = 13 Then
            '��Ժ��ϱ�����������(��δ����ʱ)
            If .TextMatrix(lngRow, Col�������) = "" Then
                If Val(.TextMatrix(lngRow - 1, colzy����)) = 13 And .TextMatrix(lngRow - 1, Col�������) = "" Then
                    Exit Function
                End If
            End If
        ElseIf lngCol = col��Ժ��� Then
            '�ǳ�Ժ���ʱ����������
            If Val(.TextMatrix(lngRow, colzy����)) <> 13 Then Exit Function
        End If
        '��Ժ����ֻ������Ҫ��Ϻ������������д
        If lngCol = col��Ժ���� Then
            If .TextMatrix(lngRow, colzy����) <> "13" Then
                Exit Function
            End If
        End If
        '���������������֤��
        If lngCol = col��ҽ֤�� Then
            If .TextMatrix(lngRow, Col�������) = "" Then Exit Function
        End If
    End With
    ZYCellEditable = True
End Function

Private Sub ZYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long, j As Long
    Dim strTmp As String
    
    With vsDiagZY
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    '�������ѡ�����ʱ�Ĵ���
                    If lngRow = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, colzy����) = .TextMatrix(lngRow, colzy����)
                    End If
                    'ȷ����ǰ��ʾ��
                    If Val(.TextMatrix(lngRow + 1, colzy����)) = Val(.TextMatrix(lngRow, colzy����)) Then
                        For j = lngRow + 1 To .Rows - 1
                            If Val(.TextMatrix(j, colzy����)) = Val(.TextMatrix(lngRow, colzy����)) Then
                                lngRow = j
                                If .TextMatrix(j, Col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, Col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, colzy����) = .TextMatrix(lngRow - 1, colzy����)
                        End If
                    Else
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy����) = .TextMatrix(lngRow - 1, colzy����)
                    End If
                End If
                
                If InStr(.TextMatrix(lngRow, Col�������), "(") > 0 And InStr(.TextMatrix(lngRow, Col�������), ")") > 0 Then
                    strTmp = Mid(.TextMatrix(lngRow, Col�������), InStrRev(.TextMatrix(lngRow, Col�������), "("))
                End If
                
                .TextMatrix(lngRow, col����) = 1
                .TextMatrix(lngRow, col��ϱ���) = "" & rsInput!����
                .TextMatrix(lngRow, Col�������) = "" & rsInput!���� & strTmp
                .Cell(flexcpData, lngRow, Col�������) = .TextMatrix(lngRow, Col�������)
                                
                
                '�������ȷ������,����ݼ���ȷ�����
                If optInput(0).value Then
                    .TextMatrix(lngRow, colzy���ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, colzy����ID) = ""
                    strSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
                Else
                    .TextMatrix(lngRow, colzy����ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, colzy���ID) = ""
                    strSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
                End If
                Set rsTmp = New ADODB.Recordset
                On Error GoTo errH
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!��ĿID))
                If Not rsTmp.EOF Then
                    If optInput(0).value Then
                        .TextMatrix(lngRow, colzy����ID) = NVL(rsTmp!id)
                    Else
                        .TextMatrix(lngRow, colzy���ID) = NVL(rsTmp!id)
                    End If
                End If
                
                '��ҽ���ݼ�����ϲο�ȡ֤��
                Call Set��ҽ֤��(lngRow, Val(.TextMatrix(lngRow, colzy���ID)))
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col��ϱ���) = ""
            .TextMatrix(lngRow, Col�������) = .EditText
            .Cell(flexcpData, lngRow, Col�������) = .TextMatrix(lngRow, Col�������)
            .TextMatrix(lngRow, colzy���ID) = ""
            .TextMatrix(lngRow, colzy����ID) = ""
            .TextMatrix(lngRow, colzy֤��ID) = ""
        End If
        .Cell(flexcpForeColor, .FixedRows, colZY����, .Rows - 1, colZY����) = vbRed
    End With
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub ZYEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiagZY
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, Col�������) To colZY����
                If ZYCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= colZY���� Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Function Set��ҽ֤��(ByVal lngRow As Long, ByVal lng���ID As Long, Optional ByVal rsInput As Recordset) As Boolean
'���ܣ���ҽ���ݼ�����ϲο�ȡ֤��
'������rsInput-�����Ϊ�գ������ָ������ҩ֤���¼��
'���أ��Ƿ��ж�Ӧ��ϵ
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String
    
    With vsDiagZY
        'ȥ�����е�֤��
        If InStr(.TextMatrix(lngRow, Col�������), "(") > 0 And InStr(.TextMatrix(lngRow, Col�������), ")") > 0 Then
            strTmp = Mid(.TextMatrix(lngRow, Col�������), 1, InStrRev(.TextMatrix(lngRow, Col�������), "(") - 1)
        Else
            strTmp = .TextMatrix(lngRow, Col�������)
        End If
        If rsInput Is Nothing Then
            If lng���ID <> 0 Then
                strSQL = "Select Distinct a.֤����� as ID,a.֤��ID,a.֤������,b.���� as ֤�����" & _
                    " From ������ϲο� A,��������Ŀ¼ B" & _
                    " Where a.֤��ID=b.ID(+) And a.���ID=[1] And a.֤������ is Not NULL" & _
                    " Order by a.֤�����"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = Nothing
                Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng���ID)
                If Not rsTmp Is Nothing Then
                    .TextMatrix(lngRow, colzy֤��ID) = NVL(rsTmp!֤��id)
                    If Not IsNull(rsTmp!֤������) Then
                        .TextMatrix(lngRow, Col�������) = strTmp
                        .Cell(flexcpData, lngRow, Col�������) = .TextMatrix(lngRow, Col�������)
                        .TextMatrix(lngRow, col��ҽ֤��) = NVL(rsTmp!֤������)
                        .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
                        If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col��ҽ֤��)
                        mblnChange = True
                        .Tag = ""
                    End If
                    Set��ҽ֤�� = True
                Else
                    If blnCancel Then
                        Set��ҽ֤�� = True
                        If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, col��ҽ֤��)
                    Else
                        Set��ҽ֤�� = False
                    End If
                End If
            Else
                Set��ҽ֤�� = False
            End If
        Else
            .TextMatrix(lngRow, colzy֤��ID) = NVL(rsInput!��ĿID)
            .TextMatrix(lngRow, Col�������) = strTmp
            .Cell(flexcpData, lngRow, Col�������) = .TextMatrix(lngRow, Col�������)
            .TextMatrix(lngRow, col��ҽ֤��) = NVL(rsInput!����)
            .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
            If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col��ҽ֤��)
        End If
    End With
End Function

Private Function Get�������(ByVal lng���ID As Long, ByVal lng����ID As Long) As String
'���ܣ��������ID�򼲲�ID��ȡ�ֵ���е����ƣ�������ϼ�¼�е����ƿ������޸ĺ��,�����ǰ׺���׺�����Ա��ٴ��޸�ʱ�ж�
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If lng���ID <> 0 Then
        strSQL = "Select ���� From �������Ŀ¼ Where ID = [1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng���ID)
        If rsTmp.RecordCount > 0 Then Get������� = "" & rsTmp!����
    ElseIf lng����ID <> 0 Then
        strSQL = "Select ���� From ��������Ŀ¼ Where ID = [1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng����ID)
        If rsTmp.RecordCount > 0 Then Get������� = "" & rsTmp!����
    End If
    
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function GetRow(ByVal lng������� As Long) As Long
'���ܣ�����ָ��������͵ĵ�һ�����
    If InStr(",11,12,13,", "," & lng������� & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng�������), , colzy����)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng�������), , col����)
    End If
End Function

Private Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Private Function GetAdviceIDByDiag(ByVal lng���ID As Long) As String
'���ܣ��������ID��ȡ������ҽ��ID
    Dim strTmp As String, strҽ��IDs As String
    Dim lngPos As Long
    If Not mrsAdvice Is Nothing Then
        mrsAdvice.Filter = "���ID=" & lng���ID
        If Not mrsAdvice.EOF Then
            mrsAdvice.MoveFirst
            strҽ��IDs = mrsAdvice!ҽ��IDs
            lngPos = InStr(strҽ��IDs, mlng��ҽ��ID & "")
            If strҽ��IDs = mlng��ҽ��ID & "" Then
            '����ҽ��Ϊ��ǰҽ���������������ؿմ�
            ElseIf lngPos <= 0 Then
            '��ǰҽ��δ������ǰ���
                strTmp = strҽ��IDs
            Else
            'ҽ��ID���������ŵ��������ͨ���ַ����滻��
                If lngPos = 1 Then
                '��ǰҽ���ڿ�ͷλ��
                    strTmp = Replace(strҽ��IDs, mlng��ҽ��ID & ",", "")
                Else
                '��ǰҽ���ڷǿ�ͷλ��
                    strTmp = Replace(strҽ��IDs, "," & mlng��ҽ��ID, "")
                End If
            End If
        End If
    End If
    
    With grsDiagConn
        .Filter = "���ID=" & lng���ID
        .Sort = "��ʶID"
        Do While Not .EOF
            If Val(!��ʶID & "") <> mlngCur��ʶ Then
                strTmp = strTmp & "," & !��ʶID
            End If
            .MoveNext
        Loop
    End With
    
    GetAdviceIDByDiag = strTmp
End Function

Private Function CheckData() As Boolean
    Dim i As Long
    Dim j As Long
    Dim curDate As Date
    
    curDate = gobjComLib.zlDatabase.Currentdate
    
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Col�������) <> "" And .TextMatrix(i - 1, Col�������) = "" _
                And Val(.TextMatrix(i, col����)) = Val(.TextMatrix(i - 1, col����)) Then
                .Row = i - 1: .Col = Col�������
                Call ShowMessage(vsDiagXY, "���������������Ϣ��")
                Exit Function
            End If
            
            If Trim(.TextMatrix(i, Col�������)) <> "" Then
                If gobjComLib.zlCommFun.ActualLen(.TextMatrix(i, Col�������)) > 200 Then
                    .Row = i: .Col = Col�������
                    Call ShowMessage(vsDiagXY, IIf(.TextMatrix(i, col�������) = "", "��Ժ���", .TextMatrix(i, col�������)) & "����̫����ֻ����200���ַ���100�����֡�")
                    Exit Function
                End If
                If .TextMatrix(i, col����ʱ��) <> "" And Not .ColHidden(col����ʱ��) Then
                    If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col����ʱ��), "YYYY-MM-DD HH:mm") Then
                         .Row = i: .Col = col����ʱ��
                        Call ShowMessage(vsDiagXY, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                        Exit Function
                    End If
                End If
                If gobjComLib.zlCommFun.ActualLen(.TextMatrix(i, col��ע)) > 50 Then
                    .Row = i: .Col = col��ע
                    Call ShowMessage(vsDiagXY, """" & .TextMatrix(i, Col�������) & """�ı�ע����̫����ֻ����50���ַ���25�����֡�")
                    Exit Function
                End If
                If Val(.TextMatrix(i, col����)) = 5 Then     'Ժ�ڸ�Ⱦ
                    If .TextMatrix(i, col��Ժ���) = "" And Not .ColHidden(col��Ժ���) Then
                        .Row = i: .Col = col��Ժ���
                        If ShowMessage(vsDiagXY, "Ժ�ڸ�Ⱦ�ĳ�Ժ���û����д���Ƿ������", True) = vbNo Then Exit Function
                    End If
                End If
                If Val(.TextMatrix(i, col����)) = 3 Then
                    If .TextMatrix(i, col��Ժ���) = "" And Not .ColHidden(col��Ժ���) Then
                        .Row = i: .Col = col��Ժ���
                        Call ShowMessage(vsDiagXY, "����д��Ժ��ϵĳ�Ժ�����")
                        Exit Function
                    ElseIf Val(.TextMatrix(i - 1, col����)) <> 3 And InStr(.TextMatrix(i, col��Ժ���), "����") > 0 And mbln���� And Not .ColHidden(col��Ժ���) Then
                        .Row = i: .Col = col��Ժ���
                        If ShowMessage(vsDiagXY, "�ò��˽���������������Ժ���ѡ��Ϊ�������Ƿ������", True) = vbNo Then Exit Function
                    ElseIf Val(.TextMatrix(i - 1, col����)) = 3 And InStr(.TextMatrix(GetRow(3), col��Ժ���), "����") = 0 And InStr(.TextMatrix(i, col��Ժ���), "����") > 0 And Not .ColHidden(col��Ժ���) Then
                        .Row = i: .Col = col��Ժ���
                        Call ShowMessage(vsDiagXY, "��Ҫ��ϵĳ�Ժ�����Ϊ��������������ϵĳ�Ժ���ȴΪ������")
                        Exit Function
                    ElseIf .TextMatrix(i, col�������) = "��Ժ���" And Not .ColHidden(col��Ժ���) Then
                        If mlng�����ж� <> 0 Then
                            '��Ҫ�����Ҫ�����˵��ⲿԭ��
                            If InStr("ST", Left(.TextMatrix(i, col��ϱ���), 1)) > 0 And Left(.TextMatrix(i, col��ϱ���), 1) <> "" Then
                                '��Ҫ�����ж��ⲿԭ��
                                If .TextMatrix(GetRow(7), Col�������) = "" Then
                                    If Not vsDiagZY.Visible Then
                                        .Row = GetRow(7): .Col = Col�������
                                        If mlng�����ж� = 1 Then
                                            Call ShowMessage(vsDiagXY, "����д�����ж���ԭ��")
                                            Exit Function
                                        Else
                                            If ShowMessage(vsDiagXY, "û����д�����ж���ԭ��,�Ƿ������", True) = vbNo Then Exit Function
                                        End If
                                    End If
                                End If
                            Else
                                If .TextMatrix(GetRow(7), Col�������) <> "" Then
                                    .Row = GetRow(7): .Col = Col�������
                                    If mlng�����ж� = 1 Then
                                        Call ShowMessage(vsDiagXY, "������д�����ж���ԭ��")
                                        Exit Function
                                    Else
                                        If ShowMessage(vsDiagXY, "��Ժ����������ж���ԭ�򲻷�,�Ƿ������", True) = vbNo Then Exit Function
                                    End If
                                End If
                            End If
                        End If
                        If mlng������� <> 0 Then
                            '��Ҫ�����Ҫ��д������ϵ��ⲿԭ��
                            If InStr("CD", Left(.TextMatrix(i, col��ϱ���), 1)) > 0 And Left(.TextMatrix(i, col��ϱ���), 1) <> "" Then
                                '��Ҫ������ϵ��ⲿԭ��
                                If .TextMatrix(GetRow(6), Col�������) = "" Then
                                    If Not vsDiagZY.Visible Then
                                        .Row = GetRow(6): .Col = Col�������
                                        If mlng������� = 1 Then
                                            Call ShowMessage(vsDiagXY, "����д������ϡ�")
                                            Exit Function
                                        Else
                                            If ShowMessage(vsDiagXY, "û����д�������,�Ƿ������", True) = vbNo Then Exit Function
                                        End If
                                    End If
                                End If
                            Else
                                If .TextMatrix(GetRow(6), Col�������) <> "" Then
                                    .Row = GetRow(6): .Col = Col�������
                                    If mlng������� = 1 Then
                                        Call ShowMessage(vsDiagXY, "������д������ϡ�")
                                        Exit Function
                                    Else
                                        If ShowMessage(vsDiagXY, "��Ժ����벡����ϲ���,�Ƿ������", True) = vbNo Then Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    For j = GetRow(3) To .Rows - 1
                        If Val(.TextMatrix(j, col����)) = 3 Then
                            If j <> i And .TextMatrix(j, Col�������) <> "" Then
                                If .TextMatrix(j, Col�������) = .TextMatrix(i, Col�������) Then
                                    .Row = i: .Col = Col�������
                                    Call ShowMessage(vsDiagXY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                    Exit Function
                                ElseIf Val(.TextMatrix(i, col����ID)) <> 0 Then
                                    If Val(.TextMatrix(j, col����ID)) = Val(.TextMatrix(i, col����ID)) Then
                                        .Row = i: .Col = Col�������
                                        Call ShowMessage(vsDiagXY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                        Exit Function
                                    End If
                                ElseIf Val(.TextMatrix(i, col���ID)) <> 0 Then
                                    If Val(.TextMatrix(j, col���ID)) = Val(.TextMatrix(i, col���ID)) Then
                                        .Row = i: .Col = Col�������
                                        Call ShowMessage(vsDiagXY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Next
    End With
        
    If vsDiagZY.Visible Then
        With vsDiagZY
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, Col�������) <> "" And .TextMatrix(i - 1, Col�������) = "" _
                    And Val(.TextMatrix(i, colzy����)) = Val(.TextMatrix(i - 1, colzy����)) Then
                    .Row = i - 1: .Col = Col�������
                    Call ShowMessage(vsDiagZY, "���������������Ϣ��")
                    Exit Function
                End If
            
                If Trim(.TextMatrix(i, Col�������)) <> "" Then
                    If gobjComLib.zlCommFun.ActualLen(.TextMatrix(i, Col�������)) > 200 Then
                        .Row = i: .Col = Col�������
                        Call ShowMessage(vsDiagZY, IIf(.TextMatrix(i, col�������) = "", "��Ժ���", .TextMatrix(i, col�������)) & "����̫����ֻ����200���ַ���100�����֡�")
                        Exit Function
                    End If
                    If .TextMatrix(i, col����ʱ��) <> "" And Not .ColHidden(col����ʱ��) Then
                        If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col����ʱ��), "YYYY-MM-DD HH:mm") Then
                             .Row = i: .Col = col����ʱ��
                            Call ShowMessage(vsDiagXY, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                            Exit Function
                        End If
                    End If
                    If gobjComLib.zlCommFun.ActualLen(.TextMatrix(i, col��ע)) > 50 Then
                        .Row = i: .Col = col��ע
                        Call ShowMessage(vsDiagZY, """" & .TextMatrix(i, Col�������) & """�ı�ע����̫����ֻ����50���ַ���25�����֡�")
                        Exit Function
                    End If
                    If Val(.TextMatrix(i, colzy����)) = 13 Then
                        If .TextMatrix(i, col��Ժ���) = "" And Not .ColHidden(col��Ժ���) Then
                            .Row = i: .Col = col��Ժ���
                            Call ShowMessage(vsDiagZY, "����д��Ժ��ϵĳ�Ժ�����")
                            Exit Function
                        ElseIf Val(.TextMatrix(i - 1, colzy����)) = 13 And InStr(.TextMatrix(GetRow(13), col��Ժ���), "����") = 0 And InStr(.TextMatrix(i, col��Ժ���), "����") > 0 And Not .ColHidden(col��Ժ���) Then
                            .Row = i: .Col = col��Ժ���
                            Call ShowMessage(vsDiagZY, "��Ҫ��ϵĳ�Ժ�����Ϊ��������������ϵĳ�Ժ���ȴΪ������")
                            Exit Function
                        End If
                        
                        For j = GetRow(13) To .Rows - 1
                            If j <> i And .TextMatrix(j, Col�������) <> "" Then
                                If .TextMatrix(j, Col�������) = .TextMatrix(i, Col�������) Then
                                    .Row = i: .Col = Col�������
                                    Call ShowMessage(vsDiagZY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                    Exit Function
                                ElseIf Val(.TextMatrix(i, colzy����ID)) <> 0 Then
                                    If Val(.TextMatrix(j, colzy����ID)) = Val(.TextMatrix(i, colzy����ID)) Then
                                        .Row = i: .Col = Col�������
                                        Call ShowMessage(vsDiagZY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End With
    End If
    CheckData = True
End Function

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'���ܣ���ʾ��ʾ��Ϣ����λ��������Ŀ��
    Dim lngColor As Long
    
    lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
    Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If

    objTmp.CellBackColor = lngColor
    If objTmp.Enabled And objTmp.Visible Then objTmp.SetFocus
    Me.Refresh
End Function

Private Function SetPublicFontSize(ByVal bytSize As Byte, Optional ByVal strOther As String)
'���ܣ����ô��弰���пؼ��������С
'������
'      bytSize:����Ϊ9������,0:����Ϊ9������,1,����Ϊ12������
'      strOther:�������������õĿؼ��������ļ���,��ʽΪ����������1,��������2,��������3,....
'˵����1.����漰��VsFlexGrid�ȱ��ؼ�����Ҫ�������ڵĻ������µ����п���и�
'      2.�������δ�г��������ؼ����Զ���ؼ�,��Ҫ���ض�����ָ�������С����ش���ģ������ⵥ������

    Dim objCtrol As Control
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Me.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In Me.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "ReportControl", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKind"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '����CommandBars�û��Զ���ؼ���ȡobjCtrol.Container�����
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.Name
            Err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Height = Me.TextHeight("��") + 20
                        'Label�����Ҫ���е���
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = Me.TextWidth("����" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = Me.TextWidth("2012-01-01    ")
                        objCtrol.Height = Me.TextHeight("��") + IIf(bytSize = 0, 100, 120)
                Case "TextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = Me.TextHeight("��")
                Case "MaskEdBox"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = Me.TextWidth(objCtrol.Mask)
                        objCtrol.Height = Me.TextHeight("��")
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        lngOldSize = objCtrol.FontSize
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                Case "IDKind"
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
            End Select
        End If
    Next
End Function

Private Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True) As String
'���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd[ HH:mm])
'������blnTime=�Ƿ���ʱ�䲿��
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = gobjComLib.zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '���봮�а������ڷָ���
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                'ֻ���������ڲ���
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                'ֻ������ʱ�䲿��
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '����Ƿ�����,����ԭ����
            strTmp = strText
        End If
    Else
        '���������ڷָ���
        If Len(strTmp) <= 2 Then
            '��������dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '��������MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '��������yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '��������MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '��������yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '��������yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) And Not blnTime Then
        strTmp = Format(strTmp, "yyyy-MM-dd")
    End If
    GetFullDate = strTmp
End Function
