VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCaseTendBodySetShowData 
   Caption         =   "����������ʾ"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBodySetShowData.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   10350
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picStb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1440
      ScaleHeight     =   360
      ScaleWidth      =   2415
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2415
      Begin VB.Label lblStb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox picThis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   240
      ScaleHeight     =   3015
      ScaleWidth      =   4935
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   4935
      Begin VB.PictureBox picShow 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   4335
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4335
         Begin VSFlex8Ctl.VSFlexGrid vfgShow 
            Height          =   615
            Left            =   0
            TabIndex        =   4
            Top             =   360
            Width           =   3975
            _cx             =   7011
            _cy             =   1085
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   0
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   2
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VB.Label lblTmp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   90
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   1095
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   3735
         _cx             =   6588
         _cy             =   1931
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   7
         FixedRows       =   2
         FixedCols       =   2
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         OwnerDraw       =   1
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
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��:2011-02-25"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1350
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6180
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBodySetShowData.frx":08CA
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15346
            Object.ToolTipText     =   "��ӡ����Ϣ"
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBodySetShowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnTemType As Boolean 'TRUE ר�����µ�,FALSE ��׼���µ�
Private mstrCurveItem As String  'ר�����µ�������������Ŀ
'���˻�����Ϣ
Private Type type_Patient
    lng����ID As Long
    lng��ҳID As Long
    lng�ļ�ID As Long
    lngӤ�� As Long
    lng����ID As Long
    lng�����ȼ� As Long
    lng��ʽID As Long
End Type
Private mT_Patient As type_Patient

'������:
Private mcbrToolBar As CommandBar
Private mrsPoint As New ADODB.Recordset
Private mrs��λ As New ADODB.Recordset
Private mrsCopy As New ADODB.Recordset '���ڻ�ԭ������Ϣ

Private Const mFontSize As Integer = 9 '���������ʼ��СΪ9������
Private mintBigSize As Integer
Private mstrActiveItem As String
Private mint����Ӧ�� As Integer
Private marrTime() As String
Private mDTime As Date
Private mDEndTime As Date
Private mblnChage As Boolean
Private mblnOK As Boolean
Private mblnMove As Boolean
Private mstrSQL As String
Private mblnInit As Boolean
Private mintColSel As Integer
Private mblnFileBack As Boolean
Private mbln��Ժ As Boolean '���˳�Ժ���ļ�����ΪTRUE
Private mbln����������ʾ As Boolean

Public Function ShowEdit(ByVal frmParent As Object, ByVal strParam As String, ByVal DTime As Date, ByVal DEndTime As Date, _
    ByVal int����Ӧ�� As Integer, Optional blnMove As Boolean = False, Optional ByVal bytSize As Byte = 0) As Boolean
'----------------------------------------------------------------------------------------------------------
'����:�������µ��༭����
'����:frmParent ������,strParam ��ʽ:����ID;��ҳId;�ļ�ID;Ӥ��;����ID;���������ȼ�
'     Dtime Ҫ�༭���µ���ʱ�� ��ʽΪ YYYY-MM-DD HH:mm:ss:DEndTime ���µ�����ʱ�� ; int����Ӧ��=2 ��ʾ���������ʹ��� blnMove ��ʷ�����Ƿ�ת��
'bytSize 0-9������ 1-12������
'----------------------------------------------------------------------------------------------------------
    Dim arrParam() As String
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then Exit Function
    
    mblnChage = False
    mblnMove = False
    mblnInit = False
    mblnOK = False
    mblnFileBack = False
    mT_Patient.lng����ID = 0
    mT_Patient.lng�����ȼ� = 3
    
    mT_Patient.lng����ID = arrParam(0)
    mT_Patient.lng��ҳID = arrParam(1)
    mT_Patient.lng�ļ�ID = arrParam(2)
    mT_Patient.lngӤ�� = arrParam(3)
    If UBound(arrParam) > 3 Then mT_Patient.lng����ID = arrParam(4)
    If UBound(arrParam) > 4 Then mT_Patient.lng�����ȼ� = arrParam(5)
    
    If mT_Patient.lng����ID = 0 And mT_Patient.lng��ҳID = 0 And mT_Patient.lng����ID = 0 Then
        MsgBox "�ļ�ID,����ID,��ҳID����Ϊ��,����!", vbInformation, gstrSysName
        Exit Function
    End If
    
    mDTime = DTime
    mDEndTime = DEndTime
    mbln����������ʾ = (Val(zlDatabase.GetPara("���������(����/����)��ʽ¼��", glngSys, 1255, 0)) = 1)
    mint����Ӧ�� = int����Ӧ��
    mblnMove = blnMove
    
    If Not OpenPatientInfo Then Exit Function
    If Not ChekPatientOut(mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��) Then Exit Function
    
    mintBigSize = bytSize   'zldatabase.GetPara("�����ļ���ʾģʽ", glngSys, 1255, 0)
    Me.Font.Size = IIf(mintBigSize = 0, 9, 12)
    '����ļ��Ƿ�鵵
    mblnFileBack = CheckFileBack(mT_Patient.lng�ļ�ID, mblnMove)
    If mblnFileBack = True Then lblStb.Caption = "�������������Ѿ��鵵,���������������޸�.": lblStb.ForeColor = 255

    Call InitCommandBars
    Call GetTableRowName
    
    mblnInit = True
    Me.Show 1
    ShowEdit = mblnOK
End Function

Public Function ChekPatientOut(ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intBaby As Long) As Boolean
'-----------------------------------------------------------------------------------------------
'����:��ȡ���µ���ʼʱ��ͽ���ʱ�� ����鲡���Ƿ��Ժ
'-----------------------------------------------------------------------------------------------
    Dim strSQL As String, strNewSql As String
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCurrDate As String
    
    mbln��Ժ = False
    On Error GoTo Errhand
        
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ),����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "(SELECT " & vbNewLine & _
                "        ����ID, ��ҳID, Ӥ��ʱ��, DECODE(NVL(Ӥ��, 0), 0, DECODE(NVL(��Ժ����, ''), '', 0, 1), DECODE(NVL(Ӥ��ʱ��, ''), '', 0, 1)) ��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID, A.��ҳID, B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����, B.Ӥ��" & vbNewLine & _
                "              FROM ������ҳ A," & vbNewLine & _
                "                   (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                     FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                     WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND NVL(B.Ӥ��, 0) <> 0 AND B.������� = 'Z'  " & vbNewLine & _
                "                      And Instr(',3,5,11,', ',' || c.�������� || ',') > 0 AND B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "              WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "              ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2) E"

    '˵��:Ŀǰ����ר�����µ������˿���ͬʱ���ڶ�����µ������µ���ʼʱ�����ֹʱ��Ĺ�������:
    '����ļ��Ŀ�ʼʱ�䲻Ϊ�ղ��Ҵ��ڵ��ڲ�����Ժʱ���Ӥ������ʱ��,���µ��Ŀ�ʼʱ�����ļ���ʼʱ��Ϊ׼,�����Բ�����Ժʱ���Ӥ������ʱ��Ϊ׼
    '����ļ�����ֹʱ�䲻Ϊ�ղ���С�ڵ��ڲ��˻�Ӥ����Ժʱ�䣨δ��Ժ���ܲ��ܴ��ڵ�ǰʱ�䣩,���µ�����ʱ�����ļ���ʼʱ��Ϊ׼���������µ�����ʱ���Բ��˻�Ӥ����Ժʱ��Ϊ׼(δ��ԺΪ��ǰʱ��)
    '����ļ�����ֹʱ��Ϊ��,����ԭ�з�ʽ,��������Ѿ���Ժ�����ѳ�Ժʱ��Ϊ׼,δ��Ժ���ѵ�ǰʱ������ݽ���ʱ��Ϊ׼.
    strSQL = " SELECT  DECODE(D.��ʼʱ��,NULL,DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��)," & vbNewLine & _
            "               DECODE(SIGN(D.��ʼʱ�� - DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��))," & vbNewLine & _
            "                      1," & vbNewLine & _
            "                      D.��ʼʱ��," & vbNewLine & _
            "                      DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��))) AS ��ʼ," & vbNewLine & _
            "       DECODE(D.����ʱ��," & vbNewLine & _
            "               NULL," & vbNewLine & _
            "               DECODE(E.��¼," & vbNewLine & _
            "                      0," & vbNewLine & _
            "                      DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹ) - D.����ʱ��), 1, NVL(E.Ӥ��ʱ��, A.��ֹ), D.����ʱ��)," & vbNewLine & _
            "                      NVL(E.Ӥ��ʱ��, A.��ֹ))," & vbNewLine & _
            "               DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹ) - D.����ʱ��), 1, D.����ʱ��, NVL(E.Ӥ��ʱ��, A.��ֹ))) ��ֹ," & vbNewLine & _
            "       DECODE(D.����ʱ��, NULL, E.��¼, 1) ��¼" & vbNewLine & _
            " FROM (SELECT ����ID, ��ҳID, MIN(��ʼʱ��) AS ��ʼ, MAX(NVL(��ֹʱ��, SYSDATE)) AS ��ֹ" & vbNewLine & _
            "       FROM ���˱䶯��¼" & vbNewLine & _
            "       WHERE ��ʼʱ�� IS NOT NULL AND ����ID = [2] AND ��ҳID = [3]" & vbNewLine & _
            "       GROUP BY ����ID, ��ҳID) A," & vbNewLine & _
            "     (SELECT ����ID, ��ҳID, ����ʱ�� FROM ������������¼ WHERE ����ID = [2] AND ��ҳID = [3] AND ��� = [4]) B," & vbNewLine & _
            "     (SELECT NVL(����ʱ��, SYSDATE) ����ʱ��, ��ʼʱ��, ����ʱ��" & vbNewLine & _
            "       FROM (SELECT MAX(B.����ʱ��) ����ʱ��, MAX(A.��ʼʱ��) ��ʼʱ��, MAX(A.����ʱ��) ����ʱ��" & vbNewLine & _
            "              FROM ���˻����ļ� A, ���˻������� B" & vbNewLine & _
            "              WHERE A.ID = B.�ļ�ID(+) AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND A.Ӥ�� = [4])) D," & vbNewLine & _
            "  " & strNewSql & vbNewLine & _
            " WHERE A.����ID = E.����ID AND A.��ҳID = E.��ҳID AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng�ļ�ID, lng����ID, lng��ҳID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        strBeginDate = Format(rsTemp!��ʼ, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!��ֹ, "YYYY-MM-DD HH:MM:SS")
        mbln��Ժ = Not (Val(rsTemp!��¼) = 0)
    Else
        MsgBox "�޴˲��˱���סԺ��Ϣ,����!", vbInformation, gstrSysName
        Exit Function '�������˱䶯��Ϣ�˳�
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    mDEndTime = strEndDate
    If CDate(mDEndTime) < CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss")) And Not mbln��Ժ Then mDEndTime = CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss"))
    
    '���˳�Ժ�ѳ�Ժʱ��Ϊ��ֹʱ��
    If mbln��Ժ = True Then
        '��Ժʱ�����Ժʱ�������ͬһ�У��򽫳�Ժʱ�����һ�У���������:��ԺҲҪ¼�����£�
        If mblnTemType = False Then '��׼���µ�
            mDEndTime = Format(RetrunEndTime(CDate(strBeginDate), CDate(mDEndTime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        Else 'ר�����µ�
            mDEndTime = Format(RetrunEndTimeNew(CDate(strBeginDate), CDate(mDEndTime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        End If
    End If
        
    If Not (CDate(mDTime) >= CDate(strBeginDate) And CDate(mDTime) <= CDate(mDEndTime)) Then
        If Int(CDate(strBeginDate)) = Int(CDate(mDEndTime)) Then
            mDTime = Format(strBeginDate, "YYYY-MM-DD HH:mm:ss")
        Else
            mDTime = Format(Int(CDate(mDEndTime)), "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    ChekPatientOut = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenPatientInfo() As Boolean
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo Errhand
    '��ȡ������Ϣ
    mstrSQL = "Select ��Ժ����ID from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng����ID, mT_Patient.lng��ҳID)
    If rsTmp.BOF = False Then
        mT_Patient.lng����ID = Val(zlCommFun.Nvl(rsTmp("��Ժ����ID").Value))
    End If
    
    '��ȡ�����ȼ�
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As �����ȼ� From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng����ID, mT_Patient.lng��ҳID)
    If rsTmp.BOF = False Then mT_Patient.lng�����ȼ� = zlCommFun.Nvl(rsTmp("�����ȼ�"), 3)
    
    '��ȡ���µ���Ϣ
    mblnTemType = False
    mstrSQL = "Select B.����,B.ID From ���˻����ļ� A,�����ļ��б� B Where A.��ʽID=B.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng�ļ�ID)
    If rsTmp.BOF = False Then
        mblnTemType = (Nvl(rsTmp!����) = "1")
        mT_Patient.lng��ʽID = rsTmp!Id
    End If
    
    If mblnTemType = True Then
        gintHourBegin = T_BodyStyle.lng��ʼʱ��
    Else
        gintHourBegin = zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4)
        T_BodyStyle.lng��ʼʱ�� = gintHourBegin
        T_BodyStyle.lngʱ���� = 4
        T_BodyStyle.lng������ = 6
        T_BodyStyle.lng���� = 7
    End If
    
    OpenPatientInfo = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitCommandBars()
'--------------------------------------------------------------------------------
'����:��ʼ��������
'--------------------------------------------------------------------------------
    Dim cbrControl As CommandBarButton
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrLable As CommandBarControl
    Dim CtlFont As StdFont
    
    On Error GoTo Errhand
      '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "�˵���"
    cbsMain.ActiveMenuBar.Visible = False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
        Set CtlFont = .Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
        Set .Font = CtlFont
    End With

  '------------------------------------------------------------------------------------------------------------------
    '����������
    Set mcbrToolBar = cbsMain.Add("��׼", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve, "���߱༭"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CurveTable, "����༭")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    
    '���ù������ı���ͼ����ʾ��ʽ
    For Each cbrControl In mcbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    '�����
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("Q"), conMenu_Edit_Curve
        .Add FCONTROL, Asc("T"), conMenu_Edit_CurveTable
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetTableRowName() As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTmpName0 As String
    Dim strTmpCurve As String
    Dim arrItem() As Variant, i As Integer
    
    On Error GoTo Errhand
    
    Call GetPainDegreeNO
    '��ȡ����������Ŀ
    mstrCurveItem = ""
    If mblnTemType = False Then
        mstrSQL = _
                " Select A.��¼��,A.��¼�� as ��Ŀ����,A.��Ŀ��� as ��Ŀ��,A.��λ" & _
                " From ���¼�¼��Ŀ A,����������Ŀ B,������¼��Ŀ C " & _
                " Where c.��ĿID=B.ID(+) And A.��Ŀ���=C.��Ŀ��� And ��Ŀ����=1 And (nvl(A.��¼��,1)<>2 Or (nvl(A.��¼��,1)=2 and A.��Ŀ���=3)) And Nvl(C.Ӧ�÷�ʽ,0)=1 AND C.�����ȼ�>=[1] And Nvl(C.���ò���,0) In (0,[3]) " & _
                " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[2]))) " & _
                " Order by Decode(A.��Ŀ���,1,0,1),A.�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng�����ȼ�, mT_Patient.lng����ID, IIf(mT_Patient.lngӤ�� = 0, 1, 2))
    Else 'ר�����µ�
        mstrCurveItem = T_BodyItem.str������Ŀ
        If InStr(1, "," & mstrCurveItem & ",", "," & gint���� & ",") = 0 Then
            arrItem = Array(T_BodyItem.str��������)
            For i = 0 To UBound(arrItem)
                If Val(arrItem(i)) = gint���� Then
                    mstrCurveItem = mstrCurveItem & "," & gint����
                    Exit For
                End If
            Next
        End If
        If Left(mstrCurveItem, 1) = "," Then mstrCurveItem = Mid(mstrCurveItem, 2)
        mstrSQL = _
                " Select /*+ RULE */ A.��¼��,A.��¼�� as ��Ŀ����,A.��Ŀ��� as ��Ŀ��,A.��λ" & _
                " From ���¼�¼��Ŀ A,����������Ŀ B,������¼��Ŀ C,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) D " & _
                " Where C.��ĿID=B.ID(+) And A.��Ŀ���=C.��Ŀ��� And (A.��¼��<>2 OR (A.��¼��=2 And A.��Ŀ���=3)) And NOT (C.Ӧ�÷�ʽ=2 And C.��Ŀ���=-1)" & _
                " And C.��Ŀ���=D.COLUMN_VALUE Order by Decode(A.��Ŀ���,1,0,1),A.�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mstrCurveItem)
    End If
    strTmpName0 = ""
    With rsTemp
        Do While Not .EOF
            strTmpName0 = strTmpName0 & ";" & zlCommFun.Nvl(!��Ŀ��) & "'" & zlCommFun.Nvl(!��Ŀ����) & IIf(zlCommFun.Nvl(!��λ) = "", "", "(" & zlCommFun.Nvl(!��λ) & ")") & "'" & zlCommFun.Nvl(!��Ŀ����)
        .MoveNext
        Loop
    End With
    
    If Left(strTmpName0, 1) = ";" Then strTmpName0 = Mid(strTmpName0, 2)
    
    Call InitTable(strTmpName0)
    'ˢ������
    Call zlRefreshData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitTable(ByVal strTmpName As String)
    Dim intCOl As Integer, intRow As Integer
    Dim strColName As String
    Dim arrColName() As String, arrColTime() As String
    
    strColName = InitTime
    arrColName = Split(strColName, "[LPF]")
    
    On Error GoTo Errhand
    
    With vfgThis
        .Clear
        .FixedCols = 3
        .FixedRows = 2
        .Rows = 3
        .Cols = .FixedCols + T_BodyStyle.lng������
        .ColHidden(0) = True
        .ColWidth(0) = 0
        .ColHidden(1) = True
        .ColWidth(1) = 0
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeCol(2) = True
        .MergeRow(0) = True
        .MergeRow(1) = True
        .Col = .FixedCols: .Row = .FixedRows
        .ColSel = .Col
        .RowSel = .Row
    
        vfgThis.Font.Size = mFontSize + mFontSize * mintBigSize / 3
       
        For intRow = 0 To .FixedRows - 1
            arrColTime = Split(arrColName(intRow), ";")
            For intCOl = .FixedCols - 1 To .Cols - 1
                .TextMatrix(intRow, intCOl) = arrColTime(intCOl + 1 - .FixedCols)
            Next intCOl
            If intRow = 0 Then
                .RowHeight(intRow) = 300 + 300 * mintBigSize / 3
            Else
                .RowHeight(intRow) = 400 + 400 * mintBigSize / 3
            End If
        Next intRow
        
        '�����п�
        For intCOl = .FixedCols - 1 To .Cols - 1
            If intCOl = .FixedCols - 1 Then
                .ColWidth(intCOl) = 1300 + 1300 * mintBigSize / 3
            Else
                .ColWidth(intCOl) = 1200 + 1200 * mintBigSize / 3
            End If
        Next intCOl
        '�̶���ͷ��ʽ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = RGB(0, 0, 255)
        .Cell(flexcpBackColor, 0, 0, .FixedRows - 1, .Cols - 1) = &H8000000F
        
        '�����е�ͷ����Ϣ
        arrColName = Split(strTmpName, ";")
        .Rows = UBound(arrColName) + .FixedRows + 1
        For intRow = .FixedRows To .Rows - 1
            arrColName(intRow - .FixedRows) = arrColName(intRow - .FixedRows) & String(3 - UBound(Split(arrColName(intRow - .FixedRows), "'")), "'")
            .RowHeight(intRow) = 300 + 300 * mintBigSize / 3
            .TextMatrix(intRow, 0) = Split(arrColName(intRow - .FixedRows), "'")(0)
            .TextMatrix(intRow, 1) = Split(arrColName(intRow - .FixedRows), "'")(2)
            .TextMatrix(intRow, 2) = Split(arrColName(intRow - .FixedRows), "'")(1)
        Next intRow
        .Cell(flexcpBackColor, .FixedRows, .FixedCols - 1, .Rows - 1, .FixedCols - 1) = &H8000000F
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With
    
    vfgThis.Cell(flexcpText, vfgThis.FixedRows, vfgThis.FixedCols, vfgThis.Rows - 1, vfgThis.Cols - 1) = ""
    
    With vfgShow
        .RowHeight(-1) = 300 + 300 * mintBigSize / 3
        .ColWidth(-1) = 1200 + 1200 * mintBigSize / 3
        .FixedRows = 0
        .FixedCols = 1
        .Rows = 2
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = &H0&
        .ScrollBars = flexScrollBarBoth
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function zlRefreshData() As Boolean
'---------------------------------------------------------------
'����:��ȡ����ĳ���ڵ�����������Ϣ
'---------------------------------------------------------------
    '��� Ϊ���˻�����ϸ��ID    IDΪ�������»���������ʱ���ʵ����� ,��ע��¼��Ϣ���ݿ����Ƿ�Ϊ��ʾ
    gstrFields = "���," & adDouble & ",18|��ֵ," & adLongVarChar & ",400|��λ," & adLongVarChar & ",200|" & _
         "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",40|" & _
         "����," & adDouble & ",1|������Դ," & adDouble & ",1|��ʾ," & adDouble & ",1|��ע," & adDouble & ",1|״̬," & adDouble & ",1|ʱ���," & adLongVarChar & ",20|�к�," & _
         adDouble & ",1|ID," & adDouble & ",18"
    Call Record_Init(mrsPoint, gstrFields)
    gstrFields = "���|��ֵ|��λ|���|ʱ��|��Ŀ���|��Ŀ����|����|������Դ|��ʾ|��ע|״̬|ʱ���|�к�|ID"
    
    Dim rsTmp As New ADODB.Recordset
    Dim strFidlds As String, strParam As String, strPart As String
    Dim arrValue() As String
    Dim lng��Ŀ��� As Long, lngCol As Long
    Dim str��Ŀ���� As String
    Dim int��ʾ As Integer, int��ע As Integer
    Dim blnAdd As Boolean, blnAllow As Boolean
    Dim intRow As Integer, intCOl As Integer
    Dim strTime As String
    Dim int��� As Integer
    Dim strEndTime As String
    
    On Error GoTo Errhand
    
    lblTime.Caption = "ʱ��:" & Format(mDTime, "YYYY-MM-DD")
    
    '��ȡ��λ
    mstrSQL = "Select ��Ŀ���,��λ,ȱʡ�� From ���²�λ"
    Call zlDatabase.OpenRecordset(mrs��λ, mstrSQL, Me.Caption)
    
    If CDate(Format(mDTime, "YYYY-MM-DD")) = CDate(Format(mDEndTime, "YYYY-MM-DD")) Then
        strEndTime = Format(CDate(mDEndTime), "YYYY-MM-DD HH:mm:ss")
    Else
        strEndTime = Format((Format(mDTime, "YYYY-MM-DD") & " 23:59:59"), "YYYY-MM-DD HH:mm:ss")
    End If
    
    '��ȡĳʱ��ε�����������������
    If mblnTemType = False Then
        mstrSQL = _
        " SELECT C.ID ���,A.����ʱ�� As ʱ��,C.��ʾ,c.��¼���� As ��ֵ,c.���²�λ,c.���Ժϸ�,D.��¼��,E.������Ŀ,D.��Ŀ���,DECODE(D.��Ŀ���,-1,1,C.��¼���) ��¼���,C.δ��˵��,C.������Դ" & vbNewLine & _
        "                    FROM ���˻����ļ� B,���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,������¼��Ŀ E" & vbNewLine & _
        "                    Where B.ID=A.�ļ�ID" & vbNewLine & _
        "                        AND A.ID = C.��¼ID" & vbNewLine & _
        "                        AND B.ID=[1]" & vbNewLine & _
        "                        AND Nvl(B.Ӥ��,0)=[4]" & vbNewLine & _
        "                        AND B.����id=[2]" & vbNewLine & _
        "                        AND B.��ҳid=[3]" & vbNewLine & _
        "                        AND D.��Ŀ���=C.��Ŀ���" & vbNewLine & _
        "                        AND C.��¼����=1" & vbNewLine & _
        "                        AND E.��Ŀ���=D.��Ŀ���" & vbNewLine & _
        "                        AND E.�����ȼ�>=[7]" & vbNewLine & _
        "                        AND (nvl(D.��¼��,1)<>2 Or (nvl(D.��¼��,1)=2 and D.��Ŀ���=3))" & _
        "                        And A.����ʱ�� BETWEEN [5] And [6] And C.��ֹ�汾 Is Null" & vbNewLine & _
        "                        AND (nvl(E.Ӧ�÷�ʽ,0)=1 OR ( -1=[10] and nvl(E.Ӧ�÷�ʽ,0)=2))" & vbNewLine & _
        "                        AND nvl(E.���ò���,0) in (0,[8]) AND (E.���ÿ���=1 or ( E.���ÿ���=2 AND Exists (select 1 from �������ÿ��� D where D.��Ŀ���=E.��Ŀ��� and D.����ID=[9])))" & vbNewLine & _
        "                    Order By A.����ʱ��,DECODE(D.��Ŀ���,-1,1,0),DECODE(D.��Ŀ���,-1,1,C.��¼���),D.��¼��"
    
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
            
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, _
             CDate(mDTime), CDate(strEndTime), mT_Patient.lng�����ȼ�, IIf(mT_Patient.lngӤ�� = 0, 1, 2), mT_Patient.lng����ID, IIf(mint����Ӧ�� = 2, -1, 0))
    Else 'ר�����µ�
        mstrSQL = _
        " SELECT /*+ RULE */ C.ID ���,A.����ʱ�� As ʱ��,C.��ʾ,c.��¼���� As ��ֵ,c.���²�λ,c.���Ժϸ�,D.��¼��,E.������Ŀ,D.��Ŀ���,DECODE(D.��Ŀ���,-1,1,C.��¼���) ��¼���,C.δ��˵��,C.������Դ" & vbNewLine & _
        "                    FROM ���˻����ļ� B,���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,������¼��Ŀ E,Table(Cast(f_num2list([7]) As zlTools.t_Numlist)) F" & vbNewLine & _
        "                    Where B.ID=A.�ļ�ID" & vbNewLine & _
        "                        AND A.ID = C.��¼ID" & vbNewLine & _
        "                        AND B.ID=[1]" & vbNewLine & _
        "                        AND Nvl(B.Ӥ��,0)=[4]" & vbNewLine & _
        "                        AND B.����id=[2]" & vbNewLine & _
        "                        AND B.��ҳid=[3]" & vbNewLine & _
        "                        AND D.��Ŀ���=C.��Ŀ���" & vbNewLine & _
        "                        AND C.��¼����=1" & vbNewLine & _
        "                        AND E.��Ŀ���=D.��Ŀ���" & vbNewLine & _
        "                        AND E.��Ŀ���=F.COLUMN_VALUE" & vbNewLine & _
        "                        AND (NVL(D.��¼��,1)<>2 Or (NVL(D.��¼��,1)=2 and D.��Ŀ���=3))" & _
        "                        AND A.����ʱ�� BETWEEN [5] And [6] And C.��ֹ�汾 Is Null" & vbNewLine & _
        "                    Order By A.����ʱ��,DECODE(D.��Ŀ���,-1,1,0),DECODE(D.��Ŀ���,-1,1,C.��¼���),D.��¼��"
    
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
            
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, _
             CDate(mDTime), CDate(strEndTime), mstrCurveItem)
    End If
    '1--������������
    '--------------------------------------------------------------------------------------
    With rsTmp
        Do While Not .EOF
            lng��Ŀ��� = zlCommFun.Nvl(!��Ŀ���)
            Select Case lng��Ŀ���
                Case gint����
                    int��� = 1
                Case Else
                    int��� = Val(Nvl(!��¼���))
            End Select
            lngCol = GetTimeCOL(Format(zlCommFun.Nvl(!ʱ��), "HH:mm:ss"))
            blnAllow = False: blnAdd = False: int��ʾ = 0
            '���ʺ���������ʱ�����������Ӧ��ʱ���Ƿ��������
            If mint����Ӧ�� = 2 And lng��Ŀ��� = -1 Then
                mrsPoint.Filter = "��Ŀ���=2 and ʱ��='" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "'"
                If mrsPoint.RecordCount > 0 Then
                    strParam = "���|" & mrsPoint("���")
                    strFidlds = "��ֵ|ID"
                    
                    '��������ʱ����δδ��˵��������Ϊδ��˵��ʱ����ʾδ��˵��
                    If UBound(Split(mrsPoint("��ֵ"), "/")) <> -1 Then
                        If IsNumeric(zlCommFun.Nvl(!��ֵ)) Then
                            If mbln����������ʾ Then
                                gstrValues = zlCommFun.Nvl(!��ֵ) & "/" & Split(mrsPoint("��ֵ"), "/")(0) & "|" & Val(zlCommFun.Nvl(!���))
                            Else
                                gstrValues = Split(mrsPoint("��ֵ"), "/")(0) & "/" & zlCommFun.Nvl(!��ֵ) & "|" & Val(zlCommFun.Nvl(!���))
                            End If
                            
                        Else
                            gstrValues = zlCommFun.Nvl(!��ֵ) & "|" & Val(zlCommFun.Nvl(!���))
                        End If
                    Else
                        gstrValues = mrsPoint("��ֵ") & "|" & Val(zlCommFun.Nvl(!���))
                    End If
                        
                    Call Record_Update(mrsPoint, strFidlds, gstrValues, strParam)
                    blnAllow = True
                Else
                    lng��Ŀ��� = 2
                End If
            End If
            
            '������������
            If (lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ��) And zlCommFun.Nvl(!��¼���) = 1 Then
                mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ʱ��='" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "' and ���<>1"
                If mrsPoint.RecordCount > 0 Then
                    strParam = "���|" & mrsPoint("���")
                    strFidlds = "��ֵ|ID"
                    gstrValues = Split(mrsPoint("��ֵ"), "/")(0) & "/" & zlCommFun.Nvl(!��ֵ) & "|" & Val(zlCommFun.Nvl(!���))
                    Call Record_Update(mrsPoint, strFidlds, gstrValues, strParam)
                End If
                blnAllow = True
            End If
            
            If blnAllow = False Then
                '����������ʾ����
                mrsPoint.Filter = "��Ŀ���=" & lng��Ŀ��� & " and �к�=" & lngCol & " and ��ʾ=1"
                If mrsPoint.RecordCount > 0 Then
                    If Val(zlCommFun.Nvl(!��ʾ)) = 1 And Val(mrsPoint!��ע) <> 1 Then
                        blnAllow = True
                    ElseIf (Val(zlCommFun.Nvl(!��ʾ)) = 1 And Val(mrsPoint!��ע) = 1) Or (Val(zlCommFun.Nvl(!��ʾ)) <> 1 And Val(mrsPoint!��ע) <> 1) Then
                        blnAllow = CheckShow(mrsPoint("ʱ��"), Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss"), lngCol)
                    Else
                        blnAllow = False
                    End If
                    
                    int��ʾ = IIf(blnAllow = True, 1, 0)
                    int��ע = Val(zlCommFun.Nvl(!��ʾ, 0))
                    
                    If blnAllow = True Then
                        Call Record_Update(mrsPoint, "��ʾ", "0", "���|" & mrsPoint!���)
                    End If
                Else
                    int��ʾ = 1
                    int��ע = Val(zlCommFun.Nvl(!��ʾ, 0))
                End If
                
                strPart = GetPart(lng��Ŀ���)
                
                gstrValues = zlCommFun.Nvl(!���) & "|" & zlCommFun.Nvl(!��ֵ, zlCommFun.Nvl(!δ��˵��, "�ܲ�")) & "|" & _
                    zlCommFun.Nvl(!���²�λ, strPart) & "|" & int��� & "|" & _
                    Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & lng��Ŀ��� & "|" & zlCommFun.Nvl(!��¼��) & "|" & Val(zlCommFun.Nvl(!���Ժϸ�)) & "|" & _
                    Val(zlCommFun.Nvl(!������Դ, 0)) & "|" & int��ʾ & "|" & int��ע & "|0|" & vfgThis.TextMatrix(0, vfgThis.FixedCols + lngCol - 1) & "|" & lngCol & "|0"
         
                Call Record_Add(mrsPoint, gstrFields, gstrValues)
            End If
        .MoveNext
        Loop
    End With
    
    '����������Ϣ
    Set mrsCopy = CopyNewRs(mrsPoint)
        
    'չʾ������Ϣ
    Call ShowData
    mblnInit = True
    zlRefreshData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CopyNewRs(ByVal rsData As ADODB.Recordset) As ADODB.Recordset
'-------------------------------------------------
'����:�����µļ�¼����Ϣ
'-------------------------------------------------
    Dim i As Integer
    Dim rsNew As New ADODB.Recordset
    On Error GoTo Errhand
    
    rsData.Filter = 0

    With rsNew
        '�����ֶ�
        For i = 0 To rsData.Fields.Count - 1
            .Fields.Append rsData.Fields(i).Name, rsData.Fields(i).Type, rsData.Fields(i).DefinedSize, adFldIsNullable
        Next i
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        '����������Ϣ
        rsData.Filter = 0
        Do While Not rsData.EOF
            .AddNew
            For i = 0 To rsData.Fields.Count - 1
                .Fields(rsData.Fields(i).Name).Value = rsData.Fields(i).Value
            Next i
            .Update
        rsData.MoveNext
        Loop
    End With
    
    rsNew.Filter = 0
    
    Set CopyNewRs = rsNew
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ShowData()
'---------------------------------------------------
'����:չʾ������Ϣ
'---------------------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim strPart As String

    '����Ƿ������ʾΪ2�ļ�¼
    For intRow = vfgThis.FixedRows To vfgThis.Rows - 1
        For intCOl = vfgThis.FixedCols To vfgThis.Cols - 1
            mrsPoint.Filter = 0
            mrsPoint.Filter = "��Ŀ���=" & Val(vfgThis.TextMatrix(intRow, 0)) & " and ��ע=2 and �к�=" & (intCOl - vfgThis.FixedCols + 1)
            If mrsPoint.RecordCount > 0 Then
                '������ʾΪ2�ļ�¼
                Do While Not mrsPoint.EOF
                    mrsPoint!��ʾ = 2
                    mrsPoint.Update
                mrsPoint.MoveNext
                Loop
                '������ʾ��Ϊ2�ļ�¼
                mrsPoint.Filter = "��Ŀ���=" & Val(vfgThis.TextMatrix(intRow, 0)) & " and ��ע<>2 and �к�=" & (intCOl - vfgThis.FixedCols + 1)
                Do While Not mrsPoint.EOF
                    mrsPoint!��ʾ = 0
                    mrsPoint.Update
                mrsPoint.MoveNext
                Loop
            End If
        Next intCOl
    Next intRow
    
    mrsPoint.Filter = 0
    '��ʾ��������
    mrsPoint.Filter = "��ʾ=1"
    mrsPoint.Sort = "���,ʱ��"
    With mrsPoint
        Do While Not .EOF
            For intRow = vfgThis.FixedRows To vfgThis.Rows - 1
                If Val(vfgThis.TextMatrix(intRow, 0)) = !��Ŀ��� Then
                    strPart = GetPart(!��Ŀ���)
                    If Nvl(!��λ) = "" Then
                        vfgThis.TextMatrix(intRow, vfgThis.FixedCols + Val(!�к�) - 1) = !��ֵ
                    Else
                        vfgThis.TextMatrix(intRow, vfgThis.FixedCols + Val(!�к�) - 1) = IIf(Trim(strPart) <> Trim(!��λ), Trim(!��λ) & ":" & !��ֵ, !��ֵ)
                    End If
                End If
            Next intRow
        .MoveNext
        Loop
    End With
    mblnInit = True
    Call vfgThis.Select(vfgThis.Row, vfgThis.Col)
    Call vfgThis_AfterRowColChange(vfgThis.Row, vfgThis.Col, vfgThis.Row, vfgThis.Col)
    mblnInit = False
End Sub

Private Function SaveData() As Boolean
'------------------------------------------------
'����:����������Ϣ
'------------------------------------------------
    Dim blnTran As Boolean
    Dim lngID As Long
    Dim strSQL As String
    Dim arrSQL() As String
    Dim i As Integer, lngItemCode As Long
    
    On Error GoTo Errhand
    Screen.MousePointer = 11
    
    ReDim Preserve arrSQL(1 To 1)
    
    With mrsPoint
        .Filter = 0
        Do While Not .EOF
            If Val(!״̬) = 2 Then
                lngID = Val(!���)
                lngItemCode = Val(!��Ŀ���)
                
                If InStr(1, !��ֵ, "/") = 0 Then
                    strSQL = "ZL_���µ�����_������ʾ("
                    strSQL = strSQL & lngID & ","
                    strSQL = strSQL & Val(!��ʾ) & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                Else
                    lngID = Val(!���)
                    
                    strSQL = "ZL_���µ�����_������ʾ("
                    strSQL = strSQL & lngID & ","
                    strSQL = strSQL & Val(!��ʾ) & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                    
                    If InStr(1, "," & gint��ʹǿ�� & "," & gint���� & "," & gint���� & ",", "," & lngItemCode & ",") <> 0 Then
                        lngID = Val(!Id)
                        
                        strSQL = "ZL_���µ�����_������ʾ("
                        strSQL = strSQL & lngID & ","
                        strSQL = strSQL & Val(!��ʾ) & ")"
                        
                        arrSQL(ReDimArray(arrSQL)) = strSQL
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    'ѭ��ִ��SQL��������
    'Debug.Print "----���濪ʼ:" & Now
    gcnOracle.BeginTrans
    blnTran = True
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������������"): ' Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    blnTran = False
    'Debug.Print "----�������:" & Now
    
    '�޸�״̬=0
    mrsPoint.Filter = 0
    Do While Not mrsPoint.EOF
        mrsPoint!״̬ = 0
        mrsPoint.Update
        mrsPoint.MoveNext
    Loop
    
    mblnChage = False
    mblnOK = True
    Screen.MousePointer = 0
    SaveData = True
    Exit Function
Errhand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
End Function

Private Function GetPart(ByVal lng��Ŀ���) As String
'����:��ȡĬ�ϵ����²�λ
    Dim strPart As String
    mrs��λ.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ȱʡ��=1"
    If mrs��λ.RecordCount > 0 Then strPart = zlCommFun.Nvl(mrs��λ("��λ"))
    GetPart = strPart
End Function

Private Function CheckShow(ByVal strBegin As String, ByVal strEnd As String, ByVal lngCol As Long) As Boolean
'-------------------------------------------------
'���ܣ��Ա�����ʱ����Ǹ��������յ�ʱ��
'strbegin �Աȵ�ʱ��  strend��ǰʱ��   lngcol-1=ʱ�䷶Χ���������
'--------------------------------------------------
    Dim strTime As String
    Dim blnAllow As Boolean
    
    If (lngCol - 1) <= UBound(marrTime) Then
        If gintHourBegin + (lngCol - 1) * T_BodyStyle.lngʱ���� = 24 Then
            strTime = Format(Format(mDTime, "YYYY-MM-DD") & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        Else
            strTime = Format(Format(mDTime, "YYYY-MM-DD") & " " & gintHourBegin + (lngCol - 1) * T_BodyStyle.lngʱ���� & ":00:00", "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    If Abs(DateDiff("s", CDate(Format(strBegin, "YYYY-MM-DD HH:mm:ss")), CDate(strTime))) > Abs(DateDiff("s", CDate(Format(strEnd, "YYYY-MM-DD HH:mm:ss")), CDate(strTime))) Then
        blnAllow = True
    Else
        blnAllow = False
    End If
    
    CheckShow = blnAllow
End Function

Private Function GetTimeCOL(ByVal strTime As String) As Integer
'--------------------------------------------------
'���ݴ����ʱ������ʱ�������Ƕ�ʱ��
'-------------------------------------------------
    Dim i As Integer
    Dim strValue As String
    
    strValue = Format(strTime, "HH:mm")
    For i = 0 To UBound(marrTime) - 1
        If strValue >= Format(Split(marrTime(i), ",")(0), "HH:mm") And strValue <= Format(Split(marrTime(i), ",")(1), "HH:mm") Then
            Exit For
        End If
    Next i
    
    GetTimeCOL = i + 1
End Function

Private Function InitTime() As String
'--------------------------------------------------------
'����:��ȡһ���ʱ�����Ϣ
'--------------------------------------------------------
    Dim i As Integer
    Dim strName As String, strTime As String
    
    Call InitDateTimeRange(marrTime, gintHourBegin, T_BodyStyle.lng������, T_BodyStyle.lngʱ����)
    For i = 0 To UBound(marrTime) - 1
        strName = strName & ";" & Format(Split(marrTime(i), ",")(0), "HH:mm") & "-" & Format(Split(marrTime(i), ",")(1), "HH:mm")
    Next i
    If Left(strName, 1) = ";" Then strName = Mid(strName, 2)
    strName = "��Ŀ\ʱ�䷶Χ" & ";" & strName
    
    For i = 0 To T_BodyStyle.lng������ - 1
        strTime = strTime & ";" & gintHourBegin + i * T_BodyStyle.lngʱ����
    Next i
    If Left(strTime, 1) = ";" Then strTime = Mid(strTime, 2)
    strTime = "��Ŀ\ʱ�䷶Χ" & ";" & strTime
    
    InitTime = strTime & "[LPF]" & strName
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strParam As String
    Dim intCOl As Integer
    Select Case Control.Id
    
        Case conMenu_Edit_Save '����
            If Not SaveData Then Exit Sub
            Set mrsCopy = CopyNewRs(mrsPoint)
            'չʾ������Ϣ
            Call ShowData
            mblnInit = True
        Case conMenu_Edit_Reuse 'ȡ��
            '����������Ϣ
            Set mrsPoint = CopyNewRs(mrsCopy)
            'չʾ������Ϣ
            Call ShowData
            mblnInit = True
            mblnChage = False
        Case conMenu_Edit_Curve, conMenu_Edit_CurveTable '���ü�¼
             If mblnChage Then
                If MsgBox("�����Ѿ������ı�,�����Ƿ���Ҫ����?", vbInformation + vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                    If Not SaveData Then Exit Sub
                End If
            End If
            intCOl = GetTimeCOL(Format(mDTime, "YYYY-MM-DD HH:mm:ss")) - 1
            If intCOl < 0 Then intCOl = 0
            strParam = Format(Format(mDTime, "YYYY-MM-DD") & " " & Split(marrTime(intCOl), ",")(0), "YYYY-MM-DD HH:mm:ss") & ";" & _
                Format(Format(mDTime, "YYYY-MM-DD") & " " & Split(marrTime(intCOl), ",")(1), "YYYY-MM-DD HH:mm:ss")
            '������ʾ�༭����
            Call gobjTendEditor.BodyEditCur(IIf(Control.Id = conMenu_Edit_Curve, 0, -1), strParam)
            Call GetTableRowName
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    
    Bottom = stbThis.Height
    With picStb
        .Top = stbThis.Top + 50
        .Left = stbThis.Panels(2).Left + 50
        .Height = stbThis.Height - 50
        .Width = stbThis.Panels(2).Width - 50
    End With
    
    With lblStb
        .Font.Size = 9 + 9 * mintBigSize / 3
        .Height = TextHeight("��")
        .Top = (picStb.Height - .Height) \ 2
        .Left = 10
    End With
End Sub

Private Sub cbsMain_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With picThis
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim frmMain As Form
    Dim blnEnable As Boolean
    
     Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Reuse
             Control.Enabled = IIf(mblnChage = True, True, False)
        Case conMenu_Edit_Curve, conMenu_Edit_CurveTable
            blnEnable = True
            For Each frmMain In Forms
                If frmMain.Name = "frmCaseTendBodySetData" Then
                    blnEnable = False
                End If
            Next
            Control.Enabled = blnEnable
    End Select
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж�����ж���
    mbln��Ժ = False
    If Not (mrsPoint Is Nothing) Then Set mrsPoint = Nothing
    If Not (mrs��λ Is Nothing) Then Set mrs��λ = Nothing
    If Not (mcbrToolBar Is Nothing) Then Set mcbrToolBar = Nothing
    If Not (mrsCopy Is Nothing) Then Set mrsCopy = Nothing
    mblnChage = False
     '���洰��
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picShow_Paint()
    picShow.BackColor = &H8000000F
End Sub

Private Sub picShow_Resize()
    lblTmp.Top = 0
    lblTmp.Left = 0
    With vfgShow
        .Top = lblTmp.Height
        .Left = 0
        .Width = picShow.Width
        .Height = picShow.Height - lblTmp.Height - lblTmp.Top
    End With
End Sub

Private Sub picThis_Paint()
    picThis.BackColor = &H8000000F
End Sub

Private Sub picThis_Resize()
    With lblTime
        .Left = 10
        .Top = 10
        .Caption = "ʱ��:" & Format(mDTime, "YYYY-MM-DD")
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With vfgThis
        .Left = 5
        .Top = lblTime.Top + lblTime.Height + 20
        .Width = picThis.Width
        .Height = (picThis.Height - .Top) * 0.65
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With picShow
        .Left = vfgThis.Left
        .Top = vfgThis.Height + vfgThis.Top + 50
        .Width = vfgThis.Width
        .Height = picThis.Height - picShow.Top
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With lblTmp
        .Top = 10
        .Left = 10
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With vfgShow
        .Left = 5
        .Top = lblTmp.Top + lblTmp.Height + 20
        .Width = picShow.Width
        .Height = picShow.Height - .Top
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    picShow.Visible = True
    lblTmp.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub vfgShow_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vfgShow
        If .Col >= .FixedCols Then
            If NewRow = .Rows - 1 Then
                .FocusRect = flexFocusHeavy
            Else
                .FocusRect = flexFocusLight
            End If
        End If
    End With
End Sub

Private Sub vfgShow_Click()
    vfgShow.ColSel = vfgShow.Col
End Sub

Private Sub vfgShow_DblClick()
    Dim intSate As Integer, int��ʾ As Integer
    Dim intCOl As Integer, intRow As Integer
    Dim intColSel As Integer
    Dim arrValue() As String
    Dim strPart As String
    Dim lngItemNO As Long
    
    If mblnInit = False Then Exit Sub
    If mblnFileBack = True Then Exit Sub
    
    With vfgShow
        If .Rows - 1 = .Row And .Col >= .FixedCols Then
            '����������Ŀ
            If .TextMatrix(.Row, .Col) = "��" Then
                
                mrsPoint.Filter = 0
                mrsPoint.Filter = "���=" & Val(.ColData(.Col))
                intSate = Val(mrsPoint!״̬)
                intCOl = Val(mrsPoint!�к�)
                lngItemNO = Val(mrsPoint!��Ŀ���)
                int��ʾ = 2
                intSate = 2
                mrsPoint!��ʾ = int��ʾ
                mrsPoint!״̬ = intSate
                mrsPoint!��ע = int��ʾ
                mrsPoint.Update
                .TextMatrix(.Row, .Col) = ""
                mrsPoint.Filter = "��Ŀ���=" & lngItemNO & " And �к�=" & intCOl & " And ���<>" & Val(.ColData(.Col))
                Do While Not mrsPoint.EOF
                    mrsPoint!��ʾ = 0
                    mrsPoint!��ע = 0
                    mrsPoint.Update
                mrsPoint.MoveNext
                Loop
            Else
                '������¼����Ϣ
                For intCOl = .FixedCols To .Cols - 1
                    If .TextMatrix(.Row, intCOl) = "��" Then
                        mrsPoint.Filter = 0
                        mrsPoint.Filter = "���=" & Val(.ColData(intCOl))
                        intSate = Val(mrsPoint!״̬)
                        int��ʾ = 0
                        Select Case intSate
                            Case 0
                                intSate = 2
                            Case 2
                                intSate = 0
                        End Select
                        mrsPoint!��ʾ = int��ʾ
                        mrsPoint!״̬ = intSate
                        mrsPoint!��ע = int��ʾ
                        mrsPoint.Update
                        .TextMatrix(.Row, intCOl) = ""
                    End If
                Next intCOl
                .TextMatrix(.Row, .Col) = "��"
                mrsPoint.Filter = 0
                mrsPoint.Filter = "���=" & Val(.ColData(.Col))
                intCOl = Val(mrsPoint!�к�)
                lngItemNO = Val(mrsPoint!��Ŀ���)
                intSate = Val(mrsPoint!״̬)
                int��ʾ = 1
                Select Case intSate
                    Case 0
                        intSate = 2
                    Case 2
                        intSate = 0
                End Select
                mrsPoint!��ʾ = int��ʾ
                mrsPoint!״̬ = intSate
                mrsPoint!��ע = int��ʾ
                mrsPoint.Update
                
                mrsPoint.Filter = "��Ŀ���=" & lngItemNO & " And �к�=" & intCOl & " And ��ʾ=2"
                Do While Not mrsPoint.EOF
                    intSate = Val(mrsPoint!״̬)
                    int��ʾ = 0
                    intSate = 2
                    mrsPoint!��ʾ = int��ʾ
                    mrsPoint!״̬ = intSate
                    mrsPoint!��ע = int��ʾ
                    mrsPoint.Update
                mrsPoint.MoveNext
                Loop
            End If
            vfgThis.Cell(flexcpText, vfgThis.FixedRows, vfgThis.FixedCols, vfgThis.Rows - 1, vfgThis.Cols - 1) = ""
            '��ʾ����
            mrsPoint.Filter = "��ʾ=1"
            mrsPoint.Sort = "���,ʱ��"
            Do While Not mrsPoint.EOF
                For intRow = vfgThis.FixedRows To vfgThis.Rows - 1
                    If Val(vfgThis.TextMatrix(intRow, 0)) = Val(mrsPoint!��Ŀ���) Then
                        strPart = GetPart(mrsPoint!��Ŀ���)
                        If Trim(mrsPoint!��λ) = "" Then
                            vfgThis.TextMatrix(intRow, vfgThis.FixedCols + Val(mrsPoint!�к�) - 1) = mrsPoint!��ֵ
                        Else
                            vfgThis.TextMatrix(intRow, vfgThis.FixedCols + Val(mrsPoint!�к�) - 1) = IIf(Trim(strPart) <> Trim(mrsPoint!��λ), Trim(mrsPoint!��λ) & ":" & mrsPoint!��ֵ, mrsPoint!��ֵ)
                        End If
                    End If
                Next intRow
            mrsPoint.MoveNext
            Loop
            mblnChage = True
        End If
    End With
End Sub

Private Sub vfgThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intCOl As Integer, intRow As Integer, i As Integer
    Dim strFind As String, strValue As String, strInfo As String
    intCOl = NewCol
    intRow = NewRow
    If mblnInit = False Then Exit Sub
    lblTmp.Caption = ""
    With vfgShow
        If NewRow >= vfgThis.FixedRows And NewCol >= vfgThis.FixedCols Then
            mintColSel = intCOl
            If Val(vfgThis.TextMatrix(intRow, 0)) = gint���� Then  '������Ŀ
                .Rows = 4
                .TextMatrix(0, 0) = "ʱ��"
                .TextMatrix(1, 0) = "��ֵ"
                .TextMatrix(2, 0) = "����"
                .TextMatrix(3, 0) = "��ʾ"
                strFind = " and �к�=" & intCOl - vfgThis.FixedCols + 1
            Else
                .Rows = 3
                .TextMatrix(0, 0) = "ʱ��"
                .TextMatrix(1, 0) = "��ֵ"
                .TextMatrix(2, 0) = "��ʾ"
                strFind = " and �к�=" & intCOl - vfgThis.FixedCols + 1
             End If
             lblTmp.Caption = vfgThis.TextMatrix(0, intCOl) & "֮����ڵ�" & vfgThis.TextMatrix(intRow, 2) & "������:"
        
             picShow.Visible = True
             mrsPoint.Filter = "��Ŀ���=" & Val(vfgThis.TextMatrix(intRow, 0)) & strFind
             mrsPoint.Sort = "ʱ��,���"
             
             .Cols = mrsPoint.RecordCount + .FixedCols
             i = .FixedCols
             Do While Not mrsPoint.EOF
                .ColWidth(-1) = 1200 + 1200 * mintBigSize / 3
                 vfgShow.TextMatrix(0, i) = Format(mrsPoint!ʱ��, "HH:mm")
                 vfgShow.TextMatrix(1, i) = mrsPoint!��ֵ
                 If Val(vfgThis.TextMatrix(intRow, 0)) = gint���� Then
                     vfgShow.TextMatrix(2, i) = IIf(mrsPoint!���� = 1, "��", "")
                     vfgShow.TextMatrix(3, i) = IIf(mrsPoint!��ʾ = 1, "��", "")
                 Else
                     vfgShow.TextMatrix(2, i) = IIf(mrsPoint!��ʾ = 1, "��", "")
                 End If
                 vfgShow.ColData(i) = Val(mrsPoint!���)
                 i = i + 1
             mrsPoint.MoveNext
             Loop
            .RowHeight(-1) = 300 + 300 * mintBigSize / 3
             .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
             .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = &H0&
             vfgThis.Cell(flexcpBackColor, vfgThis.FixedRows, vfgThis.FixedCols, vfgThis.Rows - 1, vfgThis.Cols - 1) = &H80000005
             vfgThis.Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = &H80000018
        End If
    End With
End Sub