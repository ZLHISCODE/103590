VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmDepositNormal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picDeposit 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1815
      ScaleWidth      =   4335
      TabIndex        =   3
      Top             =   240
      Width           =   4335
      Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
         Height          =   1305
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   2055
         _cx             =   3625
         _cy             =   2302
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
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
   End
   Begin VB.PictureBox picBalanceInfo 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   5520
      ScaleHeight     =   2055
      ScaleWidth      =   2295
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
      Begin VSFlex8Ctl.VSFlexGrid vsBalanceInfor 
         Height          =   1305
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2055
         _cx             =   3625
         _cy             =   2302
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
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
   End
   Begin VB.PictureBox picInvoice 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   2880
      ScaleHeight     =   2055
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
      Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
         Height          =   1305
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2055
         _cx             =   3625
         _cy             =   2302
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
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
   End
   Begin VB.PictureBox picEinvoice 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   240
      ScaleHeight     =   2055
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   2400
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsEInvoice 
         Height          =   1305
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2055
         _cx             =   3625
         _cy             =   2302
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
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
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   8760
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDepositNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mblnDateMoved As Boolean, mblnNOMoved As Boolean
Private mfrmMain As Object, mlngModule As Long
Private mstrPrivs As String, mstrEInvoicePrivs As String
Private mobjEInvoice As clsEinvoice  '����Ʊ�ݲ���
Private mblnNotRefresh As Boolean  '��ˢ������
Private mrsList As ADODB.Recordset  '�����б�
Private mlngTopRow As Long, mlngCurRow As Long '��¼�ϴ�ѡ�����
Private mint��� As Integer '0-��ѯ����;1-��ѯ����Ԥ��;2-��ѯסԺԤ��;3-��ѯѺ���¼
Public mblnGo As Boolean
Private mlngGo As Long
Private mcllFilter As Collection '������������
Public Event SelectDeposit(ByVal blnUse As Boolean, ByVal blnѺ�� As Boolean, ByVal lngԤ��ID As Long, ByVal int��¼״̬ As Integer, _
                                         ByVal bln����Ʊ�� As Boolean, ByVal lng����Ʊ��ID As Long, ByVal bln���� As Boolean, ByVal strƱ�ݺ� As String, _
                                         ByVal bln��Ԥ�� As Boolean, ByVal int���ӱ�־ As Integer, ByVal lngԭʼID As Long)  'ѡ��Ԥ���б�ĳһ��
Public Event ShowStatus(ByVal strMessage As String)     '��ʾ״̬��
Public Event FilterDeposit()                                             '�������˴���
Public Event PopupMenu()                                             '�����˵�
Public Event ViewGo()                                                    '��λ
Public Event MoneyEnum()                                             '�ֽ�㳮
Public Event RollingCurtain()                                           '�շ�����
Public Event FileLocalSet()                                               '������������
Public Event FilePrint()                                                    '��ӡ
Public Event EditDeposit()                                               '��Ԥ��
Public Event EidtBalanceDel()                                          '����˿�
Public Event ReadPati(ByVal strName As String)              '������id���˺���ز�������

Public Sub zlInit(ByVal frmMain As Object, ByVal objEInvoice As clsEinvoice, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strEInvoicePrivs As String, _
                         ByVal int��� As Integer, ByVal cllFilter As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���
    '���:objEInvoice-����Ʊ�ݴ����
    '     strPrivs-Ԥ��Ȩ�޴�
    '     strEInvoicePrivs-����Ʊ�ݲ���Ȩ�޴�
    '     int���-:0-����;1-����Ԥ��;2-סԺԤ��;3-Ѻ��
    '����:
    '����:����
    '����:2020-06-29 17:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mint��� = int���
    Set mcllFilter = cllFilter
    mstrPrivs = strPrivs: mstrEInvoicePrivs = strEInvoicePrivs
    Set mobjEInvoice = objEInvoice: mlngModule = lngModule
    Call InitDepositGrid
    Call InitInvoiceGrid
    Call InitEinvoiceGrid
    Call InitBalanceGrid
    Call InitdkpMain
End Sub

Public Function zlRefrshListData(ByVal cllFilter As Collection, ByVal blnDateMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '���:cllFilter-��������
    '     ��ʽ:array(��������,ֵ1,ֵ2,..),����
    '����:�ɹ�����true,���򷵻�False
    '����:2012-06-12 14:43:06
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mcllFilter = cllFilter
    mblnDateMoved = blnDateMoved
    On Error GoTo errHandle
    Call ShowBills(cllFilter)
    zlCommFun.StopFlash
    zlRefrshListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowBills(ByVal cllFilter As Collection)
    '����:��������ȡ�����б�(���˹���)
    '��Σ�cllFilter-��������
    Dim strWhere As String, strSQL As String, strSQLYJ As String, strYJwhere As String
    Dim lng����ID As Long, lng��ҳID As Long, lng����� As Long, lngסԺ�� As Long
    Dim str���� As String, str�տ��� As String
    Dim str��ʼ���� As String, str��������   As String, str������� As String, strѺ����� As String
    Dim str��ʼʱ�� As String, str����ʱ�� As String, str��ʼƱ�� As String, str����Ʊ��  As String
    Dim i As Integer, intCount As Integer, int���� As Integer
    Dim varData As Variant, dbl��� As Double, blnOnlyѺ�� As Boolean
    Dim strTempSQL As String, strTmp As String, strNO As String, strNoTmp As String

    On Error GoTo errH
    
    strWhere = ""
    'cllFilter-��������
    For i = 1 To cllFilter.Count
        varData = cllFilter(i)
        Select Case varData(0)
        Case "����ID"
            lng����ID = Val(varData(1))
            If lng����ID <> 0 Then strWhere = strWhere & " And A.����ID=[1]"
        Case "��ҳID"
            lng��ҳID = Val(varData(1))
            If lng��ҳID <> 0 Then strWhere = strWhere & " And A.��ҳid=[2]"
        Case "�����"
            lng����� = Val(varData(1))
            If lng����� <> 0 Then strWhere = strWhere & " And b.�����=[3]"
        Case "סԺ��"
            lngסԺ�� = Val(varData(1))
            If lngסԺ�� <> 0 Then strWhere = strWhere & " And b.����ID In (Select ����ID From ������ҳ where סԺ��=[4])"
        Case "��ʼ����"
            str��ʼ���� = Trim(varData(1))
        Case "��������"
            str�������� = Trim(varData(1))
        Case "����"
            str���� = Trim(varData(1))
            If str���� <> "" Then strWhere = strWhere & " And Upper(A.����) like [11]"
        Case "�������"
            str������� = Trim(varData(1))
            If str������� <> "" Then strWhere = strWhere & " And A.str�������  = [12]"
        Case "�տ���"
            str�տ��� = Trim(varData(1))
            If str�տ��� <> "" Then strWhere = strWhere & " And A.����Ա����  = [13]"
        Case "��ʼʱ��"
            str��ʼʱ�� = Trim(varData(1))
        Case "����ʱ��"
            str����ʱ�� = Trim(varData(1))
        Case "��ʼƱ��"
            str��ʼƱ�� = Trim(varData(1))
        Case "����Ʊ��"
            str����Ʊ�� = Trim(varData(1))
        Case "Ѻ�����"
            strѺ����� = Trim(varData(1))
            If strѺ����� <> "" Then
                blnOnlyѺ�� = True
                strYJwhere = strYJwhere & " And A.Ѻ�����=[14]"
            End If
        Case "��ѯ��¼"
            '0-����;1-������Ԥ��;2-���˿��ԭԤ����¼
             int���� = Val(varData(1))
             If int���� = 1 Then
                strWhere = strWhere & " and ��¼״̬=1"
             ElseIf int���� = 2 Then
                strWhere = strWhere & " And ��¼״̬<>1"
             End If
        End Select
    Next
    
    If str��ʼʱ�� <> "" Then strWhere = strWhere & " And A.�տ�ʱ��  between  [5] and [6]"
                
    If str��ʼ���� <> "" And str�������� <> "" Then
         strWhere = strWhere & " And a.NO between [7]  and [8]"
    ElseIf str��ʼ���� <> "" Or str�������� <> "" Then
         strWhere = strWhere & IIf(str��ʼ���� <> "", " And a.NO=[7]", " And a.NO=[8]")
    End If
    
    strSQL = ""
    If str��ʼƱ�� <> "" And str����Ʊ�� <> "" Then
        strSQL = "Select A.NO  From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B Where A.��������=2 And A.ID=B.��ӡID And B.Ʊ��=2 And B.����=1   And B.����  between [9] and [10] "
    ElseIf str��ʼƱ�� <> "" Or str����Ʊ�� <> "" Then
        strSQL = "Select A.NO  From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B Where A.��������=2 And A.ID=B.��ӡID And B.Ʊ��=2 And B.����=1   And B.����" & IIf(str��ʼƱ�� <> "", " = [9] ", " =[10]")
    End If
    If strSQL <> "" Then strWhere = strWhere & " And A.NO in (" & strSQL & ")"
            
    If strWhere = "" Then
        mblnNotRefresh = True
        vsDeposit.Clear 1: vsDeposit.Rows = 2:
        Call RefrshDataDetial
        mblnNotRefresh = False
        Exit Sub
    End If
    strWhere = strWhere & " And Nvl(A.У�Ա�־,0) =0  "
    
    strSQL = ""
    strSQL = _
        "   Select a.���ӱ�־, a.Id As ԭʼid,'' As �˿�NO,a.id,A.NO as NO,A.ʵ��Ʊ�� as Ʊ�ݺ�,A.����Ա���� as ����Ա," & _
        "           To_Char(A.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
        "           A.����ID,A.�����,A.סԺ��,A.����,A.�Ա�,A.����,D.���� as ����," & _
        "           To_Char(Sum(A.���),'9999999990.00') as ���," & _
        "           A.���㷽ʽ,A.�������,A.ժҪ,A.��¼״̬,A.���ʽ����, " & _
        "           Decode(nvl(A.Ԥ�����,2),1,'����Ԥ��', 'סԺԤ��') as Ԥ�����, nvl(A.Ԥ�����,0) as Ԥ�����ID, " & _
        "           NULL as Ѻ������,NULL as Ѻ�����,a.������ˮ��,a.����˵��,a.Ԥ������Ʊ��" & _
        " From  ����Ԥ����¼ A,���ű� D " & _
        " Where A.����ID=D.ID(+)  And a.��¼���� = 1 " & strWhere & _
                      IIf(mint��� = 0 Or mint��� = 3, "", "  And   A.Ԥ�����=" & mint���) & _
        " Group by a.id,A.NO,A.��¼״̬,A.ʵ��Ʊ�� ,Nvl(A.Ԥ�����, 0),Decode(nvl(A.Ԥ�����,2),1,'����Ԥ��', 'סԺԤ��'),A.����Ա����," & _
        "           To_Char(A.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.����ID,A.�����,A.סԺ��,A.����,A.����," & _
        "           A.�Ա� , D.����, A.���㷽ʽ,A.�������, A.ժҪ,A.���ʽ����,a.������ˮ��,a.����˵��,a.Ԥ������Ʊ��,a.���ӱ�־ "

     strSQLYJ = _
              "   Select 0 as ���ӱ�־,a.Id As ԭʼid,'' As �˿�NO,a.id,A.NO as NO,A.ʵ��Ʊ�� as Ʊ�ݺ�,A.����Ա���� as ����Ա," & _
              "           To_Char(A.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
              "           A.����ID,A.�����,A.סԺ��,A.����,A.�Ա�,A.����,D.���� as ����," & _
              "           To_Char(Sum(A.���),'9999999990.00') as ���," & _
              "           A.���㷽ʽ,A.�������,A.ժҪ,A.��¼״̬,A.���ʽ����, " & _
              "           NULL as Ԥ�����,NULL as Ԥ�����ID,Decode(nvl(A.�Ƿ�����,0),1,'����Ѻ��', 'סԺѺ��') as Ѻ������,A.Ѻ�����,a.������ˮ��,a.����˵��,0 as Ԥ������Ʊ�� " & _
              " From ����Ѻ���¼  A,���ű� D " & _
              " Where A.����ID=D.ID(+) " & strWhere & strYJwhere & _
              " Group by a.id,A.NO,A.��¼״̬,A.ʵ��Ʊ�� ,A.Ѻ�����,Decode(nvl(A.�Ƿ�����,0),1,'����Ѻ��', 'סԺѺ��'),A.����Ա����," & _
              "           To_Char(A.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.����ID,A.�����,A.סԺ��,A.����,A.����," & _
              "           A.�Ա� , D.����, A.���㷽ʽ,A.�������, A.ժҪ,A.���ʽ����,a.������ˮ��,a.����˵��"
              
    'mint���:0-��ѯ����;1-��ѯ����Ԥ��;2-��ѯסԺԤ��;3-��ѯѺ���¼
    If mint��� <> 3 Then
             strSQL = strSQL & " Union ALL " & _
              "   Select 11 as ���ӱ�־,e.Id As ԭʼID,A.no as �˿�NO,a.id,J.NO as NO,A.ʵ��Ʊ�� as Ʊ�ݺ�,A.����Ա���� as ����Ա," & _
              "           To_Char(J.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
              "           A.����ID,A.�����,A.סԺ��,A.����,A.�Ա�,A.����,D.���� as ����," & _
              "           To_Char(Sum(-1*A.��Ԥ��),'9999999990.00') as ���," & _
              "           A.���㷽ʽ,A.�������,A.ժҪ,A.��¼״̬,A.���ʽ����, " & _
              "           Decode(nvl(A.Ԥ�����,2),1,'����Ԥ��', 'סԺԤ��') as Ԥ�����, nvl(A.Ԥ�����,0) as Ԥ�����ID, " & _
              "           NULL as Ѻ������,NULL as Ѻ�����,a.������ˮ��,a.����˵��,e.Ԥ������Ʊ��" & _
              " From   ( Select Distinct a.����id, a.NO, a.�տ�ʱ��,a.Ԥ������Ʊ�� From  ����Ԥ����¼ A  " & _
              "                Where  a.��¼���� = 1 " & strWhere & "and nvl(���ӱ�־,0)>=1 )  J," & _
               "                   ����Ԥ����¼ A,���ű� D,����Ԥ����¼ E " & _
              " Where  J.����ID=A.����ID And  A.����ID=D.ID(+) And A.��¼����=11 " & _
              "            And Nvl(a.У�Ա�־, 0) = 0 And a.��Ԥ�� > 0  And a.no=e.no And e.��¼����=1 And e.��¼״̬=1 " & _
              " Group by J.no ,e.Id,a.id,A.NO,A.��¼״̬,A.ʵ��Ʊ�� ,Nvl(A.Ԥ�����, 0),Decode(nvl(A.Ԥ�����,2),1,'����Ԥ��', 'סԺԤ��'),A.����Ա����," & _
              "           To_Char(J.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.����ID,A.�����,A.סԺ��,A.סԺ��,A.����,A.����," & _
              "           A.�Ա� , D.����, A.���㷽ʽ,A.�������, A.ժҪ,A.���ʽ����,a.������ˮ��,a.����˵��,e.Ԥ������Ʊ��,a.���ӱ�־ "
    End If
       
    If mint��� = 0 Then
        If blnOnlyѺ�� Then '����ȡѺ��
            strSQL = strSQLYJ
        Else
            strSQL = strSQL & " Union all " & strSQLYJ
        End If
    ElseIf mint��� = 3 Then
        strSQL = strSQLYJ
    End If
    If mblnDateMoved Then
        strTempSQL = Replace(Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼"), "����Ѻ���¼", "H����Ѻ���¼")
        strTempSQL = Replace(Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����"), "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
        strSQL = strSQL & " Union ALL " & vbCrLf & strTempSQL
    End If
      
    strSQL = "Select a.���ӱ�־, a.ԭʼid,a.�˿�no, a.Id, a.NO, a.Ʊ�ݺ�, a.����Ա, a.����ʱ��, a.����id, a.�����, a.סԺ��, a.����, a.�Ա�, a.����, a.����, a.���, a.���㷽ʽ, a.�������," & vbNewLine & _
                 "  a.ժҪ, a.��¼״̬, a.���ʽ����, a.Ԥ�����, a.Ԥ�����id,  a.Ѻ������, a.Ѻ�����, a.������ˮ��, a.����˵��, a.Ԥ������Ʊ�� " & _
                 " From (" & strSQL & ") A Order by a.����ʱ�� desc,a.NO,a.�˿�NO desc"

    Set mrsList = New ADODB.Recordset
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, lng�����, lngסԺ��, CDate(str��ʼʱ��), CDate(str����ʱ��), str��ʼ����, str��������, str��ʼƱ��, str����Ʊ��, _
                                                                    str���� & "%", str�������, str�տ���, strѺ�����)
    mblnNotRefresh = True
    vsDeposit.Clear
    vsDeposit.Rows = 2
    If mrsList.EOF Then
        Call InitDepositGrid
        strTmp = "��ǰ����û�й��˳��κε���"
        RaiseEvent ShowStatus(strTmp)
        RaiseEvent SelectDeposit(False, False, 0, 0, False, 0, False, "", False, 0, 0)
    Else
        RaiseEvent ReadPati(Nvl(mrsList!����))
        Call InitDepositGrid
        vsDeposit.ForeColorSel = vsDeposit.CellForeColor
        mrsList.MoveFirst: dbl��� = 0
        vsDeposit.Rows = mrsList.RecordCount + 1
        With vsDeposit
            .OutlineBar = flexOutlineBarSymbolsLeaf
            .Subtotal flexSTClear
            .MultiTotals = True
            .SubtotalPosition = flexSTAbove
            .OutlineCol = .ColIndex("���ݺ�")
            For i = 1 To mrsList.RecordCount
                .TextMatrix(i, .ColIndex("�˿�NO")) = Nvl(mrsList!�˿�NO)
                .TextMatrix(i, .ColIndex("NO")) = Nvl(mrsList!NO)
                .TextMatrix(i, .ColIndex("ԭʼID")) = Val(Nvl(mrsList!ԭʼID))
                .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(mrsList!ID))
                .TextMatrix(i, .ColIndex("���ݺ�")) = IIf(Nvl(mrsList!�˿�NO) = "", Nvl(mrsList!NO), Nvl(mrsList!�˿�NO))
                .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = Nvl(mrsList!Ʊ�ݺ�)
                .TextMatrix(i, .ColIndex("����Ա")) = Nvl(mrsList!����Ա)
                .TextMatrix(i, .ColIndex("����ʱ��")) = Nvl(mrsList!����ʱ��)
                .TextMatrix(i, .ColIndex("����ID")) = Val(Nvl(mrsList!����ID))
                .TextMatrix(i, .ColIndex("�����")) = Nvl(mrsList!�����)
                .TextMatrix(i, .ColIndex("סԺ��")) = Nvl(mrsList!סԺ��)
                .TextMatrix(i, .ColIndex("����")) = Nvl(mrsList!����)
                .TextMatrix(i, .ColIndex("�Ա�")) = Nvl(mrsList!�Ա�)
                .TextMatrix(i, .ColIndex("����")) = Nvl(mrsList!����)
                .TextMatrix(i, .ColIndex("����")) = Nvl(mrsList!����)
                .TextMatrix(i, .ColIndex("���")) = Nvl(mrsList!���)
                If IsNumeric(.TextMatrix(i, .ColIndex("���"))) Then .TextMatrix(i, .ColIndex("���")) = Format(.TextMatrix(i, .ColIndex("���")), "0.00")
                .TextMatrix(i, .ColIndex("���㷽ʽ")) = Nvl(mrsList!���㷽ʽ)
                .TextMatrix(i, .ColIndex("�������")) = Nvl(mrsList!�������)
                .TextMatrix(i, .ColIndex("ժҪ")) = Nvl(mrsList!ժҪ)
                .TextMatrix(i, .ColIndex("��¼״̬")) = Nvl(mrsList!��¼״̬)
                If .TextMatrix(i, .ColIndex("��¼״̬")) = "2" Then
                    .Cell(flexcpForeColor, i, 0, i, .ColIndex("����Ʊ��")) = &HFF&
                ElseIf .TextMatrix(i, .ColIndex("��¼״̬")) = "3" Then
                    .Cell(flexcpForeColor, i, 0, i, .ColIndex("����Ʊ��")) = &HFF0000
                End If
                .TextMatrix(i, .ColIndex("ҽ�Ƹ��ʽ")) = Nvl(mrsList!���ʽ����)
                .TextMatrix(i, .ColIndex("Ԥ�����")) = Nvl(mrsList!Ԥ�����)
                .TextMatrix(i, .ColIndex("Ԥ�����ID")) = Val(Nvl(mrsList!Ԥ�����ID))
                .TextMatrix(i, .ColIndex("Ѻ������")) = Nvl(mrsList!Ѻ������)
                .TextMatrix(i, .ColIndex("Ѻ�����")) = Nvl(mrsList!Ѻ�����)
                .TextMatrix(i, .ColIndex("������ˮ��")) = Nvl(mrsList!������ˮ��)
                .TextMatrix(i, .ColIndex("����˵��")) = Nvl(mrsList!����˵��)
                .TextMatrix(i, .ColIndex("����Ʊ��")) = IIf(Val(Nvl(mrsList!Ԥ������Ʊ��)) = 1, "��", "")
                .TextMatrix(i, .ColIndex("���ӱ�־")) = Val(Nvl(mrsList!���ӱ�־))
                .IsSubtotal(i) = True
                If Nvl(mrsList!�˿�NO) = "" And Val(Nvl(mrsList!���ӱ�־)) >= 1 Then
                    strNoTmp = Nvl(mrsList!NO)
                    .Cell(flexcpBackColor, i, 0, i, .ColIndex("����Ʊ��")) = &HC0C0FF
                    .RowOutlineLevel(i) = 1
                Else
                    If Nvl(mrsList!NO) = strNoTmp And Val(Nvl(mrsList!���ӱ�־)) = 11 Then
                        .RowOutlineLevel(i) = 2
                        intCount = intCount + 1
                    Else
                        .RowOutlineLevel(i) = 1
                    End If
                End If
                If Val(Nvl(mrsList!���ӱ�־)) <> 11 Then
                    dbl��� = dbl��� + Val(Nvl(mrsList!���))
                End If
                mrsList.MoveNext
            Next
            .Outline 1
        End With
        mrsList.MoveFirst
        strTmp = "�� " & mrsList.RecordCount - intCount & " �ŵ���,�ϼ�:" & Format(dbl���, "0.00")
        RaiseEvent ShowStatus(strTmp)
    End If
    mblnNotRefresh = False
    '�ָ��ϴ���
    If mlngCurRow = 0 Then mlngCurRow = 1
    If mlngTopRow = 0 Then mlngTopRow = 1
    If mlngCurRow <= vsDeposit.Rows - 1 Then
        vsDeposit.Row = mlngCurRow
    Else
        vsDeposit.Row = vsDeposit.Rows - 1
    End If
    If mlngTopRow <= vsDeposit.Rows - 1 Then
        vsDeposit.TopRow = mlngTopRow
    Else
        vsDeposit.TopRow = vsDeposit.Row
    End If
    Call RefrshDataDetial   '������ϸ����
    Me.Refresh
    Exit Sub
errH:
    mblnNotRefresh = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefrshDataDetial()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������
    '����:���˺�
    '����:2020-04-28 10:19:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngԤ��ID As Long, lng����Ʊ��ID As Long, lngԭʼID As Long
    Dim strNO As String, strƱ�ݺ� As String
    Dim int��¼״̬ As Integer, int���ӱ�־ As Integer
    Dim bln�Ƿ�Ѻ�� As Boolean, blnԤ������Ʊ�� As Boolean, bln���� As Boolean, bln��Ԥ�� As Boolean
    
    On Error GoTo errHandle
    
     With vsDeposit
        If Not (.Row <= 0 Or .TextMatrix(.Row, .ColIndex("NO")) = "") Then
            strNO = .TextMatrix(.Row, .ColIndex("NO"))
            lngԤ��ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            bln�Ƿ�Ѻ�� = .TextMatrix(.Row, .ColIndex("Ѻ������")) <> ""
            blnԤ������Ʊ�� = Nvl(.TextMatrix(.Row, .ColIndex("����Ʊ��"))) = "��"
            int��¼״̬ = Val(Nvl(.TextMatrix(.Row, .ColIndex("��¼״̬"))))
            strƱ�ݺ� = .TextMatrix(.Row, .ColIndex("Ʊ�ݺ�"))
            bln��Ԥ�� = Val(.TextMatrix(.Row, .ColIndex("���ӱ�־"))) = 11
            int���ӱ�־ = Val(.TextMatrix(.Row, .ColIndex("���ӱ�־")))
            lngԭʼID = Val(.TextMatrix(.Row, .ColIndex("ԭʼID")))
            mlngGo = .Row: mlngCurRow = .Row: mlngTopRow = .TopRow
        End If
    End With
    If strNO = "" Then Exit Sub
    
    If mblnDateMoved Then
        mblnNOMoved = zlDatabase.NOMoved("����Ԥ����¼", strNO, , "1", Me.Caption)
    Else
        mblnNOMoved = False
    End If
    
    '������ϸ
    Call LoadEInvoiceData(strNO)
    Call LoadInvoiceData(strNO)
    Call LoadBalanceInfor(lngԤ��ID, bln�Ƿ�Ѻ��)

    RaiseEvent SelectDeposit(True, bln�Ƿ�Ѻ��, lngԤ��ID, int��¼״̬, blnԤ������Ʊ��, lng����Ʊ��ID, bln����, strƱ�ݺ�, bln��Ԥ��, int���ӱ�־, lngԭʼID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitDepositGrid()
    Dim strHead As String
    Dim i As Integer
    
    strHead = "�˿�NO,1,0|ԭʼID,1,0|NO,1,0|ID,1,0|���ݺ�,1,1200|Ʊ�ݺ�,4,1050|����Ա,1,850|����ʱ��,4,1850|����ID,1,750|�����,1,750|סԺ��,1,750|����,1,800|�Ա�,4,500|" & _
              "����,4,500|����,1,850|���,7,850|���㷽ʽ,1,850|�������,1,1500|ժҪ,1,1500|��¼״̬,1,0|ҽ�Ƹ��ʽ,1,1500|Ԥ�����,4,800|" & _
              "Ԥ�����ID,1,0|Ѻ������,4,800|Ѻ�����,1,900|������ˮ��,4,1000|����˵��,4,1500|����Ʊ��,4,900|���ӱ�־,1,0"
    With vsDeposit
        mblnNotRefresh = True
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            .ColKey(i) = UCase(Trim(.TextMatrix(0, i)))
        Next
        If Not Visible Then Call RestoreFlexState(vsDeposit, App.ProductName & "\" & Me.Name)
        .ColHidden(.ColIndex("Ԥ�����ID")) = True
        .ColHidden(.ColIndex("�˿�NO")) = True
        .ColHidden(.ColIndex("ԭʼID")) = True
        .ColHidden(.ColIndex("NO")) = True
        .ColHidden(.ColIndex("ID")) = True
        .ColHidden(.ColIndex("���ӱ�־")) = True
        .ColHidden(.ColIndex("��¼״̬")) = True
        If mint��� = 1 Or mint��� = 2 Then .ColHidden(.ColIndex("Ѻ������")) = True: .ColHidden(.ColIndex("Ѻ�����")) = True:
        .RowHeight(0) = 320
        '�ָ��ϴ���
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        .Col = 0: .ColSel = .Cols - 1
        Call vsDeposit_EnterCell
        mblnNotRefresh = False
        zl_vsGrid_Para_Restore mlngModule, vsDeposit, Me.Name, "Ԥ����Ϣ�б�", False
        If .Rows > 1 Then .Row = 1
        .Redraw = True
    End With

End Sub

Private Sub LoadInvoiceData(ByVal strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ��ʹ����ϸ
    '���:strNO-������Ϣ��
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-04-01 18:55:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    On Error GoTo errH
    If strNO = "" Then
        vsInvoice.Redraw = flexRDNone
        vsInvoice.Rows = 2
        vsInvoice.Clear 1
        vsInvoice.Redraw = flexRDBuffered: Exit Sub
    End If
    
    strSQL = _
    " Select b.Id, b.���� As Ʊ�ݺ�," & vbNewLine & _
    " Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����',7,'��Ʊ�ջ�') As ʹ��ԭ��," & vbNewLine & _
    "    To_Char(b.ʹ��ʱ��, 'MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����" & vbNewLine & _
    " From Ʊ�ݴ�ӡ���� A, Ʊ��ʹ����ϸ B" & vbNewLine & _
    " Where a.�������� = 2 And a.Id = b.��ӡid And a.No = [1] and B.Ʊ��=2" & vbNewLine & _
    " Order By ID"
    
    mblnNOMoved = zlDatabase.NOMoved("����Ԥ����¼", strNO, , 1)
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
        strSQL = Replace(strSQL, "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
    End If
    
    Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    dkpMain.Panes(3).Closed = rsInvoice.EOF
    If rsInvoice.EOF Then
        vsInvoice.Rows = 2
        vsInvoice.Clear 1
        Exit Sub
    End If
    
    With vsInvoice
        .Clear 1
        .Rows = 2
        If rsInvoice.RecordCount <> 0 Then rsInvoice.MoveFirst
        .Rows = IIf(rsInvoice.RecordCount = 0, 1, rsInvoice.RecordCount) + 1
        i = 1
        Do While Not rsInvoice.EOF
            .TextMatrix(i, .ColIndex("ID")) = Nvl(rsInvoice!ID)
            .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = Nvl(rsInvoice!Ʊ�ݺ�)
            .TextMatrix(i, .ColIndex("ʹ��ԭ��")) = Nvl(rsInvoice!ʹ��ԭ��)
            .TextMatrix(i, .ColIndex("ʹ��ʱ��")) = Nvl(rsInvoice!ʹ��ʱ��)
            .TextMatrix(i, .ColIndex("ʹ����")) = Nvl(rsInvoice!ʹ����)
            i = i + 1
            rsInvoice.MoveNext
        Loop
    End With
    vsInvoice.Redraw = flexRDBuffered
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadEInvoiceData(ByVal strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص��ӷ�Ʊ��Ϣ
    '����:���˺�
    '����:2020-03-25 17:13:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSQL As String
    Dim rsEInvoice As ADODB.Recordset
    Dim lngԭԤ��ID As Long
    On Error GoTo errHandle

    vsEInvoice.Tag = ""
    vsEInvoice.Clear 1: vsEInvoice.Rows = 2
     
    If strNO = "" Or mobjEInvoice Is Nothing Then Exit Sub
    Call mobjEInvoice.zlIsStartEinvoicFromNO(strNO, lngԭԤ��ID)
    
    If Not mobjEInvoice.zlGetEInvoiceInforFromBalanceID(lngԭԤ��ID, rsEInvoice, 2, 0) Then Exit Sub
    dkpMain.Panes(4).Closed = rsEInvoice.EOF
    If rsEInvoice.EOF Then Exit Sub
    
    With vsEInvoice
        If rsEInvoice.RecordCount <> 0 Then rsEInvoice.MoveFirst
        i = 1
        Do While Not rsEInvoice.EOF
            .TextMatrix(i, .ColIndex("ID")) = Nvl(rsEInvoice!ID)
            .TextMatrix(i, .ColIndex("��¼״̬")) = Nvl(rsEInvoice!��¼״̬)
            .TextMatrix(i, .ColIndex("����ID")) = Nvl(rsEInvoice!����ID)
            .TextMatrix(i, .ColIndex("��Ʊ����")) = Nvl(rsEInvoice!����)
            .TextMatrix(i, .ColIndex("��Ʊ����")) = Nvl(rsEInvoice!����)
            .TextMatrix(i, .ColIndex("Ʊ�ݽ��")) = Format(Nvl(rsEInvoice!Ʊ�ݽ��), "###0.00;-###0.00;;")
            .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsEInvoice!����ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("����ֽ�ʷ�Ʊ")) = IIf(Val(Nvl(rsEInvoice!�Ƿ񻻿�)) = 1, "�ѻ���", "δ����")
            .TextMatrix(i, .ColIndex("ֽ�ʷ�Ʊ��")) = Nvl(rsEInvoice!ֽ�ʷ�Ʊ��)
            .TextMatrix(i, .ColIndex("��ע")) = Nvl(rsEInvoice!��ע)
            .TextMatrix(i, .ColIndex("����Ա����")) = Nvl(rsEInvoice!����Ա����)
            If Val(Nvl(rsEInvoice!��¼״̬)) = 1 Then
                 .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = Me.ForeColor
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = IIf(Val(Nvl(rsEInvoice!��¼״̬)) = 2, vbRed, vbBlue)
            End If
            i = i + 1: .Rows = .Rows + 1
            rsEInvoice.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mfrmMain.mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyF5
            Call ShowBills(mcllFilter)
        Case vbKeyF9
            RaiseEvent MoneyEnum
        Case vbKeyF11
            RaiseEvent RollingCurtain
        Case vbKeyF12
            RaiseEvent FileLocalSet
        Case vbKeyF
            If Shift = vbCtrlMask Then RaiseEvent FilterDeposit
        Case vbKeyG
            If Shift = vbCtrlMask Then RaiseEvent ViewGo
        Case vbKeyP
            If Shift = vbCtrlMask Then RaiseEvent FilePrint
        Case vbKeyEscape
            mblnGo = False
        Case vbKeyC
            If mfrmMain.mnuEidt_CationMoney_Del.Enabled And mfrmMain.mnuEidt_CationMoney_Del.Visible Then Call ExcuteCationMoney_Del
        Case vbKeyA
            If Shift = vbCtrlMask Then
                If mfrmMain.mnuEdit_Deposit.Enabled And mfrmMain.mnuEdit_Deposit.Visible Then RaiseEvent EditDeposit
            End If
        Case vbKeyR
            If Shift = vbCtrlMask Then RaiseEvent EidtBalanceDel
        Case vbKeyDelete
            If Shift = vbShiftMask Then
                If mfrmMain.mnuEdit_Del.Enabled And mfrmMain.mnuEdit_Del.Visible Then Call ExcuteMoney_Del
            End If
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hwnd, Me.Name
    End Select
End Sub

Private Sub vsDeposit_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsDeposit, Me.Name, "Ԥ����Ϣ�б�", False
End Sub

Private Sub vsDeposit_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsDeposit, Me.Name, "Ԥ����Ϣ�б�", False
End Sub

Private Sub vsDeposit_DblClick()
    If vsDeposit.MouseRow = 0 Then Exit Sub
    If mfrmMain.mnuEdit_View.Enabled Then Call ExcuteViewDepositNO
End Sub

Private Sub vsDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call ExcuteMoney_Del
    Else
        Call Form_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub vsDeposit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RaiseEvent PopupMenu
End Sub

Private Sub vsDeposit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsDeposit.MouseRow = 0 Then
        vsDeposit.MousePointer = 99
    Else
        vsDeposit.MousePointer = 0
    End If
End Sub

Private Sub vsDeposit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = vsDeposit.MouseCol
    
    If Button = 1 And vsDeposit.MousePointer = 99 And lngCol > 0 Then
        If vsDeposit.TextMatrix(0, lngCol) = "" Then Exit Sub
        If vsDeposit.TextMatrix(1, vsDeposit.ColIndex("NO")) = "" Then Exit Sub
    End If
End Sub

Private Sub picBalanceInfo_Resize()
    Err = 0: On Error Resume Next
    With vsBalanceInfor
        .Top = picBalanceInfo.ScaleTop
        .Left = picBalanceInfo.ScaleLeft
        .Height = picBalanceInfo.ScaleHeight
        .Width = picBalanceInfo.ScaleWidth
    End With
End Sub

Private Sub picDeposit_Resize()
    Err = 0: On Error Resume Next
    With picDeposit
        vsDeposit.Left = 0
        vsDeposit.Top = 0
        vsDeposit.Height = .ScaleHeight
        vsDeposit.Width = .ScaleWidth
    End With
End Sub

Private Sub picEinvoice_Resize()
    Err = 0: On Error Resume Next
    With vsEInvoice
        .Top = picEinvoice.ScaleTop
        .Left = picEinvoice.ScaleLeft
        .Height = picEinvoice.ScaleHeight
        .Width = picEinvoice.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    Set mobjEInvoice = Nothing
    mint��� = 0
    mblnNOMoved = False: mblnDateMoved = False
    mstrPrivs = "": mstrEInvoicePrivs = ""
    mblnNotRefresh = False
    Set mrsList = Nothing
    mlngTopRow = 0: mlngCurRow = 0
End Sub

Private Sub vsBalanceInfor_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalanceInfor, Me.Name, "���������Ϣ�б�", False
End Sub

Private Sub vsBalanceInfor_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalanceInfor, Me.Name, "���������Ϣ�б�", False
End Sub

Private Sub vsEInvoice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsEInvoice, Me.Name, "����Ʊ����Ϣ�б�", False
End Sub

Private Sub vsEInvoice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsEInvoice, Me.Name, "����Ʊ����Ϣ�б�", False
End Sub

Private Sub vsInvoice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsInvoice, Me.Name, "��Ʊ��Ϣ�б�", False
End Sub

Private Sub vsInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsInvoice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsInvoice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsInvoice, Me.Name, "��Ʊ��Ϣ�б�", False
End Sub

Private Sub vsInvoice_GotFocus()
    zl_VsGridGotFocus vsInvoice, &HFFC0C0
End Sub

Private Sub vsInvoice_LostFocus()
    zl_VsGridLOSTFOCUS vsInvoice, , vsInvoice.Cell(flexcpForeColor, vsInvoice.Row, vsInvoice.Col)
End Sub
Private Sub InitInvoiceGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ʊ����ؼ�
    '����:���˺�
    '����:2020-03-25 17:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsInvoice
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        .Clear 1: .Rows = 2
        .Cols = 5
        .TextMatrix(0, i) = "ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "Ʊ�ݺ�": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "ʹ��ԭ��": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "ʹ��ʱ��": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "ʹ����": .ColWidth(i) = 1000: i = i + 1
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter: .ColAlignment(i) = flexAlignLeftCenter
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = 1200
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Select Case .ColKey(i)
            Case "ID"
                .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Case "Ʊ�ݺ�"
                .ColAlignment(i) = flexAlignCenterCenter
            End Select
        Next
        
         .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .RowHeightMin = 350
        zl_vsGrid_Para_Restore mlngModule, vsInvoice, Me.Name, "��Ʊ��Ϣ�б�", False
        If .Rows < 2 Then .Rows = 2
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub InitEinvoiceGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ӷ�Ʊ����ؼ�
    '����:���˺�
    '����:2020-03-25 17:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsEInvoice
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        .Clear 1: .Rows = 2
        .Cols = 11
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "��¼״̬": i = i + 1
        .TextMatrix(0, i) = "����ID": i = i + 1
        .TextMatrix(0, i) = "��Ʊ����": i = i + 1
        .TextMatrix(0, i) = "��Ʊ����": i = i + 1
        .TextMatrix(0, i) = "Ʊ�ݽ��": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": i = i + 1
        .TextMatrix(0, i) = "����ֽ�ʷ�Ʊ": i = i + 1
        .TextMatrix(0, i) = "ֽ�ʷ�Ʊ��": i = i + 1
        .TextMatrix(0, i) = "��ע": i = i + 1
        .TextMatrix(0, i) = "����Ա����": i = i + 1
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter: .ColAlignment(i) = flexAlignLeftCenter
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = 1200
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Select Case .ColKey(i)
            Case "��¼״̬"
                .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Case "��ע"
                .ColWidth(i) = 2000
            Case "����Ա����"
                 .ColWidth(i) = 1000
            Case "Ʊ�ݽ��"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
         .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .RowHeightMin = 350
        zl_vsGrid_Para_Restore mlngModule, vsEInvoice, Me.Name, "����Ʊ����Ϣ�б�", False
        If .Rows < 2 Then .Rows = 2
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub InitBalanceGrid()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Integer
    Dim varData As Variant

    strHead = "ID,1,0|���㷽ʽ,1,0|����,1,0|���,1,0|��Ŀ,1,1200|����,1,2000|������ˮ��,1,0 "
    
    With vsBalanceInfor
        .HighLight = flexHighlightWithFocus
        .Redraw = flexRDNone
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColKey(i) = "ID" Or .ColKey(i) = "������ˮ��" Or .ColKey(i) = "���㷽ʽ" Or .ColKey(i) = "����" Or .ColKey(i) = "���" Or .ColKey(i) = "λ��" Then .ColHidden(i) = True
        Next
        If .Rows < 2 Then .Rows = 2
        .RowHeightMin = 350
        '.Row = 1: .Col = 0: .ColSel = .COLS - 1
         .Redraw = flexRDBuffered
        If .TextMatrix(1, 0) = "" Then Exit Sub

        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        .Subtotal flexSTNone, .ColIndex("ID"), .ColIndex("��Ŀ"), gstrDec, &H8000000F
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("��Ŀ")
        .OutlineCol = .ColIndex("��Ŀ")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("��Ŀ")) = strTemp

                strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("���㷽ʽ"))
                strTemp = strTemp & "(" & Format(.Cell(flexcpTextDisplay, i + 1, .ColIndex("���")), gstrDec) & ")"
                If .Cell(flexcpTextDisplay, i + 1, .ColIndex("������ˮ��")) <> "" Then
                   strTemp = strTemp & Space(1) & "������ˮ��:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("������ˮ��"))
                End If
                
                .MergeRow(i) = True
                .MergeCells = flexMergeRestrictRows
                .Cell(flexcpAlignment, i, .ColIndex("��Ŀ"), i, .ColIndex("��Ŀ")) = 1
                
                For j = 0 To .Cols - 1
                   If j <= .ColIndex("����") Then
                       If j >= .ColIndex("��Ŀ") Then
                           .Cell(flexcpText, i, j) = strTemp
                           .Cell(flexcpFontBold, i, j) = False
                       End If
                   End If
                Next
            End If
        Next
        Call .AutoSize(.ColIndex("��Ŀ"))
        For j = 0 To .Cols - 1
            .MergeCol(j) = True
        Next
        zl_vsGrid_Para_Restore mlngModule, vsBalanceInfor, Me.Name, "���������Ϣ�б�", False
    End With
End Sub

Private Sub LoadBalanceInfor(ByVal lngԤ��ID As Long, ByVal bln�Ƿ�Ѻ�� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������㽻����ϸ
    '���:lngԤ��ID-����Ԥ����¼.ID
    '       bln�Ƿ�Ѻ��-�Ƿ�ΪѺ��
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-04-01 18:55:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsInfo As ADODB.Recordset
    
    On Error GoTo errH
    
    If lngԤ��ID = 0 Then
        vsBalanceInfor.Clear 2: vsBalanceInfor.Rows = 2
        Exit Sub
    End If
    If bln�Ƿ�Ѻ�� Then
        strSQL = _
            "Select b.����id || '_' || b.ԭԤ��id As ID, a.���㷽ʽ, Max(c.����) As ����, Sum(A.���) As ���," & vbNewLine & _
            "       b.������Ŀ, b.��������, Max(a.������ˮ��) As ������ˮ�� " & vbNewLine & _
            "From ����Ѻ���¼ A, �������㽻�� B, ҽ�ƿ���� C" & vbNewLine & _
            "Where a.Id = b.����id And a.�����id = c.Id(+) And a.ID = [1] " & vbNewLine & _
            "      And Nvl(b.����,0) = 2 " & vbNewLine & _
            "Group By b.����id, b.ԭԤ��id, a.���㷽ʽ, b.������Ŀ, b.��������" & vbNewLine & _
            "Order By ID"
    Else
        strSQL = _
            "Select b.����id || '_' || b.ԭԤ��id As ID, a.���㷽ʽ, Max(c.����) As ����, Sum(Nvl(-1 * f.���, a.��Ԥ��)) As ���," & vbNewLine & _
            "       b.������Ŀ, b.��������, Max(Nvl(f.������ˮ��, a.������ˮ��)) As ������ˮ�� " & vbNewLine & _
            "From ����Ԥ����¼ A, �������㽻�� B, ҽ�ƿ���� C, ����Ԥ����¼ E, �����˿���Ϣ F" & vbNewLine & _
            "Where a.Id = b.����id And a.�����id = c.Id(+) And a.ID = [1] " & vbNewLine & _
            "      And b.ԭԤ��id = e.Id(+) And e.id = f.��¼id(+) And f.����id(+) =  [1] And Nvl(b.����,0) = 0 " & vbNewLine & _
            "Group By b.����id, b.ԭԤ��id, a.���㷽ʽ, b.������Ŀ, b.��������" & vbNewLine & _
            "Order By ID"
    End If
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        strSQL = Replace(strSQL, "����Ѻ���¼", "H����Ѻ���¼")
        strSQL = Replace(strSQL, "�������㽻��", "H�������㽻��")
    End If
    
    Set rsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngԤ��ID)
    dkpMain.Panes(2).Closed = rsInfo.EOF
    If rsInfo.EOF Then
        vsBalanceInfor.Rows = 2
        vsBalanceInfor.Clear 1
        Exit Sub
    End If
    Set vsBalanceInfor.DataSource = rsInfo
    Call InitBalanceGrid
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picInvoice_Resize()
    Err = 0: On Error Resume Next
    With vsInvoice
        .Top = picInvoice.ScaleTop
        .Left = picInvoice.ScaleLeft
        .Height = picInvoice.ScaleHeight
        .Width = picInvoice.ScaleWidth
    End With
End Sub

Private Sub vsBalanceInfor_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsBalanceInfor, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsBalanceInfor_GotFocus()
    zl_VsGridGotFocus vsBalanceInfor, &HFFC0C0
End Sub

Private Sub vsBalanceInfor_LostFocus()
    zl_VsGridLOSTFOCUS vsBalanceInfor, , vsBalanceInfor.Cell(flexcpForeColor, vsBalanceInfor.Row, vsBalanceInfor.Col)
End Sub

Private Sub vsDeposit_AfterSort(ByVal Col As Long, Order As Integer)
    If vsDeposit.Row <= 0 Or vsDeposit.Col <= 0 Then Exit Sub
    vsDeposit.ForeColorSel = vsDeposit.CellForeColor
End Sub

Private Sub vsDeposit_EnterCell()
    If mblnNotRefresh Then Exit Sub
    vsDeposit.ForeColorSel = vsDeposit.CellForeColor
    Call RefrshDataDetial
End Sub

Private Sub InitdkpMain()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��dkpMain�ؼ�
    '����:����
    '����:2020-04-26
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(1, 1500, 3000, DockTopOf, Nothing)
        objPanel.Handle = picDeposit.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        
        Set objPanel = .CreatePane(2, 500, 1500, DockBottomOf, objPanel)
        objPanel.Title = "���������Ϣ"
        objPanel.Handle = picBalanceInfo.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        Set objPanel = .CreatePane(3, 500, 1500, DockRightOf, objPanel)
        objPanel.Title = "Ԥ��Ʊ����Ϣ"
        objPanel.Handle = picInvoice.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable

        
        Set objPanel = .CreatePane(4, 500, 1500, DockRightOf, objPanel)
        objPanel.Title = "����Ʊ����Ϣ"
        objPanel.Handle = picEinvoice.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable

        .Panes(2).Closed = True
        .Panes(3).Closed = True
        .Panes(4).Closed = True
        
        .Options.HideClient = True

    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SeekBill(blnHead As Boolean)
    Dim i As Long, strTmp As String
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    strTmp = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    RaiseEvent ShowStatus(strTmp)
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To vsDeposit.Rows - 1
        
        '�Ƚ�����
        blnFill = True
        With frmDepositFind
            If .txtNO.Text <> "" Then
                blnFill = blnFill And vsDeposit.TextMatrix(i, vsDeposit.ColIndex("NO")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And vsDeposit.TextMatrix(i, vsDeposit.ColIndex("Ʊ�ݺ�")) = .txtFact.Text
            End If
            If .cbo����Ա.ListIndex > 0 Then
                blnFill = blnFill And vsDeposit.TextMatrix(i, vsDeposit.ColIndex("����Ա")) = zlCommFun.GetNeedName(.cbo����Ա.Text)
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
            If IsNumeric(.txtסԺ��.Text) Then
                blnFill = blnFill And Val(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("סԺ��"))) = Val(.txtסԺ��.Text)
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mlngGo = i + 1
            vsDeposit.Row = i: vsDeposit.TopRow = i
            vsDeposit.Col = 0: vsDeposit.ColSel = vsDeposit.Cols - 1
            strTmp = "�ҵ�һ����¼"
            RaiseEvent ShowStatus(strTmp)
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            strTmp = "�û�ȡ����λ����"
            RaiseEvent ShowStatus(strTmp)
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    strTmp = "�Ѷ�λ���嵥β��"
    RaiseEvent ShowStatus(strTmp)
    Screen.MousePointer = 0
End Sub

Public Sub ExcuteViewDepositNO()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�鿴����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnViewCancel  As Boolean
    Dim strNO As String, str����Ա As String, bytԤ������ As Byte
    Dim int��¼״̬ As Integer, blnNOMoved As Boolean
    Dim bln�Ƿ�Ѻ�� As Boolean
    
    With vsDeposit
        strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        str����Ա = .TextMatrix(.Row, .ColIndex("����Ա"))
        bytԤ������ = Val(.TextMatrix(.Row, .ColIndex("Ԥ�����ID")))
        int��¼״̬ = Val(.TextMatrix(.Row, .ColIndex("��¼״̬")))
        blnViewCancel = int��¼״̬ = 2
        bln�Ƿ�Ѻ�� = .TextMatrix(.Row, .ColIndex("Ѻ������")) <> ""
    End With

    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        blnNOMoved = zlDatabase.NOMoved("����Ԥ����¼", strNO, , "1")
    End If
    
    If strNO = "" Then MsgBox "��ǰû�м�¼���Բ��ģ�", vbExclamation, gstrSysName: Exit Sub
    '��ʾ��������
    If bln�Ƿ�Ѻ�� Then
        Call frmCautionMoney.zlShowEdit(Me, 1, mstrPrivs, mlngModule, strNO, blnViewCancel, blnNOMoved)
    Else
        Call frmDeposit.zlShowEdit(Me, 0, 1, mobjEInvoice, mstrPrivs, mlngModule, bytԤ������, strNO, blnViewCancel, blnNOMoved)
    End If
End Sub

 Public Function ExcuteMoney_Del() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ԥ���˿��Ѻ���˿�
    '����:����
    '����:2020-06-22 11:17:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, str����Ա As String
    Dim bytԤ������ As Byte, blnѺ�� As Boolean
    
    With vsDeposit
        strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        str����Ա = .TextMatrix(.Row, .ColIndex("����Ա"))
        bytԤ������ = Val(.TextMatrix(.Row, .ColIndex("Ԥ�����ID")))
        blnѺ�� = .TextMatrix(.Row, .ColIndex("Ѻ������")) <> ""
    End With
    If strNO = "" Then
        MsgBox "��ǰû�м�¼�����˿", vbExclamation, gstrSysName
        Exit Function
    End If
        
    '����Ȩ��
    If Not BillOperCheck(6, str����Ա, _
        CDate(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("����ʱ��"))), "�˿�") Then Exit Function
    
    If Val(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("���"))) < 0 Then
        MsgBox "�ýɿ��¼���Ϊ��,��ʾ�˿�,����ִ�иò�����", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 6, Me.Caption) Then Exit Function
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    If blnѺ�� Then
         If InStr(1, mstrPrivs, ";Ѻ���˿�;") = 0 Then
            MsgBox "��û��Ȩ�޽���Ѻ���˿������", vbInformation, gstrSysName
            Exit Function
        End If
        On Error Resume Next
        Err.Clear
        ExcuteMoney_Del = frmCautionMoney.zlShowEdit(Me, 2, mstrPrivs, mlngModule, strNO)
        If ExcuteMoney_Del Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then Call ShowBills(mcllFilter)
        End If
        Exit Function
    End If
    
    If ChekcDepositDelPrivs(strNO) = False Then Exit Function

    On Error Resume Next
    Err.Clear
    ExcuteMoney_Del = frmDeposit.zlShowEdit(Me, 0, 2, mobjEInvoice, mstrPrivs, mlngModule, bytԤ������, strNO)
    If ExcuteMoney_Del Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then Call ShowBills(mcllFilter)
    End If
End Function

Private Function ChekcDepositDelPrivs(ByVal strNO As String, Optional ShowMsgbox As Boolean = True, _
                                                           Optional blnErr As Boolean = False) As Boolean
    'blnErr-�Ƿ��쳣����
    '����:���Ԥ���˿�(�����쳣�˿�)Ȩ��
    If Trim(strNO) = "" Then Exit Function
    If is���տ�(strNO) Then
         If InStr(mstrPrivs, "���տ��˿�") = 0 Then
            If ShowMsgbox Then MsgBox "��û��Ȩ�޽��д��տ��˿������", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf InStr(mstrPrivs, "Ԥ���˿�") = 0 Then
        If ShowMsgbox Then MsgBox "��û��Ȩ�޽���Ԥ���˿������", vbInformation, gstrSysName
        Exit Function
    Else
        If blnErr Then ChekcDepositDelPrivs = True: Exit Function
        If HaveSpare(strNO) = 0 And InStr(mstrPrivs, "Ԥ�������˿�") = 0 Then
            If ShowMsgbox Then MsgBox "�ò�����û��Ԥ�����,��û��Ȩ���������ŵ��ݣ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If HaveBalance(strNO) <> 0 Then
            If ShowMsgbox Then MsgBox "�ñ�Ԥ���Ѿ�������ʹ��,�㲻���������ŵ��ݣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    ChekcDepositDelPrivs = True
End Function

Public Function ExcuteCationMoney_Del() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ѻ���˿�
    '����:����
    '����:2020-06-22 11:17:33
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strNO As String, str����Ա As String
    With vsDeposit
        strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        str����Ա = .TextMatrix(.Row, .ColIndex("����Ա"))
    End With
    
    If strNO = "" Then
        MsgBox "��ǰû�м�¼�����˿", vbExclamation, gstrSysName
        Exit Function
    End If
        
    '����Ȩ��
    If Not BillOperCheck(6, str����Ա, _
        CDate(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("����ʱ��"))), "�˿�") Then Exit Function
    
    If Val(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("���"))) < 0 Then
        MsgBox "�ýɿ��¼���Ϊ��,��ʾ�˿�,����ִ�иò�����", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 6, Me.Caption) Then Exit Function
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    ExcuteCationMoney_Del = frmCautionMoney.zlShowEdit(Me, 2, mstrPrivs, mlngModule, strNO)
End Function


