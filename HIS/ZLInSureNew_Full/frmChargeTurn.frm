VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmChargeTurn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "��(��)�����תסԺ"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   8715
   ControlBox      =   0   'False
   Icon            =   "frmChargeTurn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   90
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3180
      Width           =   8400
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8715
      TabIndex        =   12
      Top             =   0
      Width           =   8715
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   6000
         TabIndex        =   2
         Top             =   95
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   3720
         TabIndex        =   1
         Top             =   90
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   609
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   225574915
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   90
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   90963971
         CurrentDate     =   36588
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   180
         Left            =   3480
         TabIndex        =   14
         Top             =   180
         Width           =   90
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շѷ���ʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   1080
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1905
      Left            =   30
      TabIndex        =   4
      ToolTipText     =   "˫�����ݲ鿴��ϸ"
      Top             =   3240
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   3360
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmChargeTurn.frx":058A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   8715
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5280
      Width           =   8715
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   60
         TabIndex        =   8
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫ��(&C)"
         Height          =   350
         Index           =   1
         Left            =   2640
         TabIndex        =   6
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Index           =   0
         Left            =   1455
         TabIndex        =   5
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&X)"
         Height          =   350
         Left            =   7380
         TabIndex        =   7
         Top             =   45
         Width           =   1100
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2535
      Left            =   30
      TabIndex        =   3
      ToolTipText     =   "˫�����ݲ鿴��ϸ"
      Top             =   600
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   4471
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmChargeTurn.frx":08A4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   5760
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeTurn.frx":0BBE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
Attribute VB_Name = "frmChargeTurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

Private mstrNOS As String
Private mlngPatient As Long

Private Enum ����Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
End Enum

Private Enum ҽԺҵ��
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    Support�൥���շѱ���ȫ�� = 39  '�൥���շѱ���ȫ��
End Enum

Private Enum COL
    C0ѡ�� = 0
    C1��� = 1
    C2ҽ�� = 2
    C3���ݺ� = 3
    C4Ʊ�ݺ� = 4
    C5������ = 5
    C6Ӧ�ս�� = 6
    C7ʵ�ս�� = 7
    C8����ʱ�� = 8
    C9����ID = 9
    C10���� = 10
End Enum
Private mbln����תסԺ����� As Boolean


Public Sub ShowME(objParent As Object, ByVal lngPatient As Long, ByRef strNOS As String)
'����:lngPatient-����ID
'����:Ҫ���з���ת��ĵ���,Ʊ��,����ID,����(��ҽ��Ϊ��):H0000001,F000023,81235,901;H0000002,F000045,81263,901;...
    mlngPatient = lngPatient
    mstrNOS = strNOS
    
    '��ʱ������ʽ�����¼�Form_Load
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
    Call SetBillSelected(strNOS)
    
    Call Me.Show(vbModal, objParent)
    
    strNOS = mstrNOS
End Sub

Private Sub SetBillSelected(ByVal strNOS As String)
'˵��:���ת�뼸���ʧ��,�ٽ���ѡ����,��ǰѡ������ѱ�ת��ĵ���������"����ת��",���Բ�Ӧ��ѡ��
    Dim i As Long
    With mshList
        For i = 1 To .Rows - 1
            If InStr(";" & strNOS, ";" & .TextMatrix(i, COL.C3���ݺ�)) > 0 And .TextMatrix(i, COL.C1���) = "��ת��" Then
                .TextMatrix(i, COL.C0ѡ��) = "��"
            Else
                .TextMatrix(i, COL.C0ѡ��) = ""
            End If
        Next
    End With
End Sub

'����:�����Ժʱ��֮���Ƿ����ת������
'����:ת�����ݵĵǼ�ʱ��
Public Function CheckExistTurn(ByVal lngPatient As Long, ByRef dat��Ժʱ�� As Date) As Boolean
    Dim rsTmp As New ADODB.Recordset, strSQL As String
        
    On Error GoTo ErrH
    strSQL = "Select Max(����ʱ��) ����ʱ�� From סԺ���ü�¼" & vbNewLine & _
            "Where ��¼���� = 2 And ��¼״̬ In(1,3) And ����id = [1] And ��ҳid Is Null And ��ʶ�� Is Null"

    'His9�Ͱ汾����������֧�ְ󶨱�����ʽ
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ������ת����", lngPatient)
    If ChkRsState(rsTmp) Then Exit Function
    If Not IsNull(rsTmp!����ʱ��) Then
        dat��Ժʱ�� = rsTmp!����ʱ��
        CheckExistTurn = True
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'����:���¼��ʵ�����ҳID
Public Sub ExecteUpdate(ByVal lngPatient As Long, ByVal strInID As String, ByVal lngPageID As Long, ByVal dat��Ժʱ�� As Date)
    Dim strSQL As String
    
    On Error GoTo ErrH
    strSQL = "Zl_�������תסԺ_Update(" & lngPatient & "," & strInID & "," & lngPageID & _
            ",To_Date('" & Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(strSQL, "���¼��ʵ�")
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'����:����ָ���ĵ��ݺ�����,ִ���������תסԺ����,��ҽ���˷ѽ������
'����:strNOS-Ҫ���з���ת��ĵ���,Ʊ��,����ID,����(��ҽ��Ϊ��):H0000001,F000023,81235,901;H0000002,F000045,81263,901;...
'     strInID-סԺ��,lngPageID-��ҳID,��������������ҽ����Ժ����Ǽ�ʱ�Ŵ���
Public Function ExecuteTurn(ByVal strNOS As String, ByVal strInID As String, ByVal lngPageID As Long, ByVal dat��Ժʱ�� As Date, ByVal lng��Ժ����ID As Long) As Boolean

    Dim DateDel As Date, arrNO As Variant, arrInfo As Variant
    Dim i As Long, j As Long, lngcnt As Long
    Dim strSQL As String, strInvoice As String, strInDate As String, strDelDate As String
    
    Dim blnTrans As Boolean, blnTransMedicare As Boolean, blnDo As Boolean
    Dim intinsure As Integer, strAdvance As String
    
    If strNOS = "" Then Exit Function
    
    strInDate = "To_Date('" & Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    strDelDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrNO = Split(strNOS, ";")
    
    On Error GoTo ErrH
    i = LBound(arrNO)
    Do While i <= UBound(arrNO)
        lngcnt = 1
        strInvoice = Trim(Split(arrNO(i), ",")(1))
        If strInvoice <> "" Then
            For j = i + 1 To UBound(arrNO)
                If strInvoice = Split(arrNO(j), ",")(1) Then
                    lngcnt = lngcnt + 1
                Else
                    Exit For
                End If
            Next
        End If
        
        'ҽ��Ҫ������һ�ſ�ʼ��
        For j = i To i + lngcnt - 1
            gcnOracle.BeginTrans: blnTrans = True
            arrInfo = Split(arrNO(j), ",")
            
            strSQL = "Zl_�������תסԺ_insert('" & arrInfo(0) & "'," & IIf(Val(strInID) = 0, "Null", strInID) & _
                "," & IIf(lngPageID = 0, "Null", lngPageID) & "," & strInDate & "," & lng��Ժ����ID & "," & _
                strDelDate & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "�������תסԺ")
            
            blnTransMedicare = False
            intinsure = Val(arrInfo(3))
            If intinsure <> 0 Then
                strAdvance = lngcnt & "|" & (j - i + 1)
                
                '$IF HIS9.19
                #If gverControl = 0 Then
                    blnDo = gclsInsure.ClinicDelSwap(Val(arrInfo(2)))
                #ElseIf gverControl = 1 Then
                '$ELSE  HIS+
                    blnDo = gclsInsure.ClinicDelSwap(Val(arrInfo(2)), , intinsure)
                #Else
                    blnDo = gclsInsure.ClinicDelSwap(Val(arrInfo(2)), , intinsure, strAdvance)
                #End If
                '$END IF
                
                If Not blnDo Then
                    GoTo ErrH
                Else
                    blnTransMedicare = True
                End If
            End If
            gcnOracle.CommitTrans: blnTrans = False
            
            #If gverControl >= 2 Then
                If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intinsure)
            #End If
        Next
               
        i = i + lngcnt
    Loop

    ExecuteTurn = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        gcnOracle.RollbackTrans
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        #If gverControl >= 2 Then
            If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intinsure)
        #End If
    End If
    
    Call SaveErrLog
End Function


Private Sub ShowBills(ByVal lngPatient As Long, ByVal datBegin As Date, ByVal datEnd As Date)
'����:��ȡ����ʾ����ָ�������ڵ�������õ���
    Dim i As Long, DatTmp As Date, strSQL As String
    Dim rsList As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    
    If datBegin > datEnd Then
        DatTmp = datEnd
        datEnd = datBegin
        datBegin = DatTmp
    End If
    strBegin = "To_Date('" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    strEnd = "To_Date('" & Format(datEnd, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    sta.Panels(2).Text = "���ڶ�ȡ�շѵ���,���Ժ� ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    
    On Error GoTo ErrH
    strSQL = "Select '��' as ѡ��,'��ת��' as ���,Decode(B.����,Null,'','��') as ҽ��, A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.Ӧ�ս��), '900090009" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '900090009" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, A.����ID, B.����" & vbNewLine & _
            "From ������ü�¼ A,���ս����¼ B" & vbNewLine & _
            "Where A.��¼���� = 1 And A.��¼״̬ = 1 And A.����id+0 = [1] And A.����ʱ�� Between [2] And [3] And A.����ID = B.��¼ID(+) And B.����(+) = 1" & vbNewLine & _
            IIf(mbln����תסԺ�����, "           And Exists(Select 1 From ������ü�¼ M,������˼�¼ J where A.ID=J.����ID and M.NO=A.NO And M.��¼����=A.��¼���� And J.������� is Not NULL and  nvl(J.��¼״̬,0)=0 and J.����=1) " & vbNewLine, "") & _
            "Group By A.NO, A.ʵ��Ʊ��, A.������, A.����ʱ��, A.����ID, B.����" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select '' as ѡ��,'����ת��' as ���,Decode(B.����,Null,'','��') as ҽ��, A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.Ӧ�ս��), '900090009" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '900090009" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��,A.����ID,0 as ����" & vbNewLine & _
            "From ������ü�¼ A,���ս����¼ B" & vbNewLine & _
            "Where Mod(A.��¼����,10)=1 And A.��¼״̬ = 3 And A.����id+0 = [1] And A.����ʱ�� Between [2] And [3] And A.����ID = B.��¼ID(+) And B.����(+) = 1" & vbNewLine & _
            "Group By A.NO, A.ʵ��Ʊ��, A.������, A.����ʱ��, A.����ID, B.����" & vbNewLine & _
            "Order By ���, Ʊ�ݺ�, ���ݺ� Desc"

    'ע��:���ڶ൥���˷�Ҫ�����һ�ſ�ʼ��,��������ܹؼ�
    Set rsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatient, datBegin, datEnd)
    mshList.Redraw = False
    mshList.ClearStructure
    mshList.Clear
    mshList.Rows = 2
    
    If rsList.EOF Then
        sta.Panels(2).Text = "û���ҵ�ָ��ʱ�䷶Χ���շѵ���!"
    Else
        Set mshList.DataSource = rsList
        sta.Panels(2).Text = "�� " & rsList.RecordCount & " ���շѵ���"
    End If
    Call SetHeader
    Call SetBillColor
    mshList.Redraw = True
    
    mshList.Row = 1
    mshList.COL = 0: mshList.ColSel = mshList.Cols - 1
    Call mshList_EnterCell
    Screen.MousePointer = 0
    Me.Refresh
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "ѡ��,4,500|���,4,850|ҽ��,4,500|���ݺ�,4,850|Ʊ�ݺ�,4,1100|������,4,800|Ӧ�ս��,7,850|ʵ�ս��,7,850|����ʱ��,4,1850|����ID,4,0|����,4,0"
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        .COL = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub SetBillColor()
    Dim i As Long, j As Long
    
    With mshList
        For i = 1 To .Rows - 1
            .Row = i
            For j = 0 To .Cols - 1
                .COL = j
                If .TextMatrix(i, COL.C1���) = "����ת��" Then
                    .CellForeColor = &H8000000C
                Else
                    .CellForeColor = 0
                End If
            Next
        Next
    End With
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim strTmp As String, Datsys As Date
    
    Call RestoreWinState(Me, App.ProductName)
    mbln����תסԺ����� = IIf(Val(GetPara("����תסԺ�����", glngSys, 1143, 0)) = 1, True, False)
        
    Datsys = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ʼʱ��")
    If IsDate(strTmp) Then
        dtpBegin.Value = CDate(strTmp)
    Else
        dtpBegin.Value = DateAdd("d", -3, Datsys)
    End If
        
    If mstrNOS <> "" Then
        strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����ʱ��")
    Else
        strTmp = ""
    End If
    If IsDate(strTmp) Then
        dtpEnd.Value = CDate(strTmp)
    Else
        dtpEnd.Value = Datsys
    End If
        
    Call SetHeader
    Call SetDetail
End Sub

Private Sub cmdExit_Click()
    Dim i As Long
    
    mstrNOS = ""
    With mshList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, COL.C1���) = "��ת��" And .TextMatrix(i, COL.C0ѡ��) = "��" Then
                mstrNOS = mstrNOS & ";" & .TextMatrix(i, COL.C3���ݺ�) & "," & .TextMatrix(i, COL.C4Ʊ�ݺ�) & _
                        "," & .TextMatrix(i, COL.C9����ID) & "," & .TextMatrix(i, COL.C10����)
            End If
        Next
    End With
    mstrNOS = Mid(mstrNOS, 2)
    
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdRefresh_Click()
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
    If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = asc("'") Then KeyAscii = 0
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    mshList.Width = Me.ScaleWidth - 100
    mshDetail.Width = Me.ScaleWidth - 100
    cmdExit.Left = picBottom.Left + picBottom.Width - cmdExit.Width - 100
    
    pic.Top = picBottom.Top - mshDetail.Height - 100
    mshDetail.Top = pic.Top + 50
    mshList.Height = pic.Top - mshList.Top - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ʼʱ��", Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss")
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����ʱ��", Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mshList.TextMatrix(mshList.Row, COL.C3���ݺ�) = "" Then Exit Sub
    Call SetRowSelected(mshList.Row, Trim(mshList.TextMatrix(mshList.Row, COL.C0ѡ��)) = "")
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    If mshList.TextMatrix(mshList.Row, COL.C3���ݺ�) = "" Then Exit Sub
    If KeyAscii = 32 Then Call SetRowSelected(mshList.Row, Trim(mshList.TextMatrix(mshList.Row, COL.C0ѡ��)) = "")
End Sub


Private Sub cmdAll_Click(Index As Integer)
    Dim i As Long
    
    With mshList
        .Redraw = False
        For i = 1 To .Rows - 1
            If Not SetRowSelected(i, Index = 0) Then
                .Row = i: .COL = 0: .ColSel = .Cols - 1
                Call mshList_EnterCell
                Exit For
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Function SetRowSelected(ByVal lngRow As Long, blnSelect As Boolean) As Boolean
'����:����һ�е�ѡ��״̬
'     ����Ƕ��ŵ����е�һ��,����ͬʱ���ö����е���������
    Dim intinsure As Integer, strNO As String, i As Long, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnDo As Boolean
    
    With mshList
        If .TextMatrix(lngRow, COL.C1���) = "��ת��" And .TextMatrix(lngRow, COL.C0ѡ��) <> IIf(blnSelect, "��", "") Then
            intinsure = Val(.TextMatrix(lngRow, COL.C10����))
            
            If intinsure > 0 And blnSelect Then
                strNO = .TextMatrix(lngRow, COL.C3���ݺ�)
                #If gverControl = 0 Then
                    blnDo = gclsInsure.GetCapability(support�����������)
                #Else
                    blnDo = gclsInsure.GetCapability(support�����������, , intinsure)
                #End If
                
                If Not blnDo Then
                    sta.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧�������������,���в�����ѡ��ת��!"
                    .TextMatrix(lngRow, COL.C0ѡ��) = ""
                    Exit Function
                Else
                    '���жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                    strTmp = GetBalanceType(strNO)
                    If strTmp <> "" Then
                        arrBalanceType = Split(strTmp, ",")
                        For i = 0 To UBound(arrBalanceType)
                            strBalanceType = arrBalanceType(i)
                            blnDo = True
                            #If gverControl >= 2 Then
                                blnDo = gclsInsure.GetCapability(support�����������, , intinsure, strBalanceType)
                            #End If
                            
                            If Not blnDo Then
                                sta.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                .TextMatrix(lngRow, COL.C0ѡ��) = ""
                                Exit Function
                            End If
                        Next
                    End If
                End If
            End If
            
            .TextMatrix(lngRow, COL.C0ѡ��) = IIf(blnSelect, "��", "")
            If intinsure > 0 Then   'ȫ��ѡ���ȡ��
                #If gverControl = 0 Then
                    blnDo = gclsInsure.GetCapability(Support�൥���շѱ���ȫ��)
                #Else
                    blnDo = gclsInsure.GetCapability(Support�൥���շѱ���ȫ��, , intinsure)
                #End If
                
                If blnDo Then
                    If Not SetMultiOther(lngRow, blnSelect, intinsure) Then Exit Function
                End If
            End If
        End If
    End With
    SetRowSelected = True
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intinsure As Integer) As Boolean
'����:���ŵ�������ѡ���ȡ��
'     ���ҽ�����ŵ���Ҫ�������˷�,ѡ������һ��ʱ,ȫѡ����,ȡ��ʱȫȡ��
    Dim i As Long, j As Long, k As Long, strNO As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnDo As Boolean
    
    With mshList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, COL.C1���) = "��ת��" And .TextMatrix(i, COL.C4Ʊ�ݺ�) = .TextMatrix(lngRow, COL.C4Ʊ�ݺ�) And i <> lngRow Then
                If .TextMatrix(i, COL.C0ѡ��) <> .TextMatrix(lngRow, COL.C0ѡ��) Then
                   If intinsure <> 0 And blnSelect Then
                        strNO = .TextMatrix(i, COL.C3���ݺ�)
                        '�жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                         strTmp = GetBalanceType(strNO)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                strBalanceType = arrBalanceType(j)
                                 
                                blnDo = True
                                #If gverControl >= 2 Then
                                    blnDo = gclsInsure.GetCapability(support�����������, , intinsure, strBalanceType)
                                #End If
                                 
                                If Not blnDo Then
                                    sta.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                    For k = 1 To .Rows - 1
                                       If .TextMatrix(k, COL.C4Ʊ�ݺ�) = .TextMatrix(i, COL.C4Ʊ�ݺ�) Then
                                           .TextMatrix(k, COL.C0ѡ��) = ""
                                       End If
                                    Next
                                    Exit Function
                                End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, COL.C0ѡ��) = IIf(blnSelect, "��", "")
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function

Private Function GetBalanceType(ByVal strNO As String) As String
'����:��ȡһ�ŵ����е�ҽ�����㷽ʽ��
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim i As Long
        
    On Error GoTo ErrH
    strSQL = "Select A.���㷽ʽ From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
            "Where A.���㷽ʽ = B.���� And B.���� In (3, 4) And A.NO =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    For i = 1 To rsTmp.RecordCount
        GetBalanceType = GetBalanceType & "," & rsTmp!���㷽ʽ
        rsTmp.MoveNext
    Next
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mshList_EnterCell()
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, COL.C3���ݺ�) = "" Then
        mshDetail.Clear
        mshDetail.Rows = 2
        Call SetDetail
    Else
        Call ShowDetail(mshList.TextMatrix(mshList.Row, COL.C3���ݺ�))
    End If
    
    If mshList.TextMatrix(mshList.Row, COL.C1���) = "����ת��" Then
        mshList.ForeColorSel = mshList.CellForeColor
    Else
        mshList.ForeColorSel = &H80000005
    End If
End Sub


Private Sub ShowDetail(ByVal strNO As String)
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo ErrH
    strSQL = "Select C.���� As ���, Nvl(E.����, B.����) As ����, B.���, A.���㵥λ As ��λ, Avg(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.��׼����), '900090.00000')) As ����, LTrim(To_Char(Sum(A.Ӧ�ս��), '90009" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '90009" & gstrDec & "')) As ʵ�ս��, D.���� As ִ�п���" & vbNewLine & _
            "From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E" & vbNewLine & _
            "Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = '" & strNO & "' And A.��¼���� = 1 And" & vbNewLine & _
            "      A.��¼״̬ In (1, 3) And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3" & vbNewLine & _
            "Group By Nvl(A.�۸񸸺�, A.���), C.����, Nvl(E.����, B.����), B.���, A.���㵥λ, D.����" & vbNewLine & _
            "Order By Nvl(A.�۸񸸺�, A.���)"
    'Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "��ʾ��ϸ")
    
    mshDetail.Redraw = False
    mshDetail.ClearStructure
    mshDetail.Clear
    mshDetail.Rows = 2
    If Not rsTmp.EOF Then Set mshDetail.DataSource = rsTmp
    Call SetDetail
    mshDetail.Redraw = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "���,1,650|����,1,1500|���,1,1450|��λ,4,500|����,7,500|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ִ�п���,4,1000"
    With mshDetail
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        .COL = 0: .ColSel = .Cols - 1
    End With
End Sub


Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If mshList.Height + y < 600 Or mshDetail.Height - y < 800 Then Exit Sub
        pic.Top = pic.Top + y
        mshList.Height = mshList.Height + y
        mshDetail.Top = mshDetail.Top + y
        mshDetail.Height = mshDetail.Height - y
        Me.Refresh
    End If
End Sub

'������˷�Ҫ�����һ�ſ�ʼ��,˳�����Ҫ,���Բ��ṩ�û�����
'Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If mshList.MouseRow = 0 Then
'        mshList.MousePointer = 99
'    Else
'        mshList.MousePointer = 0
'    End If
'End Sub
'
'Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim lngCol As Long
'
'    lngCol = mshList.MouseCol
'
'    If Button = 1 And mshList.MousePointer = 99 Then
'        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
'        If mshList.TextMatrix(mshList.Row, col.c1���) = "" Then Exit Sub
'        If mrsList Is Nothing Then Exit Sub
'
'        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
'        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
'
'        If mlngCurRow <> 0 Then mshList.Row = mlngCurRow
'        If mlngTopRow <> 0 Then mshList.TopRow = mlngTopRow
'    End If
'End Sub




