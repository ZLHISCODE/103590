VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTechnicLog 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ִ�����"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmTechnicLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboResult 
      Height          =   300
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1995
      Width           =   1095
   End
   Begin VB.TextBox txt�������� 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H80000011&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3795
      TabIndex        =   1
      Top             =   225
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4110
      TabIndex        =   7
      Top             =   2670
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2985
      TabIndex        =   6
      Top             =   2670
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -180
      TabIndex        =   14
      Top             =   2415
      Width           =   5970
   End
   Begin VB.TextBox txtִ��ժҪ 
      Height          =   945
      Left            =   1005
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   975
      Width           =   4185
   End
   Begin VB.ComboBox cboִ���� 
      Height          =   300
      Left            =   1005
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1995
      Width           =   2070
   End
   Begin MSComCtl2.DTPicker dtpִ��ʱ�� 
      Height          =   300
      Left            =   1005
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   246153219
      CurrentDate     =   38082
   End
   Begin VB.TextBox txt�������� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3795
      TabIndex        =   3
      Top             =   600
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker dtpҪ��ʱ�� 
      Height          =   300
      Left            =   1005
      TabIndex        =   0
      Top             =   225
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   246153219
      CurrentDate     =   38082
   End
   Begin VB.Label lbl����ʱ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ�н��"
      Height          =   180
      Left            =   3240
      TabIndex        =   17
      Top             =   2055
      Width           =   720
   End
   Begin VB.Label lbl��λ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ"
      Height          =   180
      Index           =   1
      Left            =   4845
      TabIndex        =   16
      Top             =   660
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   3030
      TabIndex        =   15
      Top             =   285
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ��ժҪ"
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ����"
      Height          =   180
      Left            =   420
      TabIndex        =   12
      Top             =   2055
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ��ʱ��"
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   660
      Width           =   720
   End
   Begin VB.Label lbl��λ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ"
      Height          =   180
      Index           =   0
      Left            =   4845
      TabIndex        =   10
      Top             =   285
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   3030
      TabIndex        =   9
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ҫ��ʱ��"
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   285
      Width           =   720
   End
End
Attribute VB_Name = "frmTechnicLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Enum_Inside_Program
Private mlng����ID As Long
Private mlngҽ��ID As Long
Private mlng���ͺ� As Long
Private mlngִ�п���ID As Long
Private mstrִ��ʱ�� As String
Private mbln����ִ�� As Boolean
Private mblnOK As Boolean
Private mdateִ����ֹʱ��  As Date
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstrִ�з��� As String
Private mstr�������� As String
Private mstr������� As String
Private mstrPrivs    As String
Private mstrName     As String
Private mlng������� As Long '���ڵ���ִ��ҽ������ָ��ǰҽ���Ĵ˴η��͵ĵ����б��˷ѵ���������������ָ����ҽ���Ĵ˴η��͵ĵ����б��˷ѵ�������
Private mlng��ȫִ�� As Long 'ִ�н��Ϊ��ȫִ�е��ܴ���
Private mlng����ִ�н��Old  As Long '����ʱȡԭʼ����ִ�н����Ĭ��Ϊ 1 ��ʾִ�У�0/2/3 ����ʾδִ��
Private mlng���δ���Old     As Long '����ʱȡԭʼ����ִ�д���
Private mstrNO As String '���ͼ�¼��Ӧ�ĵ��ݺ�NO
Private mintѪ���� As Integer 'һ���伸��Ѫ
Private mblnѪ������ As Boolean
Private mbln��������ִ�� As Boolean
Private mobjESign As Object '����ǩ���ӿڲ���


Public Function ShowMe(ByVal frmParent As Object, ByVal lngModul As Enum_Inside_Program, ByVal lng����ID As Long, ByVal lngҽ��ID As Long, _
    ByVal lng���ͺ� As Long, ByVal bln����ִ�� As Boolean, Optional ByVal strִ��ʱ�� As String, Optional ByVal lngִ�п���ID As Long, Optional ByVal strName As String, Optional ByVal strPrivs As String) As Boolean
'���ܣ��Ǽǻ����ִ�����
'������lng����ID=��ǰҽ������ID
'      strִ��ʱ��=����ʱ��(yyyy-MM-dd HH:mm:ss)
'���أ��Ƿ�ȡ��
    mlngModul = lngModul
    mlng����ID = lng����ID
    mlngҽ��ID = lngҽ��ID
    mlng���ͺ� = lng���ͺ�
    mlngִ�п���ID = lngִ�п���ID
    mbln����ִ�� = bln����ִ��
    mstrִ��ʱ�� = strִ��ʱ��
    mstrPrivs = strPrivs
    mstrName = strName
    
    On Error Resume Next
    Me.Show 1, frmParent
    
    ShowMe = mblnOK
End Function

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrTime As Variant, strTime As String
    Dim vDate As Date, strPause As String
    Dim lngAllCount As Long, lngCurCount As Long, strCurDate As String, lng��ID As Long
    Dim blnFind As Long, dblTmp As Double
    Dim rsѪ�� As ADODB.Recordset
    
    mblnOK = False
        
    On Error GoTo errH
    
    '��ȡ��������ִ�в���
    mbln��������ִ�� = Val(zlDatabase.GetPara("������Ҫ����ִ��", glngSys)) = 1
    
    '��ȡ����ʱ��
    If mlngҽ��ID <> 0 Then
        strSQL = "select B.����ʱ�� from ����ҽ����¼ B where B.id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID)
        If Not rsTmp.EOF Then
            lbl����ʱ��.Caption = "����ʱ�䣺" & Format(Nvl(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")
        End If
    End If
    
    '��ȡִ����(������Ա)
    strSQL = "Select A.ID,A.���,A.����,A.���� From ��Ա�� A,������Ա B" & _
        " Where A.ID=B.��ԱID And B.����ID=[1]" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    For i = 1 To rsTmp.RecordCount
        cboִ����.AddItem rsTmp!��� & "-" & rsTmp!����
        cboִ����.ItemData(cboִ����.NewIndex) = rsTmp!ID
        If mstrִ��ʱ�� = "" Then
            If rsTmp!ID = UserInfo.ID Then
                cboִ����.ListIndex = cboִ����.NewIndex
                blnFind = True
            End If
        Else
            If rsTmp!���� = mstrName Then
                cboִ����.ListIndex = cboִ����.NewIndex
                blnFind = True
            End If
        End If
        rsTmp.MoveNext
    Next
    
    If InStr(mstrPrivs, "ִ��������Ŀ") > 0 And blnFind = False Then
        cboִ����.AddItem UserInfo.��� & "-" & UserInfo.����
        cboִ����.ItemData(cboִ����.NewIndex) = UserInfo.ID
        cboִ����.ListIndex = cboִ����.NewIndex
    End If
    
    If mlngModul = pҽ������վ Then
        If Val(zlDatabase.GetPara(51, glngSys)) = 1 Then
            Me.cboִ����.Enabled = False
        End If
    End If

    'ִ�н�������˵���ʼ��
    cboResult.Clear
    cboResult.AddItem "δִ��"
    cboResult.AddItem "���"
    cboResult.AddItem "�ܾ�"
    cboResult.AddItem "���"
    cboResult.ListIndex = 1

    mlng������� = 0
    mlng��ȫִ�� = 0
    mlng����ִ�н��Old = 0
    mlng���δ���Old = 0
    '��ȡִ�����
    If mstrִ��ʱ�� = "" Then
        '�ϴ���ִ�е���һЩ����
        strSQL = "Select " & _
            " Max(ִ��ʱ��) as LastDate," & _
            " Max(Ҫ��ʱ��) as curDate," & _
            " Count(Ҫ��ʱ��) as curCount," & _
            " Sum(��������) as curNum" & _
            " From ����ҽ��ִ��" & _
            " Where ҽ��ID=[1] And ���ͺ�=[2] and ��������>0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�)
        If Not rsTmp.EOF Then
            dtpִ��ʱ��.Tag = Format(Nvl(rsTmp!LastDate), "yyyy-MM-dd HH:mm:ss") '�ϴ�ʵ��ִ��ʱ��
            txt��������.Tag = Nvl(rsTmp!curNum, 0) '�ϴ�Ϊֹʵ����ִ�е���������
            strCurDate = Format(Nvl(rsTmp!curDate), "yyyy-MM-dd HH:mm:ss") '�ϴ�ִ�е�Ҫ��ʱ��
            lngCurCount = Nvl(rsTmp!curCount, 0) '�ϴ�Ϊֹʵ����ִ�еĴ���
        End If
        
        '���㱾��ִ��Ӧ�õ�Ҫ��ʱ��
        strSQL = "Select A.��������,Nvl(B.���id, B.ID) ��ID,C.���㵥λ,A.�״�ʱ��,A.ĩ��ʱ��,Decode(B.������Դ, 2, Decode(A.��¼����, 1, 1, Decode(A.�������, 1, 1, 2)), 1) ��������," & _
            " B.��ʼִ��ʱ��,Decode(B.ҽ����Ч,0,B.ִ����ֹʱ��,null) as ִ����ֹʱ��,B.�ϴ�ִ��ʱ��,B.ִ��ʱ�䷽��," & _
            " B.ִ��Ƶ��,B.Ƶ�ʴ���,B.Ƶ�ʼ��,B.�����λ,B.����ID,b.��ҳID,c.���,c.��������,c.ִ�з���,C.���㷽ʽ,B.ҽ����Ч,Nvl(b.�ܸ�����, 1) as �ܸ�����,NVL(B.��������,1) AS ��������,A.NO " & _
            " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C" & _
            " Where A.ҽ��ID=B.ID And B.������ĿID=C.ID(+)" & _
            " And A.ҽ��ID=[1] And A.���ͺ�=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�)
        mstrNO = rsTmp!NO & ""
        '��ѯ�����������Ѿ��˷ѻ����ʵ�ҽ��ִ�д���
        mlng������� = Get�������(mbln����ִ��, mlngҽ��ID, rsTmp!��ID, rsTmp!��� & "", Val(rsTmp!�������� & ""))
        
        lbl��λ(0).Caption = Nvl(rsTmp!���㵥λ)
        lbl��λ(1).Caption = Nvl(rsTmp!���㵥λ)
        txt��������.Text = Nvl(rsTmp!��������)
        dtpִ��ʱ��.Value = zlDatabase.Currentdate
        mdateִ����ֹʱ�� = CDate(Nvl(rsTmp!ִ����ֹʱ��, 0))
        mlng����ID = Val(rsTmp!����ID & "")
        mlng��ҳID = Val(rsTmp!��ҳID & "")
        mstr������� = rsTmp!��� & ""
        mstr�������� = rsTmp!�������� & ""
        mstrִ�з��� = rsTmp!ִ�з��� & ""
        
        '����ִ�м�¼ʱ����Ѫҽ����������
        If gblnѪ��ϵͳ And mstr������� = "E" And mstr�������� = "8" Then
            mblnѪ������ = True
            strSQL = "select zl_Get_��Ѫִ�д���([1]) as ���� from dual"
            Set rsѪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!��ID & ""))
            If Not rsѪ��.EOF Then mintѪ���� = Val(rsѪ��!���� & "")
            lbl��λ(0).Caption = "��"
            lbl��λ(1).Caption = "��"
            txt��������.Text = mintѪ����
            Label1.Tag = txt��������.Tag
            txt��������.Tag = FormatEx(Val(txt��������.Tag) * mintѪ����, 0) '�ϴ�ִ���������Ѿ���5λС��������������Ϳ�������
            If mlng������� = 1 Then '������Ѫҽ�� ��� mlng������� ��ֵֻ���� 0 �� 1 ��Ϊֻ����һ�Ρ�
                MsgBox "��ҽ����ص����Ѿ��˷ѻ����� " & IIf(mbln����ִ��, "������ִ�С�", "������һ��ִ�У��뵥��ִ�С�"), vbInformation, gstrSysName
                Unload Me: Exit Sub
            ElseIf mlng������� = 0 Then
                If Val(txt��������.Tag) >= Val(txt��������.Text) Then
                    MsgBox "��ҽ�����η�������ִ�� " & txt��������.Text & "������ǰ�Ѿ�ִ���� " & Val(txt��������.Tag) & " ����������ִ�С�", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
            dtpҪ��ʱ��.Value = rsTmp!��ʼִ��ʱ�� '��Ѫҽ����Ϊһ����ִ�е�����
            txt��������.Text = 1 'ÿ��ִ��Ĭ��Ϊһ��
            Exit Sub
        Else
            mblnѪ������ = False
            mintѪ���� = 0
        End If
        
        
        '��ǰʵ���Ѿ�ִ����Ҫ��Ĵ���,��׼��ִ��
        If Val(txt��������.Tag) + mlng������� >= Val(txt��������.Text) And (Not (mbln��������ִ�� And mstrִ�з��� = "")) Then
            MsgBox "��ҽ�����η�������ִ�� " & txt��������.Text & IIf(mlng������� <> 0, " �Σ�" & "��ص����Ѿ��˷ѻ�����" & mlng�������, "") & "�Σ���ǰ�Ѿ�ִ���� " & Val(txt��������.Tag) & " �Σ�" & IIf(mbln����ִ��, "������ִ�С�", "������һ��ִ�У��뵥��ִ�С�"), vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        If rsTmp!ִ��Ƶ�� & "" = "һ����" Or rsTmp!ִ��Ƶ�� & "" = "��Ҫʱ" Or (mbln��������ִ�� And mstrִ�з��� = "") Then
            'Ϊһ����ִ�е�����
            dtpҪ��ʱ��.Value = rsTmp!��ʼִ��ʱ��
        ElseIf strCurDate = "" And lngCurCount = 0 Then
            '��һ��ִ��ʱ,��Ϊ�״�ʱ��
            dtpҪ��ʱ��.Value = rsTmp!�״�ʱ��
        Else
            '����ִ��Ƶ�ʷֽ�ʱ��
            strPause = GetAdvicePause(mlngҽ��ID)
            If IsNull(rsTmp!ִ��ʱ�䷽��) And (Nvl(rsTmp!Ƶ�ʴ���, 0) = 0 Or Nvl(rsTmp!Ƶ�ʼ��, 0) = 0 Or IsNull(rsTmp!�����λ)) Then
                '�����Գ���
                lngAllCount = 0: strTime = ""
                vDate = Format(rsTmp!�״�ʱ��, "yyyy-MM-dd")
                Do While vDate <= Format(rsTmp!ĩ��ʱ��, "yyyy-MM-dd")
                    If Not DateIsPause(vDate, strPause) Then
                        lngAllCount = lngAllCount + 1
                        If Format(vDate, "yyyy-MM-dd") > Format(strCurDate, "yyyy-MM-dd") And strTime = "" Then
                            strTime = Format(vDate, "yyyy-MM-dd")
                        End If
                    End If
                    vDate = vDate + 1
                Loop
                
                '��ǰʵ���Ѿ�ִ����Ҫ��Ĵ���,��׼��ִ��
                If lngCurCount + mlng������� >= lngAllCount And (Not (mbln��������ִ�� And mstrִ�з��� = "")) Then
                    MsgBox "��ҽ�����η�������ִ�� " & lngAllCount & "�Σ���ǰ�Ѿ�ִ���� " & lngCurCount & " �Σ�������ִ�С�", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
                
                dtpҪ��ʱ��.Value = CDate(strTime)
            Else
                vDate = Calc�����ڿ�ʼʱ��(rsTmp!��ʼִ��ʱ��, rsTmp!�״�ʱ��, rsTmp!Ƶ�ʼ��, rsTmp!�����λ)
                strTime = Calc���ڷֽ�ʱ��(vDate, rsTmp!ĩ��ʱ��, strPause, rsTmp!ִ��ʱ�䷽�� & "", rsTmp!Ƶ�ʴ���, rsTmp!Ƶ�ʼ��, rsTmp!�����λ, rsTmp!��ʼִ��ʱ��)
                arrTime = Split(strTime, ",")
                lngAllCount = 0
                For i = 0 To UBound(arrTime)
                    If CDate(arrTime(i)) >= rsTmp!�״�ʱ�� Then
                        lngAllCount = lngAllCount + 1
                    End If
                Next
                '��ǰʵ���Ѿ�ִ����Ҫ��Ĵ���,��׼��ִ��
                If lngCurCount + mlng������� >= lngAllCount Then
                    MsgBox "��ҽ�����η�������ִ�� " & lngAllCount & "�Σ���ǰ�Ѿ�ִ���� " & lngCurCount & " �Σ�������ִ�С�", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
                
                dtpҪ��ʱ��.Value = rsTmp!��ʼִ��ʱ��
                For i = 0 To UBound(arrTime)
                    If arrTime(i) > strCurDate Then
                        dtpҪ��ʱ��.Value = CDate(arrTime(i))
                        Exit For '�Ե�һ��ʱ��ΪҪ��ʱ��
                    End If
                Next
                If i > UBound(arrTime) Then
                    dtpҪ��ʱ��.Value = CDate(arrTime(0))
                End If
            End If
        End If
        If Val(rsTmp!���㷽ʽ & "") = 2 Or Val(rsTmp!���㷽ʽ & "") = 1 Then
            If rsTmp!ҽ����Ч = 0 Then
                '1��������ѡƵ�ʡ������ԡ���Ҫʱ�Ͳ���ʱ�Ե�����Ϊ���Ρ�
                txt��������.Text = Val(rsTmp!�������� & "")
            ElseIf InStr("һ����,��Ҫʱ", rsTmp!ִ��Ƶ�� & "") And rsTmp!ִ��Ƶ�� & "" <> "" Then
                '2������һ���Ժ���ҪʱƵ�ʵ�ҽ��ȡ������Ϊ���Ρ�
                If mstr������� = "E" And mstr�������� = "8" Or mstr������� = "K" Then
                    txt��������.Text = Val(rsTmp!�ܸ����� & "") - Val(txt��������.Tag)
                Else
                    txt��������.Text = 1
                End If
            Else
                txt��������.Text = Get��������(mlngҽ��ID, mlng���ͺ�, dtpҪ��ʱ��.Value, Val(rsTmp!�ܸ����� & ""), Val(rsTmp!�������� & ""))
            End If
        Else
            dblTmp = Val(txt��������.Text) - Val(txt��������.Tag) - mlng�������
            txt��������.Text = IIf(dblTmp > 1, 1, dblTmp)
        End If
        If mbln��������ִ�� And mstrִ�з��� = "" Then
            txt��������.Text = 1
        End If
        If Mid(txt��������.Text, 1, 1) = "." Then txt��������.Text = "0" & txt��������.Text
        If gblnѪ��ϵͳ And mstr������� = "K" Then txt��������.Text = Val(txt��������.Text) - Val(txt��������.Tag)
    Else
        '�ϴ���ִ�е���һЩ����(���㱾��)
        strSQL = "Select " & _
            " Max(ִ��ʱ��) as LastDate," & _
            " Sum(��������) as curNum" & _
            " From ����ҽ��ִ��" & _
            " Where ִ��ʱ��<[3] And ҽ��ID=[1] And ���ͺ�=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�, CDate(mstrִ��ʱ��))
        If Not rsTmp.EOF Then
            txt��������.Tag = Nvl(rsTmp!curNum, 0) '�ϴ�Ϊֹʵ����ִ�е���������
            dtpִ��ʱ��.Tag = Format(Nvl(rsTmp!LastDate), "yyyy-MM-dd HH:mm:ss") '�ϴ�ʵ��ִ��ʱ��
        End If
    
        strSQL = "Select A.Ҫ��ʱ��,Nvl(C.���id, C.ID) ��ID,A.ִ��ʱ��,A.��������,A.ִ��ժҪ,nvl(A.ִ�н��,1) as ִ�н��,A.ִ����,B.��������,Decode(C.������Դ, 2, Decode(B.��¼����, 1, 1, Decode(B.�������, 1, 1, 2)), 1) ��������,D.���㵥λ,Decode(c.ҽ����Ч,0,c.ִ����ֹʱ��,null) as ִ����ֹʱ�� ,d.���,d.��������,d.ִ�з���,c.����ID,c.��ҳID,B.NO" & _
            " From ����ҽ��ִ�� A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ D" & _
            " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And B.ҽ��ID=C.ID And C.������ĿID=D.ID(+)" & _
            " And A.ҽ��ID=[1] And A.���ͺ�=[2] And A.ִ��ʱ��=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�, CDate(mstrִ��ʱ��))
        
        '��ѯ���ݵ���ҽ�������ִ�д���
        mlng��ȫִ�� = Get��ȫִ��(mbln����ִ��, mlngҽ��ID, rsTmp!��ID, mlng���ͺ�)
        mstrNO = rsTmp!NO & ""
        '��ѯ�����������Ѿ��˷ѻ����ʵ�ҽ��ִ�д���
        mlng������� = Get�������(mbln����ִ��, mlngҽ��ID, rsTmp!��ID, rsTmp!��� & "", Val(rsTmp!�������� & ""))
        
        dtpҪ��ʱ��.Value = rsTmp!Ҫ��ʱ��
        txt��������.Text = Nvl(rsTmp!��������)
        lbl��λ(0).Caption = Nvl(rsTmp!���㵥λ)
        mdateִ����ֹʱ�� = CDate(Nvl(rsTmp!ִ����ֹʱ��, 0))
        mlng����ID = Val(rsTmp!����ID & "")
        mlng��ҳID = Val(rsTmp!��ҳID & "")
        mstr������� = rsTmp!��� & ""
        mstr�������� = rsTmp!�������� & ""
        mstrִ�з��� = rsTmp!ִ�з��� & ""
        mlng���δ���Old = FormatEx(Nvl(rsTmp!��������), 5)
        mlng����ִ�н��Old = Val(rsTmp!ִ�н�� & "")
        
        dtpִ��ʱ��.Value = rsTmp!ִ��ʱ��
        txt��������.Text = FormatEx(Nvl(rsTmp!��������), 5)
        
        lbl��λ(1).Caption = Nvl(rsTmp!���㵥λ)
        
        txtִ��ժҪ.Text = Nvl(rsTmp!ִ��ժҪ)
        '�޸�ʱ��ȡִ�н��
        cboResult.ListIndex = Val(rsTmp!ִ�н�� & "")
        
        mlng��ȫִ�� = mlng��ȫִ�� - IIf(Val(rsTmp!ִ�н�� & "") = 1, Val(txt��������.Text), 0)
        
        Cbo.SeekIndex cboִ����, rsTmp!ִ����
        '�޸�ִ�м�¼ʱ����Ѫҽ����������
        If gblnѪ��ϵͳ And mstr������� = "E" And mstr�������� = "8" Then
            mblnѪ������ = True
            strSQL = "select zl_Get_��Ѫִ�д���([1]) as ���� from dual"
            Set rsѪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!��ID & ""))
            If Not rsѪ��.EOF Then mintѪ���� = Val(rsѪ��!���� & "")
            'ֻ��Ҫ�������¼������� mlng���δ���Old ��mlng����ִ�н��Old �Ȳ��õ���������Щ���������������飬��Ѫҽ����������Щ���
            lbl��λ(0).Caption = "��"
            lbl��λ(1).Caption = "��"
            txt��������.Text = mintѪ����
            txt��������.Tag = FormatEx(Val(txt��������.Tag) * mintѪ����, 0)
            txt��������.Text = FormatEx(Val("" & rsTmp!��������) * mintѪ����, 0)
            Exit Sub
        Else
            mblnѪ������ = False
            mintѪ���� = 0
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get��������(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal datҪ��ʱ�� As Date, ByVal dbl���� As Double, ByVal dbl���� As Double) As Double
'���ܣ�����ҽ����Ϣִ��ʱ�䣬�����ʱ��ʱ����ҽ����������
    Dim strSQL As String, rsTmp As Recordset
    Dim lng��ǰ���� As Long, i As Long
    Dim dbl����Tmp As Double, dbl���� As Double
    
    strSQL = "Select Ҫ��ʱ�� From ҽ��ִ��ʱ�� Where ҽ��id = [1] And ���ͺ� = [2] Order By Ҫ��ʱ��"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get��������", lngҽ��ID, lng���ͺ�)
    dbl����Tmp = dbl����
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp.RecordCount = 1 Then
                dbl���� = dbl����
            Else
                If i = rsTmp.RecordCount Then
                    dbl���� = dbl����Tmp
                Else
                    If dbl����Tmp >= dbl���� Then
                        dbl���� = dbl����
                    Else
                        dbl���� = dbl����Tmp
                    End If
                    dbl����Tmp = dbl����Tmp - dbl����
                End If
            End If
            If CDate(Format(rsTmp!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:mm:ss")) = datҪ��ʱ�� Then
                Get�������� = dbl����
                Exit For
            End If
            rsTmp.MoveNext
        Next
    Else
        Get�������� = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim blnTrans As Boolean
    Dim lng���� As Long
    Dim strTmp As String
    Dim dbl�������� As Double
    Dim dbl�ѷ����� As Double
    
    If zlCommFun.ActualLen(txtִ��ժҪ.Text) > txtִ��ժҪ.MaxLength Then
        MsgBox "ִ��ժҪ���ݹ��࣬������� " & txtִ��ժҪ.MaxLength \ 2 & " �����ֻ� " & txtִ��ժҪ.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtִ��ժҪ.SetFocus: Exit Sub
    End If
    
    If cboResult.ListIndex = -1 Then
        MsgBox "��ȷ��ִ�����", vbInformation, gstrSysName
        cboResult.SetFocus
        Exit Sub
    End If
    
    dbl�������� = Val(txt��������.Text)
  
    If dbl�������� < 0 Or dbl�������� = 0 And cboResult.ListIndex <= 1 Then
        MsgBox "��ȷ�ϱ���ִ�е�" & IIf(mblnѪ������, "��Ѫ������", "���Ρ�"), vbInformation, gstrSysName
        txt��������.SetFocus: Exit Sub
    End If
     
    If cboִ����.Text = "" Then
        MsgBox "��ȷ��ִ���ˡ�", vbInformation, gstrSysName
        If cboִ����.Enabled Then cboִ����.SetFocus
        Exit Sub
    End If
    If dtpҪ��ʱ��.Value > mdateִ����ֹʱ�� And mdateִ����ֹʱ�� <> CDate(0) Then
        MsgBox "Ҫ��ʱ�䳬����ҽ����ֹʱ�䣬��ȷ��ҽ���Ƿ���ǰֹͣ��", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    '��鱾��ִ��ʱ���Ƿ�����ϴ�ִ��ʱ��
    If IsDate(dtpִ��ʱ��.Tag) Then
        If dtpִ��ʱ��.Value <= CDate(Format(dtpִ��ʱ��.Tag, "yyyy-MM-dd HH:mm:ss")) Then
            MsgBox "����ִ��ʱ��Ӧ�����ϴ�ִ��ʱ�� " & Format(dtpִ��ʱ��.Tag, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
            dtpִ��ʱ��.SetFocus: Exit Sub
        End If
    End If 
    
    '���ÿ��ִ�������Ƿ񳬹��ܵķ�������
    If Val(txt��������.Text) <> 0 And Not mblnѪ������ Then 'δ��д�������εĲ�����(������Գ���)
        If dbl�������� + Val(txt��������.Tag) > Val(txt��������.Text) And Not (mbln��������ִ�� And mstrִ�з��� = "") Then
            MsgBox "��������ִ�����ε�����ִ�����γ�����ҽ���������� " & FormatEx(txt��������.Text, 5) & " " & lbl��λ(0).Caption & "��", vbInformation, gstrSysName
            txt��������.SetFocus: Exit Sub
        End If
        '�����˷ѵĵ�����Ҫ�ж�����������ʵ��������ʵ��������ָ����������ִ�е�ִ����ɴ���+�˷����Ρ�����������ָ������������ִ�е�����ִ�еǼǴ���+�˷�����
        If mlng������� <> 0 And (mlng����ִ�н��Old <> 1 And cboResult.ListIndex = 1 Or mlng���δ���Old <> dbl��������) Then
            If (mlng��ȫִ�� + mlng������� >= Val(txt��������.Text)) Or (Val(txt��������.Tag) + mlng������� > Val(txt��������.Text)) Then
                MsgBox "��ҽ���Ѿ������˷ѻ�����,û��ʣ��ִ�д������������޸ı��ε�ִ�д�����ִ�н���� ", vbInformation, gstrSysName
                txt��������.Text = mlng���δ���Old
                cboResult.ListIndex = mlng����ִ�н��Old
                cmdCancel.SetFocus: Exit Sub
            ElseIf (mlng��ȫִ�� + mlng������� + IIf(cboResult.ListIndex = 1, dbl��������, 0) > Val(txt��������.Text)) Or (Val(txt��������.Tag) + mlng������� + dbl�������� > Val(txt��������.Text)) Then
                lng���� = IIf((Val(txt��������.Text) - mlng��ȫִ�� - mlng�������) > (Val(txt��������.Text) - Val(txt��������.Tag) - mlng�������), (Val(txt��������.Text) - Val(txt��������.Tag) - mlng�������), (Val(txt��������.Text) - mlng��ȫִ�� - mlng�������))
                If lng���� > 0 Then
                    MsgBox "��ҽ�����η�������ִ�� " & txt��������.Text & "��,��ص����Ѿ��˷ѻ�����" & mlng������� & "�Σ�" & _
                            IIf(Val(txt��������.Tag) = 0, "", "��ǰ�Ѿ�ִ���� " & Val(txt��������.Tag) & " �Σ�") & _
                            "������ִ��" & lng���� & "�Ρ�", vbInformation, gstrSysName
                    txt��������.Text = lng����
                    cboResult.ListIndex = mlng����ִ�н��Old
                    cmdCancel.SetFocus: Exit Sub
                Else
                    MsgBox "��ҽ���Ѿ������˷ѻ�����,û��ʣ��ִ�д������������޸ı��ε�ִ�д�����ִ�н���� ", vbInformation, gstrSysName
                    txt��������.Text = mlng���δ���Old
                    cboResult.ListIndex = mlng����ִ�н��Old
                    cmdCancel.SetFocus: Exit Sub
                End If
            End If
        End If
    ElseIf mblnѪ������ Then
        strTmp = FormatEx(dbl��������, 5)
        If InStr(strTmp, ".") > 0 Then
            MsgBox "��Ѫ������Ӧ����С����", vbInformation, gstrSysName
            txt��������.SetFocus: Exit Sub
        End If
        
        If dbl�������� + Val(txt��������.Tag) > Val(txt��������.Text) Then
            MsgBox "����ִ��" & dbl�������� & "�����Ѿ�ִ��" & Val(txt��������.Tag) & "����������������Ҫִ�е���Ѫ���� " & FormatEx(txt��������.Text, 5) & " " & lbl��λ(0).Caption & "��", vbInformation, gstrSysName
            txt��������.SetFocus: Exit Sub
        End If
        
        dbl�������� = FormatEx(dbl�������� / mintѪ����, 5)
        dbl�ѷ����� = Val(Label1.Tag)
        If dbl�������� + dbl�ѷ����� > 1 Then
            dbl�������� = 1 - dbl�ѷ�����
        End If
    End If
    
    If mstr������� = "E" And mstr�������� = "1" And gintCA > 0 And Mid(gstrESign, 2, 1) = "1" Then
        If Not Check����ǩ�� Then Exit Sub
    End If
    '��������
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
    If mstrִ��ʱ�� = "" Then
        If mlngִ�п���ID <> 0 Then
            strSQL = "Zl_����ҽ������_���ұ��(" & mlngҽ��ID & "," & mlng���ͺ� & "," & mlngִ�п���ID & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        
        strSQL = "ZL_����ҽ��ִ��_Insert(" & mlngҽ��ID & "," & mlng���ͺ� & "," & _
            "To_Date('" & Format(dtpҪ��ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            dbl�������� & ",'" & txtִ��ժҪ.Text & "','" & zlCommFun.GetNeedName(cboִ����.Text) & "'," & _
            "To_Date('" & Format(dtpִ��ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            IIf(mbln����ִ��, 1, 0) & "," & "0," & cboResult.ListIndex & ",'','" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngִ�п���ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        strSQL = "ZL_����ҽ��ִ��_Update(To_Date('" & mstrִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS')," & mlngҽ��ID & "," & mlng���ͺ� & "," & _
            "To_Date('" & Format(dtpҪ��ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            dbl�������� & ",'" & txtִ��ժҪ.Text & "','" & zlCommFun.GetNeedName(cboִ����.Text) & "'," & _
            "To_Date('" & Format(dtpִ��ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & "," & cboResult.ListIndex & ",NULL," & IIf(mbln����ִ��, 1, 0) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngִ�п���ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mblnѪ������ = False
    mintѪ���� = 0
    Set mobjESign = Nothing
End Sub

Private Sub txt��������_GotFocus()
    Call zlControl.TxtSelAll(txt��������)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��������_Validate(Cancel As Boolean)
    If Not IsNumeric(txt��������.Text) Then
        txt��������.Text = ""
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub txt��������_GotFocus()
    Call zlControl.TxtSelAll(txt��������)
End Sub

Private Sub txtִ��ժҪ_GotFocus()
    Call zlControl.TxtSelAll(txtִ��ժҪ)
End Sub

Private Function Get�������(ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng��ID As Long, ByVal str������� As String, ByVal int�������� As Integer) As Long
'���ܣ���ȡĳ��ҽ������ĳ��ҽ������������ʵ�ҽ��ִ�д���
'       bln����ִ�� �Ƿ񵥶�ִ�У�����������ڵ��ݵ�ҽ���ĵ���ִ��ĳһ��λ��ĳһ���ּ��
'       lngҽ��ID ����ҽ��ID
'       lng��ID û�и�ҽ�������߸�ҽ��ʱΪҽ��ID,��ҽ��Ϊ���ID
'       str������� ��ҽ�����������
'       int�������� 1-������ã�2-סԺ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTable As String
    Dim rs�Ƽ� As ADODB.Recordset
    Dim lngRes As Long
    Dim dblTmp As Double
    
    On Error GoTo errH
    strTable = IIf(int�������� = 1, "������ü�¼", "סԺ���ü�¼")
    If bln����ִ�� Then
        lng��ID = lngҽ��ID
        strSQL = "Select -1 * Sum(Nvl(a.����, 1) * a.���� / b.����) As ���������" & vbNewLine & _
                "From " & strTable & " A, ����ҽ���Ƽ� B" & vbNewLine & _
                "Where a.ҽ����� = [1] And A.NO=[3] And b.ҽ��id = a.ҽ����� And b.�շ�ϸĿid = a.�շ�ϸĿid And Nvl(B.��������,0)=0 And a.��¼״̬ = 2 And a.��¼���� in(1,2,11) And a.�۸񸸺� Is Null And" & vbNewLine & _
                "      a.�շ���� Not In ('5', '6', '7') And Not Exists" & vbNewLine & _
                " (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1)"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID, str�������, mstrNO)
        If rsTmp.RecordCount <> 0 Then
            lngRes = Val(rsTmp!��������� & "")
        End If
    Else
        strSQL = "Select a.ҽ��id,a.�շ�ϸĿid,count(1) as ���� From ҽ��ִ�мƼ� a Where a.ҽ��id = [1] And a.���ͺ� = [2] and a.����>0 group by a.ҽ��id,a.�շ�ϸĿid"
        Set rs�Ƽ� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�)
        
        'ȡ����������������һ�������շѴ������ٵ���һ��
        strSQL = "select a.ҽ��id,a.�շ�ϸĿid,a.�շѴ��� from (Select a.ҽ����� as ҽ��id,a.�շ�ϸĿid,Sum(Nvl(a.����, 1) * a.���� / b.����) As �շѴ���" & vbNewLine & _
                "       From " & strTable & " A, ����ҽ���Ƽ� B" & vbNewLine & _
                "       Where a.ҽ����� In (Select ID From ����ҽ����¼ Where (ID = [1] Or ���id = [1]) And A.NO=[3] And ������� = [2]) And b.ҽ��id = a.ҽ����� And" & vbNewLine & _
                "             b.�շ�ϸĿid = a.�շ�ϸĿid And Nvl(B.��������,0)=0  And a.��¼���� in(1,2) And a.�۸񸸺� Is Null And a.�շ���� Not In ('5', '6', '7') And" & vbNewLine & _
                "             Not Exists" & vbNewLine & _
                "        (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) " & vbNewLine & _
                "       Group By  a.ҽ�����,a.�շ�ϸĿid) a order by a.�շѴ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID, str�������, mstrNO)
        
        'ʣ��������һ����Ϊ��һ������
        If Not rsTmp.EOF Then
            rs�Ƽ�.Filter = "ҽ��id=" & rsTmp!ҽ��ID & " and �շ�ϸĿid=" & rsTmp!�շ�ϸĿid
            If Not rs�Ƽ�.EOF Then
                dblTmp = Val(rs�Ƽ�!���� & "") - Val(rsTmp!�շѴ��� & "")
                If dblTmp > 0 Then
                    lngRes = IntEx(dblTmp)
                End If
            End If
        End If
    End If
    Get������� = lngRes
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get��ȫִ��(ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng��ID As Long, ByVal lng���ͺ� As Long) As Long
'���ܣ���ȡҽ������ȫִ�д���
'       bln����ִ�� �Ƿ񵥶�ִ�У�����������ڵ��ݵ�ҽ���ĵ���ִ��ĳһ��λ��ĳһ���ּ��
'       lngҽ��ID ����ҽ��ID
'       lng��ID û�и�ҽ�������߸�ҽ��ʱΪҽ��ID,��ҽ��Ϊ���ID
'       lng���ͺ� ����ҽ�����͵ķ��ͺ�

    Dim rsTmp As New ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If bln����ִ�� Then
        lng��ID = lngҽ��ID
        strSQL = "Select Sum(Nvl(b.��������,a.��������)) ��ȫִ�д���" & vbNewLine & _
            "From ����ҽ������ A, ����ҽ��ִ�� B" & vbNewLine & _
            "Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And a.ҽ��id = [1] And a.���ͺ� = [2] And Nvl(b.ִ�н��, 1)=1"
    Else
        strSQL = "Select Max(C.��ȫִ�д���) ��ȫִ�д���" & vbNewLine & _
            "From (Select  Sum(Nvl(b.��������,a.��������)) ��ȫִ�д���" & vbNewLine & _
            "       From ����ҽ������ A, ����ҽ��ִ�� B" & vbNewLine & _
            "       Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And" & vbNewLine & _
            "             a.ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1]) And a.���ͺ� = [2] And" & vbNewLine & _
            "             Nvl(b.ִ�н��, 1)=1" & vbNewLine & _
            "       Group By a.ҽ��id) C"

    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID, lng���ͺ�)
    If rsTmp.RecordCount <> 0 Then
        Get��ȫִ�� = Val(rsTmp!��ȫִ�д��� & "")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check����ǩ��() As Boolean
    '�ж��Ƿ���������ǩ��
    Check����ǩ�� = True
    If gintCA > 0 And CheckSign(2, mlng����ID, , , , False, mobjESign) Then
        If mobjESign Is Nothing Then
            On Error Resume Next
            Set mobjESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not mobjESign Is Nothing Then
                Call mobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If mobjESign Is Nothing Then
            MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
            Check����ǩ�� = False
            Exit Function
        Else
            If Not mobjESign.CheckCertificate(UserInfo.�û���) Then
                Check����ǩ�� = False
                Exit Function
            End If
        End If
    End If
End Function
