VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmLabBarCodeBatPrint 
   Caption         =   "���������ӡ"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9390
   Icon            =   "frmLabBarCodeBatPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   9390
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl RptItem 
      Height          =   3615
      Left            =   60
      TabIndex        =   10
      Top             =   2790
      Width           =   9255
      _Version        =   589884
      _ExtentX        =   16325
      _ExtentY        =   6376
      _StockProps     =   0
      BorderStyle     =   3
      MultipleSelection=   0   'False
      SkipGroupsFocus =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdPrintSend 
      Caption         =   "��ӡ�ͼ쵥(&S)"
      Height          =   350
      Left            =   3075
      TabIndex        =   28
      Top             =   6570
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Frame fraBarCode 
      Caption         =   "ʹ����������"
      Height          =   645
      Left            =   60
      TabIndex        =   24
      Top             =   2100
      Width           =   9285
      Begin VB.TextBox txtBarCode 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1170
         TabIndex        =   25
         Top             =   210
         Width           =   7995
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "��ɨ������"
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ����(&S)"
      Height          =   350
      Left            =   120
      TabIndex        =   22
      Top             =   6570
      Width           =   1395
   End
   Begin VB.CommandButton cmdReturBill 
      Caption         =   "�걾�ͼ�(&I)"
      Height          =   350
      Left            =   4665
      TabIndex        =   21
      Top             =   6570
      Width           =   1395
   End
   Begin VB.PictureBox picBarCodePrint 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   210
      ScaleHeight     =   405
      ScaleWidth      =   675
      TabIndex        =   18
      Top             =   6540
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7830
      TabIndex        =   2
      Top             =   6570
      Width           =   1395
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   6255
      TabIndex        =   1
      Top             =   6570
      Width           =   1395
   End
   Begin VB.Frame fraFilter 
      Caption         =   "����"
      Height          =   2085
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   9285
      Begin VB.CheckBox chkCodePrint 
         Caption         =   "��ʾ�����ӡ"
         Height          =   180
         Left            =   2700
         TabIndex        =   29
         ToolTipText     =   "��ɲɼ�ʱ��ʾ�Ƿ���Ҫ��ӡ����"
         Top             =   1710
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.TextBox txtUnit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   27
         Top             =   1050
         Width           =   7845
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "�����Ѵ�ӡ"
         Height          =   180
         Left            =   1320
         TabIndex        =   20
         Top             =   1710
         Width           =   1275
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "ѡ   ��"
         Height          =   180
         Left            =   270
         TabIndex        =   19
         Top             =   1710
         Value           =   1  'Checked
         Width           =   945
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   240
         TabIndex        =   17
         Top             =   1500
         Width           =   8715
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         ItemData        =   "frmLabBarCodeBatPrint.frx":000C
         Left            =   7020
         List            =   "frmLabBarCodeBatPrint.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   660
         Width           =   1905
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   7830
         TabIndex        =   0
         Top             =   1590
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   675
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   223346691
         CurrentDate     =   39064.0416666667
      End
      Begin VB.ComboBox cboSample 
         Height          =   300
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1905
      End
      Begin VB.ComboBox cboCapture 
         Height          =   300
         Left            =   7020
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   3810
         TabIndex        =   7
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   223346691
         CurrentDate     =   39064.0416666667
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "������λ"
         Height          =   180
         Left            =   270
         TabIndex        =   23
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ִ��״̬"
         Height          =   180
         Left            =   6210
         TabIndex        =   16
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "-----"
         Height          =   180
         Left            =   3270
         TabIndex        =   15
         Top             =   705
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��    ��"
         Height          =   180
         Left            =   270
         TabIndex        =   14
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   270
         TabIndex        =   13
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��    ��"
         Height          =   180
         Left            =   3180
         TabIndex        =   12
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ɼ���ʽ"
         Height          =   180
         Left            =   6210
         TabIndex        =   11
         Top             =   300
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ImageListReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBarCodeBatPrint.frx":0010
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBarCodeBatPrint.frx":007C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLabBarCodeBatPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    ��� = 0
    ҽ��id
    ���ID
    ѡ��
    ͼ��
    �ɼ���ʽ
    ҽ������
    �걾
    ����
    �Ա�
    ����
    ����
    ����
    ��ʶ��
    ����
    ����ID
    ����
    �Թ���ɫ
    �ϲ�ҽ��
    ִ�п���
    ����ҽ��
    ����ʱ��
    ������
    ����ʱ��
    ��Ѫ��
    �Թ�����
    ����
    ������Դ
    Ӥ��
    ����
    �����ӡ
    �ͳ�ʱ��
    ������ĿID
    ����ִ�п���ID
    ��������
End Enum
Dim BlCancel As Boolean                             '������"ESC"��ʱ���Դ���
Private mstrPrivs As String                         'Ȩ��
Private mintBarCodeFormat As Integer                '�����ӡ��ʽ 1=39Code 2=128Code
Private mintExecDept As Integer                     '������ִ�п��Ҵ�ӡ
Private mblnNowConsumption As Boolean                                   '�Ƿ���������

Private Sub CrateRptHead()
    '����           ��ʼ���б�ͷ
    Dim Column As ReportColumn
    With Me.RptItem.Columns
        
        RptItem.AllowColumnRemove = False
        RptItem.ShowItemsInGroups = False
        Me.RptItem.SetImageList Me.ImageListReport
        With RptItem.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "��ѡ��ù����������������Ұ�ť..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mCol.���, "���", 120, False): Column.Visible = False
        Set Column = .Add(mCol.ѡ��, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mCol.ͼ��, "�Թ�", 18, False): Column.Icon = 1
        Set Column = .Add(mCol.����, "����", 60, True)
        Set Column = .Add(mCol.�ɼ���ʽ, "�ɼ���ʽ", 130, True)
        Set Column = .Add(mCol.ҽ������, "ҽ������", 100, True)
        Set Column = .Add(mCol.�걾, "�걾", 45, True)
        Set Column = .Add(mCol.�Ա�, "�Ա�", 45, True)
        Set Column = .Add(mCol.����, "����", 45, True)
        Set Column = .Add(mCol.����, "����", 45, True)
        Set Column = .Add(mCol.����, "����", 45, True)
        Set Column = .Add(mCol.��ʶ��, "��ʶ��", 55, True)
        Set Column = .Add(mCol.����, "����", 120, True)
        
        Set Column = .Add(mCol.ҽ��id, "ҽ��ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.���ID, "���ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����ID, "����ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����, "����", 65, True): Column.Visible = False
        Set Column = .Add(mCol.�ϲ�ҽ��, "�ϲ�ҽ��", 65, True): Column.Visible = False
        Set Column = .Add(mCol.�Թ���ɫ, "�Թ���ɫ", 75, True): Column.Visible = False
        Set Column = .Add(mCol.ִ�п���, "ִ�п���", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����ҽ��, "����ҽ��", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 65, True): Column.Visible = False
        Set Column = .Add(mCol.������, "������", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 65, True): Column.Visible = False
        Set Column = .Add(mCol.��Ѫ��, "��Ѫ��", 65, True): Column.Visible = False
        Set Column = .Add(mCol.�Թ�����, "�Թ�����", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����, "����", 65, True): Column.Visible = False
        Set Column = .Add(mCol.������Դ, "��Դ", 65, True): Column.Visible = False
        Set Column = .Add(mCol.Ӥ��, "Ӥ��", 65, True): Column.Visible = False
        Set Column = .Add(mCol.�����ӡ, "��ӡ", 65, True) ': Column.Visible = False
        Set Column = .Add(mCol.�ͳ�ʱ��, "�ͳ�ʱ��", 130, True)
        Set Column = .Add(mCol.������ĿID, "������ĿID", 130, True): Column.Visible = False
        Set Column = .Add(mCol.����ִ�п���ID, "����ִ�п���ID", 130, True): Column.Visible = False
        Set Column = .Add(mCol.��������, "��������", 130, True): Column.Visible = False
    End With
End Sub

Private Sub cboState_Click()
    Me.chkCodePrint.Visible = False
    Me.cmdPrintSend.Visible = False
    Me.cmdReturBill.Visible = (Me.cboState.Text = "�Ѳ���")
    Me.cmdPrint.Visible = True
    
    Select Case cboState.Text
        Case "δ��"
            Me.cmdPrint.Caption = "��������(&B)"
            Me.cmdReturBill.Visible = (InStr(mstrPrivs, "��ɲɼ�") > 0)
            Me.cmdReturBill.Caption = "��ɲɼ�(&M)"
        Case "�Ѱ�"
            Me.cmdPrint.Visible = (InStr(mstrPrivs, "��ɲɼ�") > 0)
            Me.cmdPrint.Caption = "��ɲɼ�(&F)"
            Me.cmdReturBill.Visible = True
            Me.cmdReturBill.Caption = "��ӡ����(&P)"
            Me.chkCodePrint.Visible = True
        Case "�Ѳ���"
            Me.cmdPrint.Caption = "��ӡ����(&P)"
            Me.cmdReturBill.Visible = True
            Me.cmdReturBill.Caption = "�걾�ͼ�(&I)"
        Case "���ͼ�"
            Me.cmdPrint.Caption = "��ӡ����(&P)"
            Me.cmdReturBill.Visible = True
            Me.cmdReturBill.Caption = "ȡ���ͼ�(&I)"
            Me.cmdPrintSend.Visible = True
        Case "��ִ��"
            Me.cmdPrint.Caption = "��ӡ����(&P)"
    End Select
    
    Call Form_Resize
    Me.TxtBarCode.Tag = ""
    Call ReadData
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkNOComple_Click()
    
End Sub

Private Sub chkPrinted_Click()
    Call cmdSelectAll_Click
End Sub

Private Sub chkPrintNO_Click()
    Call cmdSelectAll_Click
End Sub

Private Sub chkPrint_Click()
    SelectOrCancelReprotCheck Me.RptItem.Records, mCol.ѡ��, Me.chkSelect.Value
End Sub

Private Sub chkSelect_Click()
    SelectOrCancelReprotCheck Me.RptItem.Records, mCol.ѡ��, Me.chkSelect.Value
End Sub

Private Sub cmdCancel_Click()
    BlCancel = True
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    SelectOrCancelReprotCheck Me.RptItem.Records, mCol.ѡ��, False
    Me.RptItem.Populate
End Sub

Private Sub cmdFind_Click()
    Me.TxtBarCode.Tag = ""
    ReadData
    Call cmdSelectAll_Click
End Sub

Private Sub cmdPrint_Click()
    Select Case cboState.Text
        Case "δ��"
            BarCodeMake 1, True
        Case "�Ѱ�"
            If chkCodePrint.Value = 1 Then
                If MsgBox("�Ƿ���Ҫ��ӡ����?", vbYesNo + vbDefaultButton2) = vbYes Then
                    BarCodeMake 2, True
                Else
                    BarCodeMake 2, False
                End If
            Else
                BarCodeMake 2, False
            End If
        Case "�Ѳ���", "���ͼ�"
            BarCodeMake 6, True
        Case "��ִ��"
            BarCodeMake 6, True
    End Select
End Sub

Private Sub cmdPrintSend_Click()
    Dim strName As String
    Dim strID As String
    Dim strTemp As String
    Dim strAdvices As String
    Dim intLoop As Integer
    Dim astrReprot() As String
    
    With Me.RptItem
        If Me.cboState.Text = "���ͼ�" And .Rows.Count > 0 Then
            frmLabSamplingSendInfo.chkPrint.Enabled = False
            If frmLabSamplingSendInfo.ShowME(Me, strName, True) = False Then
                Exit Sub
            End If
            
            For intLoop = 0 To .Rows.Count - 1
                If .Rows(intLoop).Record(mCol.ѡ��).Checked = True Then
                    If Len(strID) >= 3800 Then  '�ַ��������ֶδ���
                        strTemp = strTemp & ";" & Mid(strID, 2)
                        strID = ""
                    Else
                        strID = strID & "," & .Rows(intLoop).Record(mCol.���ID).Value
                    End If
                End If
            Next
            
            strAdvices = strTemp & ";" & Mid(strID, 2)
            astrReprot = Split(strAdvices, ";")
            For intLoop = 0 To UBound(astrReprot)
                If astrReprot(intLoop) <> "" Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_3", Me, "ҽ���ִ�=" & astrReprot(intLoop), 2)
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdPrintSet_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1211_3", Me
End Sub

Private Sub cmdReturBill_Click()
    Select Case cboState.Text
        Case "δ��"
            BarCodeMake 3, True
        Case "�Ѱ�"
            BarCodeMake 6, True
        Case "�Ѳ���", "���ͼ�"
            SampleSend
        Case "��ִ��"
            BarCodeMake 6, True
    End Select
End Sub

Private Sub cmdSelectAll_Click()
    SelectOrCancelReprotCheck Me.RptItem.Records, mCol.ѡ��, True
    Me.RptItem.Populate
End Sub

Private Sub Form_Load()
    '�����б�ͷ
    CrateRptHead
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    '�����ʹ������
    GetInitData
End Sub
Private Sub ReadData(Optional ByVal strBarCode As String)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim Record As ReportRecord
    Dim Item As ReportRecordItem
    Dim i As Integer
    Dim strUnionID As Long                          '���Id
    Dim blnShowExec As Boolean                      '�Ƿ���ʾδ�շ����ﲡ��
    Dim strSQLbak As String
    Dim strSQLCheck As String
    Dim strҽ������ As String
    Dim blnFL As Boolean                            '�Ƿ������ʾ,������Ѫҽ��ʱ,������ʾ
    Dim strBooldSql As String                       '��Ѫ��ѯ���
    zlCommFun.ShowFlash "���ڲ���,���Ժ�...", Me
    Me.MousePointer = vbHourglass
    
    On Error GoTo errH
    
    strSQL = "Select /*+ rule */ a.���,A.ҽ��id, A.���id, A.ҽ������, A.����ʱ��, A.ִ�п���, A.����, A.�Ա�, A.����, A.��������, A.�걾, A.����id, A.�ɼ���ʽ, A.����, A.��ǰ����, A.��ʶ��," & vbNewLine & _
            "       A.����ҽ��, A.����ʱ��, A.������, A.����ʱ��, A.����, A.������Դ, A.Ӥ��, A.����, A.�����Ŀ, A.�����ӡ, A.��¼״̬, A.��¼����, A.�ز�, A.�걾�ͳ�ʱ��," & vbNewLine & _
            "       B.��ɫ As �Թ���ɫ, B.��Ѫ��, B.���� As �Թ�����,a.������ĿID,a.����ִ�п���ID,a.�������� from  " & vbCrLf & _
             " (Select  distinct decode(d.���,'K','��Ѫ','����') ���,B.Id as ҽ��ID,B.���Id,decode(d.���,'K',d.����,b.ҽ������) as ҽ������,m.����ʱ��,H.���� as ִ�п���,I.����,I.�Ա�,I.����,m.��������,b.�걾��λ as �걾, " & vbCrLf & _
             " I.����ID,decode(d.���,'K',b.ҽ������ ,d.����) As �ɼ���ʽ,E.�Թܱ��� as  ����,b.������ĿID,b.ִ�п���ID as ����ִ�п���ID, " & vbCrLf & _
             "decode(i.��ǰ����,null,decode(l.��Ժ����,null,l.��Ժ����,l.��Ժ����),i.��ǰ����) as ��ǰ����, " & vbCrLf & _
             "Decode(B.������Դ,1,I.�����,2,i.סԺ��,4,i.�����) as ��ʶ��, " & vbCrLf & _
             " b.����ҽ��,b.����ʱ��,m.������,m.����ʱ��,decode(b.������־,1,'����','') as ����, " & _
             " b.������Դ,b.Ӥ��,n.���� as ����,E.�����Ŀ,c.�����ӡ,nvl(P.��¼״̬,0) as ��¼״̬, " & vbCrLf & _
             " nvl(P.��¼����,0) as ��¼����,nvl(c.�زɱ걾,0) as �ز�,c.�걾�ͳ�ʱ��,Q.���� as ��������,decode(d.���, 'K', M.ִ��״̬,C.ִ��״̬) ִ��״̬ " & vbCrLf & _
             " From ����ҽ����¼ A, ����ҽ����¼ B,����ҽ������ C,������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
             "      ���ű� H,������Ϣ I,������ҳ L,����ҽ������ M, " & vbCrLf & _
             " (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N,סԺ���ü�¼ P,���ű� Q " & vbCrLf & _
             " Where A.ID = B.���id And A.������Ŀid = D.ID And B.������Ŀid = E.ID And (e.��� = 'E' Or e.��� = 'C') And " & vbCrLf & _
             " B.ִ�п���id = H.ID And d.��� = 'E' And d.�������� = '6'  And " & vbCrLf & _
             " A.����Id = I.����ID And a.����id = l.����ID(+) and a.��ҳid = l.��ҳid(+) and  m.ִ�в���id + 0 = [1] and M.����ʱ�� Between [2] And [3] " & vbCrLf & _
             " And A.id = M.ҽ��ID And b.id = c.ҽ��ID and E.id = N.������ĿID(+)  " & vbCrLf & _
             " And C.ҽ��ID = P.ҽ�����(+) and C.��¼���� = Mod(P.��¼����(+),10) and b.��������ID = Q.ID " & vbCrLf
    
    '�걾
    If cboSample.ItemData(cboSample.ListIndex) <> 0 Then
        strSQL = strSQL & " And B.�걾��λ= [4] "
    End If

    If Me.cboCapture.ItemData(cboCapture.ListIndex) <> 0 Then
        strSQL = strSQL & " And D.���� = [5] "
    End If
    
    '���˵�λ
    If Trim(Me.txtUnit.Text) <> "" Then
        strSQL = strSQL & " and I.������λ like [6] "
    End If
    
    '����
    If strBarCode <> "" Then
        strSQL = strSQL & " And m.�������� = [7]"
    End If
    
    If Me.cboState = "δ��" Then
        strSQL = strSQL & " And c.�������� is null) a , ��Ѫ������ b "
    ElseIf Me.cboState = "�Ѱ�" Then
        strSQL = strSQL & " And c.�������� is not null and c.������ is null) a , ��Ѫ������ B "
    ElseIf Me.cboState = "�Ѳ���" Then
        strSQL = strSQL & " and c.�������� is not null and c.������ is not null and c.�걾�ͳ�ʱ�� is null) a,��Ѫ������ B  "
    ElseIf Me.cboState = "���ͼ�" Then
        strSQL = strSQL & " and c.�������� is not null and c.������ is not null and c.�걾�ͳ�ʱ�� is not null) a,��Ѫ������ B  "
    ElseIf Me.cboState = "��ִ��" Then
        strSQL = strSQL & ") a,��Ѫ������ B "
    End If
    strSQL = strSQL & " Where a.���� = b.���� "
    If Me.cboState = "��ִ��" Then
        strSQL = strSQL & " and a.ִ��״̬ IN (1,3)"
    Else
        strSQL = strSQL & " and a.ִ��״̬ IN (0,2)"
    End If
    strSQLbak = strSQL
    
    strSQLCheck = strSQL
    
    If Me.cboState = "δ��" And mblnNowConsumption = True Then
        '����������Ҫ����ȷ�ϵ�����
        strSQLCheck = Replace$(strSQLCheck, "סԺ���ü�¼", "������ü�¼")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQLCheck, gstrSysName, cboDept.ItemData(cboDept.ListIndex), _
                CDate(Format(DtpBegin.Value, "yyyy-MM-dd hh:mm:00")), _
                CDate(Format(DTPEnd.Value, "yyyy-MM-dd hh:mm:59")), Mid(cboSample.Text, InStr(1, cboSample.Text, "-") + 1), _
                cboCapture.Text, "%" & Me.txtUnit & "%", strBarCode)
        If rsTmp.RecordCount > 0 Then
            rsTmp.filter = "��¼״̬ = 0 and ������Դ <> 2 "
            If rsTmp.RecordCount > 0 Then
                MsgBox "��¼�����������δȷ�ϵĲ��ˣ����ܽ�����������!", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
    End If
    
    strBooldSql = GetBooldReadDAtaSql(strBarCode)
    strSQLbak = Replace$(strSQLbak, "סԺ���ü�¼", "������ü�¼")
    strSQL = strSQL & " union all " & strSQLbak
    
    '����
'    strSQL = strSQL & " order by  ����ID,E.�Թܱ���,b.�걾��λ,a.ҽ������,B.���Id,b.����ʱ��,e.�����Ŀ "
    strSQL = strSQL & " Order By ���,����id, ����,���id,ִ�п���, �걾,����, ҽ������,  ����ʱ��, �����Ŀ "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, cboDept.ItemData(cboDept.ListIndex), _
                CDate(Format(DtpBegin.Value, "yyyy-MM-dd hh:mm:00")), _
                CDate(Format(DTPEnd.Value, "yyyy-MM-dd hh:mm:59")), Mid(cboSample.Text, InStr(1, cboSample.Text, "-") + 1), _
                cboCapture.Text, "%" & Me.txtUnit & "%", strBarCode)
    
    If strBarCode = "" Then
        RptItem.Records.DeleteAll
        RptItem.GroupsOrder.DeleteAll
    Else
        If rsTmp.EOF Then
            MsgBox "δ��ѯ�������Ϊ��" & strBarCode & "���ı걾��"
            TxtBarCode.SetFocus
            Me.MousePointer = vbDefault
            zlCommFun.StopFlash
            Exit Sub
        Else
            For Each Record In RptItem.Records
                If Record.Item(mCol.����).Value = strBarCode Then
                    MsgBox "�����Ϊ��" & strBarCode & "���ı걾�Ѵ��ڣ�"
                    TxtBarCode.SetFocus
                    Me.MousePointer = vbDefault
                    zlCommFun.StopFlash
                    Exit Sub
                End If
            Next
        End If
    End If
                    
    With Me.RptItem
        Do Until rsTmp.EOF
            blnShowExec = True
            '����Ȩ�����ж��Ƿ���ʾδ�շѵ������¼
            If InStr(mstrPrivs, "��ʾ���ۼ�¼") <= 0 Then
                If Nvl(rsTmp("��¼״̬"), 0) = 0 And ((Nvl(rsTmp("��¼����"), 0) Mod 10) = 1 Or Nvl(rsTmp("��¼����"), 0) = 0) Then blnShowExec = False
            End If
            
            If blnShowExec = True Then
                If strUnionID <> Nvl(rsTmp("���ID")) Then
                
                    Set Record = .Records.Add
                    For i = 0 To .Columns.Count
                        Record.AddItem ""
                    Next
                    
                    Set Item = Record(mCol.ѡ��): Item.HasCheckbox = True: Item.Checked = True
                    Record(mCol.ͼ��).BackColor = Nvl(rsTmp("�Թ���ɫ"), -1)
                    Record(mCol.�걾).Value = Nvl(rsTmp("�걾"))
                    Record(mCol.�ɼ���ʽ).Value = Nvl(rsTmp("�ɼ���ʽ"))
                    Record(mCol.����).Value = Nvl(rsTmp("ִ�п���"))
                    Record(mCol.ִ�п���).Value = Nvl(rsTmp("ִ�п���"))
                    Record(mCol.����ִ�п���ID).Value = Nvl(rsTmp("����ִ�п���ID"))
                    Record(mCol.����).Value = Nvl(rsTmp("����"))
                    Record(mCol.����).Value = Nvl(rsTmp("��������"))
                    Record(mCol.����).Value = Nvl(rsTmp("����")) & IIf(Nvl(rsTmp("Ӥ��"), "0") = 0, "", "(Ӥ��" & rsTmp("Ӥ��") & ")")
                    Record(mCol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
                    Record(mCol.ҽ��id).Value = Nvl(rsTmp("ҽ��ID"))
                    Record(mCol.���ID).Value = Nvl(rsTmp("���ID"))
                    Record(mCol.ҽ������).Value = Nvl(rsTmp("ҽ������"))
                    Record(mCol.����ID).Value = Nvl(rsTmp("����ID"))
                    Record(mCol.����).Value = Nvl(rsTmp("����"))
                    Record(mCol.����).Value = Nvl(rsTmp("��ǰ����"))
                    Record(mCol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
                    Record(mCol.�Թ���ɫ).Value = Nvl(rsTmp("�Թ���ɫ"), -1)
                    Record(mCol.����ҽ��).Value = Nvl(rsTmp("����ҽ��"))
                    Record(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                    Record(mCol.������).Value = Nvl(rsTmp("������"))
                    Record(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                    Record(mCol.��Ѫ��).Value = Nvl(rsTmp("��Ѫ��"))
                    Record(mCol.�Թ�����).Value = Nvl(rsTmp("�Թ�����"))
                    Record(mCol.����).Value = Nvl(rsTmp("����"))
                    Record(mCol.������Դ).Value = Nvl(rsTmp("������Դ"))
                    Record(mCol.Ӥ��).Value = Nvl(rsTmp("Ӥ��"), 0)
                    Record(mCol.�ͳ�ʱ��).Value = Nvl(rsTmp("�걾�ͳ�ʱ��"))
                    Record(mCol.��������).Value = Nvl(rsTmp("��������"))
                    Record(mCol.������ĿID).Value = Nvl(rsTmp("������ĿID"))
                    Record(mCol.����).Value = IIf(Trim(Nvl(rsTmp("����"))) = "", Nvl(rsTmp("ҽ������")), Nvl(rsTmp("����")))
                    Record(mCol.�����ӡ).Value = IIf(Val(Nvl(rsTmp("�����ӡ"))) = 0, "δ��ӡ", "�Ѵ�ӡ")
                    Record(mCol.���).Value = Nvl(rsTmp("���"))
                    If Nvl(rsTmp("���")) = "��Ѫ" Then blnFL = True    '��������Ѫҽ��ʱ,��Ҫ������ʾ,������ʾ��ʦ������Ѫҽ��,��Ҫ��������ҽ������
                    For i = 0 To .Columns.Count
                        Record(i).ForeColor = Nvl(rsTmp("�Թ���ɫ"), -1)
                    Next
                    
                    If Nvl(rsTmp("�ز�")) = 1 Then
                        For i = 0 To .Columns.Count
                            Record(i).Bold = True
                        Next
                    End If
                Else
                    Record(mCol.�ϲ�ҽ��).Value = Record(mCol.�ϲ�ҽ��).Value & "," & Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                    
                    strҽ������ = Nvl(rsTmp("ҽ������"))
                    If InStr(";" & Record(mCol.ҽ������).Value & ";", strҽ������) <= 0 Then
                        Record(mCol.ҽ������).Value = Record(mCol.ҽ������).Value & ";" & Nvl(rsTmp("ҽ������"))
                    End If
                        
                    strҽ������ = IIf(Trim(Nvl(rsTmp("����"))) = "", Nvl(rsTmp("ҽ������")), Nvl(rsTmp("����")))
                    If InStr(";" & Record(mCol.����).Value & ";", strҽ������) <= 0 Then
                        Record(mCol.����).Value = Record(mCol.����).Value & ";" & strҽ������
                    End If
                End If
                
                strUnionID = Nvl(rsTmp("���ID"))
            End If
            
            rsTmp.MoveNext
        Loop
        .Columns(mCol.ѡ��).TreeColumn = False
        
        If blnFL = True Then
            Call .GroupsOrder.Add(.Columns.Column(mCol.���))
        End If
        .Populate
        If strBarCode <> "" And .Records.Count > 0 Then
            Me.TxtBarCode.Tag = .Records.Count
        End If
    End With
    Call chkPrint_Click
    Me.MousePointer = vbDefault
    zlCommFun.StopFlash
    
    Exit Sub
errH:
    Me.MousePointer = vbDefault
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function GetBooldReadDAtaSql(Optional ByVal strBarCode As String) As String
    Dim strSQL
    Dim strSQLbak As String
    On Error GoTo errH
    
    strSQL = "Select /*+ rule */ a.���,A.ҽ��id, A.���id, A.ҽ������, A.����ʱ��, A.ִ�п���, A.����, A.�Ա�, A.����, A.��������, A.�걾, A.����id, A.�ɼ���ʽ, A.����, A.��ǰ����, A.��ʶ��," & vbNewLine & _
            "       A.����ҽ��, A.����ʱ��, A.������, A.����ʱ��, A.����, A.������Դ, A.Ӥ��, A.����, A.�����Ŀ, A.�����ӡ, A.��¼״̬, A.��¼����, A.�ز�, A.�걾�ͳ�ʱ��," & vbNewLine & _
            "       B.��ɫ As �Թ���ɫ, B.��Ѫ��, B.���� As �Թ�����,a.������ĿID,a.����ִ�п���ID,a.�������� from  " & vbCrLf & _
             " (Select  distinct decode(d.���,'K','��Ѫ','����') ���,B.Id as ҽ��ID,B.���Id,decode(d.���,'K',d.����,b.ҽ������) as ҽ������,m.����ʱ��,H.���� as ִ�п���,I.����,I.�Ա�,I.����,m.��������,b.�걾��λ as �걾, " & vbCrLf & _
             " I.����ID,decode(d.���,'K',b.ҽ������ ,d.����) As �ɼ���ʽ,E.�Թܱ��� as  ����,b.������ĿID,b.ִ�п���ID as ����ִ�п���ID, " & vbCrLf & _
             "decode(i.��ǰ����,null,decode(l.��Ժ����,null,l.��Ժ����,l.��Ժ����),i.��ǰ����) as ��ǰ����, " & vbCrLf & _
             "Decode(B.������Դ,1,I.�����,2,i.סԺ��,4,i.�����) as ��ʶ��, " & vbCrLf & _
             " b.����ҽ��,b.����ʱ��,m.������,m.����ʱ��,decode(b.������־,1,'����','') as ����, " & _
             " b.������Դ,b.Ӥ��,n.���� as ����,E.�����Ŀ,c.�����ӡ,nvl(P.��¼״̬,0) as ��¼״̬, " & vbCrLf & _
             " nvl(P.��¼����,0) as ��¼����,nvl(c.�زɱ걾,0) as �ز�,c.�걾�ͳ�ʱ��,Q.���� as ��������,decode(d.���, 'K', M.ִ��״̬,C.ִ��״̬) ִ��״̬ " & vbCrLf & _
             " From ����ҽ����¼ A, ����ҽ����¼ B,����ҽ������ C,������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
             "      ���ű� H,������Ϣ I,������ҳ L,����ҽ������ M, " & vbCrLf & _
             " (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N,סԺ���ü�¼ P,���ű� Q " & vbCrLf & _
             " Where A.ID = B.���id And A.������Ŀid = D.ID And B.������Ŀid = E.ID And (e.��� = 'E' Or e.��� = 'C') And " & vbCrLf & _
             " B.ִ�п���id = H.ID And d.��� = 'K'  And  e.�������� = '9' And " & vbCrLf & _
             " A.����Id = I.����ID And a.����id = l.����ID(+) and a.��ҳid = l.��ҳid(+) and  c.ִ�в���id + 0 = [1] and M.����ʱ�� Between [2] And [3] " & vbCrLf & _
             " And A.id = M.ҽ��ID And b.id = c.ҽ��ID and E.id = N.������ĿID(+)  " & vbCrLf & _
             " And C.ҽ��ID = P.ҽ�����(+) and C.��¼���� = Mod(P.��¼����(+),10) and b.��������ID = Q.ID " & vbCrLf
    
    '�걾
    If cboSample.ItemData(cboSample.ListIndex) <> 0 Then
        strSQL = strSQL & " And decode(d.���, 'K',[4],B.�걾��λ)= [4] "
    End If

    If Me.cboCapture.ItemData(cboCapture.ListIndex) <> 0 Then
        strSQL = strSQL & " And decode(d.���, 'K',E.����,D.����) = [5] "
    End If
    
    '���˵�λ
    If Trim(Me.txtUnit.Text) <> "" Then
        strSQL = strSQL & " and I.������λ like [6] "
    End If
    
    '����
    If strBarCode <> "" Then
        strSQL = strSQL & " And m.�������� = [7]"
    End If
    
    If Me.cboState = "δ��" Then
        strSQL = strSQL & " And c.�������� is null) a , ��Ѫ������ b "
    ElseIf Me.cboState = "�Ѱ�" Then
        strSQL = strSQL & " And c.�������� is not null and c.������ is null) a , ��Ѫ������ B "
    ElseIf Me.cboState = "�Ѳ���" Then
        strSQL = strSQL & " and c.�������� is not null and c.������ is not null and c.�걾�ͳ�ʱ�� is null) a,��Ѫ������ B  "
    ElseIf Me.cboState = "���ͼ�" Then
        strSQL = strSQL & " and c.�������� is not null and c.������ is not null and c.�걾�ͳ�ʱ�� is not null) a,��Ѫ������ B  "
    ElseIf Me.cboState = "��ִ��" Then
        strSQL = strSQL & ") a,��Ѫ������ B "
    End If
    strSQL = strSQL & " Where a.���� = b.���� "
    If Me.cboState = "��ִ��" Then
        strSQL = strSQL & " and a.ִ��״̬ IN (1,3)"
    Else
        strSQL = strSQL & " and a.ִ��״̬ IN (0,2)"
    End If
    strSQLbak = strSQL
    
    strSQLbak = Replace$(strSQLbak, "סԺ���ü�¼", "������ü�¼")
    strSQL = strSQL & " union all " & strSQLbak
    GetBooldReadDAtaSql = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub GetInitData()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngDeptID As Long                       '����ID
    Dim lngSampleID As Long                     '�걾ID
    Dim lngCaptureID As Long                    '�ɼ���ʽ
    Dim intUnionState As Integer                'ִ��״̬
    Dim intSpaceDate As Integer                 '���ʱ��
    Dim strNowDate As Date                      'ȡ��ǰ������ʱ��
    
    intSpaceDate = DateDiff("d", Me.DTPEnd.Value, Me.DtpBegin.Value)
   
    lngDeptID = zlDatabase.GetPara("frmLabBarCodeBatPrint_��������Id", 100, 1208, 0)
    lngSampleID = zlDatabase.GetPara("frmLabBarCodeBatPrint_�걾ID", 100, 1208, 0)
    lngCaptureID = zlDatabase.GetPara("frmLabBarCodeBatPrint_�ɼ�����", 100, 1208, 0)
    intUnionState = zlDatabase.GetPara("frmLabBarCodeBatPrint_ִ��״̬", 100, 1208, 0)
'    Me.chkComplete.Value = zlDatabase.GetPara("frmLabBarCodeBatPrint_�Ƿ���Ϊ���", 100, 1208, 0)
    intSpaceDate = zlDatabase.GetPara("frmLabBarCodeBatPrint_���ʱ��", 100, 1208, 2)
    
    On Error GoTo errH
    
    '===�������
    strSQL = _
            " Select Distinct A.ID,A.���� || '-' || A.���� as ����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(1,2,3,4) And B.�������� IN('����','����')"
            
            
    If InStr(1, mstrPrivs, "���п���") <= 0 Then
        strSQL = strSQL & " And C.��Աid = [1] "
    End If
            
    strSQL = strSQL & " Order by A.���� || '-' || A.����"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    cboDept.Clear
    Do Until rsTmp.EOF
        cboDept.AddItem rsTmp("����")
        cboDept.ItemData(cboDept.NewIndex) = rsTmp("ID")
        If rsTmp("id") = IIf(lngDeptID = 0, UserInfo.����ID, lngDeptID) Then
            cboDept.ListIndex = cboDept.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cboDept.Text = "" And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    
    '===����ɼ���ʽ(����Ѫ����Ѫ�ɼ�)
    strSQL = "select ID,���� from ������ĿĿ¼ where (����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL) And �������� in ('6','9') And ���='E'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    cboCapture.Clear
    cboCapture.AddItem "���вɼ���ʽ"
    cboCapture.ItemData(cboCapture.NewIndex) = 0
    Do Until rsTmp.EOF
        cboCapture.AddItem rsTmp("����")
        cboCapture.ItemData(cboCapture.NewIndex) = rsTmp("ID")
        If lngCaptureID = rsTmp("id") Then
            cboCapture.ListIndex = cboCapture.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cboCapture.Text = "" And cboCapture.ListCount > 0 Then cboCapture.ListIndex = 0
    
    '===�������걾
    strSQL = "select ����,���� from ���Ƽ���걾"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    cboSample.Clear
    cboSample.AddItem "���б걾"
    Do Until rsTmp.EOF
        cboSample.AddItem rsTmp("����") & "-" & rsTmp("����")
        cboSample.ItemData(cboSample.NewIndex) = rsTmp("����")
        If rsTmp("����") = lngSampleID Then
            cboSample.ListIndex = cboSample.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cboSample.Text = "" And cboSample.ListCount > 0 Then cboSample.ListIndex = 0
    
    '===ִ��״̬
    cboState.Clear
    cboState.AddItem "δ��"
    cboState.ItemData(cboState.NewIndex) = 0
    cboState.AddItem "�Ѱ�"
    cboState.ItemData(cboState.NewIndex) = 1
    cboState.AddItem "�Ѳ���"
    cboState.ItemData(cboState.NewIndex) = 2
    cboState.AddItem "���ͼ�"
    cboState.ItemData(cboState.NewIndex) = 3
    cboState.AddItem "��ִ��"
    cboState.ItemData(cboState.NewIndex) = 4
    cboState.ListIndex = intUnionState
    
    
    '===ʱ���
    strNowDate = zlDatabase.Currentdate
    Me.DTPEnd = Format(strNowDate, "yyyy-mm-dd hh:mm")
    Me.DtpBegin.Value = Format(strNowDate - intSpaceDate, "yyyy-mm-dd 00:00")
    
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With fraFilter
        .Width = Me.ScaleWidth - 7 * Screen.TwipsPerPixelX
    End With
    
    With cboCapture
        .Width = fraFilter.Width - .Left - 15 * Screen.TwipsPerPixelX
    End With
    
    With cboState
        .Width = fraFilter.Width - .Left - 15 * Screen.TwipsPerPixelX
    End With
    
    With txtUnit
        .Width = fraFilter.Width - .Left - 15 * Screen.TwipsPerPixelX
    End With
    
    With cmdFind
        .Left = fraFilter.Width - .Width - 15 * Screen.TwipsPerPixelX
    End With
    
    With Frame1
        .Width = fraFilter.Width - .Left - 15 * Screen.TwipsPerPixelX
    End With
    
    If cboState.Text = "�Ѳ���" Or cboState.Text = "���ͼ�" Or cboState.Text = "�Ѱ�" Then
        Me.fraBarCode.Visible = True
        Me.fraBarCode.Width = Me.fraFilter.Width
        Me.TxtBarCode.Width = Me.fraBarCode.Width - Me.TxtBarCode.Left - 40
        Me.RptItem.Top = Me.fraBarCode.Top + Me.fraBarCode.Height + 20
    Else
        Me.fraBarCode.Visible = False
        Me.fraBarCode.Width = Me.fraFilter.Width
        Me.TxtBarCode.Width = Me.fraBarCode.Width - Me.TxtBarCode.Left - 40
        Me.RptItem.Top = Me.fraFilter.Top + Me.fraFilter.Height + 20
    End If
    

    
    With RptItem
        .Width = Me.fraFilter.Width
        .Height = Me.ScaleHeight - .Top - Me.cmdCancel.Height - 20 * Screen.TwipsPerPixelY
    End With
    
    With cmdCancel
        .Top = Me.ScaleHeight - .Height - 10 * Screen.TwipsPerPixelY
        .Left = Me.ScaleWidth - .Width - 20 * Screen.TwipsPerPixelX
    End With
    
    With cmdPrint
        .Top = Me.cmdCancel.Top
        .Left = Me.cmdCancel.Left - .Width - 20 * Screen.TwipsPerPixelX
    End With
    
    With cmdReturBill
        .Top = Me.cmdCancel.Top
        .Left = Me.cmdPrint.Left - .Width - 20 * Screen.TwipsPerPixelX
    End With
    
    With cmdPrintSend
        .Top = Me.cmdCancel.Top
        .Left = Me.cmdReturBill.Left - .Width - 20 * Screen.TwipsPerPixelX
    End With
    
    With cmdPrintSet
        .Top = Me.cmdCancel.Top
        .Left = 300
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    Call SaveWinState(Me, App.ProductName)
    
    i = DateDiff("d", Me.DtpBegin.Value, Me.DTPEnd.Value)
    
    zlDatabase.SetPara "frmLabBarCodeBatPrint_��������Id", cboDept.ItemData(cboDept.ListIndex), 100, 1208
    zlDatabase.SetPara "frmLabBarCodeBatPrint_�걾ID", cboSample.ItemData(cboSample.ListIndex), 100, 1208
    zlDatabase.SetPara "frmLabBarCodeBatPrint_�ɼ�����", cboCapture.ItemData(cboCapture.ListIndex), 100, 1208
    zlDatabase.SetPara "frmLabBarCodeBatPrint_ִ��״̬", cboState.ItemData(cboState.ListIndex), 100, 1208
'    zlDatabase.SetPara "frmLabBarCodeBatPrint_�Ƿ���Ϊ���", chkComplete.Value, 100, 1208
    zlDatabase.SetPara "frmLabBarCodeBatPrint_���ʱ��", i, 100, 1208
End Sub

Private Sub RptItem_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim RecordC As ReportRecord
    For Each RecordC In Me.RptItem.Records
        If RecordC(mCol.����).Value = Item.Record(mCol.����).Value And Item.Record(mCol.����).Value <> "" Then
            RecordC(mCol.ѡ��).Checked = Item.Checked
        End If
    Next
End Sub

Private Sub RptItem_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim blSelect As Boolean
    Dim RepRow As ReportRow
    Dim hitColumn As ReportColumn
    With Me.RptItem
        Set hitColumn = .HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "Check" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                blSelect = Not .Records(0).Item(mCol.ѡ��).Checked
                SelectOrCancelReprotCheck .Records, mCol.ѡ��, blSelect
            End If
            .Populate
        End If
    End With
End Sub
Private Sub SelectOrCancelReprotCheck(RepObj As ReportRecords, intFiledCol As Integer, blSelect As Boolean)
    Dim Record As ReportRecord
    For Each Record In RepObj
        Record.Visible = True
        Record.Item(intFiledCol).Checked = blSelect
        If Record.Item(mCol.�����ӡ).Value = "�Ѵ�ӡ" And chkPrint.Value = 1 Then
            Record.Visible = False
        End If
    Next
    Me.RptItem.Populate
End Sub

Private Sub RptItem_SelectionChanged()
    With Me.RptItem
        If Not .FocusedRow Is Nothing And .FocusedRow.GroupRow = False Then
            .PaintManager.HighlightBackColor = Val(.FocusedRow.Record(mCol.�Թ���ɫ).Value)
            .Populate
        End If
    End With
End Sub

Public Sub ShowME(Objfrm As Object, strPrivs As String, intBarCodeFormat As Integer, intExecDept As Integer, blnNowConsumption As Boolean)
    mstrPrivs = strPrivs
    mintBarCodeFormat = intBarCodeFormat
    mintExecDept = intExecDept
    mblnNowConsumption = blnNowConsumption
    Me.Show , Objfrm
End Sub

Private Function CheckPlugIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal rsMoneyNow As ADODB.Recordset) As Boolean
'    rsNumber.Fields.Append "���", adVarChar, 20
'    rsNumber.Fields.Append "����", adVarChar, 18
'    rsNumber.Fields.Append "���ID", adBigInt
'    rsNumber.Fields.Append "��������", adVarChar, 18
'    rsNumber.Fields.Append "ִ�п���ID", adVarChar, 18
'    rsNumber.Fields.Append "������ĿID", adVarChar, 18
'    rsNumber.Fields.Append "Ӥ��", adBigInt
'    rsNumber.Fields.Append "������־", adBigInt
'    rsNumber.Fields.Append "�걾", adVarChar, 30
'    rsNumber.Fields.Append "ҽ������", adVarChar, 500
'    rsNumber.Fields.Append "�ɼ���ʽ", adVarChar, 100
'    rsNumber.Fields.Append "����ҽ��", adVarChar, 50
'    rsNumber.Fields.Append "����ʱ��", adDate
'    rsNumber.Fields.Append "������", adVarChar, 50
'    rsNumber.Fields.Append "����ʱ��", adDate
'    rsNumber.Fields.Append "��Ѫ��", adVarChar, 20
'    rsNumber.Fields.Append "�Թ�����", adVarChar, 50
'    rsNumber.Fields.Append "������Դ", adInteger
'    rsNumber.Fields.Append "ҽ��ID��", adVarChar, 500
'    rsNumber.Fields.Append "ִ�п���", adVarChar, 50
'    rsNumber.Fields.Append "Ӥ������", adVarChar, 50
'    rsNumber.Fields.Append "Ӥ���Ա�", adVarChar, 50
'    rsNumber.Fields.Append "�������", adVarChar, 50
    
    Dim blnTmp As Boolean
        On Error Resume Next
        CheckPlugIn = True
        If Not mobjZLIHISPlugIn Is Nothing Then
            blnTmp = mobjZLIHISPlugIn.LisPrintCodeBefore(lngSys, lngModual, rsMoneyNow)
            Call zlPlugInErrH(Err, "LisPrintCodeBefore")
            If Err.Number <> 0 Then
                '�ӿڳ�����,������ӡ
                blnTmp = True
            End If
        Else
            blnTmp = True
        End If
        CheckPlugIn = blnTmp
    Err.Clear: On Error GoTo 0

End Function

Private Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub BarCodeMake(intMode As Integer, bln�����ӡ As Boolean)
    '����                           д������.��û������ʱʹ��ҽ��ID��������
    '                               intMode 1=���� 2=��� 3=������� 4=�ͼ� 5=ȡ���ͼ� 6=ֻ��ӡ
    '                               bln�����ӡ  true = ��ӡ
    '                               ��������ʱ����һ������ͬ���ı걾Ϊ��λ����д��
    
    Dim lngPatientID As Long                    '����ID
    Dim lngloop As Long                         'ѭ������
    Dim intLoop As Long                         'ѭ������
    Dim strCuvetteNumber As String              '����
    Dim strUnion As String                      'ҽ��ID,���ͺ�,���� ʹ��"|"�ָ�
    Dim strNewBarCode As String                 '���ɵ�����
    Dim varAdvice As Variant                    '�ϲ���ҽ��ID
    Dim varItem As Variant                      '�ֽ��ִ��õ��ٽ�����
    Dim strSQL As String                        'SQL���
    Dim strBarCodeUnion As String               '�����ִ�
    Dim varBarCodeUnion As Variant              '�����ִ��ֽ�
    Dim i As Integer                            'ѭ������
    Dim intBaby As Integer                      'Ӥ�� >0 ��ʾӤ������
    Dim strSample As String                     '�걾
    Dim strAdviceContent As String              'ҽ������
    Dim lngConnectID As Long                    '���ID
    Dim varFilter As Variant                    '������ͬ��ҽ������
    Dim strDept As String                       'ִ�п���
    Dim str���� As String                       '����
    Dim rsNumber As ADODB.Recordset
    Dim astrSQL() As String
    Dim blnRollBak As Boolean
    Dim blnPrint As Boolean                     '�Ƿ���Ҵ�ӡ
    
    ReDim astrSQL(0)
    
    If RptItem.Records.Count = 0 Then Exit Sub

    '�ر�������ť
    Me.cmdFind.Enabled = False
    Me.cmdPrint.Enabled = False
    Me.cmdCancel.Enabled = False
    Me.cmdReturBill.Enabled = False
    On Error GoTo errH
    
    BlCancel = False
    
    zlCommFun.ShowFlash "���ڴ�ӡ����,���Ժ�...", Me
    Me.MousePointer = vbHourglass
    
    '������ӡ����
    
    InitRecordSet rsNumber
    
    With Me.RptItem
        For lngloop = 0 To .Records.Count - 1
        
            If .Records(lngloop).Item(mCol.ѡ��).Checked = True And .Records(lngloop).Visible = True Then
            
                If BlCancel = True Then Exit Sub                                    '����"ESC"ʱ�˳�
                
                Select Case intMode
                    Case 1, 3
                        MakeBarCode rsNumber, .Records(lngloop), 1, mintExecDept
                    Case 2
                        MakeBarCode rsNumber, .Records(lngloop), 3, mintExecDept
                    Case 4, 5, 6
                        MakeBarCode rsNumber, .Records(lngloop), 4, mintExecDept
                End Select
                
            End If
        Next
       
       
    End With
    
    On Error GoTo errH
    
    If rsNumber.RecordCount = 0 Then Exit Sub
    rsNumber.MoveFirst
    Select Case intMode
        Case 1, 3                                   '������������
            Do Until rsNumber.EOF
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_��������('" & rsNumber("ҽ��ID��") & "','" & rsNumber("��������") & "')"
                If intMode = 3 Then
                    'ִ�����
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',0," & IIf(rsNumber("���") & "" = "��Ѫ", 1, 0) & ")"
                End If
                rsNumber.MoveNext
            Loop
        Case 2                                      '��ɲɼ�
            Do Until rsNumber.EOF
                'ִ�����
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',0," & IIf(rsNumber("���") & "" = "��Ѫ", 1, 0) & ")"
                rsNumber.MoveNext
            Loop
        Case 4                                      '�ͼ�
            Do Until rsNumber.EOF
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_LisԤ������_�걾�ͳ�('" & rsNumber("ҽ��ID��") & "')"
                rsNumber.MoveNext
            Loop
        Case 5                                      'ȡ���ͼ�
            Do Until rsNumber.EOF
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_LisԤ������_�걾�ͳ�('" & rsNumber("ҽ��ID��") & "',1)"
                rsNumber.MoveNext
            Loop
    End Select
    
    gcnOracle.BeginTrans
    blnRollBak = True
    
    
    For intLoop = 1 To UBound(astrSQL)
        If astrSQL(intLoop) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        End If
    Next
    gcnOracle.CommitTrans
    blnRollBak = False
    
    
    If intMode = 1 Or intMode = 3 Or intMode = 2 Then
        Call WriterBarCodeToLIS(rsNumber, 3)
    End If
    '��ӡ����
    If bln�����ӡ = True Then
        blnPrint = CheckPlugIn(glngSys, glngModul, rsNumber)
        If blnPrint = True Then
            rsNumber.MoveFirst
            Do Until rsNumber.EOF
                '�������뵽PIC
                
                If mintBarCodeFormat = 1 Then
                    Bar39 Me.picBarCodePrint, 3, Nvl(rsNumber("��������")), False, True
                Else
                    Bar128 Me.picBarCodePrint, 3, Nvl(rsNumber("��������")), True
                End If
                SavePicture Me.picBarCodePrint.Image, App.path & "\BarCode.Bmp"
                '��ʼ��ӡ
                Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_1", Me, "��������=" & Nvl(rsNumber("��������")), _
                "��Ŀ=" & Replace(Nvl(rsNumber("ҽ������")), ",", " "), _
                "�������� = " & IIf(Nvl(rsNumber("����")) <> "", Nvl(rsNumber("����")) & IIf(Nvl(rsNumber("Ӥ��"), 0) = 0, "", "(Ӥ��" & Nvl(rsNumber("Ӥ��")) & ")"), "��"), _
                "�Ա� = " & IIf(Nvl(rsNumber("�Ա�")) <> "", Nvl(rsNumber("�Ա�")), "��"), _
                "���� = " & IIf(Nvl(rsNumber("����")) <> "", Nvl(rsNumber("����")), "��"), _
                "���� = " & IIf(Nvl(rsNumber("����")) <> "", Nvl(rsNumber("����")), "��"), _
                "��ʶ�� = " & IIf(Nvl(rsNumber("��ʶ��")) <> "", Nvl(rsNumber("��ʶ��")), "��"), _
                "���ڿ��� = " & IIf(Nvl(rsNumber("��������")) <> "", Nvl(rsNumber("��������")), "��"), _
                "�ɼ���ʽ = " & IIf(Nvl(rsNumber("�ɼ���ʽ")) <> "", Nvl(rsNumber("�ɼ���ʽ")), "��"), _
                "�걾 = " & IIf(Nvl(rsNumber("�걾")) <> "", Nvl(rsNumber("�걾")), "��"), _
                "ִ�п��� = " & IIf(Nvl(rsNumber("ִ�п���")) <> "", Nvl(rsNumber("ִ�п���")), "��"), _
                "����ҽ�� = " & IIf(Nvl(rsNumber("����ҽ��")) <> "", Nvl(rsNumber("����ҽ��")), "��"), _
                "����ʱ�� = " & IIf(Nvl(rsNumber("����ʱ��")) <> "", Nvl(rsNumber("����ʱ��")), "��"), _
                "������ = " & IIf(Nvl(rsNumber("������")) <> "", Nvl(rsNumber("������")), "��"), _
                "����ʱ�� = " & IIf(Nvl(rsNumber("����ʱ��")) <> "", Nvl(rsNumber("����ʱ��")), "��"), _
                "���� = " & IIf(Nvl(rsNumber("����")) <> "", Nvl(rsNumber("����")), "��"), _
                "��Ѫ�� = " & IIf(Nvl(rsNumber("��Ѫ��")) <> "", Nvl(rsNumber("��Ѫ��")), "��"), _
                "�Թ����� = " & IIf(Nvl(rsNumber("�Թ�����")) <> "", Nvl(rsNumber("�Թ�����")), "��"), _
                "���� = " & IIf(Nvl(rsNumber("������־")) <> "", Nvl(rsNumber("������־")), "��"), _
                "������Դ = " & IIf(Nvl(rsNumber("������Դ")) <> "", Nvl(rsNumber("������Դ")), "��"), _
                "����ͼ��1=" & App.path & "\BarCode.Bmp", 2)
                'ɾ������ͼ��
                Kill App.path & "\BarCode.Bmp"
                strSQL = "Zl_LisԤ������_�����ӡ('" & Replace(rsNumber("ҽ��ID��"), ",,", ",") & "')"
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
                rsNumber.MoveNext
            Loop
        End If
    End If
    
    
    
    
    
    zlCommFun.StopFlash
    Me.MousePointer = vbDefault
    '�ָ�����
    Me.cmdFind.Enabled = True
    Me.cmdPrint.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.cmdReturBill.Enabled = True
    ReadData
    Exit Sub
errH:
    If blnRollBak Then
        gcnOracle.RollbackTrans
        blnRollBak = False
    End If
    Me.cmdFind.Enabled = True
    Me.cmdPrint.Enabled = True
    Me.MousePointer = vbDefault
    zlCommFun.StopFlash
    
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog


End Sub
Private Sub SampleSend()
    '��¼�ͳ��걾ʱ��
    Dim intLoop As Integer
    Dim strIDs As String
    Dim strAllIDs As String
    Dim intRow As Integer
    Dim blnRollBak As Boolean
    Dim astrSQL() As String
    Dim astrReprot() As String
    Dim strTmp As String
    ReDim astrSQL(0)
    Dim strAdvils As String
    Dim strName As String
    Dim blnPrint As Boolean
    Dim strSQL As String
    Dim rsSampleCode As ADODB.Recordset
    
    If Me.RptItem.Rows.Count = 0 Then Exit Sub
    
    If Me.cboState.Text <> "���ͼ�" Then
        If frmLabSamplingSendInfo.ShowME(Me, strName, blnPrint) = False Then
            Exit Sub
        End If
    End If
    
    '���ɷ�������
    strSQL = "select ����ҽ������_�걾��������.NEXTVAL  from dual"
    Set rsSampleCode = zlDatabase.OpenSQLRecord(strSQL, "�걾��������", "")
    
    With Me.RptItem
    
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).GroupRow = False Then
                If .Rows(intLoop).Record(mCol.ѡ��).Checked = True Then
                    strIDs = strIDs & "," & .Rows(intLoop).Record(mCol.���ID).Value
                    If Len(strTmp) >= 3800 Then
                        strAllIDs = strAllIDs & strTmp & ";"
                        strTmp = ""
                    Else
                        If strAdvils <> "" Then
                            strTmp = strTmp & strAdvils
                            strAdvils = ""
                        End If
                    End If
                    If intRow = 5 Then
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        strIDs = Replace(Replace(strIDs, "|", ","), ";", ",")
                                           
                        astrSQL(UBound(astrSQL)) = "Zl_LisԤ������_�걾�ͳ�('" & Mid(strIDs, 2) & "'" & IIf(Me.cboState.Text = "���ͼ�", ",1", ",0") & _
                                    ",'" & strName & "','" & rsSampleCode(0) & "')"
                        strAdvils = strIDs
                        strIDs = ""
                        intRow = 0
                    End If
                    intRow = intRow + 1
                End If
            End If
        Next
        If strAdvils <> "" Then strTmp = strTmp & strAdvils
        If strIDs <> "" Then strTmp = strTmp & strIDs
        strAllIDs = strAllIDs & strTmp & ";"
        strIDs = Replace(Replace(strIDs, "|", ","), ";", ",")
    End With
    On Error GoTo errH
        
    '�����ͳ�ʱ��
    If strIDs <> "" Then
        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
        astrSQL(UBound(astrSQL)) = "Zl_LisԤ������_�걾�ͳ�('" & Mid(strIDs, 2) & "'" & IIf(Me.cboState.Text = "���ͼ�", ",1", ",0") & _
                                ",'" & strName & "','" & rsSampleCode(0) & "')"
    End If
    
    gcnOracle.BeginTrans
    blnRollBak = True
    
    For intLoop = 1 To UBound(astrSQL)
        If astrSQL(intLoop) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        End If
    Next
    gcnOracle.CommitTrans
    blnRollBak = False
    
    If strAllIDs <> "" Then
        strAllIDs = Mid(strAllIDs, 2)
    Else
        strAllIDs = 0
    End If
    astrReprot = Split(strAllIDs, ";")
    For intLoop = 0 To UBound(astrReprot)
        If astrReprot(intLoop) <> "" Then
            'д���ͼ�ʱ�䵽�������뵥��
            Call WriterSampleSendDateToLIS(astrReprot(intLoop), IIf(Me.cboState.Text = "���ͼ�", "1", "0"), strName)
        End If
    Next
    
    If Me.cboState.Text <> "���ͼ�" Then
'        If MsgBox("�Ƿ��ӡ�ͳ��嵥?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        If blnPrint = True Then
            astrReprot = Split(strAllIDs, ";")
            For intLoop = 0 To UBound(astrReprot)
                If astrReprot(intLoop) <> "" Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_3", Me, "ҽ���ִ�=" & astrReprot(intLoop), 2)
                End If
            Next
        End If
    End If
    ReadData
    Exit Sub
errH:
    If blnRollBak Then
        gcnOracle.RollbackTrans
        blnRollBak = False
    End If
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub TxtBarCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.TxtBarCode.Tag = "" Then
            Me.RptItem.Records.DeleteAll
            Me.RptItem.Populate
        End If
        Call ReadData(Me.TxtBarCode)
        Me.TxtBarCode.Text = ""
        Me.TxtBarCode.SetFocus
    End If
End Sub

Private Sub InitRecordSet(rsNumber As ADODB.Recordset)
    '��ʼ����¼��
    
    '��¼�Թܱ���
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "���", adVarChar, 20
    rsNumber.Fields.Append "����", adVarChar, 18
    rsNumber.Fields.Append "���ID", adBigInt
    rsNumber.Fields.Append "��������", adVarChar, 18
    rsNumber.Fields.Append "ִ�п���ID", adVarChar, 18
    rsNumber.Fields.Append "������ĿID", adVarChar, 18
    rsNumber.Fields.Append "Ӥ��", adBigInt
    rsNumber.Fields.Append "������־", adBigInt
    rsNumber.Fields.Append "�걾", adVarChar, 30
    rsNumber.Fields.Append "ҽ������", adVarChar, 500
    rsNumber.Fields.Append "�ɼ���ʽ", adVarChar, 100
    rsNumber.Fields.Append "����ҽ��", adVarChar, 50
    rsNumber.Fields.Append "����ʱ��", adDate
    rsNumber.Fields.Append "������", adVarChar, 50
    rsNumber.Fields.Append "����ʱ��", adDate
    rsNumber.Fields.Append "��Ѫ��", adVarChar, 20
    rsNumber.Fields.Append "�Թ�����", adVarChar, 50
    rsNumber.Fields.Append "������Դ", adInteger
    rsNumber.Fields.Append "ҽ��ID��", adVarChar, 500
    rsNumber.Fields.Append "ִ�п���", adVarChar, 50
    rsNumber.Fields.Append "����ID", adVarChar, 18
    rsNumber.Fields.Append "����", adVarChar, 50
    rsNumber.Fields.Append "�Ա�", adVarChar, 10
    rsNumber.Fields.Append "����", adVarChar, 50
    rsNumber.Fields.Append "����", adVarChar, 50
    rsNumber.Fields.Append "��ʶ��", adVarChar, 50
    rsNumber.Fields.Append "��������", adVarChar, 50
    
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
End Sub

Public Function MakeBarCode(rsNumber As ADODB.Recordset, RowRecord As ReportRecord, intMode As Integer, Optional intExecDept As Integer, Optional strBarCode As String) As Boolean
    '����                   �������벢��¼������汣�浽���ݻ��ӡ
    '����                   ���ڼ�¼�ļ�¼��
    '                       RowRecord������
    '                       'ִ�п����Ƿ�Ҫ����
    '                       Mode =0 ������ =1 �������� =2 ������� = 3 ��ɲɼ� = 4 ��ӡ������ִ��
    '                       strBarCode <> ""ʱ��ʾʹ�ð�����
    Dim strFilter As String
    Dim blnNew As Boolean
    Dim strҽ������ As String
    
    blnNew = False
    Select Case intMode
        Case 0                              '��
            If rsNumber.RecordCount = 0 Then blnNew = True
        Case 1                              '����
            strFilter = "����ID=" & RowRecord.Item(mCol.����ID).Value & " And ������ĿID=" & Val(RowRecord.Item(mCol.������ĿID).Value)
            rsNumber.filter = strFilter
            If rsNumber.EOF = False Then
                '��������Ŀ��ͬʱ����һ������
                blnNew = True
            Else
                strFilter = "����ID=" & RowRecord.Item(mCol.����ID).Value & _
                      " And ����='" & RowRecord.Item(mCol.����).Value & _
                      "' And Ӥ��=" & RowRecord.Item(mCol.Ӥ��).Value & _
                      " And ������־=" & IIf(RowRecord.Item(mCol.����).Value = "����", 1, 0) & _
                      " And �걾='" & RowRecord.Item(mCol.�걾).Value & "'"
                If intExecDept = 1 Then strFilter = strFilter & " And ִ�п���id=" & RowRecord.Item(mCol.����ִ�п���ID).Value
                rsNumber.filter = strFilter
                If rsNumber.EOF = True Then
                    '����������
                    blnNew = True
                End If
            End If
        Case 2                              'ȡ������
            If rsNumber.RecordCount = 0 Then blnNew = True
        
        Case 3, 4                           '���������ӡ
            strFilter = "��������='" & RowRecord.Item(mCol.����).Value & "'"
            rsNumber.filter = strFilter
            If rsNumber.EOF = True Then
                blnNew = True
            End If
    End Select
    If blnNew = True Then
        rsNumber.AddNew
        rsNumber!��� = RowRecord.Item(mCol.���).Value
        '�󶨺���������
        If strBarCode <> "" Then
            rsNumber!�������� = strBarCode
        Else
            If intMode = 3 Or intMode = 4 Then
                rsNumber!�������� = RowRecord.Item(mCol.����).Value
            Else
                rsNumber!�������� = zlDatabase.GetNextNo(125, Split(RowRecord.Item(mCol.ҽ��id).Value, ",")(0))
            End If
        End If
        rsNumber!�ɼ���ʽ = RowRecord.Item(mCol.�ɼ���ʽ).Value
        rsNumber!�걾 = RowRecord.Item(mCol.�걾).Value
        rsNumber!ִ�п���ID = RowRecord.Item(mCol.����ִ�п���ID).Value
        rsNumber!����ҽ�� = RowRecord.Item(mCol.����ҽ��).Value
        rsNumber!����ʱ�� = RowRecord.Item(mCol.����ʱ��).Value
        rsNumber!������ = RowRecord.Item(mCol.������).Value
        If RowRecord.Item(mCol.����ʱ��).Value <> "" Then
            rsNumber!����ʱ�� = RowRecord.Item(mCol.����ʱ��).Value
        End If
        rsNumber!���� = RowRecord.Item(mCol.����).Value
        rsNumber!��Ѫ�� = RowRecord.Item(mCol.��Ѫ��).Value
        rsNumber!�Թ����� = RowRecord.Item(mCol.�Թ�����).Value
        rsNumber!������־ = IIf(RowRecord.Item(mCol.����).Value = "����", 1, 0)
        rsNumber!������Դ = RowRecord.Item(mCol.������Դ).Value
        rsNumber!Ӥ�� = RowRecord.Item(mCol.Ӥ��).Value
        rsNumber!ִ�п��� = RowRecord.Item(mCol.ִ�п���).Value
        rsNumber!ҽ������ = RowRecord.Item(mCol.����).Value
        rsNumber!���� = RowRecord.Item(mCol.����).Value
        rsNumber!�Ա� = RowRecord.Item(mCol.�Ա�).Value
        rsNumber!���� = RowRecord.Item(mCol.����).Value
        rsNumber!���� = RowRecord.Item(mCol.����).Value
        rsNumber!��ʶ�� = RowRecord.Item(mCol.��ʶ��).Value
        rsNumber!�������� = RowRecord.Item(mCol.��������).Value
        rsNumber!����ID = RowRecord.Item(mCol.����ID).Value
        rsNumber!������ĿID = Val(RowRecord.Item(mCol.������ĿID).Value)
        rsNumber!ҽ��ID�� = Replace(Replace(RowRecord.Item(mCol.ҽ��id).Value & "," & _
                            RowRecord.Item(mCol.���ID).Value & "," & RowRecord.Item(mCol.�ϲ�ҽ��).Value, ";", ","), ",,", ",")
        rsNumber.Update
    Else
        If rsNumber.RecordCount > 0 Then
            rsNumber.MoveLast
            strҽ������ = IIf(Trim(RowRecord.Item(mCol.����).Value) = "", RowRecord.Item(mCol.ҽ������).Value, RowRecord.Item(mCol.����).Value)
            If InStr(";" & rsNumber!ҽ������ & ";", ";" & strҽ������ & ";") <= 0 Then
                rsNumber!ҽ������ = rsNumber!ҽ������ & ";" & strҽ������
                
                
            End If
            rsNumber!ҽ��ID�� = Replace(rsNumber!ҽ��ID�� & "," & Replace(RowRecord.Item(mCol.ҽ��id).Value & "," & _
                            RowRecord.Item(mCol.���ID).Value & "," & RowRecord.Item(mCol.�ϲ�ҽ��).Value, ";", ","), ",,", ",")
            rsNumber.Update
        End If
        
    End If
    rsNumber.filter = ""
End Function

Private Sub txtUnit_GotFocus()
    Me.txtUnit.SelStart = 0
    Me.txtUnit.SelLength = Len(Me.txtUnit)
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    Dim objPoint As POINTAPI
    Dim sglX As Single, sglY As Single
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    
    If KeyAscii = 13 Then
        
        If Len(Me.txtUnit.Text) < 2 Then
            MsgBox "��������1λ���ϵĵ�λ���Ʋ��ܲ�ѯ", vbInformation, "��ʾ"
            Me.txtUnit.SetFocus
            Exit Sub
        End If
        strSQL = "select /*+ rule */ distinct ����id id,������λ from ������Ϣ where �Ǽ�ʱ�� <= sysdate - (365/2) and ������λ like [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "%" & Me.txtUnit.Text & "%")
        Call ClientToScreen(txtUnit.hWnd, objPoint)
        sglX = objPoint.X * 15 - 30
        sglY = objPoint.Y * 15 + txtUnit.Height
        If frmSelectList.ShowSelect(Me, rsTmp, "������λ,3000,0,0", sglX, sglY, txtUnit.Width, 2000, Me.Name, "��ѡ���Թ�����λ") Then
            Me.txtUnit = rsTmp!������λ
            Me.txtUnit.SelStart = 0
            Me.txtUnit.SelLength = Len(Me.txtUnit)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



