VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiseaseReportSend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����걨�Ǽ�"
   ClientHeight    =   4860
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5250
   FillColor       =   &H000000C0&
   Icon            =   "frmDiseaseReportSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboRecd 
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   705
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4395
   End
   Begin VB.ComboBox cboUnit 
      Height          =   300
      Left            =   1845
      TabIndex        =   10
      Top             =   2655
      Width           =   3105
   End
   Begin VB.TextBox txtPerson 
      Height          =   300
      Left            =   1845
      TabIndex        =   6
      Top             =   1830
      Width           =   1830
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2745
      TabIndex        =   14
      Top             =   4290
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3855
      TabIndex        =   15
      Top             =   4290
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -30
      TabIndex        =   13
      Top             =   4140
      Width           =   6030
   End
   Begin VB.TextBox txtComment 
      Height          =   660
      Left            =   795
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3345
      Width           =   4140
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -90
      TabIndex        =   0
      Top             =   570
      Width           =   6030
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   300
      Left            =   1845
      TabIndex        =   8
      Top             =   2250
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   89194499
      CurrentDate     =   39668.3389814815
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "���͵�λ(&U)"
      Height          =   180
      Left            =   780
      TabIndex        =   9
      Top             =   2715
      Width           =   990
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��(&T)"
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   2310
      Width           =   990
   End
   Begin VB.Label lblPerson 
      AutoSize        =   -1  'True
      Caption         =   "������Ա(&P)"
      Height          =   180
      Left            =   780
      TabIndex        =   5
      Top             =   1890
      Width           =   990
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "���ͱ�ע(&M):"
      Height          =   180
      Left            =   780
      TabIndex        =   11
      Top             =   3105
      Width           =   1080
   End
   Begin VB.Label lblWriter 
      AutoSize        =   -1  'True
      Caption         =   "���:"
      Height          =   180
      Left            =   780
      TabIndex        =   4
      Top             =   1335
      Width           =   630
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      Height          =   180
      Left            =   780
      TabIndex        =   3
      Top             =   1050
      Width           =   450
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      Height          =   180
      Left            =   780
      TabIndex        =   2
      Top             =   765
      Width           =   450
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   150
      Picture         =   "frmDiseaseReportSend.frx":038A
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Լ�������ı���������еǼǣ��Ա������ġ�"
      Height          =   180
      Left            =   825
      TabIndex        =   1
      Top             =   120
      Width           =   4680
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDiseaseReportSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean
Private mstrId As String

Public Function ShowMe(ByVal frmParent As Object, strRecordId As String, strInfo As String, strDiagInfo1 As String, strDiagInfo2 As String) As Boolean
'strInfo=����|����|����|�Ա�|����|�����|���|�ʱ��|����ID|��ҳID|�ļ�ID
Dim rsTemp As New ADODB.Recordset, rsDocID As New ADODB.Recordset
Dim strNote As String, strDiag1 As String, strDiag2 As String, strReturn As String
Dim strFileID As String
    mstrId = strRecordId
    
    Err = 0: On Error GoTo errHand
    
    Me.lblFile.Caption = "����:" & Split(strInfo, "|")(0) & "    ����:" & Split(strInfo, "|")(1)
    Me.lblPati.Caption = "����: " & Split(strInfo, "|")(2) & "," & Split(strInfo, "|")(3) & "," & Split(strInfo, "|")(4) & "  " & Split(strInfo, "|")(5)
    Me.lblWriter.Caption = "���:" & Split(strInfo, "|")(6) & "    �ʱ��:" & Split(strInfo, "|")(7)
    
    Me.dtpTime.MinDate = Format(Split(strInfo, "|")(7), "yyyy-MM-dd HH:mm")
    Me.dtpTime.MaxDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    Me.dtpTime.Value = Me.dtpTime.MaxDate
    Me.txtPerson.Text = gstrUserName
    gstrSQL = "Select ���� From �������͵�λ Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.cboUnit.Clear
        Do While Not .EOF
            Me.cboUnit.AddItem !����
            .MoveNext
        Loop
    End With
    strFileID = Split(strInfo, "|")(10)
    strDiag1 = IIf(strDiagInfo1 = "", "", " and l.�������1 in ( '" & strDiagInfo1 & "', '" & strDiagInfo2 & "')")
    strDiag2 = IIf(strDiagInfo2 = "", "", " and l.�������2 in ( '" & strDiagInfo1 & "', '" & strDiagInfo2 & "')")
    If IsNumeric(strFileID) Then
        gstrSQL = "Select l.������, l.����ʱ��, l.�������1, l.�������2  " & vbNewLine & _
                    " From (Select s.������, s.����ʱ��, s.�������1, s.�������2,l.����id,l.����id,l.�ļ�id " & vbNewLine & _
                    " From ���Ӳ�����¼ L, �����걨��¼ S " & vbNewLine & _
                    " Where l.Id = s.�ļ�id(+) And l.�������� = 5 And l.�ļ�id='" & Split(strInfo, "|")(10) & "' " & vbNewLine & _
                    " And  s.����ʱ��>=trunc(sysdate,'yyyy') " & vbNewLine & _
                    " and s.����ʱ��<add_months(trunc(sysdate,'YYYY'),12)) L,������Ϣ P, ���ű� D " & vbNewLine & _
                    " Where l.����ID = p.����ID And l.����ID = D.ID And p.����ID ='" & Split(strInfo, "|")(8) & "' "
        gstrSQL = gstrSQL & strDiag1 & strDiag2
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsTemp
            If rsTemp.RecordCount > 0 Then
                Do While Not .EOF
                    strNote = !������ & !����ʱ�� & "�ѱ���" & NVL(strDiagInfo1, "") & NVL(strDiagInfo2, "") & vbCrLf
                    Me.lblNote.ForeColor = vbRed
                    Me.cboRecd.AddItem strNote
                    .MoveNext
                Loop
                Me.cboRecd.Visible = True
                Me.cboRecd.ListIndex = 0
                Me.cboRecd.ForeColor = vbRed
            End If
        End With
    Else
        gstrSQL = " Select rawtohex(m.Id) Docid" & vbNewLine & _
                  " From Bz_Doc_Log M, Bz_Act_Log N, Bz_Master_Codes P" & vbNewLine & _
                  " Where n.Id = m.Actlog_Id And n.Master_Id = p.Master_Id And p.Code =:bzid And m.Status >= 2 And" & vbNewLine & _
                  " m.Antetype_Id = hextoraw('" & strFileID & "')"
        strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, Split(strInfo, "|")(8) & "^16^bzid", rsDocID)
        
        If strReturn = "" And rsDocID.RecordCount > 0 Then
            Do While Not rsDocID.EOF
               gstrSQL = "Select l.������, l.����ʱ��, l.�������1, l.�������2 From �����걨��¼ l where l.�ĵ�id='" & rsDocID!docid & "'" & _
                         " and l.����ʱ��>=trunc(sysdate,'yyyy') and l.����ʱ��<add_months(trunc(sysdate,'YYYY'),12) "
               gstrSQL = gstrSQL & strDiag1 & strDiag2
               Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
               With rsTemp
                    If rsTemp.RecordCount > 0 Then
                        Do While Not .EOF
                            strNote = !������ & !����ʱ�� & "�ѱ���" & NVL(strDiagInfo1, "") & NVL(strDiagInfo2, "") & vbCrLf
                            Me.lblNote.ForeColor = vbRed
                            Me.cboRecd.AddItem strNote
                            .MoveNext
                        Loop
                        Me.cboRecd.Visible = True
                        Me.cboRecd.ListIndex = 0
                        Me.cboRecd.ForeColor = vbRed
                        Exit Do
                    End If
               End With
              rsDocID.MoveNext
            Loop
        End If
    End If
    Me.Show vbModal, frmParent
    
    ShowMe = mblnOk
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mblnOk = False: Me.Hide
End Sub

Private Sub cmdOk_Click()
    If Trim(Me.txtPerson.Text) = "" Then MsgBox "������д������Ա��", vbInformation, gstrSysName: Me.txtPerson.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txtPerson.Text), vbFromUnicode)) > 20 Then
        MsgBox "��Ա���������20���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName
        Me.txtPerson.SetFocus: Exit Sub
    End If
    
    If Trim(Me.cboUnit.Text) = "" Then MsgBox "������д���͵�λ��", vbInformation, gstrSysName: Me.cboUnit.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.cboUnit.Text), vbFromUnicode)) > 30 Then
        MsgBox "��λ���������30���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName
        Me.cboUnit.SetFocus: Exit Sub
    End If
    
    If LenB(StrConv(Trim(Me.txtComment.Text), vbFromUnicode)) > Me.txtComment.MaxLength Then
        MsgBox "��ע���������" & Me.txtComment.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName
        Me.txtComment.SetFocus: Exit Sub
    End If
    
    gstrSQL = "Zl_�����걨��¼_Send('" & mstrId & "'," & _
            "'" & Trim(Me.txtPerson.Text) & "'," & _
            "To_Date('" & Format(Me.dtpTime.Value, "yyyy-MM-dd HH:mm") & "','yyyy-mm-dd hh24:mi')," & _
            "'" & Trim(Me.cboUnit.Text) & "'," & _
            "'" & Trim(Me.txtComment.Text) & "')"
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOk = True: Me.Hide: Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optIncept_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtComment_GotFocus()
    Me.txtComment.SelStart = 0: Me.txtComment.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtComment_LostFocus()
    Me.txtComment.Text = Replace(Me.txtComment, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPerson_GotFocus()
    Me.txtPerson.SelStart = 0: Me.txtPerson.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub





