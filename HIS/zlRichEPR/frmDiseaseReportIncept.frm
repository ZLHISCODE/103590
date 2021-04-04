VERSION 5.00
Begin VB.Form frmDiseaseReportIncept 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����������"
   ClientHeight    =   3900
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5655
   Icon            =   "frmDiseaseReportIncept.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3195
      TabIndex        =   10
      Top             =   3360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4305
      TabIndex        =   11
      Top             =   3360
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -30
      TabIndex        =   9
      Top             =   3195
      Width           =   6030
   End
   Begin VB.TextBox txtComment 
      Height          =   660
      Left            =   795
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2370
      Width           =   4605
   End
   Begin VB.OptionButton optIncept 
      Caption         =   "�ܾ�����(&R)"
      Height          =   225
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   1305
   End
   Begin VB.OptionButton optIncept 
      Caption         =   "ͬ�����(&A)"
      Height          =   225
      Index           =   0
      Left            =   780
      TabIndex        =   5
      Top             =   1800
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -15
      TabIndex        =   0
      Top             =   600
      Width           =   6030
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "ͬ���ܾ�������(&M):"
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   2130
      Width           =   1800
   End
   Begin VB.Label lblWriter 
      AutoSize        =   -1  'True
      Caption         =   "���:"
      Height          =   180
      Left            =   780
      TabIndex        =   4
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      Height          =   180
      Left            =   780
      TabIndex        =   3
      Top             =   1065
      Width           =   450
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      Height          =   180
      Left            =   780
      TabIndex        =   2
      Top             =   780
      Width           =   450
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   150
      Picture         =   "frmDiseaseReportIncept.frx":038A
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ϸ����ٴ�ҽ����д�ļ������������Ƿ����Ҫ�󣬾������ջ�ܾ��ü������档"
      Height          =   360
      Left            =   780
      TabIndex        =   1
      Top             =   135
      Width           =   4680
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDiseaseReportIncept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOk As Boolean
Private mstrId As String
Private mlngPatiId As Long, mlngPageId As Long, mlngFrom As Long

Public Function ShowMe(ByVal frmParent As Object, strRecordId As String, strInfo As String) As Boolean
'strInfo=����|����|����|�Ա�|����|�����|���|�ʱ��|����ID|��ҳID
'�²��� strInfo=strInfo & |���ഫȾ��|���ഫȾ��|���ഫȾ��|��������|��������2
    mstrId = strRecordId
    
    Err = 0: On Error GoTo errHand

    Me.lblFile.Caption = "����:" & Split(strInfo, "|")(0) & "    ����:" & Split(strInfo, "|")(1)
    Me.lblPati.Caption = "����: " & Split(strInfo, "|")(2) & "," & Split(strInfo, "|")(3) & "," & Split(strInfo, "|")(4) & "  " & Split(strInfo, "|")(5)
    Me.lblWriter.Caption = "���:" & Split(strInfo, "|")(6) & "    �ʱ��:" & Split(strInfo, "|")(7)
    Me.lblFile.Tag = Split(strInfo, "|")(1)
    mlngPatiId = Split(strInfo, "|")(8): mlngPageId = Split(strInfo, "|")(9)
    If InStr(Split(strInfo, "|")(1), "����:") > 0 Then
        mlngFrom = 1
    ElseIf InStr(Split(strInfo, "|")(1), "סԺ:") > 0 Then
        mlngFrom = 2
    End If
    
    If Not IsNumeric(mstrId) Then
        lblComment.Tag = Split(strInfo, "|")(10) & ";" & Split(strInfo, "|")(11) & ";" & Split(strInfo, "|")(12)
        txtComment.Tag = Split(strInfo, "|")(13) & ";" & Split(strInfo, "|")(14)
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

Private Sub cmdOK_Click()
    On Error GoTo errHand
    Dim lngKey As Long, rsTemp As ADODB.Recordset, strContent As String, dblDocID As Double
    
    If LenB(StrConv(Trim(Me.txtComment.Text), vbFromUnicode)) > Me.txtComment.MaxLength Then
        MsgBox "˵�����������" & Me.txtComment.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName
        Me.txtComment.SetFocus: Exit Sub
    End If
    
    If IsNumeric(mstrId) Then
        strContent = ""
        gstrSQL = "Zl_�����걨��¼_Incept(" & CDbl(mstrId) & "," & IIf(Me.optIncept(0).Value, 1, -1) & ",'" & Trim(Me.txtComment.Text) & "','" & mstrId & "'," & mlngPatiId & "," & mlngPageId & "," & mlngFrom & ",'')"
    Else
        '�°没������GUIDת���ݣ���ȷ�� �����걨��¼PK
        gstrSQL = "Select �ļ�ID From �����걨��¼ Where �ĵ�ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", mstrId)
        If rsTemp.EOF Then
            dblDocID = gobjEmr.GuidToHashCode(mstrId)
        Else
            dblDocID = rsTemp!�ļ�ID
        End If
        '�°没������ȡ�걨��Ϣ
        strContent = lblComment.Tag & "|" & txtComment.Tag
        gstrSQL = "Zl_�����걨��¼_Incept(" & dblDocID & "," & IIf(Me.optIncept(0).Value, 1, -1) & ",'" & Trim(Me.txtComment.Text) & "','" & mstrId & "'," & mlngPatiId & "," & mlngPageId & "," & mlngFrom & ",'" & strContent & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If InStr(lblFile.Tag, "����") = 0 Then
        If optIncept(1).Value = True Then  '���ܾ�ʱ������һ�γ������¼
            lngKey = zlDatabase.GetNextId("����������¼")
            gstrSQL = "zl_����������¼_Update(" & lngKey & ",Null,Null," & mlngPatiId & "," & mlngPageId & ",7,'" & mstrId & "','" & _
                    txtComment.Text & "',Null,'" & gstrUserName & "',To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & _
                    ",To_Date('" & Format(zlDatabase.Currentdate + 1, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else '����ʱ,�����������ܾ����ļ�¼����ɱ�־
            gstrSQL = "Select ID From ����������¼ Where ����ID=[1] And ��ҳID=[2] and ��������=7 And �ļ�ID=[3] And ������=[4]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mstrId, gstrUserName)
            Do Until rsTemp.EOF
                gstrSQL = "zl_����������¼_Finish(" & rsTemp!ID & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                rsTemp.MoveNext
            Loop
        End If
    End If
    
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



