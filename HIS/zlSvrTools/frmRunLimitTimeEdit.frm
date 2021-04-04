VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRunLimitTimeEdit 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "��ʱʱ�䰲��"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picMark 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   80
      Index           =   0
      Left            =   885
      ScaleHeight     =   75
      ScaleWidth      =   600
      TabIndex        =   7
      Top             =   585
      Visible         =   0   'False
      Width           =   600
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   300
      Index           =   0
      Left            =   1050
      TabIndex        =   5
      Top             =   1095
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      _Version        =   393216
      Format          =   105119746
      CurrentDate     =   36494
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   5670
      TabIndex        =   1
      Top             =   1095
      Width           =   900
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   300
      Left            =   4665
      TabIndex        =   0
      Top             =   1095
      Width           =   900
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   300
      Index           =   1
      Left            =   3180
      TabIndex        =   6
      Top             =   1095
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      _Version        =   393216
      Format          =   105119746
      CurrentDate     =   36494
   End
   Begin VB.PictureBox picTime 
      BackColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   225
      ScaleHeight     =   75
      ScaleWidth      =   6285
      TabIndex        =   2
      Top             =   555
      Width           =   6345
   End
   Begin VB.Image imgRuler 
      Height          =   375
      Left            =   180
      Picture         =   "frmRunLimitTimeEdit.frx":0000
      Top             =   165
      Width           =   6450
   End
   Begin VB.Image imgCursorButtom 
      Height          =   240
      Index           =   1
      Left            =   3495
      Picture         =   "frmRunLimitTimeEdit.frx":444A
      Top             =   645
      Width           =   240
   End
   Begin VB.Image imgCursorButtom 
      Height          =   240
      Index           =   0
      Left            =   975
      Picture         =   "frmRunLimitTimeEdit.frx":4E4C
      Top             =   645
      Width           =   240
   End
   Begin VB.Label lblStop 
      AutoSize        =   -1  'True
      Caption         =   "��ֹʱ��"
      Height          =   180
      Left            =   2325
      TabIndex        =   4
      Top             =   1155
      Width           =   720
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      Caption         =   "��ʼʱ��"
      Height          =   180
      Left            =   210
      TabIndex        =   3
      Top             =   1155
      Width           =   720
   End
End
Attribute VB_Name = "frmRunLimitTimeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngRow As Long
Private mstrTimeStart As String, mstrTimeStop As String
Private mlngId As Long, mlngPlanNo As Long
Private msinRulerWidth As Single
Private X1 As Single '���ڼ�¼�ƶ����ʱ�������λ��
Private mblnOk As Boolean
Private mlngDayTime As Long

Public Function ShowMe(ByVal lngID As Long, ByVal lngPlanNo As Long, ByVal lngRow As Long, _
                    ByVal strTimeStart As String, ByVal strTimeStop As String) As Boolean
    mlngId = lngID
    mlngPlanNo = lngPlanNo
    mlngRow = lngRow
    mstrTimeStart = strTimeStart
    mstrTimeStop = strTimeStop
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'��ʱ����Ϣ���µ���Ӧ�����Ķ�Ӧ������
    On Error GoTo errH
    mblnOk = False
    If mlngId = 0 Then
        '����
        Call ExecuteProcedure("Zl_ZlRunLimitTime_Update(0,0," & mlngPlanNo & "," & mlngRow - 1 & _
                                    ", to_date('" & "1899-12-30 " & dtpTime(0).value & "','YYYY-MM-DD HH24:MI:SS'), to_date('" & _
                                    "1899-12-30 " & dtpTime(1).value & "','YYYY-MM-DD HH24:MI:SS'))", "����ʱ���")
    Else
        '�޸�
        Call ExecuteProcedure("Zl_ZlRunLimitTime_Update(1," & mlngId & "," & mlngPlanNo & "," & mlngRow - 1 & _
                                    ", to_date('" & "1899-12-30 " & dtpTime(0).value & "','YYYY-MM-DD HH24:MI:SS'), to_date('" & _
                                    "1899-12-30 " & dtpTime(1).value & "','YYYY-MM-DD HH24:MI:SS'))", "�޸�ʱ���")
    End If
    mblnOk = True
    Unload Me
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub dtpTime_Change(Index As Integer)
    If dtpTime(0).value > dtpTime(1).value And Index = 0 Then dtpTime(0).value = dtpTime(1).value
    If dtpTime(1).value < dtpTime(0).value And Index = 1 Then dtpTime(1).value = dtpTime(0).value
    Call SetPosition(0, 0, dtpTime(0).value)
    Call SetPosition(1, 1, dtpTime(1).value)
End Sub

Private Sub Form_Load()
    Call FillData
End Sub

'���Ԥ�����ݣ��Լ���ģ��Ϊ���޸ġ�ʱ����ʼ����������
Private Sub FillData()
    Dim strStartTime() As String, strStopTime() As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim lngMarkLeft As Long, lngMarkRight As Long
    
    mlngDayTime = CLng(24 * 60) * 60
    msinRulerWidth = picTime.Width - 50
    
    '���Ѿ��е�ʱ������������ϱ��Ϊ��ɫ
    strSql = "Select To_Char(��ʼʱ��, 'HH24:MI:SS') ��ʼʱ��, To_Char(����ʱ��, 'HH24:MI:SS') ����ʱ��" & vbNewLine & _
            "From ZlRunLimitTime" & vbNewLine & _
            "Where ���� = [1] And ���� = [2]" & vbNewLine & _
            "Order By ��ʼʱ��"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "��ȡһ��������ĳһ���ڵ�ʱ����Ϣ", mlngPlanNo, mlngRow - 1)
    With rsTemp
        For i = 1 To .RecordCount
            Load picMark(i)
            strStartTime = Split(!��ʼʱ��, ":")
            strStopTime = Split(!����ʱ��, ":")
            lngMarkLeft = picTime.Left + msinRulerWidth / mlngDayTime * (strStartTime(0) * 60 * 60 + strStartTime(1) * 60 + strStartTime(2))
            lngMarkRight = picTime.Left + msinRulerWidth / mlngDayTime * (strStopTime(0) * 60 * 60 + strStopTime(1) * 60 + strStopTime(2))
            picMark(i).Top = picMark(0).Top
            picMark(i).Left = lngMarkLeft
            picMark(i).Width = lngMarkRight - lngMarkLeft
            picMark(i).Visible = True
            picMark(i).ZOrder
            .MoveNext
        Next
    End With
    
    '��ʼ����������
    If mstrTimeStart <> "" Then
        Me.Caption = "�޸�ʱ���"
        strStartTime = Split(mstrTimeStart, ":")
        strStopTime = Split(mstrTimeStop, ":")
        imgCursorButtom(0).Left = picTime.Left + msinRulerWidth / mlngDayTime * (strStartTime(0) * 60 * 60 + strStartTime(1) * 60 + strStartTime(2)) - 100
        imgCursorButtom(1).Left = picTime.Left + msinRulerWidth / mlngDayTime * (strStopTime(0) * 60 * 60 + strStopTime(1) * 60 + strStopTime(2)) - 100
        dtpTime(0).value = CDate(mstrTimeStart)
        dtpTime(1).value = CDate(mstrTimeStop)
    Else
        Me.Caption = "����ʱ���"
        imgCursorButtom(0).Left = picTime.Left + msinRulerWidth / mlngDayTime * 8 * 60 * 60 - 100
        imgCursorButtom(1).Left = picTime.Left + msinRulerWidth / mlngDayTime * CLng(12 * 60) * 60 - 100
        dtpTime(0).value = CDate("8:00:00")
        dtpTime(1).value = CDate("12:00:00")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    X1 = 0
End Sub

Private Sub SetTime(ByVal intCursor As Integer, ByVal intTime As Integer, ByVal X As Single)
    '���ݱ�����α��λ�ü����ʱ�䲢��ʾ��ʱ��ؼ���
    'intCursor:�α�ؼ�������
    'intTime:���ڿؼ�������
    'x:�α��λ��
    
    Dim lngTime As Long
    Dim lngSecond As Long, lngMinute As Long, lngHour As Long
    Dim sinPosition As Single
    
    sinPosition = imgCursorButtom(intCursor).Left
    '�������ƶ�����ʱ�䣬����9:00,9:30��ʱ��ʱ����һ�ִ���Ч��
    If Abs(X - X1) <= 30 And X <> X1 Then
        lngTime = (sinPosition - picTime.Left + 100) * mlngDayTime / msinRulerWidth
        lngSecond = lngTime Mod 60
        lngMinute = (lngTime - lngSecond) / 60 Mod 60
        lngHour = Int((lngTime - lngSecond) / 60 / 60)
        
        If X - X1 > 0 Then
            If lngMinute > 20 And lngMinute < 30 Then
                Call SetPosition(intCursor, intTime, lngHour & ":30:00")
                Exit Sub
            ElseIf lngMinute > 50 And lngMinute <= 59 Then
                If lngHour = 23 Then
                    Call SetPosition(intCursor, intTime, "23:59:59")
                Else
                    Call SetPosition(intCursor, intTime, lngHour + 1 & ":00:00")
                End If
                Exit Sub
            End If
        Else
            If lngMinute < 40 And lngMinute > 30 Then
                Call SetPosition(intCursor, intTime, lngHour & ":30:00")
                Exit Sub
            ElseIf lngMinute > 0 And lngMinute < 10 Then
                Call SetPosition(intCursor, intTime, lngHour & ":00:00")
                Exit Sub
            End If
        End If
    End If

    imgCursorButtom(intCursor).Left = imgCursorButtom(intCursor).Left + X - X1
    
    lngTime = (imgCursorButtom(intCursor).Left - picTime.Left + 100) * mlngDayTime / msinRulerWidth
    If lngTime < 0 Then lngTime = 0
    lngSecond = lngTime Mod 60
    lngMinute = (lngTime - lngSecond) / 60 Mod 60
    lngHour = Int((lngTime - lngSecond) / 60 / 60)
    If lngHour >= 24 Then
        lngHour = 23
        lngMinute = 59
        lngSecond = 59
    End If
    If lngSecond <> 0 Then
        If Not (lngHour = 23 And lngMinute = 59 And lngSecond = 59) Then lngSecond = 0
        dtpTime(intTime).value = CDate(lngHour & ":" & lngMinute & ":" & lngSecond)
        Call SetPosition(intCursor, intTime)
    Else
        dtpTime(intTime).value = CDate(lngHour & ":" & lngMinute & ":" & lngSecond)
    End If
End Sub

Private Sub SetPosition(ByVal intCursor As Integer, ByVal intTime As Integer, Optional ByVal StrDate As String)
    '����ʱ��ؼ��ϵ�ʱ�������α��ڱ���ϵ�λ��
    'intCursor:�α�ؼ�������
    'intTime:���ڿؼ�������
    'strDate:����
    If StrDate <> "" Then dtpTime(intTime).value = CDate(StrDate)
    imgCursorButtom(intCursor).Left = picTime.Left + msinRulerWidth / mlngDayTime * (dtpTime(intTime).Hour * 60 * 60 + dtpTime(intTime).Minute * 60 + dtpTime(intTime).Second) - 100
End Sub

Private Sub imgCursorButtom_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    X1 = X
    If imgCursorButtom(Index).Left <= imgCursorButtom(Abs(Index - 1)).Left Then
        imgCursorButtom(Index).Tag = dtpTime(0).value
        imgCursorButtom(Abs(Index - 1)).Tag = dtpTime(1).value
    Else
        imgCursorButtom(Index).Tag = dtpTime(1).value
        imgCursorButtom(Abs(Index - 1)).Tag = dtpTime(0).value
    End If
End Sub

Private Sub imgCursorButtom_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCursor As Integer, intTime As Integer
    
    If Button = 1 Then
        If imgCursorButtom(Index).Left >= picTime.Left + msinRulerWidth - 100 And X - X1 > 0 Then Exit Sub
        If imgCursorButtom(Index).Left <= picTime.Left - 100 And X - X1 < 0 Then Exit Sub
        
        If imgCursorButtom(Index).Left <= imgCursorButtom(Abs(Index - 1)).Left Then
            Call SetTime(Index, 0, X)
            dtpTime(1).value = CDate(imgCursorButtom(Abs(Index - 1)).Tag)
        Else
            Call SetTime(Index, 1, X)
            dtpTime(0).value = CDate(imgCursorButtom(Abs(Index - 1)).Tag)
        End If
    End If
End Sub

Private Sub imgCursorButtom_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If imgCursorButtom(Index).Left > picTime.Left + msinRulerWidth - 100 Then imgCursorButtom(Index).Left = picTime.Left + msinRulerWidth - 100
        If imgCursorButtom(Index).Left < picTime.Left - 100 Then imgCursorButtom(Index).Left = picTime.Left - 100
    End If
End Sub
