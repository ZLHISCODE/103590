VERSION 5.00
Begin VB.Form frmFormatChoose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѡ����Ҫ��ӡ�ĸ�ʽ"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2700
   Icon            =   "frmFormatChoose.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�˳�ѡ��"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ  ��"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.ListBox lstMain 
      Height          =   3420
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmFormatChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrint As String '  ��Ҫ��ӡ�ĸ�ʽ "1", "1,3"  "1,4,5"  "2,4"  ÿ�����ִ���һ����ʽ

Private Sub cmdCancel_Click()
On Error GoTo errH
    mstrPrint = lstMain.ItemData(0)
    Unload Me
    Exit Sub
errH:
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

Private Sub cmdOk_Click()
On Error GoTo errH
    Dim i As Integer

    mstrPrint = ""
    For i = 0 To lstMain.ListCount - 1
        If lstMain.Selected(i) Then
            If mstrPrint <> "" Then mstrPrint = mstrPrint & ","
            mstrPrint = mstrPrint & lstMain.ItemData(i)
        End If
    Next
    
    Unload Me
    Exit Sub
errH:
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

Public Function ZLShow(ByVal strBill As String, objOwner As Form) As String
On Error GoTo errH
    Call InitList(strBill)
    
    Call Me.Move(objOwner.Left + (objOwner.Width - Me.Width) / 2, objOwner.Top + (objOwner.Height - Me.Height) / 2)
    Call Me.Show(1, objOwner)
    ZLShow = mstrPrint
    Exit Function
errH:
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Function

Private Sub InitList(ByVal strBill As String)
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Integer
    
    strSQL = "Select a.���,b.���,b.˵�� From zlreports a,zlrptfmts b Where a.Id=b.����ID And a.���=[1] Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Զ��屨���ʽ", strBill)
    
    If rsTemp.EOF Then
        mstrPrint = "0"
        Unload Me
        Exit Sub
    End If
    
    For i = 1 To rsTemp.RecordCount
        lstMain.list(i - 1) = NVL(rsTemp!˵��)
        lstMain.ItemData(i - 1) = Val(NVL(rsTemp!���))
        Call rsTemp.MoveNext
    Next
   
    lstMain.Selected(0) = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

