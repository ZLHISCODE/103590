VERSION 5.00
Begin VB.Form frmTechnicStudy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����Ŀ����"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5670
   Icon            =   "frmTechnicStudy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5670
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   375
      Left            =   3390
      Picture         =   "frmTechnicStudy.frx":000C
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5835
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   375
      Left            =   4485
      Picture         =   "frmTechnicStudy.frx":0156
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5835
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   5670
      TabIndex        =   1
      Top             =   0
      Width           =   5670
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTechnicStudy.frx":02A0
         Height          =   660
         Left            =   225
         TabIndex        =   2
         Top             =   165
         Width           =   5265
      End
   End
   Begin zl9PACSWork.ucFlexGrid ufgStudy 
      Height          =   4755
      Left            =   90
      TabIndex        =   0
      Top             =   975
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   8387
      DefaultCols     =   ""
      ColNames        =   "|ID,hide,key|��Ŀ����>����,w3000,rowcheck|��Ŀ����>����,read,w1200|Ӱ�����,w1000|����ID,hide|"
      KeyName         =   "ID"
      DisCellColor    =   16777215
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      HeadColor       =   0
      DataFontCharset =   134
      DataFontWeight  =   400
      DataColor       =   -2147483640
      RowHeightMin    =   260
      ExtendLastCol   =   -1  'True
   End
End
Attribute VB_Name = "frmTechnicStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptId As Long
Private mlngGroupId As Long
Private mblnOK As Boolean

Public Function ShowStudyAssociation(ByVal lngDeptID As Long, ByVal lngGroupId As Long, objOwner As Object) As Boolean
'��ʾ�����Ŀ����
    mlngDeptId = lngDeptID
    mlngGroupId = lngGroupId
    mblnOK = False
    
    ShowStudyAssociation = False
    
    Call LoadStudyPro
    
    Me.Show 1, objOwner
    
    ShowStudyAssociation = mblnOK
    
End Function


Private Sub LoadStudyPro()
'��������Ŀ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    '��ѯ�÷��������������ļ����Ŀ
        
    
    strSQL = " select ID,����,����,Ӱ�����,����id from (with g as (select Id,����ID,����ID,������ĿID from Ӱ�������� where ����ID=[1] ) " & vbCrLf & _
            " select distinct a.ID,a.����,a.����,b.Ӱ�����,f.����id " & _
            " from ������ĿĿ¼ a, Ӱ������Ŀ b, Ӱ���豸Ŀ¼ c, ҽ��ִ�з��� d, Ӱ��ִ�з��� e, (select ����ID,������ĿID from g where g.����ID=[2]) f " & _
            " Where a.ID = b.������Ŀid And b.Ӱ����� = c.Ӱ����� And c.�豸�� = d.����豸 And d.����id = e.ID " & _
            " and a.Id=f.������Ŀid(+) " & _
            " and not exists (select 1 from g where g.����Id<>[2] and g.������ĿID=a.Id) " & _
            " and e.id=[2])"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯӰ������Ŀ", mlngDeptId, mlngGroupId)
    
    Call ufgStudy.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "����ID Desc,���� Asc"
    Set ufgStudy.AdoData = rsData
    
    ufgStudy.GridRows = ufgStudy.AdoData.RecordCount + 1
    Call ufgStudy.RefreshData
End Sub


Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdSure_Click()
On Error GoTo ErrHandle
    Dim strSQL As String
    Dim strIds As String
    Dim i As Long
    
    strIds = ""
    
    For i = 1 To ufgStudy.GridRows - 1
        If ufgStudy.GetRowCheck(i) Then
            If strIds <> "" Then strIds = strIds & ","
            strIds = strIds & ufgStudy.KeyValue(i)
        End If
    Next i
    
    strSQL = "zl_Ӱ��������_Association(" & mlngDeptId & "," & mlngGroupId & ",'" & strIds & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "Ӱ��������")
    
    mblnOK = True
    
    Unload Me
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
'    'Debug Code
'    InitDebugObject 1290, Me, "zlhis", "HIS"
'    mlngGroupId = 29
'
'    LoadStudyPro
'    'Debug End
End Sub

Private Sub ufgStudy_OnNewRow(ByVal Row As Long)
    If Val(ufgStudy.Text(Row, "����ID")) = mlngGroupId Then Call ufgStudy.SetRowCheck(Row, True)
End Sub

