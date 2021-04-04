VERSION 5.00
Begin VB.Form frmPatholLoseEnreg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ʧ"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3165
   Icon            =   "frmPatholLoseEnreg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2895
      Begin VB.TextBox txtCount 
         Height          =   300
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "1"
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label labCount 
         Caption         =   "��ʧ������"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPatholLoseEnreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsFind As Boolean
Private mlngMaterialArchivesId As Long
Private mufgParentGrid As ucFlexGrid

Public blnIsOk As Boolean


Public Sub ShowLoseWindow(ufgMaterialGrid As ucFlexGrid, owner As Object)
'��ʾ������ʧ����
    Me.Caption = "������ʧ"
    labCount.Caption = "��ʧ������"
    
    blnIsOk = False
    mblnIsFind = False
    Set mufgParentGrid = ufgMaterialGrid
    
    If Not ufgMaterialGrid.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ������ʧ����Ĳ��ϼ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Me.Show(1, owner)
End Sub


Public Sub ShowFindWindow(ufgMaterialGrid As ucFlexGrid, owner As Object)
'��ʾ�����һش���
    Me.Caption = "�����һ�"
    labCount.Caption = "�һ�������"
    
    blnIsOk = False
    mblnIsFind = True
    Set mufgParentGrid = ufgMaterialGrid
    
    If Not ufgMaterialGrid.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����һش���Ĳ��ϼ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Me.Show(1, owner)
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    blnIsOk = False
    
    Call Me.Hide
err.Clear
End Sub

Private Sub MaterialFind()
'�����һ�
    Dim lngMaterialArchivesId As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnFind As Boolean
    Dim strValue As String
    Dim chkState As CheckState
    
    lngMaterialArchivesId = Val(mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow))
    
    strSql = "select ZL_�������_�����һ�([1],[2],[3]) as ����ֵ  from dual"
                                        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMaterialArchivesId, _
                                                                Val(txtCount.Text), _
                                                                UserInfo.����)
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "MaterialFind", "����ȡ���һش����Ĳ���״̬������ʧ�ܡ�")
        Exit Sub
    End If

'    Call mufgParentGrid.GetFieldDisplayText(gstrPatholCol_���״̬, Nvl(rsData!����ֵ), blnFind, chkState, strValue)
'    Call mufgParentGrid.SetText(mufgParentGrid.SelectionRow, gstrPatholCol_���״̬, strValue, True)
    
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_���״̬, Nvl(rsData!����ֵ), True)
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_�ڵ�����, _
                                Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_�ڵ�����)) + Val(txtCount.Text), True)
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_��ʧ����, _
                                Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_��ʧ����)) - Val(txtCount.Text), True)
End Sub

Private Sub MaterialLose()
'������ʧ
    Dim lngMaterialArchivesId As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strValue As String
    
    lngMaterialArchivesId = Val(mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow))
    
    strSql = "select ZL_�������_������ʧ([1],[2],[3], [4]) as ����ֵ  from dual"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMaterialArchivesId, _
                                                            Val(txtCount.Text), _
                                                            CDate(zlDatabase.Currentdate), _
                                                            UserInfo.����)
                                                            
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "MaterialFind", "����ȡ����ʧ�����Ĳ���״̬������ʧ�ܡ�")
        Exit Sub
    End If

'    Call mufgParentGrid.GetFieldConvertValue(gstrPatholCol_���״̬, Nvl(rsData!����ֵ), blnFind, chkState, strValue)
'    Call mufgParentGrid.SetText(mufgParentGrid.SelectRowIndex, gstrPatholCol_���״̬, strValue, True)
    
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_���״̬, Nvl(rsData!����ֵ), True)
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_�ڵ�����, _
                                Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_�ڵ�����)) - Val(txtCount.Text), True)
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_��ʧ����, _
                                Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_��ʧ����)) + Val(txtCount.Text), True)
    
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    If mblnIsFind Then
'        If mufgParentGrid.Text(mufgParentGrid.SelectRowIndex, gstrPatholCol_���״̬) = "�浵��" Then
'            Call MsgBoxD(Me, "�ò���û����ʧ�����ܽ����һش���", vbOKOnly, Me.Caption)
'            Exit Sub
'        End If

        If Val(txtCount.Text) > Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_��ʧ����)) Then
            Call MsgBoxD(Me, "�����һ��������ܴ�����ʧ������", vbOKOnly, Me.Caption)
            Exit Sub
        End If
        
        
        '�����һش���
        Call MaterialFind
    Else
'        If mufgParentGrid.Text(mufgParentGrid.SelectRowIndex, gstrPatholCol_���״̬) = "�浵��" Then
'            Call MsgBoxD(Me, "�ò����Ѿ���ʧ�����ܽ�����ʧ����", vbOKOnly, Me.Caption)
'            Exit Sub
'        End If
        
        If Val(txtCount.Text) > Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_�ڵ�����)) Then
            Call MsgBoxD(Me, "������ʧ�������ܴ����ڵ�������", vbOKOnly, Me.Caption)
            Exit Sub
        End If
        
        
        '������ʧ����
        Call MaterialLose
    End If
    
    blnIsOk = True
    
    Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call RestoreWinState(Me, App.ProductName)
    
err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
err.Clear
End Sub
