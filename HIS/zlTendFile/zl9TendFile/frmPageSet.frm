VERSION 5.00
Begin VB.Form frmPageSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҳ������"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmPageSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -120
      TabIndex        =   26
      Top             =   4590
      Width           =   6495
   End
   Begin VB.TextBox txtBegin 
      Height          =   300
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   19
      Text            =   "1"
      Top             =   4200
      Width           =   735
   End
   Begin VB.Frame fraHF 
      Caption         =   "ҳü��ҳ��"
      Height          =   2115
      Left            =   180
      TabIndex        =   24
      Top             =   1980
      Width           =   5535
      Begin VB.CommandButton cmdFooter 
         Caption         =   "�Զ���ҳ��(&F)..."
         Height          =   315
         Left            =   3210
         TabIndex        =   15
         Top             =   1020
         Width           =   2145
      End
      Begin VB.CommandButton cmdHeader 
         Caption         =   "�Զ���ҳü(&H)..."
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   1020
         Width           =   2145
      End
      Begin VB.ComboBox cmbFooter 
         Height          =   300
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1680
         Width           =   5205
      End
      Begin VB.ComboBox cmbHeader 
         Height          =   300
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   570
         Width           =   5205
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҳü(&A):"
         Height          =   180
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҳ��(&B):"
         Height          =   180
         Left            =   150
         TabIndex        =   16
         Top             =   1410
         Width           =   720
      End
   End
   Begin VB.Frame frmMargin 
      Caption         =   "ҳ��߾�"
      Height          =   1605
      Left            =   150
      TabIndex        =   23
      Top             =   150
      Width           =   5565
      Begin VB.TextBox txtFooter 
         Height          =   285
         Left            =   4050
         MaxLength       =   3
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtHeader 
         Height          =   285
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   9
         Top             =   870
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Height          =   285
         Left            =   4050
         MaxLength       =   3
         TabIndex        =   7
         Top             =   510
         Width           =   615
      End
      Begin VB.TextBox txtUp 
         Height          =   285
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   5
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Height          =   285
         Left            =   4050
         MaxLength       =   3
         TabIndex        =   3
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   1
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  ��ע�����ȵĵ�λ�Ǻ��ף�"
         Height          =   180
         Left            =   -60
         TabIndex        =   25
         Top             =   1320
         Width           =   2340
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ҳ��λ��(&T):"
         Height          =   180
         Left            =   2910
         TabIndex        =   10
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ҳüλ��(&E):"
         Height          =   180
         Left            =   180
         TabIndex        =   8
         Top             =   930
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ҳ�±߾�(&D):"
         Height          =   180
         Left            =   2910
         TabIndex        =   6
         Top             =   570
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ҳ�ϱ߾�(&U):"
         Height          =   180
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ҳ�ұ߾�(&R):"
         Height          =   180
         Left            =   2910
         TabIndex        =   2
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ҳ��߾�(&L):"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   22
      Top             =   4830
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3390
      TabIndex        =   21
      Top             =   4830
      Width           =   1100
   End
   Begin VB.CommandButton cmdOption 
      Caption         =   "ѡ��(&P)"
      Height          =   350
      Left            =   780
      TabIndex        =   20
      Top             =   4830
      Width           =   1100
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼҳ��(&R):"
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   4230
      Width           =   1080
   End
End
Attribute VB_Name = "frmPageSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��������������ĸ�������
'
'ShowError      ���ݲ�ͬ������ʾ��ͬ������Ϣ,���������Ƶ�ָ��TextBox�ؼ���
'ShowSet        ���ⲿ������ʾ������
'
'
Dim mstrHeader(5) As String      '
Dim mstrFooter(5) As String
Dim mstrHeaderTemp As String
Dim mstrFooterTemp As String
Dim mblnOK As Boolean

Dim sngWidth As Single, sngHeight As Single

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFooter_Click()
    Dim strTemp As String
    If cmbFooter.ListIndex < 6 Then
        strTemp = mstrFooter(cmbFooter.ListIndex)
    Else
        strTemp = mstrFooterTemp
    End If
    Dim frmTemp1 As New frmHF
    If frmTemp1.GetText(strTemp) Then
        mstrFooterTemp = strTemp
        If cmbFooter.ListCount = 7 Then
            If strTemp = ";;" Then
                cmbFooter.RemoveItem 6
                cmbFooter.ListIndex = 0
            Else
                cmbFooter.List(6) = ConvHF(strTemp)
            End If
        Else
            If strTemp = ";;" Then
                cmbFooter.ListIndex = 0
            Else
                cmbFooter.AddItem ConvHF(strTemp)
                cmbFooter.ListIndex = 6
            End If
        End If
    End If
    Set frmTemp1 = Nothing
End Sub

Private Sub cmdHeader_Click()
    Dim strTemp As String
    If cmbHeader.ListIndex < 6 Then
        strTemp = mstrHeader(cmbHeader.ListIndex)
    Else
        strTemp = mstrHeaderTemp
    End If
    Dim frmTemp1 As New frmHF
    If frmTemp1.GetText(strTemp) Then
        mstrHeaderTemp = strTemp
        If cmbHeader.ListCount = 7 Then
            If strTemp = ";;" Then
                cmbHeader.RemoveItem 6
                cmbHeader.ListIndex = 0
            Else
                cmbHeader.List(6) = ConvHF(strTemp)
            End If
        Else
            If strTemp = ";;" Then
                cmbHeader.ListIndex = 0
            Else
                cmbHeader.AddItem ConvHF(strTemp)
                cmbHeader.ListIndex = 6
            End If
        End If
    End If
    Set frmTemp1 = Nothing
End Sub
Private Sub ShowError(ByVal intErr As Integer, txtTemp As TextBox)
    '���ܣ����ݲ�ͬ������ʾ��ͬ������Ϣ,���������Ƶ�ָ��TextBox�ؼ���
    '������intErr       �������
    '      txtTemp      ����ý���Ŀؼ�
    '���أ���
        txtTemp.SelStart = 0
        txtTemp.SelLength = Len(txtTemp.Text)
        Select Case intErr
            Case 1
                MsgBox "���������֡�", vbCritical, gstrSysName
            Case 2
                MsgBox "�߾����ò�������ָ��ֽ�Ŵ�С��", vbCritical, gstrSysName
        End Select
        txtTemp.SetFocus
End Sub

Private Sub cmdOK_Click()
    '���ж�������ǲ�������
    If Not IsNumeric(txtUp.Text) Then
        ShowError 1, txtUp
        Exit Sub
    End If
    If Not IsNumeric(txtDown.Text) Then
        ShowError 1, txtDown
        Exit Sub
    End If
    If Not IsNumeric(txtLeft.Text) Then
        ShowError 1, txtLeft
        Exit Sub
    End If
    If Not IsNumeric(txtRight.Text) Then
        ShowError 1, txtRight
        Exit Sub
    End If
    If Not IsNumeric(txtHeader.Text) Then
        ShowError 1, txtHeader
        Exit Sub
    End If
    If Not IsNumeric(txtFooter.Text) Then
        ShowError 1, txtFooter
        Exit Sub
    End If
    
    '���ж�����ı߾��Ƿ񳬽�
    Dim sngLeft As Single, sngRight As Single, sngUp As Single, sngDown As Single
    Dim sngHeader As Single, sngFooter As Single
    sngUp = Val(txtUp.Text) * conRatemmToTwip
    sngDown = Val(txtDown.Text) * conRatemmToTwip
    sngLeft = Val(txtLeft.Text) * conRatemmToTwip
    sngRight = Val(txtRight.Text) * conRatemmToTwip
    sngFooter = Val(txtFooter.Text) * conRatemmToTwip
    sngHeader = Val(txtHeader.Text) * conRatemmToTwip
    
    If sngUp < 0 Or sngUp > sngHeight - sngDown Then
        ShowError 2, txtUp
        Exit Sub
    End If
    If sngDown < 0 Or sngDown > sngHeight - sngUp Then
        ShowError 2, txtDown
        Exit Sub
    End If
    If sngLeft < 0 Or sngLeft > sngWidth - sngRight Then
        ShowError 2, txtLeft
        Exit Sub
    End If
    If sngRight < 0 Or sngRight > sngWidth - sngLeft Then
        ShowError 2, txtRight
        Exit Sub
    End If
    If sngHeader < 0 Or sngHeader > sngHeight - sngFooter Then
        ShowError 2, txtHeader
        Exit Sub
    End If
    If sngFooter < 0 Or sngFooter > sngHeight - sngHeader Then
        ShowError 2, txtFooter
        Exit Sub
    End If
    '���б���
    gsngUp = Val(txtUp.Text)
    gsngDown = Val(txtDown.Text)
    gsngLeft = Val(txtLeft.Text)
    gsngRight = Val(txtRight.Text)
    gsngFooter = Val(txtFooter.Text)
    gsngHeader = Val(txtHeader.Text)
    
    gobjSend.EmptyDown = gsngDown
    gobjSend.EmptyLeft = gsngLeft
    gobjSend.EmptyRight = gsngRight
    gobjSend.EmptyUp = gsngUp
    
    gsngPageWidth = Printer.Width
    gsngPageHeight = Printer.Height
    gsngPageScaleWidth = Printer.ScaleWidth
    gsngPageScaleHeight = Printer.ScaleHeight
    gintSize = Printer.PaperSize
    gintOri = Printer.Orientation
    gsngScaleWidth = sngWidth
    gsngScaleHeight = sngHeight
    
    If cmbHeader.ListIndex < 6 Then
        gstrHeader = mstrHeader(cmbHeader.ListIndex)
    Else
        gstrHeader = mstrHeaderTemp
    End If
    If cmbFooter.ListIndex < 6 Then
        gstrFooter = mstrFooter(cmbFooter.ListIndex)
    Else
        gstrFooter = mstrFooterTemp
    End If
    gintBegin = Val(txtBegin.Text)
    If gintBegin < 1 Then gintBegin = 1
    zlPutPrinterSet
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdOption_Click()
    Dim frmPrintTemp As New frmPrintSet
    frmPrintTemp.Show 1
'    If Printer.Orientation = 1 Then '�����ӡ
'        sngWidth = IIf(Printer.ScaleWidth < Printer.ScaleHeight, Printer.ScaleWidth, Printer.ScaleHeight) '�ĵ���ӡ��ֽ��խ����������
'        sngHeight = IIf(Printer.ScaleWidth > Printer.ScaleHeight, Printer.ScaleWidth, Printer.ScaleHeight)
'    Else
'        sngWidth = IIf(Printer.ScaleWidth > Printer.ScaleHeight, Printer.ScaleWidth, Printer.ScaleHeight)   '�ĵ���ӡ��ֽ�Ŀ��������
'        sngHeight = IIf(Printer.ScaleWidth < Printer.ScaleHeight, Printer.ScaleWidth, Printer.ScaleHeight)
'    End If
    sngWidth = Printer.ScaleWidth
    sngHeight = Printer.ScaleHeight
End Sub

Private Sub Form_Load()
    txtFooter.Text = gsngFooter
    txtHeader.Text = gsngHeader
    txtUp.Text = gsngUp
    txtDown.Text = gsngDown
    txtLeft.Text = gsngLeft
    txtRight.Text = gsngRight

    txtBegin.Text = CStr(gintBegin)
    
    mstrFooterTemp = gstrFooter
    mstrHeaderTemp = gstrHeader
    
    sngWidth = gsngScaleWidth
    sngHeight = gsngScaleHeight
    
    mstrFooter(0) = ";;"
    mstrFooter(1) = "[��λ��]" & ";" & "" & ";" & "[�û���]"
    mstrFooter(2) = "" & ";" & "" & ";" & "[����][ʱ��]"
    mstrFooter(3) = "��ӡ�ˣ�[�û���]" & ";" & "" & ";" & "��ӡʱ�䣺[����]"
    mstrFooter(4) = "��[ҳ��]ҳ" & ";" & "��[ҳ��]ҳ" & ";" & "[����]"
    mstrFooter(5) = "[��λ��]" & ";" & "" & ";" & "��[ҳ��]ҳ"
    
    mstrHeader(0) = "" & ";" & "" & ";" & ""
    mstrHeader(1) = "[��λ��]" & ";" & "" & ";" & "[�û���]"
    mstrHeader(2) = "" & ";" & "" & ";" & "[����][ʱ��]"
    mstrHeader(3) = "��ӡ�ˣ�[�û���]" & ";" & "" & ";" & "��ӡʱ�䣺[����]"
    mstrHeader(4) = "��[ҳ��]ҳ" & ";" & "��[ҳ��]ҳ" & ";" & "[����]"
    mstrHeader(5) = "[��λ��]" & ";" & "" & ";" & "��[ҳ��]ҳ"
    
    Dim i As Integer
    For i = 0 To 5
        cmbFooter.AddItem ConvHF(mstrFooter(i))
    Next
    cmbFooter.List(0) = "��"
    For i = 0 To 5
        cmbHeader.AddItem ConvHF(mstrHeader(i))
    Next
    cmbHeader.List(0) = "��"
    
    If mstrFooterTemp = "" Or mstrFooterTemp = ";;" Then
        mstrFooterTemp = ";;"
        cmbFooter.ListIndex = 0
    Else
        cmbFooter.AddItem ConvHF(mstrFooterTemp)
        cmbFooter.ListIndex = cmbFooter.NewIndex
    End If
    If mstrHeaderTemp = "" Or mstrHeaderTemp = ";;" Then
        mstrHeaderTemp = ";;"
        cmbHeader.ListIndex = 0
    Else
        cmbHeader.AddItem ConvHF(mstrHeaderTemp)
        cmbHeader.ListIndex = cmbHeader.NewIndex
    End If
End Sub
Public Function ShowSet() As Boolean
    '���ܣ����ⲿ������ʾ������
    '���أ�����ǰ���ȡ�����򷵻�Fasle ������ȷ�����򷵻�True
    mblnOK = False
    Me.Show 1
    ShowSet = mblnOK
End Function

Private Sub txtBegin_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub
