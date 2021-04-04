VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��Ӱ����"
   ClientHeight    =   1275
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   6750
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   1275
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtEnd 
      Height          =   300
      Left            =   5370
      TabIndex        =   12
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txtStart 
      Height          =   300
      Left            =   5370
      TabIndex        =   10
      Top             =   390
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Height          =   420
      Left            =   1170
      Picture         =   "frmCine.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   708
      Width           =   900
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   1860
      Top             =   84
   End
   Begin VB.CommandButton cmdAlong 
      Height          =   420
      Left            =   2070
      Picture         =   "frmCine.frx":041E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   708
      Width           =   900
   End
   Begin VB.CommandButton cmdBackwards 
      Height          =   420
      Left            =   270
      Picture         =   "frmCine.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   708
      Width           =   900
   End
   Begin MSComctlLib.Slider sldCine 
      Height          =   324
      Left            =   252
      TabIndex        =   1
      Top             =   360
      Width           =   2736
      _ExtentX        =   4815
      _ExtentY        =   582
      _Version        =   393216
      Min             =   5
      Max             =   104
      SelStart        =   95
      TickStyle       =   3
      Value           =   95
   End
   Begin VB.Frame famPlayMode 
      Caption         =   "����ģʽ"
      Height          =   1050
      Left            =   3252
      TabIndex        =   0
      Top             =   120
      Width           =   1344
      Begin VB.OptionButton OptShuffle 
         Caption         =   "�Ӱ�"
         Height          =   216
         Left            =   204
         TabIndex        =   3
         Top             =   672
         Width           =   696
      End
      Begin VB.OptionButton OptLoop 
         Caption         =   "ѭ��"
         Height          =   180
         Left            =   204
         TabIndex        =   2
         Top             =   324
         Value           =   -1  'True
         Width           =   708
      End
   End
   Begin VB.Label lblCurrentNo 
      AutoSize        =   -1  'True
      Caption         =   "��ǰͼ��ţ�"
      Height          =   180
      Left            =   4770
      TabIndex        =   13
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "�յ㣺"
      Height          =   180
      Left            =   4770
      TabIndex        =   11
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��㣺"
      Height          =   180
      Left            =   4770
      TabIndex        =   9
      Top             =   450
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�ӿ�"
      Height          =   180
      Left            =   2568
      TabIndex        =   7
      Top             =   96
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   252
      TabIndex        =   6
      Top             =   96
      Width           =   360
   End
End
Attribute VB_Name = "frmCine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ʹ�ø���괩����ͬ�ķ��������ŵ�ӰͼƬ
Public f As frmViewer
Dim blnisBack As Boolean            'ͼ�񲥷ŵķ���True--��󲥷ţ�False--��ǰ���š�

Private Sub cmdAlong_Click()
    blnisBack = False       '��ǰ����ͼ��
    Timer.Enabled = True
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
End Sub

Private Sub cmdBackwards_Click()
    blnisBack = True        '��󲥷�ͼ��
    Timer.Enabled = True
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
End Sub

Private Sub cmdStop_Click()
    Timer.Enabled = False
    Me.txtStart.Enabled = True
    Me.txtEnd.Enabled = True
End Sub

Private Sub Form_Load()
    
    Dim thisViewer As DicomViewer
    If f.intSelectedSerial = 0 Then Exit Sub
    
    Set thisViewer = f.Viewer(f.intSelectedSerial)
    '���㲥��ͼ����Viewer�е�ƫ����
    f.intStackOffset = f.SelectedImageIndex - thisViewer.CurrentIndex    '��¼��ǰͼ���CurrentIndex֮��ľ���
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If f.SelectedImage.FrameCount > 1 Then   ''''��֡ͼ����
        f.intStackCurrentlyImage = f.SelectedImage.Frame    '��¼��ǰ������ǰͼ���
        f.blnStackisFrame = True            '��¼���ö�֡���Ż��ǵ���ѭ������
        Me.txtEnd.Text = thisViewer.Images(f.SelectedImageIndex).FrameCount
    Else                                    '��֡ͼ����
        f.blnStackisFrame = False           '��¼���ö�֡���Ż��ǵ���ѭ������
        '��¼����ǰViewer��CurrentIndex�͵�ǰͼ��
        Set f.SelectedLabel = Nothing
        f.intStackCurrentlyImage = thisViewer.CurrentIndex   '��¼��ǰ������ǰͼ���
        Set f.objStackOldImage = thisViewer.Images(f.SelectedImageIndex)
        f.intStackIndex = thisViewer.Images(f.SelectedImageIndex).Tag
        Me.txtEnd.Text = ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count
    End If
        
    Me.txtStart.Text = 1
    Timer.Interval = (105 - sldCine) * 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�ָ���Ӱ����ǰ�Ĵ���״̬
    Timer.Enabled = False
    Dim j As Integer
    If f.intSelectedSerial = 0 Then Exit Sub
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With f.Viewer(f.intSelectedSerial)
    
        If f.blnStackisFrame Then    ''''��֡ͼ����
            j = f.SelectedImage.Frame - f.intStackOffset
            f.SelectedImage.Frame = f.intStackCurrentlyImage
        Else
            '���ú�����������
            subStackEnd f.Viewer(f.intSelectedSerial), f
            j = f.intStackIndex - f.intStackOffset
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If j > ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count - .MultiColumns * .MultiRows + 1 Then
            j = ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count - .MultiColumns * .MultiRows + 1
        End If
        If j < 1 Then j = 1
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If f.VScro(f.intSelectedSerial).Visible Then f.VScro(f.intSelectedSerial).Value = j
    End With
End Sub

Private Sub sldCine_Click()
    Timer.Interval = (105 - sldCine) * 5
End Sub

Private Sub sldCine_GotFocus()
    Me.cmdStop.SetFocus
End Sub

Private Sub Timer_Timer()
    '��ʱ��������ͼ����Զ�����
    Dim objTempImage As DicomImage
    Dim thisViewer As DicomViewer
    
    Set thisViewer = f.Viewer(f.intSelectedSerial)
    If f.blnStackisFrame Then    ''''��֡ͼ����
        If Not blnisBack Then   '��ǰ����ͼ��
            If thisViewer.Images(f.SelectedImageIndex).Frame >= Val(Me.txtEnd.Text) Then
                If OptShuffle.Value Then
                   thisViewer.Images(f.SelectedImageIndex).Frame = Val(Me.txtEnd.Text) - 1
                   blnisBack = True
                Else
                    thisViewer.Images(f.SelectedImageIndex).Frame = Val(Me.txtStart.Text)
                End If
            Else
                thisViewer.Images(f.SelectedImageIndex).Frame = thisViewer.Images(f.SelectedImageIndex).Frame + 1
            End If
        Else                    '��󲥷�ͼ��
            If thisViewer.Images(f.SelectedImageIndex).Frame <= Val(Me.txtStart.Text) Then
                If OptShuffle.Value Then
                   thisViewer.Images(f.SelectedImageIndex).Frame = Val(Me.txtStart.Text) + 1
                   blnisBack = False
                Else
                    thisViewer.Images(f.SelectedImageIndex).Frame = Val(Me.txtEnd.Text)
                End If
            Else
                thisViewer.Images(f.SelectedImageIndex).Frame = thisViewer.Images(f.SelectedImageIndex).Frame - 1
            End If
        End If
        Me.lblCurrentNo.Caption = "��ǰͼ��ţ�" & thisViewer.Images(f.SelectedImageIndex).Frame
    Else        '��֡ͼ����
        '�����µ�λ��
        If Not blnisBack Then       '��ǰ����ͼ��
            If f.intStackIndex >= Val(Me.txtEnd.Text) Then
                If OptShuffle.Value Then
                   f.intStackIndex = f.intStackIndex - 1
                   blnisBack = True
                Else
                    f.intStackIndex = Val(Me.txtStart.Text)
                End If
            Else
                f.intStackIndex = f.intStackIndex + 1
            End If
        Else                        '��󲥷�ͼ��
            If f.intStackIndex <= Val(Me.txtStart.Text) Then
                If OptShuffle.Value Then
                   f.intStackIndex = Val(Me.txtStart.Text) + 1
                   blnisBack = False
                Else
                    f.intStackIndex = Val(Me.txtEnd.Text)
                End If
            Else
                 f.intStackIndex = f.intStackIndex - 1
            End If
        End If
        
        '��ָ��λ�õ�ͼ����ӵ�Viewer��
        Set objTempImage = funLoadAImage(f.intSelectedSerial, f.intStackIndex, 1)
        If Not objTempImage Is Nothing Then
            Call subInitAImage(objTempImage, f.intSelectedSerial, thisViewer)
            
            thisViewer.Images.Add objTempImage
            thisViewer.Images.Move thisViewer.Images.Count, f.SelectedImageIndex
            thisViewer.Images.Remove f.SelectedImageIndex + 1
            thisViewer.CurrentIndex = f.intStackCurrentlyImage
            
            Me.lblCurrentNo.Caption = "��ǰͼ��ţ�" & f.intStackIndex
        End If
    End If
End Sub

Private Sub txtEnd_LostFocus()
    Dim iImage As Integer
    Dim bError As Boolean
    
    iImage = Val(f.MSFViewer.TextMatrix(f.intSelectedSerial, 3))
    If Val(Me.txtEnd.Text) < Val(Me.txtStart.Text) Then bError = True
    If f.blnStackisFrame Then        '�����֡ͼ��
        If Val(Me.txtEnd.Text) > f.Viewer(f.intSelectedSerial).Images(iImage).FrameCount Then bError = True
    Else
        If Val(Me.txtEnd.Text) > ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count Then bError = True
    End If
    If bError Then
        MsgBox "��ֵֹҪ���ڿ�ʼֵ����С��ͼ��������", vbExclamation, gstrSysName
        If f.blnStackisFrame Then       '�����֡ͼ��
            Me.txtEnd.Text = f.Viewer(f.intSelectedSerial).Images(iImage).FrameCount
        Else
            Me.txtEnd.Text = ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count
        End If
    End If
End Sub

Private Sub txtStart_LostFocus()
    If Val(Me.txtStart.Text) > Val(Me.txtEnd.Text) Or Val(Me.txtStart.Text) < 1 Then
        MsgBox "��ʼֵҪС�����ֵ���Ҵ���1��", vbInformation, gstrSysName
        Me.txtStart.Text = 1
    End If
End Sub
