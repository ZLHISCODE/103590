VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''
''''���������÷�����
''''�ƽ�
'''''''''''''''''''''''''''''''''''''''''''
Public bModified As Boolean                                ''�Ƿ񱻸�����
Public lngFuncNo As Long                                   ''�������
Public lngMouseKey As Long                                 ''����λ
Public lngShift As Long                                    ''Shift��λ
Public strShowName As String                               ''��������
Public bSelected As Boolean                                ''�Ƿ�ѡ��
Public strProgramName As String                            ''��ť��������
Public ButtomID As Long                                    ''��ťID