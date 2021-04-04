VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSeriesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SeriesUID As String          '����UID�����ؼ��֣�
Public SeriesNo As String           '���кţ�(0020,0011) : Series Number
Public StudyUID As String           '���UID
Public ImageInfos As New Collection 'ͼ����Ϣ�ļ���
Public lngSource As Long            'ͼ����Դ��0-��PACSͼ����������أ�1-ֱ�Ӵ��ļ���2---��ϣ�3-�������ɵ����У�����ʸ��״λ�ؽ���ͼ��ƴ�ӡ�α�����ɵ�ͼ��
Public strHostIP As String          '�洢������IP��������FTP��Ҳ���ڹ���Ŀ¼
Public strFTPUser As String         'FTP �û���
Public strFTPPasw As String         'FTP ����
Public strFTPDir As String          'FTPĿ¼��FTP����Ŀ¼����ʽ��"/FTPĿ¼/"
Public strSaveDir As String         '����ͼ��ľ���·������"��������\���UID\"��϶��ɡ���ȡFTPĿ¼�ͱ��浽���ػ���ʱ����ʹ���������
Public strShareDir As String        '����Ŀ¼��ʽ��Ŀ¼,��������ṹ��""\\"+strHostIP+"\"+ strShareDir+"\"+ strSaveDir" ���Ǳ���ͼ��ľ���·�����ṹ��"\\IP ��ַ\����Ŀ¼��\�ļ�Ŀ¼"
Public strShareDirUser As String    '����Ŀ¼��ʽ���û���
Public strShareDirPasw As String    '����Ŀ¼��ʽ������
Public intSortType As Integer       '��¼��ǰ���е�����ʽ��0--ͼ��ţ�1--��λ����2--��λ����3--�ɼ�ʱ�䣻4--ͼ��ʱ�䣬����ZLShowSeriesInfos��ʹ�á�
Public strModality As String        'Ӱ����� (0008,0060) : Modality
Public lngWinWidth As Long          '��ǰ���еĴ������������е�ʱ�򣬴�"Ӱ��Ԥ�贰����λ"���ж�ȡԤ��Ĵ������ֹ�����F3-F12�Ĵ����󣬼�¼���ú�Ĵ���������ZLShowSeriesInfos��ʹ�á�
Public lngWinLevel As Long          '��ǰ���еĴ�λ���������е�ʱ�򣬴�"Ӱ��Ԥ�贰����λ"���ж�ȡԤ��Ĵ�λ���ֹ�����F3-F12�Ĵ�λ�󣬼�¼���ú�Ĵ�λ������ZLShowSeriesInfos��ʹ�á�
Public intRow As Integer            '��ǰ�������ڵ�����������ZLShowSeriesInfos��ʹ�á�
Public intCol As Integer            '��ǰ�������ڵ�����������ZLShowSeriesInfos��ʹ�á�
Public blnImageSyn As Boolean       '�����ڵ�ͼ���Ƿ�ͬ��������ZLShowSeriesInfos��ʹ�á�
Public StretchToFit As Boolean      '����Ӧ���š�����ZLShowSeriesInfos��ʹ�á�
Public ScrollX As Long              'X�����ƶ�������ZLShowSeriesInfos��ʹ�á�
Public ScrollY As Long              'Y�����ƶ�������ZLShowSeriesInfos��ʹ�á�
Public Zoom As Double               '���ű���������ZLShowSeriesInfos��ʹ�á�
Public FlipState As Integer         '����״̬������ZLShowSeriesInfos��ʹ�á�
Public RotateState As Integer       '��ת״̬������ZLShowSeriesInfos��ʹ�á�
Public FilterLength As Integer      'ƽ��������ZLShowSeriesInfos��ʹ�á�
Public UnsharpEnhancement As Double '��Ե��ǿǿ�ȡ�����ZLShowSeriesInfos��ʹ�á�
Public UnsharpLength As Integer     '��Ե��ǿ���ȡ�����ZLShowSeriesInfos��ʹ�á�
Public MultiFrame As Integer        '�Ƿ��֡ͼ��0-��֡ͼ��1-��֡ͼ��2-��Ƶ¼��3-��Ƶ���ݣ���ʹ�ã���
Public Selected As Boolean          '�Ƿ�ѡ������ͼ������ͬ��������ZLShowSeriesInfos��ʹ�á�
Public strCName As String           'ͼ�������ݿ��е���������ֱ�Ӵ򿪵��ļ����ֶ�Ϊ��
Public strEName As String           'ͼ�������ݿ��е�ƴ������ֱ�Ӵ򿪵��ļ����ֶ�Ϊ��
Public strAge As String             'ͼ�������ݿ��е����䣬ֱ�Ӵ򿪵��ļ����ֶ�Ϊ��
Public strStudyID As String         'ͼ�������ݿ��еļ��ţ�ֱ�Ӵ򿪵��ļ����ֶ�Ϊ��
Public strSex As String             'ͼ�������ݿ��е��Ա�ֱ�Ӵ򿪵��ļ����ֶ�Ϊ��
Public strOrderID As String         'ͼ�������ݿ��е�ҽ��ID��ֱ�Ӵ򿪵��ļ����ֶ�Ϊ��

Private Sub Class_Initialize()
    Set ImageInfos = New Collection
End Sub