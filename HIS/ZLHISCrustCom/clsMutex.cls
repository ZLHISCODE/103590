VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2019/1/16
'ģ��           clsMutex
'˵��           ������
'==================================================================================================
Private Declare Function CreateMutexA Lib "kernel32.dll" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
'@ԭ��
'    HANDLE WINAPI CreateMutex(
'      _In_opt_ LPSECURITY_ATTRIBUTES lpMutexAttributes,
'      _In_     BOOL                  bInitialOwner,
'      _In_opt_ LPCTSTR               lpName
'    );
'@����
'    ���������������δ�����������ҪΪ����ָ���������룬��ʹ��CreateMutexEx������
'@����
'lpMutexAttributes _In_opt_
'   ָ��SECURITY_ATTRIBUTES�ṹ��ָ�롣����ò���Ϊ�գ����ӽ��̲��ܼ̳о����
'   �ṹ��lpSecurityDescriptor��ԱΪ�µĻ������ָ��һ����ȫ�����������lpMutexAttributesΪ�գ�������󽫻��Ĭ�ϵİ�ȫ����������������Ĭ�ϰ�ȫ�������е�acl���Դ������������ƻ�ģ�����ơ��йظ�����Ϣ����μ�ͬ������ȫ�Ժͷ���Ȩ�ޡ�
'bInitialOwner  _In_
'   ������ֵΪ�棬���ҵ����ߴ����˻��������ô�����߳̽���û������ĳ�ʼ����Ȩ�����򣬵����̲߳����û�����������Ȩ��Ҫȷ���������Ƿ񴴽��˻�����������ķ���ֵ���֡�
'lpName _In_opt_
'   �����������ơ����Ʊ�����ΪMAX_PATH���ȡ����ƱȽ����ִ�Сд��
'   ���lpNameƥ�����������������������ƣ���˺�������MUTEX_ALL_ACCESS����Ȩ�ޡ��ڱ����У�bInitialOwner���������ԣ���Ϊ���Ѿ��ɴ����������á����lpMutexAttributes��������NULL������ȷ���Ƿ���Լ̳о�������Ǻ����䰲ȫ��������Ա��
'   ���lpNameΪ�գ��򴴽��������ʱû�����ơ�
'   ���lpNameƥ�������¼����ź������ɵȴ���ʱ������ҵ���ļ�ӳ���������ƣ���ú�����ʧ�ܣ�GetLastError����������ERROR_INVALID_HANDLE��������Ϊ��Щ��������ͬ�����ƿռ䡣
'   ���ƿ��Ծ��� "Global\"��  "Local\"ǰ׺������ʽ����ȫ�ֻ�Ự���ƿռ��д����������Ƶ����ಿ�ֿ��԰�������б���ַ�(\)֮����κ��ַ����йظ�����Ϣ����μ��ں˶������ƿռ䡣ʹ���ն˷���Ựʵ�ֿ����û��л����ں˶������Ʊ�����ѭ�ն˷����ָ��ԭ���Ա�Ӧ�ó����ܹ�֧�ֶ���û���
'   ���������˽�����ƿռ��д������йظ�����Ϣ����μ��������ƿռ䡣
'@����ֵ
'    ��������ɹ�������ֵ���´����Ļ������ľ����
'    �������ʧ�ܣ�����ֵΪNULL��Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'    ������������һ���������Ļ�����󣬲��Ҷ����������������֮ǰ�Ѿ����ڣ���ô����ֵ�������ж���ľ����GetLastError����ERROR_ALREADY_EXISTS, bInitialOwner�����ԣ������߳�û�б���������Ȩ�����ǣ���������߾������޵ķ���Ȩ�ޣ�������ʧ�ܣ�ERROR_ACCESS_DENIED��������Ӧ��ʹ��OpenMutex������
'@��ע
'    CreateMutex���صľ������MUTEX_ALL_ACCESS����Ȩ��;�����������κ���Ҫ����������ĺ�����ǰ���ǵ������ѱ��������Ȩ����������������һ�������һ��ģ�ⲻͬ�û����̴߳����ģ��������ڴ����������ʱ����Ӧ��һ����ȫ������������ͨ��������Ĭ��DACL�����Ĵ������̵�Ĭ�ϰ�ȫ���������йظ�����Ϣ����μ�ͬ������ȫ�Ժͷ���Ȩ�ޡ�
'    ���ʹ��ָ���Ļ�������Ӧ�ó�������Ϊ����ʵ���������û���������֮ǰ�����˻���������ֹӦ�ó���������Ϊ�˷�ֹ�������������һ����������Ļ����壬���洢������ƣ��Ա�ֻ����Ȩ�û����ܻ���������ߣ�������Ϊ��Ŀ��ʹ���ļ���Ҫ��Ӧ�ó�������Ϊÿ���û�һ��ʵ���������û��������ļ�Ŀ¼�д���һ���������ļ���
'    ���ý��̵��κ��̶߳������ڵ���ĳ���ȴ�����ʱָ���������������ָ�������״̬�������ź�ʱ��������ȴ��������ء�����ָʾ�����ȴ��������κ�һ��������ָ�����󶼷����ź�ʱ���ء����ȴ���������ʱ���ͷŵȴ��̼߳���ִ�С�
'    ��������״̬�����������κ��߳�ʱ�����źš������߳̿���ʹ��bInitialOwner��־���󻥳�������������Ȩ�������̱߳���ʹ��һ���ȴ���������������Ȩ������������״̬���ź�ʱ��һ���ȴ��̱߳���������Ȩ����������״̬��Ϊ���ź�״̬���ȴ��������ء��κ�ʱ��ֻ��һ���߳̿���ӵ�л�������ӵ���߳�ʹ��ReleaseMutex�������ͷ�������Ȩ��
'    ӵ�л��������߳̿������ظ��ĵȴ�����������ָ����ͬ�Ļ���������������������ִ�С�ͨ�����������ظ��ȴ�ͬһ�����������������ֻ��ƿ��Է�ֹ�߳��ڵȴ����Ѿ�ӵ�еĻ�����ʱ�����Լ������ǣ�Ϊ���ͷ���������Ȩ���̱߳���ÿ�ε���ReleaseMutexһ�Σ�ÿ�λ���������һ���ȴ���
'    �����������̿��Ե���CreateMutex��������ͬ����������������һ������ʵ�ʴ����������������㹻����Ȩ�޵ĺ�������ֻ������л������ľ������ʹ�ö�������ܹ������ͬ�������ľ����ͬʱ�������û�ȷ�����������������̵����Ρ�ʹ�ô˼���ʱ��Ӧ�ý�bInitialOwner��־����ΪFALSE;���򣬺���ȷ���ĸ����̾��г�ʼ����Ȩ��
'    ������̿��Ծ�����ͬ�������ľ�����Ӷ�֧��ʹ�øö�����н��̼�ͬ�������¶��������ƿ���:
'       ���CreateMutex��lpMutexAttributes���������˼̳У���ô��CreateProcess�����������ӽ��̿��Լ̳л������ľ�������ֻ��Ƽ��������������������Ҳ������δ�����������
'       ���̿����ڵ���DuplicateHandle����ʱָ���������ľ�����Դ���������һ������ʹ�õ��ظ���������ֻ��Ƽ��������������������Ҳ������δ�����������
'       ���̿����ڵ���OpenMutex��CreateMutex����ʱָ��һ���������Ļ�������Լ����������ľ����
'    ʹ��closeHandle�����رվ������������ֹʱ��ϵͳ�Զ��رվ�����������������һ��������ر�ʱ�����������١�
'@Ҫ��
'    Minimum supported client       Windows XP [desktop apps | UWP apps]
'    Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'    Header                         WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Synchapi.h on Windows 8 and Windows Server 2012
'    Library                        kernel32.lib
'    dll                            kernel32.dll
'    Unicode and ANSI names         CreateMutexW (Unicode) And CreateMutexA(ANSI)
Private Declare Function CreateMutexExA Lib "kernel32.dll" (ByVal lpMutexAttributes As Long, ByVal lpName As String, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long) As Long
'@ԭ��
'    HANDLE WINAPI CreateMutexEx(
'      _In_opt_ LPSECURITY_ATTRIBUTES lpMutexAttributes,
'      _In_opt_ LPCTSTR               lpName,
'      _In_     DWORD                 dwFlags,
'      _In_     DWORD                 dwDesiredAccess
'    );
'@����
'    ���������������δ����������󣬲����ظö���ľ����
'@����
'lpMutexAttributes _In_opt_
'    ָ��SECURITY_ATTRIBUTES�ṹ��ָ�롣����ò���ΪNULL�����ӽ��̲��ܼ̳л����������
'    �ṹ��lpSecurityDescriptor��ԱΪ�µĻ������ָ��һ����ȫ�����������lpMutexAttributesΪ�գ�������󽫻��Ĭ�ϵİ�ȫ����������������Ĭ�ϰ�ȫ�������е�acl���Դ������������ƻ�ģ�����ơ��йظ�����Ϣ����μ�ͬ������ȫ�Ժͷ���Ȩ�ޡ�
'lpName _In_opt_
'   �����������ơ����Ʊ�����ΪMAX_PATH���ȡ����ƱȽ����ִ�Сд��
'   ���lpNameƥ�����������������������ƣ���˺�������MUTEX_ALL_ACCESS����Ȩ�ޡ��ڱ����У�bInitialOwner���������ԣ���Ϊ���Ѿ��ɴ����������á����lpMutexAttributes��������NULL������ȷ���Ƿ���Լ̳о�������Ǻ����䰲ȫ��������Ա��
'   ���lpNameΪ�գ��򴴽��������ʱû�����ơ�
'   ���lpNameƥ�������¼����ź������ɵȴ���ʱ������ҵ���ļ�ӳ���������ƣ���ú�����ʧ�ܣ�GetLastError����������ERROR_INVALID_HANDLE��������Ϊ��Щ��������ͬ�����ƿռ䡣
'   ���ƿ��Ծ��� "Global\"��  "Local\"ǰ׺������ʽ����ȫ�ֻ�Ự���ƿռ��д����������Ƶ����ಿ�ֿ��԰�������б���ַ�(\)֮����κ��ַ����йظ�����Ϣ����μ��ں˶������ƿռ䡣ʹ���ն˷���Ựʵ�ֿ����û��л����ں˶������Ʊ�����ѭ�ն˷����ָ��ԭ���Ա�Ӧ�ó����ܹ�֧�ֶ���û���
'   ���������˽�����ƿռ��д������йظ�����Ϣ����μ��������ƿռ䡣dwFlags [��]
'dwFlags _In_
'    �������������0��Ҳ����������ֵ��
Private Const CREATE_MUTEX_INITIAL_OWNER        As Long = &H1
'    ���󴴽����ǻ������ĳ�ʼ������?
'dwDesiredAccess _In_
'    �������ķ������롣�йط���Ȩ���б�����μ�ͬ������ȫ�Ժͷ���Ȩ�ޡ�
'@����ֵ
'    ��������ɹ�������ֵ���´����Ļ������ľ����
'    �������ʧ�ܣ�����ֵΪNULL��Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'    ������������һ���������Ļ�����󣬲��Ҷ����������������֮ǰ�Ѿ����ڣ���ô����ֵ�������ж���ľ����GetLastError����ERROR_ALREADY_EXISTS, bInitialOwner�����ԣ������߳�û�б���������Ȩ�����ǣ���������߾������޵ķ���Ȩ�ޣ�������ʧ�ܣ�ERROR_ACCESS_DENIED��������Ӧ��ʹ��OpenMutex������
'@��ע
'    CreateMutex���صľ������MUTEX_ALL_ACCESS����Ȩ��;�����������κ���Ҫ����������ĺ�����ǰ���ǵ������ѱ��������Ȩ����������������һ�������һ��ģ�ⲻͬ�û����̴߳����ģ��������ڴ����������ʱ����Ӧ��һ����ȫ������������ͨ��������Ĭ��DACL�����Ĵ������̵�Ĭ�ϰ�ȫ���������йظ�����Ϣ����μ�ͬ������ȫ�Ժͷ���Ȩ�ޡ�
'    ���ʹ��ָ���Ļ�������Ӧ�ó�������Ϊ����ʵ���������û���������֮ǰ�����˻���������ֹӦ�ó���������Ϊ�˷�ֹ�������������һ����������Ļ����壬���洢������ƣ��Ա�ֻ����Ȩ�û����ܻ���������ߣ�������Ϊ��Ŀ��ʹ���ļ���Ҫ��Ӧ�ó�������Ϊÿ���û�һ��ʵ���������û��������ļ�Ŀ¼�д���һ���������ļ���
'    ���ý��̵��κ��̶߳������ڵ���ĳ���ȴ�����ʱָ���������������ָ�������״̬�������ź�ʱ��������ȴ��������ء�����ָʾ�����ȴ��������κ�һ��������ָ�����󶼷����ź�ʱ���ء����ȴ���������ʱ���ͷŵȴ��̼߳���ִ�С�
'    ��������״̬�����������κ��߳�ʱ�����źš������߳̿���ʹ��bInitialOwner��־���󻥳�������������Ȩ�������̱߳���ʹ��һ���ȴ���������������Ȩ������������״̬���ź�ʱ��һ���ȴ��̱߳���������Ȩ����������״̬��Ϊ���ź�״̬���ȴ��������ء��κ�ʱ��ֻ��һ���߳̿���ӵ�л�������ӵ���߳�ʹ��ReleaseMutex�������ͷ�������Ȩ��
'    ӵ�л��������߳̿������ظ��ĵȴ�����������ָ����ͬ�Ļ���������������������ִ�С�ͨ�����������ظ��ȴ�ͬһ�����������������ֻ��ƿ��Է�ֹ�߳��ڵȴ����Ѿ�ӵ�еĻ�����ʱ�����Լ������ǣ�Ϊ���ͷ���������Ȩ���̱߳���ÿ�ε���ReleaseMutexһ�Σ�ÿ�λ���������һ���ȴ���
'    �����������̿��Ե���CreateMutex��������ͬ����������������һ������ʵ�ʴ����������������㹻����Ȩ�޵ĺ�������ֻ������л������ľ������ʹ�ö�������ܹ������ͬ�������ľ����ͬʱ�������û�ȷ�����������������̵����Ρ�ʹ�ô˼���ʱ����Ӧ��ʹ��CREATE_MUTEX_INITIAL_OWNER��־;���򣬺���ȷ���ĸ����̾��г�ʼ����Ȩ��
'    ������̿��Ծ�����ͬ�������ľ�����Ӷ�֧��ʹ�øö�����н��̼�ͬ�������¶��������ƿ���:
'       ���CreateMutex��lpMutexAttributes���������˼̳У���ô��CreateProcess�����������ӽ��̿��Լ̳л������ľ�������ֻ��Ƽ��������������������Ҳ������δ�����������
'       ���̿����ڵ���DuplicateHandle����ʱָ���������ľ�����Դ���������һ������ʹ�õ��ظ���������ֻ��Ƽ��������������������Ҳ������δ�����������
'       ���̿����ڵ���OpenMutex��CreateMutex����ʱָ��һ���������Ļ�������Լ����������ľ����
'    ʹ��closeHandle�����رվ������������ֹʱ��ϵͳ�Զ��رվ�����������������һ��������ر�ʱ�����������١�
'@Ҫ��
'    Minimum supported client       Windows XP [desktop apps | UWP apps]
'    Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'    Header                         WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Synchapi.h on Windows 8 and Windows Server 2012
'    Library                        kernel32.lib
'    dll                            kernel32.dll
'    Unicode and ANSI names         CreateMutexExW (Unicode) And CreateMutexExA(ANSI)
Private Declare Function OpenMutexA Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
'@ԭ��
'    HANDLE WINAPI OpenMutex(
'      _In_ DWORD   dwDesiredAccess,
'      _In_ BOOL    bInheritHandle,
'      _In_ LPCTSTR lpName
'    );
'@����
'    �����е������������
'@����
'dwDesiredAccess _In_
'    �Ի������ķ��ʡ�ʹ�û�����ֻ��Ҫͬ������Ȩ;Ҫ���Ļ������İ�ȫ�ԣ���ָ��MUTEX_ALL_ACCESS�����ָ������İ�ȫ���������������ý��̵�������ʣ���ú�����ʧ�ܡ��йط���Ȩ���б�����μ�ͬ������ȫ�Ժͷ���Ȩ�ޡ�
'bInheritHandle _In_
'    �����ֵΪ�棬���ɸý��̴����Ľ��̽��̳иþ�������򣬽��̲���̳д˾����
'lpName _In_opt_
'    Ҫ�򿪵Ļ�����������?���ƱȽ����ִ�Сд?
'    �˺������Դ�˽�����ƿռ��еĶ����йظ�����Ϣ����μ��������ƿռ䡣
'    �ն˷���:���ƿ��Ծ��� "Global\"��  "Local\"ǰ׺������ʽ����ȫ�ֻ�Ự���ƿռ��д����������Ƶ����ಿ�ֿ��԰�������б���ַ�(\)֮����κ��ַ����йظ�����Ϣ����μ��ں˶������ƿռ䡣
'    ע�⣬�����û��л���ʹ���ն˷���Ựʵ�ֵġ���һ����¼���û�ʹ�ûỰ0����һ����¼���û�ʹ�ûỰ1���������ơ�ʹ���ն˷���Ựʵ�ֿ����û��л����ں˶������Ʊ�����ѭ�ն˷����ָ��ԭ���Ա�Ӧ�ó����ܹ�֧�ֶ���û���
'    ���������˽�����ƿռ��д������йظ�����Ϣ����μ��������ƿռ䡣dwFlags [��]
'@����ֵ
'    ��������ɹ�������ֵ�ǻ������ľ����
'    �������ʧ�ܣ�����ֵΪNULL��Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'    ���ָ���Ļ�����󲻴��ڣ�����ʧ�ܣ�GetLastError����ERROR_FILE_NOT_FOUND��
'@��ע
'    OpenMutex��������������̴�ͬһ��mutex����ľ����ֻ�е�ĳ�������Ѿ�ʹ��CreateMutex���������˻������ʱ���ú����Ż�ɹ������ý��̿������κ���Ҫ����������ĺ�����ʹ�÷��صľ��������wait����������dwDesiredAccess������ָ���ķ������ơ�
'    ����ʹ��DuplicateHandle�������ƾ����ʹ��closeHandle�����رվ������������ֹʱ��ϵͳ�Զ��رվ�����������������һ��������ر�ʱ�����������١�
'    ������߳�Ӧ�ó�����뷴���������򿪺͹ر�һ��ָ���Ļ����������ܻ���־�̬����������������£����ʹ��CreateMutex������OpenMutex����ΪCreateMutex�ڻ���������ʱ�������ڲ�����ʱ��������
'@Ҫ��
'    Minimum supported client       Windows XP [desktop apps | UWP apps]
'    Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'    Header                         WinBase.h (include Windows.h)
'    Library                        kernel32.lib
'    dll                            kernel32.dll
'    Unicode and ANSI names         OpenMutexW (Unicode) and OpenMutexA (ANSI)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'@ԭ��
'    BOOL WINAPI CloseHandle(
'      _In_ HANDLE hObject
'    );
'@����
'    �ر�һ���򿪵Ķ�������
'@����
'    *hObject:
'    һ����Ч�Ĵ򿪵Ķ�������
'@����ֵ
'    ��������ɹ����򷵻�ֵ�Ƿ�0��
'    �������ʧ�ܣ��򷵻�ֵΪ0��Ҫ�����չ������Ϣ�������GetLastError��
'    ���Ӧ�ó����ڵ����������У���ô�ú������׳�һ���쳣����������յ��Ĳ�����Ч�ľ��ֵ��α���ֵ������ر�һ��������Σ�����CloseHandle�رյ���FindFirstFile�������صľ���������ǵ���FindClose�������ͻᷢ�����������
'@��ע
'    CloseHandle�ر����¶�����:
'        Access token ��������
'        Communications device ͨѶ�豸
'        Console input ����̨����
'        Console screen buffer ����̨��Ļ������
'        Event �¼�()
'        File �ļ�
'        File mapping �ļ�ӳ��
'        I/O completion port I / O��ɶ˿�
'        Job ����
'        Mailslot �ʲ�
'        Memory resource notification �ڴ���Դ��֪ͨ
'        Mutex ������
'        Named Pipe �����ܵ�
'        Pipe �ܵ�
'        Process ����
'        Semaphore �ź���
'        Thread �߳�
'        Transaction ����
'        Waitable Timer �ɵȴ���ʱ��
'    ������Щ����ĺ������ĵ�������������ɸö���ʱ��Ӧ��ʹ��CloseHandle���Լ��ڸþ���رպ�Զ���Ĵ����������ᷢ��ʲô�����
'        ͨ���� CloseHandle���ָ���Ķ�����ʧЧ���Զ���ľ���������еݼ�����ִ�ж�������顣
'        ����������һ��������رպ󣬶��󽫱���ϵͳ��ɾ�����й���Щ����Ĵ����ߺ�����ժҪ�������Kernel Objects.��
'    ͨ����Ӧ�ó���Ӧ��Ϊ���򿪵�ÿ���������һ�� CloseHandle��
'        ���ʹ�þ���ĺ���ʧ�ܲ�����ERROR_INVALID_HANDLE����ôͨ��û�б�Ҫ����CloseHandle����Ϊ�������ͨ����������Ѿ�ʧЧ��
'        Ȼ����һЩ����ʹ��ERROR_INVALID_HANDLE��ָʾ������������Ч��
'        ���磬����������ӱ��жϣ���ôһ����ͼ��������ʹ�þ���ĺ���ʧ�ܲ�����ERROR_INVALID_HANDLE ����Ϊ���ļ������ٿ��á�����������£�Ӧ�ó���Ӧ�ùرվ����
'    ���һ�������������ô�������ύ֮ǰ�����а󶨵�����ľ����Ӧ�ùرա�
'        ���һ��������ͨ��ʹ��FILE_FLAG_DELETE_ON_CLOSE��־����CreateFileTransacted �������򿪣���ô��Ӧ�ó���رվ���͵��� CommitTransaction֮ǰ�����ļ����ᱻɾ����
'        �й��������ĸ�����Ϣ����μ�Working With Transactions.��
'    �ر�һ���߳̾����������ֹ��ص��̣߳�Ҳ����ɾ���̶߳��󡣹ر�һ�����̾����������ֹ��صĽ��̣�Ҳ����ɾ�����̶���
'        Ҫɾ��һ���̶߳�����������ֹ�̣߳�Ȼ��ر��߳������еľ����Ҫ��ø�����Ϣ����μ�Terminating a Thread��
'        Ҫɾ�����̶�����������ֹ���̣�Ȼ��رս��̵����о����Ҫ�˽������Ϣ����μ�Terminating a Process��
'    ��ʹ��file mapping��Ȼ�Ǵ򿪵ģ��ر�һ���ļ�ӳ��ľ��Ҳ���Գɹ���Ҫ�˽������Ϣ�������Closing a File Mapping Object.��
'    ��Ҫʹ��CloseHandle�ر�һ���׽��֡��෴��ʹ��closesocket�����������ͷ����׽��ֹ�����������Դ�������׽��ֶ���ľ����Ҫ�˽������Ϣ�������Socket Closure��
'    ��Ҫʹ��CloseHandle�ر�һ���򿪵�ע������ľ�����෴��ʹ��RegCloseKey ������CloseHandle ����رն�ע������ľ�������ǲ��᷵��һ����������ʾ���ʧ�ܡ�
'@Ҫ��
'    Minimum supported client   Windows 2000 Professional [desktop apps | UWP apps]
'    Minimum supported server   Windows 2000 Server [desktop apps | UWP apps]
'    Minimum supported phone    Windows Phone 8
'    Header                     Winbase.h (include Windows.h)
'    Library                    kernel32.lib
'    dll                        kernel32.dll
Private Declare Function ReleaseMutex Lib "kernel32.dll" (ByVal mlnghMutex As Long) As Long
'@ԭ��
'    BOOL WINAPI ReleaseMutex(
'      _In_ HANDLE hMutex
'    );
'@����
'    �ͷ�ָ��������������Ȩ��
'@����
'hMutex _In_
'    �������ľ��?CreateMutex��OpenMutex��������������?
'@����ֵ
'    ��������ɹ�������ֵΪ���㡣
'    �������ʧ�ܣ�����ֵΪ�㡣Ҫ��ȡ��չ�Ĵ�����Ϣ�������GetLastError��
'@��ע
'    ��������̲߳�ӵ�л��������ReleaseMutex������ʧ�ܡ�
'    �߳�ͨ����bInitialOwner��������ΪTRUE������������󣬻����ڵ���ĳ���ȴ�����ʱָ�����ľ�����Ӷ���û�����������Ȩ�����̲߳�����Ҫӵ�л������ʱ��������ReleaseMutex�������Ա���һ���߳̿��Ի������Ȩ��
'    �߳̿����ڵ���ĳ���ȴ�����ʱָ�����Ѿ�ӵ�еĻ���������������������ִ�С�����Է�ֹ�߳��ڵȴ����Ѿ�ӵ�еĻ������ʱ�����Լ������ǣ�Ϊ���ͷ���������Ȩ���̱߳�����ÿ�λ������Ȩʱ����ReleaseMutexһ��(ͨ��CreateMutex��ȴ�����)��
'@Ҫ��
'    Minimum supported client       Windows XP [desktop apps | UWP apps]
'    Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'    Header                         WinBase.h on Windows XP, Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 and Windows Server 2008 R2 (include Windows.h);Synchapi.h on Windows 8 and Windows Server 2012
'    Library                        kernel32.lib
'    dll                            kernel32.dll


Private Const ERROR_SHARING_VIOLATION         As Long = &H20
Private Const ERROR_INSUFFICIENT_BUFFER       As Long = 122
Private Const ERROR_MORE_DATA                 As Long = 234
Private Const ERROR_NO_MORE_ITEMS             As Long = &H103
Private Const ERROR_KEY_NOT_FOUND             As Long = &H2
Private Const ERROR_ALREADY_EXISTS            As Long = 183&
Private Const ERROR_ACCESS_DENIED             As Long = 5&
Private Const ERROR_SUCCESS                   As Long = 0&

Private mlnghMutex              As Long
'--------------------------------------------------------------------------------------------------
'����           CheckMutex
'����           ��黥�����Ƿ���ڡ��������򴴽�
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'strMutexName   String                  ����������
'-------------------------------------------------------------------------------------------------
Public Function CheckMutex(ByVal strMutexName As String) As Boolean
    Dim lngRet As Long, lnghM As Long
    On Error GoTo ErrH
    Call PushMethod("ZLHisCrustCom.clsMutex.CheckMutex", "MUTEX_" & strMutexName)
    
    If mlnghMutex = 0 Then
        mlnghMutex = CreateMutexA(ByVal 0&, 1, "MUTEX_" & strMutexName)
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) Or (Err.LastDllError = ERROR_ACCESS_DENIED) Then
            CheckMutex = True
            Call CloseMutex
        Else
            CheckMutex = False
        End If
    Else
        CloseMutex
    End If
    Call PopMethod("ZLHisCrustCom.clsMutex.CheckMutex", CheckMutex)
    Exit Function
ErrH:
    If ErrorCenter("ZLHisCrustCom.clsMutex.CheckMutex") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'����           CloseMutex
'����           �رջ�������
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Private Function CloseMutex() As Boolean
    Dim lngRet As Long
 
    On Error GoTo ErrH
    Call PushMethod("ZLHisCrustCom.clsMutex.CloseMutex")
    If (mlnghMutex <> 0) Then
        lngRet = ReleaseMutex(mlnghMutex)
        lngRet = CloseHandle(mlnghMutex)
    End If
    If lngRet = 1 Then
        CloseMutex = False
        mlnghMutex = 0
    Else
        CloseMutex = True
    End If
    Call PopMethod("ZLHisCrustCom.clsMutex.CloseMutex", CloseMutex)
    Exit Function
ErrH:
    If ErrorCenter("ZLHisCrustCom.clsMutex.CloseMutex") = 1 Then
        Resume
    End If
End Function

Private Sub Class_Terminate()
    Call CloseMutex
End Sub