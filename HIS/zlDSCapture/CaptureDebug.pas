{*******************************************************************************
���ڳ����еĴ������
�����ˣ�TJH
������ǰ��2009-11-26

������...


*******************************************************************************}
unit CaptureDebug;

interface

{$define DEBUG}

uses
  Windows, Classes, SysUtils, Forms;

Type
  TDebug = class(TObject)
  public
    class procedure DebugMsg(const className, methodName, msg: WideString);
    class procedure OutputDebug(const debug: String);
  end;

implementation

const
  DEBUG_TURN_ON: Boolean = True; //���ΪFALSE��������������TDebug.DebugMsg�ĵط�����������

{ TDebug }

class procedure TDebug.DebugMsg(const className, methodName, msg: WideString);
begin
  if DEBUG_TURN_ON then
    Application.MessageBox(PChar(String(DateTimeToStr(now) + ':' + className + '.' + methodName + ' [������Ϣ��' + msg + ']')), 'DEBUG', MB_OK + MB_ICONINFORMATION);
end;

class procedure TDebug.OutputDebug(const debug: String);
begin
  {$ifdef DEBUG}
    OutputDebugString(PAnsiChar(debug));
  {$endif}
end;

end.
