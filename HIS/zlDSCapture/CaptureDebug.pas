{*******************************************************************************
用于程序中的错误调试
创建人：TJH
创建日前：2009-11-26

描述：...


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
  DEBUG_TURN_ON: Boolean = True; //如果为FALSE，则所有设置了TDebug.DebugMsg的地方将不起作用

{ TDebug }

class procedure TDebug.DebugMsg(const className, methodName, msg: WideString);
begin
  if DEBUG_TURN_ON then
    Application.MessageBox(PChar(String(DateTimeToStr(now) + ':' + className + '.' + methodName + ' [调试信息：' + msg + ']')), 'DEBUG', MB_OK + MB_ICONINFORMATION);
end;

class procedure TDebug.OutputDebug(const debug: String);
begin
  {$ifdef DEBUG}
    OutputDebugString(PAnsiChar(debug));
  {$endif}
end;

end.
