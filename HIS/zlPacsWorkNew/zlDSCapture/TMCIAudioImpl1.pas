unit TMCIAudioImpl1;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, ActiveX, Classes, Controls, Graphics, Menus, Forms, StdCtrls,
  ComServ, StdVCL, AXCtrls, ZLDSVideoProcess_TLB, Audio, Dialogs;

type
  TTMCIAudio = class(TActiveXControl, ITMCIAudio)
  private
    { Private declarations }
    FDelphiControl: TRecorder;
    FEvents: ITMCIAudioEvents;
  protected
    { Protected declarations }
    procedure DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage); override;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    procedure InitializeControl; override;
    function Get_BackColor: OLE_COLOR; safecall;
    function Get_DrawFrequency: Integer; safecall;
    function Get_Enabled: WordBool; safecall;
    function Get_LineColor: OLE_COLOR; safecall;
    function Get_MaxColor: OLE_COLOR; safecall;
    function Get_SampleCount: Integer; safecall;
    function Get_Visible: WordBool; safecall;
    function StartRecord: WordBool; safecall;
    procedure InitiateAction; safecall;
    procedure PauseRecord; safecall;
    procedure RestartRecord; safecall;
    procedure Set_BackColor(Value: OLE_COLOR); safecall;
    procedure Set_DrawFrequency(Value: Integer); safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Set_LineColor(Value: OLE_COLOR); safecall;
    procedure Set_MaxColor(Value: OLE_COLOR); safecall;
    procedure Set_SampleCount(Value: Integer); safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    procedure SetSubComponent(IsSubComponent: WordBool); safecall;
    procedure StopRecord; safecall;
    function Get_Channels: TAxChannels; safecall;
    procedure Set_Channels(Value: TAxChannels); safecall;
    function Get_BitsPerSample: TAxBPS; safecall;
    procedure Set_BitsPerSample(Value: TAxBPS); safecall;
    function Get_SampleRate: Integer; safecall;
    procedure Set_SampleRate(Value: Integer); safecall;
    function Get_NoSamples: Integer; safecall;
    function Get_RecordFile: WideString; safecall;
    function Get_SplitChannels: WordBool; safecall;
    function Get_Triggered: WordBool; safecall;
    function Get_TrigLevel: Integer; safecall;
    procedure Set_NoSamples(Value: Integer); safecall;
    procedure Set_RecordFile(const Value: WideString); safecall;
    procedure Set_SplitChannels(Value: WordBool); safecall;
    procedure Set_Triggered(Value: WordBool); safecall;
    procedure Set_TrigLevel(Value: Integer); safecall;
    function Get_RecordSize: Double; safecall;
    function Get_RecordCurTime: Integer; safecall;
    procedure Set_RecordCurTime(Value: Integer); safecall;
    function Get_RecordPostion: Double; safecall;
    procedure Set_RecordPostion(Value: Double); safecall;
    function Get_AudioDeviceId: Integer; safecall;
    procedure Set_AudioDeviceId(Value: Integer); safecall;
    function Get_Handle: OLE_HANDLE; safecall;
    function Get_Height: Integer; safecall;
    function Get_Hint: WideString; safecall;
    function Get_Left: Integer; safecall;
    function Get_ShowHint: WordBool; safecall;
    function Get_Top: Integer; safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Height(Value: Integer); safecall;
    procedure Set_Hint(const Value: WideString); safecall;
    procedure Set_Left(Value: Integer); safecall;
    procedure Set_ShowHint(Value: WordBool); safecall;
    procedure Set_Top(Value: Integer); safecall;
    procedure Set_Width(Value: Integer); safecall;
    function Get_RecordState: TAxMCIAudioState; safecall;
    function Get_RecordTimeLen: Integer; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    function Get_AppHandle: Integer; safecall;
    procedure Set_AppHandle(Value: Integer); safecall;
    procedure FreeRes; safecall;
    function Get_BufferCount: TAxBufferCount; safecall;
    function Get_Title: WideString; safecall;
    procedure Set_BufferCount(Value: TAxBufferCount); safecall;
    procedure Set_Title(const Value: WideString); safecall;
    function Get_RecordInputCount: Integer; safecall;
    function Get_RecordInputName(index: Integer): WideString; safecall;
    procedure ShowFormatDialog; safecall;
    function Get_FormatTag: Integer; safecall;
    procedure Set_FormatTag(Value: Integer); safecall;
    function Get_CompRate: Integer; safecall;
    function Get_IsCompressWav: WordBool; safecall;
    procedure Set_CompRate(Value: Integer); safecall;
    procedure Set_IsCompressWav(Value: WordBool); safecall;
    function Get_ErrorMsg: WideString; safecall;
  end;

implementation

uses ComObj, SysUtils;

{ TTMCIAudio }

procedure TTMCIAudio.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  {TODO: Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_TMCIAudioPage); }
end;

procedure TTMCIAudio.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as ITMCIAudioEvents;
end;

procedure TTMCIAudio.InitializeControl;
begin
  FDelphiControl := Control as TRecorder;
end;


function TTMCIAudio.Get_BackColor: OLE_COLOR;
begin
  Result := OLE_COLOR(FDelphiControl.Color);
end;


function TTMCIAudio.Get_DrawFrequency: Integer;
begin
  Result := FDelphiControl.DrawFrequency;
end;

function TTMCIAudio.Get_Enabled: WordBool;
begin
  Result := FDelphiControl.Enabled;
end;

function TTMCIAudio.Get_LineColor: OLE_COLOR;
begin
  Result := OLE_COLOR(FDelphiControl.AudioLineColor);
end;

function TTMCIAudio.Get_MaxColor: OLE_COLOR;
begin
  Result := OLE_COLOR(FDelphiControl.MaxColor);
end;

function TTMCIAudio.Get_SampleCount: Integer;
begin
  Result := FDelphiControl.SampleCount;
end;

function TTMCIAudio.Get_Visible: WordBool;
begin
  Result := FDelphiControl.Visible;
end;

function TTMCIAudio.StartRecord: WordBool;
begin
  Result := FDelphiControl.StartRecord;
end;

procedure TTMCIAudio.InitiateAction;
begin
  FDelphiControl.InitiateAction;
end;

procedure TTMCIAudio.PauseRecord;
begin
  FDelphiControl.PauseRecord;
end;

procedure TTMCIAudio.RestartRecord;
begin
  FDelphiControl.RestartRecord;
end;

procedure TTMCIAudio.Set_BackColor(Value: OLE_COLOR);
begin
  FDelphiControl.Color := TColor(Value);
end;


procedure TTMCIAudio.Set_DrawFrequency(Value: Integer);
begin
  FDelphiControl.DrawFrequency := Value;
end;

procedure TTMCIAudio.Set_Enabled(Value: WordBool);
begin
  FDelphiControl.Enabled := Value;
end;

procedure TTMCIAudio.Set_LineColor(Value: OLE_COLOR);
begin
  FDelphiControl.AudioLineColor := TColor(Value);
end;

procedure TTMCIAudio.Set_MaxColor(Value: OLE_COLOR);
begin
  FDelphiControl.MaxColor := TColor(Value);
end;

procedure TTMCIAudio.Set_SampleCount(Value: Integer);
begin
  FDelphiControl.SampleCount := Value;
end;

procedure TTMCIAudio.Set_Visible(Value: WordBool);
begin
  FDelphiControl.Visible := Value;
end;

procedure TTMCIAudio.SetSubComponent(IsSubComponent: WordBool);
begin
  FDelphiControl.SetSubComponent(IsSubComponent);
end;

procedure TTMCIAudio.StopRecord;
begin
  FDelphiControl.StopRecord;
end;

function TTMCIAudio.Get_Channels: TAxChannels;
begin
  Result := TAxChannels(FDelphiControl.Channels);
end;

procedure TTMCIAudio.Set_Channels(Value: TAxChannels);
begin
  FDelphiControl.Channels := TChannels(value);
end;

function TTMCIAudio.Get_BitsPerSample: TAxBPS;
begin
  Result := TAxBPS(FDelphiControl.BitsPerSample);
end;

procedure TTMCIAudio.Set_BitsPerSample(Value: TAxBPS);
begin
  FDelphiControl.BitsPerSample := TBPS(Value);
end;

function TTMCIAudio.Get_SampleRate: Integer;
begin
  Result := FDelphiControl.SampleRate;
end;

procedure TTMCIAudio.Set_SampleRate(Value: Integer);
begin
  FDelphiControl.SampleRate := Value;
end;

function TTMCIAudio.Get_NoSamples: Integer;
begin
  Result := FDelphiControl.NoSamples;
end;

function TTMCIAudio.Get_RecordFile: WideString;
begin
  Result := FDelphiControl.RecordFile;
end;

function TTMCIAudio.Get_SplitChannels: WordBool;
begin
  Result := FDelphiControl.SplitChannels;
end;

function TTMCIAudio.Get_Triggered: WordBool;
begin
  Result := FDelphiControl.Triggered;
end;

function TTMCIAudio.Get_TrigLevel: Integer;
begin
  Result := FDelphiControl.TrigLevel;
end;

procedure TTMCIAudio.Set_NoSamples(Value: Integer);
begin
  FDelphiControl.SampleRate := Value;
end;

procedure TTMCIAudio.Set_RecordFile(const Value: WideString);
begin
  FDelphiControl.RecordFile := Value;
end;

procedure TTMCIAudio.Set_SplitChannels(Value: WordBool);
begin
  FDelphiControl.SplitChannels := Value;
end;

procedure TTMCIAudio.Set_Triggered(Value: WordBool);
begin
  FDelphiControl.Triggered := Value;
end;

procedure TTMCIAudio.Set_TrigLevel(Value: Integer);
begin
  FDelphiControl.TrigLevel := Value;
end;

function TTMCIAudio.Get_RecordSize: Double;
begin
  Result := FDelphiControl.StreamSize;
end;


function TTMCIAudio.Get_RecordCurTime: Integer;
begin
  Result := FDelphiControl.CurTime;
end;

procedure TTMCIAudio.Set_RecordCurTime(Value: Integer);
begin
  FDelphiControl.CurTime := Value;
end;

function TTMCIAudio.Get_RecordPostion: Double;
begin
  Result := FDelphiControl.StreamPostion;
end;

procedure TTMCIAudio.Set_RecordPostion(Value: Double);
begin
  FDelphiControl.StreamPostion := Round(Value);
end;

function TTMCIAudio.Get_AudioDeviceId: Integer;
begin
  Result := FDelphiControl.DeviceId;
end;

procedure TTMCIAudio.Set_AudioDeviceId(Value: Integer);
begin
  FDelphiControl.DeviceId := Value;
end;

function TTMCIAudio.Get_Handle: OLE_HANDLE;
begin
  Result := FDelphiControl.Handle;
end;

function TTMCIAudio.Get_Height: Integer;
begin
  Result := FDelphiControl.Height;
end;

function TTMCIAudio.Get_Hint: WideString;
begin
  Result := WideString(FDelphiControl.Hint);
end;

function TTMCIAudio.Get_Left: Integer;
begin
  Result := FDelphiControl.Left;
end;

function TTMCIAudio.Get_ShowHint: WordBool;
begin
  Result := FDelphiControl.ShowHint;
end;

function TTMCIAudio.Get_Top: Integer;
begin
  Result := FDelphiControl.Top;
end;

function TTMCIAudio.Get_Width: Integer;
begin
  Result := FDelphiControl.Width;
end;

procedure TTMCIAudio.Set_Height(Value: Integer);
begin
  FDelphiControl.Height := Value;
end;

procedure TTMCIAudio.Set_Hint(const Value: WideString);
begin
  FDelphiControl.Hint := String(Value);
end;

procedure TTMCIAudio.Set_Left(Value: Integer);
begin
  FDelphiControl.Left := Value;
end;

procedure TTMCIAudio.Set_ShowHint(Value: WordBool);
begin
  FDelphiControl.ShowHint := Value;
end;

procedure TTMCIAudio.Set_Top(Value: Integer);
begin
  FDelphiControl.Top := Value;
end;

procedure TTMCIAudio.Set_Width(Value: Integer);
begin
  FDelphiControl.Width := Value;
end;

function TTMCIAudio.Get_RecordState: TAxMCIAudioState;
begin
  Result := TAxMCIAudioState(FDelphiControl.AudioState);
end;

function TTMCIAudio.Get_RecordTimeLen: Integer;
begin
  Result := FDelphiControl.TimeLen;
end;

function TTMCIAudio.Get_DoubleBuffered: WordBool;
begin
  Result := FDelphiControl.DoubleBuffered;
end;

procedure TTMCIAudio.Set_DoubleBuffered(Value: WordBool);
begin
  FDelphiControl.DoubleBuffered := Value;
end;

function TTMCIAudio.Get_AppHandle: Integer;
begin
  Result := Application.Handle;
end;

procedure TTMCIAudio.Set_AppHandle(Value: Integer);
begin
  Application.Handle := Value;
end;

function TTMCIAudio.Get_BufferCount: TAxBufferCount;
begin
  Result := TAxBufferCount(FDelphiControl.BufferCount);
end;

function TTMCIAudio.Get_Title: WideString;
begin
  Result := WideString(FDelphiControl.Title);
end;

procedure TTMCIAudio.Set_BufferCount(Value: TAxBufferCount);
begin
  FDelphiControl.BufferCount := TBufferCount(Value);
end;

procedure TTMCIAudio.Set_Title(const Value: WideString);
begin
  FDelphiControl.Title := String(Value);
end;

procedure TTMCIAudio.FreeRes;
begin
  try
    if FDelphiControl.AudioState <> masStop then FDelphiControl.StopRecord;

    //if Assigned(FDelphiControl) then FreeAndNil(FDelphiControl);
  except end;
end;

function TTMCIAudio.Get_RecordInputCount: Integer;
begin
  Result := FDelphiControl.RecordInputCount;
end;

function TTMCIAudio.Get_RecordInputName(index: Integer): WideString;
begin
  Result := FDelphiControl.RecordInputName[index];
end;


procedure TTMCIAudio.ShowFormatDialog;
begin
  FDelphiControl.ShowFormatDialog;
end;

function TTMCIAudio.Get_FormatTag: Integer;
begin
  Result := FDelphiControl.FormatTag;
end;

procedure TTMCIAudio.Set_FormatTag(Value: Integer);
begin
  FDelphiControl.FormatTag := Value;
end;

function TTMCIAudio.Get_CompRate: Integer;
begin
  Result := FDelphiControl.Mp3CompressRate;
end;

function TTMCIAudio.Get_IsCompressWav: WordBool;
begin
  Result := FDelphiControl.IsCompressMp3;
end;

procedure TTMCIAudio.Set_CompRate(Value: Integer);
begin
  FDelphiControl.Mp3CompressRate := Value;
end;

procedure TTMCIAudio.Set_IsCompressWav(Value: WordBool);
begin
  FDelphiControl.IsCompressMp3 := Value;
end;

function TTMCIAudio.Get_ErrorMsg: WideString;
begin
  Result := WideString(FDelphiControl.ErrorMsg);
end;

initialization
  TActiveXControlFactory.Create(
    ComServer,
    TTMCIAudio,
    TRecorder,
    Class_TMCIAudio,
    3,
    '',
    0,
    tmApartment);
end.
