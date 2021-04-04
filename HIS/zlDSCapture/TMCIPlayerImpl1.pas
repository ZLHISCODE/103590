unit TMCIPlayerImpl1;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, ActiveX, Classes, Controls, Graphics, Menus, Forms, StdCtrls,
  ComServ, StdVCL, AXCtrls, ZLDSVideoProcess_TLB, Audio;

type
  TTMCIPlayer = class(TActiveXControl, ITMCIPlayer)
  private
    { Private declarations }
    FDelphiControl: TPlayer;
    FEvents: ITMCIPlayerEvents;
  protected
    { Protected declarations }
    procedure DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage); override;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    procedure InitializeControl; override;
    function Get_LineColor: OLE_COLOR; safecall;
    function Get_PlayState: TAxMCIAudioState; safecall;
    function Get_BitsPerSample: TAxBPS; safecall;
    function Get_BufferCount: TAxBufferCount; safecall;
    function Get_Channels: TAxChannels; safecall;
    function Get_BackColor: OLE_COLOR; safecall;
    function Get_CurTime: Integer; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    function Get_DrawFrequency: Integer; safecall;
    function Get_Enabled: WordBool; safecall;
    function Get_MaxColor: OLE_COLOR; safecall;
    function Get_OutputDeviceCount: Integer; safecall;
    function Get_SampleCount: Integer; safecall;
    function Get_SepCtrl: WordBool; safecall;
    function Get_TimeLen: Integer; safecall;
    function Get_Title: WideString; safecall;
    function Get_Visible: WordBool; safecall;
    function PlayFile(const fileName: WideString;
      NoOfRepeats: Integer): WordBool; safecall;
    procedure InitiateAction; safecall;
    procedure PausePlay; safecall;
    procedure RestartPlay; safecall;
    procedure Set_LineColor(Value: OLE_COLOR); safecall;
    procedure Set_BitsPerSample(Value: TAxBPS); safecall;
    procedure Set_BufferCount(Value: TAxBufferCount); safecall;
    procedure Set_Channels(Value: TAxChannels); safecall;
    procedure Set_BackColor(Value: OLE_COLOR); safecall;
    procedure Set_CurTime(Value: Integer); safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure Set_DrawFrequency(Value: Integer); safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Set_MaxColor(Value: OLE_COLOR); safecall;
    procedure Set_SampleCount(Value: Integer); safecall;
    procedure Set_SepCtrl(Value: WordBool); safecall;
    procedure Set_Title(const Value: WideString); safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    procedure StopPlay; safecall;
    function Get_PlayDeviceId: Integer; safecall;
    procedure Set_PlayDeviceId(Value: Integer); safecall;
    function Get_NoSamples: Integer; safecall;
    function Get_SampleRate: Integer; safecall;
    procedure Set_NoSamples(Value: Integer); safecall;
    procedure Set_SampleRate(Value: Integer); safecall;
    function Get_PlayPosition: Double; safecall;
    function Get_PlaySize: Double; safecall;
    procedure Set_PlayPosition(Value: Double); safecall;
    function Get_AppHandle: Integer; safecall;
    function Get_OutputDeviceName(Index: Integer): WideString; safecall;
    procedure FreeRes; safecall;
    procedure Set_AppHandle(Value: Integer); safecall;
    function Get_Handle: OLE_HANDLE; safecall;
    function Get_Hint: WideString; safecall;
    function Get_ShowHint: WordBool; safecall;
    procedure Set_Hint(const Value: WideString); safecall;
    procedure Set_ShowHint(Value: WordBool); safecall;
    function Get_Height: Integer; safecall;
    function Get_Left: Integer; safecall;
    function Get_Top: Integer; safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Height(Value: Integer); safecall;
    procedure Set_Left(Value: Integer); safecall;
    procedure Set_Top(Value: Integer); safecall;
    procedure Set_Width(Value: Integer); safecall;
    procedure SetSubComponent(IsSubComponent: WordBool); safecall;
    procedure GetVolume(var LeftVolume, RightVolume: Integer); safecall;
    procedure SetVolume(LeftVolume, RightVolume: Integer); safecall;
    function Get_PlayCurTime: Integer; safecall;
    procedure Set_PlayCurTime(Value: Integer); safecall;
    function Get_PlayTimeLen: Integer; safecall;
    procedure ShowFormatDialog; safecall;
    function Get_FormatTag: Integer; safecall;
    procedure Set_FormatTag(Value: Integer); safecall;
    function Get_ErrorMsg: WideString; safecall;
  end;

implementation

uses ComObj;

{ TTMCIPlayer }

procedure TTMCIPlayer.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  {TODO: Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_TMCIPlayerPage); }
end;

procedure TTMCIPlayer.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as ITMCIPlayerEvents;
end;

procedure TTMCIPlayer.InitializeControl;
begin
  FDelphiControl := Control as TPlayer;
end;

function TTMCIPlayer.Get_LineColor: OLE_COLOR;
begin
  Result := OLE_COLOR(FDelphiControl.AudioLineColor);
end;

function TTMCIPlayer.Get_PlayState: TAxMCIAudioState;
begin
  Result := TAxMCIAudioState(FDelphiControl.AudioState);
end;

function TTMCIPlayer.Get_BitsPerSample: TAxBPS;
begin
  Result := TAxBPS(FDelphiControl.BitsPerSample);
end;

function TTMCIPlayer.Get_BufferCount: TAxBufferCount;
begin
  Result := TAxBufferCount(FDelphiControl.BufferCount);
end;

function TTMCIPlayer.Get_Channels: TAxChannels;
begin
  Result := TAxChannels(FDelphiControl.Channels);
end;

function TTMCIPlayer.Get_BackColor: OLE_COLOR;
begin
  Result := OLE_COLOR(FDelphiControl.Color);
end;

function TTMCIPlayer.Get_CurTime: Integer;
begin
  Result := FDelphiControl.CurTime;
end;

function TTMCIPlayer.Get_DoubleBuffered: WordBool;
begin
  Result := FDelphiControl.DoubleBuffered;
end;

function TTMCIPlayer.Get_DrawFrequency: Integer;
begin
  Result := FDelphiControl.DrawFrequency;
end;

function TTMCIPlayer.Get_Enabled: WordBool;
begin
  Result := FDelphiControl.Enabled;
end;

function TTMCIPlayer.Get_MaxColor: OLE_COLOR;
begin
  Result := OLE_COLOR(FDelphiControl.MaxColor);
end;

function TTMCIPlayer.Get_OutputDeviceCount: Integer;
begin
  Result := FDelphiControl.OutputDeviceCount;
end;

function TTMCIPlayer.Get_SampleCount: Integer;
begin
  Result := FDelphiControl.SampleCount;
end;

function TTMCIPlayer.Get_SepCtrl: WordBool;
begin
  Result := FDelphiControl.SepCtrl;
end;

function TTMCIPlayer.Get_TimeLen: Integer;
begin
  Result := FDelphiControl.TimeLen;
end;

function TTMCIPlayer.Get_Title: WideString;
begin
  Result := WideString(FDelphiControl.Title);
end;

function TTMCIPlayer.Get_Visible: WordBool;
begin
  Result := FDelphiControl.Visible;
end;

function TTMCIPlayer.PlayFile(const fileName: WideString;
  NoOfRepeats: Integer): WordBool;
begin
  Result := FDelphiControl.PlayFile(FileName, NoOfRepeats);
end;

procedure TTMCIPlayer.InitiateAction;
begin
  FDelphiControl.InitiateAction;
end;

procedure TTMCIPlayer.PausePlay;
begin
  FDelphiControl.PausePlay;
end;

procedure TTMCIPlayer.RestartPlay;
begin
  FDelphiControl.RestartPlay;
end;

procedure TTMCIPlayer.Set_LineColor(Value: OLE_COLOR);
begin
  FDelphiControl.AudioLineColor := TColor(Value);
end;

procedure TTMCIPlayer.Set_BitsPerSample(Value: TAxBPS);
begin
  FDelphiControl.BitsPerSample := TBPS(Value);
end;

procedure TTMCIPlayer.Set_BufferCount(Value: TAxBufferCount);
begin
  FDelphiControl.BufferCount := TBufferCount(Value);
end;

procedure TTMCIPlayer.Set_Channels(Value: TAxChannels);
begin
  FDelphiControl.Channels := TChannels(Value);
end;

procedure TTMCIPlayer.Set_BackColor(Value: OLE_COLOR);
begin
  FDelphiControl.Color := TColor(Value);
end;

procedure TTMCIPlayer.Set_CurTime(Value: Integer);
begin
  FDelphiControl.CurTime := Value;
end;

procedure TTMCIPlayer.Set_DoubleBuffered(Value: WordBool);
begin
  FDelphiControl.DoubleBuffered := Value;
end;

procedure TTMCIPlayer.Set_DrawFrequency(Value: Integer);
begin
  FDelphiControl.DrawFrequency := Value;
end;

procedure TTMCIPlayer.Set_Enabled(Value: WordBool);
begin
  FDelphiControl.Enabled := Value;
end;

procedure TTMCIPlayer.Set_MaxColor(Value: OLE_COLOR);
begin
  FDelphiControl.MaxColor := TColor(Value);
end;

procedure TTMCIPlayer.Set_SampleCount(Value: Integer);
begin
  FDelphiControl.SampleCount := Value;
end;

procedure TTMCIPlayer.Set_SepCtrl(Value: WordBool);
begin
  FDelphiControl.SepCtrl := Value;
end;

procedure TTMCIPlayer.Set_Title(const Value: WideString);
begin
  FDelphiControl.Title := String(Value);
end;

procedure TTMCIPlayer.Set_Visible(Value: WordBool);
begin
  FDelphiControl.Visible := Value;
end;

procedure TTMCIPlayer.StopPlay;
begin
  FDelphiControl.StopPlay;
end;

function TTMCIPlayer.Get_PlayDeviceId: Integer;
begin
  Result := FDelphiControl.DeviceId;
end;

procedure TTMCIPlayer.Set_PlayDeviceId(Value: Integer);
begin
  FDelphiControl.DeviceId := Value;
end;

function TTMCIPlayer.Get_NoSamples: Integer;
begin
  Result := FDelphiControl.NoSamples;
end;

function TTMCIPlayer.Get_SampleRate: Integer;
begin
  Result := FDelphiControl.SampleRate;
end;

procedure TTMCIPlayer.Set_NoSamples(Value: Integer);
begin
  FDelphiControl.NoSamples := Value;
end;

procedure TTMCIPlayer.Set_SampleRate(Value: Integer);
begin
  FDelphiControl.SampleRate := Value;
end;

function TTMCIPlayer.Get_PlayPosition: Double;
begin
  Result := FDelphiControl.StreamPosition;
end;

function TTMCIPlayer.Get_PlaySize: Double;
begin
  Result := FDelphiControl.StreamSize;
end;

procedure TTMCIPlayer.Set_PlayPosition(Value: Double);
begin
  FDelphiControl.StreamPosition := Round(Value);
end;

function TTMCIPlayer.Get_AppHandle: Integer;
begin
  Result := Application.Handle;
end;

function TTMCIPlayer.Get_OutputDeviceName(Index: Integer): WideString;
begin
  Result := FDelphiControl.OutputDeviceName[index];
end;

procedure TTMCIPlayer.FreeRes;
begin
  try
    if FDelphiControl.AudioState <> masStop then FDelphiControl.StopPlay;
  except end;
end;

procedure TTMCIPlayer.Set_AppHandle(Value: Integer);
begin
  Application.Handle := Value;
end;

function TTMCIPlayer.Get_Handle: OLE_HANDLE;
begin
  Result := FDelphiControl.Handle;
end;

function TTMCIPlayer.Get_Hint: WideString;
begin
  Result := WideString(FDelphiControl.Hint);
end;

function TTMCIPlayer.Get_ShowHint: WordBool;
begin
  Result := FDelphiControl.ShowHint;
end;

procedure TTMCIPlayer.Set_Hint(const Value: WideString);
begin
  FDelphiControl.Hint := String(Value);
end;

procedure TTMCIPlayer.Set_ShowHint(Value: WordBool);
begin
  FDelphiControl.ShowHint := Value;
end;

function TTMCIPlayer.Get_Height: Integer;
begin
  Result := FDelphiControl.Height;
end;

function TTMCIPlayer.Get_Left: Integer;
begin
  Result := FDelphiControl.Left;
end;

function TTMCIPlayer.Get_Top: Integer;
begin
  Result := FDelphiControl.Top;
end;

function TTMCIPlayer.Get_Width: Integer;
begin
  Result := FDelphiControl.Width;
end;

procedure TTMCIPlayer.Set_Height(Value: Integer);
begin
  FDelphiControl.Height := Value;
end;

procedure TTMCIPlayer.Set_Left(Value: Integer);
begin
  FDelphiControl.Left := Value;
end;

procedure TTMCIPlayer.Set_Top(Value: Integer);
begin
  FDelphiControl.Top := Value;
end;

procedure TTMCIPlayer.Set_Width(Value: Integer);
begin
  FDelphiControl.Width := Value;
end;

procedure TTMCIPlayer.SetSubComponent(IsSubComponent: WordBool);
begin
  FDelphiControl.SetSubComponent(IsSubComponent);
end;

procedure TTMCIPlayer.GetVolume(var LeftVolume, RightVolume: Integer);
var
  lv, rv: Word;
begin
  FDelphiControl.GetVolume(lv, rv);
  
  LeftVolume := lv;
  RightVolume := rv;
end;

procedure TTMCIPlayer.SetVolume(LeftVolume, RightVolume: Integer);
begin
  FDelphiControl.SetVolume(LeftVolume, RightVolume); 
end;

function TTMCIPlayer.Get_PlayCurTime: Integer;
begin
  Result := FDelphiControl.CurTime;
end;

procedure TTMCIPlayer.Set_PlayCurTime(Value: Integer);
begin
  FDelphiControl.CurTime := Value;
end;

function TTMCIPlayer.Get_PlayTimeLen: Integer;
begin
  Result := FDelphiControl.TimeLen;
end;

procedure TTMCIPlayer.ShowFormatDialog;
begin
  FDelphiControl.ShowFormatDialog;
end;

function TTMCIPlayer.Get_FormatTag: Integer;
begin
  Result := FDelphiControl.FormatTag;
end;

procedure TTMCIPlayer.Set_FormatTag(Value: Integer);
begin
  FDelphiControl.FormatTag := Value;
end;

function TTMCIPlayer.Get_ErrorMsg: WideString;
begin
  Result := FDelphiControl.ErrorMsg;
end;

initialization
  TActiveXControlFactory.Create(
    ComServer,
    TTMCIPlayer,
    TPlayer,
    Class_TMCIPlayer,
    4,
    '',
    0,
    tmApartment);
end.
