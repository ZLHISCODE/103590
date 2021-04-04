program RecordAudio;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  ZLDSVideoProcess_TLB in 'ZLDSVideoProcess_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
