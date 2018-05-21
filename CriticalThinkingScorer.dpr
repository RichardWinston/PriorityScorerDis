program CriticalThinkingScorer;

uses
  Vcl.Forms,
  frmMainUnit2 in 'frmMainUnit2.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
