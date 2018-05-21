program CriticalThinkingScorerJuly3;

uses
  Vcl.Forms,
  frmMainUnit5 in 'frmMainUnit5.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
