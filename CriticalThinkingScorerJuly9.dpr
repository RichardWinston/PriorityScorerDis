program CriticalThinkingScorerJuly9;

uses
  Vcl.Forms,
  frmMainUnit6 in 'frmMainUnit6.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
