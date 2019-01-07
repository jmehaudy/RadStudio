program ExtReyWin;

uses
  Vcl.Forms,
  Main in 'Main.pas' {frmExtReyWin};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmExtReyWin, frmExtReyWin);
  Application.Run;
end.
