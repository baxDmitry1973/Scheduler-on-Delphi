program FreeBackuper;

uses
  Forms,
  UMain in 'UMain.pas' {Form1},
  UTask in 'UTask.pas' {Form2},
  UAbout in 'UAbout.pas' {AboutBox};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
end.
