program Export_dbf_Tox_xls;

uses
  Forms,
  ExportToExcel in 'ExportToExcel.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.Title:='Ёкспорт из dbf в xls';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

