unit ExportToExcel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComObj, DB, DBTables, Grids, DBGrids,
  ExtCtrls, ComCtrls, TabNotBk;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    BtnExport: TBitBtn;
    BtnDB: TBitBtn;
    TableDB: TTable;
    DataSource: TDataSource;
    OpenDialog: TOpenDialog;
    TabbedNotebook: TTabbedNotebook;
    DBGrid: TDBGrid;
    procedure BtnExportClick(Sender: TObject);
    procedure BtnDBClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.BtnExportClick(Sender: TObject);
var
 XL, XArr: Variant;
 i: Integer;
 j: Integer;
begin
 {не забудьте включить ComObj в список используемых модулей}
 // Создаем массив элементов, полученных в результате запроса
 XArr:=VarArrayCreate([1,TableDB.FieldCount],varVariant);
 XL:=CreateOLEObject('Excel.Application');     // Создание OLE объекта
 XL.WorkBooks.add;
 XL.visible:=true;

 j := 1;
 TableDB.First;
 while not TableDB.Eof do
  begin
   i:=1;
   while i<=TableDB.FieldCount do
    begin
     XArr[i] := TableDB.Fields[i-1].Value;
     i := i+1;
    end;
   XL.Range['A'+IntToStr(j),
   CHR(64+TableDB.FieldCount)+IntToStr(j)].Value := XArr;
   TableDB.Next;
   j:=j+1;
  end;
 XL.Range['A1',CHR(64+TableDB.FieldCount)+IntToStr(j)].select;
 // XL.cells.select;                     // Выбираем все
 XL.Selection.Font.Name:='Arial cur';
 XL.Selection.Font.Size:=10;
 XL.selection.Columns.AutoFit;
 XL.Range['A1','A1'].select;
end;

procedure TForm1.BtnDBClick(Sender: TObject);
begin
 if OpenDialog.Execute
 then
  begin
   TableDB.Active:=false;
   TableDB.TableName:=OpenDialog.FileName;
   TableDB.Active:=true;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
 OpenDialog.InitialDir:=ExtractFilePath(Application.ExeName);
end;

end.
