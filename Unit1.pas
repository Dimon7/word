unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, ComObj;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Edit5: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    Edit6: TEdit;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Memo1: TMemo;
    Edit7: TEdit;
    Edit8: TEdit;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Edit9: TEdit;
    Edit10: TEdit;
    Edit11: TEdit;
    SaveDialog1: TSaveDialog;
    ComboBox1: TComboBox;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure ComboBox1Click(Sender: TObject);


  private
    word : variant;
    dir : string;


    { Private declarations }
  public

    { Public declarations }
  end;

var
  Form1: TForm1;
  number : integer;
  myFile : TextFile;
  text   : string;
  dorf : string;
implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
   var saveDialog : TSaveDialog;
       Replace : integer;
       text : string;
  {************************************************}
  procedure FindAndReplace (SearchStr, ReplaceStr : string);

  begin
    word.Selection.Find.Text := SearchStr;


    word.Selection.Find.Replacement.Text := ReplaceStr;
    word.Selection.Find.Execute (Replace := 2);
  end;
    {************************************************}
begin
    try
      Word := CreateOleObject('Word.Application');
    except
      MessageBox (Handle, 'Не установлен Microsoft Office Word. Формирование ' +
                          'заявки невозможно.', 'Ошибка',
                          MB_OK or MB_ICONERROR);
      exit;
    end;


  AssignFile(myFile, 'Number.txt');

  Reset(myFile);
  while not Eof(myFile) do ReadLn(myFile, text);

  number := StrToInt(text);

  ShowMessage(IntToStr(number));

  Edit11.Text := IntToStr(number);


  dir := GetCurrentDir;
  word.documents.open(dir + '\template.docx');

  FindAndReplace ('%Number%', Edit11.Text );
  FindAndReplace ('%Dorf%', dorf); {edit1}
  FindAndReplace ('%PIB%', edit2.Text);
  FindAndReplace ('%KOD%', edit3.Text);
  FindAndReplace ('%Sum1%', edit4.Text);
  FindAndReplace ('%Sum2%', edit5.Text);
  FindAndReplace ('%Sum3%', edit6.Text);
  FindAndReplace ('%P%', edit8.Text);
  FindAndReplace ('%idk%', edit7.Text);
  FindAndReplace ('%Main%', edit9.Text);
  FindAndReplace ('%Kasa%', edit10.Text);

  saveDialog := TSaveDialog.Create(self);
  saveDialog.Title := 'Збережіть ваш файл';
  saveDialog.InitialDir := dir;
  saveDialog.Filter := 'Word file|*.docx';
  saveDialog.DefaultExt := 'docx';
  saveDialog.FilterIndex := 1;

  saveDialog.FileName := Edit11.Text + '_' + Edit2.Text;
  if saveDialog.Execute then
                         word.Activedocument.SaveAs(saveDialog.FileName)
                        else ShowMessage('Збереження було відмінено');

  saveDialog.Free;
  inc(number);

  Rewrite(myFile);
  Append(myFile);
  Write(myFile, IntToStr(number));
  ShowMessage(IntToStr(number));

  CloseFile(myFile);

end;

procedure TForm1.ComboBox1Click(Sender: TObject);
begin

   if ComboBox1.ItemIndex = 0 then dorf := 'Михальча';
   if ComboBox1.ItemIndex = 1 then dorf := 'Заволока';
   if ComboBox1.ItemIndex = 2 then dorf := 'Спаська';
   if ComboBox1.ItemIndex = 3 then dorf := 'Дубове';

end;

end.
