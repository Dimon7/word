unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, ComObj;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Label1: TLabel;
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
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  word : variant;

  procedure FindAndReplace (SearchStr, ReplaceStr : string);
  begin
    word.Selection.Find.Text := SearchStr;


    word.Selection.Find.Replacement.Text := ReplaceStr;
    word.Selection.Find.Execute (Replace := 2);
  end;
begin
    try
      Word := CreateOleObject('Word.Application');
    except
      MessageBox (Handle, '�� ���������� Microsoft Office Word. ������������ ' +
                          '������ ����������.', '������',
                          MB_OK or MB_ICONERROR);
      exit;
    end;
  word.documents.open ('F:\2.docx');
  FindAndReplace ('%Number%', edit11.Text);
  FindAndReplace ('%Dorf%', edit1.Text);
  FindAndReplace ('%PIB%', edit2.Text);
  FindAndReplace ('%KOD%', edit3.Text);
  FindAndReplace ('%Sum1%', edit4.Text);
  FindAndReplace ('%Sum2%', edit5.Text);
  FindAndReplace ('%Sum3%', edit6.Text);
  FindAndReplace ('%P%', edit8.Text);
  FindAndReplace ('%idk%', edit7.Text);
  FindAndReplace ('%Main%', edit9.Text);
  FindAndReplace ('%Kasa%', edit10.Text);


  word.ActiveDocument.SaveAs('F:\4.docx');
 // word.close;
//  word.Visible := true;
 // word := Unassigned;

end;

end.
