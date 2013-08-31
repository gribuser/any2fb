unit REEditor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TReEditorForm = class(TForm)
    Memo2: TMemo;
    Panel3: TPanel;
    Panel2: TPanel;
    Button1: TButton;
    Button2: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ReEditorForm: TReEditorForm;

implementation

{$R *.dfm}

procedure TReEditorForm.FormResize(Sender: TObject);
begin
  Edit1.Width:=Panel3.Width-Panel2.Width ;
end;

end.
