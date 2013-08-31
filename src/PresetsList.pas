unit PresetsList;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls,Registry;

type
  TManagePresets = class(TForm)
    Panel1: TPanel;
    ListBox1: TListBox;
    Panel2: TPanel;
    Edit1: TEdit;
    Panel3: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    procedure Panel2Resize(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ListBox1KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
    Reg:TRegistry;
  public
    { Public declarations }
  end;

var
  ManagePresets: TManagePresets;

implementation
uses any_2_fb_dialog;
{$R *.dfm}

procedure TManagePresets.Panel2Resize(Sender: TObject);
begin
  Edit1.Width:=Panel2.Width;
end;

procedure TManagePresets.FormCreate(Sender: TObject);
Var
 I:Integer;
begin
  Reg:=TRegistry.Create;
  Try
    Reg.OpenKeyReadOnly('Software\Grib Soft\Any to FB2\1.0\presets');
    Reg.GetKeyNames(ListBox1.Items);
    Reg.CloseKey;
    I:=0;
    While ListBox1.Items.IndexOf('Untitled'+IntToStr(I))<>-1 do
    Begin
      Edit1.Text:='Untitled'+IntToStr(I);
      inc(I);
    end;
  finally
  end;
end;

procedure TManagePresets.FormDestroy(Sender: TObject);
begin
  Reg.Free;
end;

procedure TManagePresets.Button1Click(Sender: TObject);
begin
  If UpperCase(Edit1.Text)='DEFAULT' then
  Begin
    MessageBox(Handle,'You can''t use "Default" name, sorry','Error',0);
    Exit;
  end;
  TUpenFileModalDialog(Owner).SaveSettings(Edit1.Text);
  If ListBox1.Items.IndexOf(Edit1.Text)=-1 then
    ListBox1.Items.Add(Edit1.Text);
end;

procedure TManagePresets.ListBox1Click(Sender: TObject);
Var
  I:Integer;
begin
  For I:=0 to ListBox1.Items.Count-1 do
    if ListBox1.Selected[I] then
    Begin
      Edit1.Text:=ListBox1.Items[I];
      Break;
    end;
end;

procedure TManagePresets.Button2Click(Sender: TObject);
begin
  ListBox1Click(Self);
  If MessageBox(Handle,PChar('Are you shure you want to delete preset called "'+Edit1.Text+'"?'),'Confirm',MB_YESNO or MB_ICONQUESTION)<>IDYES then
    Exit;
  Reg.DeleteKey('Software\Grib Soft\Any to FB2\1.0\presets\'+Edit1.Text);
  ListBox1.Items.Delete(ListBox1.Items.IndexOf(Edit1.Text));
end;

procedure TManagePresets.ListBox1KeyPress(Sender: TObject; var Key: Char);
begin
  ListBox1Click(Self);
end;

end.
