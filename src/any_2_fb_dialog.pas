unit any_2_fb_dialog;

interface

uses
  Windows, SysUtils, Graphics, Forms, Menus, Dialogs, ComCtrls, StdCtrls,
  Controls, ExtCtrls, Classes,MSXML2_TLB, Grids;

type

  TUpenFileModalDialog = class(TForm)
    Panel1: TPanel;
    Edit1: TEdit;
    Panel2: TPanel;
    Panel3: TPanel;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Panel4: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    OpenDialog1: TOpenDialog;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    CheckBox15: TCheckBox;
    CheckBox16: TCheckBox;
    CheckBox17: TCheckBox;
    CheckBox18: TCheckBox;
    Button5: TButton;
    CheckBox21: TCheckBox;
    ProgressBar1: TProgressBar;
    TabSheet4: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Panel5: TPanel;
    Button1: TButton;
    Panel6: TPanel;
    Edit3: TEdit;
    Label4: TLabel;
    RichEdit1: TRichEdit;
    PopupMenu1: TPopupMenu;
    Copy1: TMenuItem;
    CheckBox12: TCheckBox;
    CheckBox8: TCheckBox;
    CheckBox14: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox9: TCheckBox;
    RadioGroup1: TRadioGroup;
    CheckBox20: TCheckBox;
    CheckBox19: TCheckBox;
    Panel7: TPanel;
    GroupBox1: TGroupBox;
    CheckBox11: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    GroupBox2: TGroupBox;
    Label3: TLabel;
    Label6: TLabel;
    Label5: TLabel;
    CheckBox2: TCheckBox;
    CheckBox22: TCheckBox;
    UpDown1: TUpDown;
    Edit2: TEdit;
    CheckBox23: TCheckBox;
    Edit5: TEdit;
    Edit4: TEdit;
    SaveDialog1: TSaveDialog;
    ComboBox1: TComboBox;
    Label7: TLabel;
    CheckBox24: TCheckBox;
    StringGrid1: TStringGrid;
    StringGrid2: TStringGrid;
    Button6: TButton;
    Timer1: TTimer;
    procedure Button2Click(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Panel5Resize(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure CheckBox11Click(Sender: TObject);
    procedure CheckBox11KeyPress(Sender: TObject; var Key: Char);
    procedure CheckBox22Click(Sender: TObject);
    procedure CheckBox22KeyPress(Sender: TObject; var Key: Char);
    procedure Panel1Resize(Sender: TObject);
    procedure Panel6Resize(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox2KeyPress(Sender: TObject; var Key: Char);
    procedure StringGrid1GetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: String);
    procedure StringGrid1SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure StringGrid2TopLeftChanged(Sender: TObject);
    procedure StringGrid2Enter(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    ParentHandle:THandle;
    ConvertDone:Boolean;
    NeedSave:Boolean;
    TimerCounter:Integer;
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
    FormResult:Integer;
    OutVar:IXMLDOMDocument2;
    ResultDoc:IDispatch;
    constructor GetDOMDocument(AParent:THandle; var ModalResult:Integer;ANeedSave:Boolean);
    Procedure InformAdded(S:String);
    Procedure WarningAdded(S:String);
    Procedure ProgressState(Percent:Integer);
    Procedure LoadSettings(IsReload:Boolean;Name:String);
    Procedure SaveSettings(Name:String);
    Procedure LoadSetList;
    Procedure AddRow(Grid:TStringGrid);
  end;

var
  UpenFileModalDialog: TUpenFileModalDialog;

implementation
uses tx2fb,Registry, PresetsList,REEditor;
{$R *.dfm}
Const
  SearchPatt='+new search pattern';
  ReplacePatt='+new replace pattern';
Function LineIsEmpty(Grid:TStringGrid;I:Integer):Boolean;
Begin
  Result:=((Grid.Rows[I][0]='') or (Grid.Rows[I][0]=SearchPatt)) and
        ((Grid.Rows[I][1]='') or (Grid.Rows[I][1]=replacepatt));
end;
procedure TUpenFileModalDialog.Button2Click(Sender: TObject);
Var
  Params:TParceParams;
  Function GetREs(Grid:TStringGrid):String;
  Var
    I:Integer;
  Begin
    Result:='';
    For I:=0 to Grid.RowCount-1 do
      if not LineIsEmpty(Grid,I) then
        Result:=Result+Grid.Rows[I][0]+#13+Grid.Rows[I][1]+#13;
    If Result<>'' then
      Result[Length(Result)]:=#0;
  end;
begin
  FormResult:=mrOk;
  if ConvertDone then
  Begin
    if NeedSave then
    Begin
      If SaveDialog1.Execute then
        OutVar.Save(SaveDialog1.FileName)
      else
        exit;
    end;
    Close;
    Exit;
  end;
  If Edit1.Text='http://' then
  Begin
    MessageBox(Handle,'Choose a file or enter correct URL first!','No file to import!',mb_Ok or MB_ICONASTERISK);
    Exit;
  end;
  ComboBox1.Enabled:=False;
  Button1.Enabled:=False;
  Button2.Enabled:=False;
  Button3.Enabled:=False;
  TabSheet1.Enabled:=False;
  TabSheet2.Enabled:=False;
  TabSheet4.Enabled:=False;
  Edit1.Enabled:=False;
  PageControl1.ActivePageIndex:=3;
  RichEdit1.SelAttributes.Color:=clWindowText;
  RichEdit1.lines.Text:='Importing "'+Edit1.Text+'"...';
  Update;
  Screen.Cursor:=crHourGlass;
  Params.LoadFromString:=False;
  Params.NotDetectEncoding:=CheckBox9.Checked;
  Params.RemoveExternalLinks:=CheckBox2.Checked;
  Params.SkipAllImages:=CheckBox11.Checked;
  Params.SkipOffSiteImages:=CheckBox3.Checked;
  Params.KeepDynamicImages:=CheckBox4.Checked;
  Params.PreservForms:=CheckBox12.Checked;
  Params.NoAutodetectFileType:=RadioGroup1.ItemIndex<>0;
  Params.HeadLowCence:=RadioGroup1.ItemIndex=2;
  If Params.NoAutodetectFileType then
    Params.ConvertLong:=80
  else
    Params.ConvertLong:=2;
  Params.NoHeadDetect:=CheckBox5.Checked;
  Params.IgnoreSpaceAtStart:=CheckBox14.Checked;
  Params.NotDetectItalic:=CheckBox16.Checked;
  Params.NotDetectNotes:=CheckBox15.Checked;
  Params.NotConvertQotes:=CheckBox7.Checked;
  Params.NotConvertDef:=CheckBox8.Checked;
  Params.Skip2Lines:=CheckBox6.Checked;
  Params.NotSearchEpig:=CheckBox17.Checked;
  Params.NotRestoreParagr:=CheckBox18.Checked;
  Params.NotSearchDescription:=CheckBox19.Checked;
  Params.DownloadExternal:=CheckBox23.Checked;
  Try
    Params.LinksDeep:=StrToInt(Edit2.Text);
  except
    Params.LinksDeep:=1;
  end;
  If not CheckBox22.Checked then
    Params.LinksDeep:=0;
  If CheckBox20.Checked then
    Params.FixTrialsOver100:=1000
  else
    Params.FixTrialsOver100:=0;
  Params.SkipPoems:=CheckBox21.Checked;
  Params.REGExpOnStart:=PChar(GetREs(StringGrid2));
  Params.RegExpOnFinish:=PChar(GetREs(StringGrid1));;
  Params.HeaderMUSTRe:=PChar(Edit3.Text);
  Params.LinksFollow:=PChar(Edit4.Text);
  Params.LinksSkip:=PChar(Edit5.Text);

  OutVar:=ParseText(PChar(Edit1.Text),@Params,InformAdded,WarningAdded,ProgressState);
  Screen.Cursor:=crDefault;
  ComboBox1.Enabled:=True;
  Button2.Enabled:=True;
  Button3.Enabled:=True;
  Button1.Enabled:=True;
  TabSheet1.Enabled:=True;
  TabSheet2.Enabled:=True;
  TabSheet4.Enabled:=True;
  Edit1.Enabled:=True;
  If (OutVar <> Nil) and CheckBox24.Checked then
    Close
  else
    if OutVar = Nil  then
    Begin
      MessageBeep(mb_iconError);
      WarningAdded('Convertion failed!');
    end else
      Begin
        ConvertDone:=True;
        Button2.Caption:='Done';
        Button1.Enabled:=False;
        Edit1.Enabled:=False;
        RichEdit1.SelAttributes.Color:=clGreen;
        RichEdit1.lines.Add('Export finished.');
        FormResult:=mrCancel;
      end;
end;

procedure TUpenFileModalDialog.FormDeactivate(Sender: TObject);
begin
  BringToFront;
end;
Constructor TUpenFileModalDialog.GetDOMDocument;
Begin
  ConvertDone:=False;
  NeedSave:=ANeedSave;
  ParentHandle:=AParent;
  Create(Nil);
  FormCreate(Self);
  ShowModal;
  If FormResult=mrOk then
    ResultDoc:=OutVar
  else
    ResultDoc:=Nil;
  ModalResult:=FormResult;
//  Free;
end;

procedure TUpenFileModalDialog.Button4Click(Sender: TObject);
begin
  MessageBox(Handle,'Any2FB ActiveX control and FBE Plug-in by GribUser'#10'v 0.60 beta','About Any2FB',mb_Ok or MB_ICONINFORMATION);
end;

procedure TUpenFileModalDialog.Button1Click(Sender: TObject);
begin
  If not OpenDialog1.Execute then Exit;
  Edit1.Text:=OpenDialog1.FileName;
end;

procedure TUpenFileModalDialog.Panel5Resize(Sender: TObject);
begin
  StringGrid2.Height:=(TabSheet4.Height-Label1.Height*2-Panel6.Height) div 2;
  StringGrid1.ColWidths[0]:=(TabSheet4.Width-GetSystemMetrics(SM_CYHSCROLL)-23) div 2;
  StringGrid1.ColWidths[1]:=StringGrid1.ColWidths[0];
  StringGrid2.ColWidths[0]:=StringGrid1.ColWidths[0];
  StringGrid2.ColWidths[1]:=StringGrid1.ColWidths[0];
  Button6.Left:=StringGrid1.ColWidths[0]*2+4;
end;

procedure TUpenFileModalDialog.CheckBox1Click(Sender: TObject);
begin
  if ComboBox1.ItemIndex=0 then
    LoadSettings(True,'')
  else
    if ComboBox1.ItemIndex=1 then
      LoadSettings(True,'Default')
  else LoadSettings(True,ComboBox1.Text);
end;
procedure TUpenFileModalDialog.Button5Click(Sender: TObject);
begin
{  SaveSettings;
  ComboBox1.Checked:=True;
  CheckBox1Click(Nil);}
  ManagePresets:=TManagePresets.Create(Self);
  ManagePresets.SHowModal;
  ManagePresets.free;
  LoadSetList;
end;

Procedure TUpenFileModalDialog.InformAdded(S:String);
Begin
  RichEdit1.SelAttributes.Color:=clWindowText;
  RichEdit1.lines.Add(S);
//  RichEdit1.Selected[RichEdit1.lines.Count-1]:=True;
  ProgressBar1.Position:=0;
  Update;
end;

Procedure TUpenFileModalDialog.WarningAdded(S:String);
Begin
  RichEdit1.SelAttributes.Color:=clRed;
  RichEdit1.lines.Add(S);
  RichEdit1.SelAttributes.Color:=clWindowText;
//  RichEdit1.Selected[RichEdit1.lines.Count-1]:=True;
  ProgressBar1.Position:=0;
  Update;
end;
Procedure TUpenFileModalDialog.ProgressState;
Begin
  If Percent>99 then
    ProgressBar1.Position:=ProgressBar1.Position+1
  else
    ProgressBar1.Position:=Percent;
  Application.ProcessMessages;
end;

procedure TUpenFileModalDialog.Button3Click(Sender: TObject);
begin
  FormResult:=mrCancel;
end;

Procedure TUpenFileModalDialog.AddRow(Grid:TStringGrid);
Begin
  Grid.Rows[Grid.RowCount-1][0]:=SearchPatt;
  Grid.Rows[Grid.RowCount-1][1]:=ReplacePatt;
end;

Procedure TUpenFileModalDialog.LoadSettings;
Var
  Reg:Tregistry;
  Procedure LoadRegExps(Grid:TStringGrid);
  Var
    Items:TStringList;
    I:Integer;
  Begin
    Items:=TStringList.Create;
    Grid.RowCount:=1;
    Try
      Reg.GetValueNames(Items);
      For I:=0 to Items.Count-1 do
        If Pos('_',Items[I])=0 then
        Begin
          Grid.RowCount:=Grid.RowCount+1;
          Grid.Rows[Grid.RowCount-2][0]:=Reg.ReadString(Items[I]);
          Grid.Rows[Grid.RowCount-2][1]:=Reg.ReadString(Items[I]+'_');
        end;
    finally
      Items.Free;
    end;
    AddRow(Grid);
  end;
Begin

  Reg:=TRegistry.Create(KEY_READ);
  If (name='Default') and not Reg.KeyExists('software\Grib Soft\Any to FB2\1.0'+Name) then
  Begin
    CheckBox2.Checked:=false;
    CheckBox11.Checked:=False;
    CheckBox3.Checked:=False;
    CheckBox4.Checked:=False;
    CheckBox6.Checked:=False;
    CheckBox7.Checked:=False;
    CheckBox15.Checked:=False;
    CheckBox16.Checked:=False;
    CheckBox17.Checked:=False;
    CheckBox18.Checked:=False;
    CheckBox19.Checked:=False;
    CheckBox20.Checked:=False;
    CheckBox21.Checked:=False;
    CheckBox12.Checked:=False;
    CheckBox9.Checked:=False;
    RadioGroup1.ItemIndex:=0;
    CheckBox5.Checked:=False;
    CheckBox14.Checked:=False;
    CheckBox8.Checked:=False;
    Edit2.Text:='1';
    CheckBox23.Checked:=False;
    CheckBox22.Checked:=False;
    Edit3.Text:='';
    Edit4.Text:='';
    Edit5.Text:='';
    StringGrid2.RowCount:=1;
    StringGrid1.RowCount:=1;
    AddRow(StringGrid2);
    AddRow(StringGrid1);
    Exit;
  end;
  if Name<>'' then name:='\presets\'+name;
  Try
    Try
      if Reg.OpenKeyReadOnly('software\Grib Soft\Any to FB2\1.0'+Name) then
      Begin
        If not IsReload then
        Begin
//          CheckBox1.Checked:=Reg.ReadBool('Use defaults');
          CheckBox24.Checked:=Reg.ReadBool('Close on finish');
          Edit1.Text:=Reg.ReadString('LastOpenURI');
          PageControl1.ActivePageIndex:=Reg.ReadInteger('Active page index')
        end;
        CheckBox2.Checked:=Reg.ReadBool('Remove External Links');
        CheckBox11.Checked:=Reg.ReadBool('Remove ALL images');
        CheckBox3.Checked:=Reg.ReadBool('Remove off-site images');
        CheckBox4.Checked:=Reg.ReadBool('Preserve dinamic images');
        CheckBox6.Checked:=Reg.ReadBool('No enmty lines');
        CheckBox7.Checked:=Reg.ReadBool('Leave quotes as is');
        CheckBox15.Checked:=Reg.ReadBool('Skip footnotes');
        CheckBox16.Checked:=Reg.ReadBool('Skip _italic_');
        CheckBox17.Checked:=Reg.ReadBool('Skip epigraphs');
        CheckBox18.Checked:=Reg.ReadBool('No paragraph restore');
        CheckBox19.Checked:=Reg.ReadBool('No description');
        CheckBox20.Checked:=Reg.ReadBool('Allow 500 errors');
        CheckBox21.Checked:=Reg.ReadBool('Skip poems');
        CheckBox12.Checked:=Reg.ReadBool('Preserv forms');
        CheckBox9.Checked:=Reg.ReadBool('No encoding detection');
        RadioGroup1.ItemIndex:=Reg.ReadInteger('Header detect method');
        CheckBox5.Checked:=Reg.ReadBool('NO headers detection');
        CheckBox14.Checked:=Reg.ReadBool('Ignore spaces');
        CheckBox8.Checked:=Reg.ReadBool('Leave dashes');
        Edit2.Text:=IntToStr(Reg.ReadInteger('Links download level'));
        CheckBox23.Checked:=Reg.ReadBool('Folow external links');
        CheckBox22.Checked:=Reg.ReadBool('Folow links');
        Edit3.Text:=Reg.ReadString('Headers detect regexp');
        Edit4.Text:=Reg.ReadString('Always follow');
        Edit5.Text:=Reg.ReadString('Never Follow');
        If not IsReload then
        Begin
          Width:=Reg.ReadInteger('WindowWidth');
          Height:=Reg.ReadInteger('WindowHeight');
        end;
        Reg.OpenKey('onload',False);
        LoadRegExps(StringGrid2);
        Reg.CloseKey;
        Reg.OpenKey('software\Grib Soft\Any to FB2\1.0'+Name+'\onfinish',False);
        LoadRegExps(StringGrid1);
      end;
    Finally
      Reg.Free;
    end;
  Except
  end;
end;

Procedure TUpenFileModalDialog.SaveSettings;
Var
  Reg:Tregistry;
  Procedure StoreRegExps(Grid:TStringGrid);
  Var
    I:Integer;
  Begin
    For I:=0 to Grid.RowCount-1 do
    Begin
      if LineIsEmpty(Grid,I) then  Continue;
      Reg.WriteString(IntToStr(I),Grid.Rows[I][0]);
      Reg.WriteString(IntToStr(I)+'_',Grid.Rows[I][1]);
    end;
  end;
Begin
  if Name<>'' then name:='\presets\'+name;
  Reg:=TRegistry.Create;
  Try
    Reg.OpenKey('software\Grib Soft\Any to FB2\1.0'+Name,True);
    Reg.WriteBool('Remove External Links', CheckBox2.Checked);
    Reg.WriteBool('Remove ALL images', CheckBox11.Checked);
    Reg.WriteBool('Remove off-site images', CheckBox3.Checked);
    Reg.WriteBool('Preserve dinamic images', CheckBox4.Checked);
    Reg.WriteBool('No enmty lines', CheckBox6.Checked);
    Reg.WriteBool('Leave quotes as is', CheckBox7.Checked);
    Reg.WriteBool('Skip footnotes', CheckBox15.Checked);
    Reg.WriteBool('Skip _italic_', CheckBox16.Checked);
    Reg.WriteBool('Skip epigraphs', CheckBox17.Checked);
    Reg.WriteBool('No paragraph restore', CheckBox18.Checked);
    Reg.WriteBool('No description', CheckBox19.Checked);
    Reg.WriteBool('Allow 500 errors', CheckBox20.Checked);
    Reg.WriteBool('Skip poems', CheckBox21.Checked);
    Reg.WriteBool('Preserv forms', CheckBox12.Checked);
    Reg.WriteBool('No encoding detection', CheckBox9.Checked);
    Reg.WriteInteger('Header detect method',RadioGroup1.ItemIndex);
    Reg.WriteBool('NO headers detection', CheckBox5.Checked);
    Reg.WriteBool('Ignore spaces', CheckBox14.Checked);
    Reg.WriteBool('Leave dashes', CheckBox8.Checked);
    Try
      Reg.WriteInteger('Links download level',StrToInt(Edit2.Text));
    except
      Reg.WriteInteger('Links download level',2);
    end;
    Reg.WriteBool('Folow external links',CheckBox23.Checked);
    Reg.WriteBool('Folow links',CheckBox22.Checked);
    Reg.WriteString('Headers detect regexp',Edit3.Text);
    Reg.WriteString('Always follow',Edit4.Text);
    Reg.WriteString('Never Follow',Edit5.Text);
    Reg.DeleteKey('onload');
    Reg.DeleteKey('onfinish');
    Reg.OpenKey('onload',True);
    StoreRegExps(StringGrid2);
    Reg.CloseKey;
    Reg.OpenKey('software\Grib Soft\Any to FB2\1.0'+Name+'\onfinish',True);
    StoreRegExps(StringGrid1);
  Finally
    Reg.Free;
  end;
end;

procedure TUpenFileModalDialog.FormClose(Sender: TObject;
  var Action: TCloseAction);
Var
  Reg:Tregistry;
Begin
  SaveSettings('');
  Reg:=TRegistry.Create;
  Try
    Reg.OpenKey('software\Grib Soft\Any to FB2\1.0',True);
//    Reg.WriteBool('Use defaults', CheckBox1.Checked);
    Reg.WriteBool('Close on finish', CheckBox24.Checked);
    Reg.WriteInteger('WindowWidth',Width);
    Reg.WriteInteger('WindowHeight',Height);
    Reg.WriteString('LastOpenURI',Edit1.Text);
    Reg.WriteInteger('Active page index',PageControl1.ActivePageIndex)
  Finally
    Reg.Free;
  end;
end;

procedure TUpenFileModalDialog.FormCreate(Sender: TObject);
begin
  LoadSettings(False,'');
  LoadSetList;
  CheckBox11Click(Self);
  CheckBox22Click(Self);
  CheckBox2Click(Self);
  ComboBox1.ItemIndex:=0;
  Panel5Resize(Self);
end;

procedure TUpenFileModalDialog.CheckBox11Click(Sender: TObject);
begin
  CheckBox3.Enabled:=not (ComboBox1.itemindex=1) and not CheckBox11.Checked;
  CheckBox4.Enabled:=not (ComboBox1.itemindex=1) and not CheckBox11.Checked;
end;

procedure TUpenFileModalDialog.CheckBox11KeyPress(Sender: TObject;
  var Key: Char);
begin
  CheckBox11Click(Self);
end;

procedure TUpenFileModalDialog.CheckBox22Click(Sender: TObject);
begin
  Edit2.Enabled:=CheckBox22.Checked and not (ComboBox1.itemindex=1) and CheckBox22.Enabled;
  Edit4.Enabled:=Edit2.Enabled;
  Edit5.Enabled:=Edit2.Enabled;
  Label5.Enabled:=Edit2.Enabled;
  Label6.Enabled:=Edit2.Enabled;
  CheckBox23.Enabled:=CheckBox22.Checked and not (ComboBox1.itemindex=1) and CheckBox22.Enabled;
  UpDown1.Enabled:=CheckBox22.Checked and not (ComboBox1.itemindex=1) and CheckBox22.Enabled;
  Label3.Enabled:=CheckBox22.Checked and not (ComboBox1.itemindex=1) and CheckBox22.Enabled;
end;

procedure TUpenFileModalDialog.CheckBox22KeyPress(Sender: TObject;
  var Key: Char);
begin
  CheckBox22Click(Self);
end;

procedure TUpenFileModalDialog.Panel1Resize(Sender: TObject);
begin
  Edit1.Width:=ClientWidth-43;
  ComboBox1.Width:=ClientWidth-73
end;
procedure TUpenFileModalDialog.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.WndParent := ParentHandle;
end;

procedure TUpenFileModalDialog.Panel6Resize(Sender: TObject);
begin
  Edit3.Width:=Panel6.Width;
end;

procedure TUpenFileModalDialog.Copy1Click(Sender: TObject);
begin
  RichEdit1.SelectAll;
  RichEdit1.CopyToClipboard;
end;

procedure TUpenFileModalDialog.CheckBox2Click(Sender: TObject);
begin
  CheckBox22.Enabled:=not CheckBox2.Checked and CheckBox2.Enabled;
  CheckBox22Click(Self);
end;

procedure TUpenFileModalDialog.CheckBox2KeyPress(Sender: TObject;
  var Key: Char);
begin
  CheckBox2Click(Self);
end;

Procedure TUpenFileModalDialog.LoadSetList;
Var
  Reg:TRegistry;
Begin
  Reg:=TRegistry.Create;
  Try
    Reg.OpenKeyReadOnly('Software\Grib Soft\Any to FB2\1.0\presets');
    Reg.GetKeyNames(ComboBox1.Items);
    Reg.CloseKey;
  FInally
    Reg.Free;
  end;
  ComboBox1.Items.Insert(0,'<Defaults>');
  ComboBox1.Items.Insert(0,'<Last used>');
end;

procedure TUpenFileModalDialog.StringGrid1GetEditText(Sender: TObject;
  ACol, ARow: Integer; var Value: String);
Var
  Grid:TStringGrid absolute Sender;
begin
  If (Value=SearchPatt) or
    (Value=ReplacePatt) then
  Value:='';
  Button6.Visible:=True;
  If Grid.Focused then
    Button6.Top:=Grid.Top+(Grid.Row-Grid.TopRow)*Grid.DefaultRowHeight+3;
end;

procedure TUpenFileModalDialog.StringGrid1SetEditText(Sender: TObject;
  ACol, ARow: Integer; const Value: String);
Var
  Grid:TStringGrid absolute Sender;
  I:Integer;
begin
  if LineIsEmpty(Grid,ARow) then
  Begin
    if ARow<>Grid.RowCount-1 then
    Begin
      For I:=ARow to Grid.RowCount-2 do
      Begin
        Grid.Cols[0][I]:=Grid.Cols[0][I+1];
        Grid.Cols[1][I]:=Grid.Cols[1][I+1];
      end;
      Grid.RowCount:=Grid.RowCount-1;
    end;
    Exit;
  end;
  if ARow=Grid.RowCount-1 then
  Begin
    Grid.RowCount:=Grid.RowCount+1;
    AddRow(Grid);
  end;
end;

procedure TUpenFileModalDialog.StringGrid2TopLeftChanged(Sender: TObject);
begin
  Button6.Visible:=False;
end;

procedure TUpenFileModalDialog.StringGrid2Enter(Sender: TObject);
Var
  Grid:TStringGrid absolute Sender;
begin
  Button6.Top:=Grid.Top+(Grid.Row-Grid.TopRow)*Grid.DefaultRowHeight+3;
end;

procedure TUpenFileModalDialog.Button6Click(Sender: TObject);
Var
  Grid:TStringGrid;
begin
  If Button6.Top>StringGrid1.Top then
    Grid:=StringGrid1
  else
    Grid:=StringGrid2;
  With TReEditorForm.Create(Self) do
  Try
    Edit1.Text:=Grid.Rows[Grid.Row][0];
    Memo2.Text:=Grid.Rows[Grid.Row][1];
    If ShowModal=mrOk then
    Begin
      Grid.Rows[Grid.Row][0]:=Edit1.Text;
      Grid.Rows[Grid.Row][1]:=Memo2.Text;
    end
  Finally
    Free;
  end;
end;

procedure TUpenFileModalDialog.RadioGroup1Click(Sender: TObject);
begin
  If RadioGroup1.ItemIndex=1 then
  Begin
    if CheckBox14.Checked then
    Begin
      TimerCounter:=5;
      Timer1.Enabled:=True;
      MessageBeep(MB_ICONEXCLAMATION);
    end;
  end;
end;

procedure TUpenFileModalDialog.Timer1Timer(Sender: TObject);
begin
  if  ((TimerCounter mod 2) = 0) then
  Begin
    CheckBox8.Font.Style :=[];
    CheckBox14.Font.Style :=[];
  end else
    Begin
      if CheckBox14.Checked then
        CheckBox14.Font.Style :=[fsUnderline,fsBold];
    end;
  If (TimerCounter=0) then
  Begin
    Timer1.Enabled:=False;
  end;
  dec(TimerCounter);
end;

end.
