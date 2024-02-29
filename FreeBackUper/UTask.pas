unit UTask;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, ExtCtrls, PDirSelected, DateUtils;

type
  TForm2 = class(TForm)
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    DirDialog1: TDirDialog;
    PC: TPageControl;
    TS1: TTabSheet;
    TS2: TTabSheet;
    TS3: TTabSheet;
    TS4: TTabSheet;
    TS5: TTabSheet;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    dtp2: TDateTimePicker;
    dtp1: TDateTimePicker;
    CheckBox10: TCheckBox;
    GroupBox1: TGroupBox;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    TS6: TTabSheet;
    mc1: TMonthCalendar;
    ListBox1: TListBox;
    dtp3: TDateTimePicker;
    dtp4: TDateTimePicker;
    GroupBox2: TGroupBox;
    rbOnce: TRadioButton;
    rbEveryDay: TRadioButton;
    rbWeekDays: TRadioButton;
    rbMonthDays: TRadioButton;
    BitBtn5: TBitBtn;
    BitBtn7: TBitBtn;
    BitBtn9: TBitBtn;
    BitBtn11: TBitBtn;
    Label1: TLabel;
    edtSource: TEdit;
    SpeedButton1: TSpeedButton;
    CheckBox9: TCheckBox;
    Label2: TLabel;
    edtDestination: TEdit;
    SpeedButton2: TSpeedButton;
    Label3: TLabel;
    edtExt: TEdit;
    SpeedButton3: TSpeedButton;
    ListBox2: TListBox;
    SpeedButton4: TSpeedButton;
    Label4: TLabel;
    edtTaskName: TEdit;
    BitBtn12: TBitBtn;
    CheckBox8: TCheckBox;
    dtp22: TDateTimePicker;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    Label5: TLabel;
    edtCopies: TEdit;
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
    procedure BitBtn11Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure BitBtn12Click(Sender: TObject);
    procedure rbOnceClick(Sender: TObject);
    procedure rbEveryDayClick(Sender: TObject);
    procedure rbWeekDaysClick(Sender: TObject);
    procedure rbMonthDaysClick(Sender: TObject);
    procedure mc1DblClick(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure edtCopiesKeyPress(Sender: TObject; var Key: Char);
  private
    procedure SortNumbersInListBox;
  public
  end;

var
  Form2: TForm2;

implementation

uses UMain;

{$R *.dfm}

procedure TForm2.BitBtn3Click(Sender: TObject);
begin
PC.ActivePage := TS1;
end;

//После выбора периодичности запуска
procedure TForm2.BitBtn4Click(Sender: TObject);
begin
  TS3.TabVisible := False;
  TS4.TabVisible := False;
  TS5.TabVisible := False;
  TS6.TabVisible := False;

  if rbOnce.Checked then
  begin
    TS3.TabVisible := True;
    PC.ActivePage := TS3;
  end;

  if rbEveryDay.Checked then
  begin
    TS4.TabVisible := True;
    PC.ActivePage := TS4;
  end;

  if rbWeekDays.Checked then
  begin
    TS5.TabVisible := True;
    PC.ActivePage := TS5;
  end;

  if rbMonthDays.Checked then
  begin
    TS6.TabVisible := True;
    PC.ActivePage := TS6;
  end;
end;

procedure TForm2.BitBtn5Click(Sender: TObject);
begin
PC.ActivePage := TS2;
end;

procedure TForm2.BitBtn7Click(Sender: TObject);
begin
PC.ActivePage := TS2;
end;

procedure TForm2.BitBtn9Click(Sender: TObject);
begin
PC.ActivePage := TS2;
end;

procedure TForm2.BitBtn11Click(Sender: TObject);
begin
PC.ActivePage := TS2;
end;

procedure TForm2.SpeedButton1Click(Sender: TObject);
begin
if DirDialog1.Execute then
   edtSource.Text := DirDialog1.DirPath;
end;

procedure TForm2.SpeedButton2Click(Sender: TObject);
begin
if DirDialog1.Execute then
   edtDestination.Text := DirDialog1.DirPath;
end;

procedure TForm2.SpeedButton3Click(Sender: TObject);
begin
if pos('*.', edtExt.Text) = 0  then
       ListBox2.Items.Add('*.' + edtExt.Text)
 else
       ListBox2.Items.Add(edtExt.Text);
edtExt.Clear;

end;

procedure TForm2.SpeedButton4Click(Sender: TObject);
begin
ListBox2.DeleteSelected;
end;

procedure TForm2.BitBtn12Click(Sender: TObject);
begin
TS2.TabVisible := true;
PC.ActivePage := TS2;
end;

procedure TForm2.rbOnceClick(Sender: TObject);
begin
Form1.Period := 'Однократно';
end;

procedure TForm2.rbEveryDayClick(Sender: TObject);
begin
Form1.Period := 'Ежедневно';
end;

procedure TForm2.rbWeekDaysClick(Sender: TObject);
begin
Form1.Period := 'По дням недели';
end;

procedure TForm2.rbMonthDaysClick(Sender: TObject);
begin
Form1.Period := 'По дням месяца';
end;


procedure TForm2.mc1DblClick(Sender: TObject);
var
  i: Integer;
begin
  // Проверяем, есть ли уже такая строка в ListBox
  for i := 0 to ListBox1.Items.Count - 1 do
  begin
    if ListBox1.Items[i] = IntToStr(DayOf(mc1.Date)) then
    begin
      ShowMessage('Число ' + IntToStr(DayOf(mc1.Date)) + ' уже существует в списке!');
      Exit; // Прерываем добавление строки, если она уже существует
    end;
  end;

  // Если строка уникальна, то добавляем её в ListBox
  ListBox1.Items.Add(IntToStr(DayOf(mc1.Date)));
  //Сортировка ListBox
  SortNumbersInListBox;
end;

procedure TForm2.SortNumbersInListBox;
var
  Numbers: array of Integer;
  i, j, temp: Integer;
begin
  SetLength(Numbers, ListBox1.Items.Count);

  // Копируем числа из ListBox в массив и преобразуем их в числовой формат
  for i := 0 to ListBox1.Items.Count - 1 do
  begin
    Numbers[i] := StrToInt(ListBox1.Items[i]);
  end;

  // Применяем пузырьковую сортировку к массиву чисел
  for i := 0 to Length(Numbers) - 2 do
  begin
    for j := i + 1 to Length(Numbers) - 1 do
    begin
      if Numbers[i] > Numbers[j] then
      begin
        temp := Numbers[i];
        Numbers[i] := Numbers[j];
        Numbers[j] := temp;
      end;
    end;
  end;

  // Очищаем ListBox
  ListBox1.Clear;

  // Заполняем ListBox отсортированными числами
  for i := 0 to Length(Numbers) - 1 do
  begin
    ListBox1.Items.Add(IntToStr(Numbers[i]));
  end;
end;

procedure TForm2.SpeedButton5Click(Sender: TObject);
begin
ListBox2.Clear;
end;

procedure TForm2.SpeedButton7Click(Sender: TObject);
begin
ListBox1.Clear;
end;

procedure TForm2.SpeedButton6Click(Sender: TObject);
begin
ListBox1.DeleteSelected;
end;

procedure TForm2.BitBtn1Click(Sender: TObject);
begin
  if Form1.SaveEnabled then ModalResult := mrOK else
  begin
    Application.MessageBox('Недостаточно данных для запуска задания.' + #13#10 + 'Заполните поля ввода.',
    'Ошибка сохранения', MB_ICONERROR +  MB_OK);
    ModalResult := mrCancel;
  end;
end;

procedure TForm2.edtCopiesKeyPress(Sender: TObject; var Key: Char);
begin
  If not (Key in ['0'..'9', #8, #13]) then
     Key := #0;
end;

end.
