unit UMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, CoolTrayIcon, Menus, ComCtrls, ImgList, ActnList,
  XPStyleActnCtrls, ActnMan, ToolWin, ActnCtrls, DB, DBClient,
  DBGridEhGrouping, GridsEh, DBGridEh, ExtCtrls, StdCtrls, DateUtils, Math,
  VCLUnZip, VCLZip, Contnrs, DBCtrls, Mask;

  type  //Поток
  TMyThread = class(TThread)
  private
    TaskID: String; //У потока есть поле = GUID (CDS.ID) Задания
    PercentDone: Integer;
    VCLZip : TVCLZip;
  protected
    procedure Execute; override;
    procedure VCLZipTotalPercentDone(Sender: TObject; Percent: Integer);
    procedure VCLZip1ZipComplete(Sender: TObject; FileCount: Integer);
  public
  end;


type
  TOperation = (NewRecord, EditRecord, DeleteRecord);


type
  TForm1 = class(TForm)
    CoolTrayIcon1: TCoolTrayIcon;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    ActionToolBar1: TActionToolBar;
    ActionManager1: TActionManager;
    New: TAction;
    Edit: TAction;
    Del: TAction;
    ImageList1: TImageList;
    Start: TAction;
    DataSource1: TDataSource;
    CDS: TClientDataSet;
    Panel1: TPanel;
    Panel2: TPanel;
    DBGridEh1: TDBGridEh;
    CDSID: TStringField;
    CDSTASK_NAME: TStringField;
    CDSACTIVE: TBooleanField;
    CDSPERIOD: TStringField;
    CDSWEEK_DAY: TStringField;
    CDSMONTH_DAY: TStringField;
    CDSDATA: TStringField;
    CDSTIME: TStringField;
    CDSRUN_TIME: TStringField;
    CDSSOURCE_PATH: TStringField;
    CDSSUB_FOLDERS: TBooleanField;
    CDSDESTINATION_PATH: TStringField;
    CDSEXTENSIONS: TStringField;
    CDSKEEP_COPIES: TIntegerField;
    Memo1: TMemo;
    TimerTask: TTimer;
    Panel3: TPanel;
    Panel4: TPanel;
    lbWorking: TLabel;
    lbPercent: TLabel;
    ProgressBar1: TProgressBar;
    Label1: TLabel;
    DBEdit1: TDBEdit;
    DBCheckBox1: TDBCheckBox;
    Label2: TLabel;
    DBEdit2: TDBEdit;
    Label3: TLabel;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    Label4: TLabel;
    Label5: TLabel;
    Edit1: TEdit;
    About: TAction;
    Clear: TAction;
    procedure CoolTrayIcon1Click(Sender: TObject);
    procedure CoolTrayIcon1Startup(Sender: TObject;
      var ShowMainForm: Boolean);
    procedure N2Click(Sender: TObject);
    procedure NewExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure EditExecute(Sender: TObject);
    procedure DelExecute(Sender: TObject);
    procedure ActionManager1Update(Action: TBasicAction;
      var Handled: Boolean);
    procedure CDSBeforeDelete(DataSet: TDataSet);
    procedure DBGridEh1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumnEh; State: TGridDrawState);
    procedure DataSource1DataChange(Sender: TObject; Field: TField);
    procedure StartExecute(Sender: TObject);
    procedure TimerTaskTimer(Sender: TObject);
    procedure AboutExecute(Sender: TObject);
    procedure ClearExecute(Sender: TObject);
  private
    FOperation: TOperation;
    FPeriod: String;
    procedure CreateCDS;
    procedure EditCDS;
    procedure ClearParams;
    procedure LoadParams;
    function CheckWeekDays : String;
    procedure SetRunTime;
    procedure ScanDirExt(StartDir: string; Mask: string; List: TStrings);
    procedure DeleteFiles;
    procedure DoZip;
    procedure TaskExecute;
    function ExistingThread(id: String): integer;
    function ConvertDaysOfWeek(inputStr: string): string;
    function GetNextScheduledTimebyTime(const StartTime: string; IntervalDays: Integer): TDateTime;
    function GetNextScheduledTimebyWeek(currentTime: TDateTime; scheduledDays: array of Integer; scheduledTime: TDateTime): TDateTime;
    function GetNextScheduledTimeByDates(currentTime: TDateTime; const scheduledDates: array of Integer; scheduledTime: TDateTime): TDateTime;
  public
    scheduledDays: array of Integer;
    scheduledDates: array of Integer;
    property Operation: TOperation read FOperation write FOperation;
    property Period: String read FPeriod write FPeriod;
    function SaveEnabled : Boolean;
  end;

const
  clPaleRed = TColor($CCCCFF);
  clPaleGreen = TColor($CCFFCC);

var
  Form1: TForm1;
  Thread1, CurrentThread: TMyThread;//Процесс, выбранный процесс
  ThreadList: TObjectList;//Для хранения активных заданий
  WorkingThreads: integer;//Количество запущенных процессов
  Recalc: Boolean;//Возможность пересчета времени запуска
  CanCloseForm: Boolean = False;
implementation

uses UTask, UAbout;

{$R *.dfm}

//Сравнение строк в StringList
function MyCompareFunction(List: TStringList; Index1, Index2: Integer): Integer;
begin
 Result := CompareValue(Integer(List.Objects[Index1]), Integer(List.Objects[Index2]))
end;

procedure TMyThread.Execute;
begin
  VCLZip.Zip;
end;

//Процент архивации
procedure TMyThread.VCLZipTotalPercentDone(Sender: TObject; Percent: Integer);
begin
 {if Self = CurrentThread проверяет,
  является ли текущий объект равным объекту,
  который представляет текущий поток выполнения (CurrentThread).}
   if Self = CurrentThread then
      begin
        PercentDone := Percent;
        Form1.ProgressBar1.Position := PercentDone;
        Form1.lbPercent.Caption := IntToStr(PercentDone) + '%';
      end
end;

//По завершению архивации Поток уничтожаем
procedure TMyThread.VCLZip1ZipComplete(Sender: TObject; FileCount: Integer);
var S: String;
begin
      Terminate;
      dec(WorkingThreads);
      S := ExtractFileName(VCLZip.ZipName);
      Delete(S, 1, 20);
      Delete(S, Length(S)-3, 4);
      Form1.Memo1.Lines.Insert(0, '[' + FormatDateTime('dd.mm.yyyy hh:mm:ss', Now) + ']' + ' - Окончание задания: ' + S);
      Form1.Memo1.Lines.Insert(0, '[' + FormatDateTime('dd.mm.yyyy hh:mm:ss', Now) + ']' + ' - Создан файл: ' + VCLZip.ZipName);
      Form1.ProgressBar1.Position := 0;
      Form1.lbPercent.Caption := '0%';
end;

procedure TForm1.CoolTrayIcon1Click(Sender: TObject);
begin
Application.MainForm.Show;
ShowWindow(Application.Handle, SW_RESTORE); //Восстановить
Application.BringToFront;
end;

procedure TForm1.CoolTrayIcon1Startup(Sender: TObject;
  var ShowMainForm: Boolean);
begin
ShowMainForm := False;
end;

procedure TForm1.N2Click(Sender: TObject);
begin
CanCloseForm := True;
Close;

end;

procedure TForm1.NewExecute(Sender: TObject);
begin
 Operation := NewRecord;
 Form2.Caption := 'Новое задание';
 ClearParams;
 EditCDS;
end;

procedure TForm1.EditExecute(Sender: TObject);
begin
 Operation := EditRecord;
 Form2.Caption := 'Изменить задание';
 ClearParams;
 LoadParams;
 EditCDS;
end;

//Добавление нового задания или редактирование существующего
procedure TForm1.EditCDS;
var MyGUID : TGUID;
S : String;
begin
 if Form2.ShowModal = mrOk then
 begin
   if Operation = NewRecord then CDS.Insert;
   if (Operation = EditRecord) and (CDS.RecordCount > 0) then CDS.Edit;

   if CDS.State = dsInsert then
   begin
      CreateGUID(MyGUID);
      S := GUIDToString(MyGUID);
      Delete(S, 1, 1);
      Delete(S, Length(S), 1);
      CDS.FieldByName('ID').AsString := S;
   end;
   CDS.FieldByName('TASK_NAME').AsString := Form2.EdtTaskName.Text;
   CDS.FieldByName('ACTIVE').AsBoolean := Form2.CheckBox8.Checked;
   CDS.FieldByName('PERIOD').AsString := Period;
   CDS.FieldByName('SOURCE_PATH').AsString := Form2.EdtSource.Text;
   CDS.FieldByName('SUB_FOLDERS').AsBoolean := Form2.CheckBox9.Checked;
   CDS.FieldByName('DESTINATION_PATH').AsString := Form2.EdtDestination.Text;
   CDS.FieldByName('EXTENSIONS').AsString := Form2.ListBox2.Items.CommaText;
   CDS.FieldByName('KEEP_COPIES').AsInteger := StrToInt(Form2.EdtCopies.Text);

   if Period = 'Однократно' then
   begin
      CDS.FieldByName('DATA').AsString := DateToStr(RecodeSecond(Form2.dtp2.Date, 0));
      CDS.FieldByName('TIME').AsString := TimeToStr(RecodeSecond(Form2.dtp22.Time, 0));
   end;

   if Period = 'Ежедневно' then
   begin
      CDS.FieldByName('TIME').AsString := TimeToStr(RecodeSecond(Form2.dtp1.Time, 0));
   end;

   if Period = 'По дням недели' then
   begin
      CDS.FieldByName('WEEK_DAY').AsString := CheckWeekDays;
      CDS.FieldByName('TIME').AsString := TimeToStr(RecodeSecond(Form2.dtp4.Time, 0));
   end;

   if Period = 'По дням месяца' then
   begin
      CDS.FieldByName('MONTH_DAY').AsString := Form2.ListBox1.Items.CommaText;
      CDS.FieldByName('TIME').AsString := TimeToStr(RecodeSecond(Form2.dtp3.Time, 0));
   end;   
   CDS.Post;
     SetRunTime; //После редактирования пересчитать следующее время запуска
 end;
end;

//Создание таблицы в виде xml файла
procedure TForm1.CreateCDS;
begin
   with CDS do begin
     Active := False;
      with FieldDefs do begin
        Clear;
        //Поля таблицы
        Add('ID', ftString, 36, False);
        Add('TASK_NAME', ftString, 100, False);
        Add('ACTIVE', ftBoolean, 0, False);
        Add('PERIOD', ftString, 15, False);
        Add('WEEK_DAY', ftString, 15, False);
        Add('MONTH_DAY', ftString, 100, False);
        Add('DATA', ftString, 10, False);
        Add('TIME', ftString, 8, False);
        Add('RUN_TIME', ftString, 20, False);
        Add('SOURCE_PATH', ftString, 400, False);
        Add('SUB_FOLDERS', ftBoolean, 0, False);
        Add('DESTINATION_PATH', ftString, 400, False);
        Add('EXTENSIONS', ftString, 400, False);
        Add('KEEP_COPIES', ftInteger, 0, False);
      end;
   CreateDataSet;
   SaveToFile(ExtractFileDir(Application.ExeName) + '\Tasks.xml',dfXML);
   end;

end;

//Формирование строки в таблице по выбранным дням недели
function TForm1.CheckWeekDays : String;
var
  CheckBoxes: array[0..6] of TCheckBox; // Массив из 7 CheckBox'ов
  i, j : integer;
begin
  Result := '';
  j := 0;
  //Сколько всего выделено дней недели
  for i := 0 to Form2.ComponentCount-1 do
    if (Form2.Components[i] is TCheckBox) and (Form2.Components[i].Tag = 1)then
         if TCheckBox(Form2.Components[i]).Checked  then
              Inc(j);

if j > 0 then
begin
  SetLength(scheduledDays, j);

  // Добавляем ссылки на чекбоксы в массив CheckBoxes
  CheckBoxes[0] := Form2.CheckBox1;
  CheckBoxes[1] := Form2.CheckBox2;
  CheckBoxes[2] := Form2.CheckBox3;
  CheckBoxes[3] := Form2.CheckBox4;
  CheckBoxes[4] := Form2.CheckBox5;
  CheckBoxes[5] := Form2.CheckBox6;
  CheckBoxes[6] := Form2.CheckBox7;

  j := 0;
  // Проходим по всем CheckBox'ам
  for i := 0 to 6 do
  begin
    // Проверяем, отмечен ли текущий CheckBox
    if CheckBoxes[i].Checked then
    begin
      // Если отмечен, добавляем соответствующий день недели в массив SelectedDays
      scheduledDays[j] := i+1;
      Result := Result + IntToStr(scheduledDays[j]) + ',';
      Inc(j);
    end
  end;
  Delete(Result, Length(Result), 1);
end;
end;

//Очистка формы
procedure TForm1.ClearParams;
begin
  Form2.edtTaskName.Text := '';
  Form2.edtSource.Text := '';
  Form2.edtDestination.Text := '';
  Form2.edtExt.Text := '';
  Form2.edtCopies.Text := '3';

  Form2.ListBox1.Clear;
  Form2.ListBox2.Clear;

  Form2.CheckBox1.Checked := False; //ПН
  Form2.CheckBox2.Checked := False;
  Form2.CheckBox3.Checked := False;
  Form2.CheckBox4.Checked := False;
  Form2.CheckBox5.Checked := False;
  Form2.CheckBox6.Checked := False;
  Form2.CheckBox7.Checked := False; //ВС
  Form2.CheckBox9.Checked := False; //Включать подкаталоги
  Form2.CheckBox8.Checked := True; //Активное задание

  Form2.rbOnce.Checked := False;
  Form2.rbEveryDay.Checked := False;
  Form2.rbWeekDays.Checked := False;
  Form2.rbMonthDays.Checked := False;

  Form2.TS2.TabVisible := False;
  Form2.TS3.TabVisible := False;
  Form2.TS4.TabVisible := False;
  Form2.TS5.TabVisible := False;
  Form2.TS6.TabVisible := False;

  Form2.dtp1.DateTime := Now;
  Form2.dtp2.DateTime := Now;
  Form2.dtp22.DateTime := Now;
  Form2.dtp3.Time := Now;
  Form2.dtp4.Time := Now;

  Form2.mc1.Date := Now;
  Form2.PC.ActivePage := Form2.TS1;
end;

//Загрузка из таблицы в форму
procedure TForm1.LoadParams;
var
  Days: TStringList;
  i: Integer;
begin
  Period := CDS.FieldByName('PERIOD').AsString;
  Form2.edtTaskName.Text := CDS.FieldByName('TASK_NAME').AsString;
  Form2.CheckBox8.Checked := CDS.FieldByName('ACTIVE').AsBoolean;
  Form2.edtSource.Text := CDS.FieldByName('SOURCE_PATH').AsString;
  Form2.CheckBox9.Checked := CDS.FieldByName('SUB_FOLDERS').AsBoolean;
  Form2.edtDestination.Text := CDS.FieldByName('DESTINATION_PATH').AsString;
  Form2.ListBox2.Items.CommaText := CDS.FieldByName('EXTENSIONS').AsString;
  Form2.edtCopies.Text := CDS.FieldByName('KEEP_COPIES').AsString;

  Form2.TS2.TabVisible := True;
     if Period = 'Однократно' then
   begin
      Form2.dtp2.Date := StrToDate(CDS.FieldByName('DATA').AsString);
      Form2.dtp22.Time := StrToTime(CDS.FieldByName('TIME').AsString);
      Form2.rbOnce.Checked := True;
      Form2.TS3.TabVisible := True;
   end;

   if Period = 'Ежедневно' then
   begin
      Form2.dtp1.Time := StrToTime(CDS.FieldByName('TIME').AsString);
      Form2.rbEveryDay.Checked := True;
      Form2.TS4.TabVisible := True;
   end;

   if Period = 'По дням недели' then
   begin
      //Отмечаем выбранные дни недели
      Days := TStringList.Create;
      Days.CommaText := CDS.FieldByName('WEEK_DAY').AsString;
        for i := 0 to Days.Count - 1 do
        begin
          if StrToInt(Days[i]) = 1 then Form2.CheckBox1.Checked := True;
          if StrToInt(Days[i]) = 2 then Form2.CheckBox2.Checked := True;
          if StrToInt(Days[i]) = 3 then Form2.CheckBox3.Checked := True;
          if StrToInt(Days[i]) = 4 then Form2.CheckBox4.Checked := True;
          if StrToInt(Days[i]) = 5 then Form2.CheckBox5.Checked := True;
          if StrToInt(Days[i]) = 6 then Form2.CheckBox6.Checked := True;
          if StrToInt(Days[i]) = 7 then Form2.CheckBox7.Checked := True;
        end;

      Days.Free;
      Form2.dtp4.Time := StrToTime(CDS.FieldByName('TIME').AsString);
      Form2.rbWeekDays.Checked := True;
      Form2.TS5.TabVisible := True;
   end;

   if Period = 'По дням месяца' then
   begin
      Form2.ListBox1.Items.CommaText := CDS.FieldByName('MONTH_DAY').AsString;
      Form2.dtp3.Time := StrToTime(CDS.FieldByName('TIME').AsString);
      Form2.rbMonthDays.Checked := True;
      Form2.TS6.TabVisible := True;
   end;   
end;

//Установка запуска каждого задания
procedure TForm1.SetRunTime;
var
  bm:TBookmark;
  Week, Month: TStringList;
  i : integer;
begin
if CDS.RecordCount > 0 then
begin
  with CDS do
  begin
  bm := GetBookMark;
  DisableControls;
  try
    First;
    while not Eof do
    begin
     //--------------------------------------
        CDS.Edit;
        if CDS.FieldByName('PERIOD').AsString = 'Однократно' then
        begin
          CDS.FieldByName('WEEK_DAY').AsString := '';
          CDS.FieldByName('MONTH_DAY').AsString := '';
          CDS.FieldByName('RUN_TIME').AsString := CDS.FieldByName('DATA').AsString + ' ' + CDS.FieldByName('TIME').AsString;
        end;

        if CDS.FieldByName('PERIOD').AsString = 'Ежедневно' then
        begin
          CDS.FieldByName('WEEK_DAY').AsString := '';
          CDS.FieldByName('MONTH_DAY').AsString := '';
          CDS.FieldByName('DATA').AsString := '';
          CDS.FieldByName('RUN_TIME').AsString := DateTimeToStr(GetNextScheduledTimebyTime(CDS.FieldByName('TIME').AsString, 1));
        end;

        if CDS.FieldByName('PERIOD').AsString = 'По дням недели' then
        begin
          Week := TStringList.Create;
          Week.CommaText := CDS.FieldByName('WEEK_DAY').AsString;
          SetLength(scheduledDays, Week.Count);
          //Заносим дни недели в массив
          for i := 0 to Week.Count - 1 do  scheduledDays[i] := StrToInt(Week[i]);
          CDS.FieldByName('MONTH_DAY').AsString := '';
          CDS.FieldByName('DATA').AsString := '';
          CDS.FieldByName('RUN_TIME').AsString :=
          DateTimeToStr(GetNextScheduledTimebyWeek(Now, scheduledDays, StrToDateTime(CDS.FieldByName('TIME').AsString)));
          Week.Free;
        end;

        if CDS.FieldByName('PERIOD').AsString = 'По дням месяца' then
        begin
          Month := TStringList.Create;
          Month.CommaText := CDS.FieldByName('MONTH_DAY').AsString;
          SetLength(scheduledDates, Month.Count);
          //Заносим дни месяца в массив
          for i := 0 to Month.Count - 1 do  scheduledDates[i] := StrToInt(Month[i]);
          CDS.FieldByName('WEEK_DAY').AsString := '';
          CDS.FieldByName('DATA').AsString := '';
          CDS.FieldByName('RUN_TIME').AsString :=
          DateTimeToStr(GetNextScheduledTimeByDates(Now, scheduledDates, StrToDateTime(CDS.FieldByName('TIME').AsString)));
          Month.Free;
        end;
        CDS.Post;
     //--------------------------------------
     Next;
    end;
  finally
    if BookmarkValid(bm) then
       begin
         GotoBookmark (bm);
         FreeBookmark(bm);
       end;
    EnableControls;
  end;
  end;
end;
end;

//Следующий запуск для Ежедневно
function TForm1.GetNextScheduledTimebyTime(const StartTime: string; IntervalDays: Integer): TDateTime;
var
  CurrentTime: TDateTime;
  StartHour, StartMinute, StartSecond, StartMSec: Word;
  NextRun: TDateTime;
begin
  CurrentTime := Now;
  DecodeTime(StrToTime(StartTime), StartHour, StartMinute, StartSecond, StartMSec);
  NextRun := EncodeDateTime(YearOf(CurrentTime), MonthOf(CurrentTime), DayOf(CurrentTime), StartHour, StartMinute, StartSecond, StartMSec);
  if CurrentTime > NextRun then
    NextRun := IncDay(NextRun, IntervalDays);
  Result := NextRun;
end;


procedure TForm1.FormCreate(Sender: TObject);
begin
 WorkingThreads := 0;
 Recalc := False;
 if FileExists(ExtractFileDir(Application.ExeName) + '\Tasks.xml') then
 CDS.LoadFromFile(ExtractFileDir(Application.ExeName) + '\Tasks.xml')
 else
 begin
   CreateCDS;
    CDS.LoadFromFile(ExtractFileDir(Application.ExeName) + '\Tasks.xml');
 end;

 if FileExists(ExtractFileDir(Application.ExeName) + '\log.txt') then
      Memo1.Lines.LoadFromFile(ExtractFileDir(Application.ExeName) + '\log.txt');
 SetRunTime;
end;


procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if CanCloseForm = False then
  begin
    Action := caNone;
    Application.MainForm.Hide;
    ShowWindow(Application.Handle, SW_HIDE);
  end
  else
  begin
    CDS.SaveToFile(ExtractFileDir(Application.ExeName) + '\Tasks.xml',dfXML);
    CDS.Close;
    if (WorkingThreads = 0) and (ThreadList <> nil) then FreeAndNil(ThreadList);
        Memo1.Lines.SaveToFile(ExtractFileDir(Application.ExeName) + '\log.txt');
  end;
end;


//Следующий запуск для дней недели
function TForm1.GetNextScheduledTimebyWeek(currentTime: TDateTime; scheduledDays: array of Integer; scheduledTime: TDateTime): TDateTime;
var
  currentDay: Integer;
  i: Integer;
  daysToNextScheduledDay: Integer;
  timeDifference: TDateTime;
begin
  currentDay := DayOfTheWeek(currentTime);

  // Проверяем, если текущий день недели в списке запланированных дней
  for i := 0 to High(scheduledDays) do
  begin
    if scheduledDays[i] > currentDay then
    begin
      daysToNextScheduledDay := scheduledDays[i] - currentDay;
      Result := currentTime + daysToNextScheduledDay + scheduledTime - Frac(currentTime);
      Exit;
    end
    else if scheduledDays[i] = currentDay then
    begin
      timeDifference := scheduledTime - Frac(currentTime);
      if timeDifference > 0 then
      begin
        Result := currentTime + timeDifference;
        Exit;
      end;
    end;
  end;
  
  // Если не найдено ближайшее запланированное время, переходим к следующей неделе
  daysToNextScheduledDay := 7 - currentDay + scheduledDays[0];
  Result := currentTime + daysToNextScheduledDay + scheduledTime - Frac(currentTime);
end;

//Следующий запуск для дней месяца
function TForm1.GetNextScheduledTimeByDates(currentTime: TDateTime; const scheduledDates: array of Integer; scheduledTime: TDateTime): TDateTime;
var
  currentYear, currentMonth, currentDay: Word;
  i: Integer;
  nextScheduledTime: TDateTime;
begin
  DecodeDate(currentTime, currentYear, currentMonth, currentDay);

  // Ищем ближайшую дату из выбранных дат месяца
  nextScheduledTime := 0;
  for i := Low(scheduledDates) to High(scheduledDates) do
  begin
    if (scheduledDates[i] > currentDay) or
       ((scheduledDates[i] = currentDay) and (Frac(scheduledTime) > Frac(currentTime))) then
    begin
      nextScheduledTime := EncodeDate(currentYear, currentMonth, scheduledDates[i]) + Frac(scheduledTime);
      Break;
    end;
  end;

  // Если не найдено в текущем месяце, ищем в следующем месяце
  if nextScheduledTime <= currentTime then
  begin
    if currentMonth = 12 then
    begin
      currentYear := currentYear + 1;
      currentMonth := 1;
    end
    else
    begin
      currentMonth := currentMonth + 1;
    end;

    nextScheduledTime := EncodeDate(currentYear, currentMonth, scheduledDates[0]) + Frac(scheduledTime);
  end;

  Result := nextScheduledTime;
end;

//Удалить задание
procedure TForm1.DelExecute(Sender: TObject);
begin
CDS.Delete;
end;

//Доступность кнопок вверху
procedure TForm1.ActionManager1Update(Action: TBasicAction;
  var Handled: Boolean);
begin
  lbWorking.Caption := 'Активных процессов: ' + IntToStr(WorkingThreads);
  if (WorkingThreads = 0) and (ThreadList <> nil) then FreeAndNil(ThreadList);
  
  if CDS.RecordCount = 0 then
  begin
   Edit.Enabled := False;
   Del.Enabled := False;
   Start.Enabled := False;
  end
  else
  begin
   Edit.Enabled := True;
   Del.Enabled := True;
   Start.Enabled := True;
  end;
end;


procedure TForm1.CDSBeforeDelete(DataSet: TDataSet);
begin
if Application.MessageBox('Удалить задание?','Внимание', MB_ICONWARNING +
    MB_OKCANCEL + MB_DEFBUTTON2) = IDCancel  then
    Abort;
end;

//Проверка возможности сохранить задание
function TForm1.SaveEnabled : Boolean;
begin
  Result := False;
  if (trim(Form2.edtTaskName.Text) <> '') and
     (trim(Form2.edtSource.Text) <> '') and
     (trim(Form2.edtDestination.Text) <> '') and
     (trim(Form2.edtCopies.Text) <> '') and
     (Form2.ListBox2.Items.Count > 0) then
        Result := True;
  if (Period = 'По дням месяца') and (Form2.ListBox1.Items.Count = 0) then
        Result := False;

  if (Period = 'По дням недели') and
        (
        not (
        (Form2.CheckBox1.Checked = True) or
        (Form2.CheckBox2.Checked = True) or
        (Form2.CheckBox3.Checked = True) or
        (Form2.CheckBox4.Checked = True) or
        (Form2.CheckBox5.Checked = True) or
        (Form2.CheckBox6.Checked = True) or
        (Form2.CheckBox7.Checked = True)
             )
        ) then
        Result := False;
end;

//Поиск файлов по маске (На выходе: Имя+Время создания от старых к новым)
procedure TForm1.ScanDirExt(StartDir: string; Mask: string; List: TStrings);
  var
  SearchRec: TSearchRec;
  begin
  if Mask = '' then
    Mask := '*.*';
  if StartDir[Length(StartDir)] <> '\' then
    StartDir := StartDir + '\';
  if FindFirst(StartDir + Mask, faAnyFile, SearchRec) = 0 then
  begin
    repeat Application.ProcessMessages;
      if (SearchRec.Attr and faDirectory) <> faDirectory then
        List.AddObject(StartDir + SearchRec.Name, TObject(SearchRec.Time))
      else if (SearchRec.Name <> '..') and (SearchRec.Name <> '.')then begin
        List.AddObject(StartDir + SearchRec.Name + '\', TObject(SearchRec.Time));
      ScanDirExt(StartDir + SearchRec.Name + '\', Mask, List);
  end;
  until FindNext(SearchRec) <> 0;
  FindClose(SearchRec);
  end;
  end;

//Архивация по заданию
procedure TForm1.DoZip;
var
  fn : String;
  Src, Dst: Boolean;
begin
 Src := DirectoryExists(Form1.CDS.FieldByName('SOURCE_PATH').AsString);
 Dst := DirectoryExists(Form1.CDS.FieldByName('DESTINATION_PATH').AsString);

 if not Src then
    Memo1.Lines.Insert(0, '[' + FormatDateTime('dd.mm.yyyy hh:mm:ss', Now) + ']' + ' - Исходный каталог отсутствует: ' + Form1.CDS.FieldByName('SOURCE_PATH').AsString);

 if not Dst then
    Memo1.Lines.Insert(0, '[' + FormatDateTime('dd.mm.yyyy hh:mm:ss', Now) + ']' + ' - Каталог архива отсутствует: ' + Form1.CDS.FieldByName('DESTINATION_PATH').AsString);

 if Src and Dst then
      BEGIN
       if not Assigned(ThreadList) then
          ThreadList := TObjectList.Create(True); // Создаем TObjectList с автовысвобождением объектов
       //Создание, запуск потока, добавление его в ObjectList
       Thread1:=TMyThread.Create(True);
       inc(WorkingThreads);
       Thread1.Priority:=tpNormal;

       //Имя файла, формируется как дата+имя задания
       fn := FormatDateTime('yyy-mm-dd-hh-mm-ss', now) + '_' + Form1.CDS.FieldByName('TASK_NAME').AsString + '.zip';
       Thread1.TaskID := Form1.CDS.FieldByName('ID').AsString;
       Thread1.VCLZip := TVCLZip.Create(Self);
       Thread1.VCLZip.Recurse := Form1.CDS.FieldByName('SUB_FOLDERS').AsBoolean;
       Thread1.VCLZip.ZipName := IncludeTrailingPathDelimiter(Form1.CDS.FieldByName('DESTINATION_PATH').AsString) + fn;
       Thread1.VCLZip.RootDir := Form1.CDS.FieldByName('SOURCE_PATH').AsString;
       Thread1.VCLZip.FilesList.CommaText := Form1.CDS.FieldByName('EXTENSIONS').AsString;
       Thread1.VCLZip.ZipAction := zaReplace;
       Thread1.VCLZip.StorePaths := True;  //В архив - с каталогами
       Thread1.VCLZip.OnTotalPercentDone := Thread1.VCLZipTotalPercentDone;
       Thread1.VCLZip.OnZipComplete :=  Thread1.VCLZip1ZipComplete;
       Thread1.Resume;  //Запуск потока
       ThreadList.Add(Thread1); //Добавление его в ObjectList
       CurrentThread := Thread1;
       Memo1.Lines.Insert(0, '[' + FormatDateTime('dd.mm.yyyy hh:mm:ss', Now) + ']' + ' - Старт задания:' + ' ' + Form1.CDS.FieldByName('TASK_NAME').AsString);
      END;
end;

//Удаляем фалы по маске *Имя задания.zip. Оставляем сколько указано самых новых-1.
procedure TForm1.DeleteFiles;
var L :TStringList;
    i, n : integer;
begin
 n := CDS.FieldByName('KEEP_COPIES').AsInteger;
 L := TStringList.Create;
 ScanDirExt(CDS.FieldByName('DESTINATION_PATH').AsString, '*' + CDS.FieldByName('TASK_NAME').AsString + '.zip', L);
 L.CustomSort(MyCompareFunction);
 // Проверяем, что в списке есть больше= n строк (файлов)
  if L.Count >= n then
  begin
    // Удаляем верхние строки (файлы), оставляя только последние n
    for i := 0 to L.Count - n  do
    begin
      DeleteFile(L[i]);
    end;
  end;
end;

//Раскраска грида
procedure TForm1.DBGridEh1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumnEh;
  State: TGridDrawState);
begin
  if CDS.FieldByName('ACTIVE').AsBoolean = True
  then
  begin
    if (gdFocused in State) then //имеет ли ячейка фокус?
      DBGridEh1.Canvas.Brush.Color := clHighlight//имеет фокус
    else
      DBGridEh1.Canvas.Brush.Color := clPaleGreen;//не имеет фокуса
  end
  else
  begin
    if (gdFocused in State) then //имеет ли ячейка фокус?
      DBGridEh1.Canvas.Brush.Color := clHighlight//имеет фокус
    else
      DBGridEh1.Canvas.Brush.Color := clPaleRed;//не имеет фокуса
  end;

  //Теперь закрасим ячейку используя стандартный метод:
  DBGridEh1.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

//При перемещении по записям Поиск потока по его TaskID
//и если нашли, он будет CurrentThread
procedure TForm1.DataSource1DataChange(Sender: TObject; Field: TField);
begin
  if CDS.FieldByName('ACTIVE').AsBoolean = True then
   DBEdit4.Color := clPaleGreen
  else
   DBEdit4.Color := clPaleRed;

  if CDS.FieldByName('PERIOD').AsString = 'По дням месяца' then
     begin
       Label5.Caption := 'Дни месяца';
       Edit1.Text := CDS.FieldByName('MONTH_DAY').AsString;
       Edit1.Visible := True;
     end
  else
  if CDS.FieldByName('PERIOD').AsString = 'По дням недели' then
     begin
       Label5.Caption := 'Дни недели';
       Edit1.Text := ConvertDaysOfWeek(CDS.FieldByName('WEEK_DAY').AsString);
       Edit1.Visible := True;
     end
  else
     begin
       Label5.Caption := '';
       Edit1.Visible := False;
     end;
  //Текущй поток для ProgressBar
  if ExistingThread(CDS.FieldByName('ID').AsString) > -1 then
       CurrentThread := TMyThread(ThreadList[ExistingThread(CDS.FieldByName('ID').AsString)])
end;

//Проверяет существует ли поток по его ID (GUID) и возвращает его номер из TObjectList
function TForm1.ExistingThread(id: String): integer;
var i: integer;
begin
  Result := -1;
  if Assigned(ThreadList) then
  begin
    for i := 0 to ThreadList.Count - 1 do
    begin
       if TMyThread(ThreadList[i]).TaskID = id then
       begin
         Result := i;
         Break; // Нашли нужный поток, выходим из цикла
       end;
    end;
  end;
end;

//Запуск задания по времени
procedure TForm1.TaskExecute;
var
  bm: TBookmark;
  rec: integer;
begin
if CDS.RecordCount > 0 then
begin
  rec := -1;
  with CDS do
  begin
  bm := GetBookMark;
  DisableControls;
  try
    First;
    while not Eof do
    begin
    //Если текущее время больше времени запуска - пора запускать задание
    if (CompareDateTime(Now, CDS.FieldByName('RUN_TIME').AsDateTime)= 1) and
       (CDS.FieldByName('ACTIVE').AsBoolean = True) then
    begin
      if DirectoryExists(CDS.FieldByName('DESTINATION_PATH').AsString) then
         DeleteFiles; //Удаляем старые файлы
      DoZip;       //Запускаем архивацию
      rec:= CDS.RecNo;//Номер записи последнего запущенного процесса
      Recalc := True;//После любого запуска пересчитываем остальные
    end;
     Next;
    end;
  finally
    if BookmarkValid(bm) and (rec < 0) then
       begin
         GotoBookmark (bm);
       end else CDS.RecNo := rec;//Встаем на строку с запущенным заданием
    FreeBookmark(bm);
    EnableControls;
  end;
  end;
end;
end;

//Запуск задания немедленно
procedure TForm1.StartExecute(Sender: TObject);
begin
 DeleteFiles;
 DoZip;
end;


procedure TForm1.TimerTaskTimer(Sender: TObject);
begin
 TaskExecute;
 if Recalc then
 begin
    SetRunTime;  //Пересчитываем время следующего запуска
    Recalc := False;
 end;
end;

//Преобразование строки виде 1,3,5 в строку пн,ср,пт
function TForm1.ConvertDaysOfWeek(inputStr: string): string;
var
  inputList: TStringList;
  outputStr: string;
  i: Integer;
begin
  inputList := TStringList.Create;
  try
    inputList.Delimiter := ',';
    inputList.DelimitedText := inputStr;

    outputStr := '';
    for i := 0 to inputList.Count - 1 do
    begin
      case StrToInt(inputList[i]) of
        1: outputStr := outputStr + 'пн,';
        2: outputStr := outputStr + 'вт,';
        3: outputStr := outputStr + 'ср,';
        4: outputStr := outputStr + 'чт,';
        5: outputStr := outputStr + 'пт,';
        6: outputStr := outputStr + 'сб,';
        7: outputStr := outputStr + 'вс,';
      end;
    end;

    // Удаляем последнюю запятую из строки
    if Length(outputStr) > 0 then
      Delete(outputStr, Length(outputStr), 1);

    Result := outputStr;
  finally
    inputList.Free;
  end;
end;
procedure TForm1.AboutExecute(Sender: TObject);
begin
AboutBox.ShowModal;
end;

procedure TForm1.ClearExecute(Sender: TObject);
begin
Memo1.Clear;
end;

end.
