using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using SkHynix.HookUpToolkit.Exportor.IPC;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;


namespace SkHynix.HookUpToolkit.Exportor.GUI {
  public class ExportorGuiVM : ViewModelBase {
    private readonly string _RevitExePath = @"C:\Program Files\Autodesk\Revit 2021\Revit.exe";

    private Window _Window = null;
    private ExportorObject _ExportorObject = null;
    private DispatcherTimer _Timer = null;
    private DispatcherTimer _NwcExportTimer = null;
    private Stopwatch _InProcessStopwatch = null;
    private Process _RvtProcess = null;
    private string _ResultLogPath = string.Empty;
    private int _ProcessedCount = 0;
    private Stopwatch _RepeatTimeStopwatch = null;

    private readonly double _KillProcessTiming
        = TimeSpan.FromHours(6).TotalMilliseconds;
    private ExportFileInfo _InProgressFile = null;

    #region Binding
    private string _ExportSaveDirectory;
    private bool _IsFbxExportor;
    private bool _IsNwcExportor;
    private bool _IsGridMode;
    private bool _IsAMode;
    private bool _IsCropGrid;
    private bool _IsExportFbxAndCsv;
    private bool _IsExportFbx;
    private bool _IsExportCsv;
    private IList<string> _XGridList;
    private IList<string> _YGridList;
    private IList<string> _LevelList;
    private string _StartGrid1;
    private string _StartGrid2;
    private string _StartLevel;
    private string _EndGrid1;
    private string _EndGrid2;
    private string _EndLevel;
    private double _Offset;
    private string _ProjectCode;
    private bool _IsSingleA;
    private string _AName;
    private ObservableCollection<ExportFileInfo> _ExportFileInfos;
    private bool _IsAllChecked;
    private bool _IsAutoClose;
    private bool _IsEnableSetting;
    private bool _CannotSetting;

    private bool _CanStopMode;
    private ObservableCollection<int> _HourList;
    private ObservableCollection<int> _MinuteList;
    private int _StartHour;
    private int _StartMinute;
    private int _NumOfRepeat;
    private double _RepeatTerm;
    private string _NwcExportOptionFile;


    public string ExportSaveDirectory {
      get { return _ExportSaveDirectory; }
      set { Set(ref _ExportSaveDirectory, value); }
    }

    public bool IsFbxExportor {
      get { return _IsFbxExportor; }
      set {
        IsEnableSetting = value;
        IsNwcExportor = !value;
        Set(ref _IsFbxExportor, value);
      }
    }
    public bool IsNwcExportor {
      get { return _IsNwcExportor; }
      set { Set(ref _IsNwcExportor, value); }
    }

    public bool IsGridMode {
      get { return _IsGridMode; }
      set {
        IsAMode = !value;
        Set(ref _IsGridMode, value);
      }
    }
    public bool IsAMode {
      get { return _IsAMode; }
      set { Set(ref _IsAMode, value); }
    }
    public bool IsCropGrid {
      get { return _IsCropGrid; }
      set { Set(ref _IsCropGrid, value); }
    }
    public bool IsExportFbxAndCsv {
      get { return _IsExportFbxAndCsv; }
      set { Set(ref _IsExportFbxAndCsv, value); }
    }
    public bool IsExportFbx {
      get { return _IsExportFbx; }
      set { Set(ref _IsExportFbx, value); }
    }
    public bool IsExportCsv {
      get { return _IsExportCsv; }
      set { Set(ref _IsExportCsv, value); }
    }

    public IList<string> XGridList {
      get { return _XGridList; }
      set { Set(ref _XGridList, value); }
    }
    public IList<string> YGridList {
      get { return _YGridList; }
      set { Set(ref _YGridList, value); }
    }
    public IList<string> LevelList {
      get { return _LevelList; }
      set { Set(ref _LevelList, value); }
    }
    public string StartXGrid {
      get { return _StartGrid1; }
      set { Set(ref _StartGrid1, value); }
    }
    public string StartYGrid {
      get { return _StartGrid2; }
      set { Set(ref _StartGrid2, value); }
    }
    public string StartLevel {
      get { return _StartLevel; }
      set { Set(ref _StartLevel, value); }
    }
    public string EndXGrid {
      get { return _EndGrid1; }
      set { Set(ref _EndGrid1, value); }
    }
    public string EndYGrid {
      get { return _EndGrid2; }
      set { Set(ref _EndGrid2, value); }
    }
    public string EndLevel {
      get { return _EndLevel; }
      set { Set(ref _EndLevel, value); }
    }
    public double Offset {
      get { return _Offset; }
      set { Set(ref _Offset, value); }
    }
    public string ProjectCode {
      get { return _ProjectCode; }
      set { Set(ref _ProjectCode, value); }
    }
    public string AName {
      get { return _AName; }
      set { Set(ref _AName, value); }
    }
    public bool IsSingleA {
      get { return _IsSingleA; }
      set { Set(ref _IsSingleA, value); }
    }

    public ObservableCollection<ExportFileInfo> ExportFileInfos {
      get { return _ExportFileInfos; }
      set { Set(ref _ExportFileInfos, value); }
    }
    public bool IsAllChecked {
      get { return _IsAllChecked; }
      set {
        foreach (var info in ExportFileInfos) {
          info.IsChecked = value;
        }
        Set(ref _IsAllChecked, value);
      }
    }

    public bool IsAutoClose {
      get { return _IsAutoClose; }
      set { Set(ref _IsAutoClose, value); }
    }
    public bool IsEnableSetting {
      get { return _IsEnableSetting; }
      set { Set(ref _IsEnableSetting, value); }
    }
    public bool CannotSetting {
      get { return _CannotSetting; }
      set {
        CanStopMode = !value;
        if (!value) { IsEnableSetting = value; }
        Set(ref _CannotSetting, value);
      }
    }

    public bool CanStopMode {
      get { return _CanStopMode; }
      set { Set(ref _CanStopMode, value); }
    }
    public ObservableCollection<int> HourList {
      get { return _HourList; }
      set { Set(ref _HourList, value); }
    }
    public ObservableCollection<int> MinuteList {
      get { return _MinuteList; }
      set { Set(ref _MinuteList, value); }
    }
    public int StartHour {
      get { return _StartHour; }
      set { Set(ref _StartHour, value); }
    }
    public int StartMinute {
      get { return _StartMinute; }
      set { Set(ref _StartMinute, value); }
    }
    public int NumOfRepeat {
      get { return _NumOfRepeat; }
      set { Set(ref _NumOfRepeat, value); }
    }
    public double RepeatTerm {
      get { return _RepeatTerm; }
      set { Set(ref _RepeatTerm, value); }
    }
    public string NwcExportOptionFile {
      get { return _NwcExportOptionFile; }
      set { Set(ref _NwcExportOptionFile, value); }
    }
    #endregion

    public RelayCommand CmdSelectDirectory { get; set; }
    public RelayCommand CmdSelectNwcOption { get; set; }
    public RelayCommand CmdAddingFiles { get; set; }
    public RelayCommand CmdRemoveFiles { get; set; }
    public RelayCommand<Window> CmdFbxExport { get; set; }
    public RelayCommand CmdStopMode { get; set; }
    public RelayCommand<Window> CmdClose { get; set; }


    public ExportorGuiVM() {
      IsFbxExportor = true;

      IsGridMode = true;
      IsExportFbxAndCsv = true;

      IsAutoClose = false;
      CannotSetting = true;

      NumOfRepeat = 1;
      RepeatTerm = 0;

      Offset = 700;

      // 주열번호
      XGridList = new List<string>();
      for (int i = 1; i <= 50; i++) {
        var grid = $"X{i}";
        XGridList.Add(grid);
      }
      YGridList = new List<string>();
      for (int i = 1; i <= 50; i++) {
        var grid = $"Y{i}";
        YGridList.Add(grid);
      }
      StartXGrid = XGridList[0];
      StartYGrid = YGridList[0];
      EndXGrid = XGridList[1];
      EndYGrid = YGridList[1];

      LevelList = new List<string>();
      for (int i = 1; i <= 11; i++) {
        var level = string.Empty;

        if (i == 1) { level = $"{i}st"; } else if (i == 2) { level = $"{i}nd"; } else if (i == 3) { level = $"{i}rd"; } else { level = $"{i}th"; }

        level += " FL";
        LevelList.Add(level);
      }
      LevelList.Add($"ROOF FL");
      StartLevel = LevelList[0];
      EndLevel = LevelList[1];

      HourList = new ObservableCollection<int>();
      MinuteList = new ObservableCollection<int>();
      for (int i = 0; i < 24; i++) {
        HourList.Add(i);
      }
      for (int i = 0; i < 60; i++) {
        MinuteList.Add(i);
      }
      StartHour = DateTime.Now.AddMinutes(1).Hour;
      StartMinute = DateTime.Now.AddMinutes(1).Minute;

      ExportFileInfos = new ObservableCollection<ExportFileInfo>();

      CmdSelectDirectory = new RelayCommand(SelectSaveDirectory);
      CmdSelectNwcOption = new RelayCommand(SelectNwcOptionFile);
      CmdAddingFiles = new RelayCommand(AddFiles);
      CmdRemoveFiles = new RelayCommand(RemoveCheckedFiles);
      CmdStopMode = new RelayCommand(StopProcess);
      CmdFbxExport = new RelayCommand<Window>(RunExport);
      CmdClose = new RelayCommand<Window>(CloseWindow);
    }


    private void SelectSaveDirectory() {
      var dlg = new FolderBrowserDialog {
        Description = "내보내기 파일 저장할 폴더를 선택해주세요",
        SelectedPath = ExportSaveDirectory,
      };
      if (dlg.ShowDialog() == DialogResult.OK) {
        ExportSaveDirectory = dlg.SelectedPath;
      }
    }
    private void SelectNwcOptionFile() {
      var dialog = new OpenFileDialog() {
        Filter = "XML File(*.xml) | *.xml",
        Multiselect = false,
      };
      if (dialog.ShowDialog() == DialogResult.OK) {
        NwcExportOptionFile = dialog.FileName;
      }
    }

    private void AddFiles() {
      var dialog = new OpenFileDialog() {
        Filter = "Revit File(*.rvt) | *.rvt",
        Multiselect = true,
      };
      if (dialog.ShowDialog() == DialogResult.OK) {
        var comparer = new ExportFileInfoEqualityComparer();
        foreach (var fileName in dialog.FileNames) {
          var fileInfo = new ExportFileInfo(fileName);
          if (ExportFileInfos.Contains(fileInfo, comparer)) { continue; }

          fileInfo.Sequence = ExportFileInfos.Count + 1;
          ExportFileInfos.Add(fileInfo);
        }
      }
      ResetState();
    }

    private void RemoveCheckedFiles() {
      var tempList = new ObservableCollection<ExportFileInfo>();
      foreach (var info in ExportFileInfos) {
        if (!info.IsChecked) { continue; }

        info.Sequence = tempList.Count + 1;
        tempList.Add(info);
      }
      ExportFileInfos = tempList;
      ResetState();
    }

    private void CloseWindow(Window win) {
      win.Close();
    }

    /// <summary>
    /// 실행하기 버튼 Command
    /// </summary>
    private void RunExport(Window win) {
      if (string.IsNullOrEmpty(ExportSaveDirectory)) {
        System.Windows.Forms.MessageBox.Show("저장 위치가 설정되지 않았습니다.");
        return;
      }

      if (IsFbxExportor) {
        if (IsGridMode) {
          if (StartXGrid.Equals(EndXGrid)) {
            System.Windows.Forms.MessageBox.Show("시작 X Grid와 종료 X Grid가 일치합니다.");
            return;
          }
          if (StartYGrid.Equals(EndYGrid)) {
            System.Windows.Forms.MessageBox.Show("시작 Y Grid와 종료 Y Grid가 일치합니다.");
            return;
          }
        }

        if (IsAMode) {
          if (IsSingleA && string.IsNullOrEmpty(AName)) {
            System.Windows.Forms.MessageBox.Show("단일 A명이 지정되지 않았습니다.");
            return;
          }
        }
      }
      if (IsNwcExportor) {
        if (string.IsNullOrEmpty(NwcExportOptionFile)) {
          System.Windows.Forms.MessageBox.Show("옵션 파일이 선택되지 않았습니다.");
          return;
        }
      }

      if (!ExportFileInfos.Any()) { return; }

      var rvtProcesses = Process.GetProcessesByName("revit");
      if (rvtProcesses != null && rvtProcesses.Any()) {
        System.Windows.Forms.MessageBox.Show($"Revit이 실행 중 입니다."
            + Environment.NewLine + "Revit을 종료 후 실행해주세요.");
        return;
      }

      if (!Utils.CreateServer()) { return; }

      CanStopMode = true;

      _ExportorObject = Utils.GetShareObject();
      _ExportorObject.NeedCancel = false;

      CannotSetting = false;
      if (_Window != null) { _Window.Closing -= Window_Closing; }
      _Window = win;
      _Window.Closing += Window_Closing;

      SettingIpcObject();
      ResetState();

      if (IsFbxExportor) {
        var queue = new Queue<ExportFileInfo>(ExportFileInfos);
        Export(queue, queue.Dequeue());
      } else if (IsNwcExportor) {
        _NwcExportTimer = new DispatcherTimer();
        _NwcExportTimer.Interval = TimeSpan.FromMinutes(1);
        _NwcExportTimer.Tick += NwcProcessTimer_Tick;
        NwcProcessTimer_Tick(null, null);
        _NwcExportTimer.Start();
      }
    }

    /// <summary>
    /// 중단 버튼 Command
    /// </summary>
    private void StopProcess() {
      if (_RvtProcess != null) {
        _RvtProcess.Kill();
      }

      CannotSetting = true;
      if (IsFbxExportor) {
        IsEnableSetting = true;
      }

      CanStopMode = false;
      if (_NwcExportTimer != null) {
        _NwcExportTimer.Tick -= NwcProcessTimer_Tick;
        _NwcExportTimer = null;
      }
      if (_Timer != null) {
        _Timer.Tick -= Timer_Tick;
        _Timer = null;
      }

      Utils.UnregisterChennel();
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
      try {
        _ExportorObject.NeedCancel = true;
        if (_RvtProcess != null) { _RvtProcess.Kill(); }
        Utils.UnregisterChennel();
      } catch (Exception ex) { }
    }

    private void SettingIpcObject() {
      if (IsFbxExportor) { SettingLogPath(_ExportSaveDirectory); }
      _ExportorObject.ResultFilePath = _ResultLogPath;

      _ExportorObject.ExportSaveDirectory = ExportSaveDirectory;
      _ExportorObject.IsFbxExportor = IsFbxExportor;
      _ExportorObject.IsNwcExportor = IsNwcExportor;

      if (IsFbxExportor) {
        _ExportorObject.IsGridMode = IsGridMode;
        _ExportorObject.IsAMode = IsAMode;
        _ExportorObject.IsCropGrid = IsCropGrid;
        _ExportorObject.IsExportFbxAndCsv = IsExportFbxAndCsv;
        _ExportorObject.IsExportFbx = IsExportFbx;
        _ExportorObject.IsExportCsv = IsExportCsv;
        _ExportorObject.StartXGrid = StartXGrid;
        _ExportorObject.StartYGrid = StartYGrid;
        _ExportorObject.StartLevel = StartLevel;
        _ExportorObject.EndXGrid = EndXGrid;
        _ExportorObject.EndYGrid = EndYGrid;
        _ExportorObject.EndLevel = EndLevel;
        _ExportorObject.Offset = Offset;
        _ExportorObject.ProjectCode = ProjectCode;
        _ExportorObject.IsSingleA = IsSingleA;
        _ExportorObject.AName = AName;
      } else if (IsNwcExportor) {
        _ExportorObject.NwcExportOptionFile = NwcExportOptionFile;
      }

      _ExportorObject.TempGuid = Guid.NewGuid();
    }

    private void Export(Queue<ExportFileInfo> queue, ExportFileInfo info) {
      if (!File.Exists(info.FullPath)) {
        info.State = ProgressStateEnum.Fail;
        var str = $"[{ProgressStateEnum.Fail}] {_InProgressFile.FileName} => 해당 파일이 없습니다";
        File.AppendAllText(_ResultLogPath, str + Environment.NewLine);

        if (queue.Any()) { Export(queue, queue.Dequeue()); } else { ExitAllProcess(); }
        return;
      }

      _InProgressFile = info;
      var path = Path.Combine(_ExportorObject.ExportSaveDirectory, info.FileName);
      path = Path.ChangeExtension(path, null);
      path = $"{path}_{_ExportorObject.TempGuid}.rvt";
      var infos = new ExportInfos {
        FullPath = info.FullPath,
        FileName = info.FileName,
        ProgressPercent = info.ProgressPercent,
        State = ProgressStateEnum.Waiting,
        ElapsedTime = new TimeSpan(),
        TempSavePath = path,
      };
      _ExportorObject.TargetInfo = infos;

      if (_Timer == null) {
        _Timer = new DispatcherTimer();
        _Timer.Interval = TimeSpan.FromMilliseconds(200);
        _Timer.Tick += Timer_Tick;
      }
      _Timer.Start();
      _InProcessStopwatch = new Stopwatch();

      if (_RvtProcess == null || _RvtProcess.HasExited) {
        _RvtProcess = new Process();
        _RvtProcess.StartInfo.FileName = _RevitExePath;
        _RvtProcess.EnableRaisingEvents = true;
        _RvtProcess.Exited += (s, e) => {
          _RvtProcess = null;
          if (File.Exists(path)) { File.Delete(path); }
          if (_Timer != null) { _Timer.Stop(); }

          if (_InProgressFile.State == ProgressStateEnum.Success) {
            var str = $"[{_InProgressFile.State}] {_InProgressFile.FileName}";
            str += $"=> 소요시간: {_InProgressFile.DisplayElapsedTime}";
            File.AppendAllText(_ResultLogPath, str + Environment.NewLine);
          }

          if (queue.Any()) { Export(queue, queue.Dequeue()); } else { ExitAllProcess(); }
        };

        _RvtProcess.Start();
      }
    }

    private void ExitAllProcess() {
      if (IsFbxExportor) {
        CanStopMode = false;
        StopProcess();
      } else if (IsNwcExportor) {
        _ProcessedCount--;
        _RepeatTimeStopwatch = new Stopwatch();
        _RepeatTimeStopwatch.Start();
      }

      if (IsAutoClose) {
        _Window.Dispatcher.Invoke(DispatcherPriority.Background, new Action(delegate {
          _Window.Close();
        }));
        _Window.Dispatcher.Invoke(DispatcherPriority.Background, new Action(delegate { }));
      }

      var succCnt = 0;
      var failCnt = 0;
      var passCnt = 0;
      foreach (var item in ExportFileInfos) {
        var state = item.State;
        if (state == ProgressStateEnum.Success) { succCnt++; } else if (state == ProgressStateEnum.Fail) { failCnt++; } else if (state == ProgressStateEnum.Pass) { passCnt++; }
      }

      File.AppendAllText(_ResultLogPath, Environment.NewLine);
      var str = $"{ExportFileInfos.Count}개 중 " +
          $"Success: {succCnt}개, Fail: {failCnt}개";
      if (passCnt > 0) { str += $", Pass: {passCnt}개"; }
      File.AppendAllText(_ResultLogPath, str + Environment.NewLine);
      File.AppendAllText(_ResultLogPath, Environment.NewLine);
    }

    private void Timer_Tick(object sender, EventArgs e) {
      _InProgressFile.State = _ExportorObject.TargetInfo.State;

      switch (_InProgressFile.State) {
        case ProgressStateEnum.InProgress:
          if (!_InProcessStopwatch.IsRunning) { _InProcessStopwatch.Start(); } else { _InProgressFile.ElapsedTime = _InProcessStopwatch.Elapsed; }
          break;
        case ProgressStateEnum.Success:
          if (_InProcessStopwatch.IsRunning) { _InProcessStopwatch.Stop(); }
          _InProgressFile.ElapsedTime = _ExportorObject.TargetInfo.ElapsedTime;
          break;
        case ProgressStateEnum.Pass:
        case ProgressStateEnum.Fail:
          if (_InProcessStopwatch.IsRunning) { _InProcessStopwatch.Stop(); }
          break;
        case ProgressStateEnum.Waiting:
        default:
          break;
      }

      if (IsFbxExportor) { return; }
      var dayDirectory = Path.Combine(ExportSaveDirectory, $"{DateTime.Today:yyyyMMdd}");
      if (_RvtProcess != null) {
        if (!Directory.Exists(dayDirectory)) { Directory.CreateDirectory(dayDirectory); }
        _ExportorObject.ExportSaveDirectory = dayDirectory;
      }
      if (_InProcessStopwatch.IsRunning
          && _InProcessStopwatch.ElapsedMilliseconds > _KillProcessTiming) {
        _InProgressFile.State = ProgressStateEnum.Fail;
        var str = $"[{ProgressStateEnum.Fail}] {_InProgressFile.FileName}" +
            $"=> Passed a lot of time!";
        File.AppendAllText(_ResultLogPath, str + Environment.NewLine);
        _RvtProcess.Kill();
      }
    }

    private void SettingLogPath(string directory) {
      var logName = IsFbxExportor ? "FBX" : "NWC";
      var nowTime = $"{DateTime.Now:yy-MM-dd_HH-mm}";
      logName = $"{logName}_{nowTime} _Result.txt";
      _ResultLogPath = Path.Combine(directory, logName);
      File.AppendAllText(_ResultLogPath, nowTime + Environment.NewLine);
      File.AppendAllText(_ResultLogPath, Environment.NewLine);
    }

    /// <summary>
    /// NWC 내보내기 Timer
    /// </summary>
    /// <remarks>
    /// -1분마다 작동하여 지정된 시작시간이 됐는지 확인 후 내보내기 Process진행<br/>
    /// -작동하는 동안 날짜가 바뀔 경우 해당 날짜 폴더 생성 및 결과 경로 변경
    /// </remarks>
    private void NwcProcessTimer_Tick(object sender, EventArgs e) {
      if (!CanStopMode) { return; }

      var dayDirectory = Path.Combine(ExportSaveDirectory, $"{DateTime.Today:yyyyMMdd}");

      var now = DateTime.Now;
      if (now.Hour == StartHour && now.Minute == StartMinute) {
        ResetState();
        _ExportorObject.ExportSaveDirectory = dayDirectory;
        if (!Directory.Exists(dayDirectory)) {
          Directory.CreateDirectory(dayDirectory);
        }
        SettingLogPath(dayDirectory);
        _ExportorObject.ResultFilePath = _ResultLogPath;

        if (_RvtProcess != null) { _RvtProcess.Kill(); }

        _ProcessedCount = NumOfRepeat + 1;
        var queue = new Queue<ExportFileInfo>(ExportFileInfos);
        Export(queue, queue.Dequeue());
      } else if (_ProcessedCount > 0 && _RepeatTimeStopwatch != null) {
        if (_RepeatTimeStopwatch.Elapsed.TotalMinutes >= RepeatTerm) {
          _RepeatTimeStopwatch.Stop();
          _RepeatTimeStopwatch = null;

          ResetState();
          var queue = new Queue<ExportFileInfo>(ExportFileInfos);
          Export(queue, queue.Dequeue());
        }
      }
    }

    private void ResetState() {
      foreach (var item in ExportFileInfos) {
        item.State = ProgressStateEnum.Waiting;
      }
    }
  }
}
