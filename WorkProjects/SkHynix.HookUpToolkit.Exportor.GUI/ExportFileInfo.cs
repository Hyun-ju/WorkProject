using GalaSoft.MvvmLight;
using SkHynix.HookUpToolkit.Exportor.IPC;
using System;
using System.Collections.Generic;
using System.IO;

namespace SkHynix.HookUpToolkit.Exportor.GUI {
  public class ExportFileInfo : ObservableObject {
    private bool _IsChecked;
    private int _Sequence;
    private string _FullPath;
    private string _FileName;
    private double _ProgressPercent;
    private ProgressStateEnum _State;
    private TimeSpan _ElapsedTime;
    private string _DisplayElapsedTimeStr;

    public bool IsChecked {
      get { return _IsChecked; }
      set { Set(ref _IsChecked, value); }
    }
    public int Sequence {
      get { return _Sequence; }
      set { Set(ref _Sequence, value); }
    }
    public string FullPath {
      get { return _FullPath; }
      set { Set(ref _FullPath, value); }
    }
    public string FileName {
      get { return _FileName; }
      set { Set(ref _FileName, value); }
    }

    public double ProgressPercent {
      get { return _ProgressPercent; }
      set { Set(ref _ProgressPercent, value); }
    }
    public ProgressStateEnum State {
      get { return _State; }
      set {
        if (value != ProgressStateEnum.Success
            && value != ProgressStateEnum.InProgress) {
          DisplayElapsedTime = "-";
        }
        Set(ref _State, value);
      }
    }
    public TimeSpan ElapsedTime {
      get { return _ElapsedTime; }
      set {
        DisplayElapsedTime = TimeSpanToString(value);
        Set(ref _ElapsedTime, value);
      }
    }
    public string DisplayElapsedTime {
      get { return _DisplayElapsedTimeStr; }
      set { Set(ref _DisplayElapsedTimeStr, value); }
    }


    public ExportFileInfo() {
      IsChecked = false;
      ProgressPercent = 0;
      State = ProgressStateEnum.Waiting;
      ElapsedTime = new TimeSpan();
      DisplayElapsedTime = "-";
    }
    public ExportFileInfo(string fullPath) {
      FullPath = fullPath;
      FileName = Path.GetFileName(fullPath);
      IsChecked = false;
      ProgressPercent = 0;
      State = ProgressStateEnum.Waiting;
      ElapsedTime = new TimeSpan();
      DisplayElapsedTime = "-";
    }


    private string TimeSpanToString(TimeSpan timeSpan) {
      var time = string.Empty;
      if (timeSpan.Hours > 0) { time += $"{timeSpan.Hours}h "; }
      var milliSecond = timeSpan.Milliseconds;
      if (timeSpan.Milliseconds > 100) { milliSecond /= 10; }
      var milliSec = string.Format("{0,2:00}", milliSecond);
      time += $"{timeSpan.Minutes}m {timeSpan.Seconds}.{milliSec}s";
      return time;
    }
  }


  public class ExportFileInfoEqualityComparer : IEqualityComparer<ExportFileInfo> {
    public bool Equals(ExportFileInfo x, ExportFileInfo y) {
      return x.FullPath.Equals(y.FullPath);
    }

    public int GetHashCode(ExportFileInfo obj) {
      return 0;
    }
  }
}
