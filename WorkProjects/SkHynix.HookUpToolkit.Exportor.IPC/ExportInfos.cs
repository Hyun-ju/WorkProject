using System;

namespace SkHynix.HookUpToolkit.Exportor.IPC {
  [Serializable]
  public class ExportInfos {
    public string FullPath { get; set; }
    public string FileName { get; set; }
    public double ProgressPercent { get; set; }
    public ProgressStateEnum State { get; set; }
    public TimeSpan ElapsedTime { get; set; }

    public string TempSavePath { get; set; }
  }
}
