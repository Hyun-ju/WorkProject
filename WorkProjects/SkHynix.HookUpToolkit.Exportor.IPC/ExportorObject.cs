using System;

namespace SkHynix.HookUpToolkit.Exportor.IPC {
  [Serializable]
  public class ExportorObject : MarshalByRefObject {
    public Guid TempGuid { get; set; }

    public bool NeedCancel { get; set; } = false;

    public string ExportSaveDirectory { get; set; }

    public bool IsFbxExportor { get; set; }
    public bool IsNwcExportor { get; set; }

    public bool IsGridMode { get; set; }
    public bool IsAMode { get; set; }
    public bool IsCropGrid { get; set; }
    public bool IsExportFbxAndCsv { get; set; }
    public bool IsExportFbx { get; set; }
    public bool IsExportCsv { get; set; }

    public string StartXGrid { get; set; }
    public string StartYGrid { get; set; }
    public string StartLevel { get; set; }
    public string EndXGrid { get; set; }
    public string EndYGrid { get; set; }
    public string EndLevel { get; set; }

    public string ResultFilePath { get; set; }

    public double Offset { get; set; }
    public string ProjectCode { get; set; }
    public bool IsSingleA { get; set; }
    public string AName { get; set; }

    public string NwcExportOptionFile { get; set; }

    public ExportInfos TargetInfo { get; set; }
  }
}
