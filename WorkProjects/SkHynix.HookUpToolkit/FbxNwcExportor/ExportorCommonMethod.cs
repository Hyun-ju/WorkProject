using Autodesk.Revit.DB;
using SkHynix.HookUpToolkit.Exportor.IPC;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace SkHynix.HookUpToolkit.FbxNwcExportor {
  public class ExportorCommonMethod {
    private static readonly string _ViewName = "Export";
    private static Regex _Regex = null;

    public static void WriteLogMessage(string errorMsg,
        ProgressStateEnum state = ProgressStateEnum.Fail) {
      var exportorObject = App.Current._ExportorObject;
      var str = $"[{state}] {exportorObject.TargetInfo.FileName} => {errorMsg}";
      File.AppendAllText(exportorObject.ResultFilePath, str + Environment.NewLine);
    }

    public static View3D GetView(Document doc) {
      var viewColl = new FilteredElementCollector(doc).OfClass(typeof(View3D));
      var views = viewColl.ToList();

      foreach (var item in views) {
        if (item is View3D view3D && item.Name.Equals(_ViewName)) {
          return view3D;
        }
      }

      // 해당 뷰가 없을 시
      WriteLogMessage($"[{_ViewName}] View is not exist", ProgressStateEnum.Pass);
      return null;
    }

    public static string RemoveInvalidChars(string str) {
      if (_Regex == null) {
        var invalidChars = Path.GetInvalidFileNameChars();
        var invalidStr = string.Empty;
        foreach (var item in invalidChars) {
          invalidStr += item;
        }
        _Regex = new Regex($"[{invalidStr}]");
      }

      var newStr = _Regex.Replace(str, "_");
      return newStr;
    }
  }
}
