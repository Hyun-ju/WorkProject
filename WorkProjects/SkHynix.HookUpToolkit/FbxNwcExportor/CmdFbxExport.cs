using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SkHynix.HookUpToolkit.FbxNwcExportor {
  [Transaction(TransactionMode.Manual)]
  class CmdFbxExport : IExternalCommand {
    public Result Execute(ExternalCommandData commandData,
        ref string message, ElementSet elements) {
      var uiApp = commandData.Application;
      var uidoc = uiApp.ActiveUIDocument;
      var doc = uidoc.Document;

      if (doc.ActiveView is View3D == false) { return Result.Cancelled; }

      var view = new FbxExportOptionView();
      view.VM.GetBasicDoumentInfos(doc);
      view.Owner = App.Current.MainWindow;
      view.ShowDialog();

      return Result.Succeeded;
    }
  }
}
