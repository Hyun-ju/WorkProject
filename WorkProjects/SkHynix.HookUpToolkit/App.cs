using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Serilog;
using SkHynix.HookUpToolkit.Exportor.IPC;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Interop;
using UIFramework;
using Events = Autodesk.Revit.UI.Events;
using IpcUtils = SkHynix.HookUpToolkit.Exportor.IPC.Utils;

namespace SkHynix.HookUpToolkit {
  public class App : IExternalApplication {
    public static App Current { get; private set; }

    public MainWindow MainWindow { get; private set; }
    public UIControlledApplication UIControlledApplication { get; private set; }
    public UIApplication UIApplication { get; private set; }
    public RibbonBuilder RibbonBuilder { get; private set; }
    public ILogger Logger { get; private set; }
    public string UserName => UIApplication.Application.Username;

    public ExportorObject _ExportorObject { get; private set; }

    public Result OnShutdown(UIControlledApplication application) {
      return Result.Succeeded;
    }

    public Result OnStartup(UIControlledApplication application) {
      InitializeFields(application);

      RibbonBuilder.CreateTab();

      var app = application.ControlledApplication;
      var client = IpcUtils.CreateClient();
      if (client) {
        application.Idling += Application_Idling;
      }

      return Result.Succeeded;
    }

    private void Application_Idling(object sender, Events.IdlingEventArgs e) {
      ExportInfos tarInfo = null;
      try {
        _ExportorObject = IpcUtils.GetShareObject();

        tarInfo = _ExportorObject.TargetInfo;
        tarInfo.State = ProgressStateEnum.InProgress;
        _ExportorObject.TargetInfo = tarInfo;
      } catch (Exception ex) {
        if (sender is UIApplication uiApp) {
          uiApp.Idling -= Application_Idling;
          _ExportorObject = null;
          return;
        }
      }

      Document doc = null;
      try {
        var sw = new System.Diagnostics.Stopwatch();
        sw.Start();
        Result result = Result.Failed;

        var app = Current.UIApplication.Application;
        var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(_ExportorObject.TargetInfo.FullPath);
        var openOptions = new OpenOptions() {
          DetachFromCentralOption = DetachFromCentralOption.DetachAndDiscardWorksets
        };
        doc = app.OpenDocumentFile(modelPath, openOptions);
        //doc = app.OpenDocumentFile(_ExportorObject.TargetInfo.FullPath);
        if (_ExportorObject.IsFbxExportor) {
          var export = new FbxNwcExportor.FbxExportOptionVM();
          result = export.RunIpcFbxExport(doc);
        } else if (_ExportorObject.IsNwcExportor) {
          var export = new FbxNwcExportor.NavisworksExportVM();
          result = export.RunIpcNwcExport(doc);
        }

        sw.Stop();

        tarInfo = _ExportorObject.TargetInfo;
        tarInfo.ElapsedTime = sw.Elapsed;
        if (result == Result.Succeeded) {
          tarInfo.State = ProgressStateEnum.Success;
        } else if (result == Result.Cancelled) {
          tarInfo.State = ProgressStateEnum.Pass;
        } else {
          tarInfo.State = ProgressStateEnum.Fail;
        }
        _ExportorObject.TargetInfo = tarInfo;
      } catch (Exception ex) {
        tarInfo = _ExportorObject.TargetInfo;
        tarInfo.State = ProgressStateEnum.Fail;
        _ExportorObject.TargetInfo = tarInfo;

        var str = $"[{ProgressStateEnum.Fail}] {_ExportorObject.TargetInfo.FileName} => {ex}";
        File.AppendAllText(_ExportorObject.ResultFilePath, str + Environment.NewLine);
      }
      return;
    }

    private void InitializeFields(UIControlledApplication application) {
      Current = this;
      UIControlledApplication = application;

      {
        var fieldName = "m_uiapplication";
        var fi = UIControlledApplication.GetType().GetField(
            fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
        UIApplication = (UIApplication)fi.GetValue(UIControlledApplication);
      }

      {
        var hwndSrc = HwndSource.FromHwnd(
            UIControlledApplication.MainWindowHandle);
        MainWindow = hwndSrc.RootVisual as MainWindow;
      }
    }
  }
}
