using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using System.Drawing;
using adwin = Autodesk.Windows;

namespace SkHynix.HookUpToolkit {
  public sealed class RibbonBuilder {
    internal static class Utils {
      public static PushButtonData CreatePushButtonData(
          Type commandClass, string text,
          string tooltip, string description,
          Bitmap img16, Bitmap img32) {
        var cmdClassFullName = commandClass.FullName;
        var dataName = $"Button.Data.{cmdClassFullName}";
        return new PushButtonData(
            dataName, text,
            commandClass.Assembly.Location,
            cmdClassFullName) {
          ToolTip = tooltip,
          LongDescription = description,
          Image = null,
          LargeImage = null,
        };
      }

      public static PushButton CreatePushButton(
          RibbonPanel panel, PushButtonData data,
          string tooltip = null) {
        var button = (PushButton)panel.AddItem(data);
        button.ToolTip = tooltip ?? data.ToolTip;
        return button;
      }

      public static PushButton CreatePushButton(
          RibbonPanel panel, Type commandClass,
          string text, string tooltip, string description,
          Bitmap img16, Bitmap img32) {
        var data = CreatePushButtonData(
            commandClass, text, tooltip,
            description, img16, img32);
        return CreatePushButton(panel, data);
      }
    }

    UIControlledApplication _UICtrlApp = null;

    private static readonly string _TabName = "Tool";

    public RibbonBuilder(UIControlledApplication uiCtrlApp) {
      _UICtrlApp = uiCtrlApp;
    }

    public void CreateTab() {
      var newTab = CreateRibbonTab(_TabName);
      BuildTestPanel("Test Tab");
      var panel = _UICtrlApp.CreateRibbonPanel(_TabName, "도구");
    }

    public void Application_DocumentOpened(object sender,
        Autodesk.Revit.DB.Events.DocumentOpenedEventArgs e) {
      //SetCustomRibbonTab();
    }

    [Conditional("DEBUG")]
    private void BuildTestPanel(string tabName) {
      var panel = _UICtrlApp.CreateRibbonPanel(tabName, "Test");
      var _TestButton = Utils.CreatePushButton(
          panel, typeof(FbxNwcExportor.CmdFbxExport),
          "Test",
          null,
          null,
          null,
          null);
    }

    private adwin.RibbonTab CreateRibbonTab(string newTabName) {
      var tabs = adwin.ComponentManager.Ribbon.Tabs;
      foreach (var tab in tabs) {
        if (tab.AutomationName.Equals(newTabName)) { return null; }
      }
      var uiCtrlApp = App.Current.UIControlledApplication;
      uiCtrlApp.CreateRibbonTab(newTabName);

      tabs = adwin.ComponentManager.Ribbon.Tabs;
      adwin.RibbonTab newTab = null;
      foreach (var tab in tabs) {
        if (tab.AutomationName.Equals(newTabName)) {
          newTab = tab;
          break;
        }
      }
      return newTab;
    }
  }
}