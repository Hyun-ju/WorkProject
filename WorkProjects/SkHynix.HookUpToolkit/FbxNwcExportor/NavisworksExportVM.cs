using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace SkHynix.HookUpToolkit.FbxNwcExportor {
  public class NavisworksExportVM : ViewModelBase {
    private class OptionNames {
      public const string NwExportRevit = "nwexportrevit";
      public const string Parameters = "nwexportrevit_element_params";
      public const string ElementIds = "nwexportrevit_element_ids";
      public const string FindMissingMaterials = "nwexportrevit_element_find_missing_materials";
      public const string ExportScope = "nwexportrevit_section_extract";
      public const string Coordinates = "nwexportrevit_coordinates";
      public const string ConvertElementProperties = "nwexportrevit_element_generic_properties";
      public const string ExportUrls = "nwexportrevit_urls";
      public const string ExportRoomAsAttribute = "nwexportrevit_room";
      public const string ExportRoomGeometry = "nwexportrevit_room_geometry";
      public const string ExportLinks = "nwexportrevit_linked_files";
      public const string ExportParts = "nwexportrevit_construction_parts";
      public const string DivideFileIntoLevels = "nwexportrevit_divide_file_into_levels";
      public const string FacetingFactor = "nwexportrevit_param_faceting_factor";
      public const string ConvertLinkedCADFormats = "nwexportrevit_linked_CAD_formats";
      public const string ConvertLights = "nwexportrevit_lights";
    }

    private UIDocument _UiDoc = null;
    private Document _Doc = null;

    private string _ExportSaveDirectory;
    private string _ExportOptionFile;

    public string ExportSaveDirectory {
      get { return _ExportSaveDirectory; }
      set { Set(ref _ExportSaveDirectory, value); }
    }
    public string ExportOptionFile {
      get { return _ExportOptionFile; }
      set { Set(ref _ExportOptionFile, value); }
    }


    public RelayCommand CmdSelectDirectory { get; set; }
    public RelayCommand CmdNwcExport { get; set; }


    public NavisworksExportVM() {
      ExportSaveDirectory = string.Empty;

      CmdSelectDirectory = new RelayCommand(SelectSaveDirectory);
      CmdNwcExport = new RelayCommand(RunExportNwc);
    }


    public void GetDocument(UIDocument uiDoc) {
      _UiDoc = uiDoc;
      _Doc = uiDoc.Document;

      ExportSaveDirectory = Path.GetDirectoryName(uiDoc.Document.PathName);
    }


    public Result RunIpcNwcExport(Document doc) {
      var exportorObject = App.Current._ExportorObject;
      _Doc = doc;

      var view = ExportorCommonMethod.GetView(_Doc);
      if (view == null) { return Result.Cancelled; }

      ExportSaveDirectory = exportorObject.ExportSaveDirectory;
      ExportOptionFile = exportorObject.NwcExportOptionFile;
      try {
        ExportNwc(view.Id);
      } catch (Exception ex) {
        ExportorCommonMethod.WriteLogMessage($"{ex}");
        return Result.Failed;
      }

      return Result.Succeeded;
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

    private void RunExportNwc() {
      try {
        ExportNwc(_Doc.ActiveView.Id);
      } catch (Exception ex) {
        var exMsg = string.Format(
            "Err => Navisworks Export Error!");
        Serilog.Log.Error(ex, exMsg);
      }
    }

    /// <summary>
    /// NMC 내보내기
    /// </summary>
    /// <param name="viewId"></param>
    /// <remarks>Transcation을 열어 실행할 시 내보내기 진행안됨</remarks>
    private void ExportNwc(ElementId viewId) {
      var doc = _Doc;

      var linkTypes = new FilteredElementCollector(doc)
          .OfClass(typeof(RevitLinkType)).ToElements();
      foreach (var link in linkTypes) {
        if (link is RevitLinkType linkType) {
          if (linkType.GetLinkedFileStatus() == LinkedFileStatus.Loaded) {
            linkType.Unload(null);
          }
        }
      }

      var navisworksOptions = SetExportOptions();
      navisworksOptions.ViewId = viewId;

      var directory = ExportSaveDirectory;
      if (App.Current._ExportorObject != null) {
        directory = App.Current._ExportorObject.ExportSaveDirectory;
      }
      doc.Export(directory, $"{doc.Title}", navisworksOptions);
    }

    private NavisworksExportOptions SetExportOptions() {
      var options = new NavisworksExportOptions();

      var file = ExportOptionFile;
      XmlDocument doc = new XmlDocument();
      doc.Load(file);

      var nodeList = doc.GetElementsByTagName("option");
      foreach (XmlNode node in nodeList) {
        var attris = node.Attributes;
        var name = string.Empty;
        foreach (XmlAttribute attri in attris) {
          if (attri.Value.Contains(OptionNames.NwExportRevit)) {
            name = attri.Value;
            break;
          }
        }

        var value = node.InnerText;
        foreach (XmlElement data in node) {
          var elementAttr = data.GetAttribute("type");
          if (elementAttr.Equals("name")) {
            foreach (XmlElement innerXml in data) {
              value = innerXml.GetAttribute("internal");
              if (value.Contains($"{name}:")) {
                value = value.Replace($"{name}:", string.Empty);
              }
            }
          }
        }

        var isBoolean = bool.TryParse(value, out var boolean);
        var isInt = int.TryParse(value, out var intValue);
        var IsDouble = double.TryParse(value, out var doubleValue);
        switch (name) {
          case OptionNames.Parameters:
            if (isInt) { options.Parameters = (NavisworksParameters)intValue; }
            break;
          case OptionNames.ElementIds:
            if (isBoolean) { options.ExportElementIds = boolean; }
            break;
          case OptionNames.FindMissingMaterials:
            if (isBoolean) { options.FindMissingMaterials = boolean; }
            break;
          case OptionNames.ExportScope:
            if (isInt) { options.ExportScope = (NavisworksExportScope)intValue; }
            break;
          case OptionNames.Coordinates:
            if (isInt) { options.Coordinates = (NavisworksCoordinates)intValue; }
            break;
          case OptionNames.ConvertElementProperties:
            if (isBoolean) { options.ConvertElementProperties = boolean; }
            break;
          case OptionNames.ExportUrls:
            if (isBoolean) { options.ExportUrls = boolean; }
            break;
          case OptionNames.ExportRoomAsAttribute:
            if (isBoolean) { options.ExportRoomAsAttribute = boolean; }
            break;
          case OptionNames.ExportRoomGeometry:
            if (isBoolean) { options.ExportRoomGeometry = boolean; }
            break;
          case OptionNames.ExportLinks:
            if (isBoolean) { options.ExportLinks = boolean; }
            break;
          case OptionNames.ExportParts:
            if (isBoolean) { options.ExportParts = boolean; }
            break;
          case OptionNames.DivideFileIntoLevels:
            if (isBoolean) { options.DivideFileIntoLevels = boolean; }
            break;
          case OptionNames.FacetingFactor:
            if (IsDouble) { options.FacetingFactor = doubleValue; }
            break;
          case OptionNames.ConvertLinkedCADFormats:
            if (isBoolean) { options.ConvertLinkedCADFormats = boolean; }
            break;
          case OptionNames.ConvertLights:
            if (isBoolean) { options.ConvertLights = boolean; }
            break;
        }
      }

      return options;
    }
  }
}