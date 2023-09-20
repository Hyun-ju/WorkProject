using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using SkHynix.HookUpToolkit.Exportor.IPC;
using SkHynix.HookUpToolkit.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SkHynix.HookUpToolkit.FbxNwcExportor {
  public class FbxExportOptionVM : ViewModelBase {
    private Document _Doc = null;
    private View3D _View = null;
    private BoundingBoxXYZ _ViewBoundBox = null;
    private ExportorObject _ExportorObject = null;

    private Dictionary<string, List<Element>> _ANameDic;
    private Dictionary<string, List<RevitLinkInstance>> _CategoryDic;
    private const string _NA = "NA";
    private static readonly List<string> _FileCodes = new List<string>() {
      "ABC",
      "DEF",
      "GHI",
    };
    private string _FileCode = _NA;
    private string _GridInfo = string.Empty;
    private const string _Roof = "ROOF FL";
    private static readonly List<string> _LevelNaming = new List<string>() {
      "st FL",
      "nd FL",
      "rd FL",
      "th FL",
      _Roof
    };


    #region 바인딩 목록
    private string _ExportSaveDirectory;
    private bool _IsGridMode;
    private bool _IsAMode;
    private bool _IsCropGrid;
    private bool _IsExportFbxAndCsv;
    private bool _IsExportFbx;
    private bool _IsExportCsv;

    private Grid _SelectedXGrid0;
    private Grid _SelectedXGrid1;
    private Grid _SelectedYGrid0;
    private Grid _SelectedYGrid1;
    private List<Grid> _XGridList;
    private List<Grid> _YGridList;
    private List<Level> _Levels;
    private Level _StartLevel;
    private Level _EndLevel;
    private double _Offset;
    private string _ProjectCode;
    private bool _IsSingleA;
    private string _AName;

    public string ExportSaveDirectory {
      get { return _ExportSaveDirectory; }
      set { Set(ref _ExportSaveDirectory, value); }
    }
    public bool IsGridMode {
      get { return _IsGridMode; }
      set { Set(ref _IsGridMode, value); }
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

    public Grid StartXGrid {
      get { return _SelectedXGrid0; }
      set { Set(ref _SelectedXGrid0, value); }
    }
    public Grid EndXGrid {
      get { return _SelectedXGrid1; }
      set { Set(ref _SelectedXGrid1, value); }
    }
    public Grid StartYGrid {
      get { return _SelectedYGrid0; }
      set { Set(ref _SelectedYGrid0, value); }
    }
    public Grid EndYGrid {
      get { return _SelectedYGrid1; }
      set { Set(ref _SelectedYGrid1, value); }
    }
    public List<Grid> XGridList {
      get { return _YGridList; }
      set { Set(ref _YGridList, value); }
    }
    public List<Grid> YGridList {
      get { return _XGridList; }
      set { Set(ref _XGridList, value); }
    }
    public List<Level> Levels {
      get { return _Levels; }
      set { Set(ref _Levels, value); }
    }
    public Level StartLevel {
      get { return _StartLevel; }
      set { Set(ref _StartLevel, value); }
    }
    public Level EndLevel {
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
    public bool IsSingleA {
      get { return _IsSingleA; }
      set { Set(ref _IsSingleA, value); }
    }
    public string AName {
      get { return _AName; }
      set { Set(ref _AName, value); }
    }

    public RelayCommand CmdSelectDirectory { get; set; }
    public RelayCommand CmdFbxExport { get; set; }
    #endregion


    public FbxExportOptionVM() {
      YGridList = new List<Grid>();
      XGridList = new List<Grid>();
      Levels = new List<Level>();

      IsGridMode = true;
      IsAMode = false;
      IsExportFbxAndCsv = true;

      Offset = 700;

      CmdSelectDirectory = new RelayCommand(SelectSaveDirectory);
      CmdFbxExport = new RelayCommand(RunButtonFbxExport);
    }

    /// <summary>
    /// document에 있는 기본 정보들을 로딩
    /// </summary>
    public void GetBasicDoumentInfos(Document doc) {
      _Doc = doc;
      if (doc.ActiveView is View3D view3D) { _View = view3D; }

      var grids = new FilteredElementCollector(doc).OfClass(typeof(Grid));
      foreach (Grid grid in grids) {
        var name = grid.Name;
        if (!name.First().Equals('X') && !name.First().Equals('Y')) { continue; }
        var gridInx = name.Remove(0, 1);
        if (!int.TryParse(gridInx, out var inx)) { continue; }

        if (grid.Curve is Line == false) { continue; }
        var line = (Line)grid.Curve;
        var direc = line.Direction.Normalize();
        if (direc.IsAlmostEqualTo(XYZ.BasisX)
            || direc.IsAlmostEqualTo(-XYZ.BasisX)) {
          YGridList.Add(grid);
        } else if (direc.IsAlmostEqualTo(XYZ.BasisY)
              || direc.IsAlmostEqualTo(-XYZ.BasisY)) {
          XGridList.Add(grid);
        }
      }

      for (int i = 0; i < XGridList.Count - 1; i++) {
        var grid0 = XGridList[i];
        for (int j = i + 1; j < XGridList.Count; j++) {
          var grid1 = XGridList[j];

          if (grid0.Curve is Line line0 && grid1.Curve is Line line1) {
            if (line0.Origin.X > line1.Origin.X) {
              var temp = grid0;
              XGridList[i] = grid1;
              XGridList[j] = temp;
              grid0 = grid1;
            }
          }
        }
      }
      for (int i = 0; i < YGridList.Count - 1; i++) {
        var grid0 = YGridList[i];
        for (int j = i + 1; j < YGridList.Count; j++) {
          var grid1 = YGridList[j];

          if (grid0.Curve is Line line0 && grid1.Curve is Line line1) {
            if (line0.Origin.Y > line1.Origin.Y) {
              var temp = grid0;
              YGridList[i] = grid1;
              YGridList[j] = temp;
              grid0 = grid1;
            }
          }
        }
      }
      if (XGridList.Count > 1) {
        StartXGrid = XGridList[0];
        EndXGrid = XGridList[1];
      }
      if (YGridList.Count > 1) {
        StartYGrid = YGridList[0];
        EndYGrid = YGridList[1];
      }

      var levels = new FilteredElementCollector(doc).OfClass(typeof(Level));
      foreach (Level level in levels) {
        var levelName = level.Name.ToUpper();

        foreach (var item in _LevelNaming) {
          if (levelName.Contains(item.ToUpper())) {
            Levels.Add(level);
            break;
          }
        }
      }
      Levels.Sort((a, b) => a.Elevation.CompareTo(b.Elevation));
      if (Levels.Count > 1) {
        StartLevel = Levels[0];
        EndLevel = Levels[1];
      }

      ExportSaveDirectory = Path.GetDirectoryName(doc.PathName);

      foreach (var naming in _FileCodes) {
        if (_Doc.Title.ToUpper().Contains(naming.ToUpper())) {
          _FileCode = naming;
        }
      }
      return;
    }

    public Result RunIpcFbxExport(Document doc) {
      _ExportorObject = App.Current._ExportorObject;
      _Doc = doc;

      var exportView = ExportorCommonMethod.GetView(doc);
      if (exportView == null) { return Result.Cancelled; }
      GetBasicDoumentInfos(doc);

      StartXGrid = null;
      EndXGrid = null;
      StartYGrid = null;
      EndYGrid = null;

      IsGridMode = _ExportorObject.IsGridMode;
      IsAMode = _ExportorObject.IsAMode;
      IsCropGrid = _ExportorObject.IsCropGrid;
      IsExportFbxAndCsv = _ExportorObject.IsExportFbxAndCsv;
      IsExportFbx = _ExportorObject.IsExportFbx;
      IsExportCsv = _ExportorObject.IsExportCsv;
      if (IsGridMode || IsCropGrid) {
        foreach (var item in XGridList) {
          var name = item.Name.ToUpper();
          if (_ExportorObject.StartXGrid.ToUpper().Equals(name)) {
            StartXGrid = item;
          }
          if (_ExportorObject.EndXGrid.ToUpper().Equals(name)) {
            EndXGrid = item;
          }
        }
        foreach (var item in YGridList) {
          var name = item.Name.ToUpper();
          if (_ExportorObject.StartYGrid.ToUpper().Equals(name)) {
            StartYGrid = item;
          }
          if (_ExportorObject.EndYGrid.ToUpper().Equals(name)) {
            EndYGrid = item;
          }
        }

        if (StartXGrid == null || EndXGrid == null ||
            StartYGrid == null || EndYGrid == null) {
          if (StartXGrid == null || StartYGrid == null) {
            var notExist = string.Empty;
            if (StartXGrid == null) { notExist = $"{_ExportorObject.StartXGrid}"; }
            if (StartYGrid == null) {
              if (!string.IsNullOrEmpty(notExist)) { notExist += $", "; }
              notExist += $"{_ExportorObject.StartYGrid}";
            }
            ExportorCommonMethod.WriteLogMessage($"선택한 시작 그리드[{notExist}]가 없습니다");
          }
          if (EndXGrid == null || EndYGrid == null) {
            var notExist = string.Empty;
            if (EndXGrid == null) { notExist = $"{_ExportorObject.EndXGrid}"; }
            if (EndYGrid == null) {
              if (!string.IsNullOrEmpty(notExist)) { notExist += $", "; }
              notExist += $"{_ExportorObject.EndYGrid}";
            }
            ExportorCommonMethod.WriteLogMessage($"선택한 종료 그리드[{notExist}]가 없습니다");
          }
          return Result.Failed;
        }
      }
      if (IsAMode) {
        ProjectCode = _ExportorObject.ProjectCode;
        IsSingleA = _ExportorObject.IsSingleA;
        AName = _ExportorObject.AName;
      }

      StartLevel = null;
      EndLevel = null;
      foreach (var item in Levels) {
        var name = item.Name.ToUpper();
        if (_ExportorObject.StartLevel.ToUpper().Equals(name)) {
          StartLevel = item;
        }
        if (_ExportorObject.EndLevel.ToUpper().Equals(name)) {
          EndLevel = item;
        }
      }
      if (_ExportorObject.EndLevel.ToUpper().Equals(_Roof.ToUpper()) && EndLevel == null) {
        EndLevel = Levels[Levels.Count - 1];
      }
      if (StartLevel == null || EndLevel == null) {
        if (StartLevel == null) {
          ExportorCommonMethod.WriteLogMessage($"선택한 시작 Level[{_ExportorObject.StartLevel}]이 없습니다");
        }
        if (EndLevel == null) {
          ExportorCommonMethod.WriteLogMessage($"선택한 종료 Level[{_ExportorObject.EndLevel}]이 없습니다");
        }
        return Result.Failed;
      }

      _View = exportView;
      Offset = _ExportorObject.Offset;

      using (var tx = new Transaction(doc,
          "FBX Export")) {
        try {
          tx.Start();
          if (_View.IsSectionBoxActive) {
            _View.IsSectionBoxActive = false;
            doc.Regenerate();
          }
          ExportByOptions();
          tx.RollBack();
        } catch (Exception ex) {
          tx.RollBack();
          ExportorCommonMethod.WriteLogMessage($"{ex}");
          return Result.Failed;
        }
      }

      return Result.Succeeded;
    }

    private void RunButtonFbxExport() {
      using (var tx = new Transaction(_Doc,
           "FBX Export")) {
        try {
          tx.Start();
          ExportByOptions();
          tx.RollBack();
        } catch (Exception ex) {
          tx.RollBack();
          var exMsg = string.Format(
              "Err Transaction Name => {0}", tx.GetName());
          Serilog.Log.Error(ex, exMsg);
        }
      }
    }

    private void SelectSaveDirectory() {
      var dlg = new FolderBrowserDialog {
        Description = "내보내기 파일 저장할 폴더를 선택해주세요",
        SelectedPath = ExportSaveDirectory,
      };
      var show = dlg.ShowDialog();
      if (show == DialogResult.OK) {
        ExportSaveDirectory = dlg.SelectedPath;
      }
    }

    private static bool IsExportElement(Element elem) {
      if (elem is FamilyInstance instance && instance.MEPModel != null) {
        return true;
      } else if (elem is MEPCurve) {
        return true;
      } else if (elem?.Category?.Id?.IntegerValue == ((int)BuiltInCategory.OST_GenericModel)) {
        return true;
      }

      return false;
    }

    private void ExportByOptions() {
      if (StartLevel == null || EndLevel == null) { return; }
      _GridInfo = _NA;

      if (IsGridMode) { GetLinkFiles(); }
      var allElems = new FilteredElementCollector(_Doc, _View.Id);
      var hideCates = new HashSet<ElementId>();
      var hideElems = new HashSet<ElementId>();
      foreach (var item in allElems) {
        if (IsAMode && IsExportElement(item)) { continue; }

        var cateId = item?.Category?.Id;
        if (cateId == null) { continue; }

        if (IsGridMode) {
          if (cateId.IntegerValue == ((int)BuiltInCategory.OST_RvtLinks)) {
            continue;
          }
        }

        if (!item.CanBeHidden(_View)) { continue; }
        if (!hideCates.Contains(cateId)) {
          hideCates.Add(cateId);
        }
        if (!hideElems.Contains(item.Id)) {
          hideElems.Add(item.Id);
        }
      }
      if (IsGridMode && hideElems.Any()) {
        _View.HideElements(hideElems);
      } else if (IsAMode && hideCates.Any()) {
        _View.HideCategoriesTemporary(hideCates);
      }
      _View.ConvertTemporaryHideIsolateToPermanent();

      _View.IsSectionBoxActive = true;
      var sb = _View.GetSectionBox();
      var ori = sb.Transform.Origin;
      var min = ori + sb.Min;
      var max = ori + sb.Max;
      var bb = new BoundingBoxXYZ {
        Min = min,
        Max = max
      };
      _ViewBoundBox = bb;
      _View.IsSectionBoxActive = false;

      if (StartLevel.Elevation > EndLevel.Elevation) {
        var tempLevel = StartLevel;
        StartLevel = EndLevel;
        EndLevel = tempLevel;
      }

      CsvFileExport._Doc = _Doc;
      CsvFileExport._ExportSaveDirectory = _ExportSaveDirectory;
      CsvFileExport._IsGridMode = _IsGridMode;
      CsvFileExport._IsAMode = _IsAMode;
      CsvFileExport._View = _View;

      if (IsGridMode) { IterateGrid(); } else if (IsAMode) { GetAById(); }
    }

    private void GetAById() {
      var elems = new FilteredElementCollector(_Doc, _View.Id);
      _ANameDic = new Dictionary<string, List<Element>>();

      foreach (var item in elems) {
        var name = item.Name;

        if (string.IsNullOrEmpty(name) &&
            item is FamilyInstance fi && TypeUtils.IsMechanicalEquipment(fi)) {
          name = item.Name;
        }
        if (string.IsNullOrEmpty(name)) { continue; }

        var temp = name.Split('_');
        if (!temp.Any()) { continue; }
        name = temp.First();

        if (IsSingleA && !AName.Equals(name)) { continue; }

        AddDictionary(name, item, ref _ANameDic);
      }

      Offset = 0;
      if (IsCropGrid) {
        StartXGrid = XGridList[0];
        EndXGrid = XGridList[XGridList.Count - 1];
        StartYGrid = YGridList[0];
        EndYGrid = YGridList[YGridList.Count - 1];
        IterateGrid();
      } else if (IsAMode) {
        var min = _ViewBoundBox.Min;
        var max = _ViewBoundBox.Max;
        CropSectionBox(min.X, min.Y, max.X, max.Y);
      }
    }

    private void GetLinkFiles() {
      var doc = _Doc;
      _CategoryDic = new Dictionary<string, List<RevitLinkInstance>>();

      var linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
      foreach (RevitLinkInstance link in linkInstances) {
        var typeId = link.GetTypeId();
        var type = doc.GetElement(typeId);
        if (type is RevitLinkType linkType) {
          if (linkType.GetLinkedFileStatus() != LinkedFileStatus.Loaded) { continue; }
        }

        // 보안내용 부분 삭제

        var categories = $"{_NA}_{_NA}";
        var names = $"{_NA}_{_NA}";

        categories = $"{names}_{categories}";
        AddDictionary(categories, link, ref _CategoryDic);
      }
    }


    private void IterateGrid() {
      var xInx0 = XGridList.IndexOf(StartXGrid);
      var xInx1 = XGridList.IndexOf(EndXGrid);
      var startX = xInx0 < xInx1 ? xInx0 : xInx1;
      var endX = xInx0 > xInx1 ? xInx0 : xInx1;

      var yInx0 = YGridList.IndexOf(StartYGrid);
      var yInx1 = YGridList.IndexOf(EndYGrid);
      var startY = yInx0 < yInx1 ? yInx0 : yInx1;
      var endY = yInx0 > yInx1 ? yInx0 : yInx1;

      for (int x = startX; x < endX; x++) {
        var xGrid0 = XGridList[x];
        var xGrid1 = XGridList[x + 1];
        var xCurve0 = ((Line)xGrid0.Curve).Origin;
        var xCurve1 = ((Line)xGrid1.Curve).Origin;
        for (int y = startY; y < endY; y++) {
          var yGrid0 = YGridList[y];
          var yGrid1 = YGridList[y + 1];
          var yCurve0 = ((Line)yGrid0.Curve).Origin;
          var yCurve1 = ((Line)yGrid1.Curve).Origin;

          var minX = xCurve0.X < xCurve1.X ? xCurve0.X : xCurve1.X;
          var maxX = xCurve0.X > xCurve1.X ? xCurve0.X : xCurve1.X;
          var minY = yCurve0.Y < yCurve1.Y ? yCurve0.Y : yCurve1.Y;
          var maxY = yCurve0.Y > yCurve1.Y ? yCurve0.Y : yCurve1.Y;

          _GridInfo = $"{xGrid0.Name}{yGrid0.Name}";
          CropSectionBox(minX, minY, maxX, maxY);
        }
      }
    }

    private void CropSectionBox(double minX, double minY, double maxX, double maxY) {
      var startInx = Levels.IndexOf(StartLevel);
      var endInx = Levels.IndexOf(EndLevel);
      for (int levelInx = startInx; levelInx <= endInx; levelInx++) {
        double minZ = 0.0;
        double maxZ = 0.0;
        var level0 = Levels[levelInx];

        if (levelInx != Levels.Count() - 1) {
          var level1 = Levels[levelInx + 1];
          minZ = level0.Elevation < level1.Elevation
              ? level0.Elevation : level1.Elevation;
          maxZ = level0.Elevation > level1.Elevation
              ? level0.Elevation : level1.Elevation;
        } else {
          minZ = level0.Elevation;
          maxZ = _ViewBoundBox.Max.Z;
          if (minZ > maxZ) {
            System.Diagnostics.Debug.WriteLine($"{level0.Name}");
            return;
          }
        }

        var levelForge = level0.get_Parameter(BuiltInParameter.LEVEL_ELEV)?.GetUnitTypeId();
        var offset = UnitUtils.ConvertToInternalUnits(Offset, levelForge);
        var minPnt = new XYZ(minX - offset, minY - offset, minZ - offset);
        var maxPnt = new XYZ(maxX + offset, maxY + offset, maxZ + offset);
        var bb = new BoundingBoxXYZ {
          Min = minPnt,
          Max = maxPnt
        };
        _View.SetSectionBox(bb);
        _Doc.Regenerate();

        ExportInSectionBox(level0);
      }
    }


    private void ExportInSectionBox(Level level) {
      var levelName = string.Empty;
      foreach (var item in level.Name) {
        if (!char.IsDigit(item)) { break; }
        levelName += item;
      }
      if (string.IsNullOrEmpty(levelName)) { levelName = level.Name; } else { levelName += "F"; }

      if (IsGridMode) {
        if (_CategoryDic != null && _CategoryDic.Any()) {
          BasisLinkFileName($"{levelName}_{_GridInfo}");
        }
      } else if (IsAMode) {
        var doc = _Doc;

        foreach (var kvp in _ANameDic) {
          var AName = kvp.Key;

          var exportElems = new List<Element>();
          var viewedElems = new FilteredElementCollector(doc, _View.Id).ToList();
          var ids = new HashSet<ElementId>();
          foreach (var item in kvp.Value) {
            ids.Add(item.Id);
          }

          foreach (var item in viewedElems) {
            if (ids.Contains(item.Id)) { exportElems.Add(item); }
          }

          if (!IsCropGrid) {
            Element equip = null;
            foreach (var item in kvp.Value) {
              if (item is FamilyInstance instance && TypeUtils.IsMechanicalEquipment(instance)) {
                equip = item;
              }
            }
            if (equip?.Location is LocationPoint lp) {
              _GridInfo = string.Empty;
              var pnt = lp.Point;
              for (int i = 0; i < XGridList.Count - 1; i++) {
                var xGrid0 = XGridList[i];
                var xGrid1 = XGridList[i + 1];
                if (xGrid0.Curve is Line line0 && xGrid1.Curve is Line line1) {
                  if (pnt.X > line0.Origin.X && pnt.X < line1.Origin.X) {
                    _GridInfo += xGrid0.Name;
                    break;
                  }
                }
              }
              for (int i = 0; i < YGridList.Count - 1; i++) {
                var yGrid0 = YGridList[i];
                var yGrid1 = YGridList[i + 1];
                if (yGrid0.Curve is Line line0 && yGrid1.Curve is Line line1) {
                  if (pnt.Y > line0.Origin.Y && pnt.Y < line1.Origin.Y) {
                    _GridInfo += yGrid0.Name;
                    break;
                  }
                }
              }
            }
            if (string.IsNullOrEmpty(_GridInfo)) { _GridInfo = _NA; }
          }

          if (!exportElems.Any()) { continue; }
          var name = $"{ProjectCode}_{_FileCode}_{AName}_{levelName}_{_GridInfo}";
          IsolateByWorkType(name, exportElems);
        }
      }
    }

    private void IsolateByWorkType(string name, List<Element> elems) {
      var dicionary = new Dictionary<string, List<Element>>();
      foreach (var item in elems) {
        if (item.Location == null) { continue; }

        var category = _NA;
        // 보안

        AddDictionary(category, item, ref dicionary);
      }

      name = ExportorCommonMethod.RemoveInvalidChars(name);
      foreach (var kvp in dicionary) {
        var fbxName = $"{name}_{kvp.Key}";
        var ids = new HashSet<ElementId>();
        foreach (var item in kvp.Value) {
          ids.Add(item.Id);
        }
        if (!ids.Any()) { continue; }

        ExoortFbx(fbxName, ids.ToList(), null);
      }

      if (!IsExportFbx) { CsvFileExport.ExportCsvInfo(name, dicionary); }
    }

    private void BasisLinkFileName(string levelGrid) {
      var categoriesDictionary = new Dictionary<string, Dictionary<string, List<Element>>>();
      foreach (var item in _CategoryDic) {
        var splits = item.Key.Split('_');
        var fileName = $"{splits[0]}_{splits[1]}";
        var cate = $"{splits[2]}_{splits[3]}";
        var fbxName = $"{fileName}_{levelGrid}_{cate}";

        var elemInLinkFile = new List<Element>();
        var viewedIds = new List<ElementId>();
        var hideIds = new List<ElementId>();
        foreach (var kvp in _CategoryDic) {
          if (!item.Key.Equals(kvp.Key)) {
            foreach (var link in kvp.Value) {
              hideIds.Add(link.Id);
            }
            continue;
          }
          foreach (var link in kvp.Value) {
            viewedIds.Add(link.Id);
            View3D view = null;
            var linkViews = new FilteredElementCollector(link.GetLinkDocument()).OfClass(typeof(View3D));
            foreach (View3D linkView in linkViews) {
              if (view == null) { view = linkView; }
              if (!linkView.IsTemplate) { view = linkView; }
            }

            var section = _View.GetSectionBox();
            var outline = new Outline(section.Min, section.Max);
            var boundingFilter = new BoundingBoxIntersectsFilter(outline);
            var linkElems = new FilteredElementCollector(link.GetLinkDocument(), view.Id).WherePasses(boundingFilter);
            elemInLinkFile.AddRange(linkElems.ToList());
          }
        }
        if (!elemInLinkFile.Any()) { continue; }

        if (categoriesDictionary.ContainsKey(fileName)) {
          categoriesDictionary[fileName].Add(cate, elemInLinkFile);
        } else {
          var midCateDic = new Dictionary<string, List<Element>>();
          midCateDic.Add(cate, elemInLinkFile);
          categoriesDictionary.Add(fileName, midCateDic);
        }

        ExoortFbx(fbxName, viewedIds, hideIds);
      }

      if (!IsExportFbx) {
        foreach (var kvp in categoriesDictionary) {
          CsvFileExport.ExportCsvInfo($"{kvp.Key}_{levelGrid}", kvp.Value);
        }
      }
      return;
    }

    private void ExoortFbx(string fbxName, List<ElementId> viewedIds, List<ElementId> hideIds) {
      if (IsExportCsv || viewedIds == null || !viewedIds.Any()) { return; }
      var doc = _Doc;
      fbxName = ExportorCommonMethod.RemoveInvalidChars(fbxName);

      using (var tx = new SubTransaction(doc)) {
        try {
          tx.Start();

          if (hideIds != null && hideIds.Any()) {
            _View.HideElements(hideIds);
          }
          _View.UnhideElements(viewedIds);
          _View.IsolateElementsTemporary(viewedIds);
          _View.ConvertTemporaryHideIsolateToPermanent();
          doc.Regenerate();

          var viewSet = new ViewSet();
          viewSet.Insert(_View);

          var directory = ExportSaveDirectory;
          if (_ExportorObject != null) { directory = _ExportorObject.ExportSaveDirectory; }
          doc.Export(directory, fbxName, viewSet, new FBXExportOptions());

          tx.RollBack();
        } catch (Exception ex) {
          var exMsg = string.Format(
              "Err Transaction Name => SubTx");
          Serilog.Log.Error(ex, exMsg);
          if (_ExportorObject != null) {
            ExportorCommonMethod.WriteLogMessage($"{ex}");
          }
          tx.RollBack();
        }
      }
    }

    private void AddDictionary<T>(string key, T value,
                                  ref Dictionary<string, List<T>> dictionary) {
      if (dictionary.ContainsKey(key)) { dictionary[key].Add(value); } else { dictionary.Add(key, new List<T>() { value }); }
    }
  }
}
