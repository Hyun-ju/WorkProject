using Autodesk.Revit.DB;
using CsvHelper;
using System.Collections.Generic;
using System.IO;
using SkHynix.HookUpToolkit.Utils;

namespace SkHynix.HookUpToolkit.FbxNwcExportor {
  public class CsvFileExport {
    public class CsvHeaderNames {
      public const string FileName = "FILENAME"; 
      public const string FbxId = "FBX_ID"; 
      public const string RevitId = "REVIT_ID"; 
      public const string FamilyName = "FAMILY_NAME";
      public const string FamilyType = "FAMILY_TYPE";

      public const string ObjectMinX = "OBJECT_MIN_X";
      public const string ObjectMinY = "OBJECT_MIN_Y";
      public const string ObjectMinZ = "OBJECT_MIN_Z";
      public const string ObjectMaxX = "OBJECT_MAX_X";
      public const string ObjectMaxY = "OBJECT_MAX_Y";
      public const string ObjectMaxZ = "OBJECT_MAX_Z";
    }

    public static Document _Doc = null;
    public static View3D _View = null;
    public static string _ExportSaveDirectory;
    public static bool _IsGridMode;
    public static bool _IsAMode;

    public static void ExportCsvInfo(string fbxName,
        Dictionary<string, List<Element>> workTypeDic) {
      if (App.Current._ExportorObject != null) {
        _ExportSaveDirectory = App.Current._ExportorObject.ExportSaveDirectory;
      }

      // 보안 내용 모두 삭제

      var type = string.Empty;
      if (_IsGridMode) { type = "Grid"; } else if (_IsAMode) { type = "AMode"; }
      var path = Path.Combine(_ExportSaveDirectory, $"{fbxName}_{type}.csv");

      using (var file = new StreamWriter(path, false, System.Text.Encoding.UTF8))
      using (var csv = new CsvWriter(file, System.Globalization.CultureInfo.CurrentCulture)) {
        WriteCsvHeader(csv);
        foreach (var kvp in workTypeDic) {
          foreach (var item in kvp.Value) {
            if (_IsGridMode) {
              ConstructionCsv($"{fbxName}_{kvp.Key}", item, csv);
            } else if (_IsAMode) {
              HookupCsv($"{fbxName}_{kvp.Key}", item, csv);
            }
          }
        }
        csv.Flush();
      }
    }

    private static void WriteCsvHeader(CsvWriter csv) {
      if (_IsGridMode) {
        csv.WriteField(CsvHeaderNames.FileName);
        csv.WriteField(CsvHeaderNames.RevitId);
        csv.WriteField(CsvHeaderNames.FamilyName);
        csv.WriteField(CsvHeaderNames.ObjectMinX);
        csv.WriteField(CsvHeaderNames.ObjectMinY);
        csv.WriteField(CsvHeaderNames.ObjectMinZ);
        csv.WriteField(CsvHeaderNames.ObjectMaxX);
        csv.WriteField(CsvHeaderNames.ObjectMaxY);
        csv.WriteField(CsvHeaderNames.ObjectMaxZ);
        csv.NextRecord();
      } else if (_IsAMode) {
        csv.WriteField(CsvHeaderNames.FileName);
        csv.WriteField(CsvHeaderNames.FbxId);
        csv.WriteField(CsvHeaderNames.FamilyName);
        csv.WriteField(CsvHeaderNames.ObjectMinX);
        csv.WriteField(CsvHeaderNames.ObjectMinY);
        csv.WriteField(CsvHeaderNames.ObjectMinZ);
        csv.WriteField(CsvHeaderNames.ObjectMaxX);
        csv.WriteField(CsvHeaderNames.ObjectMaxY);
        csv.WriteField(CsvHeaderNames.ObjectMaxZ);
        csv.NextRecord();
      }
    }

    private static void HookupCsv(string fbxName, Element elem, CsvWriter csv) {
      var valueNAN = "N/A";

      var familyType = elem
          .get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsValueString();
      var bb = elem.get_BoundingBox(_View);
      // 파라미터가 없을 시 null / 값이 없을 N/A
      if (elem.Category.Id.IntegerValue == ((int)BuiltInCategory.OST_MechanicalEquipment)) {

        csv.WriteField($"{CommaCheck(fbxName)}");
        csv.WriteField($"{elem.Id.IntegerValue}");
        var familyName = elem is FamilyInstance instance ? instance.Symbol.FamilyName : valueNAN;
        csv.WriteField($"{CommaCheck(familyName)}");

        csv.WriteField(ConvertUnit(bb?.Min?.X));
        csv.WriteField(ConvertUnit(bb?.Min?.Y));
        csv.WriteField(ConvertUnit(bb?.Min?.Z));
        csv.WriteField(ConvertUnit(bb?.Max?.X));
        csv.WriteField(ConvertUnit(bb?.Max?.Y));
        csv.WriteField(ConvertUnit(bb?.Max?.Z));
      } else {
        csv.WriteField($"{CommaCheck(fbxName)}");
        csv.WriteField($"{elem.Id.IntegerValue}");
        csv.WriteField($"{CommaCheck(familyType)}");

        csv.WriteField(ConvertUnit(bb?.Min?.X));
        csv.WriteField(ConvertUnit(bb?.Min?.Y));
        csv.WriteField(ConvertUnit(bb?.Min?.Z));
        csv.WriteField(ConvertUnit(bb?.Max?.X));
        csv.WriteField(ConvertUnit(bb?.Max?.Y));
        csv.WriteField(ConvertUnit(bb?.Max?.Z));
      }
      csv.NextRecord();
    }
    private static void ConstructionCsv(string fbxName, Element elem, CsvWriter csv) {
      var doc = elem.Document;
      var type = doc.GetElement(elem.GetTypeId());

      csv.WriteField($"{CommaCheck(fbxName)}");
      csv.WriteField($"{elem.Id.IntegerValue}");      
      var familyName = elem
          .get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?
          .AsValueString();
      csv.WriteField($"{CommaCheck(familyName)}");

      var bb = elem.get_BoundingBox(_View);
      csv.WriteField(ConvertUnit(bb?.Min?.X)); // EQUIPMENT_MIN_X
      csv.WriteField(ConvertUnit(bb?.Min?.Y)); // EQUIPMENT_MIN_Y
      csv.WriteField(ConvertUnit(bb?.Min?.Z)); // EQUIPMENT_MIN_Z
      csv.WriteField(ConvertUnit(bb?.Max?.X)); // EQUIPMENT_MAX_X
      csv.WriteField(ConvertUnit(bb?.Max?.Y)); // EQUIPMENT_MAX_Y
      csv.WriteField(ConvertUnit(bb?.Max?.Z)); // EQUIPMENT_MAX_Z

      csv.NextRecord();
    }

    private static string ConvertUnit(double? value) {
      if (value == null) { return string.Empty; }

      var convertValue = value.Value;
      if (convertValue.IsAlmostEqualTo(0)) { convertValue = 0; }
      convertValue = UnitUtils
          .ConvertFromInternalUnits(convertValue, UnitTypeId.Millimeters);
      return convertValue.ToString();
    }

    private static string CommaCheck(string value) {
      //if (value == null) { return value; }
      //var newValue = value;

      //if (newValue.Contains(",")) {
      //  newValue = newValue.Replace(",", "_");
      //}

      //return newValue;
      return value;
    }
  }
}
