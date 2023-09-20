using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

namespace SkHynix.HookUpToolkit.Utils {
  public static class TypeUtils {

    public static bool IsMechanicalEquipment(FamilyInstance fi) {
      if (fi == null || !fi.IsValidObject || fi.Category == null) {
        return false;
      }

      var catId = fi.Category.Id.IntegerValue;
      if (catId != (int)BuiltInCategory.OST_MechanicalEquipment) {
        return false;
      }

      if (fi.MEPModel is MechanicalEquipment == false ||
          fi?.MEPModel?.ConnectorManager?.Connectors == null) {
        return false;
      }
      return true;
    }
  }
}
