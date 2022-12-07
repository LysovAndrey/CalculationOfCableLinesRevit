using System.Linq;

using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;

namespace CalculationCable;

[Transaction(TransactionMode.Manual)]
internal class CommandCopyCabel : IExternalCommand {
  public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {

    BDCableSet cables = new BDCableSet();

    var copyElement = cables.GetElementsToCopy();

    if (copyElement != null) {

      using (Transaction t = new Transaction(cables.ActiveDocument)) {

        t.Start("Копирование элементов");

        for (int i = 0; i < copyElement.Count; i++) {
          var item = copyElement[i];
          ElementId cable_id;
          if (item.IsCopy) {
            cable_id = ElementTransformUtils.CopyElement(Plugin.ActiveDocument, item.Element.Id, new XYZ(0, 0, 0)).First();
            item.IsCopy = false;
          }
          else {
            cable_id = item.Element.Id;
          }
          Element cabel = Plugin.ActiveDocument.GetElement(cable_id);
          cabel = Plugin.ActiveDocument.GetElement(cable_id);
          cabel.get_Parameter(Global.Parameters["BD_Состав кабельной продукции"]).Set(item.Source);
          cabel.get_Parameter(Global.Parameters["BD_Марка кабеля"]).Set(item.Group);
          cabel.get_Parameter(Global.Parameters["BD_Обозначение кабеля"]).Set(item.CableType);
          cabel.get_Parameter(Global.Parameters["BD_Длина вручную"]).Set(-1);
          cabel.get_Parameter(Global.Parameters["BD_Длина кабеля"]).Set(item.Length);
          try {
            cabel.get_Parameter(Global.Parameters["ADSK_Количество"]).Set(item.Quantity);
          }
          catch (System.Exception) { }
        }
        t.Commit();
      }
    }
    return Result.Succeeded;
  }
}
