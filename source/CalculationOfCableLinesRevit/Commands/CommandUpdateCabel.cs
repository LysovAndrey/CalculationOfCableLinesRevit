using System;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace CalculationCable;

/// <summary>
/// Копирование параметра BD_Состав кабельно продукции из элемента в буфер
/// </summary>
[Transaction(TransactionMode.Manual)]
internal class CommandUpdateCabel : IExternalCommand {

  public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {

    BDCableSet cables = new BDCableSet();

    var updated = cables.GetUpdatedElements();

    // var pagsing = cables.GetElementsToErrorParsing();

    if (updated != null && updated.Count != 0) {

      using (Transaction t = new Transaction(cables.ActiveDocument)) {
        t.Start("Заполнение параметров");
        foreach (var item in updated) {
          var elem = Plugin.ActiveDocument.GetElement(item.Element.Id);
          elem.get_Parameter(Global.Parameters["BD_Марка кабеля"]).Set(item.Group);
          elem.get_Parameter(Global.Parameters["BD_Обозначение кабеля"]).Set(item.CableType);
          elem.get_Parameter(Global.Parameters["BD_Длина кабеля"]).Set(item.Length);
          try {
            elem.get_Parameter(Global.Parameters["ADSK_Количество"]).Set(item.Quantity);
          }
          catch (Exception) {
          }
        }
        t.Commit();
      }
    }
    return Result.Succeeded;
  }
}