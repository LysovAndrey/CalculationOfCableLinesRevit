using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI.Selection;

namespace CalculationCable {

  /// <summary>
  /// Копирование параметра BD_Состав кабельно продукции из элемента в буфер
  /// </summary>
  [Transaction(TransactionMode.Manual)]
  internal class CopyToBufferBD_СableСomposition : IExternalCommand {

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {

      BDCableSet bd = new BDCableSet();
 
      return Result.Succeeded;

      /*
      Document doc = commandData.Application.ActiveUIDocument.Document;
      Selection selection = commandData.Application.ActiveUIDocument.Selection;

      IList<Reference> references = null;

      try {
        references = selection.PickObjects(ObjectType.Element, new SelectCable(), "Выбирете элементы с параметром BD_Состав кабельной продукции");
      }
      catch (Autodesk.Revit.Exceptions.OperationCanceledException) {
        return Result.Cancelled;
      }

      List<Element> warning_elements = new List<Element>();
      List<Element> ok_elements = new List<Element>();

      foreach (var r in references) {
        Element elem = doc.GetElement(r);
        string value = elem.get_Parameter(Options.Parameter["BD_Состав кабельной продукции"]).AsString();
        if (string.IsNullOrEmpty(value)) {
          continue;
        }
        if (value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length > 1) {
          warning_elements.Add(elem);
          continue;
        }
        ok_elements.Add(elem);
      }


      Debug.WriteLine($"Хороших элементов - \t{ok_elements.Count}");
      Debug.WriteLine($"Плохих элементов - \t{warning_elements.Count}");
      */
    }
  }


  /// <summary>
  /// Фильтр выбора элементов с параметром BD_Состав кабельной продукции
  /// </summary>
  //internal class SelectCable : ISelectionFilter {

  //  public bool AllowElement(Element elem) {

  //    if (elem.Category != null && Options.AllCategory.Exists(x => x == (BuiltInCategory)elem.Category.Id.IntegerValue)) {
  //      Parameter paramBD = elem.get_Parameter(new Guid("f08f11b0-abe7-4ef0-bbaf-b80fd9243814"));
  //      Parameter paramLen = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
  //      if (paramBD != null) {
  //        if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_ElectricalEquipment && paramLen == null) {
  //          return false;
  //        }
  //        return true;
  //      }
  //    }
  //    return false;
  //  }

  //  public bool AllowReference(Reference reference, XYZ position) {
  //    return true;
  //  }
  //}
}
