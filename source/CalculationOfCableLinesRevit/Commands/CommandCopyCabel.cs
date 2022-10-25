using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;

using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using System.Diagnostics;

namespace CalculationCable {
  [Transaction(TransactionMode.Manual)]
  internal class CommandCopyCabel : IExternalCommand {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {

      // Обязательные параметры в проекте
      Dictionary<string, Guid> parametrs = new Dictionary<string, Guid>() {
          {"BD_Состав кабельной продукции", new Guid("f08f11b0-abe7-4ef0-bbaf-b80fd9243814")},
          {"BD_Длина вручную", new Guid("4955cbe3-6068-46d6-a588-df76ac45e30e")},
          {"BD_Длина кабеля", new Guid("1d8966b3-d27c-4358-a95a-ad32dd46dc63")},
          {"ADSK_Количество", new Guid("8d057bb3-6ccd-4655-9165-55526691fe3a")}
        };

      List<BuiltInCategory> cats = new List<BuiltInCategory>(Options.AllCategory);

      var doc = commandData.Application.ActiveUIDocument.Document;

      // Удаление из набора отсутствующей категории в модели
      cats.RemoveAll(x => new FilteredElementCollector(doc).OfCategory(x).WhereElementIsNotElementType().FirstElement() == null);

      // Проверка наличия параметров в категориях
      Dictionary<BuiltInCategory, string> chekParam = new Dictionary<BuiltInCategory, string>();
      foreach (var cat in cats) {
        var fe = new FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().FirstElement();
        string value = "";
        foreach (var item in parametrs) {
          var p = fe.get_Parameter(item.Value);
          if (p == null) {
            if (chekParam.TryGetValue((BuiltInCategory)fe.Category.Id.IntegerValue, out value)) {
              value += "\n - " + item.Key;
              chekParam[(BuiltInCategory)fe.Category.Id.IntegerValue] = value;
            }
            else {
              value += " - " + item.Key;
              chekParam.Add((BuiltInCategory)fe.Category.Id.IntegerValue, value);
            }
          }
        }
      }

      if (chekParam.Count != 0) {
        string value = "";
        foreach (var item in chekParam) {
          value += Category.GetCategory(doc, item.Key).Name + ":\n\t" + item.Value + "\n";
          cats.Remove(item.Key); // удаляем категории где нет параметров
        }
        TaskDialog.Show("Предупреждение", $"Отсутствуют параметры в категориях:\n{value}");
        return Result.Succeeded;
      }

      // Выбор элементов с параметром BD_Состав кабельной продукции
      IList<ElementFilter> categorySet = new List<ElementFilter>(cats.Count());
      foreach (var item in cats) {
        categorySet.Add(new ElementCategoryFilter(item));
      }
      var catFilter = new LogicalOrFilter(categorySet);
      var element = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(catFilter).FirstElement();

      IList<FilterRule> rules = new List<FilterRule>();

      var parameter = element.get_Parameter(parametrs["BD_Состав кабельной продукции"]);
      if (parameter != null) {
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Состав кабельной продукции"));
        rules.Add(ParameterFilterRuleFactory.CreateContainsRule(parameter.Id, ";", false));
      }

      var filterElement = new ElementParameterFilter(rules);
      var elementSet = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(catFilter).WherePasses(filterElement).ToList();

      // Отобрать элементы у которых более одной строки в BD_Состав кабельной продукции, исключить электрооборудование без системного параметра "Длина"
      Dictionary<ElementId, string> bd_elements = new Dictionary<ElementId, string>();
      List<ElementId> clearBDParameter = new List<ElementId>();
      foreach (var item in elementSet) {
        if ((BuiltInCategory)item.Category.Id.IntegerValue == BuiltInCategory.OST_ElectricalEquipment) {
          parameter = item.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
          if (parameter == null) {
            clearBDParameter.Add(item.Id);
            continue;
          }
        }

        var count = item.get_Parameter(parametrs["BD_Состав кабельной продукции"]).AsString().Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
        if (count == 1) {
          continue;
        }
        var value = item.get_Parameter(parametrs["BD_Состав кабельной продукции"]).AsString();
        bd_elements.Add(item.Id, value);
      }

      using (var transaction = new Transaction(doc)) {
        transaction.Start("Удаление значения BD_Состав кабельной продукции");
        foreach (var item in clearBDParameter) {
          doc.GetElement(item).get_Parameter(parametrs["BD_Состав кабельной продукции"]).Set("");
        }

        foreach (var item in bd_elements) {
          var compositionCabel = item.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
          for (var i = 0; i < compositionCabel.Length; i++) {
            if (i == 0) {
              doc.GetElement(item.Key).get_Parameter(parametrs["BD_Состав кабельной продукции"]).Set(compositionCabel[i]);
            }
            else {
              var newCabel = doc.GetElement(ElementTransformUtils.CopyElement(doc, item.Key, new XYZ(0, 0, 0)).First());
              newCabel.get_Parameter(parametrs["BD_Состав кабельной продукции"]).Set(compositionCabel[i]);
            }
          }
        }

        transaction.Commit();
      }
      return Result.Succeeded;
    }
  }
}
