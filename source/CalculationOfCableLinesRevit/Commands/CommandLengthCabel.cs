using System;
using System.Linq;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Collections.Generic;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Globalization;

namespace CalculationCable {

  [Transaction(TransactionMode.Manual)]
  internal class CommandLengthCabel : IExternalCommand {
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

      // Выбор всех элементов с параметром ADSK_Количество, BD_Состав кабельной продукции и BD_Длина вручную
      IList<ElementFilter> categorySet = new List<ElementFilter>(cats.Count());
      foreach (var item in cats) {
        categorySet.Add(new ElementCategoryFilter(item));
      }
      var catFilter = new LogicalOrFilter(categorySet);
      var element = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(catFilter).FirstElement();

      IList<FilterRule> rules = new List<FilterRule>();

      var parameter = element.get_Parameter(parametrs["ADSK_Количество"]);
      if (parameter != null) {
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("ADSK_Количество"));
      }
      parameter = element.get_Parameter(parametrs["BD_Состав кабельной продукции"]);
      if (parameter != null) {
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Состав кабельной продукции"));
        rules.Add(ParameterFilterRuleFactory.CreateContainsRule(parameter.Id, ";", false));
      }
      parameter = element.get_Parameter(parametrs["BD_Длина вручную"]);
      if (parameter != null) {
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Длина вручную"));
      }
      var filterElement = new ElementParameterFilter(rules);
      var elementSet = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(catFilter).WherePasses(filterElement).ToList();

      // Значения параметра "ADSK_Количество"
      Dictionary<ElementId, double> adsk_count = new Dictionary<ElementId, double>();
      foreach (var item in elementSet) {
        double value = 0;
        double get_value = 0;
        switch ((BuiltInCategory)item.Category.Id.IntegerValue) {
          case BuiltInCategory.OST_DuctFitting:
          case BuiltInCategory.OST_PipeFitting:
          case BuiltInCategory.OST_ConduitFitting:
            value = item.get_Parameter(parametrs["ADSK_Количество"]).AsDouble();
            value = value != 0 ? value : 0.05;
            get_value = value;
            value = UnitUtils.ConvertToInternalUnits(value, DisplayUnitType.DUT_METERS);
            break;
          default:
            value = item.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
            break;
        }
        adsk_count.Add(item.Id, value);
        Debug.WriteLine($"{item.Category.Name} - get: {value} - вставили {value}");
      }

      rules.Clear();

      // Выбор элементов у которых есть параметр "BD_Состав кабельной продукции" 
      parameter = element.get_Parameter(parametrs["BD_Состав кабельной продукции"]);
      if (parameter != null) { //
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Состав кабельной продукции"));
        rules.Add(ParameterFilterRuleFactory.CreateContainsRule(parameter.Id, ";", false));
      }

      filterElement = new ElementParameterFilter(rules);
      elementSet = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(catFilter).WherePasses(filterElement).ToList();

      Dictionary<ElementId, double> elementLength = new Dictionary<ElementId, double>();
      string pattern = @";(\s+)?(?<ln>((\d+[.,])?\d+))$";
      foreach (var item in elementSet) {
        switch ((BuiltInCategory)item.Category.Id.IntegerValue) {
          case BuiltInCategory.OST_DuctFitting:
          case BuiltInCategory.OST_PipeFitting:
          case BuiltInCategory.OST_ConduitFitting:
            continue;
        }
        double value = -1;
        var valueStr = item.get_Parameter(parametrs["BD_Состав кабельной продукции"]).AsString();
        try {
          if (Regex.IsMatch(valueStr, pattern, RegexOptions.IgnoreCase)) {
            foreach (Match match in Regex.Matches(valueStr, pattern, RegexOptions.IgnoreCase)) {
              var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator.ToString();
              var number = match.Groups["ln"].Value;
              if (decimalSeparator == ",") {
                number = number.Trim(';').Replace('.', ',');
              }
              else {
                number = number.Trim(';').Replace(',', '.');
              }
              double.TryParse(number, out value);
              elementLength.Add(item.Id, UnitUtils.ConvertToInternalUnits(value, DisplayUnitType.DUT_METERS));
              item.get_Parameter(parametrs["BD_Длина вручную"]).Set(value);
            }
          }
        }
        catch (Exception) {
        }
      }

      foreach (var item in elementLength) {
        adsk_count[item.Key] = item.Value;
      }

      // Выбор элементов у которых есть параметр "BD_Длина вручную" 
      parameter = element.get_Parameter(parametrs["BD_Длина вручную"]);
      if (parameter != null) {
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Длина вручную"));
        rules.Add(ParameterFilterRuleFactory.CreateEqualsRule(parameter.Id, 1));
      }

      filterElement = new ElementParameterFilter(rules);
      var elementIds = new FilteredElementCollector(doc).WhereElementIsNotElementType()
        .WherePasses(catFilter).WherePasses(filterElement).ToElementIds();

      foreach (var item in elementIds) {
        double value = doc.GetElement(item).get_Parameter(parametrs["BD_Длина кабеля"]).AsDouble();
        var unit = doc.GetElement(item).get_Parameter(parametrs["BD_Длина кабеля"]);
        if (unit.DisplayUnitType == DisplayUnitType.DUT_MILLIMETERS) {
          UnitUtils.ConvertToInternalUnits(value, DisplayUnitType.DUT_METERS);
        }
        adsk_count[item] = value;
      }

      FormatValueOptions valueOptions = new FormatValueOptions();
      valueOptions.SetFormatOptions(
      // Настройка отображения в метрах. Округление до 0.01, подавление нулей
      new FormatOptions(DisplayUnitType.DUT_METERS) {
        Accuracy = 0.01,
        SuppressTrailingZeros = true
      });

      using (Transaction t = new Transaction(doc)) {
        t.Start("Добавление параметра");
        foreach (var item in adsk_count) {
          string length_meters = UnitFormatUtils.Format(doc.GetUnits(), UnitType.UT_Length, item.Value, false, true, valueOptions);
          try {
            Debug.WriteLine("+");
            var d = UnitUtils.ConvertFromInternalUnits(item.Value, DisplayUnitType.DUT_METERS);
            doc.GetElement(item.Key).get_Parameter(parametrs["ADSK_Количество"]).Set(d);
            var p = doc.GetElement(item.Key).get_Parameter(parametrs["ADSK_Количество"]);
            doc.GetElement(item.Key).get_Parameter(parametrs["BD_Длина кабеля"]).Set(UnitUtils.ConvertToInternalUnits(p.AsDouble(), DisplayUnitType.DUT_METERS));
          }
          catch (Exception) {
            var p = doc.GetElement(item.Key).get_Parameter(parametrs["ADSK_Количество"]);
            doc.GetElement(item.Key).get_Parameter(parametrs["BD_Длина кабеля"]).Set(UnitUtils.ConvertToInternalUnits(p.AsDouble(), DisplayUnitType.DUT_METERS));
          }
        }
        t.Commit();
      }

      return Result.Succeeded;
    }
  }
}