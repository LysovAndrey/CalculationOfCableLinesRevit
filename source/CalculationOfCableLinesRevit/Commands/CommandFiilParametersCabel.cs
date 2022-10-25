using System;
using System.Linq;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;

using System.Text.RegularExpressions;

namespace CalculationCable {

  class Cables {
    public string Group { get; set; }
    public string Marks { get; set; }
    public string Length { get; set; }
  }

  [Transaction(TransactionMode.Manual)]
  internal class CommandFiilParametersCabel : IExternalCommand {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {

      // Обязательные параметры в проекте
      Dictionary<string, Guid> parametrs = new Dictionary<string, Guid>() {
        {"BD_Состав кабельной продукции", new Guid("f08f11b0-abe7-4ef0-bbaf-b80fd9243814")},
        {"BD_Марка кабеля", new Guid("049d1803-85a6-4dee-be4b-fe2eb7e5700f")},
        {"BD_Обозначение кабеля", new Guid("8e952e6b-3e8b-46f0-80c2-992ed0acd387")}
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

      // Выбор всех элементов с параметром BD_Состав кабельной продукции и BD_Марка кабеля
      IList<ElementFilter> categorySet = new List<ElementFilter>(cats.Count());
      foreach (var item in cats) {
        categorySet.Add(new ElementCategoryFilter(item));
      }
      var catFilter = new LogicalOrFilter(categorySet);
      var element = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(catFilter).FirstElement();

      IList<FilterRule> rules = new List<FilterRule>();

      var parameter = element.get_Parameter(parametrs["BD_Марка кабеля"]);
      if (parameter != null) {
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Марка кабеля"));
      }
      parameter = element.get_Parameter(parametrs["BD_Состав кабельной продукции"]);
      if (parameter != null) {
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Состав кабельной продукции"));
        rules.Add(ParameterFilterRuleFactory.CreateContainsRule(parameter.Id, ";", false));
      }

      var filterElement = new ElementParameterFilter(rules);
      var elementSet = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(catFilter).WherePasses(filterElement).ToList();

      Dictionary<ElementId, Cables> asd = new Dictionary<ElementId, Cables>();
      foreach (var item in elementSet) {
        var valueStr = item.get_Parameter(parametrs["BD_Состав кабельной продукции"]).AsString();
        try {
          string p = @";";
          Regex regex = new Regex(p);
          string[] substrings = regex.Split(valueStr);
          string pattern = "";
          switch (substrings.Length) {
            case 2:
              pattern = @"(?<mr>(.+));(\s+)?(?<ln>((\d+.)?\d+))$";
              if (Regex.IsMatch(valueStr, pattern, RegexOptions.IgnoreCase)) {
                break;
              }
              pattern = @"(?<gp>(.+));(\s+)?(?<mr>(.+))$";
              if (Regex.IsMatch(valueStr, pattern, RegexOptions.IgnoreCase)) {
                break;
              }
              pattern = @"(?<mr>(.+));";
              break;
            case 3:
              pattern = @"(?<gp>(.+));(\s+)?(?<mr>(.+));(\s+)?(?<ln>((\d+.)?\d+))";
              break;
            default:
              asd.Add(item.Id, new Cables { Group = valueStr, Marks = "", Length = "" });
              pattern = @"\w+;";
              continue;
          }

          if (Regex.IsMatch(valueStr, pattern, RegexOptions.IgnoreCase)) {
            foreach (Match match in Regex.Matches(valueStr, pattern, RegexOptions.IgnoreCase)) {
              string group = match.Groups["gp"].Value;
              group = group.Trim(';');
              if (group.Length == 0) {
                group = "0";
              }
              var marks = match.Groups["mr"].Value;
              marks = marks.Trim(';');
              var length = match.Groups["ln"].Value;
              length = length.Trim(';');
              asd.Add(item.Id, new Cables { Group = group, Marks = marks, Length = length });
            }
          }
        }
        catch (Exception) { }
      }

      using (Transaction t = new Transaction(doc)) {
        t.Start("Добавление параметра");
        foreach (var item in asd) {
          try {
            doc.GetElement(item.Key).get_Parameter(parametrs["BD_Марка кабеля"]).Set(item.Value.Group);
            doc.GetElement(item.Key).get_Parameter(parametrs["BD_Обозначение кабеля"]).Set(item.Value.Marks);
          }
          catch (Exception) {
          }
        }
        t.Commit();
      }

      return Result.Succeeded;
    }

    private int GetObjectsCount(Dictionary<BuiltInCategory, string> categoriesWithNames, Document doc)
        => categoriesWithNames.Sum(categoryWithName => FillElements(GetElements(categoryWithName.Key, doc), categoryWithName.Value));


    private IDictionary<Element, string[]> GetElements(BuiltInCategory categoryElement, Document doc) {
      var elementBunch = new FilteredElementCollector(doc)
          .OfCategory(categoryElement)
          .WhereElementIsNotElementType()
          .Where(f => !string.IsNullOrEmpty(f.LookupParameter(Options.PARAMETER_NAME)?.AsString()))
          .ToList();

      var elements = elementBunch.Select(x => new { Element = x, StrParts = x.LookupParameter(Options.PARAMETER_NAME).AsString().Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries) })
              .Where(x => x.StrParts.Length < Options.LENGTH_STR + 1 && x.StrParts.Length > 1)
              .ToDictionary(x => x.Element, y => y.StrParts);

      return elements;
    }

    private int FillElements(IDictionary<Element, string[]> elements, string errorText) {
      try {
        int i = 0;
        foreach (var element in elements) {
          element.Key.LookupParameter(Options.PARAMETER_MARK)?.Set(element.Value[0]);
          element.Key.LookupParameter(Options.PARAMETER_GROUP)?.Set(element.Value[1]);
          i++;
        }
        return i;
      }
      catch (Exception ex) {
        TaskDialog.Show($"Ошибка {errorText}", ex.Message);
      }
      return 0;
    }
  }
}
