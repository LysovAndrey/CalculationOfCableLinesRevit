using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CalculationCable {

  internal enum BDStatus {
    None = 5000,     // не определенно
    Normal = 5100,   // Нормально
    Error = -5001,   // Ошибка
    Cancel = -5002,  // Отмена
    Failed = -5003   // Неудача
  }

  internal class MyCable {

    public ElementId Id { get; set; }
    public string Source { get; set; } // Исходное значение параметра BD_Состав кабельно продукции
    public string Group { get; set; } // Номер группы
    public string CableType { get; set; } // Тип, марка кабеля
    public string Length { get; set; } // Длина участка кабеля

    /// <summary>
    /// Сравнение параметра "BD_Состав кабельной продукции"
    /// </summary>
    /// <param name="source">Значение парметра "BD_Состав кабельной продукции"</param>
    /// <returns>true если параметры равны, иначе false</returns>
    public bool CompareSource(string source) {
      if (Source.Contains(source)) { return true; }
      return false;
    }
  }

  internal class BDCableSet {

    public BDCableSet(Document document) {

      m_document = document;

      m_categories.RemoveAll(c => new FilteredElementCollector(document)
      .OfCategory(c)
      .WhereElementIsNotElementType()
      .FirstElement() == null);

      if (HasParameters() == BDStatus.Normal) { }
    }

    public BDCableSet(List<Element> elements) {
      foreach (var elem in elements) {
        Add(elem);
      }
    }

    public MyCable Current => m_cables.GetEnumerator().Current;

    public void Add(Element elem) {

      var value = elem.get_Parameter(Options.Parameter["BD_Состав кабельной продукции"]).AsString();

      if (value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length > 1) {
        return;
      }
      Regex regex = new Regex(@";");
      string[] substrings = regex.Split(value);
      string pattern;
      try {
        switch (substrings.Length) {
          case 2:
            //    <mr>       <ln>
            // Тип кабеля; 200{{,.}0}
            pattern = @"(?<mr>(.+));(\s+)?(?<ln>((\d+[.,])?\d+))$";
            if (Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase)) {
              break;
            }
            // <gp>    <mr>   
            // Гр1; Тип кабеля
            pattern = @"(?<gp>(.+));(\s+)?(?<mr>(.+))$";
            if (Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase)) {
              break;
            }
            // <gp>
            // Гр1;
            pattern = @"(?<gp>(.+));";
            break;
          case 3:
            // <gp>    <mr>       <ln>
            // Гр1; Тип кабеля; 200{{,.}0}
            pattern = @"(?<gp>(.+));(\s+)?(?<mr>(.+));(\s+)?(?<ln>((\d+[.,])?\d+))";
            break;
          default:
            // Список с неправильным заполнением парметра BD_Состав кабельной продукции
            //_cables.Add(new MyCable { Group = value, Marks = "", Length = "", Id = elem.Id });
            pattern = @"\w+;";
            break;
        }

        if (Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase)) {
          foreach (Match match in Regex.Matches(value, pattern, RegexOptions.IgnoreCase)) {
            var сableType = match.Groups["mr"].Value;
            string group = match.Groups["gp"].Value;
            if (group.Length == 0) {
              group = "0";
            }
            var length = match.Groups["ln"].Value; // Перевод в double

            m_cables.Add(new MyCable { Id = elem.Id, Source = value, Group = group, CableType = сableType, Length = length });
          }
        }
      }
      catch (Exception) { }

    }

    public List<MyCable> ToList() { return m_cables; }

    public IEnumerator GetEnumerator() => m_cables.GetEnumerator();

    /// <summary>
    /// Проверка обязательных параметров в модели.
    /// </summary>
    BDStatus HasParameters() {

      if (m_categories.Count == 0) {
        return BDStatus.Cancel;
      }

      // string наименование параметра
      // Если List<Category> = null параметра в модели нет
      Dictionary<string, List<Category>> bind_params = new Dictionary<string, List<Category>>();

      foreach (var categoryId in m_categories.Reverse<BuiltInCategory>()) {
        List<Category> categories = null;
        foreach (var param in m_params) {
          Definition definition_ = FindParameterInProject(m_document, param.Value);
          if (definition_ == null) {
            if (!bind_params.TryGetValue(param.Key, out categories)) {
              bind_params.Add(param.Key, null);
            }
            continue;
          }
          Category category = Category.GetCategory(m_document, categoryId);
          ElementBinding elem_bind = m_document.ParameterBindings.get_Item(definition_) as ElementBinding;
          if (elem_bind.Categories.Contains(category)) {
            // Включаем категорию "Электрооборудование" если семейство выпонено на основе линии
            Element element = new FilteredElementCollector(m_document).OfCategory(categoryId).WhereElementIsCurveDriven().FirstElement();
            if (element == null && category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_ElectricalEquipment)) {
              m_categories.Remove(categoryId);
              break;
            }
            // TODO: Целесообразность!!!
            //else { 
            //  if (!m_categories.Contains(categoryId)) {
            //    m_categories.Add(categoryId);
            //  }
            //}
            continue;
          }
          if (bind_params.TryGetValue(param.Key, out categories)) {
            categories.Add(category);
          }
          else {
            bind_params.Add(param.Key, new List<Category>() { { category } });
          }
        }
      }

      if (bind_params.Count == 0) {
        return BDStatus.Normal;
      }

      // Добавление параметров в модель
      string error_messages = "";
      string warning_messages = "";

      foreach (var bp in bind_params) {
        if (bp.Value == null) {
          error_messages += "   - " + bp.Key + ";\n";
        }
        else {
          warning_messages += $"\nПарметр \"{bp.Key}\" не назнечен категориям:\n";
          bp.Value.ForEach(x => warning_messages += "   - " + x.Name + ";\n");
        }
      }

      string info = "\n";
      if (!string.IsNullOrEmpty(error_messages)) {
        info += $"В модели отсутствуют параметры:\n{error_messages}";
      }
      if (!string.IsNullOrEmpty(warning_messages)) {
        info += warning_messages;
      }

      TaskDialog dialog = new TaskDialog("Предупреждение.");
      dialog.MainInstruction = $"{info}";
      dialog.CommonButtons = TaskDialogCommonButtons.Close;
      dialog.DefaultButton = TaskDialogResult.Close;
      dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Добавить параметры");
      TaskDialogResult res = dialog.Show();
      if (res == TaskDialogResult.CommandLink1) {

        Dictionary<Definition, CategorySet> add_params = new Dictionary<Definition, CategorySet>();
        CategorySet add_categories = m_document.Application.Create.NewCategorySet();

        // Привязка параметров из модели
        foreach (var item in bind_params.Reverse()) {
          if (item.Value != null) {
            Definition definition_ = FindParameterInProject(m_document, m_params[item.Key]);
            ElementBinding element_bind = m_document.ParameterBindings.get_Item(definition_) as ElementBinding;
            element_bind.Categories.Cast<Category>().ToList().ForEach(c => add_categories.Insert(c));
            item.Value.ForEach(c => add_categories.Insert(c));
            add_params.Add(definition_, add_categories);
            bind_params.Remove(item.Key);
          }
        }

        Binding typeBinding = m_document.Application.Create.NewInstanceBinding(add_categories);

        using (Transaction tr = new Transaction(m_document)) {
          tr.Start("Добавление параметров - CalculationCable");
          foreach (var item in add_params) {
            m_document.ParameterBindings.ReInsert(item.Key, typeBinding);
          }
          tr.Commit();
        }

        if (bind_params.Count == 0) { 
          return BDStatus.Normal; 
        }

        add_categories.Clear();
        add_params.Clear();

        m_categories.ForEach(c => add_categories.Insert(Category.GetCategory(m_document, c)));

        // Добавление пармаетров из ФОП
        foreach (var item in bind_params.Reverse()) {
          if (item.Value == null) {
            Definition definition_ = FindParameterInFile(m_document, m_params[item.Key]);
            if (definition_ == null) { continue; }
            add_params.Add(definition_, add_categories);
            bind_params.Remove(item.Key);
          }
        }

        using (Transaction tr = new Transaction(m_document)) {
          tr.Start("Добавление параметров - CalculationCable");
          foreach (var item in add_params) {
            m_document.ParameterBindings.Insert(item.Key, typeBinding);
          }
          tr.Commit();
        }

        BDStatus status = BDStatus.Normal;

        if (bind_params.Count != 0) {
          string file_name = m_document.Application.SharedParametersFilename;
          var error_message = "";
          bind_params.ToList().ForEach(s => error_message += "\nПараметр: " + s.Key + "\nGuid: " + m_params[s.Key].ToString() + "\n");
          TaskDialog dialog_file = new TaskDialog("Ошибка.");
          dialog_file.MainInstruction = $"В файле общих параметров \"{file_name}\" отсутствуют параметры:\n{error_message}";
          dialog_file.CommonButtons = TaskDialogCommonButtons.Close;
          dialog_file.DefaultButton = TaskDialogResult.Close;
          dialog_file.Show();
          status = BDStatus.Error;
        }
        return status;
      }
      return BDStatus.Failed;
    }

    /// <summary>
    /// Поиск параметра в проекте
    /// </summary>
    /// <param name="document">Текущий документ</param>
    /// <param name="guid">Guid парметра</param>
    /// <returns>null если параметр не найден</returns>
    Definition FindParameterInProject(Document document, Guid guid) {
      Definition definition = null;
      try {
        BindingMap bindingMap = document.ParameterBindings;
        var it = bindingMap.ForwardIterator();
        it.Reset();
        while (it.MoveNext()) {
          var def = (InternalDefinition)it.Key;
          var sharedParameter = document.GetElement(def.Id) as SharedParameterElement;
          if (sharedParameter != null) {
            if (sharedParameter.GuidValue.Equals(guid))
              return definition = def;
          }
        }
      }
      catch (Exception) {
      }
      return definition;
    }

    /// <summary>
    /// Проверяет привязку парметра в категории.
    /// </summary>
    /// <param name="document">Текущий документ</param>
    /// <param name="category">Проверяемая категория</param>
    /// <param name="guid">Guid общего парметра для проверки</param>
    /// <returns>True, если параметр имеется, в противном случае - false</returns>
    bool HasCategoryParameterBinding(Document document, BuiltInCategory category, Guid guid) {

      try {
        BindingMap bindingMap = document.ParameterBindings; //Получение карты привязки текущего документа
        var it = bindingMap.ForwardIterator();
        it.Reset();
        while (it.MoveNext()) {
          var definition = (InternalDefinition)it.Key;
          var sharedParameter = document.GetElement(definition.Id) as SharedParameterElement;
          if (sharedParameter != null) {
            if (sharedParameter.GuidValue.Equals(guid)) {
              var inst_bind = it.Current as InstanceBinding;
              return inst_bind.Categories.Contains(Category.GetCategory(document, category));
            }
          }
        }
      }
      catch (Exception) {
      }
      return false;
    }

    /// <summary>
    /// Поиск определения по GUID в ФОП
    /// </summary>
    /// <param name="document">Текущий документ</param>
    /// <param name="guid">GUID параметра</param>
    /// <returns>null если определение не найдено</returns>
    Definition FindParameterInFile(Document document, Guid guid) {

      DefinitionFile fileSharedParameter = document.Application.OpenSharedParameterFile();
      Definition definition = null;

      try {
        foreach (DefinitionGroup definitionGroup in fileSharedParameter.Groups.Reverse()) {
          foreach (Definition def in definitionGroup.Definitions) {
            if (def != null) {
              ExternalDefinition guidDef = def as ExternalDefinition;
              if (guidDef.GUID.Equals(guid)) {
                definition = def;
                break;
              }
            }
          }
        }
      }
      catch (Exception) {
      }
      return definition;
    }

    /// <summary>
    /// Список кабелей
    /// </summary>
    List<MyCable> m_cables = new List<MyCable>();

    /// <summary>
    /// Обязательные категории
    /// </summary>
    List<BuiltInCategory> m_categories = new List<BuiltInCategory>() {
      BuiltInCategory.OST_Conduit,
      BuiltInCategory.OST_PipeCurves,
      BuiltInCategory.OST_DuctCurves,
      BuiltInCategory.OST_DuctFitting,
      BuiltInCategory.OST_PipeFitting,
      BuiltInCategory.OST_ElectricalEquipment,
      BuiltInCategory.OST_ConduitFitting
    };

    /// <summary>
    /// Обязательные параметры
    /// </summary>
    Dictionary<string, Guid> m_params = new Dictionary<string, Guid>() {
      {"BD_Состав кабельной продукции", new Guid("f08f11b0-abe7-4ef0-bbaf-b80fd9243814")},
      {"BD_Марка кабеля", new Guid("049d1803-85a6-4dee-be4b-fe2eb7e5700f")},
      {"BD_Обозначение кабеля", new Guid("8e952e6b-3e8b-46f0-80c2-992ed0acd387")},
      {"BD_Длина вручную", new Guid("4955cbe3-6068-46d6-a588-df76ac45e30e")},
      {"BD_Длина кабеля", new Guid("1d8966b3-d27c-4358-a95a-ad32dd46dc63")},
      {"ADSK_Количество", new Guid("8d057bb3-6ccd-4655-9165-55526691fe3a")}
    };

    Document m_document;
  }
}