using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Globalization;

namespace CalculationCable {

  internal enum BDStatus {
    None = 5000,     /// не определенно
    Normal = 5100,   /// Нормально
    Error = -5001,   /// Ошибка
    Cancel = -5002,  /// Отмена 
    Failed = -5003   /// Неудача
  }

  internal interface IBDCable {
    public Element Element { get; }   /// Элемент Revit
    public Category Category { get; } /// Категория Revit
    public string Source { get; }     /// BD_Состав кабельной продукции
    public string Group { get; }      /// BD_Обозначение кабеля
    public string CableType { get; }  /// BD_Марка кабеля
    public double Length { get; }     /// BD_Длина кабеля
    public double Quantity { get; }   /// ADSK_Количество
  }

  internal class BDCableSet : HashSet<IBDCable> {

    public new int Count => base.Count;
    public Document ActiveDocument { get; set; }

    public BDCableSet() {

      ActiveDocument = Start.ActiveDocument;

      if (HasParameters() == BDStatus.Normal) {
        if (BDCableMagazine.GetInstance().TryGetValue(ActiveDocument, out var cables)) {
          Update(ref cables);
        }
        else {
          BDCableMagazine.GetInstance().Add(ActiveDocument, this);
        }
      }
    }

    /// <summary>
    /// Обновление текущего набора
    /// </summary>
    public void Update() {
      if (BDCableMagazine.GetInstance().TryGetValue(ActiveDocument, out var cables)) {
        Update(ref cables);
      }
    }

    /// <summary>
    /// Получение элементов для копирования
    /// </summary>
    public List<IBDCable> GetElementsToCopy() {

      List<IBDCable> list = null;

      List<IBDCable> tmp = null;

      if (BDCableMagazine.GetInstance().TryGetValue(ActiveDocument, out var cables)) {
        tmp = cables.Where(c => {
          var cable = c as BDCableImpl;
          return cable.State == BDState.Copy;
        }).ToList();

        if (tmp.Count != 0) {

          foreach (var item in tmp) {
            var compositionCabel = item.Source.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            list = new List<IBDCable>();
            for (var i = 0; i < compositionCabel.Length; i++) {
              var c = new BDCableImpl();
              c.Element = item.Element;
              c.Category = item.Category;
              c.Source = compositionCabel[i];
              c.State = BDState.Ok;
              list.Add(c);
            }
          }
        }
      }

      return list;
    }

    /// <summary>
    /// Получение элементов с ошибками заполнения параметра "BD_Состав кабельной продукции"
    /// </summary>
    public List<IBDCable> GetElementsToErrorParsing() {

      List<IBDCable> list = null;

      if (BDCableMagazine.GetInstance().TryGetValue(ActiveDocument, out var cables)) {
        list = cables.Where(c => {
          var cable = c as BDCableImpl;
          return cable.State == BDState.Parsing;
        }).ToList();
      }

      return list.Count == 0 ? null : list; 
    }

    /// <summary>
    /// Получение обновленых элементов
    /// </summary>
    public List<IBDCable> GetUpdatedElements() {

      List<IBDCable> list = null;

      if (BDCableMagazine.GetInstance().TryGetValue(ActiveDocument, out var cables)) {
        list = cables.Where(c => {
          var cable = c as BDCableImpl;
          return (cable.HasUpdate == true && cable.State == BDState.Ok);
        }).ToList();
      }

      return list.Count == 0 ? null : list;
    }

    BDStatus Update(ref BDCableSet cables) {

      // Проверить производительность при ссылке
      BDStatus status = BDStatus.Failed;

      var elementSet = SelectElement();

      if (elementSet != null && elementSet.Count != 0) {

        HashSet<IBDCable> cur = new HashSet<IBDCable>(elementSet.Count);

        foreach (var item in elementSet) {
          BDCableImpl cable = new BDCableImpl(item);
          if (cables.TryGetValue(cable, out var actual_cable)) {
            var update_cable = actual_cable as BDCableImpl;
            update_cable.Update(cable);
            cur.Add(update_cable);
          }
          else {
            cur.Add(cable);
          }
        }
        cables.Clear();
        cables.UnionWith(cur);
        cables.TrimExcess();
        this.UnionWith(cables);
        status = BDStatus.Normal;
      }
      else {
        cables.Clear();
        cables.TrimExcess();
        BDCableMagazine.GetInstance().Remove(ActiveDocument);
        status = BDStatus.None;
      }
      return status;
    }


    /// <summary>
    /// Выбор элементов в модели
    /// </summary>
    List<Element> SelectElement() {

      List<Element> elementSet = null;

      m_categories = Global.Categories as List<BuiltInCategory>;
      m_categories.RemoveAll(c => new FilteredElementCollector(ActiveDocument)
                                        .OfCategory(c)
                                        .WhereElementIsNotElementType()
                                        .FirstElement() == null);

      if (m_categories.Count != 0) {

        IList<ElementFilter> categorySet = new List<ElementFilter>(m_categories.Count());
        m_categories.ForEach(c => categorySet.Add(new ElementCategoryFilter(c)));
        var catFilter = new LogicalOrFilter(categorySet);

        /// Фильтр заполненого параметра
        var element = new FilteredElementCollector(ActiveDocument)
          .WherePasses(catFilter)
          .WhereElementIsNotElementType()
          .FirstElement();
        var parametr = element.get_Parameter(Global.Parameters["BD_Состав кабельной продукции"]);
        IList<FilterRule> rules = new List<FilterRule>();
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Состав кабельной продукции"));
        rules.Add(ParameterFilterRuleFactory.CreateContainsRule(parametr.Id, ";", false));
        var parametr_filter = new ElementParameterFilter(rules);

        /// Фильтры выбора элементов
        var elements = new FilteredElementCollector(ActiveDocument)
          .WherePasses(catFilter)
          .WherePasses(parametr_filter)
          .WhereElementIsCurveDriven();
        if (elements.GetElementCount() != 0) {
          elementSet = elements.ToList();
        }

        categorySet.Clear();

        m_categories.Remove(BuiltInCategory.OST_ElectricalEquipment);

        if (m_categories.Count != 0) {
          m_categories.ForEach(c => categorySet.Add(new ElementCategoryFilter(c)));
          catFilter = new LogicalOrFilter(categorySet);

          elements = new FilteredElementCollector(ActiveDocument)
            .OfClass(typeof(FamilyInstance))
            .WherePasses(catFilter)
            .WherePasses(parametr_filter)
            .WhereElementIsNotElementType();
          if (elements.GetElementCount() != 0 && elementSet != null) {
            elementSet.AddRange(elements.ToList());
          }
          else if (elements.GetElementCount() != 0 && elementSet == null) {
            elementSet = elements.ToList();
          }
        }
      }

      return elementSet;
    }



    /// <summary>
    /// Проверка обязательных параметров в модели.
    /// </summary>
    BDStatus HasParameters() {

      BDStatus status = BDStatus.Failed;

      if (m_categories.Count == 0) {
        status = BDStatus.Cancel;
      }
      else {
        // string наименование параметра
        // Если List<Category> = null параметра в модели нет
        Dictionary<string, List<Category>> bind_params = new Dictionary<string, List<Category>>();

        foreach (var categoryId in m_categories.Reverse<BuiltInCategory>()) {
          List<Category> categories = null;
          foreach (var param in Global.Parameters) {
            Definition definition_ = Global.FindParameterInProject(ActiveDocument, param.Value);
            if (definition_ == null) {
              if (!bind_params.TryGetValue(param.Key, out categories)) {
                bind_params.Add(param.Key, null);
              }
              continue;
            }
            Category category = Category.GetCategory(ActiveDocument, categoryId);
            ElementBinding elem_bind = ActiveDocument.ParameterBindings.get_Item(definition_) as ElementBinding;
            if (elem_bind.Categories.Contains(category)) {
              // Включаем категорию "Электрооборудование" если семейство выпонено на основе линии
              Element element = new FilteredElementCollector(ActiveDocument).OfCategory(categoryId).WhereElementIsCurveDriven().FirstElement();
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
          status = BDStatus.Normal;
        }
        else {
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

          if (res != TaskDialogResult.CommandLink1) {
            status = BDStatus.Cancel;
          }
          else {
            Dictionary<Definition, CategorySet> add_params = new Dictionary<Definition, CategorySet>();
            CategorySet add_categories = ActiveDocument.Application.Create.NewCategorySet();

            // Привязка параметров из модели
            foreach (var item in bind_params.Reverse()) {
              if (item.Value != null) {
                Definition definition_ = Global.FindParameterInProject(ActiveDocument, Global.Parameters[item.Key]);
                ElementBinding element_bind = ActiveDocument.ParameterBindings.get_Item(definition_) as ElementBinding;
                element_bind.Categories.Cast<Category>().ToList().ForEach(c => add_categories.Insert(c));
                item.Value.ForEach(c => add_categories.Insert(c));
                add_params.Add(definition_, add_categories);
                bind_params.Remove(item.Key);
              }
            }

            Binding typeBinding = ActiveDocument.Application.Create.NewInstanceBinding(add_categories);

            using (Transaction tr = new Transaction(ActiveDocument)) {
              tr.Start("Добавление параметров - CalculationCable");
              foreach (var item in add_params) {
                ActiveDocument.ParameterBindings.ReInsert(item.Key, typeBinding);
              }
              tr.Commit();
            }

            if (bind_params.Count == 0) {
              status = BDStatus.Normal;
            }
            else {
              add_categories.Clear();
              add_params.Clear();

              m_categories.ForEach(c => add_categories.Insert(Category.GetCategory(ActiveDocument, c)));

              // Добавление пармаетров из ФОП
              foreach (var item in bind_params.Reverse()) {
                if (item.Value == null) {
                  Definition definition_ = Global.FindParameterInFile(ActiveDocument, Global.Parameters[item.Key]);
                  if (definition_ == null) {
                    continue;
                  }
                  add_params.Add(definition_, add_categories);
                  bind_params.Remove(item.Key);
                }
              }

              using (Transaction tr = new Transaction(ActiveDocument)) {
                tr.Start("Добавление параметров - CalculationCable");
                foreach (var item in add_params) {
                  ActiveDocument.ParameterBindings.Insert(item.Key, typeBinding);
                }
                tr.Commit();
              }

              if (bind_params.Count != 0) {
                string file_name = ActiveDocument.Application.SharedParametersFilename;
                var error_message = "";
                bind_params.ToList().ForEach(s => error_message += "\nПараметр: " + s.Key + "\nGuid: " + Global.Parameters[s.Key].ToString() + "\n");
                TaskDialog dialog_file = new TaskDialog("Ошибка.");
                dialog_file.MainInstruction = $"В файле общих параметров \"{file_name}\" отсутствуют параметры:\n{error_message}";
                dialog_file.CommonButtons = TaskDialogCommonButtons.Close;
                dialog_file.DefaultButton = TaskDialogResult.Close;
                dialog_file.Show();
                status = BDStatus.Error;
              }
              else {
                status = BDStatus.Normal;
              }
            }
          }
        }
      }
      return status;
    }

    enum BDState {
      None,   /// Не определено
      Ok,     /// Элемент для обновления
      Copy,   /// Элемент для копирования
      Parsing /// Элемент с неправильным заполнением параметра
    }

    class BDCableImpl : IBDCable {

      public Element Element { get; set; }        /// Элемент Revit
      public Category Category { get; set; }      /// Категория Revit
      public string Source { get; set; }          /// BD_Состав кабельной продукции
      public string Group { get; set; }           /// BD_Обозначение кабеля
      public string CableType { get; set; }       /// BD_Марка кабеля
      public double Length { get; set; }          /// BD_Длина кабеля
      public double Quantity { get; set; }        /// ADSK_Количество
      public bool IsManuallyLength { get; set; }  /// Длина в ручную
      public bool HasUpdate { get; set; }         /// Элемент для обновления
      public BDState State { get; set; }          /// Статус элемента

      internal BDCableImpl() {
      }

      internal BDCableImpl(Element element) {
        State = BDState.None;
        Element = element;
        Category = Element.Category;
        Source = Element.get_Parameter(Global.Parameters["BD_Состав кабельной продукции"]).AsString();
        Group = "";
        CableType = "";
        Quantity = 0;
        IsManuallyLength = Element.get_Parameter(Global.Parameters["BD_Длина вручную"]).AsInteger() != 0;
        this.HasUpdate = true;

        if (!HasCopy(Source)) {
          State = BDState.Copy;
          this.HasUpdate = false;
        }
        else if (SupportParsingBDParameter(Source, out string pattern)) {
          ParsingParametr(pattern);
          State = BDState.Ok;
        }
        else {
          State = BDState.Parsing;
          this.HasUpdate = false;
        }
      }

      public override bool Equals(object obj) {

        if (obj is BDCableImpl && obj == null) { return false; }
        if (this.GetHashCode() == obj.GetHashCode()) { return true; }
        else {
          return false;
        }
      }

      public override int GetHashCode() {
        return 1515200279 + Element.Id.IntegerValue.GetHashCode();
      }


      /// <summary>
      /// Обновление кабеля
      /// </summary>
      public void Update(BDCableImpl c) {

        if (c.State == BDState.Ok && this.State == c.State) {
          
          var group = c.Element.get_Parameter(Global.Parameters["BD_Марка кабеля"]).AsString();
          var cableType = c.Element.get_Parameter(Global.Parameters["BD_Обозначение кабеля"]).AsString();
          var len = c.Element.get_Parameter(Global.Parameters["BD_Длина кабеля"]).AsDouble();
          var qt = c.Element.get_Parameter(Global.Parameters["ADSK_Количество"]).AsDouble();

          if (!this.Group.Equals(c.Group) || !this.Group.Equals(group)) { this.HasUpdate = true; }
          else if (!this.CableType.Equals(c.CableType) || !this.CableType.Equals(cableType)) { this.HasUpdate = true; }
          else if (!this.Length.Equals(c.Length) || !this.Length.Equals(len)) { this.HasUpdate = true; }
          else if (!this.Quantity.Equals(c.Quantity) || !this.Quantity.Equals(qt)) { this.HasUpdate = true; }
          else { this.HasUpdate = false; }
        }
        else {
          this.HasUpdate = true;
        }

        if (this.HasUpdate) {
          this.Source = c.Source;
          this.Group = c.Group;
          this.CableType = c.CableType;
          this.Length = c.Length;
          this.Quantity = c.Quantity;
          this.State = c.State;
        }
      }

      /// <summary>
      /// Парсинг параметра
      /// </summary>
      void ParsingParametr(string pattern) {

        if (Regex.IsMatch(Source, pattern, RegexOptions.IgnoreCase)) {
          foreach (Match match in Regex.Matches(Source, pattern, RegexOptions.IgnoreCase)) {
            Group = match.Groups["gp"].Value;
            CableType = match.Groups["mr"].Value;

            switch ((BuiltInCategory)Category.Id.IntegerValue) {
              case BuiltInCategory.OST_DuctFitting:
              case BuiltInCategory.OST_PipeFitting:
              case BuiltInCategory.OST_ConduitFitting:
                Length = Element.get_Parameter(Global.Parameters["ADSK_Количество"]).AsDouble();
                Length = Length != 0 ? Length : 0.05;
                Quantity = Length;
                Length = Global.ConvertToInternaMeters(Length);
                break;
              default:
                if (match.Groups["ln"].Value.Length != 0) {
                  // Из BD_Состав кабельной продукции
                  var str_value = match.Groups["ln"].Value;
                  double result = Double.NaN;
                  if (!double.TryParse(str_value, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out result)) {
                    if (!double.TryParse(str_value, NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out result)) {
                      State = BDState.Parsing;
                    }
                  }

                  Length = Global.ConvertToInternaMeters(result);
                  Quantity = result;
                }
                else if (match.Groups["ln"].Value.Length == 0 && IsManuallyLength) {
                  // Из BD_Длина кабеля
                  Length = Element.get_Parameter(Global.Parameters["BD_Длина кабеля"]).AsDouble();
                  Quantity = Global.ConvertFromInternaMeters(Length);
                }
                else {
                  Length = Element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                  Quantity = Global.ConvertFromInternaMeters(Length);
                }
                break;
            }
          }
        }
      }

      /// <summary>
      /// Проверка парсинга параметра BD_Состав кабельной продукции
      /// </summary>
      /// <param name="value">Значение параметра BD_Состав кабельной продукции</param>
      /// <returns>true, парсинг возможен, иначе false</returns>
      bool SupportParsingBDParameter(string source, out string pattern) {

        bool status = false;

        Regex regex = new Regex(@";");
        string[] substrings = regex.Split(source);
        pattern = null;
        switch (substrings.Length) {
          case 2:
            // <gp>    <mr>   
            // Гр1; Тип кабеля
            pattern = @"(?<gp>(.+));(\s+)?(?<mr>(.+))$";
            if (Regex.IsMatch(source, pattern, RegexOptions.IgnoreCase)) {
              status = true;
              break;
            }
            //    <mr>       <ln>
            // Тип кабеля; 200{{,.}0}
            pattern = @"(?<mr>(.+));(\s+)?(?<ln>((\d+[.,])?\d+))$";
            if (Regex.IsMatch(source, pattern, RegexOptions.IgnoreCase)) {
              status = true;
              break;
            }
            //// <gp>
            //// Гр1;
            //pattern = @"(?<gp>(.+));";
            //if (Regex.IsMatch(source, pattern, RegexOptions.IgnoreCase)) {
            //  status = true;
            //  break;
            //}
            break;
          case 3:
            // <gp>    <mr>       <ln>
            // Гр1; Тип кабеля; 200{{,.}0}
            pattern = @"(?<gp>(.+));(\s+)?(?<mr>(.+));(\s+)?(?<ln>((\d+[.,])?\d+))";
            if (Regex.IsMatch(source, pattern, RegexOptions.IgnoreCase)) {
              status = true;
            }
            break;
        }

        return status;
      }

      /// <summary>
      /// Проверка элемента для копирования
      /// </summary>
      /// <returns>true, копировать не нужно, иначе false</returns>
      bool HasCopy(string source) {
        return source.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length == 1 ? true : false;
      }
    }

    List<BuiltInCategory> m_categories = Global.Categories.ToList();
  }

  internal class BDCableMagazine : Dictionary<Document, BDCableSet> {

    public BDCableSet ActiveBDCableSet { get => _activeBDCableSet; }
    public new int Count => base.Count;

    public new BDStatus Add(Document document, BDCableSet cables) {

      BDStatus status = BDStatus.Failed;

      if (!ContainsKey(document)) {
        base.Add(document, cables);
        _activeBDCableSet = cables;
        cables.Update();
        status = BDStatus.Normal;
      }

      return status;
    }

    public new BDStatus Remove(Document document) {
      BDStatus status = BDStatus.Failed;
      if (base.Remove(document)) {
        status = BDStatus.Normal;
      }
      return status;
    }

    public void Update(Document document) {
      TryGetValue(document, out var bds);
      _activeBDCableSet = bds;
    }

    public static BDCableMagazine GetInstance() {
      if (_instance == null) {
        _instance = new BDCableMagazine();
      }
      return _instance;
    }

    BDCableMagazine() : base(new DocumentEqualityComparer()) { }

    static BDCableMagazine _instance;
    BDCableSet _activeBDCableSet;
  }

  internal static class Global {

    public static double ConvertFromInternaMeters(double length) {

#if RTV2020
      return UnitUtils.ConvertFromInternalUnits(length, DisplayUnitType.DUT_METERS);
#endif

#if RTV2022
      return -1;
#endif
    }

    public static double ConvertToInternaMeters(double length) {

#if RTV2020
      return UnitUtils.ConvertToInternalUnits(length, DisplayUnitType.DUT_METERS);
#endif

#if RTV2022
      return -1;
#endif
    }

    /// <summary>
    /// Поиск определения параметра в проекте по Guid
    /// </summary>
    /// <param name="document">Текущий документ</param>
    /// <param name="guid">Guid парметра</param>
    /// <returns>null если параметр не найден</returns>
    public static Definition FindParameterInProject(Document document, Guid guid) {
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
      catch (NullReferenceException) {
      }
      return definition;
    }

    /// <summary>
    /// Поиск определения параметра в в файле общих параметров по Guid
    /// </summary>
    /// <param name="document">Текущий документ</param>
    /// <param name="guid">Guid парметра</param>
    /// <returns>null если определение не найдено</returns>
    public static Definition FindParameterInFile(Document document, Guid guid) {

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
    /// Обязательные параметры
    /// </summary>
    public static IReadOnlyDictionary<string, Guid> Parameters =>
      new Dictionary<string, Guid>() {
        { "BD_Состав кабельной продукции", new Guid("f08f11b0-abe7-4ef0-bbaf-b80fd9243814")},
        { "BD_Марка кабеля", new Guid("049d1803-85a6-4dee-be4b-fe2eb7e5700f")},
        { "BD_Обозначение кабеля", new Guid("8e952e6b-3e8b-46f0-80c2-992ed0acd387")},
        { "BD_Длина вручную", new Guid("4955cbe3-6068-46d6-a588-df76ac45e30e")},
        { "BD_Длина кабеля", new Guid("1d8966b3-d27c-4358-a95a-ad32dd46dc63")},
        { "ADSK_Количество", new Guid("8d057bb3-6ccd-4655-9165-55526691fe3a")}
      };

    /// <summary>
    /// Обязательные категории
    /// </summary>
    public static IReadOnlyList<BuiltInCategory> Categories =>
      new List<BuiltInCategory>() {
        BuiltInCategory.OST_Conduit,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_ConduitFitting
      };
  }

  #region EqualityComparer

  internal class BDCableEqualityComparer : IEqualityComparer<BDCableBak> {

    public bool Equals(BDCableBak cable1, BDCableBak cable2) {

      if (cable1 == null && cable2 == null) { return true; }
      if (cable1 == null || cable2 == null) { return false; }
      if (cable1.Id.IntegerValue == cable2.Id.IntegerValue) { return true; }
      else { return false; }
    }

    public int GetHashCode(BDCableBak cable) {
      int hCode1 = cable.Id.IntegerValue;
      int hCode2 = cable.Source.GetHashCode();
      int hCode = cable.Id.IntegerValue.GetHashCode();
      return hCode;
    }
  }

  internal class DocumentEqualityComparer : IEqualityComparer<Document> {

    public bool Equals(Document document1, Document document2) {

      if (document1 == null && document2 == null) { return true; }
      if (document1 == null || document2 == null) { return false; }
      if (document1.GetHashCode() == document2.GetHashCode()) { return true; }
      else { return false; }
    }

    public int GetHashCode(Document document) { return document.GetHashCode(); }
  }

  #endregion

  #region Старая реализация

  internal class BDCableBak {
    public ElementId Id { get; set; }
    public string Source { get; set; } // Исходное значение параметра BD_Состав кабельно продукции
    public string Group { get; set; } // Номер группы
    public string CableType { get; set; } // Тип, марка кабеля
    public string Length { get; set; } // Длина участка кабеля

    public BDCableBak() { }

    public BDCableBak(Element element, string parsing) {
      Id = element.Id;
      Source = element.get_Parameter(Global.Parameters["BD_Состав кабельной продукции"]).AsString();
    }

    public void Update() {

    }

    public override bool Equals(object obj) {
      return obj is BDCableBak cable &&
             EqualityComparer<ElementId>.Default.Equals(Id, cable.Id);
    }

    public override int GetHashCode() {
      return 2108858624 + EqualityComparer<ElementId>.Default.GetHashCode(Id);
    }
  }

  internal class BDCableSetBak {

    public BDCableSetBak() {

      m_document = Start.ActiveDocument;

      m_categories.RemoveAll(c => new FilteredElementCollector(m_document)
      .OfCategory(c)
      .WhereElementIsNotElementType()
      .FirstElement() == null);

      m_status = HasParameters();
      if (m_status == BDStatus.Normal) {
        // Выбор всех элементов у которых заполнен парметр "BD_Состав кабельной продукции
        // Выбор элементов с параметром BD_Состав кабельной продукции
        IList<ElementFilter> categorySet = new List<ElementFilter>(m_categories.Count());
        m_categories.ForEach(c => categorySet.Add(new ElementCategoryFilter(c)));

        var catFilter = new LogicalOrFilter(categorySet);
        var element = new FilteredElementCollector(m_document).WhereElementIsNotElementType().WherePasses(catFilter).FirstElement();
        var param = element.get_Parameter(m_params["BD_Состав кабельной продукции"]);

        IList<FilterRule> rules = new List<FilterRule>();
        rules.Add(ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("BD_Состав кабельной продукции"));
        rules.Add(ParameterFilterRuleFactory.CreateContainsRule(param.Id, ";", false));

        var filterElement = new ElementParameterFilter(rules);
        var elementSet = new FilteredElementCollector(m_document).WhereElementIsNotElementType().WherePasses(catFilter).WherePasses(filterElement);

#if DEBUG
        Stopwatch stopwatch = new Stopwatch();
        Debug.WriteLine("\n----------------------------------------------");
        Debug.WriteLine($"Активный документ - {m_document.Title}");
        stopwatch.Start();
#endif
        Parallel.ForEach(elementSet, CreateBDCable);
#if DEBUG
        stopwatch.Stop();
        Debug.WriteLine($"Количество элементов: {elementSet.Count()}");
        Debug.WriteLine($"Количество тактов: {stopwatch.ElapsedTicks}");
        var resultTime = stopwatch.Elapsed;
        string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:000}",
          resultTime.Hours,
          resultTime.Minutes,
          resultTime.Seconds,
          resultTime.Milliseconds);
        Debug.WriteLine($"Время выполнения: {elapsedTime}");
        Debug.WriteLine("----------------------------------------------");
        Debug.WriteLine($"Количество элементов для копирования: {m_copy_elements.Count}");

        m_copy_elements.ForEach(e => Debug.WriteLine($"{e.Category.Name}"));
#endif
      }
      //BDCableMagazine.GetInstance().Add(m_document, this);
    }

    public BDCableSetBak(List<Element> elements) {
      m_document = elements.First().Document;
      foreach (var elem in elements) {
        Add(elem);
      }
    }

    public BDStatus Status { get; private set; }

    public CategorySet ToCategorySet() {
      CategorySet cats = null;
      if (m_categories.Count != 0) {
        cats = m_document.Application.Create.NewCategorySet();
        m_categories.ForEach(c => cats.Insert(Category.GetCategory(m_document, c)));
      }
      return cats;
    }

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
            //_cables.Add(new BDCable { CableType = value, CableType = "", Length = "", Id = elem.Id });
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

            m_cables.Add(new BDCableBak { Id = elem.Id, Source = value, Group = group, CableType = сableType, Length = length });
          }
        }
      }
      catch (Exception) { }

    }

    public void Update() {

    }

    public IEnumerator GetEnumerator() => m_cables.GetEnumerator();

    /// <summary>
    /// Для внутреннего использования.
    /// Создание BDCable
    /// </summary>
    void CreateBDCable(Element element, ParallelLoopState pls) {
      var value = element.get_Parameter(m_params["BD_Состав кабельной продукции"]).AsString();
      //Проверка на количество строк, должно быть 1
      var count_str = value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
      if (count_str == 1) {
        string parsing;
        // Проверка правильность заполнения параметра
        if (SupportParsingBDParameter(value, out parsing)) {
          //BDCable cable = new BDCable(element, parsing);
          //BDCable actualCable = null;
          //if (!m_cables_new.TryGetValue(cable, out actualCable)) {

          //  // Проверка изменений параметров
          //  // Проверка BD_Состав кабельной продукции
          //  // Проверка длины
          //};
        }
        else {
          m_bad_parsing.Add(element);
        }
      }
      else {
        m_copy_elements.Add(element);
      }
    }

    /// <summary>
    /// Возможность парсинга параметра BD_Состав кабельной продукции
    /// </summary>
    /// <param name="value">Значение параметра BD_Состав кабельной продукции</param>
    /// <param name="pattern">Шаблон регулярного выражения</param>
    /// <returns>true, парсинг возможен, иначе false</returns>
    bool SupportParsingBDParameter(string value, out string pattern) {

      bool status = false;

      Regex regex = new Regex(@";");
      string[] substrings = regex.Split(value);
      pattern = null;
      switch (substrings.Length) {
        case 2:
          //    <mr>       <ln>
          // Тип кабеля; 200{{,.}0}
          pattern = @"(?<mr>(.+));(\s+)?(?<ln>((\d+[.,])?\d+))$";
          if (Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase)) {
            status = true;
            break;
          }
          // <gp>    <mr>   
          // Гр1; Тип кабеля
          pattern = @"(?<gp>(.+));(\s+)?(?<mr>(.+))$";
          if (Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase)) {
            status = true;
            break;
          }
          // <gp>
          // Гр1;
          pattern = @"(?<gp>(.+));";
          if (Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase)) {
            status = true;
            break;
          }
          break;
        case 3:
          // <gp>    <mr>       <ln>
          // Гр1; Тип кабеля; 200{{,.}0}
          pattern = @"(?<gp>(.+));(\s+)?(?<mr>(.+));(\s+)?(?<ln>((\d+[.,])?\d+))";
          if (Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase)) {
            status = true;
          }
          break;
      }

      return status;
    }

    /// <summary>
    /// Проверка обязательных параметров в модели.
    /// </summary>
    BDStatus HasParameters() {

      BDStatus status = BDStatus.Failed;

      if (m_categories.Count == 0) {
        status = BDStatus.Cancel;
      }
      else {
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
          status = BDStatus.Normal;
        }
        else {
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

          if (res != TaskDialogResult.CommandLink1) {
            status = BDStatus.Cancel;
          }
          else {
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
              status = BDStatus.Normal;
            }
            else {
              add_categories.Clear();
              add_params.Clear();

              m_categories.ForEach(c => add_categories.Insert(Category.GetCategory(m_document, c)));

              // Добавление пармаетров из ФОП
              foreach (var item in bind_params.Reverse()) {
                if (item.Value == null) {
                  Definition definition_ = FindParameterInFile(m_document, m_params[item.Key]);
                  if (definition_ == null) {
                    continue;
                  }
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
              else {
                status = BDStatus.Normal;
              }
            }
          }
        }
      }
      return status;
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
      catch (NullReferenceException) {
      }
      return definition;
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
    HashSet<BDCableBak> m_cables_new = new HashSet<BDCableBak>(new BDCableEqualityComparer());

    /// <summary>
    /// Список кабелей
    /// </summary> 
    List<BDCableBak> m_cables = new List<BDCableBak>();

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

    /// <summary>
    /// Список элемент для коппирования, больше 1-ой строки 
    /// в параметре "BD_Состав кабельной продукции"
    /// </summary>
    List<Element> m_copy_elements = new List<Element>();

    /// <summary>
    /// Ошибочное заполнение параметра "BD_Состав кабельной продукции"
    /// </summary>
    List<Element> m_bad_parsing = new List<Element>();

    Document m_document;
    BDStatus m_status;
  }

  #endregion


}