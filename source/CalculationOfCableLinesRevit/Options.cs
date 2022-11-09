using Autodesk.Revit.DB;

using System;
using System.Collections.Generic;

namespace CalculationCable {

  internal readonly struct Options {
    public static Dictionary<string, Guid> Parameter = new Dictionary<string, Guid>() {
        {"BD_Состав кабельной продукции", new Guid("f08f11b0-abe7-4ef0-bbaf-b80fd9243814")},
        {"BD_Марка кабеля", new Guid("049d1803-85a6-4dee-be4b-fe2eb7e5700f")},
        {"BD_Обозначение кабеля", new Guid("8e952e6b-3e8b-46f0-80c2-992ed0acd387")},
        {"BD_Длина вручную", new Guid("4955cbe3-6068-46d6-a588-df76ac45e30e")},
        {"BD_Длина кабеля", new Guid("1d8966b3-d27c-4358-a95a-ad32dd46dc63")}
      };

    internal const string PARAMETER_DESCRIPTION = "BD_Длина вручную"; // 4955cbe3-6068-46d6-a588-df76ac45e30e
    internal const string PARAMETER_NAME = "BD_Состав кабельной продукции"; // f08f11b0-abe7-4ef0-bbaf-b80fd9243814
    internal const string PARAMETER_LENGTH = "BD_Длина кабеля"; // 1d8966b3-d27c-4358-a95a-ad32dd46dc63
    internal const string PARAMETER_MARK = "BD_Марка кабеля"; // 049d1803-85a6-4dee-be4b-fe2eb7e5700f
    internal const string PARAMETER_GROUP = "BD_Обозначение кабеля"; // 8e952e6b-3e8b-46f0-80c2-992ed0acd387
    internal const string PARAMETER_ADSK_QUANTITY = "ADSK_Количество"; // 8d057bb3-6ccd-4655-9165-55526691fe3a
    internal const int LENGTH_STR = 3;

    internal readonly static List<BuiltInCategory> AllCategory = new List<BuiltInCategory>() {
          BuiltInCategory.OST_Conduit,
          BuiltInCategory.OST_PipeCurves,
          BuiltInCategory.OST_DuctCurves,
          BuiltInCategory.OST_DuctFitting,
          BuiltInCategory.OST_PipeFitting,
          BuiltInCategory.OST_ElectricalEquipment,   // Электрооборудование
          BuiltInCategory.OST_ConduitFitting
        };
  }
}
