using System;
using System.Linq;

using Autodesk.Revit.DB;

namespace CalculationCable;

public static partial class Global {

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
    if (fileSharedParameter == null) {
      return null;
    }

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
}
