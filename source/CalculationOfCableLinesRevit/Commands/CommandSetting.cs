using System;
using System.Linq;
using System.Diagnostics;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

using static CalculationCable.Options;


//Review: удалены неиспользуемые сборки  
namespace CalculationCable {

  [Transaction(TransactionMode.Manual)]
  internal class CommandSetting : IExternalCommand {
    internal static CommandSetting _app = null;
    public static CommandSetting Instance => _app;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {

      Debug.WriteLine("Комманда - \"Setting\"");

      var doc = commandData.Application.ActiveUIDocument.Document;
      //DialogSetting window = new DialogSetting();


      //window.ShowDialog();
      return Result.Succeeded;
    }

  }


}
