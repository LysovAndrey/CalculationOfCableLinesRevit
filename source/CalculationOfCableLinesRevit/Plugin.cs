using System;
using System.Reflection;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using CalculationCable.Properties;

namespace CalculationCable;

public class Plugin : IExternalApplication {

  public static Document ActiveDocument { get; set; }
  public static UIApplication UIApplication { get; set; }
  public static UIControlledApplication UIControlledApplication { get; set; }

  public Result OnStartup(UIControlledApplication application) {

#if DEBUG
    string version = String.Format("\nversion: {0}", Assembly.GetExecutingAssembly().GetName().Version.ToString(3));
#else
    string version = String.Format("\nversion: {0}", Assembly.GetExecutingAssembly().GetName().Version.ToString(2));
#endif

    UIControlledApplication = application;

    string _tabName = "BIMDATA";

    _ribbon = new Ribbon(UIControlledApplication);

    if (!_ribbon.FindTab(_tabName, out var tab)) {
      tab = _ribbon.CreateTab(_tabName);
    }

    var panel = tab.CreatePanel("CalculationCable");
    panel.CreatePushButton<CommandCopyCabel>("Cabel copy", bt => bt
      .SetLargeImage(Resources.BD_Copy)
      .SetSmallImage(Resources.BD_Copy)
      .SetDescription(version));
    panel.CreatePushButton<CommandUpdateCabel>("Update cabels", bt => bt
      .SetLargeImage(Resources.BD_Parameter)
      .SetSmallImage(Resources.BD_Parameter)
      .SetDescription(version));

    UIControlledApplication.ViewActivated += Application_ViewActivated;
    UIControlledApplication.ControlledApplication.DocumentClosing += ControlledApplication_DocumentClosing;

#if DEBUG
    UIControlledApplication.Idling += Application_Idling;
#endif

    return Result.Succeeded;
  }

  public Result OnShutdown(UIControlledApplication application) {
    return Result.Succeeded;
  }

  void Application_ViewActivated(object sender, Autodesk.Revit.UI.Events.ViewActivatedEventArgs e) {
    UIApplication = sender as UIApplication;
    ActiveDocument = e.Document;
    _ribbon.Update(e);
    BDCableMagazine.GetInstance().Update(ActiveDocument);
  }

#if DEBUG
  void Application_Idling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e) {

    //BDCableSet bDCables = BDCableMagazine.GetInstance().ActiveBDCableSet;
    //if (bDCables != null) {
    //  bDCables.Update();
    //}
    //UIControlledApplication.Idling -= Application_Idling;
  }
#endif

  void ControlledApplication_DocumentClosing(object sender, Autodesk.Revit.DB.Events.DocumentClosingEventArgs e) {

    if (BDCableMagazine.GetInstance().Count != 0) {
      BDCableMagazine.GetInstance().Remove(e.Document);
    }
    if (UIApplication.Application.Documents.Size == 1) {
      _ribbon.Reset();
    }
  }

  Ribbon _ribbon;
}
