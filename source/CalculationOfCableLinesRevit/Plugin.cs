using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;

using CalculationCable.Properties;

namespace CalculationCable;

public class Plugin : IExternalApplication {

  public static Document ActiveDocument { get; set; }
  public static UIApplication UIApplication { get; set; }
  public static UIControlledApplication UIControlledApplication { get; set; }

  public Result OnStartup(UIControlledApplication application) {

    UIControlledApplication = application;

    UIControlledApplication.ViewActivated += Application_ViewActivated;
    UIControlledApplication.ApplicationClosing += Application_ApplicationClosing;
    UIControlledApplication.Idling += Application_Idling;
    UIControlledApplication.ControlledApplication.ApplicationInitialized += ControlledApplication_ApplicationInitialized;
    UIControlledApplication.ControlledApplication.DocumentClosing += ControlledApplication_DocumentClosing;

//#if !RTV2020
//  string tabPanelName = "BIMDATA";
//  try {
//    application.GetRibbonPanels(tabPanelName);
//  }
//  catch {
//    application.CreateRibbonTab(tabPanelName);
//  }

//  Autodesk.Revit.UI.RibbonPanel ribbonPanel = application.GetRibbonPanels(tabPanelName).FirstOrDefault(x => x.Name == tabPanelName) ??
//    application.CreateRibbonPanel(tabPanelName, tabPanelName);
//  ribbonPanel.Name = tabPanelName;
//  ribbonPanel.Title = tabPanelName;
//  AddSplitButtonGroup(ribbonPanel);

//#endif
    return Result.Succeeded;
  }

  private void ControlledApplication_ApplicationInitialized(object sender, Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs e) {

    string _tabName = "BIMDATA";

    _ribbon = new Ribbon(UIControlledApplication);

    if (!_ribbon.FindTab(_tabName, out var tab)) {
      tab = _ribbon.CreateTab(_tabName);
    }
    if (!tab.FindPanel(_tabName, out var panel)) {
      tab.CreatePanel(_tabName);
    }
    panel = tab.CreatePanel("CalculationCable");
    
    string version = String.Format("\nversion: {0}", Assembly.GetExecutingAssembly().GetName().Version.ToString(2));

    panel.CreatePushButton<CommandUpdateCabel>("Cabel copy", bt => bt
      .SetLargeImage(Resources.BD_Copy)
      .SetSmallImage(Resources.BD_Copy)
      .SetDescription(version));

    panel.CreatePushButton<CommandUpdateCabel>("Update cabels", bt => bt
      .SetLargeImage(Resources.BD_Parameter)
      .SetSmallImage(Resources.BD_Parameter)
      .SetDescription(version));
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

  void Application_Idling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e) {

    BDCableSet bDCables = BDCableMagazine.GetInstance().ActiveBDCableSet;
    if (bDCables != null) {
      bDCables.Update();
    }

    UIControlledApplication.Idling -= Application_Idling;
  }

  void Application_ApplicationClosing(object sender, Autodesk.Revit.UI.Events.ApplicationClosingEventArgs e) { }

  void ControlledApplication_DocumentClosing(object sender, Autodesk.Revit.DB.Events.DocumentClosingEventArgs e) {
    if (BDCableMagazine.GetInstance().Count != 0) {
      BDCableMagazine.GetInstance().Remove(e.Document);
    }
    if (UIApplication.Application.Documents.Size == 1) {
#if !RTV2020
      _ribbon.Reset();
#endif
    }
  }

//#if RTV2020
//  void AddSplitButtonGroup(Autodesk.Revit.UI.RibbonPanel panel) {
//    SplitButtonData group1Data = new SplitButtonData("Cabel", "Split CableType Cabel\nВерсия: {curver}");
//    SplitButton splitButton = panel.AddItem(group1Data) as SplitButton;
//    splitButton.IsSynchronizedWithCurrentItem = false;
//    AddButtonInGroup(splitButton, "BD_Copy.png", "Cabel copy", "Copying cable", typeof(CommandCopyCabel));
//    //AddButtonInGroup(splitButton, "BD_Length.png", "Cabel length", "Calculation of cable length", typeof(CommandLengthCabel));
//    //AddButtonInGroup(splitButton, "BD_Fill.png", "Cabel fill parameters", "Fill of cable parameters", typeof(CommandFiilParametersCabel));
//    AddButtonInGroup(splitButton, "BD_Copy.png", "Update cabels", "Update cabels", typeof(CommandUpdateCabel));
//  }

//  PushButton AddButtonInGroup(SplitButton group, string icoName, string battonName, string buttonDescription, Type commandType) {
//    var cabelDataFill = new PushButtonData(battonName, buttonDescription, Assembly.GetExecutingAssembly().Location, commandType.FullName);
//    var button = group.AddPushButton(cabelDataFill) as PushButton;

//#if DEBUG
//    string version = String.Format("\nversion: {0} DEBUG", Assembly.GetExecutingAssembly().GetName().Version.ToString(3));
//#else
//  string version = String.Format("\nversion: {0}", Assembly.GetExecutingAssembly().GetName().Version.ToString(2));
//#endif

//    button.ToolTip = cabelDataFill.Text + version;   // Can be changed to a more descriptive text.
//#if RTV2020
//  button.Image = new BitmapImage(new Uri($"pack://application:,,,/BDCalculationCable2020;component/Resources/{icoName}"));
//#endif
//#if RTV2021
//    button.Image = new BitmapImage(new Uri($"pack://application:,,,/BDCalculationCable2021;component/Resources/{icoName}"));
//#endif
//#if RTV2022
//  button.Image = new BitmapImage(new Uri($"pack://application:,,,/BDCalculationCable2022;component/Resources/{icoName}"));
//#endif
//#if RTV2023
//    button.Image = new BitmapImage(new Uri($"pack://application:,,,/BDCalculationCable2023;component/Resources/{icoName}"));
//#endif
//    button.LargeImage = button.Image;
//    return button;
//  }
//#endif

  CalculationCable.Ribbon _ribbon;
}
