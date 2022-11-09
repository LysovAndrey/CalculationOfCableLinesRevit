using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace CalculationCable {

  public class Start : IExternalApplication {

    public static Document ActiveDocument { get; set; }
    public static UIApplication UIApplication => m_ui_app;

    public Result OnStartup(UIControlledApplication application) {

      application.ViewActivated += Application_ViewActivated;
      application.ApplicationClosing += Application_ApplicationClosing;
      application.Idling += Application_Idling;
      application.ControlledApplication.DocumentClosing += ControlledApplication_DocumentClosing;

      string tabPanelName = "BIMDATA";
      try {
        application.GetRibbonPanels(tabPanelName);
      }
      catch {
        application.CreateRibbonTab(tabPanelName);
      }

      var ribbonPanel = application.GetRibbonPanels(tabPanelName).FirstOrDefault(x => x.Name == tabPanelName) ??
        application.CreateRibbonPanel(tabPanelName, tabPanelName);
      ribbonPanel.Name = tabPanelName;
      ribbonPanel.Title = tabPanelName;

      AddSplitButtonGroup(ribbonPanel);
      return Result.Succeeded;
    }

    public Result OnShutdown(UIControlledApplication application) {
      return Result.Succeeded;
    }

    void Application_ViewActivated(object sender, Autodesk.Revit.UI.Events.ViewActivatedEventArgs e) {
      ActiveDocument = e.Document;
      Debug.WriteLine($"Application_ViewActivated - {ActiveDocument.Title}");
      BDCableMagazine.GetInstance().Update(ActiveDocument);
    }

    void Application_Idling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e) {

      var str = ActiveDocument?.Title == null ? "Проекта нет" : ActiveDocument.Title;
      Debug.WriteLine($"\nApplication_Idling - {str}");
      if (BDCableMagazine.GetInstance().ActiveBDCableSet != null) {
        // Обновление BDCableSet активного документа
        Debug.WriteLine($"\tКоличество документов - {BDCableMagazine.GetInstance().Count}");
        BDCableMagazine.GetInstance().ActiveBDCableSet.Update();
      } else {
        // Набора BDCableSet в документе нет
        Debug.WriteLine($"Базы нет");
      }
    }

    void Application_ApplicationClosing(object sender, Autodesk.Revit.UI.Events.ApplicationClosingEventArgs e) {
      Debug.WriteLine("Application_ViewActivated");
    }

    void ControlledApplication_DocumentClosing(object sender, Autodesk.Revit.DB.Events.DocumentClosingEventArgs e) {
      if (BDCableMagazine.GetInstance().Count != 0) {
        BDCableMagazine.GetInstance().Remove(e.Document);
      }
    }

    void AddSplitButtonGroup(RibbonPanel panel) {
      SplitButtonData group1Data = new SplitButtonData("Cabel", "Split Group Cabel\nВерсия: {curver}");
      SplitButton splitButton = panel.AddItem(group1Data) as SplitButton;
      splitButton.IsSynchronizedWithCurrentItem = false;
      AddButtonInGroup(splitButton, "BD_Copy.png", "Cabel copy", "Copying cable", typeof(CommandCopyCabel));
      AddButtonInGroup(splitButton, "BD_Length.png", "Cabel length", "Calculation of cable length", typeof(CommandLengthCabel));
      AddButtonInGroup(splitButton, "BD_Fill.png", "Cabel fill parameters", "Fill of cable parameters", typeof(CommandFiilParametersCabel));
#if DEBUG
      AddButtonInGroup(splitButton, "BD_Copy.png", "Cabel to buffer", "Cabel to buffer", typeof(CopyToBufferBD_СableСomposition));
#endif
    }

    PushButton AddButtonInGroup(SplitButton group, string icoName, string battonName, string buttonDescription, Type commandType) {
      var cabelDataFill = new PushButtonData(battonName, buttonDescription, Assembly.GetExecutingAssembly().Location, commandType.FullName);
      var button = group.AddPushButton(cabelDataFill) as PushButton;
      button.ToolTip = cabelDataFill.Text +
        String.Format("\nversion: {0}", Assembly.GetExecutingAssembly().GetName().Version.ToString(2)); ;  // Can be changed to a more descriptive text.
      button.Image = new BitmapImage(new Uri($"pack://application:,,,/CalculationCable;component/Resources/{icoName}"));
      button.LargeImage = button.Image;
      return button;
    }

    static UIApplication m_ui_app = null;
  }
}
