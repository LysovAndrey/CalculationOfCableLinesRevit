using Autodesk.Revit.UI;

using System;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace CalculationCable {

  public class Start : IExternalApplication {

    internal static Start _app = null;
    public static Start Instance => _app;

    public static bool FittingCalc = false;
    public static bool ParameterCheck = false;

    public Result OnStartup(UIControlledApplication application) {
      _app = this;
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

    private void AddSplitButtonGroup(RibbonPanel panel) {
      SplitButtonData group1Data = new SplitButtonData("Cabel", "Split Group Cabel");
      SplitButton splitButton = panel.AddItem(group1Data) as SplitButton;
      splitButton.IsSynchronizedWithCurrentItem = false;
      AddButtonInGroup(splitButton, "BD_Copy.png", "Cabel copy", "Copying cable", typeof(CommandCopyCabel));
      AddButtonInGroup(splitButton, "BD_Length.png", "Cabel length", "Calculation of cable length", typeof(CommandLengthCabel));
      AddButtonInGroup(splitButton, "BD_Fill.png", "Cabel fill parameters", "Fill of cable parameters", typeof(CommandFiilParametersCabel));
    }

    private PushButton AddButtonInGroup(SplitButton group, string icoName, string battonName, string buttonDescription, Type commandType) {
      var cabelDataFill = new PushButtonData(battonName, buttonDescription, Assembly.GetExecutingAssembly().Location, commandType.FullName);
      var button = group.AddPushButton(cabelDataFill) as PushButton;
      button.ToolTip = cabelDataFill.Text;  // Can be changed to a more descriptive text.
      button.Image = new BitmapImage(new Uri($"pack://application:,,,/CalculationCable;component/Resources/{icoName}"));
      // pack://application:,,,/assemblyname;component
      button.LargeImage = button.Image;
      return button;
    }
  }
}
