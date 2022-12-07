using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace CalculationCable;

public class RibbonTab {

  public string Name { get => _tab.Name; set => _tab.Name = value; }
  public string Title { get => _tab.Title; set => _tab.Title = value; }
  public string Id { get => _tab.Id; set => _tab.Id = value; }
  public bool Visible { get => _tab.IsVisible; set => _tab.IsVisible = value; }

  public void Reset() {
    foreach (var item in _panels) {
      item.Value.Visible = true;
    }
  }

  public void Update(ViewActivatedEventArgs e) {
    foreach (var item in _panels) {
      item.Value.Update(e);
    }
  }

  public RibbonPanel CreatePanel(string name) => CreatePanel(name, RibbonPanel.VisibilityInDocument.None);

  public RibbonPanel CreatePanelForProject(string name) => CreatePanel(name, RibbonPanel.VisibilityInDocument.Project);

  public RibbonPanel CreatePanelForFamily(string name) => CreatePanel(name, RibbonPanel.VisibilityInDocument.Family); 

  RibbonPanel CreatePanel(string name, RibbonPanel.VisibilityInDocument document) {

    if (string.IsNullOrWhiteSpace(name)) {
      throw new ArgumentNullException("name");
    }

    if (FindPanel(name, out _)) {
      throw new ArgumentException(name, $"Панель {name} существует");
    }

    var panel = _ribbon.Application.CreateRibbonPanel(_tab.Name, name);
    RibbonPanel panel2 = new(panel);
    panel2.TypeDocument = document;
    _panels.Add(name, panel2); 
    return panel2; 
  }

  public bool FindPanel(string name, out RibbonPanel panel) {

    if (string.IsNullOrWhiteSpace(name)) {
      throw new ArgumentNullException("name");
    }

    if (_panels.TryGetValue(name, out panel)) {
      return true;
    }
    return false;
  }

  internal Autodesk.Windows.RibbonTab get_RibbonTab() => _tab;

  internal RibbonTab(Ribbon ribbon, Autodesk.Windows.RibbonTab tab) {

    _ribbon = ribbon;
    _tab = tab;

    var panels = _ribbon.Application.GetRibbonPanels(Name);
    if (panels.Count != 0) {
      foreach (var item in panels) {
        RibbonPanel panel = new(item);
        _panels.Add(panel.Name, panel);
      }
    }
  }

  Dictionary<string, RibbonPanel> _panels = new();

  Autodesk.Windows.RibbonTab _tab;
  Ribbon _ribbon;

}