using System;
using System.Collections.Generic;

using Autodesk.Revit.UI.Events;

namespace CalculationCable;

public partial class Ribbon {

  public class Tab {

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

    public Panel CreatePanel(string name) => CreatePanel(name, Panel.VisibilityInDocument.None);

    public Panel CreatePanelForProject(string name) => CreatePanel(name, Panel.VisibilityInDocument.Project);

    public Panel CreatePanelForFamily(string name) => CreatePanel(name, Panel.VisibilityInDocument.Family);

    public bool FindPanel(string name, out Panel panel) {

      if (string.IsNullOrWhiteSpace(name)) {
        throw new ArgumentNullException("name");
      }

      if (_panels.TryGetValue(name, out panel)) {
        return true;
      }
      return false;
    }

    internal Panel CreatePanel(string name, Panel.VisibilityInDocument document) {

      if (string.IsNullOrWhiteSpace(name)) {
        throw new ArgumentNullException("name");
      }

      if (FindPanel(name, out _)) {
        throw new ArgumentException(name, $"Панель {name} существует");
      }

      var panel = _ribbon.Application.CreateRibbonPanel(_tab.Name, name);
      Panel panel2 = new(panel);
      panel2.TypeDocument = document;
      _panels.Add(name, panel2);
      return panel2;
    }

    internal Autodesk.Windows.RibbonTab get_RibbonTab() => _tab;

    internal Tab(Ribbon ribbon, Autodesk.Windows.RibbonTab tab) {

      _ribbon = ribbon;
      _tab = tab;

      var panels = _ribbon.Application.GetRibbonPanels(Name);
      if (panels.Count != 0) {
        foreach (var item in panels) {
          Panel panel = new(item);
          _panels.Add(panel.Name, panel);
        }
      }
    }

    Dictionary<string, Panel> _panels = new();

    Autodesk.Windows.RibbonTab _tab;
    Ribbon _ribbon;
  }
}

