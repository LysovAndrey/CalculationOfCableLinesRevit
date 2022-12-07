using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace CalculationCable;

public class Ribbon {

  public UIControlledApplication Application { get; }

  public void Reset() { _tabs.ToList().ForEach(t => t.Value.Reset()); }

  public void Update(ViewActivatedEventArgs e) {

    _tabs.ToList().ForEach(t => t.Value.Update(e));

    //_tabs.TryGetValue("BIMDATA", out var tab);

    //if (e.Document.IsFamilyDocument) {
    //  tab.Visible = false;
    //}
    //else {
    //  tab.Visible = true;
    //}
  }

  public RibbonTab CreateTab(string name) {

    if (string.IsNullOrWhiteSpace(name)) {
      throw new ArgumentNullException("name");
    }

    if (FindTab(name, out _)) {
      throw new ArgumentException(name, $"Вкладка {name} существует");
    }

    var tabs = Autodesk.Windows.ComponentManager.Ribbon.Tabs;

    int num = 0;
    if (0 < tabs.Count) {
      do {
        Autodesk.Windows.RibbonTab ribbonTab = tabs[num];
        if (ribbonTab.IsContextualTab || ribbonTab.Id == "Modify") {
          break;
        }
        num++;
      }
      while (num < tabs.Count);
    }
    Application.CreateRibbonTab(name);
    var tab = Autodesk.Windows.ComponentManager.Ribbon.FindTab(name);
    RibbonTab tab2 = new(this, tab);
    //tabs.Insert(num, tab2.get_RibbonTab());
    _tabs.Add(name, tab2);

    return tab2;
  }

  public bool FindTab(string name, out RibbonTab tab) {

    if (string.IsNullOrWhiteSpace(name)) {
      throw new ArgumentNullException("name");
    }

    if (_tabs.TryGetValue(name, out tab)) {
      return true;
    }

    var tabs = Autodesk.Windows.ComponentManager.Ribbon.Tabs;

    foreach (var item in tabs) {
      if (string.CompareOrdinal(item.Id, name) == 0) {
        tab = new(this, item);
        _tabs.Add(name, tab);
        return true;
      }
    }
    return false;
  }

  public Ribbon(UIControlledApplication application) {
    Application = application;
  }

  Dictionary<string, RibbonTab> _tabs = new();

}
