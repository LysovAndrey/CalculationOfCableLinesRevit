using Autodesk.Revit.UI.Events;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalculationCable;

public class RibbonPanel {

  public enum VisibilityInDocument {
    None,
    Project,
    Family
  }

  public VisibilityInDocument TypeDocument { get; set; }
  public string Name { get => _panel.Name; set => _panel.Name = value; }
  public bool Visible { get => _panel.Visible; set => _panel.Visible = value; }
  public bool Enabled { get => _panel.Enabled; set => _panel.Enabled = value; }


  public RibbonPanel CreatePushButton<TExternalCommandClass>
    (string text, Action<RibbonButton> action)
        where TExternalCommandClass : class, Autodesk.Revit.UI.IExternalCommand {

    var commandClassType = typeof(TExternalCommandClass);
    string name = text;
    RibbonButton bt = new RibbonButton(name, text, commandClassType);

    if (action != null) {
      action.Invoke(bt);
    }

    var buttonData = bt.Create();

    _panel.AddItem(buttonData);
    return this;
  }

    public void AddSeparator() => _panel.AddSeparator();

  public void Update(ViewActivatedEventArgs e) {
    if (TypeDocument == VisibilityInDocument.Project && e.Document.IsFamilyDocument == true) {
      _panel.Visible = false;
    }
    else if (TypeDocument == VisibilityInDocument.Project && e.Document.IsFamilyDocument != true) {
      _panel.Visible = true;
    }
    else if (TypeDocument == VisibilityInDocument.Family && e.Document.IsFamilyDocument == true) {
      _panel.Visible = true;
    }
    else if (TypeDocument == VisibilityInDocument.Family && e.Document.IsFamilyDocument != true) {
      _panel.Visible = false;
    }
    else {
      _panel.Visible = true;
    }

  }

  internal RibbonPanel(Autodesk.Revit.UI.RibbonPanel panel) {
    _panel = panel;
  }

  Autodesk.Revit.UI.RibbonPanel _panel;
}

