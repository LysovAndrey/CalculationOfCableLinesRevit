using System;
using System.Collections.Generic;

using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace CalculationCable;

public partial class Ribbon {

  public class Panel {

    public enum VisibilityInDocument {
      None,
      Project,
      Family
    }

    public VisibilityInDocument TypeDocument { get; set; }
    public string Name { get => _panel.Name; set => _panel.Name = value; }
    public bool Visible { get => _panel.Visible; set => _panel.Visible = value; }
    public bool Enabled { get => _panel.Enabled; set => _panel.Enabled = value; }


    public void CreateSplitButtonGroup(string name, string text, Action<ButtonStacke> action) {

      if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(text)) {
        throw new ArgumentNullException("name or text");
      }

      if (action == null) {
        throw new ArgumentNullException("action");
      }

      SplitButtonData groupData = new SplitButtonData(name, text);
      SplitButton group = _panel.AddItem(groupData) as SplitButton;
      group.IsSynchronizedWithCurrentItem = false;

      ButtonStacke bts = new(group);

      action.Invoke(bts);

    }

    public void CreatePushButtonsStacke(Action<ButtonStacke> action) {

      if (action == null) {
        throw new ArgumentNullException("action");
      }

      ButtonStacke bts = new();

      action.Invoke(bts);

      if (bts.Count < 2 || bts.Count > 3) {
        throw new ArgumentNullException("Должно создаваться две или три кнопки");
      }

      var bt1 = bts.Buttons[0].Create();
      var bt2 = bts.Buttons[1].Create();
      if (bts.Count == 3) {
        var bt3 = bts.Buttons[2].Create();
        _panel.AddStackedItems(bt1, bt2, bt3);
      }
      else {
        _panel.AddStackedItems(bt1, bt2);
      }
    }

    public void CreatePushButton<TExternalCommandClass>(string text, Action<Button> action)
          where TExternalCommandClass : class, IExternalCommand {

      if (string.IsNullOrWhiteSpace(text)) {
        throw new ArgumentNullException("text");
      }
      if (action == null) {
        throw new ArgumentNullException("action");
      }

      string name = text;
      var commandClassType = typeof(TExternalCommandClass);

      Button bt = new Button(name, text, commandClassType);

      action.Invoke(bt);

      var buttonData = bt.Create();

      _panel.AddItem(buttonData);
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

    internal Panel(Autodesk.Revit.UI.RibbonPanel panel) {
      _panel = panel;
    }

    RibbonPanel _panel;
  }
}



