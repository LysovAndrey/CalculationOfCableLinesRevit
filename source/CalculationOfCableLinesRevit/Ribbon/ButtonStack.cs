using Autodesk.Revit.UI;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalculationCable;

public partial class Ribbon {

  public class ButtonStacke {

    public List<Button> Buttons { get => _buttons; }
    public int Count { get => _buttons.Count; }

    public ButtonStacke AddSeparator() {
      if (_parent is SplitButton) {
        _parent.AddSeparator();

      }
      return this;
    }

    public ButtonStacke CreateButton<TExternalCommandClass>(string text, Action<Button> action)
       where TExternalCommandClass : class, Autodesk.Revit.UI.IExternalCommand {

      if (action == null) {
        throw new ArgumentNullException("action");
      }

      CreateButton(text, text, typeof(TExternalCommandClass), action);

      return this;
    }

    internal void CreateButton(string name, string text, Type externalCommandType, Action<Button> action) {

      Button bt = new Button(name, text, externalCommandType);

      if (action != null) {
        action.Invoke(bt);
      }
      if (_parent is SplitButton) {
        _parent?.AddPushButton(bt.Create());

      }
      else {
        _buttons.Add(bt);
      }
    }

    internal ButtonStacke(SplitButton parent) { _parent = parent; }
    internal ButtonStacke() { }

    SplitButton _parent;
    List<Button> _buttons = new();

  }
}



