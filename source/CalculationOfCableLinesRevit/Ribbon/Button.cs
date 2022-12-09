using System;
using System.Drawing;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using Autodesk.Revit.UI;

namespace CalculationCable;

public partial class Ribbon {

  public class Button {

    public Button SetDescription(string description) {
      _description = description;
      return this;
    }

    public Button SetContextualHelp(ContextualHelpType contextualHelpType, string helpPath) {
      _contextualHelp = new ContextualHelp(contextualHelpType, helpPath);
      return this;
    }

    public Button SetHelpUrl(string url) {
      _contextualHelp = new ContextualHelp(ContextualHelpType.Url, url);
      return this;
    }

    public Button SetSmallImage(ImageSource smallImage) {
      _smallImage = smallImage;
      return this;
    }

    public Button SetSmallImage(Bitmap smallImage) {
      _smallImage = ConvertBitmap(smallImage);
      return this;
    }

    public Button SetLargeImage(Bitmap largeImage) {
      _largeImage = ConvertBitmap(largeImage);
      return this;
    }

    public Button SetLargeImage(ImageSource largeImage) {
      _largeImage = largeImage;
      return this;
    }

    internal BitmapSource ConvertToBitmapSource(Bitmap source) {
      return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
        source.GetHbitmap(),
        IntPtr.Zero,
        Int32Rect.Empty,
        BitmapSizeOptions.FromEmptyOptions());
    }

    internal PushButtonData Create() {

      PushButtonData bt = new PushButtonData(_name, _text, _assemblyLocation, _className);

      if (_largeImage != null) { bt.LargeImage = _largeImage; }
      if (_smallImage != null) { bt.Image = _smallImage; }
      if (_description != null) { bt.LongDescription = _description; }
      if (_contextualHelp != null) { bt.SetContextualHelp(_contextualHelp); }

      return bt;
    }

    private BitmapSource ConvertBitmap(Bitmap source) {
      return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
        source.GetHbitmap(),
        IntPtr.Zero,
        Int32Rect.Empty,
        BitmapSizeOptions.FromEmptyOptions());
    }

    internal Button(string name, string text, Type externalCommandType) {
      _name = name;
      _text = text;
      if (externalCommandType != null) {
        _assemblyLocation = externalCommandType.Assembly.Location;
        _className = externalCommandType.FullName;
      }
    }

    ContextualHelp _contextualHelp;
    string _description;
    ImageSource _smallImage;
    ImageSource _largeImage;
    string _className;
    string _assemblyLocation;
    string _text;
    string _name;
  }
}



