using Autodesk.Revit.UI;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace CalculationCable;

public class RibbonButton {

  public RibbonButton SetDescription(string description) {
    _description = description;
    return this;
  }

  public RibbonButton SetContextualHelp(ContextualHelpType contextualHelpType, string helpPath) {
    _contextualHelp = new ContextualHelp(contextualHelpType, helpPath);
    return this;
  }

  public RibbonButton SetHelpUrl(string url) {
    _contextualHelp = new ContextualHelp(ContextualHelpType.Url, url);
    return this;
  }

  public RibbonButton SetSmallImage(ImageSource smallImage) {
    _smallImage = smallImage;
    return this;
  }

  public RibbonButton SetSmallImage(Bitmap smallImage) {
    _smallImage = ConvertBitmap(smallImage);
    return this;
  }

  public RibbonButton SetLargeImage(Bitmap largeImage) {
    _largeImage = ConvertBitmap(largeImage);
    return this;
  }

  public RibbonButton SetLargeImage(ImageSource largeImage) {
    _largeImage = largeImage;
    return this;
  }

  private BitmapSource ConvertBitmap(Bitmap source) {
    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
      source.GetHbitmap(),
      IntPtr.Zero,
      Int32Rect.Empty,
      BitmapSizeOptions.FromEmptyOptions());
  }

  public static BitmapSource ConvertToBitmapSource(Bitmap source) {
    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
      source.GetHbitmap(),
      IntPtr.Zero,
      Int32Rect.Empty,
      BitmapSizeOptions.FromEmptyOptions());
  }

  public virtual ButtonData Create() {

    PushButtonData bt = new PushButtonData(_name, _text, _assemblyLocation, _className);

    if (_largeImage != null) { bt.LargeImage = _largeImage; }
    if (_smallImage != null) { bt.Image = _smallImage; }
    if (_description != null) { bt.LongDescription = _description; }
    if (_contextualHelp != null) { bt.SetContextualHelp(_contextualHelp); }
    return bt;
  }

  public RibbonButton(string name, string text, Type externalCommandType) {
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

