using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SkHynix.HookUpToolkit.Exportor.GUI {
  class VisibilityValueConverter : IValueConverter {
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
      if (value is bool boolean) {
        if (boolean) {
          return Visibility.Visible;
        } else {
          return Visibility.Collapsed;
        }
      }
      return null;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
      return null;
    }
  }
}
