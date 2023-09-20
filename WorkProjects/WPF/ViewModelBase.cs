using System.ComponentModel;
using System.Windows;

namespace WPF.Mvvm {
  public class ViewModelBase : ObservableObject {
    private static readonly DependencyObject _DependencyObject
        = new DependencyObject();

    public bool IsInDesignMode { get; private set; }
        = DesignerProperties.GetIsInDesignMode(_DependencyObject);

    public System.Windows.Threading.Dispatcher Dispatcher {
      get { return _DependencyObject.Dispatcher; }
    }
  }
}