using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WPF.Mvvm {
  public class ObservableObject : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;

    public void VerifyPropertyName(string propertyName) {
      var type = GetType();
      if (string.IsNullOrEmpty(propertyName)) { return; }

      if (type.GetProperty(propertyName) == null) {
        throw new ArgumentException("Property not found", propertyName);
      }
    }

    public virtual void RaisePropertyChanged(
        [CallerMemberName] string propertyName = "") {
      this.PropertyChanged?.Invoke(this,
          new PropertyChangedEventArgs(propertyName));
    }

    protected bool Set<T>(string propertyName, ref T field, T newValue) {
      if (EqualityComparer<T>.Default.Equals(field, newValue)) {
        return false;
      }

      field = newValue;
      RaisePropertyChanged(propertyName);
      return true;
    }

    protected bool Set<T>(ref T field, T newValue,
        [CallerMemberName] string propertyName = null) {
      return Set(propertyName, ref field, newValue);
    }
  }
}
