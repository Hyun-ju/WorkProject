using System;
using System.Windows.Input;

namespace WPF.Mvvm {
  public class RelayCommand<T> : ICommand {
    public event EventHandler CanExecuteChanged;


    private readonly Action<T> _Execute;
    private readonly Func<T, bool> _CanExecute;


    public RelayCommand(Action<T> execute)
      : this(execute, null) {
    }

    public RelayCommand(Action<T> execute, Func<T, bool> canExecute) {
      if (execute == null) {
        throw new ArgumentNullException("execute");
      }

      _Execute = execute;
      if (canExecute != null) {
        _CanExecute = canExecute;
      } else {
        _CanExecute = (T) => true;
      }
    }

    public void RaiseCanExecuteChanged() {
      this.CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    public bool CanExecute(object parameter) {
      return _CanExecute((T)parameter);
    }

    public virtual void Execute(object parameter) {
      if (!CanExecute(parameter) || _Execute == null) {
        return;
      }
      _Execute((T)parameter);
    }
  }
}
