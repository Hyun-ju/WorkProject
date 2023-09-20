using System;
using System.Windows.Input;

namespace WPF.Mvvm {
  public class RelayCommand : ICommand {
    public event EventHandler CanExecuteChanged;


    private readonly Action _Execute;
    private readonly Func<bool> _CanExecute;


    public RelayCommand(Action execute)
      : this(execute, null) {
    }
    public RelayCommand(Action execute, Func<bool> canExecute) {
      if (execute == null) {
        throw new ArgumentNullException("execute");
      }

      _Execute = execute;
      if (canExecute != null) {
        _CanExecute = canExecute;
      } else {
        _CanExecute = () => true;
      }
    }

    public void RaiseCanExecuteChanged() {
      this.CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    public bool CanExecute(object parameter) {
      return _CanExecute();
    }

    public virtual void Execute(object parameter) {
      if (!CanExecute(parameter) && _Execute == null) {
      }
      _Execute();
    }
  }
}
