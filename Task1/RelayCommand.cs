using System;
using System.Windows;
using System.Windows.Input;

namespace Task1 {

    public class RelayCommand : ICommand {

        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute)
            : this(execute, null!) { }

        private RelayCommand(Action execute, Func<bool> canExecute) {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter) => _canExecute == null! || _canExecute();

        public void Execute(object? parameter) {
            _execute();
        }
    }
}