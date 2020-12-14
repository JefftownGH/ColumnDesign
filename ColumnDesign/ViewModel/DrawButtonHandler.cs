using System;
using System.Windows.Input;

namespace ColumnDesign.ViewModel
{
    public class DrawButtonHandler : ICommand
    {
        private readonly Action _action;
        private readonly Func<bool> _canExecute;

        public DrawButtonHandler(Action action, Func<bool> canExecute)
        {
            _action = action;
            _canExecute = canExecute;
        }
        
        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
        
        public bool CanExecute(object parameter)
        {
            return _canExecute.Invoke();
        }

        public void Execute(object parameter)
        {
            _action();
        }
    }
}