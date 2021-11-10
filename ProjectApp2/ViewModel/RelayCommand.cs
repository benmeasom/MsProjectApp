using System;
using System.Windows.Input;

namespace ProjectApp2.ViewModel
{
    public class RelayCommand : ICommand
    {
        private Action mAction;

        public event EventHandler CanExecuteChanged = (sender, e) => { };

        public bool CanExecute(object parameter)
        {
            return true;
        }
        public void Execute(object parameter)
        {
            mAction();
        }

        public RelayCommand(Action action)
        {
            mAction = action;
        }
    }
}
