// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Windows.Input;

namespace MsGraph_Samples.Helpers
{
    public class RelayCommand : ICommand
    {
        readonly Action _executeMethod;
        readonly Func<bool>? _canExecuteMethod;

        public RelayCommand(Action executeMethod, Func<bool>? canexecuteMethod = null)
        {
            _executeMethod = executeMethod;
            _canExecuteMethod = canexecuteMethod;
        }

        public void Execute(object _)
        {
            _executeMethod();
        }

        public bool CanExecute(object _)
        {
            if (_canExecuteMethod == null)
                return true;

            return _canExecuteMethod();
        }

        public event EventHandler CanExecuteChanged
        {
            add
            {
                CommandManager.RequerySuggested += value;
            }
            remove
            {
                CommandManager.RequerySuggested -= value;
            }
        }
    }

    public class RelayCommand<T> : ICommand
    {
        readonly Action<T> _executeMethod;
        readonly Func<T, bool>? _canExecuteMethod;

        public RelayCommand(Action<T> executeMethod, Func<T, bool>? canexecuteMethod = null)
        {
            _executeMethod = executeMethod;
            _canExecuteMethod = canexecuteMethod;
        }

        public void Execute(object parameter)
        {
            _executeMethod((T)parameter);
        }

        public bool CanExecute(object parameter)
        {
            if (_canExecuteMethod == null)
                return true;

            return _canExecuteMethod((T)parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add
            {
                CommandManager.RequerySuggested += value;
            }
            remove
            {
                CommandManager.RequerySuggested -= value;
            }
        }
    }
}