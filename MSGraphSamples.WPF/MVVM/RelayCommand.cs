// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace MsGraph_Samples.MVVM
{
    public abstract class BaseRelayCommand : ICommand
    {
        public event EventHandler? CanExecuteChanged
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

        public static void RaiseCanExecuteChanged() => CommandManager.InvalidateRequerySuggested();

        abstract public bool CanExecute(object? parameter);

        abstract public void Execute(object? parameter);
    }

    public class RelayCommand : BaseRelayCommand
    {
        private readonly Action _execute;
        private readonly Func<bool>? _canExecute;

        public RelayCommand(Action execute, Func<bool>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        override public void Execute(object? _ = null) => _execute();

        override public bool CanExecute(object? _ = null) => _canExecute == null || _canExecute();

    }

    public class RelayCommand<T> : BaseRelayCommand
    {
        private readonly Action<T> _execute;
        private readonly Predicate<T>? _canExecute;

        public RelayCommand(Action<T> execute, Predicate<T>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        override public void Execute(object? parameter)
        {
            if (parameter is T tParam)
                _execute(tParam);
        }

        override public bool CanExecute(object? parameter)
        {
            if (_canExecute == null)
                return true;

            if (parameter is T tParam)
                return _canExecute(tParam);

            return false;
        }
    }

    public class AsyncRelayCommand : BaseRelayCommand
    {
        private bool _isExecuting;
        private readonly Func<Task> _execute;
        private readonly Func<bool>? _canExecute;

        public AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        override public void Execute(object? _ = null) => ExecuteAsync().Await();
        override public bool CanExecute(object? _ = null) => !_isExecuting && (_canExecute == null || _canExecute());

        public async Task ExecuteAsync()
        {
            if (!CanExecute())
                return;

            _isExecuting = true;

            try
            {
                await _execute();
            }
            finally
            {
                _isExecuting = false;
            }

            RaiseCanExecuteChanged();
        }

    }

    public class AsyncRelayCommand<T> : BaseRelayCommand
    {
        private bool _isExecuting;
        private readonly Func<T, Task> _executeTask;
        private readonly Predicate<T>? _canExecute;

        public AsyncRelayCommand(Func<T, Task> executeTask, Predicate<T>? canExecute = null)
        {
            _executeTask = executeTask;
            _canExecute = canExecute;
        }

        override public void Execute(object? parameter) => ExecuteAsync(parameter).Await();
        override public bool CanExecute(object? parameter)
        {
            if (_isExecuting)
                return false;

            if (_canExecute == null)
                return true;

            if (parameter is T tParam)
                return _canExecute(tParam);

            return false;
        }

        public async Task ExecuteAsync(object? parameter)
        {
            if (parameter is not T tParam)
                return;

            if (!CanExecute(tParam))
                return;

            _isExecuting = true;

            try
            {
                await _executeTask(tParam);
            }
            finally
            {
                _isExecuting = false;
            }

            RaiseCanExecuteChanged();
        }
    }
}