using System;
using System.Windows.Input;

namespace UiBase;

public class Command : ICommand
{
    public event EventHandler? CanExecuteChanged;

    public Action Action { get; }

    public Command(Action action)
    {
        Action = action;
    }

    public bool CanExecute(object? parameter)
    {
        return true;
    }

    public void Execute(object? parameter)
    {
        Action();
    }
}