using System.ComponentModel;
using System.Runtime.CompilerServices;
namespace UiBase;

public class BaseNotifyObject : INotifyPropertyChanged, INotifyPropertyChanging
{
    /// <inheritdoc cref="INotifyPropertyChanged" />
    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>
    /// Raise property changed event.
    /// </summary>
    /// <param name="propertyName">Property name. Optional.</param>
    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    /// <inheritdoc cref="INotifyPropertyChanging"/>
    public event PropertyChangingEventHandler? PropertyChanging;

    /// <summary>
    /// Raise property changing event.
    /// </summary>
    /// <param name="propertyName">Property name. Optional.</param>
    protected void OnPropertyChanging([CallerMemberName] string? propertyName = null)
    {
        PropertyChanging?.Invoke(this, new PropertyChangingEventArgs(propertyName));
    }

    protected void SetProperty<T>(ref T prop, T value, [CallerMemberName] string? propertyName = null)
    {
        if (!Equals(prop, value))
        {
            OnPropertyChanging(propertyName);
            prop = value;
            OnPropertyChanged(propertyName);
        }
    }
}