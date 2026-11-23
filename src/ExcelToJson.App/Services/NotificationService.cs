namespace ExcelToJson.App.Services;

public enum ToastKind
{
    Info,
    Success,
    Warning,
    Error
}

public sealed record ToastAction(string Label, Action Callback);

public sealed class Toast
{
    public Guid Id { get; } = Guid.NewGuid();
    public required ToastKind Kind { get; init; }
    public required string Title { get; init; }
    public string? Message { get; init; }
    public IReadOnlyList<ToastAction> Actions { get; init; } = [];
}

/// <summary>Sağ altta gösterilen hafif toast bildirimleri.</summary>
public sealed class NotificationService
{
    private readonly List<Toast> _toasts = [];
    private readonly object _lock = new();

    public IReadOnlyList<Toast> Toasts
    {
        get { lock (_lock) return _toasts.ToList(); }
    }

    public event Action? Changed;

    public void Show(ToastKind kind, string title, string? message = null, params ToastAction[] actions)
    {
        var toast = new Toast { Kind = kind, Title = title, Message = message, Actions = actions };
        lock (_lock)
        {
            _toasts.Add(toast);
            if (_toasts.Count > 4)
                _toasts.RemoveAt(0);
        }
        Changed?.Invoke();

        var timeout = kind switch
        {
            ToastKind.Error => TimeSpan.FromSeconds(12),
            _ when actions.Length > 0 => TimeSpan.FromSeconds(12),
            _ => TimeSpan.FromSeconds(5)
        };
        _ = DismissLaterAsync(toast.Id, timeout);
    }

    public void Info(string title, string? message = null) => Show(ToastKind.Info, title, message);
    public void Success(string title, string? message = null, params ToastAction[] actions) => Show(ToastKind.Success, title, message, actions);
    public void Warning(string title, string? message = null) => Show(ToastKind.Warning, title, message);
    public void Error(string title, string? message = null) => Show(ToastKind.Error, title, message);

    public void Dismiss(Guid id)
    {
        bool removed;
        lock (_lock)
            removed = _toasts.RemoveAll(t => t.Id == id) > 0;
        if (removed)
            Changed?.Invoke();
    }

    private async Task DismissLaterAsync(Guid id, TimeSpan delay)
    {
        await Task.Delay(delay);
        Dismiss(id);
    }
}
