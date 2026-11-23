using ExcelToJson.App.Services;
using Microsoft.AspNetCore.Components;

namespace ExcelToJson.App.Components;

/// <summary>
/// <see cref="AppState.Changed"/> olayına abone olan temel bileşen. Parametresi olmayan bileşenler
/// ebeveyn yeniden çizildiğinde güncellenmediği için adımlar ve paneller bu sınıftan türetilir.
/// </summary>
public abstract class StateComponent : ComponentBase, IDisposable
{
    [Inject] protected AppState State { get; set; } = default!;

    protected override void OnInitialized()
    {
        State.Changed += OnStateChanged;
        base.OnInitialized();
    }

    private void OnStateChanged() => InvokeAsync(StateHasChanged);

    public virtual void Dispose()
    {
        State.Changed -= OnStateChanged;
        GC.SuppressFinalize(this);
    }
}
