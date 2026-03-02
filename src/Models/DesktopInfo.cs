namespace VirtualDesktopUtils.Models;

internal sealed record DesktopInfo(Guid Id, int Index, string Name, bool IsCurrent)
{
    public string DisplayName => $"{Name} ({Index})";
}
