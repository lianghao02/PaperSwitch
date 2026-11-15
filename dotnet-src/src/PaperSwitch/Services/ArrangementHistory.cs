using PaperSwitch.Models;

namespace PaperSwitch.Services;

internal sealed class ArrangementHistory
{
    public Stack<ArrangementSnapshot> Undo { get; } = new();
    public Stack<ArrangementSnapshot> Redo { get; } = new();
    public IEnumerable<string> SourcePaths => Undo.Concat(Redo).SelectMany(s => s.PageStates).Select(s => s.Page.SourceFilePath);
}

internal sealed class ArrangementSnapshot
{
    public IReadOnlyList<PageState> PageStates { get; }
    public PaperItem? SelectedPage { get; }

    public ArrangementSnapshot(IEnumerable<PaperItem> pages, PaperItem? selectedPage)
    {
        PageStates = pages.Select(page => new PageState(page)).ToList();
        SelectedPage = selectedPage;
    }

    public bool Matches(IEnumerable<PaperItem> pages, PaperItem? selectedPage)
    {
        var currentPages = pages.ToList();
        return SelectedPage == selectedPage
            && PageStates.Count == currentPages.Count
            && PageStates.Zip(currentPages, (state, page) => state.Matches(page)).All(matches => matches);
    }
}

internal sealed class PageState
{
    public PaperItem Page { get; }
    private int Rotation { get; }
    private bool IsSelected { get; }
    private string SourceFileName { get; }

    public PageState(PaperItem page)
    {
        Page = page;
        Rotation = page.Rotation;
        IsSelected = page.IsSelected;
        SourceFileName = page.SourceFileName;
    }

    public bool Matches(PaperItem page)
    {
        return ReferenceEquals(Page, page)
            && Rotation == page.Rotation
            && IsSelected == page.IsSelected
            && SourceFileName == page.SourceFileName;
    }

    public void Restore()
    {
        Page.Rotation = Rotation;
        Page.IsSelected = IsSelected;
        Page.SourceFileName = SourceFileName;
    }
}
