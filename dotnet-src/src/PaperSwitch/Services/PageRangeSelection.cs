namespace PaperSwitch.Services;

public static class PageRangeSelection
{
    public static HashSet<int> Parse(string text, int pageCount)
    {
        var result = new HashSet<int>();
        string normalized = text
            .Replace('\u3000', ' ')
            .Replace('，', ',')
            .Replace('、', ',')
            .Replace('–', '-')
            .Replace('—', '-')
            .Replace('－', '-')
            .Replace('～', '-')
            .Replace('~', '-');

        foreach (var part in normalized.Split(','))
        {
            var ends = part.Trim().Split('-');
            if (ends.Length > 2 || !int.TryParse(ends[0].Trim(), out int first) || first < 1 || first > pageCount)
                throw new FormatException("請輸入有效的畫布頁序，例如 1-5,8；頁序不可超出目前頁數。");
            int last = first;
            if (ends.Length == 2 && (!int.TryParse(ends[1].Trim(), out last) || last < first || last > pageCount))
                throw new FormatException("範圍結尾必須大於或等於起始頁，且不可超出目前頁數。");
            for (int page = first; page <= last; page++) result.Add(page - 1);
        }
        return result;
    }
}
