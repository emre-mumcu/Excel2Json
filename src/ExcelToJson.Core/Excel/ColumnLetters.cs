namespace ExcelToJson.Core.Excel;

/// <summary>Excel sütun numarası ↔ harf dönüşümü.</summary>
public static class ColumnLetters
{
    public static string FromNumber(int column)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 1);
        Span<char> buffer = stackalloc char[4];
        var pos = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--pos] = (char)('A' + column % 26);
            column /= 26;
        }
        return new string(buffer[pos..]);
    }

    public static int ToNumber(string letters)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(letters);
        var result = 0;
        foreach (var ch in letters.Trim().ToUpperInvariant())
        {
            if (ch is < 'A' or > 'Z')
                throw new ArgumentException($"Geçersiz sütun harfi: {letters}", nameof(letters));
            result = result * 26 + (ch - 'A' + 1);
        }
        return result;
    }
}
