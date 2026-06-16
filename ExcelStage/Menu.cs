namespace ExcelStage;

/// <summary>
/// A simple arrow-key driven selection list for the console. Up/Down move the
/// highlight, Enter confirms, Esc cancels. Falls back to a numbered prompt when
/// the output is redirected (so it still works in non-interactive shells).
/// </summary>
public static class Menu
{
    /// <summary>
    /// Shows <paramref name="items"/> and returns the chosen index, or -1 if the
    /// user pressed Esc to cancel.
    /// </summary>
    public static int Select(string title, IReadOnlyList<string> items)
    {
        if (items.Count == 0)
        {
            return -1;
        }

        if (Console.IsOutputRedirected || Console.IsInputRedirected)
        {
            return SelectFallback(title, items);
        }

        Console.WriteLine();
        Console.WriteLine(title);
        Console.WriteLine("(Use the Up/Down arrow keys, Enter to choose, Esc to cancel)");

        var listTop = Console.CursorTop;
        var index = 0;
        var previousCursorVisible = true;
        try { previousCursorVisible = Console.CursorVisible; } catch { /* not supported everywhere */ }
        try { Console.CursorVisible = false; } catch { /* ignore */ }

        try
        {
            Render(items, index, listTop);
            while (true)
            {
                var key = Console.ReadKey(intercept: true).Key;
                switch (key)
                {
                    case ConsoleKey.UpArrow:
                        index = (index - 1 + items.Count) % items.Count;
                        Render(items, index, listTop);
                        break;
                    case ConsoleKey.DownArrow:
                        index = (index + 1) % items.Count;
                        Render(items, index, listTop);
                        break;
                    case ConsoleKey.Enter:
                        return index;
                    case ConsoleKey.Escape:
                        return -1;
                }
            }
        }
        finally
        {
            // Park the cursor below the list so later output starts on a clean line.
            try { Console.SetCursorPosition(0, listTop + items.Count); } catch { /* ignore */ }
            try { Console.CursorVisible = previousCursorVisible; } catch { /* ignore */ }
            Console.WriteLine();
        }
    }

    private static void Render(IReadOnlyList<string> items, int selected, int listTop)
    {
        var width = SafeWidth();
        for (var i = 0; i < items.Count; i++)
        {
            try { Console.SetCursorPosition(0, listTop + i); } catch { /* ignore */ }

            var marker = i == selected ? ">" : " ";
            var line = $" {marker} {items[i]}";
            if (line.Length > width)
            {
                line = line[..width];
            }

            if (i == selected)
            {
                Console.ForegroundColor = ConsoleColor.Black;
                Console.BackgroundColor = ConsoleColor.Gray;
            }

            // Pad to the line width so the previous highlight is fully cleared.
            Console.Write(line.PadRight(Math.Max(0, width - 1)));
            Console.ResetColor();
        }
    }

    private static int SafeWidth()
    {
        try
        {
            return Console.WindowWidth > 0 ? Console.WindowWidth : 80;
        }
        catch
        {
            return 80;
        }
    }

    private static int SelectFallback(string title, IReadOnlyList<string> items)
    {
        Console.WriteLine();
        Console.WriteLine(title);
        for (var i = 0; i < items.Count; i++)
        {
            Console.WriteLine($"  {i + 1}. {items[i]}");
        }

        while (true)
        {
            Console.Write($"Enter a number (1-{items.Count}), or blank to cancel: ");
            var input = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(input))
            {
                return -1;
            }

            if (int.TryParse(input, out var choice) && choice >= 1 && choice <= items.Count)
            {
                return choice - 1;
            }

            Console.WriteLine("  ! Not a valid choice.");
        }
    }
}
