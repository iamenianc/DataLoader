namespace ExcelStage;

/// <summary>
/// An arrow-key driven selection list for the console. Up/Down move the
/// highlight (PgUp/PgDn/Home/End jump), Enter confirms, Esc cancels. The list
/// scrolls within a fixed on-screen window, so it works correctly even when
/// there are far more items than fit on the screen. Falls back to a numbered
/// prompt when the console is redirected.
/// </summary>
public static class Menu
{
    public static Nav<int> Select(string title, IReadOnlyList<string> items)
    {
        if (items.Count == 0)
        {
            return Nav<int>.Back;
        }

        if (Console.IsOutputRedirected || Console.IsInputRedirected)
        {
            return SelectFallback(title, items);
        }

        Console.WriteLine();
        Console.WriteLine(items.Count > 1 ? $"{title}  ({items.Count} items)" : title);
        Console.WriteLine("(Up/Down move, Enter select | B back, R restart, Esc/Q cancel)");

        // Size the visible window to the console, leaving room for the headers.
        var windowHeight = Math.Min(items.Count, Math.Max(3, SafeHeight() - 5));

        // Reserve the rows up front so the viewport stays on-screen even if
        // printing them scrolls the buffer; then compute the (stable) top row.
        for (var i = 0; i < windowHeight; i++)
        {
            Console.WriteLine();
        }

        var viewTop = Math.Max(0, Console.CursorTop - windowHeight);

        var previousCursorVisible = true;
        try { previousCursorVisible = Console.CursorVisible; } catch { /* not supported everywhere */ }
        try { Console.CursorVisible = false; } catch { /* ignore */ }

        var index = 0;
        var offset = 0;

        try
        {
            Render(items, index, offset, windowHeight, viewTop);
            while (true)
            {
                var key = Console.ReadKey(intercept: true).Key;
                switch (key)
                {
                    case ConsoleKey.UpArrow:
                        index = Math.Max(0, index - 1);
                        break;
                    case ConsoleKey.DownArrow:
                        index = Math.Min(items.Count - 1, index + 1);
                        break;
                    case ConsoleKey.PageUp:
                        index = Math.Max(0, index - windowHeight);
                        break;
                    case ConsoleKey.PageDown:
                        index = Math.Min(items.Count - 1, index + windowHeight);
                        break;
                    case ConsoleKey.Home:
                        index = 0;
                        break;
                    case ConsoleKey.End:
                        index = items.Count - 1;
                        break;
                    case ConsoleKey.Enter:
                        return Nav<int>.FromValue(index);

                    // Navigation
                    case ConsoleKey.LeftArrow:
                    case ConsoleKey.Backspace:
                    case ConsoleKey.B:
                        return Nav<int>.Back;
                    case ConsoleKey.R:
                        return Nav<int>.Restart;
                    case ConsoleKey.Escape:
                    case ConsoleKey.Q:
                        return Nav<int>.Quit;

                    default:
                        continue;
                }

                // Keep the highlighted item inside the visible window.
                if (index < offset)
                {
                    offset = index;
                }
                else if (index >= offset + windowHeight)
                {
                    offset = index - windowHeight + 1;
                }

                Render(items, index, offset, windowHeight, viewTop);
            }
        }
        finally
        {
            try { Console.SetCursorPosition(0, viewTop + windowHeight); } catch { /* ignore */ }
            try { Console.CursorVisible = previousCursorVisible; } catch { /* ignore */ }
            Console.WriteLine();
        }
    }

    /// <summary>
    /// Like <see cref="Select"/>, but lets the user pick several items. Space
    /// toggles the highlighted item, Enter confirms. The returned indexes are in
    /// the order they were selected, so callers that care about "the first one
    /// chosen" (e.g. which sheet sets the column types) can rely on it. Pressing
    /// Enter with nothing ticked selects just the highlighted item.
    /// </summary>
    public static Nav<List<int>> SelectMultiple(string title, IReadOnlyList<string> items)
    {
        if (items.Count == 0)
        {
            return Nav<List<int>>.Back;
        }

        if (Console.IsOutputRedirected || Console.IsInputRedirected)
        {
            return SelectMultipleFallback(title, items);
        }

        // Selection order is tracked explicitly so the first ticked item stays first.
        var order = new List<int>();

        Console.WriteLine();
        Console.WriteLine($"{title}  ({items.Count} items)");
        Console.WriteLine("(Up/Down move, Space toggle, A all/none, Enter confirm | B back, R restart, Esc/Q cancel)");

        var windowHeight = Math.Min(items.Count, Math.Max(3, SafeHeight() - 5));
        for (var i = 0; i < windowHeight; i++)
        {
            Console.WriteLine();
        }

        var viewTop = Math.Max(0, Console.CursorTop - windowHeight);

        var previousCursorVisible = true;
        try { previousCursorVisible = Console.CursorVisible; } catch { /* not supported everywhere */ }
        try { Console.CursorVisible = false; } catch { /* ignore */ }

        var index = 0;
        var offset = 0;

        try
        {
            RenderMulti(items, order, index, offset, windowHeight, viewTop);
            while (true)
            {
                var key = Console.ReadKey(intercept: true).Key;
                switch (key)
                {
                    case ConsoleKey.UpArrow:
                        index = Math.Max(0, index - 1);
                        break;
                    case ConsoleKey.DownArrow:
                        index = Math.Min(items.Count - 1, index + 1);
                        break;
                    case ConsoleKey.PageUp:
                        index = Math.Max(0, index - windowHeight);
                        break;
                    case ConsoleKey.PageDown:
                        index = Math.Min(items.Count - 1, index + windowHeight);
                        break;
                    case ConsoleKey.Home:
                        index = 0;
                        break;
                    case ConsoleKey.End:
                        index = items.Count - 1;
                        break;

                    case ConsoleKey.Spacebar:
                        Toggle(order, index);
                        break;
                    case ConsoleKey.A:
                        if (order.Count == items.Count)
                        {
                            order.Clear();
                        }
                        else
                        {
                            for (var i = 0; i < items.Count; i++)
                            {
                                if (!order.Contains(i)) order.Add(i);
                            }
                        }
                        break;

                    case ConsoleKey.Enter:
                        var chosen = order.Count > 0 ? new List<int>(order) : new List<int> { index };
                        return Nav<List<int>>.FromValue(chosen);

                    case ConsoleKey.LeftArrow:
                    case ConsoleKey.Backspace:
                    case ConsoleKey.B:
                        return Nav<List<int>>.Back;
                    case ConsoleKey.R:
                        return Nav<List<int>>.Restart;
                    case ConsoleKey.Escape:
                    case ConsoleKey.Q:
                        return Nav<List<int>>.Quit;

                    default:
                        continue;
                }

                if (index < offset)
                {
                    offset = index;
                }
                else if (index >= offset + windowHeight)
                {
                    offset = index - windowHeight + 1;
                }

                RenderMulti(items, order, index, offset, windowHeight, viewTop);
            }
        }
        finally
        {
            try { Console.SetCursorPosition(0, viewTop + windowHeight); } catch { /* ignore */ }
            try { Console.CursorVisible = previousCursorVisible; } catch { /* ignore */ }
            Console.WriteLine();
        }
    }

    private static void Toggle(List<int> order, int index)
    {
        if (!order.Remove(index))
        {
            order.Add(index);
        }
    }

    private static void RenderMulti(
        IReadOnlyList<string> items, List<int> order, int selected, int offset, int windowHeight, int viewTop)
    {
        var maxWidth = Math.Max(1, SafeWidth() - 1);

        for (var r = 0; r < windowHeight; r++)
        {
            try { Console.SetCursorPosition(0, viewTop + r); } catch { /* ignore */ }

            var itemIndex = offset + r;
            var isSelected = itemIndex == selected;

            string line;
            if (itemIndex < items.Count)
            {
                var marker = isSelected ? ">" : " ";
                // Ticked items show their 1-based selection order; the "1" is the
                // sheet whose data drives the column types.
                var pos = order.IndexOf(itemIndex);
                var box = pos >= 0 ? $"[{pos + 1}]" : "[ ]";
                var scroll = ScrollHint(itemIndex, offset, windowHeight, items.Count);
                line = $" {marker} {box} {items[itemIndex]}{scroll}";
            }
            else
            {
                line = string.Empty;
            }

            if (line.Length > maxWidth)
            {
                line = line[..maxWidth];
            }

            line = line.PadRight(maxWidth);

            if (isSelected)
            {
                Console.ForegroundColor = ConsoleColor.Black;
                Console.BackgroundColor = ConsoleColor.Gray;
            }

            Console.Write(line);
            Console.ResetColor();
        }
    }

    private static Nav<List<int>> SelectMultipleFallback(string title, IReadOnlyList<string> items)
    {
        Console.WriteLine();
        Console.WriteLine(title);
        for (var i = 0; i < items.Count; i++)
        {
            Console.WriteLine($"  {i + 1}. {items[i]}");
        }

        while (true)
        {
            Console.Write(
                $"Enter number(s) 1-{items.Count} in order, comma-separated (e.g. 2,1) | B back, R restart, Q cancel: ");
            var input = Console.ReadLine()?.Trim();
            switch (input?.ToLowerInvariant())
            {
                case "b" or "back":
                    return Nav<List<int>>.Back;
                case "r" or "restart":
                    return Nav<List<int>>.Restart;
                case "q" or "quit" or "cancel" or null or "":
                    return Nav<List<int>>.Quit;
            }

            var parts = input!.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var chosen = new List<int>();
            var ok = true;
            foreach (var part in parts)
            {
                if (int.TryParse(part, out var n) && n >= 1 && n <= items.Count)
                {
                    if (!chosen.Contains(n - 1)) chosen.Add(n - 1); // preserve typed order
                }
                else
                {
                    ok = false;
                    break;
                }
            }

            if (ok && chosen.Count > 0)
            {
                return Nav<List<int>>.FromValue(chosen);
            }

            Console.WriteLine("  ! Enter one or more valid numbers, separated by commas.");
        }
    }

    private static void Render(
        IReadOnlyList<string> items, int selected, int offset, int windowHeight, int viewTop)
    {
        // Never write into the final column - that forces a line wrap, which is
        // what produced the "staircase" effect.
        var maxWidth = Math.Max(1, SafeWidth() - 1);

        for (var r = 0; r < windowHeight; r++)
        {
            try { Console.SetCursorPosition(0, viewTop + r); } catch { /* ignore */ }

            var itemIndex = offset + r;
            var isSelected = itemIndex == selected;

            string line;
            if (itemIndex < items.Count)
            {
                var marker = isSelected ? ">" : " ";
                var scroll = ScrollHint(itemIndex, offset, windowHeight, items.Count);
                line = $" {marker} {items[itemIndex]}{scroll}";
            }
            else
            {
                line = string.Empty;
            }

            if (line.Length > maxWidth)
            {
                line = line[..maxWidth];
            }

            line = line.PadRight(maxWidth);

            if (isSelected)
            {
                Console.ForegroundColor = ConsoleColor.Black;
                Console.BackgroundColor = ConsoleColor.Gray;
            }

            Console.Write(line);
            Console.ResetColor();
        }
    }

    // Shows a small marker on the first/last visible rows when more items exist
    // above/below the current window.
    private static string ScrollHint(int itemIndex, int offset, int windowHeight, int total)
    {
        if (itemIndex == offset && offset > 0)
        {
            return "   (more above)";
        }

        if (itemIndex == offset + windowHeight - 1 && offset + windowHeight < total)
        {
            return "   (more below)";
        }

        return string.Empty;
    }

    private static int SafeWidth()
    {
        try { return Console.WindowWidth > 0 ? Console.WindowWidth : 80; }
        catch { return 80; }
    }

    private static int SafeHeight()
    {
        try { return Console.WindowHeight > 0 ? Console.WindowHeight : 25; }
        catch { return 25; }
    }

    private static Nav<int> SelectFallback(string title, IReadOnlyList<string> items)
    {
        Console.WriteLine();
        Console.WriteLine(title);
        for (var i = 0; i < items.Count; i++)
        {
            Console.WriteLine($"  {i + 1}. {items[i]}");
        }

        while (true)
        {
            Console.Write($"Enter a number (1-{items.Count}) | B back, R restart, Q cancel: ");
            var input = Console.ReadLine()?.Trim();
            switch (input?.ToLowerInvariant())
            {
                case "b" or "back":
                    return Nav<int>.Back;
                case "r" or "restart":
                    return Nav<int>.Restart;
                case "q" or "quit" or "cancel" or null or "":
                    return Nav<int>.Quit;
            }

            if (int.TryParse(input, out var choice) && choice >= 1 && choice <= items.Count)
            {
                return Nav<int>.FromValue(choice - 1);
            }

            Console.WriteLine("  ! Not a valid choice.");
        }
    }
}
