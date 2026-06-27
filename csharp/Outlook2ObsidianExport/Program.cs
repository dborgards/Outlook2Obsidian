using System;
using System.Globalization;
using System.IO;

namespace Outlook2Obsidian.Export
{
    public static class Program
    {
        public static int Main(string[] args)
        {
            try
            {
                var opts = Options.Parse(args);
                if (opts.ShowHelp) { PrintHelp(); return 0; }

                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string configPath = opts.ConfigPath ?? Path.Combine(baseDir, "config.txt");
                string statePath = opts.StatePath ?? Path.Combine(baseDir, "state.txt");

                var cfg = Config.Load(configPath);
                var state = ExportState.Load(statePath);

                DateTime? start, end;
                ResolveRange(opts, state, out start, out end);

                Console.WriteLine("Outlook -> Obsidian export");
                Console.WriteLine("  vault : " + cfg.VaultPath);
                Console.WriteLine("  folder: " + cfg.FolderPath + (cfg.Recurse ? " (recursive)" : ""));
                Console.WriteLine("  range : " + Describe(start) + " .. " + Describe(end));
                if (opts.DryRun) Console.WriteLine("  mode  : DRY-RUN (no files written)");
                Console.WriteLine();

                int written;
                using (var exporter = new OutlookExporter(cfg, state, Console.WriteLine))
                    written = exporter.Export(start, end, opts.DryRun);

                if (!opts.DryRun)
                    state.Save(statePath);

                Console.WriteLine();
                Console.WriteLine($"Done. {written} note(s) {(opts.DryRun ? "would be " : "")}written.");
                return 0;
            }
            catch (OptionException oe)
            {
                Console.Error.WriteLine("Error: " + oe.Message);
                Console.Error.WriteLine("Run with --help for usage.");
                return 2;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return 1;
            }
        }

        private static void ResolveRange(Options opts, ExportState state,
                                         out DateTime? start, out DateTime? end)
        {
            DateTime now = DateTime.Now;

            if (opts.SinceLastExport)
            {
                // First ever run with no watermark falls back to a full export.
                start = state.LastExport;
                end = null;
                return;
            }

            switch ((opts.Range ?? "").ToLowerInvariant())
            {
                case "day": start = now.AddDays(-1); end = null; return;
                case "week": start = now.AddDays(-7); end = null; return;
                case "month": start = now.AddMonths(-1); end = null; return;
                case "year": start = now.AddYears(-1); end = null; return;
                case "all": start = null; end = null; return;
            }

            // Explicit "YYYY-MM-DD..YYYY-MM-DD" (end date inclusive).
            if (!string.IsNullOrEmpty(opts.Range) && opts.Range.Contains(".."))
            {
                var bounds = opts.Range.Split(new[] { ".." }, StringSplitOptions.None);
                if (bounds.Length == 2)
                {
                    start = ParseDate(bounds[0], "start");
                    DateTime? e = string.IsNullOrWhiteSpace(bounds[1]) ? (DateTime?)null
                                  : ParseDate(bounds[1], "end");
                    end = e?.AddDays(1); // make the end date inclusive
                    return;
                }
            }

            throw new OptionException(
                "Specify --range <day|week|month|year|all|YYYY-MM-DD..YYYY-MM-DD> or --since-last-export.");
        }

        private static DateTime ParseDate(string s, string which)
        {
            s = (s ?? "").Trim();
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d))
                return d;
            throw new OptionException($"Invalid {which} date '{s}'. Use YYYY-MM-DD.");
        }

        private static string Describe(DateTime? d) =>
            d == null ? "(open)" : d.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

        private static void PrintHelp()
        {
            Console.WriteLine(@"o2o-export — export Outlook mail to Obsidian Markdown notes.

USAGE:
  o2o-export --range <window>
  o2o-export --since-last-export

RANGE WINDOWS:
  day | week | month | year   rolling window ending now
  all                         the whole folder
  YYYY-MM-DD..YYYY-MM-DD       explicit dates (end inclusive); open end allowed

OPTIONS:
  --since-last-export   export everything newer than the last successful run
  --config <path>       config file (default: config.txt next to the exe)
  --state  <path>       state file  (default: state.txt next to the exe)
  --dry-run             list what would be written without touching disk
  --help                show this help

NOTES:
  * Requires classic Outlook for Windows, signed in, with macros/automation
    allowed by policy. Does NOT work with the 'new Outlook'.
  * --since-last-export tracks a high-water mark plus exported message ids, so
    re-running an overlapping range never creates duplicates.");
        }
    }

    internal sealed class OptionException : Exception
    {
        public OptionException(string msg) : base(msg) { }
    }

    internal sealed class Options
    {
        public string Range;
        public bool SinceLastExport;
        public string ConfigPath;
        public string StatePath;
        public bool DryRun;
        public bool ShowHelp;

        public static Options Parse(string[] args)
        {
            var o = new Options();
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                switch (a.ToLowerInvariant())
                {
                    case "--range": o.Range = Next(args, ref i, a); break;
                    case "--since-last-export": o.SinceLastExport = true; break;
                    case "--config": o.ConfigPath = Next(args, ref i, a); break;
                    case "--state": o.StatePath = Next(args, ref i, a); break;
                    case "--dry-run": o.DryRun = true; break;
                    case "--help":
                    case "-h":
                    case "/?": o.ShowHelp = true; break;
                    default: throw new OptionException("Unknown argument: " + a);
                }
            }

            if (!o.ShowHelp && o.Range == null && !o.SinceLastExport)
                o.ShowHelp = true; // no action requested -> show usage

            if (o.Range != null && o.SinceLastExport)
                throw new OptionException("Use either --range or --since-last-export, not both.");

            return o;
        }

        private static string Next(string[] args, ref int i, string flag)
        {
            if (i + 1 >= args.Length) throw new OptionException("Missing value for " + flag);
            return args[++i];
        }
    }
}
