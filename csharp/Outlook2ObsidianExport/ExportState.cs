using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Outlook2Obsidian.Export
{
    /// <summary>
    /// Persisted memory between runs. Stores the high-water mark used by
    /// --since-last-export plus the set of already-exported message ids so a
    /// re-run over an overlapping date range never produces duplicates.
    ///
    /// Format (plain text, no JSON dependency):
    ///   LastExport=2026-06-27T10:00:00.0000000+02:00
    ///   &lt;internet-message-id-1&gt;
    ///   &lt;internet-message-id-2&gt;
    ///   ...
    /// </summary>
    public sealed class ExportState
    {
        /// <summary>Local time of the newest message exported so far, or null if never run.</summary>
        public DateTime? LastExport;

        private readonly HashSet<string> _exportedIds =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public bool AlreadyExported(string messageId) =>
            !string.IsNullOrEmpty(messageId) && _exportedIds.Contains(messageId);

        public void MarkExported(string messageId, DateTime receivedLocal)
        {
            if (!string.IsNullOrEmpty(messageId))
                _exportedIds.Add(messageId);
            if (LastExport == null || receivedLocal > LastExport.Value)
                LastExport = receivedLocal;
        }

        public static ExportState Load(string path)
        {
            var state = new ExportState();
            if (!File.Exists(path)) return state;

            foreach (var raw in File.ReadAllLines(path))
            {
                var line = raw.Trim();
                if (line.Length == 0) continue;
                if (line.StartsWith("LastExport=", StringComparison.OrdinalIgnoreCase))
                {
                    string v = line.Substring("LastExport=".Length).Trim();
                    if (DateTime.TryParse(v, CultureInfo.InvariantCulture,
                            DateTimeStyles.RoundtripKind, out var dt))
                        state.LastExport = dt;
                }
                else
                {
                    state._exportedIds.Add(line);
                }
            }
            return state;
        }

        public void Save(string path)
        {
            var lines = new List<string>();
            if (LastExport != null)
                lines.Add("LastExport=" + LastExport.Value.ToString("o", CultureInfo.InvariantCulture));
            lines.AddRange(_exportedIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));

            // Write atomically: a crash mid-write must not corrupt the watermark.
            string tmp = path + ".tmp";
            File.WriteAllLines(tmp, lines);
            if (File.Exists(path)) File.Delete(path);
            File.Move(tmp, path);
        }
    }
}
