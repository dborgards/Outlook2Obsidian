using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook2Obsidian.Export
{
    /// <summary>
    /// Drives the locally installed (classic) Outlook via COM to export mail
    /// items in a date range to Markdown notes. No Azure / Graph involved — it
    /// uses the Outlook session the user is already signed into.
    /// </summary>
    public sealed class OutlookExporter : IDisposable
    {
        private const string PrInternetMessageId =
            "http://schemas.microsoft.com/mapi/proptag/0x1035001F";

        private readonly Config _cfg;
        private readonly ExportState _state;
        private readonly Action<string> _log;

        private Outlook.Application _app;
        private Outlook.NameSpace _ns;

        public OutlookExporter(Config cfg, ExportState state, Action<string> log)
        {
            _cfg = cfg;
            _state = state;
            _log = log ?? (_ => { });
            _app = new Outlook.Application();          // attaches to running Outlook
            _ns = _app.GetNamespace("MAPI");
        }

        /// <summary>
        /// Export everything received in [start, end). A null bound means open-ended.
        /// Returns the number of notes written.
        /// </summary>
        public int Export(DateTime? startLocal, DateTime? endLocal, bool dryRun)
        {
            var root = ResolveFolder(_cfg.FolderPath);
            int written = 0;
            foreach (var folder in EnumerateFolders(root, _cfg.Recurse))
            {
                written += ExportFolder(folder, startLocal, endLocal, dryRun);
                if (!ReferenceEquals(folder, root)) Release(folder);
            }
            Release(root);
            return written;
        }

        private int ExportFolder(Outlook.MAPIFolder folder, DateTime? startLocal,
                                 DateTime? endLocal, bool dryRun)
        {
            Outlook.Items items = folder.Items;
            string filter = BuildDaslFilter(startLocal, endLocal);
            Outlook.Items scope = string.IsNullOrEmpty(filter) ? items : items.Restrict(filter);
            scope.Sort("[ReceivedTime]", false); // oldest first => watermark advances safely

            _log($"Folder '{folder.Name}': scanning {scope.Count} item(s)...");

            int written = 0;
            foreach (object obj in scope)
            {
                var mail = obj as Outlook.MailItem;
                if (mail == null) { Release(obj); continue; } // skip meetings/contacts/etc.

                try
                {
                    string messageId = GetInternetMessageId(mail) ?? mail.EntryID;
                    if (_state.AlreadyExported(messageId))
                        continue; // dedup across overlapping ranges

                    var record = ToRecord(mail, dryRun);
                    string path = WriteNote(record, dryRun);
                    _state.MarkExported(messageId, record.ReceivedLocal);
                    written++;
                    _log((dryRun ? "[dry-run] " : "") + "Saved: " + Path.GetFileName(path));
                }
                catch (Exception ex)
                {
                    _log("  ! skipped one item: " + ex.Message);
                }
                finally
                {
                    Release(mail);
                }
            }

            if (!string.IsNullOrEmpty(filter)) Release(scope);
            Release(items);
            return written;
        }

        // ---- Mapping -------------------------------------------------------

        private MailRecord ToRecord(Outlook.MailItem mail, bool dryRun)
        {
            var rec = new MailRecord
            {
                Subject = mail.Subject ?? "",
                SenderName = mail.SenderName ?? "",
                ReceivedLocal = mail.ReceivedTime,
                Body = mail.Body ?? ""
            };

            Outlook.Recipients recips = mail.Recipients;
            try
            {
                for (int i = 1; i <= recips.Count; i++)
                {
                    Outlook.Recipient r = recips[i];
                    try { rec.Recipients.Add(r.Name ?? ""); }
                    finally { Release(r); }
                }
            }
            finally { Release(recips); }

            if (_cfg.IncludeAttachments)
                SaveAttachments(mail, rec, dryRun);

            return rec;
        }

        private void SaveAttachments(Outlook.MailItem mail, MailRecord rec, bool dryRun)
        {
            Outlook.Attachments atts = mail.Attachments;
            try
            {
                if (atts.Count == 0) return;
                string dir = Path.Combine(_cfg.VaultPath, "attachments");
                if (!dryRun) Directory.CreateDirectory(dir);

                string baseName = _cfg.FileNamePrefix +
                    rec.ReceivedLocal.ToString("yyyy-MM-dd HHmm", CultureInfo.InvariantCulture);

                for (int i = 1; i <= atts.Count; i++)
                {
                    Outlook.Attachment a = atts[i];
                    try
                    {
                        string safe = NameFormatter.SanitizeFileName(a.FileName ?? ("attachment" + i));
                        string fileName = baseName + " " + safe;
                        string full = Path.Combine(dir, fileName);
                        full = EnsureUnique(full);
                        if (!dryRun) a.SaveAsFile(full);
                        rec.AttachmentLinks.Add("attachments/" + Path.GetFileName(full));
                    }
                    finally { Release(a); }
                }
            }
            finally { Release(atts); }
        }

        // ---- File output ---------------------------------------------------

        private string WriteNote(MailRecord rec, bool dryRun)
        {
            string baseFile = _cfg.FileNamePrefix +
                rec.ReceivedLocal.ToString("yyyy-MM-dd HHmm", CultureInfo.InvariantCulture) +
                " " + NameFormatter.SanitizeFileName(rec.Subject) + ".md";

            string path = EnsureUnique(Path.Combine(_cfg.VaultPath, baseFile));
            if (!dryRun)
            {
                Directory.CreateDirectory(_cfg.VaultPath);
                // UTF-8 without BOM, matching the original tool's SaveAsUTF8 intent.
                File.WriteAllText(path, MarkdownBuilder.BuildNote(rec, _cfg), new UTF8Encoding(false));
            }
            return path;
        }

        /// <summary>Append " (n)" until the path does not collide with an existing file.</summary>
        private static string EnsureUnique(string path)
        {
            if (!File.Exists(path)) return path;
            string dir = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);
            for (int n = 2; ; n++)
            {
                string candidate = Path.Combine(dir, $"{name} ({n}){ext}");
                if (!File.Exists(candidate)) return candidate;
            }
        }

        // ---- Outlook plumbing ----------------------------------------------

        private static string BuildDaslFilter(DateTime? startLocal, DateTime? endLocal)
        {
            // DASL @SQL with an explicit yyyy-MM-dd HH:mm string is locale-independent,
            // unlike the [ReceivedTime] >= '...' Jet syntax which depends on regional
            // date formats.
            const string prop = "\"urn:schemas:httpmail:datereceived\"";
            var parts = new List<string>();
            if (startLocal != null)
                parts.Add(prop + " >= '" +
                    startLocal.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + "'");
            if (endLocal != null)
                parts.Add(prop + " < '" +
                    endLocal.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + "'");
            return parts.Count == 0 ? null : "@SQL=" + string.Join(" AND ", parts);
        }

        private static string GetInternetMessageId(Outlook.MailItem mail)
        {
            try { return (string)mail.PropertyAccessor.GetProperty(PrInternetMessageId); }
            catch { return null; }
        }

        private Outlook.MAPIFolder ResolveFolder(string spec)
        {
            switch ((spec ?? "").Trim().ToLowerInvariant())
            {
                case "":
                case "inbox": return _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                case "sent":
                case "sentmail": return _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                case "drafts": return _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
                case "outbox": return _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
                case "deleteditems": return _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                case "junk":
                case "junkemail": return _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);
            }

            // Treat anything else as a backslash path under the default store root,
            // e.g. "Inbox\Projects\ACME".
            Outlook.MAPIFolder current = _ns.DefaultStore.GetRootFolder();
            foreach (string part in spec.Split('\\'))
            {
                if (part.Length == 0) continue;
                Outlook.MAPIFolder next = null;
                Outlook.Folders children = current.Folders;
                try
                {
                    foreach (Outlook.MAPIFolder child in children)
                    {
                        if (string.Equals(child.Name, part, StringComparison.OrdinalIgnoreCase))
                        { next = child; break; }
                        Release(child);
                    }
                }
                finally { Release(children); }
                if (next == null)
                    throw new InvalidOperationException("Folder not found: '" + spec + "' (at '" + part + "')");
                if (!ReferenceEquals(current, next)) Release(current);
                current = next;
            }
            return current;
        }

        private IEnumerable<Outlook.MAPIFolder> EnumerateFolders(Outlook.MAPIFolder root, bool recurse)
        {
            yield return root;
            if (!recurse) yield break;
            Outlook.Folders children = root.Folders;
            try
            {
                foreach (Outlook.MAPIFolder child in children)
                    foreach (var f in EnumerateFolders(child, true))
                        yield return f;
            }
            finally { Release(children); }
        }

        private static void Release(object com)
        {
            if (com != null && Marshal.IsComObject(com))
                Marshal.ReleaseComObject(com);
        }

        public void Dispose()
        {
            Release(_ns); _ns = null;
            Release(_app); _app = null;
        }
    }
}
