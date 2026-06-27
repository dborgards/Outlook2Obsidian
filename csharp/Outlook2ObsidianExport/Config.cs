using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace Outlook2Obsidian.Export
{
    /// <summary>
    /// Plain key=value configuration so there is no dependency on a JSON library
    /// (keeps the tool buildable on a locked-down corporate box with no feed).
    /// Lines starting with '#' are comments; blank lines are ignored.
    /// </summary>
    public sealed class Config
    {
        /// <summary>Absolute folder the .md notes are written to. Trailing slash optional.</summary>
        public string VaultPath = @"C:\Users\YourUsername\Obsidian\Vault\Emails\";

        /// <summary>Prefix added before names in [[wikilinks]], e.g. "@".</summary>
        public string PersonNameStartChar = "@";

        /// <summary>Prefix for the generated file name, e.g. "Email_".</summary>
        public string FileNamePrefix = "Email_";

        /// <summary>
        /// Which mail folder to export. "Inbox" / "SentMail" / "DeletedItems" map to
        /// the well-known default folders; anything else is treated as a folder path
        /// like "Inbox\Projects\ACME" resolved from the default store root.
        /// </summary>
        public string FolderPath = "Inbox";

        /// <summary>Recurse into subfolders of <see cref="FolderPath"/>.</summary>
        public bool Recurse = false;

        /// <summary>Save attachments next to the note and link them in the frontmatter.</summary>
        public bool IncludeAttachments = false;

        /// <summary>
        /// Wrap the email body in a fenced code block so any Dataview/Templater code
        /// inside a (potentially hostile) message cannot execute when Obsidian opens
        /// the note. Off by default for readability; turn on if you value safety over
        /// rendered formatting.
        /// </summary>
        public bool WrapBodyInCodeBlock = false;

        public static Config Load(string path)
        {
            var cfg = new Config();
            if (!File.Exists(path))
                throw new FileNotFoundException(
                    "Config file not found. Copy config.example.txt to config.txt and edit it.", path);

            foreach (var raw in File.ReadAllLines(path))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#")) continue;
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                string key = line.Substring(0, eq).Trim();
                string val = line.Substring(eq + 1).Trim();

                switch (key.ToLowerInvariant())
                {
                    case "vaultpath": cfg.VaultPath = val; break;
                    case "personnamestartchar": cfg.PersonNameStartChar = val; break;
                    case "filenameprefix": cfg.FileNamePrefix = val; break;
                    case "folderpath": cfg.FolderPath = val; break;
                    case "recurse": cfg.Recurse = ParseBool(val); break;
                    case "includeattachments": cfg.IncludeAttachments = ParseBool(val); break;
                    case "wrapbodyincodeblock": cfg.WrapBodyInCodeBlock = ParseBool(val); break;
                    // Unknown keys are ignored so old config files keep working.
                }
            }

            if (string.IsNullOrWhiteSpace(cfg.VaultPath))
                throw new InvalidOperationException("VaultPath must be set in the config file.");

            // Require an absolute path. A relative path would silently write notes
            // next to the exe instead of into the vault.
            if (!Path.IsPathRooted(cfg.VaultPath))
                throw new InvalidOperationException(
                    "VaultPath must be an absolute path, e.g. C:\\Users\\You\\Obsidian\\Vault\\Emails\\.");

            // Normalise to a trailing separator so path joins are predictable.
            if (!cfg.VaultPath.EndsWith("\\") && !cfg.VaultPath.EndsWith("/"))
                cfg.VaultPath += "\\";

            return cfg;
        }

        private static bool ParseBool(string v)
        {
            v = (v ?? "").Trim().ToLowerInvariant();
            return v == "true" || v == "1" || v == "yes" || v == "y";
        }
    }
}
