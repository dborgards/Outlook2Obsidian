using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Outlook2Obsidian.Export
{
    /// <summary>
    /// Filename sanitising, display-name formatting and YAML escaping.
    /// The formatName logic is a port of the original VBA Utilities module,
    /// hardened against the injection issues found in the security review.
    /// </summary>
    public static class NameFormatter
    {
        // ---- YAML -----------------------------------------------------------

        /// <summary>
        /// Escape an arbitrary string for use inside a double-quoted YAML scalar.
        /// Fixes the frontmatter-injection finding: a sender/recipient display
        /// name containing a quote, backslash or newline can no longer break out
        /// of the value or inject new keys.
        /// </summary>
        public static string YamlQuote(string s)
        {
            s = s ?? "";
            var sb = new StringBuilder(s.Length + 2);
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\r': sb.Append(' '); break;   // keep scalars single-line
                    case '\n': sb.Append(' '); break;
                    case '\t': sb.Append("\\t"); break;
                    default: sb.Append(c); break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        // ---- File names -----------------------------------------------------

        private static readonly char[] InvalidNameChars =
            "/\\:?\"<>|[]*".ToCharArray();

        private static readonly Regex ReservedDevice = new Regex(
            @"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Strip characters that are illegal in a Windows file name (and control
        /// chars), collapse whitespace, and guard against reserved device names.
        /// Path separators are removed, so a crafted subject cannot traverse out
        /// of the vault folder.
        /// </summary>
        public static string SanitizeFileName(string subject)
        {
            subject = subject ?? "";
            var sb = new StringBuilder(subject.Length);
            foreach (char c in subject)
            {
                if (c < 32) continue;                       // control chars
                if (InvalidNameChars.Contains(c)) continue; // illegal on Windows
                sb.Append(c);
            }

            string cleaned = sb.ToString().Trim();
            // Windows ignores trailing dots/spaces; strip them to avoid surprises.
            cleaned = cleaned.TrimEnd('.', ' ');
            if (cleaned.Length == 0) cleaned = "untitled";
            if (ReservedDevice.IsMatch(cleaned)) cleaned = "_" + cleaned;

            // Keep the whole path comfortably under MAX_PATH.
            if (cleaned.Length > 150) cleaned = cleaned.Substring(0, 150).TrimEnd();
            return cleaned;
        }

        // ---- Display-name formatting (ported from VBA formatName) -----------

        private static readonly Regex RxFirstLast =
            new Regex(@"^\w+\s\w+$", RegexOptions.Compiled);
        private static readonly Regex RxLastFirstDomain =
            new Regex(@"^\w+,\s\w+@\w+(\.\w+)+", RegexOptions.Compiled);
        private static readonly Regex RxPlainEmail =
            new Regex(@"^\w+@\w+\.\w+", RegexOptions.Compiled);
        private static readonly Regex RxLastFirstAgency =
            new Regex(@"^[a-zA-Z_\-]+,\s[a-zA-Z_\-]+\s\(\w+\)$", RegexOptions.Compiled);
        private static readonly Regex RxSingleNameAgency =
            new Regex(@"^([a-zA-Z_\-]+\s)+\([a-zA-Z_\-]+\)", RegexOptions.Compiled);

        /// <summary>
        /// Turn an Outlook display name into an Obsidian [[wikilink]], mirroring
        /// the original VBA heuristics for the common Outlook name shapes.
        /// </summary>
        public static string FormatName(string name, string startChar)
        {
            name = (name ?? "").Trim();
            startChar = startChar ?? "";

            // "Doe, John@domain.or.gov"
            if (RxLastFirstDomain.IsMatch(name))
            {
                int comma = name.IndexOf(", ", StringComparison.Ordinal);
                int at = name.IndexOf('@');
                if (comma >= 0 && at > comma + 2)
                {
                    string fname = name.Substring(comma + 2, at - (comma + 2));
                    string lname = name.Substring(0, name.IndexOf(','));
                    return "[[" + startChar + fname + " " + lname + "]]";
                }
            }

            // "Doe, John (Agency)"
            if (RxLastFirstAgency.IsMatch(name))
            {
                int comma = name.IndexOf(", ", StringComparison.Ordinal);
                int paren = name.IndexOf(" (", StringComparison.Ordinal);
                if (comma >= 0 && paren > comma + 2)
                {
                    string fname = name.Substring(comma + 2, paren - (comma + 2));
                    string lname = name.Substring(0, name.IndexOf(','));
                    return "[[" + startChar + fname + " " + lname + "]]";
                }
            }

            // "JohnDoe@gmail.com"
            if (RxPlainEmail.IsMatch(name))
            {
                int at = name.IndexOf('@');
                return "[[" + startChar + name.Substring(0, at) + "]]";
            }

            // "Payroll (Agency)" — single/distribution-list name, no startChar.
            if (RxSingleNameAgency.IsMatch(name))
            {
                int paren = name.IndexOf(" (", StringComparison.Ordinal);
                if (paren > 0) return "[[" + name.Substring(0, paren) + "]]";
            }

            // "John Doe"
            if (RxFirstLast.IsMatch(name))
                return "[[" + startChar + name + "]]";

            // Anything else: link verbatim.
            return "[[" + name + "]]";
        }
    }
}
