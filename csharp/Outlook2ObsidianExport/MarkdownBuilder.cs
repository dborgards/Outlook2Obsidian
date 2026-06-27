using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Outlook2Obsidian.Export
{
    /// <summary>Plain data for one email, decoupled from the COM objects.</summary>
    public sealed class MailRecord
    {
        public string Subject = "";
        public string SenderName = "";
        public DateTime ReceivedLocal;
        public List<string> Recipients = new List<string>();
        public string Body = "";
        public List<string> AttachmentLinks = new List<string>();
    }

    /// <summary>
    /// Builds the Markdown note (YAML frontmatter + body), matching the layout of
    /// the original VBA tool but with all untrusted fields safely escaped.
    /// </summary>
    public static class MarkdownBuilder
    {
        public static string BuildNote(MailRecord mail, Config cfg)
        {
            var sb = new StringBuilder();

            // ---- YAML frontmatter (all dynamic values are YAML-escaped) ----
            sb.Append("---\n");
            sb.Append("class: email\n");
            sb.Append("area: \n");
            sb.Append("project: \n");
            sb.Append("title: ").Append(NameFormatter.YamlQuote(mail.Subject)).Append('\n');
            sb.Append("date: ")
              .Append(mail.ReceivedLocal.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture))
              .Append('\n');
            sb.Append("from: ")
              .Append(NameFormatter.YamlQuote(NameFormatter.FormatName(mail.SenderName, cfg.PersonNameStartChar)))
              .Append('\n');

            sb.Append("to:\n");
            foreach (var r in mail.Recipients)
            {
                sb.Append("  - ")
                  .Append(NameFormatter.YamlQuote(NameFormatter.FormatName(r, cfg.PersonNameStartChar)))
                  .Append('\n');
            }

            sb.Append("tags:\n");
            sb.Append("related:\n");

            sb.Append("attachments: [");
            for (int i = 0; i < mail.AttachmentLinks.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append(NameFormatter.YamlQuote(mail.AttachmentLinks[i]));
            }
            sb.Append("]\n");
            sb.Append("---\n\n");

            // ---- Task + separator ----
            sb.Append("- [ ] ").Append(OneLine(mail.Subject)).Append("\n\n");
            sb.Append("---\n\n");

            // ---- Body ----
            sb.Append(BuildBody(mail.Body, cfg));

            return sb.ToString();
        }

        private static string BuildBody(string body, Config cfg)
        {
            body = body ?? "";
            if (!cfg.WrapBodyInCodeBlock)
                return body;

            // Wrap the (untrusted) body in a fence long enough that any backtick
            // run inside it cannot terminate the block early. This neutralises
            // Dataview/Templater code blocks so they render as text instead of
            // executing when Obsidian opens the note.
            int longest = 0, run = 0;
            foreach (char c in body)
            {
                if (c == '`') { run++; if (run > longest) longest = run; }
                else run = 0;
            }
            string fence = new string('`', Math.Max(3, longest + 1));
            return fence + "text\n" + body + "\n" + fence + "\n";
        }

        private static string OneLine(string s) =>
            (s ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
    }
}
