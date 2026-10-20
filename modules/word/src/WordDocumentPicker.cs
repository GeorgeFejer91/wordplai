using System;
using System.Collections.Generic;
using System.IO;

namespace UnderTree.WordAddin
{
    internal static class WordDocumentPicker
    {
        internal const string Uri = "ui://wordplai/document-picker-v2.html";
        internal const string SessionKey = "\u0000wordplai-session";
        private const string Mime = "text/html;profile=mcp-app";

        internal static object OutputSchema()
        {
            object text = new { type = "string" };
            object nullableText = new { type = new[] { "string", "null" } };
            object choice = new { type = "object", properties = new {
                manuscript_id = text, documentId = text, title = text, active = new { type = "boolean" }
            }, required = new[] { "manuscript_id", "documentId", "title", "active" }, additionalProperties = false };
            return new { type = "object", oneOf = new object[] {
                new { type = "object", properties = new {
                    selection_required = new { type = "boolean" },
                    manuscripts = new { type = "array", items = choice },
                    selected_manuscript_id = nullableText, title = nullableText
                }, required = new[] { "selection_required", "manuscripts", "selected_manuscript_id", "title" }, additionalProperties = false },
                new { type = "object", properties = new {
                    error = new { type = "object", properties = new {
                        code = text, message = text,
                        details = new { description = "Operation-specific refusal details, or null when no details are supplied." }
                    }, required = new[] { "code", "message", "details" }, additionalProperties = false }
                }, required = new[] { "error" }, additionalProperties = false }
            } };
        }

        internal static object ToolMetadata()
        {
            return new Dictionary<string, object> {
                { "ui", new { resourceUri = Uri } },
                { "openai/outputTemplate", Uri },
                { "openai/widgetAccessible", true },
                { "openai/toolInvocation/invoking", "Opening Word documents" },
                { "openai/toolInvocation/invoked", "Word document ready" }
            };
        }

        internal static object List()
        {
            return new { resources = new[] { new { uri = Uri, name = "Choose a Word document", mimeType = Mime,
                description = "Choose which open Word document to edit. The chosen document remains the writing target." } } };
        }

        internal static object Read(IDictionary<string, object> parameters)
        {
            foreach (string key in parameters.Keys)
                if (key != "uri" && key != "_meta") throw new McpOperationException("Unknown Word UI resource.");
            object supplied;
            if (!parameters.TryGetValue("uri", out supplied) || !Object.Equals(supplied, Uri) ||
                (parameters.TryGetValue("_meta", out supplied) && !(supplied is IDictionary<string, object>)))
                throw new McpOperationException("Unknown Word UI resource.");
            using (Stream stream = typeof(WordDocumentPicker).Assembly.GetManifestResourceStream("wordplai.DocumentPicker"))
            {
                if (stream == null) throw new McpOperationException("The document picker is unavailable; choose a document by name in chat.");
                using (var reader = new StreamReader(stream))
                    return new { contents = new[] { new { uri = Uri, mimeType = Mime, text = reader.ReadToEnd(),
                        _meta = new Dictionary<string, object> {
                            { "openai/widgetDescription", "Choose one Word document, then continue the existing writing task. Do not choose on the user's behalf." },
                            { "ui", new { prefersBorder = true, csp = new { connectDomains = new string[0], resourceDomains = new string[0] } } },
                            { "openai/widgetPrefersBorder", true },
                            { "openai/widgetCSP", new { connect_domains = new string[0], resource_domains = new string[0] } }
                        } } } };
            }
        }
    }
}
