using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace Wordplai.Cli
{
    internal static class Program
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 16777216 };
        private static readonly Dictionary<string, string> Local = new Dictionary<string, string> {
            { "documents", "documents" }, { "state", "state" }, { "context", "context_preview" },
            { "rule-save", "save_rule" }, { "mark", "mark_range" }, { "unmark", "remove_mark" },
            { "accept", "accept" }, { "reject", "reject" }, { "revoke", "revoke_context" } };
        private static readonly Dictionary<string, string> Connections = new Dictionary<string, string> {
            { "status", "connection_status" }, { "connect-local", "connection_local" },
            { "connect-public", "connection_start" }, { "stop", "connection_stop" } };

        [STAThread]
        public static int Main(string[] args)
        {
            Console.OutputEncoding = new UTF8Encoding(false);
            Console.InputEncoding = new UTF8Encoding(false);
            WordConnection connection = null;
            try
            {
                if (args.Length == 0 || args[0] == "help" || args[0] == "--help") { Help(); return 0; }
                string command = args[0];
                if (!Local.ContainsKey(command) && !Connections.ContainsKey(command) && command != "mcp")
                    throw new ArgumentException("Unknown command. Use wordplai.exe help.");
                Dictionary<string, string> options = Options(args.Skip(1).ToArray());
                var input = ReadInput(options);
                int pid = 0;
                if (options.ContainsKey("--word-pid") && (!Int32.TryParse(options["--word-pid"], out pid) || pid <= 0))
                    throw new ArgumentException("--word-pid must be a positive process ID.");
                string document = Option(options, "--document");
                if (Local.ContainsKey(command) && command != "documents" && String.IsNullOrWhiteSpace(document))
                    throw new ArgumentException("An explicit --document ID from documents is required.");
                if (input.ContainsKey("command") || (command != "mcp" && input.ContainsKey("documentId")))
                    throw new ArgumentException("Use command and --document arguments; input cannot override their target.");
                if (command == "mcp" && (Option(options, "--method") == null || Option(options, "--session-file") == null))
                    throw new ArgumentException("mcp requires --method and --session-file.");
                if (command != "mcp" && (options.ContainsKey("--method") || options.ContainsKey("--session-file")))
                    throw new ArgumentException("--method and --session-file are only valid for mcp.");
                using (connection = new WordConnection(pid))
                {
                    Dictionary<string, object> response;
                    if (command == "mcp") response = Mcp(connection, options, input);
                    else if (Connections.ContainsKey(command)) response = Execute(connection, Connections[command], input);
                    else
                    {
                        input["command"] = Local[command];
                        if (document != null) input["documentId"] = document;
                        response = Execute(connection, "automation", input);
                    }
                    response["wordProcessId"] = connection.ProcessId;
                    if (!options.ContainsKey("--include-secrets")) Redact(response);
                    Console.WriteLine(Json.Serialize(response));
                    return response.ContainsKey("ok") && (bool)response["ok"] ? 0 : 1;
                }
            }
            catch (ArgumentException error) { return Error(error.Message, 2, false); }
            catch (Exception error)
            {
                return Error(error is WebException ? "MCP transport failed. Inspect the original request before retrying." : error.Message,
                    3, connection != null && connection.RequestSent);
            }
        }

        private static int Error(string message, int code, bool uncertain)
        {
            Console.WriteLine(Json.Serialize(new { ok = false, error = message, outcome_unknown = uncertain }));
            return code;
        }

        private static Dictionary<string, string> Options(string[] args)
        {
            var result = new Dictionary<string, string>();
            var values = new[] { "--document", "--input", "--word-pid", "--method", "--session-file" };
            for (int index = 0; index < args.Length; index++)
            {
                string key = args[index];
                if (result.ContainsKey(key)) throw new ArgumentException("Duplicate option: " + key);
                if (key == "--include-secrets") result.Add(key, "true");
                else if (values.Contains(key) && index + 1 < args.Length && !args[index + 1].StartsWith("--", StringComparison.Ordinal))
                    result.Add(key, args[++index]);
                else throw new ArgumentException("Unknown option or missing value: " + key);
            }
            return result;
        }

        private static string Option(Dictionary<string, string> options, string name)
        { string value; return options.TryGetValue(name, out value) ? value : null; }

        private static Dictionary<string, object> ReadInput(Dictionary<string, string> options)
        {
            string path = Option(options, "--input");
            if (path == null) return new Dictionary<string, object>();
            string text;
            using (TextReader reader = path == "-" ? (TextReader)Console.In : new StreamReader(path, Encoding.UTF8, true))
            {
                var buffer = new char[131073];
                int used = 0, count;
                while (used < buffer.Length && (count = reader.Read(buffer, used, buffer.Length - used)) > 0) used += count;
                if (used > 131072) throw new ArgumentException("Input exceeds 131072 characters.");
                text = new string(buffer, 0, used);
            }
            object value;
            try { value = Json.DeserializeObject(text); }
            catch (ArgumentException) { throw new ArgumentException("Input must be a UTF-8 JSON object."); }
            var input = value as Dictionary<string, object>;
            if (input == null) throw new ArgumentException("Input must be a UTF-8 JSON object.");
            return input;
        }

        private static Dictionary<string, object> Execute(WordConnection connection, string operation, Dictionary<string, object> input)
        { return (Dictionary<string, object>)Json.DeserializeObject(connection.Execute(operation, Json.Serialize(input))); }

        private static Dictionary<string, object> Mcp(WordConnection connection, Dictionary<string, string> options, Dictionary<string, object> input)
        {
            string method = options["--method"], file = Path.GetFullPath(options["--session-file"]);
            if (!new[] { "initialize", "ping", "tools/list", "tools/call", "resources/list", "resources/read" }.Contains(method))
                throw new ArgumentException("Unsupported MCP method.");
            var status = Execute(connection, "connection_status", new Dictionary<string, object>());
            if (!(bool)status["ok"]) return status;
            var state = (Dictionary<string, object>)status["result"];
            Uri endpoint;
            if (!(bool)state["running"] || !Uri.TryCreate((string)state["localMcpUrl"], UriKind.Absolute, out endpoint) ||
                endpoint.Scheme != "http" || !endpoint.IsLoopback)
                throw new InvalidOperationException("Start local access explicitly with connect-local before using mcp.");
            string digest;
            using (var sha = SHA256.Create()) digest = Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(endpoint.AbsoluteUri)));
            string session = null;
            if (method == "initialize")
            {
                if (File.Exists(file)) throw new ArgumentException("Choose a new session file; initialization never overwrites an existing session.");
                if (input.Count == 0) input = new Dictionary<string, object> { { "protocolVersion", "2025-03-26" },
                    { "capabilities", new Dictionary<string, object>() }, { "clientInfo", new { name = "wordplai CLI", version = "1" } } };
            }
            else
            {
                var saved = (Dictionary<string, object>)Json.DeserializeObject(File.ReadAllText(file, Encoding.UTF8));
                if ((string)saved["endpointDigest"] != digest) throw new InvalidOperationException("Connection changed. Retain the old session and initialize a new session file.");
                session = (string)saved["session"];
            }
            var request = (HttpWebRequest)WebRequest.Create(endpoint);
            request.Method = "POST"; request.ContentType = "application/json"; request.Accept = "application/json, text/event-stream";
            request.Timeout = 20000; request.ReadWriteTimeout = 20000; request.KeepAlive = false;
            request.AllowAutoRedirect = false; request.Proxy = null;
            if (session != null) request.Headers["Mcp-Session-Id"] = session;
            byte[] bytes = Encoding.UTF8.GetBytes(Json.Serialize(new { jsonrpc = "2.0", id = 1, method = method, @params = input }));
            request.ContentLength = bytes.Length;
            using (Stream stream = request.GetRequestStream()) stream.Write(bytes, 0, bytes.Length);
            using (var response = (HttpWebResponse)request.GetResponse())
            using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
            {
                Dictionary<string, object> reply;
                try { reply = Json.DeserializeObject(reader.ReadToEnd()) as Dictionary<string, object>; }
                catch (ArgumentException) { throw new InvalidOperationException("MCP returned invalid JSON. Inspect the original request before retrying."); }
                if (reply == null) throw new InvalidOperationException("MCP returned no result object. Inspect the original request before retrying.");
                if (method == "initialize" && !reply.ContainsKey("error"))
                {
                    session = response.Headers["Mcp-Session-Id"];
                    if (String.IsNullOrEmpty(session)) throw new InvalidOperationException("MCP initialization returned no session identity.");
                    using (var stream = new FileStream(file, FileMode.CreateNew, FileAccess.Write))
                    using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                        writer.Write(Json.Serialize(new { endpointDigest = digest, session = session }));
                }
                bool ok = !reply.ContainsKey("error");
                object result;
                if (reply.TryGetValue("result", out result))
                {
                    var tool = result as Dictionary<string, object>;
                    if (tool != null && tool.ContainsKey("isError") && (bool)tool["isError"]) ok = false;
                }
                return new Dictionary<string, object> { { "ok", ok }, { "response", reply } };
            }
        }

        private static void Redact(object value)
        {
            var dictionary = value as Dictionary<string, object>;
            if (dictionary != null)
                foreach (string key in dictionary.Keys.ToArray())
                {
                    if (key == "localMcpUrl" || key == "publicMcpUrl") dictionary[key] = "[redacted]";
                    else Redact(dictionary[key]);
                }
            var array = value as object[];
            if (array != null) foreach (object child in array) Redact(child);
        }

        private static void Help()
        {
            Console.WriteLine("wordplai <command> [--document ID] [--input FILE|-] [--word-pid PID] [--include-secrets]");
            Console.WriteLine("Commands: documents, state, context, rule-save, mark, unmark, accept, reject, revoke,");
            Console.WriteLine("          status, connect-local, connect-public, stop, mcp");
            Console.WriteLine("mcp also requires --method METHOD --session-file FILE. Each file retains one MCP session.");
            Console.WriteLine("Document commands require an exact ID. Mutations require the current snapshot in JSON input.");
            Console.WriteLine("Word must already be open with the add-in connected. No Word launch, mouse, keyboard or focus injection.");
            Console.WriteLine("JSON output; exit 0 success, 1 add-in/MCP refusal, 2 usage/input, 3 runtime/transport error.");
        }
    }
}
