using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using UnderTree.WordAddin;

internal static class PolicyStateJson
{
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 16777216 };

    internal static string Serialize(PolicyState state)
    {
        return Canonical(Json.DeserializeObject(Json.Serialize(state)));
    }

    internal static string Canonical(object value)
    {
        var fields = value as IDictionary<string, object>;
        if (fields != null)
            return "{" + String.Join(",", fields.OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .Select(pair => Json.Serialize(pair.Key) + ":" + Canonical(pair.Value))) + "}";
        var items = value as object[];
        if (items != null) return "[" + String.Join(",", items.Select(Canonical)) + "]";
        return Json.Serialize(value);
    }
}
