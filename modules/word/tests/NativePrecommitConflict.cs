using System;
using System.IO;
using UnderTree.WordAddin;
internal static partial class NativeHighlightingSmoke
{
    private static void RequestedColors(dynamic app, WordEditor editor, string address, string output)
    {
        dynamic document=app.ActiveDocument;
        string original=WordSnapshot.Capture(document).MainText;
        string id=UiState.Text(UiState.Map(Local(editor,"state")["document"]),"id");
        var context=Share(editor,address,id);
        var response=Rpc(address,"highlight_text",new{documentId=id,marking=new{requestId="precommit_conflict",revision=context["revision"],contextId=context["context_id"],ranges=new[]{new{from=0,to=3,quote=original.Substring(0,3)}}}});
        Check(UiState.Flag(response,"isError"),"A real native precommit text change rejects highlighting.");
        Check(WordSnapshot.Capture(document).MainText=="User change. "+original,"Conflict rejection never undoes the intervening native user edit.");
        var state=Local(editor,"state");
        Check(UiState.Records(state,"marks").Count==0&&UiState.Records(state,"rules").Count==4,"A precommit conflict leaves no staged color or mark.");
        Check(!UiState.Flag(UiState.Map(state["document"]),"shared"),"A precommit conflict revokes the stale witness.");
        File.WriteAllText(Path.Combine(output,"conflict-result.txt"),"PASS "+assertions+" native precommit conflict assertions.");
    }
}
