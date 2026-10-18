using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using UnderTree.WordAddin;

internal static class NativeDoubleClickSmoke
{
    private delegate bool Enumerate(IntPtr window,IntPtr state);
    private delegate void DoubleClick(object selection,ref bool cancel);
    private delegate void AfterDoubleClick(object selection,ref bool cancel);
    [StructLayout(LayoutKind.Sequential)] private struct Point { public int X,Y; }
    [DllImport("user32.dll")] private static extern bool EnumWindows(Enumerate callback,IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent,Enumerate callback,IntPtr state);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("user32.dll")] private static extern int GetClassName(IntPtr window,StringBuilder name,int count);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window,uint id,ref Guid iid,[MarshalAs(UnmanagedType.IDispatch)] out object value);
    [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr window,ref Point point);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window,uint message,IntPtr wparam,IntPtr lparam);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window,uint command);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window,uint kind);
    private static readonly JavaScriptSerializer Json=new JavaScriptSerializer{MaxJsonLength=16777216};
    private static readonly List<string> Checks=new List<string>();
    private static readonly List<object> Events=new List<object>();
    private static readonly List<Dictionary<string,object>> Dialogs=new List<Dictionary<string,object>>();
    private static readonly HashSet<Form> Seen=new HashSet<Form>();
    private static readonly List<bool> Cancellations=new List<bool>();
    private static readonly List<string> DocumentIds=new List<string>();
    private static readonly List<string> ThreadErrors=new List<string>();
    private static dynamic app,doc;
    private static WordEditor editor;
    private static ComAddin addin;
    private static Control dispatcher;
    private static IntPtr textWindow;
    private static string output,documentId,documentFile,mode;
    private static Delegate production;
    private static DoubleClick observer;
    private static AfterDoubleClick completed;
    private static int eventCount,openedCount;
    private static int ownedPid;
    private static bool inEvent,rebind,chooseOverlap,choiceMade,choiceChildSeen,keepRuleWindows;
    private static bool switchDocument,closeDocument,disconnect;
    private static dynamic otherDocument;
    private static string failure;
    private static readonly Guid EventsId=new Guid("00020A01-0000-0000-C000-000000000046");
    [STAThread] public static int Main(string[] args)
    {
        Process owned=null; WordTestAutomation filter=null;
        try {
            mode=args.Length>2?args[2]:"navigation-only";
            if(mode!="navigation-only"&&mode!="real-input")throw new ArgumentException("Choose navigation-only or explicitly requested real-input.");
            SetProcessDPIAware(); output=Path.GetFullPath(args[1]);Directory.CreateDirectory(output);filter=new WordTestAutomation();
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException+=delegate(object sender,ThreadExceptionEventArgs error){ThreadErrors.Add(error.Exception.ToString());};
            string bootstrap=Path.Combine(output,"isolated-double-click.rtf");File.WriteAllText(bootstrap,@"{\rtf1\ansi Disposable event bootstrap.\par}");
            owned=Process.Start(new ProcessStartInfo(args[0],"/a /x /q \""+bootstrap+"\""){UseShellExecute=false,WindowStyle=ProcessWindowStyle.Hidden});
            app=AttachOwned(owned);
            Check((int)app.Documents.Count==1&&(string)app.ActiveDocument.FullName==bootstrap,"The exact owned instance contains only the synthetic bootstrap.");
            app.DisplayAlerts=0;app.ActiveDocument.Close(0);doc=app.Documents.Add();doc.Content.Text="Outside text before.\rA green citation remains unchanged.\rOutside text after.\r";
            documentFile=Path.Combine(output,"event-fixture.docx");doc.SaveAs2(documentFile,16);
            editor=new WordEditor((object)app,operation=>operation(),Path.Combine(output,"profile"),Path.Combine(output,"unused-zrok.exe"));
            addin=new ComAddin();dispatcher=new Control();IntPtr ignored=dispatcher.Handle;
            Set("editor",editor);Set("dispatcher",dispatcher);Set("wordApplication",(object)app);Set("ownerThread",Thread.CurrentThread.ManagedThreadId);
            var state=Local("state");documentId=Text(Map(state["document"]),"id");DocumentIds.Add(documentId);
            Local("save_rule",Values("id","event-green","label","Native green citation","color","#33AA66","action","protect","accessMode","suggestions","prompt",""));
            Mark("citation","event-green");
            if(mode=="navigation-only") { NavigationCases();Write("PASS",null);return 0; }
            MethodInfo callback=typeof(ComAddin).GetMethod("OpenHighlightedRules",BindingFlags.Instance|BindingFlags.NonPublic);
            Type delegateType=typeof(ComAddin).GetField("documentDoubleClick",BindingFlags.Instance|BindingFlags.NonPublic).FieldType;
            production=Delegate.CreateDelegate(delegateType,addin,callback);Set("documentDoubleClick",production);
            observer=Observe;
            ComEventsHelper.Combine((object)app,EventsId,14,observer);
            ComEventsHelper.Combine((object)app,EventsId,14,production);
            completed=Completed;ComEventsHelper.Combine((object)app,EventsId,14,completed);
            app.Visible=true;app.ActiveWindow.WindowState=0;app.ActiveWindow.View.Type=3;
            using(var closer=new System.Windows.Forms.Timer{Interval=40}) {
                closer.Tick+=delegate {
                    foreach(Form form in Application.OpenForms.Cast<Form>().ToArray()) {
                        if(form is RuleDialog||form is HighlightedRulesDialog) {
                            if(Seen.Add(form)) {
                                openedCount++;Dialogs.Add(Values("type",form.GetType().Name,"title",form.Text,"deferred",!inEvent,
                                    "nativeOwner",GetWindow(form.Handle,4).ToInt64(),"preserve",Descendants(form).OfType<CheckBox>().Any(box=>box.Checked&&box.Text.StartsWith("Preserve text",StringComparison.Ordinal)),
                                    "color",form is RuleDialog?((TextBox)typeof(RuleDialog).GetField("colorInput",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form)).Text:""));
                                if(form is RuleDialog&&chooseOverlap)choiceChildSeen=true;
                                if(form is HighlightedRulesDialog&&chooseOverlap) {
                                    ListBox list=Descendants(form).OfType<ListBox>().Single();
                                    Check(list.Items.Count==2,"Overlapping native identities require an explicit choice between both colors.");
                                    list.SelectedIndex=1;choiceMade=true;
                                    form.BeginInvoke((Action)delegate { form.AcceptButton.PerformClick(); });
                                    continue;
                                }
                            }
                            if(form is RuleDialog&&keepRuleWindows)continue;
                            if(form is HighlightedRulesDialog&&chooseOverlap&&(!choiceChildSeen||Application.OpenForms.Cast<Form>().Any(item=>item is RuleDialog)))continue;
                            if(form.CancelButton!=null)form.CancelButton.PerformClick();else form.Close();
                        }
                    }
                };closer.Start();
                string nativeBefore=new WordNativeEditState().NativeIdentity(doc),rulesBefore=Json.Serialize(Local("state")["rules"]),marksBefore=Json.Serialize(Local("state")["marks"]);
                doc.Range(0,0).Select();
                ClickWord("citation");
                PumpUntil(()=>eventCount>0&&openedCount>0,5000);
                Check(eventCount==1,"A genuine owned Word window double-click raises exactly one native BeforeDoubleClick event.");
                Check(openedCount==1&&Dialogs.Count==1,"The genuine text double-click opens one native RuleDialog and Cancel closes it.");
                Check((bool)Dialogs[0]["deferred"],"The native rule dialog opens after Word completes its input event callback.");
                Check((long)Dialogs[0]["nativeOwner"]==(long)(int)app.ActiveWindow.Hwnd,"The text rule dialog has the exact owned Word document as its native modal owner.");
                Check((bool)Dialogs[0]["preserve"]&&Cancellations.Single(),"A single protected green mark opens its existing Preserve rule and cancels Word's default text action.");
                Check(new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&Json.Serialize(Local("state")["rules"])==rulesBefore&&Json.Serialize(Local("state")["marks"])==marksBefore,
                    "Opening and cancelling the native text rule leaves source, native IDs, marks and rules unchanged.");
                Check(IsWindowEnabled(new IntPtr((int)app.ActiveWindow.Hwnd)),"Cancel restores the owned Word window's normal input state.");

                Local("save_rule",Values("id","event-blue","label","Native blue wording","color","#AACCEE","action","prompt","accessMode","direct","prompt","Use concise wording."));
                Mark("after","event-blue");nativeBefore=new WordNativeEditState().NativeIdentity(doc);rulesBefore=Json.Serialize(Local("state")["rules"]);marksBefore=Json.Serialize(Local("state")["marks"]);
                keepRuleWindows=true;int before=openedCount;ClickWord("citation");PumpUntil(()=>openedCount==before+1,5000);
                RuleDialog greenWindow=Application.OpenForms.OfType<RuleDialog>().Single();
                var draft=(TextBox)typeof(RuleDialog).GetField("nameInput",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(greenWindow);draft.Text="Retained green draft";
                ClickWord("after");PumpUntil(()=>openedCount==before+2,5000);
                Check(Application.OpenForms.OfType<RuleDialog>().Count()==2&&Dialogs[before]["color"].Equals("#33AA66")&&Dialogs[before+1]["color"].Equals("#AACCEE")&&IsWindowEnabled(new IntPtr((int)app.ActiveWindow.Hwnd)),
                    "Two genuine disjoint passage double-clicks keep Word enabled and open separate modeless rule windows with the exact green and blue colors.");
                ClickWord("citation");PumpUntil(()=>false,150);
                Check(openedCount==before+2&&draft.Text=="Retained green draft","Returning to the first actual passage reuses its existing color window without discarding its draft.");
                keepRuleWindows=false;foreach(Form form in Application.OpenForms.Cast<Form>().ToArray())if(form is RuleDialog)form.CancelButton.PerformClick();
                Check(new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&Json.Serialize(Local("state")["rules"])==rulesBefore&&Json.Serialize(Local("state")["marks"])==marksBefore,
                    "Opening both per-color windows and cancelling retained drafts does not change native source, marker IDs or permissions.");

                before=openedCount;doc.Range(29,29).Select();addin.RibbonSelectionRules(null);PumpUntil(()=>openedCount==before+1,5000);
                Check(openedCount==before+1&&Text(Dialogs.Last(),"type")=="RuleDialog"&&(bool)Dialogs.Last()["preserve"]&&
                    (long)Dialogs.Last()["nativeOwner"]==(long)(int)app.ActiveWindow.Hwnd,"The keyboard-equivalent ribbon callback opens the same protected color dialog with the same native Word owner.");
                Check(new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&Json.Serialize(Local("state")["rules"])==rulesBefore&&Json.Serialize(Local("state")["marks"])==marksBefore,
                    "Keyboard rule opening and Cancel preserve native source, rules and marks.");

                before=openedCount;int beforeEvents=eventCount;doc.Range(60,60).Select();ClickWord("Outside");
                PumpUntil(()=>eventCount>beforeEvents&&Cancellations.Count==eventCount,5000);PumpUntil(()=>false,200);
                Check(eventCount==beforeEvents+1&&!Cancellations.Last()&&openedCount==before&&(int)app.Selection.End>(int)app.Selection.Start,
                    "A genuine unmarked double-click keeps Word's normal word selection and opens no rule dialog.");

                before=openedCount;beforeEvents=eventCount;rebind=true;ClickWord("citation");
                PumpUntil(()=>eventCount>beforeEvents&&Cancellations.Count==eventCount,5000);PumpUntil(()=>false,200);
                Check(eventCount==beforeEvents+1&&Cancellations.Last()&&openedCount==before&&(int)app.Selection.Start==0,
                    "Moving the real selection before the deferred callback prevents a stale clicked color from opening.");
                Check(new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&Json.Serialize(Local("state")["rules"])==rulesBefore&&Json.Serialize(Local("state")["marks"])==marksBefore,
                    "The refused stale event leaves all native content, identities and local policy unchanged.");

                Local("save_rule",Values("id","event-gold","label","Native gold wording","color","#E8B949","action","prompt","accessMode","suggestions","prompt","Make the paragraph clear."));
                Mark("citation","event-gold");nativeBefore=new WordNativeEditState().NativeIdentity(doc);rulesBefore=Json.Serialize(Local("state")["rules"]);marksBefore=Json.Serialize(Local("state")["marks"]);
                chooseOverlap=true;before=openedCount;beforeEvents=eventCount;ClickWord("citation");
                PumpUntil(()=>eventCount>beforeEvents&&choiceChildSeen&&Application.OpenForms.Cast<Form>().All(form=>!(form is RuleDialog)&&!(form is HighlightedRulesDialog)),5000);
                Check(choiceMade&&choiceChildSeen&&openedCount==before+2&&Dialogs[before]["type"].Equals("HighlightedRulesDialog")&&Dialogs[before+1]["type"].Equals("RuleDialog"),
                    "An actual overlapping double-click opens the chooser, then its explicitly selected per-color RuleDialog.");
                Check((bool)Dialogs[before]["deferred"]&&(long)Dialogs[before]["nativeOwner"]==(long)(int)app.ActiveWindow.Hwnd,
                    "The overlap chooser also opens after native input with the exact owned Word parent.");
                Check(new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&Json.Serialize(Local("state")["rules"])==rulesBefore&&Json.Serialize(Local("state")["marks"])==marksBefore&&IsWindowEnabled(new IntPtr((int)app.ActiveWindow.Hwnd)),
                    "Explicit overlap inspection and both Cancel actions leave source, native IDs and permissions unchanged and Word enabled.");

                otherDocument=app.Documents.Add();otherDocument.Content.Text="Another owned document.\r";doc.Activate();
                string otherBefore=new WordNativeEditState().NativeIdentity(otherDocument);
                before=openedCount;beforeEvents=eventCount;switchDocument=true;ClickWord("citation");
                PumpUntil(()=>eventCount>beforeEvents&&Cancellations.Count==eventCount,5000);PumpUntil(()=>false,200);
                Check(eventCount==beforeEvents+1&&openedCount==before&&new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&new WordNativeEditState().NativeIdentity(otherDocument)==otherBefore,
                    "Switching to another actual native document before the queued action opens no stale dialog and changes neither source.");
                doc.Activate();otherDocument.Close(0);otherDocument=null;

                string diskBefore=Hash(documentFile);before=openedCount;beforeEvents=eventCount;closeDocument=true;ClickWord("citation");
                PumpUntil(()=>eventCount>beforeEvents&&Cancellations.Count==eventCount,5000);PumpUntil(()=>false,200);
                Check(eventCount==beforeEvents+1&&openedCount==before&&(int)app.Documents.Count==0&&Hash(documentFile)==diskBefore&&ThreadErrors.Count==0,
                    "Closing the actual document before the queued action opens no dialog, saves nothing and raises no unhandled UI exception.");

                chooseOverlap=false;doc=app.Documents.Add();doc.Content.Text="A green citation stays intact.\r";
                documentFile=Path.Combine(output,"disconnect-fixture.docx");doc.SaveAs2(documentFile,16);
                documentId=Text(Map(Local("state")["document"]),"id");DocumentIds.Add(documentId);
                Local("save_rule",Values("id","event-green","label","Native green citation","color","#33AA66","action","protect","accessMode","suggestions","prompt",""));Mark("citation","event-green");
                nativeBefore=new WordNativeEditState().NativeIdentity(doc);diskBefore=Hash(documentFile);
                before=openedCount;beforeEvents=eventCount;disconnect=true;ClickWord("citation");
                PumpUntil(()=>eventCount>beforeEvents&&Cancellations.Count==eventCount,5000);PumpUntil(()=>false,200);
                Check(eventCount==beforeEvents+1&&openedCount==before&&new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&Hash(documentFile)==diskBefore&&dispatcher.IsDisposed&&ThreadErrors.Count==0,
                    "Actual add-in disconnection before the queued action disposes its dispatcher without a dialog, source mutation, save or unhandled exception.");
                Check(!(bool)typeof(ComAddin).GetField("highlightedRulesPending",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(addin),
                    "Disconnection clears a queued pending flag even when its disposed dispatcher cannot run the delegate's finally.");
                editor=new WordEditor((object)app,operation=>operation(),Path.Combine(output,"profile-reconnected"),Path.Combine(output,"unused-zrok.exe"));
                dispatcher=new Control();IntPtr reconnectedHandle=dispatcher.Handle;
                Set("editor",editor);Set("dispatcher",dispatcher);Set("wordApplication",(object)app);Set("ownerThread",Thread.CurrentThread.ManagedThreadId);
                ComEventsHelper.Remove((object)app,EventsId,14,observer);ComEventsHelper.Remove((object)app,EventsId,14,completed);
                production=Delegate.CreateDelegate(delegateType,addin,callback);Set("documentDoubleClick",production);
                ComEventsHelper.Combine((object)app,EventsId,14,observer);ComEventsHelper.Combine((object)app,EventsId,14,production);ComEventsHelper.Combine((object)app,EventsId,14,completed);
                before=openedCount;beforeEvents=eventCount;ClickWord("citation");
                PumpUntil(()=>eventCount>beforeEvents&&openedCount>before,5000);
                Check(eventCount==beforeEvents+1&&openedCount==before+1&&Text(Dialogs.Last(),"type")=="RuleDialog"&&(bool)Dialogs.Last()["deferred"]&&
                    new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&Hash(documentFile)==diskBefore&&ThreadErrors.Count==0,
                    "The same add-in object opens its protected color again after isolated reconnection without stale pending state or source changes.");
            }
            Write("PASS",null);return 0;
        }catch(Exception error){failure=error.ToString();Console.Error.WriteLine(error);Write("FAIL",failure);return 1;}
        finally {
            if(app!=null){try{if(observer!=null)ComEventsHelper.Remove((object)app,EventsId,14,observer);if(production!=null)ComEventsHelper.Remove((object)app,EventsId,14,production);if(completed!=null)ComEventsHelper.Remove((object)app,EventsId,14,completed);}catch(Exception error){Console.Error.WriteLine(error.Message);}}
            if(editor!=null)editor.Dispose();if(dispatcher!=null)dispatcher.Dispose();
            if(app!=null)app.Quit(0);else if(owned!=null&&!owned.HasExited)owned.Kill();
            foreach(string id in DocumentIds){string receipt=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"wrAIterWord","reviews",id+".bin");if(File.Exists(receipt))File.Delete(receipt);}
            if(filter!=null)filter.Dispose();
        }
    }
    private static void Observe(object selection,ref bool cancel)
    {
        inEvent=true;eventCount++;dynamic received=selection;
        Events.Add(new{eventStart=(int)received.Range.Start,eventEnd=(int)received.Range.End,applicationStart=(int)app.Selection.Range.Start,applicationEnd=(int)app.Selection.Range.End,
            story=(int)received.Range.StoryType,depth=(int)typeof(ComAddin).GetField("authorityDepth",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(addin),cancel=cancel});
    }
    private static void Completed(object selection,ref bool cancel)
    {
        Cancellations.Add(cancel);
        if(rebind){rebind=false;doc.Range(0,0).Select();}
        if(switchDocument){switchDocument=false;otherDocument.Activate();}
        if(closeDocument){closeDocument=false;doc.Close(0);doc=null;}
        if(disconnect){disconnect=false;Array custom=new object[0];addin.OnDisconnection(0,ref custom);editor=null;}
        inEvent=false;
    }
    private static IEnumerable<Control> Descendants(Control control){foreach(Control child in control.Controls){yield return child;foreach(Control nested in Descendants(child))yield return nested;}}
    private static void ClickWord(string word)
    {
        IntPtr frame=new IntPtr((int)app.ActiveWindow.Hwnd);uint frameProcess;GetWindowThreadProcessId(frame,out frameProcess);
        Check(frameProcess==ownedPid&&IsWindowEnabled(frame),"Only the exact enabled synthetic Word window receives native input activation.");
        textWindow=IntPtr.Zero; EnumChildWindows(new IntPtr((int)app.ActiveWindow.Hwnd),delegate(IntPtr child,IntPtr ignored){var name=new StringBuilder(80);GetClassName(child,name,name.Capacity);if(name.ToString()=="_WwG")textWindow=child;return true;},IntPtr.Zero); Check(textWindow!=IntPtr.Zero,"The current synthetic document has its own live native text child.");
        WordSnapshot snap=WordSnapshot.Capture(doc);int start=snap.MainText.IndexOf(word,StringComparison.Ordinal);
        doc.Repaginate(); app.ActiveWindow.ScrollIntoView(snap.Range(doc,start,start+1),true); Application.DoEvents(); Thread.Sleep(250);
        int left,top,width,height;app.ActiveWindow.GetPoint(out left,out top,out width,out height,snap.Range(doc,start,start+1));
        Point point=new Point{X=left+Math.Max(1,width/2),Y=top+Math.Max(1,height/2)};Check(ScreenToClient(textWindow,ref point),"The owned text position maps to its exact native document window.");
        File.WriteAllText(Path.Combine(output,"click-position.json"),Json.Serialize(new{mode="Explicit real input required; this harness sends no mouse input",pid=ownedPid,title=(string)doc.Name,word=word,expectedStart=start,left=left,top=top,width=width,height=height,window=frame.ToInt64()}));
        Console.WriteLine("REAL INPUT REQUIRED: double-click '"+word+"' in owned Word PID "+ownedPid+" at "+left+","+top+" (see click-position.json).");
        int before=eventCount;PumpUntil(()=>eventCount>before,60000);
        Check(eventCount>before,"The explicitly requested physical double-click reaches the actual Word event subscription.");
    }
    private static void PumpUntil(Func<bool> done,int milliseconds){DateTime end=DateTime.UtcNow.AddMilliseconds(milliseconds);while(!done()&&DateTime.UtcNow<end){Application.DoEvents();Thread.Sleep(15);}}
    private static object AttachOwned(Process process)
    {
        DateTime end=DateTime.UtcNow.AddSeconds(30);
        while(textWindow==IntPtr.Zero&&DateTime.UtcNow<end){EnumWindows(delegate(IntPtr window,IntPtr data){uint pid;GetWindowThreadProcessId(window,out pid);if(pid!=process.Id)return true;
            EnumChildWindows(window,delegate(IntPtr child,IntPtr ignored){var name=new StringBuilder(80);GetClassName(child,name,name.Capacity);if(name.ToString()=="_WwG")textWindow=child;return true;},IntPtr.Zero);return true;},IntPtr.Zero);Application.DoEvents();Thread.Sleep(30);}
        Check(textWindow!=IntPtr.Zero,"The separately spawned Word process exposes an owned text window.");
        object native=null;Guid dispatch=new Guid("00020400-0000-0000-C000-000000000046");
        while(native==null&&DateTime.UtcNow<end){if(AccessibleObjectFromWindow(textWindow,0xfffffff0,ref dispatch,out native)!=0)native=null;Application.DoEvents();Thread.Sleep(30);}
        Check(native!=null,"The exact owned Word automation endpoint becomes ready.");dynamic result=((dynamic)native).Application;uint actual;GetWindowThreadProcessId(new IntPtr((int)result.ActiveWindow.Hwnd),out actual);
        Check(actual==process.Id,"Native automation is bound to the exact spawned PID, without the user's Word instance.");ownedPid=process.Id;Console.WriteLine("Owned Word PID "+process.Id);return result;
    }
    private static void Mark(string quote,string rule)
    {
        WordSnapshot snap=WordSnapshot.Capture(doc);int start=snap.MainText.IndexOf(quote,StringComparison.Ordinal);snap.Range(doc,start,start+quote.Length).Select();
        var preview=Local("selection_preview");preview["ruleId"]=rule;preview["prompt"]="";Local("mark_selection",preview);
    }
    private static void NavigationCases()
    {
        Local("save_rule",Values("id","event-yellow","label","Yellow wording","color","#E8B949","action","prompt","accessMode","suggestions","prompt","Retain the meaning."));
        Local("save_rule",Values("id","event-blue","label","Blue direct wording","color","#AACCEE","action","prompt","accessMode","direct","prompt","Use concise wording."));
        Mark("before","event-yellow");Mark("after","event-blue");
        var reviews=(IDictionary)typeof(WordEditor).GetField("documents",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(editor);
        var review=reviews.Values.Cast<WordDocumentReview>().Single(value=>value.Id==documentId);
        PassageMark green=review.Policy.Marks.Single(mark=>mark.RuleId=="event-green");
        PassageMark yellow=review.Policy.Marks.Single(mark=>mark.RuleId=="event-yellow");
        PassageMark blue=review.Policy.Marks.Single(mark=>mark.RuleId=="event-blue");
        doc.Range(green.From,green.From).Select();
        int selection=(int)app.Selection.Start;
        var blueChoices=UiState.Records(Local("selection_rules",Values("nativeFrom",blue.From,"nativeTo",blue.From)),"choices");
        Check(blueChoices.Count==1&&UiState.Text(blueChoices[0],"ruleId")==blue.RuleId&&(int)app.Selection.Start==selection,
            "An exact native clicked point resolves the second disjoint color without moving the old green selection.");
        blue.RequiresRemarking=true;blue.Stale=true;
        Local("state");
        string blueBefore=Json.Serialize(blue),nativeBefore=new WordNativeEditState().NativeIdentity(doc),diskBefore=Hash(documentFile);
        string rulesBefore=Json.Serialize(Local("state")["rules"]);
        blueChoices=UiState.Records(Local("selection_rules",Values("nativeFrom",blue.From,"nativeTo",blue.From)),"choices");
        Check(blueChoices.Count==1&&UiState.Text(blueChoices[0],"id")==blue.Id&&(bool)blueChoices[0]["stale"]&&Json.Serialize(blue)==blueBefore,
            "An explicitly ambiguous exact-quote marker remains navigable through its surviving native bookmark without changing its policy geometry or quote.");
        WordSnapshot snap=WordSnapshot.Capture(doc);review.Policy.Share(documentId);
        string context=review.Policy.ReadContext(documentId,snap.Revision,snap.MainText,snap.FullText);
        bool denied=false;
        try { review.Policy.Submit(new Proposal{Id="stale_blue_no_authority",DocumentId=documentId,Revision=snap.Revision,From=blue.From,To=blue.To,Quote=blue.Quote,Replacement="later",Rationale=""},context,snap.MainText,snap.FullText); }
        catch(InvalidOperationException){denied=true;}
        Check(denied&&blue.Stale&&!review.Policy.Proposals.Any(proposal=>proposal.Id=="stale_blue_no_authority"),
            "Navigating the stale direct color does not restore suggestion or direct editing authority.");
        int yellowFrom=(int)doc.Bookmarks.Item(yellow.Id).Range.Start;
        doc.Range(yellowFrom+1,yellowFrom+2).Text="X";doc.Bookmarks.Item(yellow.Id).Range.Shading.BackgroundPatternColor=-16777216;Local("state");
        Check(yellow.Stale&&(bool)doc.Bookmarks.Exists(yellow.Id)&&(int)doc.Bookmarks.Item(yellow.Id).Range.End>(int)doc.Bookmarks.Item(yellow.Id).Range.Start,
            "Changed wording with removed visible permission leaves a real noncollapsed marker while native policy remains stale.");
        string yellowBefore=Json.Serialize(yellow);nativeBefore=new WordNativeEditState().NativeIdentity(doc);
        var yellowChoices=UiState.Records(Local("selection_rules",Values("nativeFrom",yellowFrom,"nativeTo",yellowFrom)),"choices");
        Check(yellowChoices.Count==1&&UiState.Text(yellowChoices[0],"ruleId")==yellow.RuleId&&Json.Serialize(yellow)==yellowBefore,
            "Changed-quote stale coloration opens only the rule attached to its exact surviving native marker, retaining original policy fields.");
        doc.Bookmarks.Item(yellow.Id).Delete();Local("state");
        yellowChoices=UiState.Records(Local("selection_rules",Values("nativeFrom",yellowFrom,"nativeTo",yellowFrom)),"choices");
        Check(yellowChoices.Count==0,"Visible shading and stale old offsets never recover a missing native marker.");
        doc.Bookmarks.Add(yellow.Id,doc.Range(yellowFrom,yellowFrom));
        yellowChoices=UiState.Records(Local("selection_rules",Values("nativeFrom",yellowFrom,"nativeTo",yellowFrom)),"choices");
        Check(yellowChoices.Count==0,"A collapsed native marker cannot provide a guessed rule target.");
        doc.Bookmarks.Item(yellow.Id).Delete();
        dynamic header=doc.Sections.Item(1).Headers.Item(1).Range;header.Text="Header text";
        doc.Bookmarks.Add(yellow.Id,header);
        yellowChoices=UiState.Records(Local("selection_rules",Values("nativeFrom",yellowFrom,"nativeTo",yellowFrom)),"choices");
        Check(yellowChoices.Count==0,"An exact bookmark name in another Word story cannot navigate a main-body color.");
        doc.Bookmarks.Add("WraiO_Unclaimed",doc.Range(blue.From,blue.To));blue.NativeBookmark="WraiO_Unclaimed";
        blueChoices=UiState.Records(Local("selection_rules",Values("nativeFrom",blue.From,"nativeTo",blue.From)),"choices");
        Check(blueChoices.Count==0,"A native alias without the exact persisted marker-witness binding cannot select a rule.");
        blue.NativeBookmark=null;
        nativeBefore=new WordNativeEditState().NativeIdentity(doc);string marksBefore=Json.Serialize(Local("state")["marks"]);
        denied=false;try{Local("selection_rules",Values("nativeFrom",-1,"nativeTo",1));}catch(InvalidOperationException){denied=true;}
        Check(denied&&new WordNativeEditState().NativeIdentity(doc)==nativeBefore&&Json.Serialize(Local("state")["marks"])==marksBefore,
            "An out-of-body clicked range is refused without native content or marker mutation.");
        Check(Json.Serialize(Local("state")["rules"])==rulesBefore&&Hash(documentFile)==diskBefore&&ThreadErrors.Count==0,
            "All local navigation and denials keep rules and saved document bytes unchanged without unhandled errors or foreground input.");
    }
    private static void Set(string field,object value){typeof(ComAddin).GetField(field,BindingFlags.Instance|BindingFlags.NonPublic).SetValue(addin,value);}
    private static Dictionary<string,object> Local(string name,Dictionary<string,object> args=null){return Map(Json.DeserializeObject(Json.Serialize(editor.Dispatch(name,args??new Dictionary<string,object>()))));}
    private static Dictionary<string,object> Map(object value){return (Dictionary<string,object>)value;}
    private static string Text(Dictionary<string,object> value,string key){return Convert.ToString(value[key]);}
    private static string Hash(string file){using(var stream=new FileStream(file,FileMode.Open,FileAccess.Read,FileShare.ReadWrite))using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-","").ToLowerInvariant();}
    private static Dictionary<string,object> Values(params object[] args){var result=new Dictionary<string,object>();for(int i=0;i<args.Length;i+=2)result.Add((string)args[i],args[i+1]);return result;}
    private static void Check(bool pass,string message){if(!pass)throw new Exception(message);Checks.Add(message);Console.WriteLine("PASS: "+message);}
    private static void Write(string status,string error){if(output!=null)File.WriteAllText(Path.Combine(output,"result.json"),Json.Serialize(new{status=status,mode=mode,checks=Checks,events=Events,cancellations=Cancellations,dialogs=Dialogs,threadErrors=ThreadErrors,error=error,normalWord="NOT TOUCHED",installed="NOT RUN"}));}
}
