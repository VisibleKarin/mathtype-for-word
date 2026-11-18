// Compiled locally for the installed MT6.dll architecture. No vendor binaries are distributed.
using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Web.Script.Serialization;

public static class MathTypeNativeWorker {
    const string Dll = "MT6.dll";
    [StructLayout(LayoutKind.Sequential)] public struct Dims { public int baseline, left, top, right, bottom; }
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr LoadLibrary(string path);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool SetDllDirectory(string path);
    [DllImport(Dll)] static extern int MTAPIConnect(short flags, short timeout);
    [DllImport(Dll)] static extern int MTAPIDisconnect();
    [DllImport(Dll)] static extern int MTCloseOleObject(int save, [MarshalAs(UnmanagedType.IDispatch)] object eq);
    [DllImport(Dll)] static extern int MTSetEqnFromLangBytes([MarshalAs(UnmanagedType.IDispatch)] object eq, short lang, byte[] bytes, int length);
    [DllImport(Dll, CharSet=CharSet.Ansi)] static extern int MTGetPrefsFromClipboard(StringBuilder prefs, short length);
    [DllImport(Dll)] static extern int MTXFormReset();
    [DllImport(Dll, CharSet=CharSet.Ansi)] static extern int MTXFormSetPrefs(short kind, string prefs);
    [DllImport(Dll, CharSet=CharSet.Ansi)] static extern int MTXFormEqn(short src, short srcFmt, byte[] bytes, int length,
        short dst, short dstFmt, IntPtr result, int resultLength, string path, ref Dims dims);
    [DllImport(Dll)] static extern int MTXFormGetStatus(short index);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern uint RegisterClipboardFormat(string name);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr memory);
    [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr memory);
    [DllImport("kernel32.dll")] static extern UIntPtr GlobalSize(IntPtr memory);
    [DllImport("ole32.dll")] static extern void ReleaseStgMedium(ref STGMEDIUM medium);
    static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);

    static void Check(int code, string step) {
        if (code != 0) throw new InvalidOperationException(step + " returned " + code);
    }
    static string MathML(object equation) {
        IDataObject data = equation as IDataObject;
        if (data == null) throw new InvalidOperationException("MathType has no IDataObject.");
        foreach (string name in new [] {"MathML Presentation", "MathML", "application/mathml+xml"}) {
            FORMATETC f = new FORMATETC {cfFormat=unchecked((short)RegisterClipboardFormat(name)),
                dwAspect=DVASPECT.DVASPECT_CONTENT, lindex=-1, tymed=TYMED.TYMED_HGLOBAL};
            if (data.QueryGetData(ref f) != 0) continue;
            STGMEDIUM medium;
            data.GetData(ref f, out medium);
            try {
                IntPtr p = GlobalLock(medium.unionmember);
                if (p == IntPtr.Zero) throw new InvalidOperationException("Cannot lock MathML data.");
                try {
                    ulong length = GlobalSize(medium.unionmember).ToUInt64();
                    if (length < 2 || length > 32 * 1024 * 1024) throw new InvalidOperationException("Invalid MathML length.");
                    byte[] bytes = new byte[(int)length];
                    Marshal.Copy(p, bytes, 0, bytes.Length);
                    // UTF-16 may have a BOM or begin with '<\0'; UTF-8 NUL terminators are byte-aligned.
                    bool wide = (bytes[0] == 0xff && bytes[1] == 0xfe) || bytes[1] == 0;
                    if (wide) {
                        int count = 0;
                        while (count + 1 < bytes.Length && (bytes[count] != 0 || bytes[count+1] != 0)) count += 2;
                        return Encoding.Unicode.GetString(bytes, 0, count).TrimStart('\uFEFF');
                    }
                    int n = Array.IndexOf(bytes, (byte)0);
                    return Utf8.GetString(bytes, 0, n < 0 ? bytes.Length : n).TrimStart('\uFEFF');
                } finally { GlobalUnlock(medium.unionmember); }
            } finally { ReleaseStgMedium(ref medium); }
        }
        throw new InvalidOperationException("MathType exposes no MathML clipboard format.");
    }
    static string Prefs(dynamic word, dynamic shape) {
        shape.Select();
        word.Selection.Copy();
        int length = -1;
        for (int retry=0; retry<6 && length<1; retry++) {
            System.Threading.Thread.Sleep(120 * (retry+1));
            length = MTGetPrefsFromClipboard(null, 0);
        }
        if (length < 1 || length > short.MaxValue) throw new InvalidOperationException("Cannot read embedded preferences: " + length);
        var result = new StringBuilder(length);
        Check(MTGetPrefsFromClipboard(result, (short)length), "read preferences");
        return result.ToString();
    }
    static byte[] Transform(byte[] input, string prefs) {
        // Never set a translator for MTEF -> MTEF; some MT6 builds crash on a null translator.
        // Use a bounded growing output buffer, rather than the historical fixed 30 KB buffer.
        for (int capacity=Math.Max(65536, input.Length*2); capacity<=16*1024*1024; capacity*=2) {
            Check(MTXFormReset(), "reset transform");
            Check(MTXFormSetPrefs(3, prefs), "set equation preferences");
            IntPtr memory = Marshal.AllocHGlobal(capacity);
            try {
                Dims dims = new Dims();
                int code = MTXFormEqn(-3, 4, input, input.Length, -3, 4, memory, capacity, "", ref dims);
                int length = MTXFormGetStatus(-1);
                if (code == 0 && length > 0 && length <= capacity) {
                    byte[] result = new byte[length];
                    Marshal.Copy(memory, result, 0, length);
                    return result;
                }
                // Retry only when the SDK explicitly reports output larger than the supplied buffer.
                if (length <= capacity) throw new InvalidOperationException("Native transform failed: " + code + ", length=" + length);
            } finally { Marshal.FreeHGlobal(memory); }
        }
        throw new InvalidOperationException("Native equation output exceeds 16 MiB.");
    }
    static dynamic Shape(dynamic doc, int wanted) {
        int index = 0;
        for (int k=1; k<=doc.InlineShapes.Count; k++) {
            dynamic shape = doc.InlineShapes[k];
            string prog = "";
            try { prog = (string)shape.OLEFormat.ProgID; } catch (COMException) {}
            if (prog == "Equation.DSMT4" && ++index == wanted) return shape;
        }
        throw new InvalidOperationException("Inline MathType equation not found: " + wanted);
    }
    static string ReadMathML(dynamic shape) {
        shape.OLEFormat.DoVerb(2); // MathType's hidden conversion verb, never an interactive edit verb.
        object obj = shape.OLEFormat.Object;
        try { return MathML(obj); }
        finally { Check(MTCloseOleObject(1, obj), "close equation"); }
    }
    static string ChangeFull(string prefs, double size) {
        string current = "";
        int count = 0;
        string[] lines = prefs.Replace("\r\n", "\n").Split('\n');
        for (int i=0; i<lines.Length; i++) {
            string line = lines[i].Trim();
            if (line.StartsWith("[") && line.EndsWith("]")) current = line;
            if (current == "[Sizes]" && System.Text.RegularExpressions.Regex.IsMatch(line, @"^Full\s*=")) {
                lines[i] = "Full=" + size.ToString("0.###", CultureInfo.InvariantCulture) + " pt";
                count++;
            }
        }
        if (count != 1) throw new InvalidOperationException("Expected exactly one internal Full preference.");
        return string.Join("\r\n", lines);
    }
    static void Write(string folder, int index, string suffix, string value) {
        File.WriteAllText(Path.Combine(folder, index + "-" + suffix), value, Utf8);
    }
    [STAThread] public static int Main(string[] args) {
        dynamic word = null, doc = null;
        bool connected = false, owned = false;
        var serializer = new JavaScriptSerializer();
        try {
            var job = serializer.Deserialize<Dictionary<string, object>>(File.ReadAllText(args[0], Encoding.UTF8));
            string root = args[1], journal = args[2], folder = (string)job["work_dir"];
            string input = Path.GetFullPath((string)job["input_path"]), output = Path.GetFullPath((string)job["output_path"]);
            if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase) || File.Exists(output) ||
                !string.Equals(Path.GetDirectoryName(output), Path.GetFullPath(folder), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Native worker output must be a new file in its private work directory.");
            string dll = Path.Combine(root, "System", Dll);
            string blank = Path.Combine(root, "Office Support", "BlankEqn.doc");
            if (!File.Exists(dll) || !File.Exists(blank)) throw new FileNotFoundException("MathType SDK or blank equation template is missing.");
            if (!SetDllDirectory(Path.GetDirectoryName(dll)) || LoadLibrary(dll) == IntPtr.Zero)
                throw new InvalidOperationException("Cannot load the installed MathType SDK: " + Marshal.GetLastWin32Error());
            Check(MTAPIConnect(0, 15), "connect MathType SDK");
            connected = true;
            var before = Process.GetProcessesByName("WINWORD").Select(p => p.Id).ToArray();
            word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
            var launched = Process.GetProcessesByName("WINWORD").Where(p => !before.Contains(p.Id)).ToArray();
            owned = launched.Length == 1;
            if (!owned) throw new InvalidOperationException("Word did not create an isolated process; existing sessions were left untouched.");
            int pid = launched[0].Id;
            Console.Error.WriteLine("[INFO] WORD_PID=" + pid);
            Directory.CreateDirectory(Path.GetDirectoryName(journal));
            File.WriteAllText(journal, serializer.Serialize(new [] {new Dictionary<string, object> {
                {"pid", pid}, {"start", launched[0].StartTime.ToFileTimeUtc()}}}), Utf8);
            word.Visible = false;
            word.DisplayAlerts = 0;
            word.AutomationSecurity = 3;
            doc = word.Documents.Open((string)job["input_path"], false, true, false);
            foreach (Dictionary<string, object> op in (IEnumerable)job["operations"]) {
                int index = Convert.ToInt32(op["index"]);
                double size = Convert.ToDouble(op["full_size_pt"], CultureInfo.InvariantCulture);
                Console.Error.WriteLine("[INFO] Reading equation #" + index);
                dynamic shape = Shape(doc, index);
                Write(folder, index, "before.mml", ReadMathML(shape));
                string prefs = Prefs(word, shape);
                Write(folder, index, "before.prefs", prefs);
                byte[] formatted = Transform(File.ReadAllBytes((string)op["mtef_path"]), ChangeFull(prefs, size));
                Console.Error.WriteLine("[INFO] Transformed native data for equation #" + index);
                int end = (int)shape.Range.End;
                int originalShapeCount = (int)doc.InlineShapes.Count;
                dynamic insertion = doc.Range(end, end);
                insertion.InsertFile(blank, "MTBlankEqn");
                if ((int)doc.InlineShapes.Count != originalShapeCount + 1) throw new InvalidOperationException("Blank template inserted an unexpected object count.");
                dynamic replacement = Shape(doc, index + 1);
                if ((string)replacement.OLEFormat.ProgID != "Equation.DSMT4") throw new InvalidOperationException("Expected an empty native MathType object.");
                shape.Delete();
                if ((bool)doc.Bookmarks.Exists("MTBlankEqn")) doc.Bookmarks["MTBlankEqn"].Delete();
                replacement.OLEFormat.DoVerb(2);
                object eq = replacement.OLEFormat.Object;
                try {
                    // Never write into a nonempty object: the SDK otherwise appends instead of replacing.
                    string blankML = MathML(eq);
                    var blankTree = System.Xml.Linq.XDocument.Parse(blankML);
                    if (blankTree.Descendants().Any(n => new [] {"mi", "mn", "mo", "mtext"}.Contains(n.Name.LocalName) && n.Value.Trim().Length > 0))
                        throw new InvalidOperationException("Vendor blank equation is not empty.");
                    Check(MTSetEqnFromLangBytes(eq, 3, formatted, formatted.Length), "write formatted MTEF");
                } finally { Check(MTCloseOleObject(1, eq), "save replacement equation"); }
                replacement.ScaleWidth = 100;
                replacement.ScaleHeight = 100;
                Console.Error.WriteLine("[INFO] Formatted equation #" + index);
            }
            doc.SaveAs2((string)job["output_path"], 16);
            doc.Close(0);
            doc = null;
            doc = word.Documents.Open((string)job["output_path"], false, true, false);
            foreach (Dictionary<string, object> op in (IEnumerable)job["operations"]) {
                int index = Convert.ToInt32(op["index"]);
                dynamic shape = Shape(doc, index);
                Write(folder, index, "saved.mml", ReadMathML(shape));
                Write(folder, index, "saved.prefs", Prefs(word, shape));
            }
            Console.WriteLine("{\"ok\":true,\"action\":\"format-native\"}");
            return 0;
        } catch (Exception ex) {
            Console.Error.WriteLine(ex.ToString());
            Console.WriteLine(serializer.Serialize(new Dictionary<string, object> {{"ok", false}, {"action", "format-native"}, {"error", ex.GetType().Name + ": " + ex.Message}}));
            return 1;
        } finally {
            if (doc != null) try { doc.Close(0); } catch {}
            if (word != null && owned) try {
                if (!(bool)word.Visible && (int)word.Documents.Count == 0) word.Quit(0);
            } catch {}
            if (connected) try { MTAPIDisconnect(); } catch {}
        }
    }
}
