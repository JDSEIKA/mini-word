using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;

namespace WordPadClone.Controls
{
    internal sealed class SpellError
    {
        public int Paragraph;
        public int Start;
        public int Length;
        public string Word;
        public int TableRow = -1;
        public int TableColumn = -1;
    }

    /// <summary>
    /// Windows 8+ spell checking (MsSpellCheckingFacility) through its native
    /// COM vtable.  COM is single-threaded here: all calls happen on one MTA
    /// worker, and results are posted back to the UI SynchronizationContext.
    /// No dictionary or third-party library is shipped with the application.
    /// </summary>
    internal sealed class SystemSpellChecker : IDisposable
    {
        private const string FactoryClsid = "7AB36653-1796-484B-BDFA-E74F1DB7C1DC";

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int IsSupportedFn(IntPtr self, string tag, out int supported);
        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int CreateCheckerFn(IntPtr self, string tag, out IntPtr checker);
        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int CheckFn(IntPtr self, string text, out IntPtr errors);
        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int SuggestFn(IntPtr self, string word, out IntPtr suggestions);
        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int WordFn(IntPtr self, string word);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int NextErrorFn(IntPtr self, out IntPtr error);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int NextStringFn(IntPtr self, int count, out IntPtr value, out int fetched);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetIntFn(IntPtr self, out int value);


        private sealed class CheckRequest
        {
            public string[] Paragraphs;
            public int Version;
            public Action<int, List<SpellError>> Completed;
        }

        private sealed class SuggestRequest
        {
            public string Word;
            public Action<List<string>> Completed;
        }

        private sealed class WordRequest
        {
            public string Word;
            public bool IsAdd;
        }

        private readonly BlockingCollection<object> _queue = new BlockingCollection<object>();
        private readonly SynchronizationContext _context;
        private readonly Thread _worker;
        private readonly string _language;
        private volatile bool _disposed;
        private volatile bool _available;
        private string _activeLanguage = string.Empty;

        public SystemSpellChecker(string language)
        {
            _context = SynchronizationContext.Current;
            _language = string.IsNullOrEmpty(language) ? CultureInfo.CurrentCulture.Name : language;
            _worker = new Thread(Worker);
            _worker.IsBackground = true;
            _worker.Name = "MiniWord spell checker";
            _worker.SetApartmentState(ApartmentState.MTA);
            _worker.Start();
        }

        public bool IsAvailable { get { return _available; } }
        public string Language { get { return _activeLanguage; } }

        public void Check(string[] paragraphs, int version, Action<int, List<SpellError>> completed)
        {
            if (_disposed || !_available) return;
            _queue.Add(new CheckRequest { Paragraphs = paragraphs, Version = version, Completed = completed });
        }

        public void Suggest(string word, Action<List<string>> completed)
        {
            if (_disposed || !_available) return;
            _queue.Add(new SuggestRequest { Word = word, Completed = completed });
        }

        public void Ignore(string word)
        {
            if (_disposed || !_available) return;
            _queue.Add(new WordRequest { Word = word, IsAdd = false });
        }

        public void Add(string word)
        {
            if (_disposed || !_available) return;
            _queue.Add(new WordRequest { Word = word, IsAdd = true });
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { _queue.CompleteAdding(); } catch { }
            try { if (!_worker.Join(800)) _worker.Abort(); } catch { }
        }

        private void Post(Action action)
        {
            if (_context != null) _context.Post(delegate { action(); }, null);
            else action();
        }

        private void Worker()
        {
            IntPtr factory = IntPtr.Zero;
            IntPtr checker = IntPtr.Zero;
            object factoryObject = null;
            try
            {
                NativeMethods.CoInitializeEx(IntPtr.Zero, 0);
                Type type = Type.GetTypeFromCLSID(new Guid(FactoryClsid));
                if (type == null) return;
                factoryObject = Activator.CreateInstance(type);
                factory = Marshal.GetIUnknownForObject(factoryObject);
                if (factory == IntPtr.Zero) return;

                IsSupportedFn supported = Function<IsSupportedFn>(factory, 4);
                CreateCheckerFn create = Function<CreateCheckerFn>(factory, 5);
                string[] candidates = new string[] { _language, "es-MX", "es-ES", "en-US", "en-GB" };
                for (int i = 0; i < candidates.Length; i++)
                {
                    if (string.IsNullOrEmpty(candidates[i])) continue;
                    int yes;
                    int hr;
                    try { hr = supported(factory, candidates[i], out yes); }
                    catch { hr = -1; yes = 0; }
                    if (hr != 0 || yes == 0) continue;
                    hr = create(factory, candidates[i], out checker);
                    if (hr == 0 && checker != IntPtr.Zero)
                    {
                        _activeLanguage = candidates[i];
                        break;
                    }
                }
                if (checker == IntPtr.Zero) return;
                _available = true;

                CheckFn check = Function<CheckFn>(checker, 4);
                SuggestFn suggest = Function<SuggestFn>(checker, 5);
                WordFn add = Function<WordFn>(checker, 6);
                WordFn ignore = Function<WordFn>(checker, 7);

                foreach (object item in _queue.GetConsumingEnumerable())
                {
                    if (_disposed) break;
                    CheckRequest checkRequest = item as CheckRequest;
                    if (checkRequest != null)
                    {
                        List<SpellError> result = new List<SpellError>();
                        try { RunCheck(checker, check, checkRequest.Paragraphs, result); }
                        catch { }
                        if (checkRequest.Completed != null)
                        {
                            int version = checkRequest.Version;
                            Action<int, List<SpellError>> callback = checkRequest.Completed;
                            Post(delegate { callback(version, result); });
                        }
                        continue;
                    }
                    SuggestRequest suggestRequest = item as SuggestRequest;
                    if (suggestRequest != null)
                    {
                        List<string> result = new List<string>();
                        try { RunSuggest(checker, suggest, suggestRequest.Word, result); }
                        catch { }
                        if (suggestRequest.Completed != null)
                        {
                            Action<List<string>> callback = suggestRequest.Completed;
                            Post(delegate { callback(result); });
                        }
                        continue;
                    }
                    WordRequest wordRequest = item as WordRequest;
                    if (wordRequest != null)
                    {
                        try
                        {
                            if (wordRequest.IsAdd) add(checker, wordRequest.Word);
                            else ignore(checker, wordRequest.Word);
                        }
                        catch { }
                    }
                }
            }
            catch
            {
                _available = false;
            }
            finally
            {
                if (checker != IntPtr.Zero) Marshal.Release(checker);
                if (factory != IntPtr.Zero) Marshal.Release(factory);
                if (factoryObject != null)
                {
                    try { Marshal.ReleaseComObject(factoryObject); }
                    catch { }
                }
                NativeMethods.CoUninitialize();
            }
        }

        private static T Function<T>(IntPtr obj, int slot) where T : class
        {
            IntPtr vtable = Marshal.ReadIntPtr(obj);
            IntPtr function = Marshal.ReadIntPtr(vtable, slot * IntPtr.Size);
            return (T)(object)Marshal.GetDelegateForFunctionPointer(function, typeof(T));
        }

        private static void RunCheck(IntPtr checker, CheckFn check, string[] paragraphs, List<SpellError> result)
        {
            for (int pi = 0; pi < paragraphs.Length; pi++)
            {
                string text = paragraphs[pi] ?? string.Empty;
                if (text.Length == 0) continue;
                IntPtr errors;
                if (check(checker, text, out errors) != 0 || errors == IntPtr.Zero) continue;
                NextErrorFn next = Function<NextErrorFn>(errors, 3);
                IntPtr error;
                int guard = 0;
                while (guard++ < 4000 && next(errors, out error) == 0 && error != IntPtr.Zero)
                {
                    GetIntFn getStart = Function<GetIntFn>(error, 3);
                    GetIntFn getLength = Function<GetIntFn>(error, 4);
                    GetIntFn getAction = Function<GetIntFn>(error, 5);
                    int start, length, action;
                    getStart(error, out start);
                    getLength(error, out length);
                    getAction(error, out action);
                    if (start >= 0 && length > 0 && start + length <= text.Length)
                    {
                        result.Add(new SpellError
                        {
                            Paragraph = pi,
                            Start = start,
                            Length = length,
                            Word = text.Substring(start, length)
                        });
                    }
                    Marshal.Release(error);
                    if (result.Count > 3000) break;
                }
                Marshal.Release(errors);
            }
        }

        private static void RunSuggest(IntPtr checker, SuggestFn suggest, string word, List<string> result)
        {
            IntPtr enumPtr;
            if (suggest(checker, word, out enumPtr) != 0 || enumPtr == IntPtr.Zero) return;
            NextStringFn next = Function<NextStringFn>(enumPtr, 3);
            for (int i = 0; i < 12; i++)
            {
                IntPtr text;
                int fetched;
                if (next(enumPtr, 1, out text, out fetched) != 0 || fetched == 0 || text == IntPtr.Zero) break;
                string value = Marshal.PtrToStringUni(text);
                if (!string.IsNullOrEmpty(value)) result.Add(value);
                Marshal.FreeCoTaskMem(text);
            }
            Marshal.Release(enumPtr);
        }

        private static class NativeMethods
        {
            [DllImport("ole32.dll")]
            public static extern int CoInitializeEx(IntPtr reserved, int coInit);
            [DllImport("ole32.dll")]
            public static extern void CoUninitialize();
        }
    }
}
