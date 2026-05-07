namespace MailMergeKit.WordAddin
{
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    internal sealed partial class Globals
    {
        private Globals()
        {
        }

        private static ThisAddIn _ThisAddIn;
        private static global::Microsoft.Office.Tools.Word.ApplicationFactory _factory;

        internal static ThisAddIn ThisAddIn
        {
            get { return _ThisAddIn; }
            set
            {
                if (_ThisAddIn == null)
                    _ThisAddIn = value;
                else
                    throw new System.NotSupportedException();
            }
        }

        internal static global::Microsoft.Office.Tools.Word.ApplicationFactory Factory
        {
            get { return _factory; }
            set
            {
                if (_factory == null)
                    _factory = value;
                else
                    throw new System.NotSupportedException();
            }
        }
    }
}
