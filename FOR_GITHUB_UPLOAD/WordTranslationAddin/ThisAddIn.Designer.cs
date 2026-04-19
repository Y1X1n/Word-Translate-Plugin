namespace WordTranslationAddin
{
    public partial class ThisAddIn : Microsoft.Office.Tools.Word.WordAddInBase
    {
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        protected override void Initialize()
        {
            base.Initialize();
            this.Application = this.GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), "Application");
            Globals.ThisAddIn = this;
            InternalStartup();
        }

        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        protected override void InitializeData()
        {
        }

        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        private void IsInitialized()
        {
        }

        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        private void IsMatched()
        {
        }

        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        private void OnStartup()
        {
            this.Application = this.GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), "Application");
        }

        public Microsoft.Office.Interop.Word.Application Application;
    }

    internal sealed partial class Globals
    {
        private Globals()
        {
        }

        private static ThisAddIn _ThisAddIn;

        internal static ThisAddIn ThisAddIn
        {
            get
            {
                return _ThisAddIn;
            }
            set
            {
                if ((_ThisAddIn == null))
                {
                    _ThisAddIn = value;
                }
                else
                {
                    throw new System.NotSupportedException();
                }
            }
        }
    }
}
