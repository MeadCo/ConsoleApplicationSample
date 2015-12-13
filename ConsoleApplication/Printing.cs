using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScriptXPrinting
{
    /// <summary>
    /// A simple wrapper in the classic 'factory' and 'factory.printing' objects 
    /// instead of code such as factory.printing.PrintHTML() we write 
    ///     var factory = new Printing();
    ///     factory.Printing.PrintHtml()
    /// 
    /// Some simple wrappers on functions available on factory are also provided.
    /// </summary>
    class Printing : IDisposable
    {
        private ScriptX.Factory _factory = null;
        private ScriptX.printing _printer = null;
        private string _lastError = string.Empty;
        private bool _disposed = false;

        public Printing()
        {
        }

        ~Printing()
        {
            Dispose(false);
        }

        private void Initialise()
        {
            if (_printer == null)
            {
                try
                {
                    _factory = new ScriptX.Factory();
                    _printer = (ScriptX.printing)_factory.printing;
                }
                catch (Exception e)
                {
                    _lastError = e.Message;
                }
            }

        }

        public string ErrorReason
        {
            get
            {
                return _lastError;
            }
        }

        public string Version
        {
            get
            {
                string sv = string.Empty;
                object vmaj = new object();
                object vmin = new object();
                object vbuild = new object();
                object vpart = new object();

                Initialise();
                if (_factory != null)
                {
                    vmaj = "";
                    vmin = "";
                    vbuild = "";
                    vpart = "";
                    _factory.GetComponentVersion("ScriptX.Factory", ref vmaj, ref vmin, ref vbuild, ref vpart);
                    sv = vmaj.ToString() + "." + vmin.ToString() + "." + vbuild.ToString() + "." + vpart.ToString();
                }

                return sv;
            }
        }

        public ScriptX.printing Printer
        {
            get
            {
                Initialise();
                return _printer;
            }
        }

        public void Close()
        {
            Dispose();
        }

        // IDisposable

        // Implement IDisposable.
        // Do not make this method virtual.
        // A derived class should not be able to override this method.
        public void Dispose()
        {
            Dispose(true);
            // Take yourself off the Finalization queue 
            // to prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }

        // Dispose(bool disposing) executes in two distinct scenarios.
        // If disposing equals true, the method has been called directly
        // or indirectly by a user's code. Managed and unmanaged resources
        // can be disposed.
        // If disposing equals false, the method has been called by the 
        // runtime from inside the finalizer and you should not reference 
        // other objects. Only unmanaged resources can be disposed.
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this._disposed)
            {
                // Note that this is not thread safe.
                // Another thread could start disposing the object
                // after the managed resources are disposed,
                // but before the disposed flag is set to true.
                // If thread safety is necessary, it must be
                // implemented by the client.

                if (_factory != null)
                {
                    int i;

                    _factory.Shutdown();

                    if (_printer != null)
                    {
                        i = System.Runtime.InteropServices.Marshal.ReleaseComObject(_printer);
                        while (i > 0)
                        {
                            i = System.Runtime.InteropServices.Marshal.ReleaseComObject(_printer);
                        }
                    }

                    i = System.Runtime.InteropServices.Marshal.ReleaseComObject(_factory);
                    while (i > 0)
                    {
                        i = System.Runtime.InteropServices.Marshal.ReleaseComObject(_factory);
                    }
                }
            }

            _disposed = true;
        }

    }
}
