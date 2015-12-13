using System;

namespace ScriptXPrinting
{
    /// <summary>
    /// A wrapper on the COM Object MeadCo.SecMgr
    /// The object provides licensing to a process. This wrapper simplifies 
    /// instantiating and disposing of the object.
    /// </summary>
    class Licensing : IDisposable
    {
        private SecMgr.SecMgr _secMgr = null;
        private string _errorDescription = string.Empty;
        private bool _licensed = false;
        private bool _disposed = false;

        public Licensing()
        {
        }

        ~Licensing()
        {
            Dispose(false);
        }

        /// <summary>
        /// A description of why atttempting to use the license failed.
        /// </summary>
        public string ErrorReason
        {
            get
            {
                return _errorDescription;
            }
        }

        /// <summary>
        /// Attempts to apply a license that has already been installed to the machine
        /// to this process. 
        /// </summary>
        /// <param name="licenseGuid">GUID of the license as supplied by MeadCo.</param>
        /// <param name="licenseRevision">Revision of the license as supplied by MeadCo</param>
        /// <returns>true if the application is now licensed, if failed see the ErrorReason property</returns>
        public bool UseMachineLicense(Guid licenseGuid, int licenseRevision)
        {
            return ApplyLicense("",licenseGuid,licenseRevision);
        }

        /// <summary>
        /// Attempts to download/load the license from the given location (url/filepath) and use it with 
        /// this process. The license must describe this application (name and company)
        /// </summary>
        /// <param name="licensePath">Location of the license file (sxlic.mlf as supplied by MeadCo)</param>
        /// <param name="licenseGuid">GUID of the license as supplied by MeadCo.</param>
        /// <param name="licenseRevision">Revision of the license as supplied by MeadCo</param>
        /// <returns>true if the application is now licensed, if failed see the ErrorReason property</returns>
        public bool ApplyApplicationLicense(string licensePath, Guid licenseGuid, int licenseRevision)
        {
            return ApplyLicense(licensePath, licenseGuid, licenseRevision);
        }

        /// <summary>
        /// Attempts to download/load the license from the given location (url/filepath) and use it with 
        /// this process.
        /// </summary>
        /// <param name="licensePath">Location of the license file (sxlic.mlf as supplied by MeadCo)</param>
        /// <param name="licenseGuid">GUID of the license as supplied by MeadCo.</param>
        /// <param name="licenseRevision">Revision of the license as supplied by MeadCo</param>
        /// <returns>true if the application is now licensed, if failed see the ErrorReason property</returns>
        private bool ApplyLicense(string licensePath, Guid licenseGuid, int licenseRevision)
        {
            if (!_licensed)
            {
                try
                {
                    _secMgr = new SecMgr.SecMgr();
                    _secMgr.Apply(licensePath, "{" + licenseGuid.ToString() + "}", licenseRevision);
                    _licensed = true;
                }
                catch (Exception e)
                {
                    _errorDescription = "Unable to apply license: " + e.Message;
                }
            }
            return _licensed;
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

                if (_secMgr != null)
                {
                    // secmgr is relatively lightweight and one could wait for garbage collection
                    int i;

                    i = System.Runtime.InteropServices.Marshal.ReleaseComObject(_secMgr);
                    while (i > 0)
                    {
                        i = System.Runtime.InteropServices.Marshal.ReleaseComObject(_secMgr);
                    }
                }
            }

            _disposed = true;
        }

    }

}
