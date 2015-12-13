using System;
using System.Globalization;
using System.Reflection;

namespace ScriptXPrinting
{
    //
    // see: http://limbioliong.wordpress.com/2011/08/07/misunderstanding-intptr-and-system-object/
    //

    // PrintHtmlExCallback implements an IDispatch interface that can be passed
    // to PrintHtmlEx as the call back paramater.
    //
    // Usage :: 
    //
    // Derive from PrintHtmlExCallback and override the the implementation of Notified
    // Note that if an framework object is passed as callbackData you must box it..
    //
    // PrintHtmlEx progress callback implementation
    //
    //private class PrintProgress : PrintHtmlExCallback
    //{
    //    /// <summary>
    //    /// Method called on events from the PrintHtmlEx method execution
    //    /// </summary>
    //    /// <param name="status">The event status</param>
    //    /// <param name="statusData">information relevant to the event</param>
    //    /// <param name="callbackData">the object passed to PrintHtmlEx</param>
    //    public override void Notified(PrintHtmlExCallback.CallBackCode status, string statusData, object callbackData)
    //    {
    //        ((BoxObject)callbackData)._M.Log("Status: " + Enum.GetName(typeof(PrintHtmlExCallback.CallBackCode), status) + " - " + statusData);
    //    }
    //}

    //private class BoxObject
    //{
    //    public sample_master _M;
    //    public BoxObject(sample_master M) { _M = M; }
    //}

    // ==>  printer.PrintHTMLEx(sReportUrl, false, new PrintProgress(), new BoxObject(Master));

    /// <summary>
    /// Implementation of a callback object for use with the MeadCo ScriptX PrintHtmlEx API
    /// </summary>
    class PrintHtmlExCallback : IReflect
    {
        private static class Utils
        {
            public static Exception NotImplemented(string message)
            {
                return new Exception(message);
            }            
        }

        public enum CallBackCode
        {
            // queue call back opcodes ...
            BatchQueued = 1,
            BatchStarted = 2,
            BatchDownloading = 3,
            BatchDownloaded = 4,
            BatchPrinting = 5,
            BatchCompleted = 6,
            BatchPaused = 7,
            BatchPrintpdf = 8,

            BatchError = -1,
            BatchAbandon = -2
        }

        public virtual void Notified(CallBackCode status, string statusData, object callbackData)
        {
        }

        #region IReflect Members

        // Unimplemented methods
        public FieldInfo GetField(string name, BindingFlags bindingAttr) { throw Utils.NotImplemented("In ManagedIDispatch.GetField"); }
        public MemberInfo[] GetMember(string name, BindingFlags bindingAttr) { throw Utils.NotImplemented("In ManagedIDispatch.GetMember"); }
        public MethodInfo GetMethod(string name, BindingFlags bindingAttr) { throw Utils.NotImplemented("In ManagedIDispatch.GetMethod"); }
        public MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers) { throw Utils.NotImplemented("In ManagedIDispatch.GetMethod"); }
        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
        {
            throw Utils.NotImplemented("In ManagedIDispatch.GetProperty");
        }
        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr) { throw Utils.NotImplemented("In ManagedIDispatch.GetProperty"); }
        public Type UnderlyingSystemType { get { throw Utils.NotImplemented("In ManagedIDispatch.UnderlyingSystemType"); } }

        // Methods that need to be implemented 
        //public FieldInfo[] GetFields(BindingFlags bindingAttr) { return new FieldInfo[0]; }
        //public MemberInfo[] GetMembers(BindingFlags bindingAttr) { return new MemberInfo[0]; }
        //public PropertyInfo[] GetProperties(BindingFlags bindingAttr) { return new PropertyInfo[0]; }
        //public MethodInfo[] GetMethods(BindingFlags bindingAttr) { return new MethodInfo[0]; }

        public FieldInfo[] GetFields(BindingFlags bindingAttr)
        {
            return GetType().GetFields(bindingAttr);
        }

        public MethodInfo[] GetMethods(BindingFlags bindingAttr)
        {
            MethodInfo[] r = GetType().GetMethods(bindingAttr);
            return r;
        }

        public PropertyInfo[] GetProperties(BindingFlags bindingAttr)
        {
            return GetType().GetProperties(bindingAttr);
        }

        public MemberInfo[] GetMembers(BindingFlags bindingAttr)
        {
            throw new NotImplementedException();
        }


        public object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
        {
            
            switch (name)
            {
                case "[DISPID=0]":
                    Notified((CallBackCode)args[0], args[1] != null ? args[1].ToString() : "", args[2]);
                    return this;
                default:
                    throw Utils.NotImplemented("This is unreachable");
            };
        }

        #endregion

    }
}
