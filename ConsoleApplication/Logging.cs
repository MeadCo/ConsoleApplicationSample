using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScriptXPrinting
{
    internal static class Logging
    {
        public enum Entry
        {
            Info,
            Warning,
            Error,
            Always = 100
        }

        public static Entry Level = Entry.Info;

        private static void WriteEntry(Entry entryType)
        {
            switch (entryType)
            {
                case Entry.Error:
                    Console.Write("*** ERROR *** > ");
                    break;

                case Entry.Warning:
                    Console.Write("* Warning * > ");
                    break;
            }
        }

        private static void WriteLine(Entry entryType, string strText)
        {
            if (entryType >= Level)
            {
                WriteEntry(entryType);
                Console.WriteLine(strText);
            }
        }

        private static void WriteLine(Entry entryType, string strFormat, params object[] argsObjects)
        {
            WriteLine(entryType, string.Format(strFormat, argsObjects));
        }

        public static void Write(Entry entryType, string strText)
        {
            if (entryType >= Level)
            {
                WriteEntry(entryType);
                Console.Write(strText);
            }
        }

        public static void Write(Entry entryType, string strFormat, params object[] argsObjects)
        {
            Write(entryType, string.Format(strFormat, argsObjects));
        }

        public static void Always(string strText)
        {
            WriteLine(Entry.Always, strText);
        }

        public static void Always(string strFormat, params object[] argsObjects)
        {
            Always(string.Format(strFormat, argsObjects));
        }

        public static void Info(string strText)
        {
            WriteLine(Entry.Info, strText);
        }

        public static void Info(string strFormat, params object[] argsObjects)
        {
            Info(string.Format(strFormat, argsObjects));
        }

        public static void Warning(string strText)
        {
            WriteLine(Entry.Warning, strText);
        }

        public static void Warning(string strFormat, params object[] argsObjects)
        {
            Warning(string.Format(strFormat, argsObjects));
        }

        public static void Error(string strText)
        {
            WriteLine(Entry.Error, strText);
        }

        public static void Error(string strFormat, params object[] argsObjects)
        {
            Error(string.Format(strFormat, argsObjects));
        }
    }
}
