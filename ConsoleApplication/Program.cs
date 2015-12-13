using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using CommandLine;
using CommandLine.Text;

// A sample console application for MeadCo's ScriptX HTML Printing 
//
// There is no copyright, there is no encumberance. Do with this sample source code whatever you like.
//
// What you must do is update the program name, company and copyright and obtain a license in order to
// use binary compilations of the code.
//
// This code is based upon code in our own post build test and verification suite.
//

// Setup
//
// Nuget package used: CommandLineParser
//
// Add references to COM Type libraries:
//  MeadCo ScriptX 7.0 Type Library
//  MeadCo Security Manager 7.1 Type Library
//
// What happens here?
//
// Test application for testing licensing and printing in applications
//
// Requires a test license obtained from the Warehouse
//
// License:
//      license -[tv]
//
// Verifies the app can be licensed, optionally testing threads can be licensed.
//
// Printing:
//      print -[v]j 500
//
//  Selects the XPS Printer Driver
//  Selects output to file (to c:\temp)
//  Prints a url (not cached) n times, each time to a new file name
//  Verifies that all of the created files are of the expected size.
//
// Logging:
//   defaults to verbose
//   -v :: verbose: info, warnings and errors
//   -w :: warning: warnings and errors
//   -e :: error: errors only
//
//

// This sample is configured and built with VS 2013, .NET Framework 4.5. 
//
// It will build with .NET 3.5. To do so:
//      remove the CommandLineParser NuGet package.
//      configure the app to use .NET 3.5
//      add the CommandLineParser package again
//

namespace ScriptXPrinting
{
    class Program
    {
        private const int ErrorCodeExitFailApply = 100;
        private const int ErrorCodeExitTestFail = 200;

        // the number of threads to use (currently used for license installed correctly test only)
        private const int maxThreads = 2;

        // NOTE: To make the application license work must enable native code debugging or disable the Visual Studio hosting process
        // (which can be done on the debug page) - otherwise it is the VS host process that runs and it will not match the license.

        // Each application requires a unique license. The code here is illustrative (and might temporarily work 
        // but the license is short-lived).
        //
        // To test for yourself, please contact us for a license (feedback@meadroid.com). The information required for an application license is 
        // (all entries from the assembly information dialog / assemblyinfo.cs):
        //      Product name
        //      Company (this must also be included in the copyright)
        //  
        // The application copyright must include the company name (as an exact match).
        // 
        // For example
        //  [assembly: AssemblyCompany("Mead & Co Ltd.")]
        //  [assembly: AssemblyProduct("ScriptX Printing Sample")]
        //  [assembly: AssemblyCopyright("Copyright © 2015, Mead & Co Ltd.")]
        
        // To simplify distruibution of test licenses, we allow use from our server ...
        private const string LicenseServerUrl = "http://licenses.meadroid.com";

        // You will be provided with the license GUID and revision.
        private static Guid _licenseGuid = new Guid("{29F95807-5316-40B1-AF54-323D2519949B}");
        private const int LicenseRevision = 0;

        // the license can be obtained from the server ...
        private static readonly string LicenseUrl = LicenseServerUrl + "/download/" + _licenseGuid.ToString() + "/mlf";

        // In production you would use your own server or distrubute the license as an application resource
        // or perhaps on some folder on disk ....
        // private const string LocalLicenseUrl = "c:\\scriptxdev\\applic.mlf";

        // Command line options parsing
        // see this: https://github.com/gsscoder/commandline/wiki/Grammar-Details
        //
        // e.g. app simple -t 
        // or app simple --threaded
        //

        internal class SimpleOptions
        {
            [Option('t', "threaded", DefaultValue = false, HelpText = "Use threads on the test")]
            public bool Threaded { get; set; }

            [Option('v', "verbose", DefaultValue = false, HelpText = "verbose (informational) output")]
            public bool Verbose { get; set; }

            [Option('w', "Warnings", DefaultValue = false, HelpText = "terse (warnings) output")]
            public bool Warnings { get; set; }

            [Option('e', "Errors", DefaultValue = false, HelpText = "terse (errors) output")]
            public bool Errors { get; set; }
        }

        private sealed class PrintOptions : SimpleOptions
        {
            [Option('j', "jobs", DefaultValue = 1, HelpText = "Number of jobs to use")]
            public int Jobs { get; set; }
        }

        private sealed class Options
        {
            [VerbOption("license", HelpText = "Perform a Apply() license test")]
            public SimpleOptions LicenseVerb { get; set; }

            [VerbOption("print", HelpText = "Perform a print test")]
            public PrintOptions PrintVerb { get; set; }

            [HelpVerbOption]
            public string GetUsage(string verb)
            {
                return HelpText.AutoBuild(this, verb);
            }

            [HelpOption]
            public string GetUsage()
            {
                return HelpText.AutoBuild(this);
            }
        }

        // The main thread must be STA for ScriptX.
        [STAThread]
        static void Main(string[] args)
        {
            var options = new Options();

            Logging.Level = Logging.Entry.Info;

            // parse and dispatch the commnand line
            if (!CommandLine.Parser.Default.ParseArgumentsStrict(args, options, Run, () => Console.WriteLine("Parse failed")))
            {
                AppEnd(CommandLine.Parser.DefaultExitCodeFail);
            }

            AppEnd();
        }

        // AppEnd
        //
        // Exits the app with a pause for keypress in debug builds.
        //
        private static void AppEnd(int? exitCode = null)
        {
            Logging.Always("Test completed.");
#if (DEBUG)
            {
                Console.ReadKey(true);
            }
#endif
            if (exitCode.HasValue)
            {
                Environment.Exit(exitCode.Value);
            }
        }

        /// <summary>
        /// Execute the required action
        /// </summary>
        /// <param name="verb">the action to be executed</param>
        /// <param name="options">options for the action</param>
        private static void Run(string verb, object options)
        {
            SimpleOptions simpleOptions = (SimpleOptions)options;

            if (simpleOptions != null)
            {
                if (simpleOptions.Warnings)
                {
                    Logging.Level = Logging.Entry.Warning;
                }
                else if (simpleOptions.Verbose)
                {
                    Logging.Level = Logging.Entry.Info;
                }
                else if (simpleOptions.Errors)
                {
                    Logging.Level = Logging.Entry.Error;
                }
            }

            switch (verb)
            {
                case "license":
                    if (!LicenseTest(simpleOptions))
                    {
                        AppEnd(ErrorCodeExitTestFail);
                    }
                    break;

                case "print":
                    if (!PrintTest((PrintOptions)options))
                    {
                        AppEnd(ErrorCodeExitTestFail);
                    }
                    break;

            }
        }

        /// <summary>
        /// License this process
        /// </summary>
        /// <returns>true if the process has been licensed.</returns>
        private static bool ApplyLicense()
        {
            bool bApplied = false;
            string url = LicenseUrl;

            Logging.Info("applying license {0} ...", url);

            using (Licensing l = new Licensing())
            {
                if (l.ApplyApplicationLicense(url, _licenseGuid, LicenseRevision))
                {
                    Logging.Info("license applied succesfully");
                    bApplied = true;
                }
                else
                {
                    Logging.Error("Failed to apply license: {0}", l.ErrorReason);
                }
            }

            return bApplied;
        }

        #region LicenseTest

        /// <summary>
        /// Tests that the license can be applied to this process -- tests single threaded and multi-threaded use.
        /// </summary>
        /// <param name="options"></param>
        /// <returns></returns>
        private static bool LicenseTest(SimpleOptions options)
        {
            bool bPassed;

            Logging.Always("Starting license test ...");

            bPassed = ApplyLicense();
            if (bPassed)
            {
                if (options == null || !options.Threaded)
                {
                    Console.WriteLine("Single threaded ... ");
                    bPassed = CheckLicensed(0);
                }
                else
                {
                    Console.WriteLine("Multi threaded ... ");

                    for (int threadIndex = 1; threadIndex <= maxThreads; threadIndex++)
                    {
                        int thisIndex = threadIndex;

                        new Thread(() =>
                        {
                            Console.WriteLine(string.Format("Thread[{0}] TestLicense starting ... ", thisIndex));

                            for (int i = 0; i < 10; i++)
                            {
                                bPassed = CheckLicensed(thisIndex);
                                if (!bPassed)
                                    break;

                                Thread.Sleep(500);
                            }

                            Console.WriteLine("Thread TestLicense completed ... ");

                        }).Start();
                    }

                    for (int i = 0; i < 10; i++)
                    {
                        Console.WriteLine("Multi threaded main sleeping ... ");
                        Thread.Sleep(500);
                    }
                }
            }

            Logging.Always("Leaving license test ...: " + bPassed.ToString());
            return bPassed;

        }

        /// <summary>
        /// A simple check that licensing has been applied by calling a lightweight licensed
        /// function that does nothing in its implementation than check a license
        /// is available and then some state flags. In application scenarios
        /// this does start the host process MCSXHost75.exe
        /// </summary>
        /// <param name="id">an informational id (e.g.thread number)</param>
        /// <returns></returns>
        private static bool CheckLicensed(int id)
        {
            bool bLicensed = false;
            using (Printing printing = new Printing())
            {
                Logging.Always(string.Format("[{0}] Factory in use is version: {1}", id, printing.Version));

                if (printing.Printer != null)
                {
                    try
                    {
                        bLicensed = !printing.Printer.IsSpooling();
                        Logging.Always("\tIs licensed: " + bLicensed);
                    }
                    catch (Exception e)
                    {
                        Logging.Error(e.Message);
                    }
                }
                else
                {
                    Logging.Error("No printer object");
                }
            }

            return bLicensed;
        }
        #endregion

        #region PrintTest

        /// <summary>
        /// A class to box the data we want to pass to scriptx for return as callback data
        /// </summary>
        private class BoxCallbackObject
        {
            public int _j;
            public BoxCallbackObject(int jobNo) { _j = jobNo; }
        }

        /// <summary>
        /// Implementation of callbacks from PrintHtmlEx 
        /// </summary>
        private class PrintProgress : PrintHtmlExCallback
        {
            public PrintProgress()
            {
                ErrorJobsList = new List<int>();
                JobsDone = 0;
            }

            public int TotalJobs { get; set; }
            public int JobsDone { private set; get; }
            public List<int> ErrorJobsList { private set; get; }

            public bool Complete
            {
                get { return JobsDone + ErrorJobsList.Count >= TotalJobs; }
            }

            /// <summary>
            /// Method called on events from the PrintHtmlEx method execution
            /// </summary>
            /// <param name="status">The event status</param>
            /// <param name="statusData">information relevant to the event</param>
            /// <param name="callbackData">the object passed to PrintHtmlEx</param>
            public override void Notified(PrintHtmlExCallback.CallBackCode status, string statusData, object callbackData)
            {
                int jobNo = ((BoxCallbackObject)callbackData)._j;

                Logging.Info("Print job " + jobNo.ToString() + " - status: " + Enum.GetName(typeof(PrintHtmlExCallback.CallBackCode), status) + " - " + statusData);
                switch (status)
                {
                    case CallBackCode.BatchCompleted:
                        JobsDone++;
                        break;

                    case CallBackCode.BatchAbandon:
                    case CallBackCode.BatchError:
                        ErrorJobsList.Add(jobNo);
                        break;

                }
            }
        }

        /// <summary>
        /// We print to file using the XPS Printer Driver and then check the size of the print
        /// files to check the print worked (crude but effective - output size does vary).
        /// </summary>
        /// <param name="options"></param>
        /// <returns></returns>
        private static bool PrintTest(PrintOptions options)
        {
            bool bPassed;

            // WARNING:: The folder musts exist. If it doesnt you'll see a confusing RPC failure error when printing.
            const string outputFilename = "c:\\temp\\print_{0}.xps";

#if (WIN64)
            const int minimumAcceptableFileSize = 112095;
            const int maximumAcceptableFileSize = 112128;
        
#else
            const int minimumAcceptableFileSize = 168450;
            const int maximumAcceptableFileSize = 169500;
#endif
            Logging.Always("Starting print test ...");

            bPassed = ApplyLicense();
            if (bPassed)
            {
                if (options == null || !options.Threaded)
                {
                    int jobs = options == null ? 1 : options.Jobs;

                    using (var printing = new Printing())
                    {
                        Logging.Info(string.Format("Factory in use is version: {0}, will queue {1} jobs", printing.Version, jobs));

                        var printer = printing.Printer;
                        var progressMonitor = new PrintProgress { TotalJobs = jobs };

                        try
                        {
                            printer.footer = "&b&b&p of &P";
                            printer.currentPrinter = "Microsoft XPS Document Writer";

                            Logging.Info("The printer that will be used is: " + printer.printer);

                            // Queue all the prints into Scriptx -- it has its own queue manager that
                            // will pull down the document to print and then print it. Note that printing
                            // runs in a separate process.
                            for (int job = 1; bPassed && job <= jobs; job++)
                            {
                                string jobFileName = string.Format(outputFilename, job);

                                // When printing to file, the file must not already exist.
                                // ScriptX will not delete it for us/
                                if (File.Exists(jobFileName))
                                {
                                    File.Delete(jobFileName);
                                }
                                printer.printToFileName = jobFileName;

                                // we can have different "everything" per job, but just change the header
                                printer.header = string.Format("Print from Application&b&bJob: {0}",job);

                                // print a well known url
                                bPassed = printer.PrintHTMLEx(
                                    "http://scriptxsamples.meadroid.com/Licensed/salesinfo?_=" +
                                    DateTime.Now.Millisecond, false,
                                    progressMonitor, new BoxCallbackObject(job));

                                if (!bPassed)
                                {
                                    Logging.Error("Failed to queue job {0}", job);
                                }
                            }

                            // wait for the queue to complete, we give it 3 seconds per job.
                            //
                            // If you leave it as infinite the app could hang
                            //
                            if (!printer.WaitForSpoolingComplete(3000 * jobs))
                            {
                                Logging.Error("Wait abandoned, isSpooling? " + printer.IsSpooling());
                                bPassed = false;
                            }
                            else
                            {
                                // Success, check the output sizes are in range ...
                                Logging.Always(string.Format("Jobs completed: {0}, jobs in error: {1}",
                                    progressMonitor.JobsDone, progressMonitor.ErrorJobsList.Count));
                                for (int job = 1; bPassed && job <= jobs; job++)
                                {
                                    if (!progressMonitor.ErrorJobsList.Contains(job))
                                    {
                                        string jobFileName = string.Format(outputFilename, job);
                                        long fileSize = new FileInfo(jobFileName).Length;

                                        if (fileSize < minimumAcceptableFileSize ||
                                            fileSize > maximumAcceptableFileSize)
                                        {
                                            Logging.Error(
                                                string.Format("The output file size {0} on job {1} is wrong",
                                                    fileSize,
                                                    job));
                                            // bPassed = false;
                                        }
                                        else
                                        {
                                            Logging.Info("The output file on job {0} is correct", job);
                                        }
                                    }


                                    if (bPassed)
                                        bPassed = progressMonitor.JobsDone == 0 || (progressMonitor.Complete && progressMonitor.ErrorJobsList.Count == 0);
                                }
                            }

                        }
                        catch (Exception e)
                        {
                            Logging.Error("An error occurred: " + e.Message);
                            bPassed = false;
                        }
                    }
                }
            }

            Logging.Always("completed print test ... " + bPassed.ToString());

            return bPassed;
        }
        #endregion


    }
}
