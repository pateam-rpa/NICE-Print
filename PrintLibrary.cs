using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Printing;
using System.Printing;
using System.IO;

using Direct.Shared;
using log4net;

using RawPrint;
using System.Management;

namespace Pateam.Print.Library
{
    [DirectSealed]
    [DirectDom("PATeam Print Library")]
    [ParameterType(false)]
    public class PrintLibrary
    {
        public static readonly ILog logArchitect = LogManager.GetLogger(Loggers.LibraryObjects);

        [DirectDom]
        public delegate void PrintJobStarted (IDirectComponentBase sender, JobStartedEventArgs e);

        [DirectDom("Printing Job Started")]
        public static event PrintJobStarted PrintJobStartedEvent;

        [DirectDom("Prints file from Path")]
        [DirectDomMethod("Print {file} at {folder} using {Printer Name} ")]
        [MethodDescriptionAttribute("Prints file from Path")]
        public static void PrintDocument(string fileName, string pathToFile, string PrinterName)
        {
            if (logArchitect.IsDebugEnabled)
            {
                logArchitect.DebugFormat("PrintLibrary.PrintDocument - entered with filename {0} and path {1}", fileName, pathToFile);
            }

            //Input validation
            if (!File.Exists(Path.Combine(pathToFile, fileName)))
            {
                logArchitect.ErrorFormat("PrintLibrary.PrintDocument - Can't find file at {0}.", Path.Combine(pathToFile, fileName));
            }
            else if (string.IsNullOrEmpty(PrinterName))
            {
                logArchitect.Error("PrintLibrary.PrintDocument - must specify printer name.");
            }

            //Check if the printer exists.
            bool exists=false;
            foreach (string printerName in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                if (printerName == PrinterName)
                {
                    exists = true;
                    break;
                }
            }
            if (!exists){
                logArchitect.ErrorFormat("PrintLibrary.PrintDocument - can't find printer with name {0}.", PrinterName);
            }


            try
            {
                IPrinter printer = new Printer();
                printer.OnJobCreated += (sender, eventArgs) => {
                    //Console.WriteLine("Job started.");
                    JobStartedEventArgs eArgs = new JobStartedEventArgs((int)eventArgs.Id);
                    PrintJobStartedEvent(null, eArgs);

                };
                printer.PrintRawFile(PrinterName, Path.Combine(pathToFile, fileName), fileName);
                

            }
            catch (Exception ex)
            {
                if (logArchitect.IsErrorEnabled)
                {
                    logArchitect.Error("PrintLibrary.PrintDocument - failed with exception", ex);
                }
            }
        }

        [DirectDom("Check Print Job Status")]
        [DirectDomMethod("Check Status of {Job ID} in {Printr Name}")]
        [MethodDescriptionAttribute("Checks Print Job Status")]
        public static string CheckJobStatus(int id, string qName)
        {
            if (logArchitect.IsDebugEnabled)
            {
                logArchitect.DebugFormat("PrintLibrary.CheckJobStatus - Given parameters are job id: {0} Printer name: {1}", id, qName);
            }
            PrintServer printServer = new PrintServer();
            PrintQueueCollection myPrintQueues = printServer.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            foreach (PrintQueue pq in myPrintQueues)
            {
               if (logArchitect.IsDebugEnabled)
                {
                    logArchitect.DebugFormat("PrintLibrary.CheckJobStatus - queue name is {0}", pq.Name);
                }
               if (pq.FullName == qName)
                {
                    pq.Refresh();
                    PrintJobInfoCollection jobs = pq.GetPrintJobInfoCollection();
                    foreach (PrintSystemJobInfo job2 in jobs)
                    {
                        if (logArchitect.IsDebugEnabled)
                        {
                            logArchitect.DebugFormat("PrintLibrary.CheckJobStatus - job id is {0}", job2.JobIdentifier);
                        }
                        if (job2.JobIdentifier == id)
                        {
                            return SpotTroubleUsingJobAttributes(job2);
                        }
                    }
                   // pq.Refresh();
                   // PrintSystemJobInfo job = pq.GetJob(id);
                  //  return SpotTroubleUsingJobAttributes(job);
                }
               
            }
            // suspected to not work
            // PrintQueue pq = printServer.GetPrintQueue(qName);
            // pq.Refresh();
            //PrintSystemJobInfo job = pq.GetJob(id);
            return "Job id not queue";
        }

        internal static string SpotTroubleUsingJobAttributes(PrintSystemJobInfo theJob)
        {
            if ((theJob.JobStatus & PrintJobStatus.Blocked) == PrintJobStatus.Blocked)
            {
                return "The job is blocked.";
            }
            if (((theJob.JobStatus & PrintJobStatus.Completed) == PrintJobStatus.Completed)
                ||
                ((theJob.JobStatus & PrintJobStatus.Printed) == PrintJobStatus.Printed))
            {
                return "The job has finished.";
            }
            if (((theJob.JobStatus & PrintJobStatus.Deleted) == PrintJobStatus.Deleted)
                ||
                ((theJob.JobStatus & PrintJobStatus.Deleting) == PrintJobStatus.Deleting))
            {
                return "The user or someone with administration rights to the queue has deleted the job. It must be resubmitted.";
            }
            if ((theJob.JobStatus & PrintJobStatus.Error) == PrintJobStatus.Error)
            {
                return "The job has errored.";
            }
            if ((theJob.JobStatus & PrintJobStatus.Offline) == PrintJobStatus.Offline)
            {
                return "The printer is offline. Have user put it online with printer front panel.";
            }
            if ((theJob.JobStatus & PrintJobStatus.PaperOut) == PrintJobStatus.PaperOut)
            {
                return "The printer is out of paper of the size required by the job. Have user add paper.";
            }

            if (((theJob.JobStatus & PrintJobStatus.Paused) == PrintJobStatus.Paused)
                ||
                ((theJob.HostingPrintQueue.QueueStatus & PrintQueueStatus.Paused) == PrintQueueStatus.Paused))
            {
                //HandlePausedJob(theJob);
                return "The job is printing paused.";
            }

            if ((theJob.JobStatus & PrintJobStatus.Printing) == PrintJobStatus.Printing)
            {
                return "The job is printing now.";
            }
            if ((theJob.JobStatus & PrintJobStatus.Spooling) == PrintJobStatus.Spooling)
            {
                return "The job is spooling now.";
            }
            if ((theJob.JobStatus & PrintJobStatus.UserIntervention) == PrintJobStatus.UserIntervention)
            {
                return "The printer needs human intervention.";
            }

            return "";

        }

        internal static void HandlePausedJob(PrintSystemJobInfo theJob)
        {
            // If there's no good reason for the queue to be paused, resume it and 
            // give user choice to resume or cancel the job.
            Console.WriteLine("The user or someone with administrative rights to the queue" +
                 "\nhas paused the job or queue." +
                 "\nResume the queue? (Has no effect if queue is not paused.)" +
                 "\nEnter \"Y\" to resume, otherwise press return: ");
            String resume = Console.ReadLine();
            if (resume == "Y")
            {
                theJob.HostingPrintQueue.Resume();

                // It is possible the job is also paused. Find out how the user wants to handle that.
                Console.WriteLine("Does user want to resume print job or cancel it?" +
                    "\nEnter \"Y\" to resume (any other key cancels the print job): ");
                String userDecision = Console.ReadLine();
                if (userDecision == "Y")
                {
                    theJob.Resume();
                }
                else
                {
                    theJob.Cancel();
                }
            }//end if the queue should be resumed

        }//end HandlePausedJob

        [DirectDom("Get available printers")]
        [DirectDomMethod("Get available printers")]
        [MethodDescriptionAttribute("Gets the available printers.")]
        public static DirectCollection<string> GetAvailablePrinterNames()
        {
            DirectCollection<string> Printers = new DirectCollection<string>();
            foreach (string printerName in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                Printers.Add(printerName);
            }
            return Printers;
        }

        internal static void FurtherCheckPrinterStatus(ref string status, string _printerName, ManagementScope scope, ManagementObjectCollection MOC)
        {

            // Set management scope
            

            // Select Printers from WMI Object Collections
            

            string printerName = "";
            foreach (ManagementObject printer in MOC)
            {
                printerName = printer["Name"].ToString();
                if (printerName.Equals(_printerName))
                {
                    PropertyDataCollection pdc = printer.Properties;
                   // Console.WriteLine("Printer = " + printer["Name"]);
                    if (printer["WorkOffline"].ToString().ToLower().Equals("true"))
                    {
                        // printer is offline by user
                        status = "Is off line.";
                    }
                    else
                    {
                        // printer is not offline
                        //Console.WriteLine("Your Plug-N-Play printer is connected.");
                    }
                }
            }
        }
            
        [DirectDom("Get available printers statuses")]
        [DirectDomMethod("Get available printers statuses")]
        [MethodDescriptionAttribute("Get available printers statuses if the status is empty, you can print.")]
        public static DirectCollection<PrinterStatus> GetAvailablePrinterQs()
        {
            DirectCollection<PrinterStatus> Queues = new DirectCollection<PrinterStatus>();

            //PrintQueue printQueue = null;

            PrintServer printServer = new PrintServer();
            PrintQueueCollection myPrintQueues = printServer.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            ManagementScope scope = new ManagementScope(@"\root\cimv2");
            scope.Connect();
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
            ManagementObjectCollection MOC = searcher.Get();
            
            foreach (PrintQueue pq in myPrintQueues)
            {

                pq.Refresh();
                string status = "";
                SpotTroubleUsingQueueAttributes(ref status, pq);
                if (string.IsNullOrEmpty(status))
                {
                    FurtherCheckPrinterStatus(ref status, pq.Name, scope, MOC);
                }
                Queues.Add(new PrinterStatus(pq.FullName, status));
            }

            //LocalPrintServer localPrintServer = new LocalPrintServer();

            //// Retrieving collection of local printer on user machine
            //PrintQueueCollection localPrinterCollection =
            //    localPrintServer.GetPrintQueues();

            //System.Collections.IEnumerator localPrinterEnumerator =
            //localPrinterCollection.GetEnumerator();

            //if (localPrinterEnumerator.MoveNext())
            //{
            //    // Get PrintQueue from first available printer
            //    printQueue = (PrintQueue)localPrinterEnumerator.Current;
            //}
            //else
            //{
            //    // No printer exist, return null PrintTicket
            //    return null;
            //}

            //foreach (string printerName in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            //{
            //    Printers.Add(printerName);
            //}
            return Queues;
        }

        // Check for possible trouble states of a printer using the flags of the QueueStatus property
        internal static void SpotTroubleUsingQueueAttributes(ref String statusReport, PrintQueue pq)
        {
            if ((pq.QueueStatus & PrintQueueStatus.PaperProblem) == PrintQueueStatus.PaperProblem)
            {
                statusReport = statusReport + "Has a paper problem. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.NoToner) == PrintQueueStatus.NoToner)
            {
                statusReport = statusReport + "Is out of toner. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.DoorOpen) == PrintQueueStatus.DoorOpen)
            {
                statusReport = statusReport + "Has an open door. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Error) == PrintQueueStatus.Error)
            {
                statusReport = statusReport + "Is in an error state. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.NotAvailable) == PrintQueueStatus.NotAvailable)
            {
                statusReport = statusReport + "Is not available. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Offline) == PrintQueueStatus.Offline)
            {
                statusReport = statusReport + "Is off line. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.OutOfMemory) == PrintQueueStatus.OutOfMemory)
            {
                statusReport = statusReport + "Is out of memory. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.PaperOut) == PrintQueueStatus.PaperOut)
            {
                statusReport = statusReport + "Is out of paper. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.OutputBinFull) == PrintQueueStatus.OutputBinFull)
            {
                statusReport = statusReport + "Has a full output bin. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.PaperJam) == PrintQueueStatus.PaperJam)
            {
                statusReport = statusReport + "Has a paper jam. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Paused) == PrintQueueStatus.Paused)
            {
                statusReport = statusReport + "Is paused. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.TonerLow) == PrintQueueStatus.TonerLow)
            {
                statusReport = statusReport + "Is low on toner. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.UserIntervention) == PrintQueueStatus.UserIntervention)
            {
                statusReport = statusReport + "Needs user intervention. ";
            }

            // Check if queue is even available at this time of day
            // The method below is defined in the complete example.
            ReportAvailabilityAtThisTime(ref statusReport, pq);
        }

        private static void ReportAvailabilityAtThisTime(ref String statusReport, PrintQueue pq)
        {
            if (pq.StartTimeOfDay != pq.UntilTimeOfDay) // If the printer is not available 24 hours a day
            {
                DateTime utcNow = DateTime.UtcNow;
                Int32 utcNowAsMinutesAfterMidnight = (utcNow.TimeOfDay.Hours * 60) + utcNow.TimeOfDay.Minutes;

                // If now is not within the range of available times . . .
                if (!((pq.StartTimeOfDay < utcNowAsMinutesAfterMidnight)
                   &&
                   (utcNowAsMinutesAfterMidnight < pq.UntilTimeOfDay)))
                {
                    statusReport = statusReport + " Is not available at this time of day. ";
                }
            }
        }
    }

    [DirectDom]
    [DirectSealed]
    public class JobStartedEventArgs : DirectEventArgs
    {
        int jobId = 0;

        //string status = string.Empty;

        public JobStartedEventArgs(int jobId)
        {
            this.jobId = jobId;
        }

        [DirectDom("Job ID")]
        public int JobId
        {
            get { return jobId; }
        }

    }

    [DirectDom("Printer Status", "PATeam Print Library", false)]
    public class PrinterStatus : DirectComponentBase
    {
        protected PropertyHolder<string> _Name = new PropertyHolder<string>("Name");
        protected PropertyHolder<string> _Status = new PropertyHolder<string>("Status");

        public PrinterStatus()
        {

        }

        public PrinterStatus(string name, string value)
        {
            this.Name = name;
            this.Status = value;
        }
        public PrinterStatus(IProject project)
        : base(project)
        {

        }


        [DirectDom("Printer Name")]
        public string Name
        {
            get { return _Name.TypedValue; }
            set { this._Name.TypedValue = value; }
        }


        [DirectDom("Status")]
        public string Status
        {
            get { return _Status.TypedValue; }
            set { this._Status.TypedValue = value; }
        }



    }
}
