using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace HyperVSamples
{
    enum JobState
    {
        New = 2,
        Starting = 3,
        Running = 4,
        Suspended = 5,
        ShuttingDown = 6,
        Completed = 7,
        Terminated = 8,
        Killed = 9,
        Exception = 10,
        CompletedWithWarnings = 32768
    }

    class WmiUtility
    {
        private static bool
        IsJobComplete(
            object jobStateObj)
        {
            JobState jobState = (JobState)((ushort)jobStateObj);

            return (jobState == JobState.Completed) ||
                (jobState == JobState.CompletedWithWarnings) || (jobState == JobState.Terminated) ||
                (jobState == JobState.Exception) || (jobState == JobState.Killed);
        }

        private static bool
        IsJobSuccessful(
            object jobStateObj)
        {
            JobState jobState = (JobState)((ushort)jobStateObj);

            return (jobState == JobState.Completed) || (jobState == JobState.CompletedWithWarnings);
        }

        public static void
        PrintMsvmErrors(
            ManagementObject job)
        {
            string[] errorList;

            using (ManagementBaseObject inParams = job.GetMethodParameters("GetErrorEx"))
            using (ManagementBaseObject outParams = job.InvokeMethod("GetErrorEx", inParams, null))
            {
                if ((uint)outParams["ReturnValue"] != 0)
                {
                    throw new ManagementException(string.Format(CultureInfo.CurrentCulture,
                                                                "GetErrorEx() call on the job failed"));
                }

                errorList = (string[])outParams["Errors"];
            }

            if (errorList == null)
            {
                Console.WriteLine("No errors found.");
                return;
            }

            Console.WriteLine("Detailed errors: \n");

            foreach (string error in errorList)
            {
                string errorSource = string.Empty;
                string errorMessage = string.Empty;
                int propId = 0;

                XmlReader reader = XmlReader.Create(new StringReader(error));

                while (reader.Read())
                {
                    if (reader.Name.Equals("PROPERTY", StringComparison.OrdinalIgnoreCase))
                    {
                        propId = 0;

                        if (reader.HasAttributes)
                        {
                            string propName = reader.GetAttribute(0);

                            if (propName.Equals("ErrorSource", StringComparison.OrdinalIgnoreCase))
                            {
                                propId = 1;
                            }
                            else if (propName.Equals("Message", StringComparison.OrdinalIgnoreCase))
                            {
                                propId = 2;
                            }
                        }
                    }
                    else if (reader.Name.Equals("VALUE", StringComparison.OrdinalIgnoreCase))
                    {
                        if (propId == 1)
                        {
                            errorSource = reader.ReadElementContentAsString();
                        }
                        else if (propId == 2)
                        {
                            errorMessage = reader.ReadElementContentAsString();
                        }

                        propId = 0;
                    }
                    else
                    {
                        propId = 0;
                    }
                }

                Console.WriteLine("Error Message: {0}", errorMessage);
                Console.WriteLine("Error Source:  {0}\n", errorSource);
            }
        }

        public static bool
        ValidateOutput(ManagementBaseObject outputParameters, ManagementScope scope, bool throwIfFailed, bool printErrors)
        {
            bool succeeded = true;
            string errorMessage = "The method call failed.";

            if ((uint)outputParameters["ReturnValue"] == 4096)
            {
                //
                // The method invoked an asynchronous operation. Get the Job object
                // and wait for it to complete. Then we can check its result.
                //

                using (ManagementObject job = new ManagementObject((string)outputParameters["Job"]))
                {
                    job.Scope = scope;

                    while (!IsJobComplete(job["JobState"]))
                    {
                        Thread.Sleep(TimeSpan.FromSeconds(1));

                        // 
                        // ManagementObjects are offline objects. Call Get() on the object to have its
                        // current property state.
                        //
                        job.Get();
                    }

                    if (!IsJobSuccessful(job["JobState"]))
                    {
                        succeeded = false;

                        //
                        // In some cases the Job object can contain helpful information about
                        // why the method call failed. If it did contain such information,
                        // use it instead of a generic message.
                        //
                        if (!string.IsNullOrEmpty((string)job["ErrorDescription"]))
                        {
                            errorMessage = (string)job["ErrorDescription"];
                        }

                        if (printErrors)
                        {
                            PrintMsvmErrors(job);
                        }

                        if (throwIfFailed)
                        {
                            throw new ManagementException(errorMessage);
                        }
                    }
                }
            }
            else if ((uint)outputParameters["ReturnValue"] != 0)
            {
                succeeded = false;

                if (throwIfFailed)
                {
                    throw new ManagementException(errorMessage);
                }
            }

            return succeeded;
        }
    }
}
