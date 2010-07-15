using System;
using System.Collections.Generic;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;
using X4U.Outlook;

namespace SpamGrabberCommon
{
    public static class Reporting
    {
        
        /// <summary>
        /// Gets the headers of a MailItem
        /// </summary>
        /// <param name="objMailItem">The Outlook.MailItem object to use</param>
        /// <returns>String of the entire outlook headers</returns>
        public static string GetHeaders(Outlook.MailItem objMailItem)
        {
            string transportHeader;
            object missing = Missing.Value;
            // get the Outlook Application Object
            Outlook.Application outlookApplication = objMailItem.Application;//this._explorer.Application;
            

            // get the namespace object
            Outlook.NameSpace nameSpace = outlookApplication.GetNamespace("MAPI");

            // Logon to Session, here we use an already opened Outlook
            nameSpace.Logon(missing, missing, true, true);

            transportHeader = X4UMapi.GetMessageProperty(objMailItem.MAPIOBJECT, X4UMapi.PR_TRANSPORT_MESSAGE_HEADERS);

            // Release MailItem
            Marshal.ReleaseComObject(objMailItem);

            // logoff from namespace
            nameSpace.Logoff();

            // release resources
            Marshal.ReleaseComObject(nameSpace);
            Marshal.ReleaseComObject(outlookApplication.Application);

            // run garbagecollection
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // return the mail transport header
            return transportHeader;
        }
    }
}
