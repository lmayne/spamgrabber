using System;
using System.Collections.Generic;
using System.Text;
using System.Net;


namespace SpamGrabberCommon
{
    /// <summary>
    /// Utility class used to get the IP/net information from the client.
    /// </summary>
    public static class SpamGrapperUtility
    {
        public static string GetIPAdressesFromLocalhost()
        {
            List<string> lstIpAdress = new List<string>();
            String strHostName = string.Empty;


            // Getting Ip address of local machine...
            // First get the host name of local machine.
            strHostName = Dns.GetHostName();


            // Then using host name, get the IP address list..
            IPHostEntry ipEntry = Dns.GetHostEntry(strHostName);
            IPAddress[] addr = ipEntry.AddressList;

            for (int i = 0; i < addr.Length; i++)
            {
                if (addr[i].AddressFamily== System.Net.Sockets.AddressFamily.InterNetwork)
                    lstIpAdress.Add(addr[i].ToString());
            }
            //return string.Join(".", lstIpAdress.ToArray());
            return string.Join("_", lstIpAdress.ToArray()).Replace(".","_");// changed to ensure that we dont create multi extention
        }

        /// <summary>
        /// Greates the attachment name for the reports
        /// </summary>
        /// <param name="action_spam">is the report Spam?</param>
        /// <param name="sendIp">should the name include the client IP</param>
        /// <returns>the attachment name</returns>
        public static string CreateAttachmentName(SGGlobals.ReportAction pAction, bool pblnSendIp)
        {
            string s = string.Empty;
            string att_name = string.Empty;
            string programVersion = GlobalSettings.PROG_VERSION;

            programVersion = programVersion.Replace(".", "_");

            //' get date and time
            s = DateTime.Now.ToString();
            s = s.Replace(":", "");
            s = s.Replace("-", "");
            s = s.Replace(" ", "");
            s = s.Replace("/", "");
            s = s.Replace("\\", "");

            if (pAction==SGGlobals.ReportAction.ReportSpam)// SPAM
            {
                if (pblnSendIp)
                {
                    att_name = "blacklist_" + s + "_" + GetIPAdressesFromLocalhost() + "_" + programVersion + ".txt";
                }
                else
                {
                    att_name = ("blacklist_" + s + "_unknown_ip" + "_" + programVersion + ".txt");
                }
            }
            else if (pAction==SGGlobals.ReportAction.ReportHam)
            { // HAM
                if (pblnSendIp)
                {
                    att_name = ("whitelist_" + s + "_" + GetIPAdressesFromLocalhost() + "_" + programVersion + ".txt");
                }
                else
                {
                    att_name = ("whitelist_" + s + "_unknown_ip" + "_" + programVersion + ".txt");
                }
            }
            else if (pAction == SGGlobals.ReportAction.ReportSupport)
            {
                if (pblnSendIp)
                {
                    att_name = ("Support_" + s + "_" + GetIPAdressesFromLocalhost() + "_" + programVersion + ".txt");
                }
                else
                {
                    att_name = ("Support_" + s + "_unknown_ip" + "_" + programVersion + ".txt");
                }
            }

            return att_name;
        }

      
    }

}
