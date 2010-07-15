using System;
using System.Collections.Generic;
using System.Text;

namespace SpamGrabberCommon
{
    public class Base64
    {
        /// <summary>
        /// Base64 Encoding
        /// </summary>
        /// <param name="str">the string to encode</param>
        /// <returns>base 64 endcoded string</returns>
        public static string Encode(string pstrSource)
        {
            byte[] encbuff = System.Text.Encoding.UTF8.GetBytes(pstrSource);
            return Convert.ToBase64String(encbuff);
        }
        /// <summary>
        /// Base64 Encode
        /// </summary>
        /// <param name="encbuff">byte array to convert</param>
        /// <returns>base 64 encoded string</returns>
        public static string Encode(byte[] pobjEncbuff)
        {
            return Convert.ToBase64String(pobjEncbuff);
        }
        /// <summary>
        /// Base64 Decoding
        /// </summary>
        /// <param name="str">string to decode</param>
        /// <returns>the decoded string</returns>
        public static string Decode(string pstrSource)
        {
            byte[] decbuff = Convert.FromBase64String(pstrSource);
            return System.Text.Encoding.UTF8.GetString(decbuff);
        }
    }
}
