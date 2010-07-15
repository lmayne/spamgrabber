using System;
using System.Collections;
using System.Text;

/// <summary>
/// Class for encoding and decoding a string to QuotedPrintable
/// RFC 1521 http://www.ietf.org/rfc/rfc1521.txt
/// RFC 2045 http://www.ietf.org/rfc/rfc2045.txt
/// Date: 2006-03-23
/// Author: Kevin Spaun
/// Company: SPAUN Informationstechnik GmbH - http://www.spaun-it.com/
/// Feedback: kspaun@spaun-it.de
/// License: This piece of code comes with no guaranties. You can use it for whatever you want for free.
/// </summary>
/// 
namespace SpamGrabberCommon
{
    public class QuotedPrintable
    {
        private const byte EQUALS = 61;
        private const byte CR = 13;
        private const byte LF = 10;
        private const byte SPACE = 32;
        private const byte TAB = 9;

        /// <summary>
        /// Encodes a string to QuotedPrintable
        /// </summary>
        /// <param name="_ToEncode">String to encode</param>
        /// <returns>QuotedPrintable encoded string</returns>
        public static string Encode(string pstrToEncode)
        {
            StringBuilder Encoded = new StringBuilder();
            string hex = string.Empty;
            byte[] bytes = Encoding.Default.GetBytes(pstrToEncode);
            int count = 0;

            for (int i = 0; i < bytes.Length; i++)
            {
                //these characters must be encoded
                if ((bytes[i] < 33 || bytes[i] > 126 || bytes[i] == EQUALS) && bytes[i] != CR && bytes[i] != LF && bytes[i] != SPACE)
                {
                    if (bytes[i].ToString("X").Length < 2)
                    {
                        hex = "0" + bytes[i].ToString("X");
                        Encoded.Append("=" + hex);
                    }
                    else
                    {
                        hex = bytes[i].ToString("X");
                        Encoded.Append("=" + hex);
                    }
                }
                else
                {
                    //check if index out of range
                    if ((i + 1) < bytes.Length)
                    {
                        //if TAB is at the end of the line - encode it!
                        if ((bytes[i] == TAB && bytes[i + 1] == LF) || (bytes[i] == TAB && bytes[i + 1] == CR))
                        {
                            Encoded.Append("=0" + bytes[i].ToString("X"));
                        }
                        //if SPACE is at the end of the line - encode it!
                        else if ((bytes[i] == SPACE && bytes[i + 1] == LF) || (bytes[i] == SPACE && bytes[i + 1] == CR))
                        {
                            Encoded.Append("=" + bytes[i].ToString("X"));
                        }
                        else
                        {
                            Encoded.Append(System.Convert.ToChar(bytes[i]));
                        }
                    }
                    else
                    {
                        Encoded.Append(System.Convert.ToChar(bytes[i]));
                    }
                }
                if (count == 75)
                {
                    Encoded.Append("=\r\n"); //insert soft-linebreak
                    count = 0;
                }
                count++;
            }

            return Encoded.ToString();
        }

        /// <summary>
        /// Decodes a QuotedPrintable encoded string 
        /// </summary>
        /// <param name="_ToDecode">The encoded string to decode</param>
        /// <returns>Decoded string</returns>
        public static string Decode(string pstrToDecode)
        {
            StringBuilder Decoded = new StringBuilder();

            //remove soft-linebreaks first
            pstrToDecode = pstrToDecode.Replace("=\r\n", "");

            char[] chars = pstrToDecode.ToCharArray();

            for (int i = 0; i < chars.Length; i++)
            {
                // if encoded character found decode it
                if (chars[i] == '=')
                {
                    Decoded.Append(System.Convert.ToChar(int.Parse(chars[i + 1].ToString() + chars[i + 2].ToString(), System.Globalization.NumberStyles.HexNumber)));
                    //jump ahead
                    i += 2;
                }
                else
                {
                    Decoded.Append(System.Convert.ToChar(chars[i]));
                }
            }

            return Decoded.ToString();
        }
    }

}