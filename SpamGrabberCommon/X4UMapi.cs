/// http://www.x4u.de
/// (c) Helmut Obertanner [X4U electronix]
/// flash@x4u.de
/// This sample demonstrates how to retrieve Outlook / MAPI Properties that could not be accessed 
/// by OOM (Outlook Object Model) or are subject to the Outlook Security Guard
/// 
/// Slightly modified by Per Baggesen
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;
using System.IO;


namespace X4U.Outlook
{
    public class X4UMapi
    {
        #region Public Functions
        /// <summary>
        /// The <b>GetMessageSenderAddress</b> function is used to retrieve a messagebody of a email without hitting the Outlook Security Guard. 
        /// </summary>
        /// <param name="mapiObject">The Outlook Item MAPIOBJECT property</param>
        /// <returns>The sender EmailAddress as string</returns>
        /// <example>
        /// object missing = Missing.Value; 
        ///
        /// get the Outlook Application Object
        /// Outlook.Application outlookApplication = new Outlook.Application();
        ///
        /// get the namespace object
        /// Outlook.NameSpace nameSpace = outlookApplication.GetNamespace("MAPI");
        ///
        /// Logon to Session, here we use an already opened Outlook
        /// nameSpace.Logon(missing, missing, false, false);
        /// 
        /// get the InboxFolder
        /// Outlook.MAPIFolder inboxFolder = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        ///
        /// get the first email
        /// Outlook.MailItem mailItem = ( Outlook.MailItem ) inboxFolder.Items[1];
        ///
        /// get mailbody
        /// string body = X4UMapi.GetMessageBody(mailItem.MAPIOBJECT);
        ///
        /// release used resources
        /// Marshal.ReleaseComObject(mailItem);
        /// Marshal.ReleaseComObject(inboxFolder);
        ///
        /// logof from namespace
        /// nameSpace.Logoff();
        ///
        /// release resources
        /// Marshal.ReleaseComObject( nameSpace );
        /// Marshal.ReleaseComObject(outlookApplication.Application); 
        /// </example>
        public static string GetMessageSenderAddress(object mapiObject)
        {
            // try to get the message body
            return GetMessageProperty(mapiObject, PR_SENDER_EMAIL_ADDRESS);
        }

        public static string GetMessageHeaders(object mapiObject)
        {
            // try to get the message body
            return GetMessageProperty(mapiObject, PR_TRANSPORT_MESSAGE_HEADERS);
        }

        #endregion

        #region Internal Functions

        /// <summary>
        /// Returns the Propertyvalue as string from the given Property Tag 
        /// </summary>
        /// <param name="mapiObject">[in] The Outlook Item MAPIOBJECT Property</param>
        /// <param name="propertyTag">[in] The Property Tag to retrieve</param>
        /// <returns>The Item Body as string.</returns>
        public static string GetMessageProperty(object mapiObject, uint propertyTag)
        {
            string body = "";

            // Pointer to IUnknown Interface
            IntPtr IUnknown = NULL;

            // Pointer to IMessage Interface
            IntPtr IMessage = NULL;

            // Pointer to IMAPIProp Interface
            IntPtr IMAPIProp = NULL;

            // Structure that will hold the Property Value
            SPropValue propValue;

            // A pointer that points to the SPropValue structure 
            IntPtr ptrPropValue = NULL;

            // if we have no MAPIObject everything is senseless...
            if (mapiObject == null) return "";

            try
            {
                // We can pass NULL here as parameter, so we do it. 
                MAPIInitialize(NULL);

                // retrive the IUnknon Interface from our MAPIObject comming from Outlook.
                IUnknown = Marshal.GetIUnknownForObject(mapiObject);

                // since HrGetOneProp needs a IMessage Interface, we must query our IUnknown interface for the IMessage interface.
                // create a Guid that we pass to retreive the IMessage Interface.
                Guid guidIMessage = new Guid(IID_IMessage);

                // try to retrieve the IMessage interface, if we don't get it, everything else is sensless.
                if (Marshal.QueryInterface(IUnknown, ref guidIMessage, out IMessage) != S_OK) return "";

                // create a Guid that we pass to retreive the IMAPIProp Interface.
                Guid guidIMAPIProp = new Guid(IID_IMAPIProp);

                // try to retrieve the IMAPIProp interface from IMessage Interface, everything else is sensless.
                if (Marshal.QueryInterface(IMessage, ref guidIMAPIProp, out IMAPIProp) != S_OK) return "";

                // double check, if we wave no pointer, exit...
                if (IMAPIProp == NULL) return "";

                // try to get the Property ( Property Tags can be found with Outlook Spy from Dmitry Streblechenko )
                // we pass the IMAPIProp Interface, the PropertyTag and the pointer to the SPropValue to the function.
                HrGetOneProp(IMAPIProp, propertyTag, out ptrPropValue);

                if (ptrPropValue == NULL) return string.Empty;

                // connect the pointer to our structure holding the value
                propValue = (SPropValue)Marshal.PtrToStructure(ptrPropValue, typeof(SPropValue));

                // now get the property
                // mark, that the result could also be a pointer to a stream if the messagebody is > 64K
                // the property value could also of another type
                body = Marshal.PtrToStringAnsi(new IntPtr(propValue.Value));
                return body;
            }
            catch (System.Exception ex)
            {
                return string.Empty;
            }
            finally
            {
                // Free used Memory structures
                if (ptrPropValue != NULL) MAPIFreeBuffer(ptrPropValue);

                // cleanup all references to COM Objects
                if (IMAPIProp != NULL) Marshal.Release(IMAPIProp);
                if (IMessage != NULL) Marshal.Release(IMessage);
                if (IUnknown != NULL) Marshal.Release(IUnknown);
                MAPIUninitialize();
            }
        }

        /// <summary>
        /// The <b>GetMessageBody</b> function is used to retrieve a messagebody of a email without hitting the Outlook Security Guard. 
        /// </summary>
        /// <param name="mapiObject"></param>
        /// <returns>The messagebody as string</returns>
        /// <example>
        /// object missing = Missing.Value; 
        ///
        /// get the Outlook Application Object
        /// Outlook.Application outlookApplication = new Outlook.Application();
        ///
        /// get the namespace object
        /// Outlook.NameSpace nameSpace = outlookApplication.GetNamespace("MAPI");
        ///
        /// Logon to Session, here we use an already opened Outlook
        /// nameSpace.Logon(missing, missing, false, false);
        /// 
        /// get the InboxFolder
        /// Outlook.MAPIFolder inboxFolder = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        ///
        /// get the first email
        /// Outlook.MailItem mailItem = ( Outlook.MailItem ) inboxFolder.Items[1];
        ///
        /// get mailbody
        /// string body = X4UMapi.GetMessageBody(mailItem.MAPIOBJECT);
        ///
        /// release used resources
        /// Marshal.ReleaseComObject(mailItem);
        /// Marshal.ReleaseComObject(inboxFolder);
        ///
        /// logof from namespace
        /// nameSpace.Logoff();
        ///
        /// release resources
        /// Marshal.ReleaseComObject( nameSpace );
        /// Marshal.ReleaseComObject(outlookApplication.Application); 
        /// </example>
        /// <param name="mapiObject">[in] The Outlook Item MAPIOBJECT Property</param>
        /// <returns>The Item Body as string.</returns>
        public static string GetMessageBody(object mapiObject)
        {
            string body = "";

            // Pointer to IUnknown Interface
            IntPtr IUnknown = NULL;

            // Pointer to IMessage Interface
            IntPtr IMessage = NULL;

            // Pointer to IMAPIProp Interface
            IntPtr IMAPIProp = NULL;

            // Structure that will hold the Property Value
            SPropValue propValue;

            // A pointer that points to the SPropValue structure 
            IntPtr ptrPropValue = NULL;

            // if we have no MAPIObject everything is senseless...
            if (mapiObject == null) return "";

            try
            {
                // We can pass NULL here as parameter, so we do it. 
                MAPIInitialize(NULL);

                // retrive the IUnknon Interface from our MAPIObject comming from Outlook.
                IUnknown = Marshal.GetIUnknownForObject(mapiObject);

                // since HrGetOneProp needs a IMessage Interface, we must query our IUnknown interface for the IMessage interface.
                // create a Guid that we pass to retreive the IMessage Interface.
                Guid guidIMessage = new Guid(IID_IMessage);

                // try to retrieve the IMessage interface, if we don't get it, everything else is sensless.
                if (Marshal.QueryInterface(IUnknown, ref guidIMessage, out IMessage) != S_OK) return "";

                // create a Guid that we pass to retreive the IMAPIProp Interface.
                Guid guidIMAPIProp = new Guid(IID_IMAPIProp);

                // try to retrieve the IMAPIProp interface from IMessage Interface, everything else is sensless.
                if (Marshal.QueryInterface(IMessage, ref guidIMAPIProp, out IMAPIProp) != S_OK) return "";

                // double check, if we wave no pointer, exit...
                if (IMAPIProp == NULL) return "";

                // try to get the Property ( Property Tags can be found with Outlook Spy from Dmitry Streblechenko )
                // we pass the IMAPIProp Interface, the PropertyTag and the pointer to the SPropValue to the function.
                HrGetOneProp(IMAPIProp, PR_BODY, out ptrPropValue);

                // if that also fails we have no such property
                if (ptrPropValue == NULL)
                {
                    // we will try another prop
                    HrGetOneProp(IMAPIProp, PR_BODY_HTML, out ptrPropValue);

                    if (ptrPropValue == NULL)
                    {
                        // if that fails we will try the last possibility
                        HrGetOneProp(IMAPIProp, PR_HTML, out ptrPropValue);
                    }
                }

                // if that also fails we have no such property
                if (ptrPropValue == NULL) return "";

                // connect the pointer to our structure holding the value
                propValue = (SPropValue)Marshal.PtrToStructure(ptrPropValue, typeof(SPropValue));

                // now get the property
                // mark, that the result could also be a pointer to a stream if the messagebody is > 64K
                // the property value could also of another type
                body = Marshal.PtrToStringAnsi(new IntPtr(propValue.Value));
                return body;
            }
            catch (System.Exception ex)
            {
                return "";
            }
            finally
            {
                // Free used Memory structures
                if (ptrPropValue != NULL) MAPIFreeBuffer(ptrPropValue);

                // cleanup all references to COM Objects
                if (IMAPIProp != NULL) Marshal.Release(IMAPIProp);
                if (IMessage != NULL) Marshal.Release(IMessage);
                if (IUnknown != NULL) Marshal.Release(IUnknown);
                MAPIUninitialize();
            }
        }

        /// <summary>
        /// Gets the datastream from an attachment
        /// </summary>
        /// <param name="mapiObject">the attachments mapi object.</param>
        /// <returns>string</returns>
        private static string GetMessageAttachments(object mapiObject)
        {
            string Attachmentstream = "";

            // Pointer to IUnknown Interface
            IntPtr IUnknown = NULL;

            // Pointer to IMessage Interface
            IntPtr IAttachment = NULL;

            // Pointer to IMAPIProp Interface
            IntPtr IMAPIProp = NULL;

            // Structure that will hold the Property Value
            SPropValue propValue;

            // A pointer that points to the SPropValue structure 
            IntPtr ptrPropValue = NULL;

            // if we have no MAPIObject everything is senseless...
            if (mapiObject == null) return "";

            try
            {
                // We can pass NULL here as parameter, so we do it. 
                MAPIInitialize(NULL);

                // retrive the IUnknon Interface from our MAPIObject comming from Outlook.
                IUnknown = Marshal.GetIUnknownForObject(mapiObject);

                // since HrGetOneProp needs a IID_IAttachment Interface, we must query our IUnknown interface for the IID_IAttachment interface.
                // create a Guid that we pass to retreive the IID_IAttachment Interface.
                Guid guidIMessage = new Guid(IID_IAttachment);

                // try to retrieve the IID_IAttachment interface, if we don't get it, everything else is sensless.
                if (Marshal.QueryInterface(IUnknown, ref guidIMessage, out IAttachment) != S_OK) return "";

                // create a Guid that we pass to retreive the IMAPIProp Interface.
                Guid guidIMAPIProp = new Guid(IID_IMAPIProp);

                // try to retrieve the IMAPIProp interface from IMessage Interface, everything else is sensless.
                if (Marshal.QueryInterface(IAttachment, ref guidIMAPIProp, out IMAPIProp) != S_OK) return "";

                // double check, if we wave no pointer, exit...
                if (IMAPIProp == NULL) return "";

                // get the size of the attachment.
                HrGetOneProp(IMAPIProp, PR_ATTACH_SIZE, out ptrPropValue);

                // try to get the Property ( Property Tags can be found with Outlook Spy from Dmitry Streblechenko )
                // we pass the IMAPIProp Interface, the PropertyTag and the pointer to the SPropValue to the function.
                if (ptrPropValue == NULL) return "";

                // connect the pointer to our structure holding the value
                propValue = (SPropValue)Marshal.PtrToStructure(ptrPropValue, typeof(SPropValue));
                int attSize = (int)propValue.Value;
                // if that also fails we have no such property

                HrGetOneProp(IMAPIProp, PR_ATTACH_DATA_BIN, out ptrPropValue);


                // if that also fails we have no such property
                if (ptrPropValue == NULL) return "";

                // connect the pointer to our structure holding the value
                propValue = (SPropValue)Marshal.PtrToStructure(ptrPropValue, typeof(SPropValue));

                // now get the property
                // mark, that the result could also be a pointer to a stream if the messagebody is > 64K
                // the property value could also of another type
                //Attachmentstream = Marshal.PtrToStringAnsi(new IntPtr(propValue.Value),attSize);
                byte[] filedata = new byte[attSize];
                for (int i = 0; i < attSize; i++)
                {
                    filedata[i] = Marshal.ReadByte(new IntPtr(propValue.Value), i);
                }
                Attachmentstream = filedata.ToString();


                return Attachmentstream;
            }
            catch (System.Exception ex)
            {
                return "";
            }
            finally
            {
                // Free used Memory structures
                if (ptrPropValue != NULL) MAPIFreeBuffer(ptrPropValue);

                // cleanup all references to COM Objects
                if (IMAPIProp != NULL) Marshal.Release(IMAPIProp);
                if (IAttachment != NULL) Marshal.Release(IAttachment);
                if (IUnknown != NULL) Marshal.Release(IUnknown);
                MAPIUninitialize();
            }
        }


        /// <summary>
        /// The <b>SetMessageBody</b> Method sets the Body property with the given text.
        /// </summary>
        /// <param name="mapiObject">The mapi message object</param>
        /// <param name="text">the Text that should be set on the property.</param>
        public static void SetMessageBody(object mapiObject, string text)
        {
            // Pointer to IUnknown Interface
            IntPtr IUnknown = NULL;

            // Pointer to IMessage Interface
            IntPtr IMessage = NULL;

            // Pointer to IMAPIProp Interface
            IntPtr IMAPIProp = NULL;

            // Structure that will hold the Property Value
            SPropValue propValue;

            // A pointer that points to the SPropValue structure 
            IntPtr ptrPropValue = NULL;

            // if we have no MAPIObject everything is senseless...
            if (mapiObject == null) return;

            try
            {
                // We can pass NULL here as parameter, so we do it. 
                MAPIInitialize(NULL);

                // retrive the IUnknon Interface from our MAPIObject comming from Outlook.
                IUnknown = Marshal.GetIUnknownForObject(mapiObject);

                // since HrGetOneProp needs a IMessage Interface, we must query our IUnknown interface for the IMessage interface.
                // create a Guid that we pass to retreive the IMessage Interface.
                Guid guidIMessage = new Guid(IID_IMessage);

                // try to retrieve the IMessage interface, if we don't get it, everything else is sensless.
                if (Marshal.QueryInterface(IUnknown, ref guidIMessage, out IMessage) != S_OK) return;

                // create a Guid that we pass to retreive the IMAPIProp Interface.
                Guid guidIMAPIProp = new Guid(IID_IMAPIProp);

                // try to retrieve the IMAPIProp interface from IMessage Interface, everything else is sensless.
                if (Marshal.QueryInterface(IMessage, ref guidIMAPIProp, out IMAPIProp) != S_OK) return;

                // double check, if we wave no pointer, exit...
                if (IMAPIProp == NULL) return;

                // Alloc memory for the text and create a pointer to it
                IntPtr ptrToValue = Marshal.StringToHGlobalAnsi(text);

                // Create our structure with data
                propValue = new SPropValue();
                // Wich property should be set
                propValue.ulPropTag = PR_BODY;
                propValue.dwAlignPad = 0;
                propValue.Value = (long)ptrToValue;

                ptrPropValue = Marshal.AllocHGlobal(Marshal.SizeOf(propValue));
                Marshal.StructureToPtr(propValue, ptrPropValue, false);

                // try to set the Property ( Property Tags can be found with Outlook Spy from Dmitry Streblechenko )
                HrSetOneProp(IMAPIProp, ptrPropValue);

                Marshal.FreeHGlobal(ptrPropValue);
                Marshal.FreeHGlobal(ptrToValue);
            }
            catch (System.Exception ex)
            {
            }
            finally
            {
                // Free used Memory structures
                if (ptrPropValue != NULL) MAPIFreeBuffer(ptrPropValue);

                // cleanup all references to COM Objects
                if (IMAPIProp != NULL) Marshal.Release(IMAPIProp);
                if (IMessage != NULL) Marshal.Release(IMessage);
                if (IUnknown != NULL) Marshal.Release(IUnknown);
                MAPIUninitialize();
            }
        }

        /// <summary>
        /// Returns the Propertyvalue as string from the given Property Tag 
        /// </summary>
        /// <param name="mapiObject">[in] The Outlook Item MAPIOBJECT Property</param>
        /// <returns>The Item Body as string.</returns>
        public static string GetMessageEntryID(object mapiObject)
        {
            SEntryID entryID = new SEntryID();

            // Pointer to IUnknown Interface
            IntPtr IUnknown = NULL;

            // Pointer to IMessage Interface
            IntPtr IMessage = NULL;

            // Pointer to IMAPIProp Interface
            IntPtr IMAPIProp = NULL;

            // Structure that will hold the Property Value
            SPropValue propValue;

            // A pointer that points to the SPropValue structure 
            IntPtr ptrPropValue = NULL;

            // if we have no MAPIObject everything is senseless...
            if (mapiObject == null) return "";

            try
            {
                // We can pass NULL here as parameter, so we do it. 
                MAPIInitialize(NULL);

                // retrive the IUnknon Interface from our MAPIObject comming from Outlook.
                IUnknown = Marshal.GetIUnknownForObject(mapiObject);

                // since HrGetOneProp needs a IMessage Interface, we must query our IUnknown interface for the IMessage interface.
                // create a Guid that we pass to retreive the IMessage Interface.
                Guid guidIMessage = new Guid(IID_IMessage);

                // try to retrieve the IMessage interface, if we don't get it, everything else is sensless.
                if (Marshal.QueryInterface(IUnknown, ref guidIMessage, out IMessage) != S_OK) return "";

                // create a Guid that we pass to retreive the IMAPIProp Interface.
                Guid guidIMAPIProp = new Guid(IID_IMAPIProp);

                // try to retrieve the IMAPIProp interface from IMessage Interface, everything else is sensless.
                if (Marshal.QueryInterface(IMessage, ref guidIMAPIProp, out IMAPIProp) != S_OK) return "";

                // double check, if we wave no pointer, exit...
                if (IMAPIProp == NULL) return "";

                // try to get the Property ( Property Tags can be found with Outlook Spy from Dmitry Streblechenko )
                // we pass the IMAPIProp Interface, the PropertyTag and the pointer to the SPropValue to the function.
                HrGetOneProp(IMAPIProp, PR_ENTRYID, out ptrPropValue);

                // if that also fails we have no such property
                if (ptrPropValue == NULL) return "";

                // connect the pointer to our structure holding the value
                propValue = (SPropValue)Marshal.PtrToStructure(ptrPropValue, typeof(SPropValue));

                // now get the property
                // mark, that the result could also be a pointer to a stream if the messagebody is > 64K
                // the property value could also of another type
                // entryID = (SEntryID) propValue.Value;

                return "";
            }
            catch (System.Exception ex)
            {
                return "";
            }
            finally
            {
                // Free used Memory structures
                if (ptrPropValue != NULL) MAPIFreeBuffer(ptrPropValue);

                // cleanup all references to COM Objects
                if (IMAPIProp != NULL) Marshal.Release(IMAPIProp);
                if (IMessage != NULL) Marshal.Release(IMessage);
                if (IUnknown != NULL) Marshal.Release(IUnknown);
                MAPIUninitialize();
            }
        }
        #endregion

        #region Private Properties

        /// <summary>
        /// A Variable used as C-Style NULL Pointer;
        /// </summary>
        private static readonly IntPtr NULL = IntPtr.Zero;

        /// <summary>
        /// Used for checking returncodes.
        /// </summary>
        private const int S_OK = 0;

        #endregion

        #region Initialization / Cleanup

        /// <summary>
        /// The construction Code.
        /// </summary>
        public X4UMapi()
        {

        }

        #endregion

        #region MAPI Interface ID'S


        // The Interface ID's are used to retrieve the specific MAPI Interfaces from the IUnknown Object

        public const string IID_IMAPISession = "00020300-0000-0000-C000-000000000046";
        public const string IID_IMAPIProp = "00020303-0000-0000-C000-000000000046";
        public const string IID_IMAPITable = "00020301-0000-0000-C000-000000000046";
        public const string IID_IMAPIMsgStore = "00020306-0000-0000-C000-000000000046";
        public const string IID_IMAPIFolder = "0002030C-0000-0000-C000-000000000046";
        public const string IID_IMAPISpoolerService = "0002031E-0000-0000-C000-000000000046";
        public const string IID_IMAPIStatus = "0002031E-0000-0000-C000-000000000046";
        public const string IID_IMessage = "00020307-0000-0000-C000-000000000046";
        public const string IID_IAddrBook = "00020309-0000-0000-C000-000000000046";
        public const string IID_IProfSect = "00020304-0000-0000-C000-000000000046";
        public const string IID_IMAPIContainer = "0002030B-0000-0000-C000-000000000046";
        public const string IID_IABContainer = "0002030D-0000-0000-C000-000000000046";
        public const string IID_IMsgServiceAdmin = "0002031D-0000-0000-C000-000000000046";
        public const string IID_IProfAdmin = "0002031C-0000-0000-C000-000000000046";
        public const string IID_IMailUser = "0002030A-0000-0000-C000-000000000046";
        public const string IID_IDistList = "0002030E-0000-0000-C000-000000000046";
        public const string IID_IAttachment = "00020308-0000-0000-C000-000000000046";
        public const string IID_IMAPIControl = "0002031B-0000-0000-C000-000000000046";
        public const string IID_IMAPILogonRemote = "00020346-0000-0000-C000-000000000046";
        public const string IID_IMAPIForm = "00020327-0000-0000-C000-000000000046";
        // added by per
        public const string IID_IStream = "0000000C-0000-0000-C000-000000000046";


        #endregion

        #region Property Tags
        /// <summary>
        /// Used to get the Emailheaders
        /// </summary>
        public const uint PR_TRANSPORT_MESSAGE_HEADERS = 0x007D001E;

        /// <summary>
        /// Used to read the Body of an Email
        /// </summary>
        public const uint PR_BODY = 0x1000001E;

        /// <summary>
        /// Used to read the HTML Body of the Email
        /// </summary>
        public const uint PR_BODY_HTML = 0x1013001E;

        /// <summary>
        /// Used to read the HTML Body of the Email
        /// </summary>
        public const uint PR_HTML = 0x10130102;

        /// <summary>
        /// Used to read the smtp / exchange sender address of an Email
        /// </summary>
        public const uint PR_SENDER_EMAIL_ADDRESS = 0x0C1F001E;

        /// <summary>
        /// 
        /// </summary>
        public const uint PR_ENTRYID = 0x0FFF0102;

        // attachment handling constants
        // added PB/NextUs
        public const uint PR_ATTACH_ADDITIONAL_INFO = 0x370F0102;
        public const uint PR_ATTACH_DATA_BIN = 0x37010102;
        public const uint PR_ATTACH_DATA_OBJ = 0x3701000D;
        public const uint PR_ATTACH_ENCODING = 0x37020102;
        public const uint PR_ATTACH_EXTENSION = 0x3703001E;
        public const uint PR_ATTACH_FILENAME = 0x3704001E;
        public const uint PR_ATTACH_LONG_FILENAME = 0x3707001E;
        public const uint PR_ATTACH_LONG_PATHNAME = 0x370D001E;
        public const uint PR_ATTACH_METHOD = 0x37050003;
        public const uint PR_ATTACH_MIME_TAG = 0x370E001E;
        public const uint PR_ATTACH_NUM = 0x0E210003;
        public const uint PR_ATTACH_PATHNAME = 0x3708001E;
        public const uint PR_ATTACH_RENDERING = 0x37090102;
        public const uint PR_ATTACH_SIZE = 0x0E200003;
        public const uint PR_ATTACH_TAG = 0x370A0102;
        public const uint PR_ATTACH_TRANSPORT_NAME = 0x370C001E;
        // end attachment


        public const uint PT_NULL = 1;	/* NULL property value */
        public const uint PT_I2 = 2;	/* Signed 16-bit value */
        public const uint PT_LONG = 3;	/* Signed 32-bit value */
        public const uint PT_R4 = 4;	/* 4-byte floating point */
        public const uint PT_DOUBLE = 5;	/* Floating point double */
        public const uint PT_CURRENCY = 6;	/* Signed 64-bit int (decimal w/	4 digits right of decimal pt) */
        public const uint PT_APPTIME = 7;	/* Application time */
        public const uint PT_ERROR = 10;	/* 32-bit error value */
        public const uint PT_BOOLEAN = 11;	/* 16-bit boolean (non-zero true) */
        public const uint PT_OBJECT = 13;	/* Embedded object in a property */
        public const uint PT_I8 = 20;	/* 8-byte signed integer */
        public const uint PT_STRING8 = 30;	/* Null terminated 8-bit character string */
        public const uint PT_UNICODE = 31;	/* Null terminated Unicode string */
        public const uint PT_SYSTIME = 64;	/* FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601 */
        public const uint PT_CLSID = 72;	/* OLE GUID */
        public const uint PT_BINARY = 258;	/* Uninterpreted (counted byte array) */

        #endregion

        #region MAPI Structures


        /// <summary>
        /// The SPropValue structure describes a MAPI property.
        /// </summary>
        public struct SPropValue
        {
            /// <summary>
            /// Property tag for the property. Property tags are 32-bit unsigned integers consisting of the property's unique identifier in the high-order 16 bits and the property's type in the low-order 16 bits.
            /// </summary>
            public uint ulPropTag;

            /// <summary>
            /// Reserved for MAPI; do not use.
            /// </summary>
            public uint dwAlignPad;

            /// <summary>
            /// Union of data values, the specific value dictated by the property type.
            /// </summary>
            public long Value;
        }

        /* ENTRYID */
        public struct SEntryID
        {
            byte[] abFlags;
            byte[] ab;
        }

        #endregion

        #region MAPI DLL Imports

        /// <summary>
        /// The MAPIInitialize function increments the MAPI subsystem reference count and initializes global data for the MAPI DLL.
        /// </summary>
        /// <param name="lpMapiInit">[in] Pointer to a MAPIINIT_0 structure. The lpMapiInit parameter can be set to NULL.</param>
        /// <returns>
        /// S_OK
        /// The MAPI subsystem was initialized successfully.
        /// </returns>
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
        private static extern int MAPIInitialize(IntPtr lpMapiInit);

        /// <summary>
        /// The MAPIUninitialize function decrements the reference count, cleans up, and deletes per-instance global data for the MAPI DLL.
        /// </summary>
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
        private static extern void MAPIUninitialize();

        /// <summary>
        /// The HrGetOneProp function retrieves the value of a single property from a property interface, that is, an interface derived from IMAPIProp.
        /// </summary>
        /// <param name="pmp">[in] Pointer to the IMAPIProp interface from which the property value is to be retrieved.</param>
        /// <param name="ulPropTag">[in] Property tag of the property to be retrieved.</param>
        /// <param name="ppprop">[out] Pointer to a pointer to the returned SPropValue structure defining the retrieved property value.</param>
        /// <remarks>
        /// Unlike the IMAPIProp::GetProps method, the HrGetOneProp function never returns any warning.
        /// Because it retrieves only one property, it simply either succeeds or fails. For retrieving multiple properties,
        /// GetProps is faster. 
        ///
        /// You can set or change a single property with the HrSetOneProp function.
        /// </remarks>
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "HrGetOneProp@12")]
        private static extern int HrGetOneProp(IntPtr pmp, uint ulPropTag, out IntPtr ppprop);

        /// <summary>
        /// The HrSetOneProp function sets or changes the value of a single property on a property interface, that is, an interface derived from IMAPIProp.
        /// </summary>
        /// <param name="pmp">[in] Pointer to an IMAPIProp interface on which the property value is to be set or changed.</param>
        /// <param name="pprop">[in] Pointer to the SPropValue structure defining the property to be set or changed.</param>
        /// <remarks>
        /// Unlike the IMAPIProp::SetProps method, the HrSetOneProp function never returns any warning.
        /// Because it sets only one property, it simply either succeeds or fails.
        /// For setting or changing multiple properties, SetProps is faster. 
        /// 
        /// You can retrieve a single property with the HrGetOneProp function.
        /// </remarks>
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "HrSetOneProp@8")]
        private static extern int HrSetOneProp(IntPtr pmp, IntPtr pprop);

        /// <summary>
        /// The MAPIFreeBuffer function frees a memory buffer allocated with a call to the MAPIAllocateBuffer function or the MAPIAllocateMore function.
        /// </summary>
        /// <param name="lpBuffer">[in] Pointer to a previously allocated memory buffer. If NULL is passed in the lpBuffer parameter, MAPIFreeBuffer does nothing.</param>
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "MAPIFreeBuffer@4")]
        private static extern void MAPIFreeBuffer(IntPtr lpBuffer);

        //[DllImport("OLE32.DLL", EntryPoint = "CreateStreamOnHGlobal")]
        //extern public static int CreateStreamOnHGlobal(int hGlobalMemHandle, bool
        //    fDeleteOnRelease, out UCOMIStream pOutStm);



        #endregion
    }
}