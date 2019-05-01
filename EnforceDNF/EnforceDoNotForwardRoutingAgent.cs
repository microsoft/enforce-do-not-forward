using System;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.ContentTypes.Tnef;
using Microsoft.Exchange.Data.Transport.Routing;
using System.Diagnostics;

namespace EnforceDNF
{

    public class EnforceDoNotForwardRoutingAgent : RoutingAgent, IDisposable
    {
        private static readonly Guid PsPublicStrings = new Guid("00020329-0000-0000-C000-000000000046");
        private const string PropertyTagName = "DoNotForward";
        private const string EventLogSource = "EnforceDoNotForwardRoutingAgent";
        private const string EventLogName = "Application";

        internal EventLog EventLog;

        public EnforceDoNotForwardRoutingAgent()
        {
            EventLog = new EventLog {Log = EventLogName, Source = EventLogSource};

            OnRoutedMessage += EnforceDoNotForwardRoutingAgent_OnRoutedMessage;
        }

        private void EnforceDoNotForwardRoutingAgent_OnRoutedMessage(RoutedMessageEventSource esEvtSource, QueuedMessageEventArgs qmQueuedMessage)
        {

            // determine whether or not we care about this message
            if (qmQueuedMessage.MailItem.Message.MapiMessageClass == "IPM.Schedule.Meeting.Request")
            {
                ProcessMailItem(qmQueuedMessage.MailItem);
            }
        
        }

        internal void ProcessMailItem(MailItem item)
        {
            bool isPropertySet = false;
            int tag;
            unchecked
            {
                tag = (int)0x8000000B;
            }

            try
            {
                MimePart tnefPart = item.Message.TnefPart;
                if (tnefPart != null)
                {
                    TnefReader reader = new TnefReader(tnefPart.GetContentReadStream());
                    TnefWriter writer = new TnefWriter(
                        tnefPart.GetContentWriteStream(tnefPart.ContentTransferEncoding),
                        reader.AttachmentKey);

                    while (reader.ReadNextAttribute())
                    {
                        if (reader.AttributeTag == TnefAttributeTag.MapiProperties)
                        {
                            writer.StartAttribute(TnefAttributeTag.MapiProperties, TnefAttributeLevel.Message);
                            while (reader.PropertyReader.ReadNextProperty())
                            {
                                if (reader.PropertyReader.IsNamedProperty)
                                {
                                    switch (reader.PropertyReader.PropertyNameId.Name)
                                    {
                                        case PropertyTagName:
                                            isPropertySet = true;
                                            writer.StartProperty(new TnefPropertyTag(tag), PsPublicStrings, PropertyTagName);
                                            writer.WritePropertyValue(true);
                                            break;
                                        default:
                                            writer.WriteProperty(reader.PropertyReader);
                                            break;
                                    }
                                }
                                else
                                {
                                    writer.WriteProperty(reader.PropertyReader);
                                }
                            }

                            if (!isPropertySet)
                            {
                                writer.StartProperty(new TnefPropertyTag(tag), PsPublicStrings, PropertyTagName);
                                writer.WritePropertyValue(true);
                            }
                        }
                        else
                        {
                            writer.WriteAttribute(reader);
                        }
                    }
                    if (null != writer)
                    {
                        writer.Close();
                    }
                }
                else
                {
                    WriteLog("Attempted to process item with null TnefPart. Subject: " + item.Message.Subject, EventLogEntryType.Warning, 2000, EventLogSource);
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message + ". Stack Trace: " + ex.StackTrace, EventLogEntryType.Error, 5000, EventLogSource);
            }


        }
        internal void WriteLog(string message, EventLogEntryType entryType,
                                int eventId, string processName)
        {

            try
            {
                EventLog.WriteEntry(message, entryType, eventId);
            }
            catch
            {
            }
        }

        #region IDisposable Support
        private bool _disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    try
                    {
                        EventLog?.Dispose();
                    }
                    catch
                    {
                    }
                }

                _disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
        }
        #endregion
    }


}
