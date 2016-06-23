using SKYPE4COMLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Reflection;

namespace SkypeLib
{
    /// <summary> 
    /// SkypeHelper class is used to interact with Skype
    /// </summary> 
    public class SkypeHelper
    {
        private Skype SkypeClient;

        /// <summary> 
        /// SkypeHelper class is used to interact with Skype
        /// Allow access in skype to operate.
        /// </summary>
        public SkypeHelper()
        {
            SkypeClient = new Skype();
            if (!SkypeClient.Client.IsRunning)
            {
                SkypeClient.Client.Start(true, false);
            }          
            SkypeClient.Attach(7, false);          
        }

        /// <summary> 
        /// Get a list of all Contact's UserName
        /// </summary
        public List<string> GetAllContact()
        {
            List<string> contacts = new List<string>();
            IEnumerable<SKYPE4COMLib.User> users = SkypeClient.Friends.OfType<SKYPE4COMLib.User>();

            foreach (var item in users)
            {
                contacts.Add(item.Handle);
            }
            contacts.Sort();

            return contacts;
        }

        /// <summary> 
        /// Get a list of Online Contact UserName
        /// </summary>
        public List<string> GetOnlineContact()
        {
            List<string> contacts = new List<string>();
            IEnumerable<SKYPE4COMLib.User> users = SkypeClient.Friends.OfType<SKYPE4COMLib.User>();

            users
                .Where(u => u.OnlineStatus == TOnlineStatus.olsOnline)
                .OrderBy(u => u.FullName)
                .ToList()
                .ForEach(u => contacts.Add(u.Handle));

            return contacts;
        }

        /// <summary> 
        /// Send message to a specific person
        /// </summary>
        public void SendMessageToTheSpecificContact(string id, string message)
        {          
            SkypeClient.SendMessage(id, message);
        }

        /// <summary> 
        /// Send message to all contacts
        /// </summary>
        public void SendGenericMessage(string message)
        {
            List<string> contacts = new List<string>();
            contacts = GetAllContact(); 
            foreach(var id in contacts)
            {
                SendMessageToTheSpecificContact(id, message);
            }
        }

        /// <summary> 
        /// Send message according to specific list
        /// </summary>
        public void SendGenericMessage(List<string> str, string message)
        {
            foreach (var id in str)
            {
                SendMessageToTheSpecificContact(id, message);
            }
        }

        /// <summary> 
        /// Send group message
        /// </summary>
        private void SendMessageToGroup(string GroupName)
        {
              
        }

        /// <summary> 
        /// Send SMS to mobile number
        /// </summary>
        public void SendSMS(string mobileNumber, string msg)
        {
            try
            {
                SkypeClient.Timeout = 120 * 1000;

                var messageType = TSmsMessageType.smsMessageTypeOutgoing;
                var message = SkypeClient.CreateSms(messageType, mobileNumber);

                message.Body = msg;
                message.Send();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary> 
        /// Send IM and SMS based on Message.xml file
        /// </summary>
        public void SendNotification()
        {
            string fileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "NotificationConfig.xml");

            string logDir = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "log");
            string logFile = Path.Combine(logDir, "Messages.log");

            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);       
            }

            if (!File.Exists(logFile))
            {
                FileStream file = File.Create(logFile);
                file.Close();
            }

            CheckFileSize(logFile);

            XmlDocument doc = new XmlDocument();
            if (File.Exists(fileName))
            {
                doc.Load(fileName);
            }
            else
            {
                throw new Exception("Can't find Messages.xml");
            }

            XmlElement root = doc.DocumentElement;

            //send IM from xml file
            XmlNode imNode = root.SelectSingleNode("IM");
            if (imNode.Attributes["enable"].Value == "true")
            {
                XmlNodeList imNodeList = imNode.SelectNodes("person");
                foreach (XmlNode node in imNodeList)
                {
                    SkypeClient.SendMessage(node.Attributes["id"].Value, node.Attributes["message"].Value);
                    using (StreamWriter writer = new StreamWriter(logFile, true))
                    {
                        string info = DateTime.Now.ToString() + " Send IM to " + node.Attributes["id"].Value + " -> " + node.Attributes["message"].Value;
                        writer.WriteLine(info);
                    }
                }              
            }

            //Send sms from xml file
            XmlNode smsNode = root.SelectSingleNode("SMS");
            if (smsNode.Attributes["enable"].Value == "true")
            {
                XmlNodeList smsNodeList = smsNode.SelectNodes("person");
                foreach (XmlNode node in smsNodeList)
                {
                    SendSMS(node.Attributes["Num"].Value, node.Attributes["message"].Value);
                    using (StreamWriter writer = new StreamWriter(logFile, true))
                    {
                        string info = DateTime.Now.ToString() + " Send SMS to " + node.Attributes["Num"].Value + " -> " + node.Attributes["message"].Value;
                        writer.WriteLine(info);
                    }
                }            
            }           
        }

        private void CheckFileSize(string logFile)
        {
            string oldLogFile = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "log", "Messages_Old.log");
            FileInfo fileInfo = new FileInfo(logFile);
            long size = (fileInfo.Length) / (1024 * 1024);
            if (size >= 10)
            {
                if (File.Exists(oldLogFile))
                    File.Delete(oldLogFile);

                fileInfo.MoveTo(oldLogFile);
                FileStream file = File.Create(logFile);
                file.Close();
            }
        }

        /// <summary> 
        /// Make a call to specific contact
        /// </summary>
        public void MakeACall(string id)
        {
            //Process.Start("callto:" + id);

            if (SkypeClient.User[id].OnlineStatus == TOnlineStatus.olsOnline)
            {

                Call call = SkypeClient.PlaceCall(id);
                do
                {
                    System.Threading.Thread.Sleep(1);
                } while (call.Status != TCallStatus.clsInProgress);

                

                call.StartVideoSend();



                int duration = call.Duration;
                int vmDuraiton = call.VmDuration;
            }
            else
            {
                throw new Exception("User is not online.");
            }
        }

        /// <summary> 
        /// Make group call
        /// </summary>
        public void MakeGroupCall(string GroupName)
        {
            foreach (Group group in SkypeClient.Groups)
            {
                
            }
        }
        
        private void AutoResponse()
        {
            while (true)
            {
                foreach (IChatMessage msg in SkypeClient.MissedMessages)
                {
                    string handle = msg.Sender.Handle;
                    string message = "This is an auto response message.";
                    SkypeClient.SendMessage(handle, message);
                }
                System.Threading.Thread.Sleep(2000);
            }
        }
    }
}
