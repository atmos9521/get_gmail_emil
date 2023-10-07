using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Security;
using System.IO;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;

using Google.Apis.Gmail.v1.Data;

namespace Get_Gmail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //https://console.cloud.google.com/apis
        //已啟用的API和服務
        //啟用的API和服務
        //GmailAPI
        //憑證
        //建立憑證
        //OAuth用戶端ID
        //下載json改名成credentials.json放到bin\Debug內

        private void button1_Click(object sender, EventArgs e)
        {
            string host = "imap.gmail.com";
            int port = 993;
            string username = "willychen0206@gmail.com";
            string password = "!QAZ2wsx9521";

            using (TcpClient client = new TcpClient())
            {
                client.Connect(host, port);
                using (SslStream sslStream = new SslStream(client.GetStream()))
                {
                    sslStream.AuthenticateAsClient(host);
                    StreamReader reader = new StreamReader(sslStream);
                    StreamWriter writer = new StreamWriter(sslStream);

                    // Read the greeting message from the server
                    Console.WriteLine(reader.ReadLine());

                    // Login to the server
                    writer.WriteLine($"LOGIN {username} {password}");
                    writer.Flush();

                    // Read the response from the server
                    string response = reader.ReadLine();
                    Console.WriteLine(response);

                    // List the available mailboxes
                    writer.WriteLine("LIST \"\" \"*\"");
                    writer.Flush();

                    // Read the list of mailboxes
                    response = reader.ReadLine();
                    Console.WriteLine(response);

                    // Select the mailbox (inbox in this case)
                    writer.WriteLine("SELECT INBOX");
                    writer.Flush();

                    // Read the mailbox information
                    response = reader.ReadLine();
                    Console.WriteLine(response);

                    // Search for all unseen messages
                    writer.WriteLine("SEARCH UNSEEN");
                    writer.Flush();

                    // Read the list of unseen messages
                    response = reader.ReadLine();
                    Console.WriteLine(response);

                    // Logout and close the connection
                    writer.WriteLine("LOGOUT");
                    writer.Flush();
                }
            }

            Console.ReadLine();
        }




        string[] Scopes = { GmailService.Scope.GmailReadonly };
        string ApplicationName = "Gmail API C# Quickstart";

        private void button2_Click(object sender, EventArgs e)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // 创建 Gmail 服务
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // 使用服务进行操作，例如获取邮件列表
            // 参考 Gmail API 文档：https://developers.google.com/gmail/api

            //// 读取Gmail邮件
            //var service = GetGmailService(credentialsFilePath);

            //// 读取邮件
            ///is:unread：获取所有未读邮件。
            //is:read：获取所有已读邮件。
            //in:inbox：获取收件箱中的所有邮件。
            //from: example @example.com：获取特定发件人的邮件（将 example@example.com 替换为实际的发件人地址）。
            //subject: "Your Subject"：获取主题为 "Your Subject" 的邮件。
            //before: 2023 / 01 / 01：获取在 2023 年 1 月 1 日之前收到的所有邮件。
            //after: 2023 / 01 / 01：获取在 2023 年 1 月 1 日之后收到的所有邮件。
            ListMessages(service, "me", "from:willy77423@gmail.com");

            Console.ReadLine();
        }

        // 获取Gmail API服务
        public static GmailService GetGmailService(string credentialsFilePath)
        {
            GoogleCredential credential;
            using (var stream = new FileStream(credentialsFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(GmailService.Scope.GmailReadonly);
            }

            // 创建Gmail API服务
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Gmail API C# Quickstart",
            });

            return service;
        }

        // 获取符合条件的邮件列表
        public static void ListMessages(GmailService service, string userId, string query)
        {
            var request = service.Users.Messages.List(userId);
            request.Q = query;

            var messages = request.Execute().Messages;
            if (messages != null && messages.Count > 0)
            {
                Console.WriteLine("符合条件的邮件：");
                foreach (var message in messages)
                {
                    var email = service.Users.Messages.Get(userId, message.Id).Execute();
                    // 获取主题
                    string subject = email.Payload.Headers.FirstOrDefault(h => h.Name == "Subject")?.Value;
                    // 打印主题
                    Console.WriteLine($"主题： {subject}");
                    // 获取发件人
                    string from = email.Payload.Headers.FirstOrDefault(h => h.Name == "From")?.Value;
                    // 打印发件人
                    Console.WriteLine($"发件人： {from}");
                    // 获取时间
                    string internalDate = email.InternalDate.ToString();
                    // 打印时间
                    Console.WriteLine($"时间： {internalDate}");
                    // 获取邮件内容
                    string body = GetEmailBody(email.Payload, service, userId, message.Id);
                    Console.WriteLine("---------");
                }
            }
            else
            {
                Console.WriteLine("没有找到符合条件的邮件。");
            }

        }

        // 获取邮件内容
        public static string GetEmailBody(MessagePart payload, GmailService service, string userId, string msgId)
        {
            if (payload.Body.Data != null)
            {
                byte[] data = Convert.FromBase64String(payload.Body.Data.Replace('-', '+').Replace('_', '/'));
                return Encoding.UTF8.GetString(data);
            }
            else if (payload.Body.AttachmentId != null)
            {
                string attachment = GetAttachment(service, userId, payload.Body.AttachmentId, msgId);
                return $"[Attachment: {attachment}]";
            }
            else if (payload.Parts != null)
            {
                string return_string = "";
                // 遍历所有部分，查找包含邮件内容的部分
                foreach (var part in payload.Parts)
                {
                    if (part.Body != null && part.Body.Data != null)
                    {
                        // 如果邮件内容在 part 中
                        return_string = DecodeBase64String(part.Body.Data);
                    }
                }
                return return_string;
            }
            else
            {
                return "";
            }
        }

        // 获取附件内容
        public static string GetAttachment(GmailService service, string userId, string attachmentId, string messageId)
        {
            var attachment = service.Users.Messages.Attachments.Get(userId, messageId, attachmentId).Execute();
            byte[] data = Convert.FromBase64String(attachment.Data.Replace('-', '+').Replace('_', '/'));
            return Encoding.UTF8.GetString(data);
        }

        // 解码 base64 编码的字符串
        public static string DecodeBase64String(string base64String)
        {
            byte[] data = Convert.FromBase64String(base64String.Replace('-', '+').Replace('_', '/'));
            return Encoding.UTF8.GetString(data);
        }
    }
}
