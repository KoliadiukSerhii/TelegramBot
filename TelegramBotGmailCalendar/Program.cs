using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TelegramBotGmailCalendar;
using TelegramBotGmailCalendar.APIHelper;
using TelegramBotGmailCalendar.APIHelpers;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using GmailAPI;
using Telegram.Bot;
using Telegram.Bot.Extensions;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using Telegram.Bot.Extensions.Polling;
using Telegram.Bot.Exceptions;

namespace GmailAPI
{
    class Program
    {
        static TelegramBotClient telegramBotClient = new TelegramBotClient
            ("5443198139:AAHp2DTIfQx2ydVh--iOnxvZ8dxoYno5Iv4");
        static ReceiverOptions receiverOptions = new ReceiverOptions { AllowedUpdates = { } };
        static CancellationTokenSource cts = new CancellationTokenSource();
        static GmailService GmailService = null;
        static CalendarService CalendarService = null;
        static string username = "";
        static string recipient = "";
        static string subject = "";
        static string mailText = "";
        static string CalendarEvents = "";

        static void Main(string[] args)
        {
            try
            {
                telegramBotClient.StartReceiving(HandlerUpdateAsync, HandlerError, receiverOptions, cancellationToken: cts.Token);
                Console.WriteLine("Bot GmailCalendarBot started!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
            Console.ReadKey();
        }

        public static List<Gmail> GetAllEmails(string HostEmailAddress)
        {
            try
            {
                List<Gmail> EmailList = new List<Gmail>();
                UsersResource.MessagesResource.ListRequest ListRequest = GmailService.Users.Messages.List(HostEmailAddress);
                ListRequest.LabelIds = "INBOX";
                ListRequest.IncludeSpamTrash = false;
                ListRequest.Q = "is:unread";

                ListMessagesResponse ListResponse = ListRequest.Execute();

                if (ListResponse != null && ListResponse.Messages != null)
                {
                    foreach (Google.Apis.Gmail.v1.Data.Message Msg in ListResponse.Messages)
                    {
                        GmailAPIHelper.MsgMarkAsRead(HostEmailAddress, Msg.Id);

                        UsersResource.MessagesResource.GetRequest Message = GmailService.Users.Messages.Get(HostEmailAddress, Msg.Id);
                        Console.WriteLine("\n-----------------NEW MAIL----------------------");
                        Console.WriteLine("Message ID:" + Msg.Id);

                        Google.Apis.Gmail.v1.Data.Message MsgContent = Message.Execute();

                        if (MsgContent != null)
                        {
                            string FromAddress = string.Empty;
                            string ToAddress = string.Empty;
                            string Date = string.Empty;
                            string Subject = string.Empty;
                            string MailBody = string.Empty;
                            string ReadableText = string.Empty;

                            foreach (var MessageParts in MsgContent.Payload.Headers)
                            {
                                if (MessageParts.Name == "From")
                                {
                                    FromAddress = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "To")
                                {
                                    ToAddress = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Date")
                                {
                                    Date = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Subject")
                                {
                                    Subject = MessageParts.Value;
                                }
                            }

                            Console.WriteLine("Reading Mail Body:");
                            List<string> FileName = GmailAPIHelper.GetAttachments(HostEmailAddress, Msg.Id, Convert.ToString(ConfigurationManager.AppSettings["GmailAttach"]));

                            if (FileName.Count() > 0)
                            {
                                foreach (var EachFile in FileName)
                                {
                                    string[] RectifyFromAddress = FromAddress.Split(' ');
                                    string FromAdd = RectifyFromAddress[RectifyFromAddress.Length - 1];

                                    if (!string.IsNullOrEmpty(FromAdd))
                                    {
                                        FromAdd = FromAdd.Replace("<", string.Empty);
                                        FromAdd = FromAdd.Replace(">", string.Empty);
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Mail has no attachments.");
                            }

                            MailBody = string.Empty;
                            if (MsgContent.Payload.Parts == null && MsgContent.Payload.Body != null)
                            {
                                MailBody = MsgContent.Payload.Body.Data;
                            }
                            else
                            {
                                MailBody = GmailAPIHelper.MsgNestedParts(MsgContent.Payload.Parts);
                            }

                            ReadableText = string.Empty;
                            ReadableText = GmailAPIHelper.Base64Decode(MailBody);

                            Console.WriteLine("Identifying & Configure Mails.");

                            if (!string.IsNullOrEmpty(ReadableText))
                            {
                                Gmail GMail = new Gmail();
                                GMail.From = FromAddress;
                                GMail.To = ToAddress;
                                GMail.Body = ReadableText;
                                GMail.MailDateTime = Convert.ToDateTime(Date);
                                EmailList.Add(GMail);
                            }
                        }
                    }
                }
                return EmailList;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
                return null;
            }
        }

        private static async Task HandlerUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            if (update.Type == UpdateType.Message && update?.Message?.Text != null)
            {
                username = update.Message.Chat.Username;
                await MessageHandlerAsync(botClient, update.Message);
            }
            if (update?.Type == UpdateType.CallbackQuery)
            {
                await HandlerCallbackQuery(botClient, update.CallbackQuery);
            }
        }

        private static Task HandlerError(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
        {
            var errorMessage = exception switch
            {
                ApiRequestException apiRequestException => $"Telegram Bot API error:\n{apiRequestException.ErrorCode}" +
                $"\n{apiRequestException.Message}",
                _ => exception.ToString()
            };
            Console.WriteLine(errorMessage);
            return Task.CompletedTask;
        }

        private static async Task MessageHandlerAsync(ITelegramBotClient botClient, Telegram.Bot.Types.Message message)
        {
            if (message.Text.StartsWith('@'))
            {
                string[] TEXT = message.Text.Split(' ');

                if (TEXT[1] == "To:")
                {
                    recipient = string.Empty;
                    recipient = TEXT[2];
                }
                else if (TEXT[1] == "Subject:")
                {
                    subject = string.Empty;
                    for (int i = 2; i < TEXT.Length; i++)
                        subject += TEXT[i] + " ";
                }
                else if (TEXT[1] == "Message:")
                {
                    mailText = string.Empty;
                    for (int i = 2; i < TEXT.Length; i++)
                        mailText += TEXT[i] + " ";
                }
                return;
            }

            switch (message.Text)
            {
                case "/start":
                    InlineKeyboardMarkup keyboardMarkup = new
                        (
                            new[]
                            {
                            new[]
                            {
                            InlineKeyboardButton.WithCallbackData("Authorize me", "AuthorizeMe")
                            }
                            }
                        );
                    await botClient.SendTextMessageAsync(message.Chat.Id, $"Hi {username}! Please authorize me to set up a Gmail integration:", replyMarkup: keyboardMarkup);
                    break;
                case "/draft":
                    if (GmailService != null)
                    {
                        InlineKeyboardMarkup inlineKeyboardMarkup = new(
                                new[]
                                {
                                    new[]
                                    {
                                        InlineKeyboardButton.WithSwitchInlineQueryCurrentChat("To:", "To: "),
                                        InlineKeyboardButton.WithSwitchInlineQueryCurrentChat("Subject:", "Subject: "),
                                        InlineKeyboardButton.WithSwitchInlineQueryCurrentChat("Message:", "Message: "),
                                    }
                                });
                        await botClient.SendTextMessageAsync(message.Chat.Id, "Please specify a recipient, an optional subject and the text of the email.", replyMarkup: inlineKeyboardMarkup);
                    }
                    else
                        await botClient.SendTextMessageAsync(message.Chat.Id, "You must log in first.");
                    break;
                case "/send":
                    if (GmailService != null)
                    {
                        Console.WriteLine($"Recipient: {recipient}\r\nSubject: {subject}\r\nMessage: {mailText}");
                        string messageToSend = $"To: {recipient}\r\nSubject: {subject}\r\nContent-Type: text/html;charset=utf-8\r\n{mailText}";
                        Google.Apis.Gmail.v1.Data.Message msg = new Google.Apis.Gmail.v1.Data.Message();
                        msg.Raw = Base64UrlEncode(messageToSend.ToString());
                        GmailService.Users.Messages.Send(msg, "me").Execute();
                        await botClient.SendTextMessageAsync(message.Chat.Id, $"Message sent to {recipient}");
                        Console.WriteLine($"Message sent to {recipient}");
                    }
                    else
                        await botClient.SendTextMessageAsync(message.Chat.Id, "You must log in first.");
                    break;
                case "/mail":
                    if (GmailService != null)
                    {
                        try
                        {
                            List<Gmail> MailLists = GetAllEmails(Convert.ToString(ConfigurationManager.AppSettings["HostAddress"]));
                            if (MailLists.Count == 0 || MailLists == null)
                            {
                                await botClient.SendTextMessageAsync(message.Chat.Id, "You have no unread emails.");
                                return;
                            }

                            string mailText = "";
                            foreach (Gmail mail in MailLists)
                            {
                                mailText += mail.To; mailText += "\n";
                                mailText += mail.From; mailText += "\n";
                                mailText += mail.MailDateTime; mailText += "\n";
                                mailText += mail.Body; mailText += "\n";

                                await botClient.SendTextMessageAsync(message.Chat.Id, mailText);

                                Console.WriteLine("------------------------------------");
                                Console.WriteLine("To: " + mail.To);
                                Console.WriteLine("From: " + mail.From);
                                Console.WriteLine("MailDateTime: " + mail.MailDateTime);
                                Console.WriteLine("Body: " + mail.Body);
                                Console.WriteLine("------------------------------------");
                            }
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine("Error: " + exception);
                        }
                    }
                    else
                        await botClient.SendTextMessageAsync(message.Chat.Id, "You must log in first.");
                    break;
                case "/agenda":
                    if (CalendarService != null)
                    {
                        string upcomingEvents = GetUpcomingEvents();
                        await botClient.SendTextMessageAsync(message.Chat.Id, upcomingEvents);
                    }
                    else
                        await botClient.SendTextMessageAsync(message.Chat.Id, "You must log in first.");
                    break;
                case "/stop":
                    await botClient.SendTextMessageAsync(message.Chat.Id, "You will no longer receive notifications about new emails/events");
                    //cts.Cancel();
                    break;
            }
        }

        private static async Task HandlerCallbackQuery(ITelegramBotClient botClient, CallbackQuery? callbackQuery)
        {
            if (callbackQuery.Data.StartsWith("AuthorizeMe"))
            {
                try
                {
                    GmailService = GmailAPIHelper.GetService();
                    CalendarService = CalendarHelpers.GetCalendarService();
                }
                catch (Exception exception)
                {
                    Console.WriteLine("Error: " + exception);
                }
                return;
            }
            else
                await botClient.SendTextMessageAsync(callbackQuery.Message.Chat.Id, text: "Please press the button");
            return;
        }

        private static string Base64UrlEncode(string input)
        {
            var data = Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(data).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }

        public static string GetUpcomingEvents()
        {
            return "No Upcoming Events";
            /*
            try
            {
                EventsResource.ListRequest request = CalendarHelpers.GetRequest(CalendarService);
                Events events = request.Execute();

                if (events == null)
                    return "No Upcoming Events";

                if (events.Items != null && events.Items.Count > 0)
                {
                    foreach (var item in events.Items)
                        CalendarEvents += item.Summary + Environment.NewLine;
                }
                else
                    CalendarEvents = "No Upcoming Events";

                return CalendarEvents;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
                return null;
            }*/
        }
    }
}

