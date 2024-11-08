using ClosedXML.Excel;
using ClosedXML.Report.Options;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

namespace MyGridBot
{
    internal class TG
    {
        #region Переменные
        public static int SendReport { get; set; } = 100; // Если значение не указано, то отправлять отчет через 100 циклов Buy/Sell
        public static int Sorting { get; set; } = 50; // Если значение не указано, то сортировать монеты через 50 циклов Buy/Sell
        public static string Token { get; set; } = "";
        public static string ReportMini = "📊 Подготовка отчета";
        public static TelegramBotClient Client = new(Token);
        public static Chat Chat = new();
        private static readonly string PathTG = @"..\\..\\..\\..\\Telegram.xlsx"; // Расположение конфигурационного файла
        private static ITelegramBotClient _botClient;
        private static ReceiverOptions _receiverOptions;
        #endregion

        #region Отправка сообщений
        public static async Task SendMessageAsync(string message)
        {
            if (!string.IsNullOrEmpty(Token))
            {
                int retryCount = 0;
                while (retryCount < 3)
                {
                    try
                    {
                        await Client.SendTextMessageAsync(Chat.Id, message);
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при отправке сообщения: {ex.Message}");
                        await Task.Delay(1000);
                        retryCount++;
                    }
                }
            }
        }
        #endregion

        #region Чтение параметров из Telegram.xlsx
        public static async Task TGConfig()
        {
            using (var workbookTG = new XLWorkbook(PathTG))
            {
                var sheetTG = workbookTG.Worksheet(1);

                if (!sheetTG.Cell(1, 2).IsEmpty() && !sheetTG.Cell(2, 2).IsEmpty())
                {
                    Token = sheetTG.Cell(1, 2).GetString();
                    Chat.Id = Convert.ToInt64(sheetTG.Cell(2, 2).Value);
                    Client = new TelegramBotClient(Token);
                }
                else
                {
                    Console.WriteLine("Не указан Token или Id");
                }
            }
            SendMessageAsync("🤖 GridBoviBot подключен.\nНажмите /start чтобы включить ⌨").Wait();
            await TG.WaitMessage();
        }
        #endregion

        #region Ожидание сообщений от пользователя
        public static async Task WaitMessage()
        {
            _botClient = new TelegramBotClient(Token);
            _receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = new[] { UpdateType.Message },
                ThrowPendingUpdates = true,
            };
            using var cts = new CancellationTokenSource();
            _botClient.StartReceiving(UpdateHandler, ErrorHandler, _receiverOptions, cts.Token);
        }
        #endregion

        #region Подключение кнопок и обработка сообщений телеграм
        private static async Task SendReplyKeyboardAsync(long chatId, string message)
        {
            var replyKeyboard = new ReplyKeyboardMarkup(new[]
            {
                new KeyboardButton[] { "📊 Отчет", "💬 BOVI Флудилка" }
            })
            {
                ResizeKeyboard = true,
            };
            await Client.SendTextMessageAsync(chatId, message, replyMarkup: replyKeyboard);
        }
        private static async Task UpdateHandler(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            try
            {
                if (update.Type == UpdateType.Message && update.Message?.Type == MessageType.Text)
                {
                    var message = update.Message;
                    var chat = message.Chat;

                    switch (message.Text)
                    {
                        case "/start":
                            await SendReplyKeyboardAsync(chat.Id, "⌨️ Клавиатура подключена");
                            break;
                        case "📊 Отчет":
                            if (!string.IsNullOrEmpty(ReportMini))
                            {
                                await botClient.SendTextMessageAsync(chat.Id, ReportMini);
                            }
                            break;
                        case "💬 BOVI Флудилка":
                            await botClient.SendTextMessageAsync(chat.Id, "Перейдите по ссылке\n https://t.me/c/2046625015/1");
                            break;
                        default:
                            await botClient.SendTextMessageAsync(chat.Id, "Используй только кнопки!");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка в обработчике обновлений: {ex.Message}");
            }
        }
        private static Task ErrorHandler(ITelegramBotClient botClient, Exception error, CancellationToken cancellationToken)
        {
            var errorMessage = error switch
            {
                ApiRequestException apiRequestException => $"Telegram API Error:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
                _ => error.ToString()
            };

            Console.WriteLine(errorMessage);
            return Task.CompletedTask;
        }
        #endregion
    }
}
