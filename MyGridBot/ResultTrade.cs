﻿using Bybit.Net.Clients;
using Bybit.Net.Objects.Models.V5;
using ClosedXML.Excel;
using CryptoExchange.Net.Objects;
using Mexc.Net.Clients;
using Mexc.Net.Objects.Models.Spot;

namespace MyGridBot
{
    internal class ResultTrade
    {
        public static ulong Buy { get; set; } = 0;
        public static ulong Sell { get; set; } = 0;
        public static decimal TotalBalanceUSDT { get; set; } = 0;
        public static decimal TotalBalanceUSDC { get; set; } = 0;
        public static decimal ExpectedProfitUSDT { get; set; } = 0;
        public static decimal ExpectedProfitUSDC { get; set; } = 0;
        public static decimal USDTtotal { get; set; } = 0;
        public static decimal USDCtotal { get; set; } = 0;
        static int Copy { get; set; } = 0;
        public static long Flag = 0; // Кол-во сортировок до уведомления
        public static long Tflag = 0;
        public static decimal Tprofit = 0;
        public static decimal Cprofit = 0;

        public static async Task BalanceByBit(BybitRestClient bybitRestClient, DateTime dateTime)
        {
            Copy++;
            string Profit = "";
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(" Делаю запрос баланса");
            WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitBalance>> balance = null;
            while (true)
            {
                try { balance = await bybitRestClient.V5Api.Account.GetBalancesAsync(Bybit.Net.Enums.AccountType.Unified); }
                catch (Exception ex) { Console.WriteLine(ex.Message); await Task.Delay(1000); }
                if (balance.Error == null) { break; }
                else if (balance.Error.Code == 10002)
                {
                    Console.WriteLine($" {balance.Error.Code} {balance.Error.Message}");
                    await Task.Delay(2000);
                    continue;
                }
                else if (balance.Error.Code == 502)
                {
                    Console.WriteLine($" {balance.Error.Code} {balance.Error.Message}");
                    await Task.Delay(2000);
                    continue;
                }
                else if (balance.Error.Code != null && balance.Error.Message != null)
                {
                    await TG.SendMessageAsync($"🚨 Внимание!\nОшибка при запросе баланса.\n⛔️ Бот остановлен.");
                    Console.WriteLine($" Ошибка при запросе баланса \n" +
                                      $" {balance.Error.Code} {balance.Error.Message}");
                    Console.ReadLine();
                }
                else
                {
                    await Task.Delay(2000);
                }
            }

            TotalBalanceUSDT = 0;
            TotalBalanceUSDC = 0;
            ExpectedProfitUSDT = 0;
            ExpectedProfitUSDC = 0;

            Console.WriteLine();
            var dt = DateTime.Now;
            var timeElapsed = dt - dateTime;
            Console.WriteLine($" Время работы: \n" +
                              $" Дни: {timeElapsed.Days}  {timeElapsed.Hours:00}:{timeElapsed.Minutes:00}:{timeElapsed.Seconds:00}");
            foreach (var Symbol in SettingStart.SymbolList)
            {
                //USDT
                string asset = "";
                for (int i = 0; i < Symbol.Length - 4; i++)
                {
                    asset += Symbol[i];
                }
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{Symbol}.xlsx"))
                        {
                            var sheet = workbook.Worksheet(1);

                            await ComissionByBit(bybitRestClient, Symbol, Convert.ToDecimal(sheet.Cell(13, 16).Value));

                            if ('T' == Symbol.Last())
                            {
                                TotalBalanceUSDT += Convert.ToDecimal(sheet.Cell(1, 12).Value);
                                ExpectedProfitUSDT += Convert.ToDecimal(sheet.Cell(1, 18).Value);
                            }
                            else
                            {
                                TotalBalanceUSDC += Convert.ToDecimal(sheet.Cell(1, 12).Value);
                                ExpectedProfitUSDC += Convert.ToDecimal(sheet.Cell(1, 18).Value);
                            }

                            foreach (var coins in balance.Data.List)
                            {
                                foreach (var coin in coins.Assets)
                                {
                                    if (coin.Asset == asset)
                                    {
                                        if (coin.WalletBalance < Convert.ToDecimal(sheet.Cell(1, 13).Value))
                                        {
                                            await TG.SendMessageAsync($"🚨 Внимание!\n {asset} на счете меньше чем нужно\n для сетки, на {Convert.ToDecimal(sheet.Cell(1, 13).Value) - coin.WalletBalance}\n ⛔️ Бот остановлен.");
                                            Console.WriteLine($" Монета:{asset} меньше в наличии чем в ексель,на {Convert.ToDecimal(sheet.Cell(1, 13).Value) - coin.WalletBalance}");
                                            Console.ReadLine();
                                        }
                                        else
                                        {
                                            Profit = $"🪙 {coin.WalletBalance - Convert.ToDecimal(sheet.Cell(1, 13).Value)} {asset}\n";
                                            Console.Write($" Монета: ");
                                            Console.ForegroundColor = ConsoleColor.Gray;
                                            Console.Write($"{asset} ");
                                            Console.ForegroundColor = ConsoleColor.Magenta;
                                            Console.Write(" Профит: ");
                                            Console.ForegroundColor = ConsoleColor.Yellow;
                                            Console.Write($"{coin.WalletBalance - Convert.ToDecimal(sheet.Cell(1, 13).Value)}");
                                            Console.ForegroundColor = ConsoleColor.Magenta;
                                            Console.WriteLine();
                                            break;
                                        }
                                    }
                                }

                            }
                        }
                        break;
                    }
                    catch
                    {
                        Console.WriteLine($" Не смог открыть файл {Symbol}.xlsx метод Balance");
                        await TG.SendMessageAsync($"🚨 Внимание!\n Не смог открыть файл {Symbol}.xlsx\n метод Balance. Требуется Ваша помощь.");
                        Thread.Sleep(10000);
                    }
                }
            }
            foreach (var coins in balance.Data.List)
            {
                foreach (var coin in coins.Assets)
                {
                    if (coin.Asset == "USDT")
                    {
                        if (coin.WalletBalance < TotalBalanceUSDT)
                        {
                            Console.WriteLine($" USDT на счете меньше чем нужно для сетки, на {TotalBalanceUSDT - coin.WalletBalance}");
                            await TG.SendMessageAsync($"🚨 Внимание!\n USDT на счете меньше чем нужно\n для сетки, на {TotalBalanceUSDT - coin.WalletBalance}\n ⛔️ Бот остановлен.");
                            Console.ReadLine();
                        }
                        else
                        {
                            Profit += $"💵 {Math.Round(coin.WalletBalance - TotalBalanceUSDT, 2)} USDT\n";
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($" USDT");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.Write(" Профит: ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{coin.WalletBalance - TotalBalanceUSDT} $");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine();

                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.Write(" Ожидаемая стомость портфеля: ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{coin.WalletBalance + ExpectedProfitUSDT} $");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine();
                            USDTtotal = coin.WalletBalance;
                        }
                    }
                    if (coin.Asset == "USDC")
                    {
                        if (coin.WalletBalance < TotalBalanceUSDC)
                        {
                            Console.WriteLine($" USDC на счете меньше чем нужно для сетки,на {TotalBalanceUSDC - coin.WalletBalance}");
                            await TG.SendMessageAsync($"🚨 Внимание!\n USDC на счете меньше чем нужно\n для сетки, на {TotalBalanceUSDC - coin.WalletBalance}\n ⛔️ Бот остановлен.");
                            Console.ReadLine();
                        }
                        else
                        {
                            Profit += $"💵 {Math.Round(coin.WalletBalance - TotalBalanceUSDT, 2)} USDC\n";
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($" USDC");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.Write(" Профит: ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{coin.WalletBalance - TotalBalanceUSDC} $");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine();

                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.Write(" Ожидаемая стомость портфеля: ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{coin.WalletBalance + ExpectedProfitUSDC} $");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine();
                            USDCtotal = coin.WalletBalance;
                        }
                    }
                }
            }
            Console.WriteLine();
            Console.Write($" Сделки: Buy: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($"{Buy}");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.Write(" Sell: ");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write($"{Sell}");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine();

            // Проверка на конец сетки
            string EndOrder = "";
            if (Trader.BuyEndOrder != "")
            {
                EndOrder += $"{Trader.BuyEndOrder}";
            }
            if (Trader.SellEndOrder != "")
            {
                EndOrder += $"{Trader.SellEndOrder}";
            }

            #region Основная логика формирования отчетов в телеграм
            decimal
            rTtotal = Math.Round(USDTtotal, 2),
            rCtotal = Math.Round(USDCtotal, 2),
            rUSDT = Math.Round(rTtotal + ExpectedProfitUSDT, 2),
            rUSDC = Math.Round(rCtotal + ExpectedProfitUSDC, 2);
            ulong rBuy = Buy;
            ulong rSell = Sell;
            if (Tflag == 0)
            {
                Tprofit = rUSDT;
                Cprofit = rUSDC;
                Tflag++;
            }
            Tprofit = rUSDT - Tprofit;
            Cprofit = rUSDC - Cprofit;

            string RMiniUSD;
            if (TotalBalanceUSDC == 0)
            {
                RMiniUSD = $"📊 USDT: {rUSDT}  🚀 {Tprofit}\n" +
                           $"💼 USDT: {rTtotal}\n";
            }
            else
            {
                RMiniUSD = $"📊 USDC: {rUSDC}  🚀 {Cprofit}\n" +
                           $"💼 USDC: {rCtotal}\n";
            }

            // Формирование отчетов
            TG.ReportMini = GenerateMiniReport("BYBIT", timeElapsed, RMiniUSD, Profit, rBuy, rSell, EndOrder);

            // Уведомление о сортировке
            if (TG.Sorting - Copy > 0)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine();
                Console.Write(" Сортировка будет через (");
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write($"{TG.Sorting - Copy}");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(") прокрутов");
                Console.ForegroundColor = ConsoleColor.Magenta;
            }
            else
            {
            // Отправка отчета
            await TG.SendMessageAsync(TG.ReportMini);
            }
            #endregion

            // Сортировка и резервное копирование
            if (Copy >= TG.Sorting)
            {
                //Сортировка
                Console.ForegroundColor = ConsoleColor.Blue;
                await NewExcel.SortBuySellByBitAsync(bybitRestClient);
                Copy = 0;
                CopyTable.Copy(@"..\\..\\..\\..\\Work", @"..\\..\\..\\..\\WorkCopy");
            }
        }
        public static async Task BalanceMexc(MexcRestClient mexcRestClient, DateTime dateTime)
        {
            Copy++;
            string Profit = "";
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(" Делаю запрос баланса");
            WebCallResult<MexcAccountInfo> balance = null;
            while (true)
            {
                try { balance = await mexcRestClient.SpotApi.Account.GetAccountInfoAsync(); }
                catch (Exception ex) { Console.WriteLine(ex.Message); await Task.Delay(1000); }
                if (balance.Error == null) { break; }
                else if (balance.Error.Code == 700022) { continue; }
                else if (balance.Error.Code == 700003) { continue; }
                else if (balance.Error.Code == 504) { await Task.Delay(3000); continue; }
                else if (balance.Error.Code != null && balance.Error.Message != null)
                {
                    await TG.SendMessageAsync($"🚨 Внимание!\nОшибка при запросе баланса.\n⛔️ Бот остановлен.");
                    Console.WriteLine($" Ошибка при запросе баланса \n" +
                                      $" {balance.Error.Code} {balance.Error.Message}");
                    Console.ReadLine();
                }
                else
                {
                    await Task.Delay(2000);
                }
            }

            TotalBalanceUSDT = 0;
            TotalBalanceUSDC = 0;
            ExpectedProfitUSDT = 0;
            ExpectedProfitUSDC = 0;

            Console.WriteLine();
            var dt = DateTime.Now;
            var timeElapsed = dt - dateTime;
            Console.WriteLine($" Время работы: \n" +
                              $" Дни: {timeElapsed.Days}  {timeElapsed.Hours:00}:{timeElapsed.Minutes:00}:{timeElapsed.Seconds:00}");
            foreach (var Symbol in SettingStart.SymbolList)
            {
                //USDT
                string asset = "";
                for (int i = 0; i < Symbol.Length - 4; i++)
                {
                    asset += Symbol[i];
                }
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\WorkMexc\\{Symbol}.xlsx"))
                        {
                            var sheet = workbook.Worksheet(1);

                            await ComissionMexc(mexcRestClient, Symbol, Convert.ToDecimal(sheet.Cell(13, 16).Value));

                            if ('T' == Symbol.Last())
                            {
                                TotalBalanceUSDT += Convert.ToDecimal(sheet.Cell(1, 12).Value);
                                ExpectedProfitUSDT += Convert.ToDecimal(sheet.Cell(1, 18).Value);
                            }
                            else
                            {
                                TotalBalanceUSDC += Convert.ToDecimal(sheet.Cell(1, 12).Value);
                                ExpectedProfitUSDC += Convert.ToDecimal(sheet.Cell(1, 18).Value);
                            }

                            foreach (var coin in balance.Data.Balances)
                            {
                                if (coin.Asset == asset)
                                {
                                    if (coin.Total < Convert.ToDecimal(sheet.Cell(1, 13).Value))
                                    {
                                        await TG.SendMessageAsync($"🚨 Внимание!\n {asset} на счете меньше чем нужно\n для сетки, на {Convert.ToDecimal(sheet.Cell(1, 13).Value) - coin.Total}\n ⛔️ Бот остановлен.");
                                        Console.WriteLine($" Монета:{asset} меньше в наличии чем в ексель,на {Convert.ToDecimal(sheet.Cell(1, 13).Value) - coin.Total}");
                                        Console.ReadLine();
                                    }
                                    else
                                    {
                                        Profit = $"🪙 {coin.Total - Convert.ToDecimal(sheet.Cell(1, 13).Value)} {asset}\n";
                                        Console.Write($" Монета: ");
                                        Console.ForegroundColor = ConsoleColor.Gray;
                                        Console.Write($"{asset} ");
                                        Console.ForegroundColor = ConsoleColor.Magenta;
                                        Console.Write(" Профит: ");
                                        Console.ForegroundColor = ConsoleColor.Yellow;
                                        Console.Write($"{coin.Total - Convert.ToDecimal(sheet.Cell(1, 13).Value)}");
                                        Console.ForegroundColor = ConsoleColor.Magenta;
                                        Console.WriteLine();
                                        break;
                                    }
                                }
                            }
                        }
                        break;
                    }
                    catch
                    {
                        Console.WriteLine($" Не смог открыть файл {Symbol}.xlsx метод BalanceMexc");
                        await TG.SendMessageAsync($"🚨 Внимание!\n Не смог открыть файл {Symbol}.xlsx\n метод Balance. Требуется Ваша помощь.");
                        Thread.Sleep(10000);
                    }
                }
            }
            foreach (var coin in balance.Data.Balances)
            {
                if (coin.Asset == "USDT")
                {
                    if (coin.Total < TotalBalanceUSDT)
                    {
                        Console.WriteLine($" USDT на счете меньше чем нужно для сетки,на {TotalBalanceUSDT - coin.Total}");
                        await TG.SendMessageAsync($"🚨 Внимание!\n USDT на счете меньше чем нужно\n для сетки, на {TotalBalanceUSDT - coin.Total}\n ⛔️ Бот остановлен.");
                        Console.ReadLine();
                    }
                    else
                    {
                        Profit += $"💵 {Math.Round(coin.Total - TotalBalanceUSDT, 2)} USDT\n";
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($" USDT");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(" Профит: ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{coin.Total - TotalBalanceUSDT} $");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine();

                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(" Ожидаемая стомость портфеля: ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{coin.Total + ExpectedProfitUSDT} $");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine();
                        USDTtotal = coin.Total;
                    }
                }
                if (coin.Asset == "USDC")
                {
                    if (coin.Total < TotalBalanceUSDC)
                    {
                        Console.WriteLine($" USDC на счете меньше чем нужно для сетки,на {TotalBalanceUSDC - coin.Total}");
                        await TG.SendMessageAsync($"🚨 Внимание!\n USDC на счете меньше чем нужно\n для сетки, на {TotalBalanceUSDT - coin.Total}\n ⛔️ Бот остановлен.");
                        Console.ReadLine();
                    }
                    else
                    {
                        Profit += $"💶 {Math.Round(coin.Total - TotalBalanceUSDT, 2)} USDC\n";
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($" USDC");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(" Профит: ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{coin.Total - TotalBalanceUSDC} $");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine();

                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(" Ожидаемая стомость портфеля: ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{coin.Total + ExpectedProfitUSDC} $");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine();
                        USDCtotal = coin.Total;
                    }
                }
            }
            Console.WriteLine();
            Console.Write($" Сделки: Buy: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($"{Buy}");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.Write(" Sell: ");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write($"{Sell}");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine();

            // Проверка на конец сетки
            string EndOrder = "";
            if (Trader.BuyEndOrder != "")
            {
                EndOrder += $"{Trader.BuyEndOrder}";
            }
            if (Trader.SellEndOrder != "")
            {
                EndOrder += $"{Trader.SellEndOrder}";
            }


            #region Основная логика формирования отчетов в телеграм
            decimal
            rTtotal = Math.Round(USDTtotal, 2),
            rCtotal = Math.Round(USDCtotal, 2),
            rUSDT = Math.Round(rTtotal + ExpectedProfitUSDT, 2),
            rUSDC = Math.Round(rCtotal + ExpectedProfitUSDC, 2);
            ulong rBuy = Buy;
            ulong rSell = Sell;
            if (Tflag == 0)
            {
                Tprofit = rUSDT;
                Cprofit = rUSDC;
                Tflag++;
            }
            Tprofit = rUSDT - Tprofit;
            Cprofit = rUSDC - Cprofit;

            string RMiniUSD;
            if (TotalBalanceUSDC == 0)
            {
                RMiniUSD = $"📊 USDT: {rUSDT}  🚀 {Tprofit}\n" +
                           $"💼 USDT: {rTtotal}\n";
            }
            else
            {
                RMiniUSD = $"📊 USDC: {rUSDC}  🚀 {Cprofit}\n" +
                           $"💼 USDC: {rCtotal}\n";
            }

            // Формирование отчетов
            TG.ReportMini = GenerateMiniReport("MEXC", timeElapsed, RMiniUSD, Profit, rBuy, rSell, EndOrder);

            // Уведомление о сортировке
            if (TG.Sorting - Copy > 0)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine();
                Console.Write(" Сортировка будет через (");
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write($"{TG.Sorting - Copy}");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(") прокрутов");
                Console.ForegroundColor = ConsoleColor.Magenta;
            }
            else
            {
                // Отправка отчета
                await TG.SendMessageAsync(TG.ReportMini);
            }
            #endregion

            // Сортировка и резервное копирование
            if (Copy >= TG.Sorting)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                await NewExcel.SortBuySellMexcAsync();
                Copy = 0;
                CopyTable.Copy(@"..\\..\\..\\..\\WorkMexc", @"..\\..\\..\\..\\WorkCopyMexc");
            }
        }
        public static async Task TimerReversAsync(int seconds, BybitRestClient bybitRestClient)
        {
            bool isPaused = false;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();
            Console.WriteLine(" Нажмите ПРОБЕЛ для остановки\n Что бы редактировать все ексель файлы");
            var dateTime = DateTime.Now;
            DateTime dt = dateTime.AddSeconds(-seconds);

            ConsoleKeyInfo keyInfo;
            while (dateTime >= dt)
            {
                if (Console.KeyAvailable)
                {
                    keyInfo = Console.ReadKey(true);
                    if (keyInfo.Key == ConsoleKey.Spacebar)
                    {
                        isPaused = !isPaused;
                        Console.WriteLine(isPaused ? " Таймер приостановлен. \n Можно редактировать ексель\n Нажмите ПРОБЕЛ для продолжения." : " Таймер продолжает работу.");
                        if (!isPaused)
                        {
                            await SettingStart.StartNewExelAsync(bybitRestClient);
                        }
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }

                if (!isPaused)
                {
                    var ticks = (dateTime - dt).Ticks;
                    Console.WriteLine(new DateTime(ticks).ToString("      HH:mm:ss"));
                    Thread.Sleep(850);
                    dt = dt.AddSeconds(1);
                }
            }
        }
        public static async Task TimerReversAsyncMexc(int seconds, MexcRestClient mexcRestClient)
        {
            bool isPaused = false;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();
            Console.WriteLine(" Нажмите ПРОБЕЛ для остановки\n Что бы редактировать все ексель файлы");
            var dateTime = DateTime.Now;
            DateTime dt = dateTime.AddSeconds(-seconds);

            ConsoleKeyInfo keyInfo;
            while (dateTime >= dt)
            {
                if (Console.KeyAvailable)
                {
                    keyInfo = Console.ReadKey(true);
                    if (keyInfo.Key == ConsoleKey.Spacebar)
                    {
                        isPaused = !isPaused;
                        Console.WriteLine(isPaused ? " Таймер приостановлен. \n Можно редактировать ексель\n Нажмите ПРОБЕЛ для продолжения." : " Таймер продолжает работу.");
                        if (!isPaused)
                        {
                            await SettingStart.StartNewExelAsyncMexc(mexcRestClient);
                        }
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }

                if (!isPaused)
                {
                    var ticks = (dateTime - dt).Ticks;
                    Console.WriteLine(new DateTime(ticks).ToString("      HH:mm:ss"));
                    Thread.Sleep(850);
                    dt = dt.AddSeconds(1);
                }
            }
        }
        public static async Task<bool> ComissionByBit(BybitRestClient bybitRestClient, string symbol, decimal comission)
        {
            BybitFeeRate FeeByBit = null;
            while (true)
            {
                try
                {
                    FeeByBit = ((await bybitRestClient.V5Api.Account.GetFeeRateAsync(Bybit.Net.Enums.Category.Spot, symbol)).Data.List.First());
                    if (FeeByBit == null || FeeByBit.BaseAsset == null) { await Task.Delay(1000); continue; }
                    else
                    {
                        if (FeeByBit.TakerFeeRate * 100 != comission)
                        {
                            if (FeeByBit.TakerFeeRate * 100 > comission)
                            {
                                await TG.SendMessageAsync($"🚨 Внимание! ⛔️ Бот остановлен.\n" +
                                                          $"Комиссия биржи стала выше,\n" +
                                                          $"а именно: {FeeByBit.TakerFeeRate * 100}%\n" +
                                                          $"Измените комиссию в таблице:\n" +
                                                          $"{symbol}.xlsx и нажмите ENTER");
                                Console.WriteLine();
                                Console.WriteLine($" Комиссия биржи стала выше, а именно: {FeeByBit.TakerFeeRate * 100} %\n" +
                                                  $" Измените комиссию в ексель: {symbol}.xlsx\n" +
                                                  $" Нажмите ENTER");
                                Console.ReadLine();
                                return true;
                            }
                            else if (FeeByBit.TakerFeeRate * 100 < comission)
                            {
                                Console.WriteLine();
                                Console.WriteLine($" Комиссия биржи стала ниже, а именно: {FeeByBit.TakerFeeRate * 100} %\n" +
                                                  $" Можно изменить комиссию в ексель: {symbol}.xlsx");
                                return true;
                            }
                        }
                        else
                        {
                            return true;
                        }
                    }

                }
                catch
                {
                    await Task.Delay(1000);
                    continue;
                }
            }

        }
        public static async Task<bool> ComissionMexc(MexcRestClient mexcRestClient, string symbol, decimal comission)
        {
            MexcTradeFee FeeMexc = null;
            while (true)
            {
                try
                {
                    FeeMexc = (await mexcRestClient.SpotApi.Account.GetTradeFeeAsync(symbol)).Data;
                    if (FeeMexc == null) { await Task.Delay(1000); continue; }
                    else
                    {
                        if (FeeMexc.TakerFee * 100 != comission)
                        {
                            if (FeeMexc.TakerFee * 100 > comission)
                            {
                                await TG.SendMessageAsync($"🚨 Внимание! ⛔️ Бот остановлен.\n" +
                                                          $"Комиссия биржи стала выше,\n" +
                                                          $"а именно: {FeeMexc.TakerFee * 100}%\n" +
                                                          $"Измените комиссию в таблице:\n" +
                                                          $"{symbol}.xlsx и нажмите ENTER");
                                Console.WriteLine();
                                Console.WriteLine($" Комиссия биржи стала выше, а именно: {FeeMexc.TakerFee * 100} %\n" +
                                                  $" Измените комиссию в ексель: {symbol}.xlsx\n" +
                                                  $" Нажмите ENTER");
                                Console.ReadLine();
                                return true;
                            }
                            else if (FeeMexc.TakerFee * 100 < comission)
                            {
                                Console.WriteLine();
                                Console.WriteLine($" Комиссия биржи стала ниже, а именно: {FeeMexc.TakerFee * 100} %\n" +
                                                  $" Можно изменить комиссию в ексель: {symbol}.xlsx");
                                return true;
                            }
                        }
                        else
                        {
                            return true;
                        }
                    }

                }
                catch
                {
                    await Task.Delay(1000);
                    continue;
                }
            }
        }
        private static string GenerateMiniReport(string exchangeName, TimeSpan timeElapsed,
                                           string miniUSD, string profit, ulong rBuy, ulong rSell, string endOrder)
        {
            return $"🏦 Биржа {exchangeName}\n" +
                   $"📆 Дни: {timeElapsed.Days} ⌚️ {timeElapsed.Hours:00}:{timeElapsed.Minutes:00}:{timeElapsed.Seconds:00}\n" +
                   $"💰 Профит:\n" +
                   $"{profit}" +
                   $"{miniUSD}" +
                   $"⚖️ 📉 Buy: {rBuy} 📈 Sell: {rSell}\n" +
                   $"{endOrder}";
        }
    }
}
