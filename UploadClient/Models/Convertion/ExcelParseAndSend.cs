﻿using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using ServiceReference1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.ServiceModel;
using System.Threading.Tasks;
using UploadClient.Models.AbsDTO;

namespace UploadClient.Models.Convertion
{
    public class ExcelParseAndSend : IExcelParseAndSend
    {
        public async Task<string> ParseAndSend(IFormFile file)
        {
            var workbook2 = new XLWorkbook(file.OpenReadStream());
            // 5. Распарсили Excel-файл, получили список платёжек для загрузки в Обработку платежей (АБС: Бэк-офис/Кредиты/Обработка платежей)
            var paymentsFromExcel = LoadFromExcel(workbook2, "Документы").ToArray();

            // 6. Отправили все платёжки в буфер необработанных платежей
            var massInsertPayments = GetPaymentsForMassInsert(paymentsFromExcel);
            var response = await UploadPaymentsToBufferAsync(massInsertPayments);
            var uploadResult = response.DsPaymentBufferMassInsertRes;

            // 7. Обработать результат отправки платёжек на предмет ошибок загрузки
            ProcessUploadPaymentErrors(paymentsFromExcel, uploadResult);

            // 8. Получить список пар: (Id загруженной платёжки, Id кредитного договора) на основе списка платежей, для которых заполнено поле НомерКредитногоДоговора
            var paymentsToBind = await GetPaymentsToBind(paymentsFromExcel, massInsertPayments, uploadResult);

            // 9. Сделать автоматическую привязку к кредитным договорам платежей, для которых заполнено поле НомерКредитногоДоговора
            await BindPaymentsToCreditContractsAsync(paymentsToBind);
        
            return "Ok";
        }

        /// <summary> Загрузка платёжных поручений из Excel. Столбец </summary>
        static IEnumerable<PaymentOrder> LoadFromExcel(IXLWorkbook workbook, string worksheetName)
        {
            var table = workbook.Worksheet(worksheetName).Tables.Single();

            foreach (var row in table.RowsUsed())
            {
                if (row.RowNumber() == 1) continue;
                yield return new PaymentOrder
                {
                    CreditContractNumber = row.Cell(1).GetString(),
                    Number = row.Cell(2).GetString(),
                    IncomeDate = row.Cell(3).GetDateTime(),
                    Amount = Convert.ToDecimal(row.Cell(4).GetString().Replace(" ", "").Replace(',', '.'), CultureInfo.InvariantCulture),
                    Description = row.Cell(5).GetString(),
                    PayerName = row.Cell(6).GetString()
                };
            }
        }
        static TPaymentListTypeForDSPaymentBufferMassInsert[] GetPaymentsForMassInsert(IEnumerable<PaymentOrder> paymentOrders)
        {
            var massInsertPayments = paymentOrders
                                     .Select((payment, i) =>
                                                 new TPaymentListTypeForDSPaymentBufferMassInsert
                                                 {
                                                     GUID = BitConverter.ToInt64(Guid.NewGuid().ToByteArray(), 0),
                                                     ModuleBrief = "VTBR2",
                                                     Date = payment.IncomeDate,
                                                     Amount = payment.Amount,
                                                     Comment = payment.Description,
                                                     Number = payment.Number,
                                                     PayerName = payment.PayerName,
                                                     PaymentKind = 3,
                                                     LinkID = i // ВАЖНО: обработка ошибок будет ссылаться на этот порядковый номер.
                                                                // Связь TPaymentListTypeForDSPaymentBufferMassInsert[] и PaymentOrder[] - по порядковому номеру i
                                                  })
                                     .ToArray();
            return massInsertPayments;
        }

        /// <summary>
        /// Загрузка платёжек в буфер необработанных платежей: Бэк-офис/Кредиты/Обработка платежей
        /// </summary>
        /// <param name="massInsertPayments"></param>
        /// <returns></returns>
        static async Task<dsPaymentBufferMassInsertResponse> UploadPaymentsToBufferAsync(TPaymentListTypeForDSPaymentBufferMassInsert[] massInsertPayments)
        {
            var client = GetServiceClient();
            return await client.dsPaymentBufferMassInsertAsync(massInsertPayments);
        }

        /// <summary>
        /// Обработка ошибок загрузки платёжек в АБС, в реестр неразнесённых платежей (Бэк-офис/Кредиты/Обработка платежей)
        /// </summary>
        static void ProcessUploadPaymentErrors(IReadOnlyList<PaymentOrder> paymentsFromExcel, DsPaymentBufferMassInsertRes uploadResult)
        {
            var dsPaymentBufferMassInsertRes = uploadResult;
            if (dsPaymentBufferMassInsertRes.Status != "OK")
            {
                var errors = dsPaymentBufferMassInsertRes.NotificationList.Select(
                    x => $"Ошибка: {x.NTFMessage}\n Платёж:\n {JsonConvert.SerializeObject(paymentsFromExcel[Convert.ToInt32(x.LinkID)], Formatting.Indented)}");
                var errorMessage = dsPaymentBufferMassInsertRes.ReturnMsg + "\n\n " + string.Join(";\n", errors);
               // Console.WriteLine(errorMessage);
            }
            else
            {
              //  Console.WriteLine("Загрузка успешно завершена");
            }
        }

        /// <summary> Получить список пар: (Id загруженной платёжки, Id кредитного договора) на основе списка платежей, для которых заполнено поле НомерКредитногоДоговора </summary>
        static async Task<IEnumerable<(long PaymentId, long CreditContractId)>> GetPaymentsToBind(
            PaymentOrder[] paymentFromExcel,
            TPaymentListTypeForDSPaymentBufferMassInsert[] massInsertPayments,
            DsPaymentBufferMassInsertRes uploadResult)
        {
            var errorLinks = uploadResult?.NotificationList?.ToDictionary(x => x.LinkID, x => x.LinkID) ?? new Dictionary<long, long>();
            var paymentsToBind = paymentFromExcel.Select((p, i) => (p.CreditContractNumber, massInsertPayments[i].GUID, massInsertPayments[i].LinkID))
                                            .Where(x => !string.IsNullOrEmpty(x.CreditContractNumber))
                                            .Where(x => !errorLinks.ContainsKey(x.LinkID))
                                            .Select(x => (PaymentGuid: x.GUID, x.CreditContractNumber))
                                            .ToList();

            if (!paymentsToBind.Any())
            {
               // Console.WriteLine("Список платёжек с указанным номером договора пуст.");
                return new (long, long)[0];
            }

           // Console.WriteLine("Список кандидатов на автоматическую привязку (GUID Платёжки, НомерДоговора):");
         //   paymentsToBind.ForEach(x => Console.WriteLine(x.ToString()));

            var paymentIds = await GetPaymentIdsAsync(paymentsToBind.Select(x => x.PaymentGuid));
            var creditContractIds = await GetCreditContractIdsAsync(paymentsToBind.Select(x => x.CreditContractNumber));

            var paymentContractPairs = paymentsToBind.Where(x => paymentIds.ContainsKey(x.PaymentGuid) && creditContractIds.ContainsKey(x.CreditContractNumber))
                                           .Select(x => (PaymentId: paymentIds[x.PaymentGuid], CreditContractId: creditContractIds[x.CreditContractNumber]))
                                           .ToList();

          //  Console.WriteLine("Оставшиеся кандидаты на автоматическую привязку (Id Платёжки, Id Договора):");
        //    paymentContractPairs.ForEach(x => Console.WriteLine(x.ToString()));

            return paymentContractPairs;
        }

        /// <summary>
        /// Получить словарь { GUID платёжки, ID платёжки } по списку GUID-ов платёжек из реестра необработанных платежей (АБС: Бэк-офис/Кредиты/Обработка платежей)
        /// </summary>
        static async Task<Dictionary<long, long>> GetPaymentIdsAsync(IEnumerable<long> paymentGuids)
        {
            var client = GetServiceClient();

            var requests = paymentGuids.Distinct().Select(async x => await client.dsPaymentBufferFindListByParamAsync(new DsPaymentBufferFindListByParamReq { GUID = x, GUIDSpecified = true }));
            var responses = await Task.WhenAll(requests);

            var paymentsMap = new Dictionary<long, long>();
            foreach (var response in responses)
            {
                var responseResult = response.DsPaymentBufferFindListByParamRes;
                if (responseResult.Status != "OK")
                {
                    throw new Exception("Ошибка запроса списка неразнесённых платежей.", new Exception(responseResult.ReturnMsg));
                }
                var paymentByGuid = responseResult.PaymentList?.SingleOrDefault();
                if (paymentByGuid != null)
                {
                    paymentsMap.Add(paymentByGuid.GUID, paymentByGuid.PaymentID);
                }
            }

            return paymentsMap;
        }
        /// <summary>
        /// Получить словарь { Номер кредитного договора, ID кредитного договора } по номерам кредитных договоров
        /// </summary>
        static async Task<Dictionary<string, long>> GetCreditContractIdsAsync(IEnumerable<string> creditContractNumbers)
        {
            var client = GetServiceClient();

            var requests = creditContractNumbers.Distinct().Select(async x => await client.dsLoanBrowseListByParamAsync(new DsLoanBrowseListByParamReq { Number = x }));
            var responses = await Task.WhenAll(requests);

            var contractsMap = new Dictionary<string, long>();
            foreach (var response in responses)
            {
                var responseResult = response.DsLoanBrowseListByParamRes;
                if (responseResult.Status != "OK")
                {
                    throw new Exception("Ошибка запроса списка кредитных договоров.", new Exception(responseResult.ReturnMsg));
                }
                var contractByNumber = responseResult.Result?.SingleOrDefault();
                if (contractByNumber != null)
                {
                    contractsMap.Add(contractByNumber.Number, contractByNumber.LoanID);
                }
            }

            return contractsMap;
        }

        /// <summary>
        /// Привязать Платёжное поручение к Кредитному договору
        /// </summary>
        static async Task BindPaymentsToCreditContractsAsync(IEnumerable<(long PaymentId, long CreditContractId)> paymentToBind)
        {
            var client = GetServiceClient();

            var requests = paymentToBind.Distinct().Select(async x => await client.dsPaymentBufferExecuteOperationAsync(
                                                               new DsPaymentBufferExecuteOperationReq
                                                               {
                                                                   LoanID = x.CreditContractId,
                                                                   PaymentID = x.PaymentId,
                                                                   // Признак автоматического сторнирования и выполнения операций, выполненных по договору после даты платежа:
                                                                   // 1 - не выполнять (по умолчанию)
                                                                   // 2 - выполнять.
                                                                   AutomaticReversalFlag = 2,
                                                                   AutomaticReversalFlagSpecified = true
                                                               }));
            var responses = await Task.WhenAll(requests);
            foreach (var result in responses.Select(x => x.DsPaymentBufferExecuteOperationRes))
            {
                var message = $"{(result.ReturnCodeSpecified ? "Ошибка" : "Успех")} выполнения операции привязки Платёжного поручения к Кредитному договору.\n" +
                              $"{(result.ReturnCodeSpecified ? result.ReturnMsg : JsonConvert.SerializeObject(result))}\n";

               // Console.ForegroundColor = result.ReturnCodeSpecified ? ConsoleColor.Red : ConsoleColor.White;
             //   Console.WriteLine(message + "\n======\n\n");
              //  Console.ResetColor();
            }
        }

        static LOANCREDITWSPORTTYPEClient GetServiceClient()
        {
            var client = new LOANCREDITWSPORTTYPEClient();
            client.Endpoint.Address = new EndpointAddress("http://cbs.dev3.mmk.local:9081/ftbpmadt/ftbpmadt");
            //client.Endpoint.Address = new EndpointAddress("http://cbs.hotfix.mmk.local:9081/ftbpmadt/ftbpmadt");
            client.ClientCredentials.UserName.UserName = "ESB_ADMIN";
            client.ClientCredentials.UserName.Password = "!Qwerty1";
            return client;
        }
    }
}
