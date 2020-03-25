using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace UploadClient.Models.Convertion
{
    public sealed class ConvertToExcel : IConvertToExcel
    {

        public byte[] Convert(IFormFile formFile)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var inputFileName = formFile.Name;
            var inputFile = ReadString(formFile);

            // 2. Распарсили 1С-файл, получили список платёжек. Отфильтровали список, отсортировали: платёжки, требующие привязки, - в начале списка
            var paymentsFrom1C = LoadFrom1C(inputFile)
                                .Where(x => x["ДатаПоступило"] != null)           // Исходящие платежи
                                .Where(x => x["ПлательщикИНН"]?.ToString() != x["ПолучательИНН"]?.ToString()) // Внутренние переводы ООО МФК "Быстроденьги"
                                .OrderBy(x => x["Плательщик1"] != null && x["Плательщик1"].ToString().StartsWith("УФК") ? 1 : 2) // Взыскания по просуженным займам через Пенсионный фонд - в первую очередь
                                .ThenBy(x => x["Плательщик1"] != null && x["Плательщик1"].ToString().StartsWith("ПАО") ? 1 : 2) // Самостоятельная оплата от заёмщика - во вторую очередь 
                                .ThenBy(x => x["НазначениеПлатежа"] != null && x["НазначениеПлатежа"].ToString().ToLower().Contains("инкас") ? 2 : 1); // Платежи от инкассации - в последнюю очередь 

            // 3. Экспортировали список в Excel для последующей ручной обработки - оператор должен заполнить поле НомерКредитногоДоговора
            var workbook = GetExcelWorkbook(paymentsFrom1C, "Документы");
            var response = SaveWorkbookToMemoryStream(workbook);
           // var excelFileName = Path.ChangeExtension(inputFileName, ".xlsx");
          //  workbook.SaveAs(response);

            return response;
        }

        /// <summary> Загрузка платежей через разбор 1с-формата выгрузки платёжек </summary>
        /// <param name="input1CDocument">выгрузка платёжек в формате 1CClientBankExchange версии 1.02 и выше</param>
        private static IEnumerable<JObject> LoadFrom1C(string input1CDocument)
        {
            var regex = new Regex(@"СекцияДокумент(?:(?:\s|\t)*=(?:\s|\t)*(?<doctype>(?:\w|\s)*)){0,1}\n(?:(?<kv>(.)*?)\n)*?КонецДокумента", RegexOptions.Singleline);
            var matches = regex.Matches(input1CDocument).OfType<Match>().ToList();

            foreach (var match in matches)
            {
                var payment = new JObject { { "@Тип", match.Groups["doctype"].Value.Trim() } };

                foreach (var s in match.Groups["kv"].Captures.OfType<Capture>().Select(x => x.Value))
                {
                    var kvPair = s.Split('=');
                    payment.Add(kvPair[0].Trim(), kvPair.Length == 2 ? kvPair[1].Trim() : null);
                }

                yield return payment;
            }
        }

        private static IXLWorkbook GetExcelWorkbook(IEnumerable<JObject> incomePayments, string worksheetName)
        {
            if (incomePayments == null) throw new ArgumentNullException(nameof(incomePayments));
            if (!incomePayments.Any()) throw new ArgumentOutOfRangeException(nameof(incomePayments));

            var jArray = new JArray(incomePayments);
            var tempTable = JsonConvert.DeserializeObject<DataTable>(jArray.ToString());

            var dataTable = new DataTable(worksheetName);
            dataTable.Columns.Add("НомерКредитногоДоговора");
            var dataColumns = new List<DataColumn>();
            foreach (DataColumn column in tempTable.Columns) dataColumns.Add(column);
            dataTable.Columns.AddRange(dataColumns
                                       .Select(x => new DataColumn(x.ColumnName, x.DataType))
                                       .OrderBy(x => x.ColumnName == "Номер" ? 1 : 2)
                                       .ThenBy(x => x.ColumnName == "ДатаПоступило" ? 1 : 2)
                                       .ThenBy(x => x.ColumnName == "Сумма" ? 1 : 2)
                                       .ThenBy(x => x.ColumnName == "НазначениеПлатежа" ? 1 : 2)
                                       .ThenBy(x => x.ColumnName == "Плательщик1" ? 1 : 2)
                                       .ThenBy(x => x.ColumnName == "@Тип" ? 2 : 1)
                                       .ToArray());

            foreach (DataRow row in tempTable.Rows)
            {
                var newRow = dataTable.NewRow();
                foreach (DataColumn column in tempTable.Columns)
                {
                    newRow[column.ColumnName] = row[column.ColumnName];
                }

                // Если плательщик - Управление Федеральным Казначейством, пытаемся выковырять 8-значный номер кредитного договора,
                // который начинается с "9" из поля Назначение платежа
                if (row["Плательщик1"].ToString().StartsWith("УФК"))
                {
                    var matchCollection = Regex.Matches(row["НазначениеПлатежа"].ToString(), @"\D+(?<ndoc>9\d{7,7})\D");
                    if (matchCollection.Count == 1)
                    {
                        newRow["НомерКредитногоДоговора"] = matchCollection[0].Groups["ndoc"];
                    }
                }

                dataTable.Rows.Add(newRow);
            }

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(dataTable);
            worksheet.Columns().AdjustToContents();
            return workbook;
        }

        private string ReadString(IFormFile file)
        {
            using var reader = new StreamReader(file.OpenReadStream(), Encoding.GetEncoding(1251));
            string result = reader.ReadToEnd();
            return result;
        }

        public static byte[] SaveWorkbookToMemoryStream(IXLWorkbook workbook)
        {
            using MemoryStream stream = new MemoryStream();

            workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = false, GenerateCalculationChain = false, ValidatePackage = false });
            var bytes = stream.ToArray();
           // var tmp = ByteArrayToObject(bytes);

            string path = @"C:\test1.xlsx";
            File.WriteAllBytes(path, bytes);

            return bytes;
        }

        //public static byte[] SaveWorkbookToMemoryStream(IXLWorkbook workbook)
        //{
        //    using MemoryStream stream = new MemoryStream();

        //    workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = false, GenerateCalculationChain = false, ValidatePackage = false });
        //   // var bytes = stream.ToArray();
        //    return stream;
        //}
        private static ExcelPackage ByteArrayToObject(byte[] arrBytes)
        {
            using (MemoryStream memStream = new MemoryStream(arrBytes))
            {
                ExcelPackage package = new ExcelPackage(memStream);
                return package;
            }
        }
    }
}
