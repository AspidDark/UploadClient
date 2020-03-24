using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace UploadClient.Models.AbsDTO
{
    /// <summary> Платёжное поручение </summary>
    class PaymentOrder
    {
        /// <summary> Номер кредитного договора для автоматической привязки Платёжного поручения </summary>
        public string CreditContractNumber { get; set; }

        /// <summary> Номер  </summary>
        public string Number { get; set; }

        /// <summary> Дата поступления </summary>
        public DateTime IncomeDate { get; set; }

        /// <summary> Сумма </summary>
        public decimal Amount { get; set; }

        /// <summary> Назначение платежа </summary>
        public string Description { get; set; }

        /// <summary> Плательщик </summary>
        public string PayerName { get; set; }
    }
}
