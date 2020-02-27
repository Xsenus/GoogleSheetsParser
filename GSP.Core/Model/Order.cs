using System;

namespace GSP.Core.Model
{
    /// <summary>
    /// Заказ.
    /// </summary>
    public class Order
    {
        /// <summary>
        /// Дата заказа.
        /// </summary>
        public DateTime? Date { get; set; }

        /// <summary>
        /// Номер запроса.
        /// </summary>
        public int? RequestNumber { get; set; }

        /// <summary>
        /// Партномер.
        /// </summary>
        public string Number { get; set; }

        /// <summary>
        /// Наименование.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Количество.
        /// </summary>
        public int? Count { get; set; }

        public Manager Manager { get; set; }

        public override string ToString()
        {
            var dateTime = Date is null ? "Нет даты" : Convert.ToDateTime(Date).ToShortDateString();

            return $"{dateTime} [{RequestNumber}]: {Number} {Name} -> {Count}";
        }
    }
}
