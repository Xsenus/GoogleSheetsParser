using GSP.Core.Controller;
using GSP.Core.Model;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Threading;

namespace GSP.ConsoleApp
{
    class Program
    {
        private static List<Order> ListOrder = new List<Order>();

        static void Main(string[] args)
        {
            var nameAccessDataBase = "GS.mdb";
            var connectionAccess = new ConnectionAccess(nameAccessDataBase);

            var selectAccess = $"SELECT Order.[{nameof(Order.Date)}], Order.[{nameof(Order.RequestNumber)}], " +
                $"Order.[{nameof(Order.Number)}], Order.[{nameof(Order.Name)}], Order.[{nameof(Order.Count)}] FROM [{nameof(Order)}];";

            using (var cmd = new OleDbCommand(selectAccess, connectionAccess.OleDbConnection))
            {
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var date = reader[0] is null ? default(DateTime?) : Convert.ToDateTime(reader[0]);
                        var requestNumber = reader[1] is null ? default(int?) : Convert.ToInt32(reader[1]);
                        var number = reader[2].ToString();
                        var name = reader[3].ToString();
                        var count = reader[4] is null ? default(int?) : Convert.ToInt32(reader[4]);

                        var order = new Order()
                        {
                            Date = date,
                            RequestNumber = requestNumber,
                            Number = number,
                            Name = name,
                            Count = count
                        };
                        ListOrder.Add(order);
                    }
                }
            }

            var jsonFile = "client_id.json";
            var googleSheetsAPI = new GoogleSheetsAPI(jsonFile);

            var spreadsheetId = "1iBrC-3ViNLCtv4fIUBDuZnfzaIoF2sa5G_dxE9sm6Zc";
            var range = "A3:Q";

            while (true)
            {
                var values = googleSheetsAPI.GetValue(spreadsheetId, range);

                if (values != null && values.Count > 0)
                {
                    var tempDate = default(DateTime?);
                    var tempRequestNumber = default(int?);

                    foreach (var row in values)
                    {
                        var date = default(DateTime?);
                        var requestNumber = default(int?);
                        var number = default(string);
                        var name = default(string);
                        var count = default(int?);

                        if (row.Count > 1)
                        {
                            date = string.IsNullOrWhiteSpace(row[0].ToString()) ? tempDate : Convert.ToDateTime(row[0]);
                            tempDate = date;
                        }

                        if (row.Count > 3)
                        {
                            requestNumber = string.IsNullOrWhiteSpace(row[2].ToString()) ? tempRequestNumber : Convert.ToInt32(row[2]);
                            tempRequestNumber = requestNumber;
                        }

                        if (row.Count > 4)
                        {
                            number = row[4] is null ? default : row[4].ToString();
                        }

                        if (row.Count > 5)
                        {
                            name = row[5] is null ? default : row[5].ToString();
                        }

                        if (row.Count > 6)
                        {
                            count = row[6] is null ? default(int?) : Convert.ToInt32(row[6]);
                        }

                        if (date == null && requestNumber == null && string.IsNullOrWhiteSpace(number) && string.IsNullOrWhiteSpace(name) && count == null)
                        {
                            tempDate = default;
                            tempRequestNumber = default;
                            continue;
                        }

                        var order = new Order()
                        {
                            Date = tempDate,
                            RequestNumber = tempRequestNumber,
                            Number = number,
                            Name = name,
                            Count = count
                        };

                        if (ListOrder.Where(w => w.Date == order.Date && 
                                            w.RequestNumber == order.RequestNumber && 
                                            w.Number == order.Number && 
                                            w.Name == order.Name && 
                                            w.Count == order.Count).FirstOrDefault() == null)
                        {
                            ListOrder.Add(order);                            

                            var quersy = $"INSERT INTO [{nameof(Order)}] " +
                                $"([{nameof(Order.Date)}], [{nameof(Order.RequestNumber)}], [{nameof(Order.Number)}], [{nameof(Order.Name)}], [{nameof(Order.Count)}])" +
                                $" VALUES " +
                                $"('{order.Date}', {order.RequestNumber}, '{order.Number}', '{order.Name}', {order.Count});";
                            connectionAccess.Injection(quersy);

                            Console.WriteLine(order);
                        }
                    }
                }

                Thread.Sleep(10000);
            }
        }
    }
}
