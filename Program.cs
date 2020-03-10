using Dadata;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SqlServer.Types;
using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Globalization;
using Dadata.Model;
using System.Data.SqlClient;

namespace ParseAddress
{
    class Program
    {
        
        static void Main(string[] args)  
        {
            ArrayList data;
            if (args.Length == 0)
                data = Parsexcel();
            else
                data = Parsexcel(args[0]);
            Console.WriteLine("Количество записей без шапки: {0}", data.Count - 1);
            data = Fulldata(data);
            Console.WriteLine("Количество распарсеных записей: {0}", data.Count);
            InsertTable(data);
            System.Threading.Thread.Sleep(2000);

        }

        public static void InsertTable(ArrayList data)
        {
            string connectionString = @"Data Source=HP15-CE073UR\SQLEXPRESS;Initial Catalog=Testparse;Integrated Security=True;User ID=admin;Password=123";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            foreach (ArrayList datarow in data)
            {
                String CommandText = String.Format("INSERT INTO Testparse.dbo.addrgeo(full_address, number_email, fias, lat, lon, geo) VALUES('{0}', {1}, '{2}', '{3}', '{4}', '{5}')", 
                    datarow[0], datarow[1], datarow[2], datarow[3], datarow[4], datarow[5]);
                SqlCommand comm = new SqlCommand(CommandText, connection);
                comm.ExecuteNonQuery();
            }
            Console.WriteLine("Вставка в БД закончена без ошибок!");
            connection.Close();
        }


        public static ArrayList Fulldata(ArrayList data, bool light=false)
        {
            if (data.Count == 0)
            {
                Console.WriteLine("Список для отправки в Dadata пуст!");
                System.Environment.Exit(0);
            }
            var client = new SuggestClient("b717e3a82d3e964d4e2f37ffe9777d9cad101217");
            ArrayList clearsheet = new ArrayList();
            ArrayList errorsheet = new ArrayList();
            foreach (ArrayList datarow in data)
            {   
                try
                {
                    var response = client.SuggestAddress(datarow[0].ToString(), 1);
                    if (response.suggestions[0].data.house_fias_id != null)
                        datarow.Add(response.suggestions[0].data.house_fias_id);
                    else if (response.suggestions[0].data.street_fias_id != null)
                        datarow.Add(response.suggestions[0].data.street_fias_id);
                    else if(response.suggestions[0].data.settlement_fias_id != null)
                        datarow.Add(response.suggestions[0].data.settlement_fias_id);
                    else
                        datarow.Add(response.suggestions[0].data.city_fias_id);

                    double lat;
                    double lon;
                    NumberFormatInfo nfi = new NumberFormatInfo();
                    nfi.NumberDecimalSeparator = ".";
                    lat = Convert.ToDouble(response.suggestions[0].data.geo_lat, nfi);
                    lon = Convert.ToDouble(response.suggestions[0].data.geo_lon, nfi);
                    datarow.Add(lat);
                    datarow.Add(lon);
                    var routeBuilder = SqlGeography.Point(lat, lon, 4326);
                    datarow.Add(routeBuilder);
                    clearsheet.Add(datarow);
                }
                catch
                {
                    Console.WriteLine("Ошибка с: {0}", datarow[0]);
                    errorsheet.Add(datarow);
                }
            }
            if (light == true)
            {
                return clearsheet;
            }
            if(errorsheet.Count > 1)
            {
                Console.WriteLine("Записи приведённый выше не распарсились, возможно в них есть ошибка или недопустимые знаки.");
                Console.WriteLine("Отправить на стандартизацию, каждая стандартизация стоит 10 копеек: (Ввести y-Y)");
                Console.WriteLine("\n Или хотите исправить ошибки в исходном документе вручную: (Ввести n-N или другой символ)");
                
                string input = Console.ReadLine();
                if (input == "y" ^ input == "Y")
                {
                    clearsheet.AddRange(Fulldata(DaDatacon(errorsheet)));
                }
                else
                {
                    Console.ReadLine();
                    System.Environment.Exit(0);
                }
            }
            return clearsheet;
        }

        public static ArrayList DaDatacon(ArrayList data)
        {
            var client = new CleanClient("b717e3a82d3e964d4e2f37ffe9777d9cad101217", "336ae4b9a1e5cd37b02a9d0c1538392561623662");
            Console.WriteLine("    Старый адрес               ||       Адрес после стандартизации   ");
            data.RemoveAt(0);
            foreach (ArrayList datarow in data)
            {
                try
                {
                    
                    var resp = client.Clean<Address>(datarow[0].ToString());
                    Console.WriteLine("{0} ||  {1}", datarow[0].ToString(), resp.result);
                    datarow.RemoveAt(0);
                    datarow.Insert(0, resp.result);
                }
                catch
                {
                    Console.WriteLine("Ошибка с: {0}", datarow[0]);
                    Console.WriteLine("Стоит поправить исходный файл");
                    Console.ReadLine();
                    System.Environment.Exit(0);
                }
                
            }
            return data;            
        }

        public static ArrayList Parsexcel(string fileName = @"D:\C# project\example.xlsx")
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    var rows = sheet.Descendants<Row>();

                    ArrayList myAL = new ArrayList();

                    // Or... via each row
                    foreach (Row row in rows)
                    {
                        ArrayList mydata = new ArrayList();
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                mydata.Add(str);
                            }
                            else if (c.CellValue != null)
                            {
                                mydata.Add(c.CellValue.Text);
                            }
                                                    }
                        myAL.Add(mydata);
                    }
                    return myAL;
                }
            }

        }

    }
}
