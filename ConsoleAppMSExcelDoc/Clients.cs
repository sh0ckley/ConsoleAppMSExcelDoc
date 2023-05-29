using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Text.RegularExpressions;

namespace ConsoleAppMSExcelDoc
{
    class Clients
    {
        private static string GetConnectionString()
        {
            return @"Data Source=DESKTOP-DN7NRNQ;Initial Catalog=dbStore;Integrated Security=True";
        }

        private string connectionString = GetConnectionString();

        public void ClientsInfo()
        {
            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                try
                {
                    sqlConnection.Open();

                    using (SqlCommand sqlCommand = new SqlCommand("SELECT * FROM Clients", sqlConnection))
                    {
                        using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
                        {
                            Console.WriteLine();
                            while (sqlDataReader.Read())
                            {
                                int id = sqlDataReader.GetInt32(0);
                                string nameOrg = sqlDataReader.GetString(1);
                                string address = sqlDataReader.GetString(2);
                                string surname = sqlDataReader.GetString(3);
                                string firstname = sqlDataReader.GetString(4);
                                string patronymic = sqlDataReader.GetString(5);

                                Console.WriteLine(id + "\t" + nameOrg + "\t" + address + "\t" + surname + "\t" + firstname + "\t" + patronymic);
                            }
                        }
                    }

                    using (SqlCommand sqlCommand = new SqlCommand("SELECT Product_Name, Price FROM Product", sqlConnection))
                    {
                        using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
                        {
                            Console.WriteLine();
                            while (sqlDataReader.Read())
                            {
                                string name = sqlDataReader.GetString(0);
                                int price = sqlDataReader.GetInt32(1);

                                Console.Write(name +  " " + price + "\n");
                            }
                        }
                    }


                    using (SqlCommand sqlCommand = new SqlCommand("SELECT Required_Quantity, Date_Placement  FROM Applications", sqlConnection))
                    {
                        using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
                        {
                            Console.WriteLine();
                            while (sqlDataReader.Read())
                            {
                                int quantity = sqlDataReader.GetInt32(0);
                                string date = sqlDataReader.GetString(1);

                                Console.Write(quantity + "\t" + date + "\n");
                            }
                        }
                    }
                }
                catch (SqlException ex)
                {
                    Console.WriteLine(ex);
                }
            }
        }

        public void ChangeClient(string filepath)
        {
            try
            {
                Console.Write("\nВведите название организации: ");
                string nameOrg = Console.ReadLine();

                Console.Write("Введите фамилию клиента: ");
                string surname = Console.ReadLine();

                Console.Write("Введите имя клиента: ");
                string firstname = Console.ReadLine();

                Console.Write("Введите отчество клиента: ");
                string patronymic = Console.ReadLine();

                string connectionString = GetConnectionString();

                using (SqlConnection sqlConnection = new SqlConnection(connectionString))
                {
                    sqlConnection.Open();
                    using (SqlCommand sqlCommand = new SqlCommand(
                        "UPDATE Clients " +
                        $"SET Name_Org = '{nameOrg}', " +
                        $"Surname = '{surname}', " +
                        $"Firstname = '{firstname}', " +
                        $"Patronymic = '{patronymic}' " +
                        "WHERE ID_Client=287", sqlConnection))
                    {
                        int rowsAffected = sqlCommand.ExecuteNonQuery();
                        Console.WriteLine("Изменения внесены.", rowsAffected);
                    }
                }

                try
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                        FileVersion fv = new FileVersion();
                        fv.ApplicationName = "Microsoft Office Excel";
                        worksheetPart.Worksheet = new Worksheet(new SheetData());
                        WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();

                        //Создаем лист в книге
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Клиенты" };
                        sheets.Append(sheet);

                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        //Добавим заголовки в первую строку
                        Row row = new Row() { RowIndex = 1 };
                        sheetData.Append(row);

                        InsertCell(row, 1, "Код клиента", CellValues.String, 5);
                        InsertCell(row, 2, "Наименование организации", CellValues.String, 5);
                        InsertCell(row, 3, "Адрес", CellValues.String, 5);
                        InsertCell(row, 4, "Контактное лицо (ФИО)", CellValues.String, 5);

                        // Добавляем в строку все стили подряд.
                        row = new Row() { RowIndex = 2 };
                        sheetData.Append(row);

                        InsertCell(row, 1, "287", CellValues.Number, 1);
                        InsertCell(row, 2, ReplaceHexadecimalSymbols($"{nameOrg}"), CellValues.String, 2);
                        InsertCell(row, 2, ReplaceHexadecimalSymbols("пензенская область, город клин, пл. сталина, 74}"),
                            CellValues.String, 3);
                        InsertCell(row, 3, ReplaceHexadecimalSymbols($"{surname}" + " " + $"{firstname}" + " " + $"{patronymic}"),
                            CellValues.String, 4);

                        Row row2 = new Row() { RowIndex = 3 };
                        sheetData.Append(row2);

                        InsertCell(row2, 1, "820", CellValues.Number, 1);
                        InsertCell(row2, 2, ReplaceHexadecimalSymbols("ООО Звезда"), CellValues.String, 2);
                        InsertCell(row2, 2, ReplaceHexadecimalSymbols("Брянская область, город Красногорск, проезд Ленина, 53"),
                            CellValues.String, 3);
                        InsertCell(row2, 3, ReplaceHexadecimalSymbols("Журавлёв Давид Александрович"),
                            CellValues.String, 4);

                        Row row3 = new Row() { RowIndex = 4 };
                        sheetData.Append(row3);

                        InsertCell(row3, 1, "748", CellValues.Number, 1);
                        InsertCell(row3, 2, ReplaceHexadecimalSymbols("ООО День"), CellValues.String, 2);
                        InsertCell(row3, 2, ReplaceHexadecimalSymbols("Оренбургская область, город Воскресенск, проезд Чехова, 76"),
                            CellValues.String, 3);
                        InsertCell(row3, 3, ReplaceHexadecimalSymbols("Муравьёвa Жанна Львовна"),
                            CellValues.String, 4);

                        Row row4 = new Row() { RowIndex = 5 };
                        sheetData.Append(row4);

                        InsertCell(row4, 1, "633", CellValues.Number, 1);
                        InsertCell(row4, 2, ReplaceHexadecimalSymbols("ООО Снег"), CellValues.String, 2);
                        InsertCell(row4, 2, ReplaceHexadecimalSymbols("Ивановская область, город Орехово-Зуево, пер. Чехова, 59"),
                            CellValues.String, 3);
                        InsertCell(row4, 3, ReplaceHexadecimalSymbols("Андреев Кирилл Дмитриевич"), CellValues.String, 4);

                        workbookPart.Workbook.Save();
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("Error");
                }
            }
            catch (SqlException ex)
            {
                Console.WriteLine(ex);
            }
        }

        static void InsertCell(Row row, int cell_num, string val, CellValues type, uint styleIndex)
        {
            Cell refCell = null;
            Cell newCell = new Cell() { CellReference = cell_num.ToString() + ":" + row.RowIndex.ToString(), StyleIndex = styleIndex };
            row.InsertBefore(newCell, refCell);

            // Устанавливает тип значения.
            newCell.CellValue = new CellValue(val);
            newCell.DataType = new EnumValue<CellValues>(type);

        }

        //Метод убирает из строки запрещенные спец символы.
        static string ReplaceHexadecimalSymbols(string txt)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(txt, r, "", RegexOptions.Compiled);
        }

        public void GoldClient()
        {
            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                try
                {
                    sqlConnection.Open();

                    Console.WriteLine("\nВведите количество заказов: ");
                    string quantity = Console.ReadLine();

                    Console.WriteLine("\nВведите дату: ");
                    string date = Console.ReadLine();

                    using (SqlCommand sqlCommand = new SqlCommand("SELECT Surname, Firstname, Patronymic FROM Clients " +
                        "\nJOIN Applications ON Clients.ID_Client = Applications.ID_Client" +
                        $"\nWHERE Required_Quantity = {quantity} AND Date_Placement = '{date}'"
                        , sqlConnection))
                    {
                        using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
                        {


                            while (sqlDataReader.Read())
                            {
                                string surname = sqlDataReader.GetString(0);
                                string firstname = sqlDataReader.GetString(1);
                                string patronymic = sqlDataReader.GetString(2);
                                Console.WriteLine("\n" + surname + " " + firstname + " " + patronymic);
                            }
                        }
                    }
                }
                catch (SqlException ex)
                {
                    Console.WriteLine(ex);
                }
            }
        }
    }
}
