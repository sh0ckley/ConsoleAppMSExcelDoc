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
    class Program
    {
        static void Main(string[] args)
        {
            Console.WindowHeight = Console.LargestWindowHeight;
            Console.WindowWidth = Console.LargestWindowWidth;

            Clients clients = new Clients();

            string input = "";

            Console.WriteLine("Введите путь до файла с данными: ");
            string filepath = Console.ReadLine();

            Console.WriteLine("\nДобро пожаловать!!!");
            Console.WriteLine("\nq - Указать путь до файла с данными");
            Console.WriteLine("w - Информация о клиентах, заказавших товар");
            Console.WriteLine("e - Изменить контактное лицо клиента");
            Console.WriteLine("r - Определить золотого клиента");
            Console.WriteLine("t - очистить консоль");
            Console.WriteLine("y - выход из программы");

            while (input != "y")
            {
                Console.Write("\n> ");

                input = Console.ReadLine();

                switch (input)
                {
                    case "q":
                        Console.WriteLine("Введите путь до файла с данными: ");
                        filepath = Console.ReadLine();
                        Console.WriteLine("Путь успешно сохранен.");
                        break;
                    case "w":
                        clients.ClientsInfo();
                        break;
                    case "e":
                        clients.ChangeClient(filepath);
                        break;
                    case "r":
                        clients.GoldClient();
                        break;
                    case "t":
                        Console.Clear();
                        break;
                    default:
                        Console.WriteLine("\nq - Указать путь до файла с данными");
                        Console.WriteLine("w - Информация о клиентах, заказавших товар");
                        Console.WriteLine("e - Изменить контактное лицо клиента");
                        Console.WriteLine("r - Определить золотого клиента");
                        Console.WriteLine("t - очистить консоль");
                        Console.WriteLine("y - выход из программы");
                        break;
                }
            }
        }
    }
}
