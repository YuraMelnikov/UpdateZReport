using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using NLog;

namespace UpdateZConsol
{
    class Program
    {
        static void Main(string[] args)
        {
            string tabel_01USR = "M:\\Производственный отдел\\Отчеты мастеров\\!Новая структура\\01_УСР\\Табель_УСР.xlsm";
            string tabel_02USMK = "M:\\Производственный отдел\\Отчеты мастеров\\!Новая структура\\02_Участок сборки модулей\\Табель_УСМК.xlsx";
            string tabel_04UIS = "M:\\Производственный отдел\\Отчеты мастеров\\!Новая структура\\04_Шинный участок\\Табель_УИШ.xlsm";
            string tabel_03USS = "M:\\Производственный отдел\\Отчеты мастеров\\!Новая структура\\03_Участок сборки шкафов\\Табель_УСШ.xlsm";
            string tabel_05EMU = "M:\\Производственный отдел\\Отчеты мастеров\\!Новая структура\\05_Электромонтажный участок (2)\\табель_ЭМУ.xlsm";
            Logger logger = LogManager.GetCurrentClassLogger();
            Excel.Application ObjWorkExcel = new Excel.Application();


            DateTime dateTime = DateTime.Now.AddDays(-7);
            string textConsole = "";

            for (int i = 0; i < 15; i++)
            {
                DateTime date = dateTime.AddDays(i);
                ReadExcel readExcel = new ReadExcel(date);


                textConsole = DateTime.Now.ToString() + " | " + "Зафиксирована дата загрузки данных: " + readExcel.DateUploadData.ToString().Substring(0, 10) + Environment.NewLine;
                Console.WriteLine(textConsole);

                //01_УСР
                UploadDump uploadDump = new UploadDump(date, "УСР");
                uploadDump = readExcel.GetUploadDump01_USR(uploadDump, tabel_01USR);
                readExcel.AddUploadDumpList(uploadDump);
                textConsole = DateTime.Now.ToString() + " | " + "Успешно считаны данные по 01_УСР: " + Environment.NewLine;
                Console.WriteLine(textConsole);

                //02_УСМК
                UploadDump uploadDump1 = new UploadDump(date, "УСМК");
                uploadDump = readExcel.GetUploadDump01_USR(uploadDump1, tabel_02USMK);
                readExcel.AddUploadDumpList(uploadDump1);
                textConsole = DateTime.Now.ToString() + " | " + "Успешно считаны данные по 02_УСМК: " + Environment.NewLine;
                Console.WriteLine(textConsole);

                //03_УСШ
                UploadDump uploadDump2 = new UploadDump(date, "УСШ");
                uploadDump = readExcel.GetUploadDump01_USR(uploadDump2, tabel_03USS);
                readExcel.AddUploadDumpList(uploadDump2);
                textConsole = DateTime.Now.ToString() + " | " + "Успешно считаны данные по 03_УСШ: " + Environment.NewLine;
                Console.WriteLine(textConsole);

                //04_УИШ
                UploadDump uploadDump3 = new UploadDump(date, "УИШ");
                uploadDump = readExcel.GetUploadDump01_USR(uploadDump3, tabel_04UIS);
                readExcel.AddUploadDumpList(uploadDump3);
                textConsole = DateTime.Now.ToString() + " | " + "Успешно считаны данные по 04_УИШ: " + Environment.NewLine;
                Console.WriteLine(textConsole);

                //05_ЭМУ
                UploadDump uploadDump4 = new UploadDump(date, "ЭМУ");
                uploadDump = readExcel.GetUploadDump01_USR(uploadDump4, tabel_05EMU);
                readExcel.AddUploadDumpList(uploadDump4);
                textConsole = DateTime.Now.ToString() + " | " + "Успешно считаны данные по 05_ЭМУ: " + Environment.NewLine;
                Console.WriteLine(textConsole);
                readExcel.WriteUploadDataInReport();
                textConsole = DateTime.Now.ToString() + " | Данные записаны в отчет успешно! " + Environment.NewLine;
                Console.WriteLine(textConsole);
                logger.Debug(" class Form1 : Form - buttonStart_Click ");
            }

        }
    }
}
