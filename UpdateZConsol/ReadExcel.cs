using NLog;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace UpdateZConsol
{
    class ReadExcel
    {
        DateTime dateUploadData;
        List<UploadDump> uploadDumpList = new List<UploadDump>();
        string excelListName = "";

        public DateTime DateUploadData
        {
            get { return dateUploadData; }
            set
            {
                dateUploadData = new DateTime(value.Year, value.Month, value.Day);
            }
        }

        private static Logger logger = LogManager.GetCurrentClassLogger();

        public ReadExcel(DateTime dateUploadData)
        {
            this.dateUploadData = dateUploadData;
            this.excelListName = GetNameDate(dateUploadData);
        }

        public UploadDump GetUploadDump01_USR(UploadDump uploadDump, string path)
        {
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[GetNumListExcel(ObjWorkBook)]; //получить 1 лист

            int column = dateUploadData.Day + 3;
            int row = 13;
            if (uploadDump.devision == "ЭМУ")
                row = 22;
            string textData = "";
            string textCorrectData = "";

            //Get WorkData & Set worker
            do
            {
                try
                {
                    textCorrectData = ObjWorkSheet.Cells[row, 2].Text.ToString();
                }
                catch
                {
                    textCorrectData = "";
                }
                try
                {
                    textData = ObjWorkSheet.Cells[row, column].Text.ToString();
                }
                catch
                {
                    textData = "";
                }

                if (textData == "Б" || textData == "б")
                {
                    uploadDump.bol += 1;
                    uploadDump.worker += 1;
                }

                if (textData == "П" || textData == "п")
                {
                    uploadDump.prostoy += 1;
                    uploadDump.worker += 1;
                }
                if (textData == "О" || textData == "о" || textData == "o" || textData == "O")
                {
                    uploadDump.po += 1;
                    uploadDump.worker += 1;

                }
                if (textData == "C" || textData == "c" || textData == "С" || textData == "с" || textData == "В" || textData == "в")
                {
                    uploadDump.so += 1;
                    uploadDump.worker += 1;
                }
                try
                {
                    if (textData.Substring(0, 1) == "K" || textData.Substring(0, 1) == "k" || textData.Substring(0, 1) == "К" || textData.Substring(0, 1) == "к")
                    {
                        uploadDump.kom += 1;
                        uploadDump.worker += 1;
                    }
                }
                catch
                {

                }
                if (textData == "У" || textData == "у")
                {
                    uploadDump.uvo += 1;
                    uploadDump.worker += 1;
                }
                try
                {
                    if (Convert.ToInt32(textData) <= 8 & Convert.ToInt32(textData) > 0 & Convert.ToInt32(textCorrectData) > 0)
                    {
                        uploadDump.work8 += 1;
                        uploadDump.worker += 1;
                    }
                    if (Convert.ToInt32(textData) > 8)
                    {
                        uploadDump.work10 += 1;
                        uploadDump.worker += 1;
                    }
                }
                catch
                {

                }
                if (textData.ToUpper() == "УСР")
                    uploadDump.setUSR += 1;
                if (textData.ToUpper() == "УСМК")
                    uploadDump.setUSMK += 1;
                if (textData.ToUpper() == "УСШ")
                    uploadDump.setUSS += 1;
                if (textData.ToUpper() == "УИШ")
                    uploadDump.setUIS += 1;
                if (textData.ToUpper() == "ЭМУ")
                    uploadDump.setEMU += 1;
                row++;
            } while (textCorrectData != "Итого по бригадам");

            //Get Start position Get worker
            do
            {
                textCorrectData = ObjWorkSheet.Cells[row, 2].Text.ToString();
                textData = ObjWorkSheet.Cells[row, column].Text.ToString();
                row++;
            } while (textCorrectData != "Перемещения");
            row += 3;
            //getUSR
            do
            {
                textCorrectData = ObjWorkSheet.Cells[row, 2].Text.ToString();
                textData = ObjWorkSheet.Cells[row, column].Text.ToString();
                try
                {
                    if (Convert.ToInt32(textData) > 0)
                        uploadDump.getUSR += 1;
                }
                catch
                {

                }
                row++;
            } while (textCorrectData != "УСМК");

            //getUSMK
            do
            {
                textCorrectData = ObjWorkSheet.Cells[row, 2].Text.ToString();
                textData = ObjWorkSheet.Cells[row, column].Text.ToString();
                try
                {
                    if (Convert.ToInt32(textData) > 0)
                        uploadDump.getUSMK += 1;
                }
                catch
                {

                }
                row++;
            } while (textCorrectData != "УСШ");

            //getUSS
            do
            {
                textCorrectData = ObjWorkSheet.Cells[row, 2].Text.ToString();
                textData = ObjWorkSheet.Cells[row, column].Text.ToString();
                try
                {
                    if (Convert.ToInt32(textData) > 0)
                        uploadDump.getUSS += 1;
                }
                catch
                {

                }
                row++;
            } while (textCorrectData != "УИШ");

            //getUIS
            do
            {
                textCorrectData = ObjWorkSheet.Cells[row, 2].Text.ToString();
                textData = ObjWorkSheet.Cells[row, column].Text.ToString();
                try
                {
                    if (Convert.ToInt32(textData) > 0)
                        uploadDump.getUIS += 1;
                }
                catch
                {

                }
                row++;
            } while (textCorrectData != "ЭМУ");

            //getEMU
            do
            {
                textCorrectData = ObjWorkSheet.Cells[row, 2].Text.ToString();
                textData = ObjWorkSheet.Cells[row, column].Text.ToString();
                try
                {
                    if (Convert.ToInt32(textData) > 0)
                        uploadDump.getEMU += 1;
                }
                catch
                {

                }
                row++;
            } while (row < 300);

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой

            return uploadDump;
        }

        public void AddUploadDumpList(UploadDump uploadDump)
        {
            uploadDumpList.Add(uploadDump);
        }

        private string GetNameDate(DateTime date)
        {
            try
            {
                string data = "";
                data = date.Year.ToString() + ".";
                if (date.Month.ToString().Length == 1)
                    data += "0" + date.Month;
                else
                    data += date.Month;
                logger.Debug(" class ReadExcel - GetNameDate(DateTime date) ");
                return data;
            }
            catch (Exception ex)
            {
                logger.Error(" class ReadExcel - GetNameDate(DateTime date) " + ex.Message.ToString());
                return "error";
            }
        }

        private int GetNumListExcel(Excel.Workbook ObjWorkBook)
        {
            int size = ObjWorkBook.Sheets.Count;
            for (int i = 1; i <= size; i++)
            {
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[i];
                if (ObjWorkSheet.Name == excelListName)
                {
                    logger.Debug(" class ReadExcel - GetNumListExcel(Excel.Worksheet ObjWorkSheet, DateTime date) ");
                    return i;
                }
            }
            logger.Error(" class ReadExcel - GetNumListExcel(Excel.Worksheet ObjWorkSheet, DateTime date) " + "не найден такой лист: " + excelListName);
            return 0;
        }

        public void WriteUploadDataInReport()
        {
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open("M:\\Производственный отдел\\Z отчет\\отчёт по мощности ПО с 2018_07.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист

            DateTime dateCorrect = new DateTime(1901, 1, 1);
            int column = 1;
            int row = 1;
            do
            {
                string rangeTmp = "A" + row.ToString();
                string time = ObjWorkSheet.Range[rangeTmp].Text.ToString();
                try
                {
                    dateCorrect = Convert.ToDateTime(ObjWorkSheet.Cells[row, column].Text.ToString());
                    row++;
                }
                catch
                {
                    dateCorrect = new DateTime(1901, 1, 1);
                    row++;
                }

            } while (dateCorrect.ToString().Substring(0, 10) != dateUploadData.ToString().Substring(0, 10) || row > 10000);

            foreach (var data in uploadDumpList)
            {
                ObjWorkSheet.Cells[row, 2] = (data.worker + data.setUSR + data.setUSMK + data.setUIS + data.setUSS + data.setEMU).ToString();
                int work = data.worker - data.prostoy - data.po - data.so - data.kom - data.bol - data.uvo;
                ObjWorkSheet.Cells[row, 3] = work.ToString();
                ObjWorkSheet.Cells[row, 5] = data.prostoy.ToString();
                ObjWorkSheet.Cells[row, 6] = data.po.ToString();
                ObjWorkSheet.Cells[row, 7] = data.so.ToString();
                ObjWorkSheet.Cells[row, 8] = data.kom.ToString();
                ObjWorkSheet.Cells[row, 9] = data.bol.ToString();
                ObjWorkSheet.Cells[row, 10] = data.uvo.ToString();

                ObjWorkSheet.Cells[row, 11] = data.work8.ToString();
                ObjWorkSheet.Cells[row, 12] = data.work10.ToString();

                ObjWorkSheet.Cells[row, 13] = data.setUSR.ToString();
                ObjWorkSheet.Cells[row, 14] = data.setUSMK.ToString();
                ObjWorkSheet.Cells[row, 15] = data.setUIS.ToString();
                ObjWorkSheet.Cells[row, 16] = data.setUSS.ToString();
                ObjWorkSheet.Cells[row, 17] = data.setEMU.ToString();

                ObjWorkSheet.Cells[row, 18] = data.getUSR.ToString();
                ObjWorkSheet.Cells[row, 19] = data.getUSMK.ToString();
                ObjWorkSheet.Cells[row, 20] = data.getUIS.ToString();
                ObjWorkSheet.Cells[row, 21] = data.getUSS.ToString();
                ObjWorkSheet.Cells[row, 22] = data.getEMU.ToString();
                row++;
                //if (data.devision == "УСР")
                //    row += 2;
                //else
                //    row++;
            }
            string pathVersion = "M:\\Производственный отдел\\Z отчет\\" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" +
                DateTime.Now.Day.ToString();
            ObjWorkBook.Save();
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой
        }

        //public void ReadDataInTabelMasters()
        //{
        //    GetStartFinishDate();
        //    GetStartFinishDateList();

        //    Excel.Application ObjWorkExcel = new Excel.Application();
        //    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"C:\Users\Yura.Melnikau\Desktop\!Новая структура\01_УСР\" + @"новый табель_Сварочный участок.xlsm", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[GetNumListExcel(ObjWorkBook, dateStart)]; //получить 1 лист

        //    string devisionName = "УСР";

        //    foreach (var data in dateList)
        //    {
        //        uploadDumpList.Add(new UploadDump(data, devisionName));
        //    }



        //    //var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
        //    //string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
        //    //for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
        //    //    for (int j = 0; j < lastCell.Row; j++) // по всем строкам
        //    //        list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку


        //    ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
        //    ObjWorkExcel.Quit(); // выйти из экселя
        //    GC.Collect(); // убрать за собойqs
        //}

        //private void GetStartFinishDate()
        //{
        //    try
        //    {
        //        dateStart = DateTime.Now.AddDays(-7);
        //        dateFinish = DateTime.Now;
        //        logger.Debug(" class ReadExcel - GetStartFinishDate() ");
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error(" class ReadExcel - GetStartFinishDate() " + ex.Message.ToString());
        //    }

        //}

        //private void GetStartFinishDateList()
        //{
        //    try
        //    {
        //        for (int i = 0; i < 7; i++)
        //        {
        //            dateList.Add(dateStart.AddDays(i));
        //        }
        //        logger.Debug(" class ReadExcel - GetStartFinishDateList() ");
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error(" class ReadExcel - GetStartFinishDateList() " + ex.Message.ToString());
        //    }

        //}
    }
}
