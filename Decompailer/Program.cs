using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

class E
{

    public static void Main(string[] arg)
    {
        int stratHour = 6;
        int endHour = 9;
        bool dateTimeActivation = false; // for all or timed data 

        Console.WriteLine("Location open file");
        string path = Console.ReadLine();
        //Console.WriteLine("Do you want to add START and END time (HOUR ONLY)--- y/n");
        //var mline = Console.ReadLine();
        //if (mline == "y" || mline == "Y")
        //{
        //    Console.WriteLine("Add Start HOUR");
        //    var stratHOUR = Console.ReadLine();
        //    Console.WriteLine("Add End HOUR");
        //    var EndHour = Console.ReadLine();
        //    int.TryParse(stratHOUR, out stratHour);
        //    int.TryParse(EndHour, out endHour);
        //}





        Excel.Application excel;
        Excel.Workbook workbook;
        Excel.Worksheet sheet;

        excel = new Excel.Application();
        excel.Visible = true;

        workbook = excel.Workbooks.Add();
        sheet = (Excel.Worksheet)workbook.ActiveSheet;

        //string[] lines = File.ReadAllLines(@"01.03.2021.csv");
        string[] lines = File.ReadAllLines(path);

        int index = 0;
        foreach (string line in lines)
        {
            // Use a tab to indent each line of the file.
            Console.WriteLine("\t" + line);
            string[] result = line.Split(',');

            if (result.Length > 0)
            {
                for (int i = 0; i < result.Length; i++)
                {
                    string miniVar = result[i];

                    if(i == 6)
                    {

                    }
                    else
                    {
                        miniVar = miniVar.TrimStart('"');
                        miniVar = miniVar.TrimEnd('"');
                        Console.WriteLine(miniVar);
                    }

                    result[i] = miniVar;
                }
            }

            int miniIndex = index + 1;

            if (result.Length == 1)
            {
                string getVal = result[0];
                sheet.Cells[index + 1, 1] = getVal;
                string ACol = "A" + miniIndex;
                string GCol = "G" + miniIndex;

                sheet.get_Range(ACol, GCol).Merge();
            }


            if (result.Length > 1)
            {
                bool print = false;
                for (int i = 0; i < result.Length; i++)
                {
                    string miniResult = result[i];

                    if (i == 0)
                    {
                        string[] trimedTime = miniResult.Split(':');
                        string hour = trimedTime[0];

                        int Hour = 0;

                        int.TryParse(hour, out Hour);

                        //int Hour = Int32.Parse(hour);
                        if (Hour >= stratHour && Hour < endHour)
                        {
                            //sheet.Cells[miniIndex, i + 1] = result[i];
                            print = true;
                        }
                        else
                        {
                            print = false;
                        }
                    }

                    if (print)
                    {
                        sheet.Cells[miniIndex, i + 1] = result[i];
                    }

                }
            }

            if (dateTimeActivation)
            {

            }
            else
            {
                index++;//making index
            }
           
        }

        string FileName = "NEW" + path;
        sheet.SaveAs2(FileName);
        workbook.Close();
        excel.Quit();

    }

}