using Microsoft.Office.Interop.Excel;
using _excel = Microsoft.Office.Interop.Excel;

namespace ALI
{
    class Excel
    {
        _Application ExcelApp = new _excel.Application();
        Workbook WB;
        Worksheet WS;
                
        public void CreateFile()
        {
            this.WB = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.WS = this.WB.Worksheets[1];

            WS.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            WS.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
        }

        public void CreateSheet(String Input)
        {
            String Raw = File.ReadAllText(@Input);

            Worksheet Sheet1 = ExcelApp.Worksheets.Add(After: this.WS);

            Sheet1.Rows.RowHeight = 25;
            Sheet1.StandardWidth = 5;
            Sheet1.Cells.Font.Size = 12;

            int R_Start = 2;
            int C_Start = 2;

            int R = R_Start - 1;
            int C = C_Start;

            for (int i = 0; i < 40; i++)
            {
                Sheet1.Cells[R, C].Font.Bold = true;
                Sheet1.Cells[R, C] = (i + 1).ToString();
                C += 1;
            }

            R = R_Start;
            C = C_Start - 1;

            for (int i = 0; i <= 20; i++)
            {
                Sheet1.Cells[R, C].Font.Bold = true;
                Sheet1.Cells[R, C] = i.ToString();
                R += 1;
            }

            R = R_Start;
            C = C_Start;

            for (int i = 0; i < Raw.Length - 1; i++)
            {
                String Current = Raw.Substring(i, 1);
                String Next = Raw.Substring(i + 1, 1);

                String Output = "";

                Sheet1.Cells[R, C].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                Sheet1.Cells[R, C].Font.Bold = true;

                switch (Current)
                {
                    case "\r":
                        Output = "CR";
                        break;

                    case "\n":
                        Output = "LN";
                        break;

                    case "\u0002":
                        Output = "STX";
                        break;

                    case "\u0003":
                        Output = "ETX";
                        break;

                    default:
                        Output = Current;
                        Sheet1.Cells[R, C].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                        Sheet1.Cells[R, C].Font.Bold = false;
                        break;
                }

                Console.WriteLine($"[{R},{C}] = {i} {Output} {Current}");
                Sheet1.Cells[R, C] = Output;

                if (Current == "\r" & Next != "\n")
                {
                    R += 1;
                    C = C_Start;
                }
                else if (Current == "\n")
                {
                    R += 1;
                    C = C_Start;
                }
                else
                {
                    C += 1;
                }
            }
        }

        public void SaveAs(string filepath)
        {
            WB.SaveAs(filepath);
        }

        public static void Main(string[] args)
        {
            String Input = args[0];
            String Output = "C:\\" + Input + ".xlsx";

            Console.WriteLine("Output File = " + Output);

            Excel ExcelFile = new Excel();
            ExcelFile.CreateFile();
            ExcelFile.CreateSheet(Input);
            ExcelFile.SaveAs(@Output);
            ExcelFile.WB.Close();
        }
    }
}
