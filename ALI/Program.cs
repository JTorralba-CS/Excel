using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace ALI
{
    class Excel
    {
        _Application _MSExcel = new MSExcel.Application();

        Workbook WB = null;
        Worksheet WS = null;
                
        public void CreateFile()
        {
            this.WB = _MSExcel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.WS = this.WB.Worksheets[1];

            WS.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            WS.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
        }

        public void CreateSheet(String Input)
        {
            String Raw = File.ReadAllText(Input);
            
            Worksheet Sheet = this.WS;

            Sheet.Name = Input;

            Sheet.Rows.RowHeight = 30;
            Sheet.StandardWidth = 5;
            Sheet.Cells.Font.Size = 12;

            int R_Start = 2;
            int C_Start = 2;

            int R = R_Start - 1;
            int C = C_Start;

            for (int I = 0; I < 50; I++)
            {
                Sheet.Cells[R, C].Font.Bold = true;
                Sheet.Cells[R, C] = (I + 1).ToString();
                C += 1;
            }

            R = R_Start;
            C = C_Start - 1;

            for (int I = 0; I <= 25; I++)
            {
                Sheet.Cells[R, C].Font.Bold = true;
                Sheet.Cells[R, C] = I.ToString();
                R += 1;
            }

            R = R_Start;
            C = C_Start;

            for (int I = 0; I < Raw.Length - 1; I++)
            {
                String Current = Raw.Substring(I, 1);
                String Next = Raw.Substring(I + 1, 1);

                String Data = "";

                Sheet.Cells[R, C].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                Sheet.Cells[R, C].Font.Bold = true;

                switch (Current)
                {
                    case "\r":
                        Data = "CR";
                        break;

                    case "\n":
                        Data = "LN";
                        break;

                    case "\u0002":
                        Data = "STX";
                        break;

                    case "\u0003":
                        Data = "ETX";
                        break;

                    default:
                        Data = Current;
                        Sheet.Cells[R, C].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                        Sheet.Cells[R, C].Font.Bold = false;
                        break;
                }

                //Console.WriteLine($"[{R},{C}] = {I} {Data}");
                Console.WriteLine("{0,-7} {1,-11} {2,-7}", I, $"[{R},{C}]", Data);

                Sheet.Cells[R, C] = Data;

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

        public void SaveAs(String Output)
        {

            FileInfo F = new FileInfo(Output);
            F.Delete();

            try
            {
                WB.SaveAs(Output);
                WB.Close(false);
                _MSExcel.Quit();
            }
            catch (Exception E)
            {
            }
            finally
            {
                Marshal.ReleaseComObject(WS);
                Marshal.ReleaseComObject(WB);
                Marshal.ReleaseComObject(_MSExcel);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static void Main(String[] Args)
        {
            String Input = @Args[0];
            String Output = @Path.Combine(Environment.CurrentDirectory, Input + ".xlsx");

            Console.WriteLine("Output File = " + Output);
            Console.WriteLine();

            Excel _Excel = new Excel();

            _Excel.CreateFile();
            _Excel.CreateSheet(Input);
            _Excel.SaveAs(Output);
        }
    }
}
