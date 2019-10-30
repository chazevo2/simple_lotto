using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string str = "메뉴 || 1.예상번호 출력 || 2.당첨여부 확인 || 0.종료 : ";
            int menu = 1;
            int[] numList = new int[6] { 0, 0, 0, 0, 0, 0 };

            while (menu != 0)
            {
                Console.Write(str);
                menu = int.Parse(Console.ReadLine());

                switch (menu)
                {
                    case 1:
                        Console.WriteLine("예상번호를 출력합니다");
                        for (int i = 0; i < 5; i++)
                            RandomNumber();
                        Console.WriteLine();
                        break;
                    case 2:
                        Console.WriteLine("당첨여부를 조회하기 위한 번호 입력을 준비합니다.");
                        for (int i = 0; i < 6; i++)
                        {
                            Console.Write((i + 1) + " 번째 번호를 입력하세요 : ");
                            numList[i] = int.Parse(Console.ReadLine());
                        }
                        Console.WriteLine("입력된 번호의 당첨여부를 조회합니다.");
                        bool check = ReadExcelData(numList);
                        Console.WriteLine(check ? "미당첨번호입니다." : "당첨된 번호입니다.");
                        Console.WriteLine();
                        break;
                    case 0:
                        Console.WriteLine("종료합니다.");
                        break;
                    default:
                        break;
                }

            }

        }

        public static void RandomNumber()
        {
            int[] numList = new int[6] { 0, 0, 0, 0, 0, 0 };
            Random r = new Random();
            int num = 0;
            int cnt = 0;
            bool flag = true;

            num = r.Next(45) + 1;
            numList[0] = num;

            while (true)
            {
                if (cnt == 5)
                    break;

                num = r.Next(45) + 1;
                for (int i = 0; i < cnt + 1; i++)
                {
                    if (numList[i] == num)
                        flag = false;
                }
                if (flag)
                {
                    numList[cnt + 1] = num;
                    cnt++;
                }
                flag = true;
            }

            Array.Sort(numList);

            for (int i = 0; i < numList.Length; i++)
            {
                Console.Write(numList[i] + " ");
            }

            bool check = ReadExcelData(numList);
            Console.WriteLine(check ? "미당첨번호입니다." : "이미 당첨된 번호입니다.");
        }

        public static bool ReadExcelData(int[] array)
        {
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application();

                wb = excelApp.Workbooks.Open(@"D:\data.xlsx");

                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

                Excel.Range rng = ws.UsedRange;

                object[,] data = rng.Value;

                int cnt = 0;
                for (int r = 2; r <= data.GetLength(0); r++)
                {
                    for (int c = 2; c <= data.GetLength(1); c++)
                    {
                        //Console.Write(data[r, c].ToString() + " ");
                        if (array[c - 2] == int.Parse(data[r, c].ToString()))
                            cnt++;
                    }
                    //Console.WriteLine("");
                    if (cnt == 6)
                    {
                        wb.Close(true);
                        excelApp.Quit();
                        return false;
                    }
                    cnt = 0;
                }

                wb.Close(true);
                excelApp.Quit();
                return true;

            }
            finally
            {
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }
        }

        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
