using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Porag_R
{
    class Program
    {
        private static List<Roll> rolls = new List<Roll>();
        private static List<Item> items = new List<Item>();
        private static int totalRollArea = 0;
        private static int totalItemArea = 0;
        static void Main(string[] args)
        {
            readFile();
            if (totalItemArea > totalRollArea)
            {
                Console.WriteLine("Item area is bigger than total roll area. Problem cannot be solved");
                /*Console.ReadKey();
                System.Environment.Exit(0);*/
            }
            rolls.Sort((x, y) => x.width.CompareTo(y.width));
            rolls.Reverse();
            items.Sort((x, y) => x.width.CompareTo(y.width));
            items.Reverse();
            for (int i = 0; i < items.Count; i++)
            {
                for (int j = 0; j < rolls.Count; j++)
                {
                    if (items[i].width > rolls[j].width)
                    {
                        continue;
                    }
                    else
                    {

                    }
                    
                }
                
            }
            Console.WriteLine(totalRollArea);
            Console.WriteLine(totalItemArea);
            Console.ReadKey();
        }

        static void readFile()
        {
            FileInfo intialInfo = new FileInfo("SIze-Samples.xlsx");
            var excel = new ExcelPackage(intialInfo);
            foreach (ExcelWorksheet workSheet in excel.Workbook.Worksheets)
            {
                var start = workSheet.Dimension.Start;
                var end = workSheet.Dimension.End;
                int rollNumber = 1;
                //read rolls
                for (int i = 2; i <= 5; i++)
                {
                    Object cellValue = workSheet.Cells[2, i].Value;
                    Object cellValue2 = workSheet.Cells[3, i].Value;
                    if (cellValue != null && cellValue2 != null)
                    {
                        Roll r = new Roll();
                        r.Number = rollNumber++;
                        r.height = Convert.ToInt32(cellValue2);
                        r.width = Convert.ToInt32(cellValue);
                        totalRollArea += (r.height * r.width);
                        rolls.Add(r);
                    }
                }

                //read Bedrooms 
                for (int i = 8; i <= 68; i++)
                {
                    Object cellValue = workSheet.Cells[i, 2].Value;
                    Object cellValue2 = workSheet.Cells[i, 3].Value;
                    Object cellValue3 = workSheet.Cells[i, 4].Value;
                    Object cellValue4 = workSheet.Cells[i, 5].Value;
                    if (cellValue != null && cellValue2 != null)
                    {
                        Item item = new Item();

                        item.quantity = Convert.ToInt32(cellValue);
                        item.width = Convert.ToInt32(cellValue2);
                        item.height = Convert.ToInt32(cellValue3);
                        item.number = Convert.ToInt32(cellValue4);
                        totalItemArea += item.width * item.height;
                        items.Add(item);
                    }
                }

                //read livingRooms 
                for (int i = 8; i <= 94; i++)
                {
                    Object cellValue = workSheet.Cells[i, 9].Value;
                    Object cellValue2 = workSheet.Cells[i, 10].Value;
                    Object cellValue3 = workSheet.Cells[i, 11].Value;
                    Object cellValue4 = workSheet.Cells[i, 12].Value;
                    if (cellValue != null && cellValue2 != null)
                    {
                        Item item = new Item();

                        item.quantity = Convert.ToInt32(cellValue);
                        item.width = Convert.ToInt32(cellValue2);
                        item.height = Convert.ToInt32(cellValue3);
                        item.number = Convert.ToInt32(cellValue4);
                        totalItemArea += item.width * item.height;
                        items.Add(item);
                    }
                }

            }
               
        }
    }
}
