namespace TriedExcel
{
    using System;
    using System.Collections.Generic;

    public class Information
    {
        public string Name { get; set; }
        public int Sum { get; set; }
        public int LastRow { get; set; }

        public Information(string name, int sum, int lastRow)
        {
            Name = name;
            Sum = sum;
            LastRow = lastRow;
        }

        public static void PrintInformation(List<Information> information)
        {
            Console.WriteLine("Printing the whole information at once:");
            foreach (Information info in information)
            {
                Console.WriteLine(info.ToString());
            }
        }
        
        public override string ToString()
        {
            return ($"{this.Name} \tlast row - {this.LastRow} - \tsum - {this.Sum}");
        }
    }
}