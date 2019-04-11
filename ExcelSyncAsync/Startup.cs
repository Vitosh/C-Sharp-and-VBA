namespace TriedExcel
{
    using System;
    using System.Diagnostics;
    using TriedExcel.Reader;

    public class Startup
    {
        public const string filePath = @"C:\Users\gropc\Desktop\Sample.xlsx";

        static void Main()
        {
            Stopwatch stopWatch = Stopwatch.StartNew();
            AsyncReader asyncReader = new AsyncReader(filePath);
            asyncReader.MainAsync().GetAwaiter().GetResult();
            stopWatch.Stop();
            string resultAsync = ($"{stopWatch.Elapsed}");

            stopWatch.Start();
            SyncReader syncReader = new SyncReader(filePath);
            syncReader.MainSync();
            stopWatch.Stop();
            string resultSync = ($"{stopWatch.Elapsed}");

            Console.WriteLine($"\nAsync \t\t {resultAsync}\nSync \t\t {resultSync}");
        }
    }
}

