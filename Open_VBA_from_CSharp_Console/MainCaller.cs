using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel; //References>Add Reference>COM>Type Libraries>Microsoft Office 14.0 Project Library


class MainCaller
{
    static void open_file(string str_path,Application xlApp)
    {
        Workbooks xlWb = xlApp.Workbooks;
        Workbook xlWbSh = xlWb.Open(str_path);
    }

    static Application get_file(string str_path)
    {
        Application xlApp = new Application();
        xlApp.Visible = true;
        open_file(str_path, xlApp);
        return xlApp;
    }

    static void Main()
    {
        string str_path = @"C:\Users\v.doynov\Desktop\CodeMe.xlsb";
        Application xlMyApp = get_file(str_path);
        xlMyApp.Run("Main");
        //get_file(str_path).Run("Main");
        Console.WriteLine("Operation finished!");
    }
}