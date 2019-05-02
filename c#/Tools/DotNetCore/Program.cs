using DotNetCore.TextLog;
using System;

namespace DotNetCore
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!"); 
            LogHelper.WriteLog(new Exception("---"));
        }
    }
}
