﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;


namespace peteli.flaus
{
    
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string HelloDna(string name)
        {
            return "Hello " + name;
        }
    }
}
