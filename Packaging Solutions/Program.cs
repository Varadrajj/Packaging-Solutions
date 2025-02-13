using System;
using System.Collections.Generic;
using System.Linq;
using PackagingSolutions.Services;
using PackagingSolutions.Model;
using PackagingSolutions.View;


static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());  // Now correctly references Views.MainForm
    }
}