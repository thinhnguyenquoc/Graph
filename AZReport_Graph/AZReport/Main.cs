using Ninject;
using System;
using System.Windows.Forms;

namespace AZReport
{
    static class MainProgram
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var kernel = new StandardKernel(new NinjectBindingModel());
            var form = kernel.Get<Form1>();
            Application.Run(form);
        }
    }
}