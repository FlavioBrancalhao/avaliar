using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace avaliar
{
    class HandleApp
    {
        //Enumera os tipos para usar com o switch (o 1,2,3 são da API do user32)
        public enum Actions { Normal = 1, Minimize = 2, Maximize = 3 };

        //Importa o user32.dll para poder usar as APIs nativas
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        //Busca um aplicativo pelo nome
        public static IntPtr FindWindow(string title)
        {
            Process[] pros = Process.GetProcessesByName(title);

            if (pros.Length == 0)
                return IntPtr.Zero;

            return pros[0].MainWindowHandle;
        }

        //Dispara a ação desejada, só tem 3 opções no exemplo
        public static void Action(string name, Actions act)
        {
            IntPtr hWnd = FindWindow(name);

            if (!hWnd.Equals(IntPtr.Zero))
                ShowWindowAsync(hWnd, (int)act);
        }
    }
}
