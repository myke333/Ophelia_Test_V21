using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbicarpe
    {

        public static void KBiCarpeCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> val, int flag, int fieldIndex, string editValue)
        {

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

            int i = 0;


            if (flag == 0)
            {
                foreach (var elem in editFields)
                {

                    if (elem.Enabled)
                    {
                        elem.ClearCache();
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        desktopSession.Keyboard.SendKeys(val[i]);
                        i++;
                    }
                    Thread.Sleep(1000);
                }
            }
            else
            {
                editFields[fieldIndex].Clear();
                desktopSession.Mouse.MouseMove(editFields[fieldIndex].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[fieldIndex].SendKeys(editValue);
            }


        }
        static public void ventanaSum(WindowsDriver<WindowsElement> desktopSession)
        {







            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");







            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1000, 415);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);



        }
    }
}
