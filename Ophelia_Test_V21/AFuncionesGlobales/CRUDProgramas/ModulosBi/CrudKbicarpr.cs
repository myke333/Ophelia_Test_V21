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
    class CrudKbicarpr
    {
        public static void KBiCarprCRUD(WindowsDriver<WindowsElement> desktopSession, int fieldIndex, string editValue)
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


           
            editFields[fieldIndex].Clear();
            desktopSession.Mouse.MouseMove(editFields[fieldIndex].Coordinates);
            desktopSession.Mouse.Click(null);
            editFields[fieldIndex].SendKeys(editValue);
            


        }
    }
}
