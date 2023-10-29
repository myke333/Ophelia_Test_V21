using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System.Drawing;
using OpenQA.Selenium;
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmacarb
    {
        public static void KNmAcArbCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;

            PruebaCRUD.LupaDinamica(desktopSession, variables);
            Thread.Sleep(1000);
            List<int> lupasOff = new List<int>() { 0, 2 };
            List<string> editData = new List<string>() { variables[2], variables[3] };
            newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, editData, lupasOff);
            Thread.Sleep(1000);

        }


        public static List<string> BotonAceptar(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////
            WebDriverWait printScreen;
            WindowsElement printerField;
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> botones = desktopSession.FindElementsByClassName("TBitBtn");
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();

            foreach (var elem in botones)
            {
                if (elem.Text == "Aceptar")
                {
                    elem.Click();
                    Thread.Sleep(3000);
                    newFunctions_4.ScreenshotNuevo("Barra de progreso", true, file);
                    Thread.Sleep(5000);
                    rootSession = PruebaCRUD.RootSession();
                    Thread.Sleep(5000);
                    newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
                    Thread.Sleep(2000);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                }
            }

            return errorMessages;

        }

    }
}
