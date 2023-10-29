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


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoprogi
    {
        public static void KSoProgiCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> fieldYear = desktopSession.FindElementsByClassName("TSpinEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> fieldMonth = desktopSession.FindElementsByClassName("TComboBox");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButtons = desktopSession.FindElementsByClassName("TGroupButton");

            fieldYear[0].Click();
            fieldYear[0].Clear();
            Thread.Sleep(1000);
            fieldYear[0].SendKeys(variables[0]);
            Thread.Sleep(1000);

            fieldMonth[0].Click();
            fieldMonth[0].Clear();
            Thread.Sleep(1000);
            fieldMonth[0].SendKeys(variables[1]);
            Thread.Sleep(1000);

            foreach (var elem in radioButtons) {
                if (elem.Text == variables[2]) {
                    elem.Click();
                    Thread.Sleep(1000);
                }
            }

        }

        public static List<string> BotonAceptarConBarraProgreso(WindowsDriver<WindowsElement> desktopSession, string file)
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
                    rootSession = PruebaCRUD.RootSession();
                    Thread.Sleep(5000);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    newFunctions_4.ScreenshotNuevo("Progreso del Boton Aceptar", true, file);
                    Thread.Sleep(20000);
                    newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                }
            }

            return errorMessages;

        }

    }
}
