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
    class CrudKnmsapic
    {
        public static void KNmSapicCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButtons = desktopSession.FindElementsByClassName("TGroupButton");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> checkBoxes = desktopSession.FindElementsByClassName("TCheckBox");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> dates = desktopSession.FindElementsByClassName("TEdit");

            if (flag == 0)
            {
                foreach (var radio in radioButtons) {
                    if (radio.Text == variables[0]) {
                        radio.Click();
                        Thread.Sleep(1000);
                    }
                }

                foreach (var box in checkBoxes) {
                    if (box.Text == variables[1]) {
                        box.Click();
                        Thread.Sleep(1000);
                    }
                }

                for (int i = 0; i < dates.Count; i++) {
                    dates[i].Click();
                    dates[i].Clear();
                    Thread.Sleep(1000);
                    dates[i].SendKeys(variables[2 + i]);
                }

            }
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
                    rootSession = PruebaCRUD.RootSession();
                    Thread.Sleep(5000);
                    newFunctions_4.ScreenshotNuevo("Barra de progreso", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(4000);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(5000);
                    newFunctions_4.ScreenshotNuevo("Error al terminar proceso", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                }
            }

            return errorMessages;

        }


        public static List<string> BotonErrores(WindowsDriver<WindowsElement> desktopSession, string ruta, string file)
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
                if (elem.Text == "Errores")
                {
                    elem.Click();
                    Thread.Sleep(2000);
                    try
                    {
                        errors = newFunctions_5.GenerarReportes("Errores", file, ruta);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                    }
                    catch
                    {
                        Thread.Sleep(3000);
                        newFunctions_4.ScreenshotNuevo("Error al generar preliminar", true, file);
                        rootSession = PruebaCRUD.RootSession();
                        Thread.Sleep(10000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2000);
                    }

                }
            }

            return errorMessages;

        }

    }
}
