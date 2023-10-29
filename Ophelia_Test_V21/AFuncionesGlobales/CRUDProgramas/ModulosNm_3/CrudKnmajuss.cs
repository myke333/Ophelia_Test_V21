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
    class CrudKnmajuss
    {
        public static void KNmAjussCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TEdit");
            List<int> fieldIndex = new List<int>() { 3, 2 };
            for (int i = 0; i < fieldIndex.Count; i++) {
                editFields[fieldIndex[i]].Click();
                editFields[fieldIndex[i]].Clear();
                Thread.Sleep(1000);
                editFields[fieldIndex[i]].SendKeys(variables[1+i]);
                Thread.Sleep(1000);
            }

            PruebaCRUD.LupaDinamica(desktopSession, variables);
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
                    newFunctions_4.ScreenshotNuevo("Error", true, file);
                    Thread.Sleep(1000);
                    rootSession = PruebaCRUD.RootSession();
                    Thread.Sleep(5000);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                }
            }

            return errorMessages;

        }
    }
}
