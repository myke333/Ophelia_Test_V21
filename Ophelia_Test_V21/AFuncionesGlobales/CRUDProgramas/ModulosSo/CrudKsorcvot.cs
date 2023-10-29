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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsorcvot
    {
        public static void KSoRcvotCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> dates = desktopSession.FindElementsByClassName("TEdit");

            PruebaCRUD.LupaDinamica(desktopSession, variables);
            Thread.Sleep(1000);
            for (int i = 0; i < variables.Count - 1; i++) {
                dates[2 + i].Click();
                dates[2 + i].Clear();
                Thread.Sleep(1000);
                dates[2 + i].SendKeys(variables[1 + i]);
                Thread.Sleep(1000);
            }

        }
    }
}
