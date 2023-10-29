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
    class CrudKsopcont
    {
        public static void KSoPcontCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButtons = desktopSession.FindElementsByClassName("TGroupButton");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

            if (flag == 0)
            {
                foreach (var button in radioButtons) {
                    if (button.Text == variables[0]) {
                        desktopSession.Mouse.MouseMove(button.Coordinates);
                        Thread.Sleep(1000);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }

                editFields[2].Click();
                editFields[2].Clear();
                Thread.Sleep(1000);
                editFields[2].SendKeys(variables[1]);
                Thread.Sleep(1000);

            }
            else {
                Thread.Sleep(1000);
                editFields[2].Click();
                editFields[2].Clear();
                Thread.Sleep(1000);
                editFields[2].SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }
    }
}
