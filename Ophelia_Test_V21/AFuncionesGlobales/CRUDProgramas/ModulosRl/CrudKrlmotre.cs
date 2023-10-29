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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosRl
{
    class CrudKrlmotre
    {
        public static void KRlMotreCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

            if (flag == 0)
            {
                for (int i = 2; i <= 3; i++)
                {
                    editFields[i].Click();
                    Thread.Sleep(1000);
                    editFields[i].Clear();
                    desktopSession.Keyboard.SendKeys(variables[i - 2]);
                    Thread.Sleep(1000);

                }
            }
            else
            {
                editFields[2].Click();
                Thread.Sleep(1000);
                editFields[2].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 1]);
                Thread.Sleep(1000);
            }
        }
    }
}
