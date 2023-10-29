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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmpamde
    {
        public static void KNmPamdeCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButtons = desktopSession.FindElementsByClassName("TGroupButton");

            if (flag == 0)
            {
                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);
                foreach (var button in radioButtons)
                {
                    if (button.Text == variables[variables.Count - 2]) {
                        desktopSession.Mouse.MouseMove(button.Coordinates);
                        Thread.Sleep(1000);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }
            }
            else {
                foreach (var button in radioButtons) {
                    if (button.Text == variables[variables.Count - 1])
                    {
                        desktopSession.Mouse.MouseMove(button.Coordinates);
                        Thread.Sleep(1000);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }
            }
            
        }
    }
}
