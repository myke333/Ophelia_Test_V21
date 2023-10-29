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
    class CrudKbiexame
    {
        public static void KBiexameCRUDDet(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int indexWindows, int funcion) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBGrid");



            if (indexWindows == 1)
            {
                if (funcion == 0)
                {
                    desktopSession.Mouse.MouseMove(editFields[indexWindows].Coordinates, 50, 27);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);

                    for (int index = 0; index < variables.Count - 1; index++)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(1000);
                        desktopSession.Keyboard.SendKeys(variables[index]);
                        Thread.Sleep(1000);
                    }
                }
                else
                {
                    desktopSession.Keyboard.SendKeys(variables[variables.Count - 1]);
                    Thread.Sleep(1000);
                }

            }
            else {
                if (indexWindows == 0) {
                    if (funcion == 0)
                    {
                        desktopSession.Mouse.MouseMove(editFields[indexWindows].Coordinates, 50, 27);
                        Thread.Sleep(1000);
                        desktopSession.Mouse.Click(null);

                        for (int index = 0; index < variables.Count - 1; index++)
                        {
                            if (index == 2) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            desktopSession.Keyboard.SendKeys(variables[index]);
                            Thread.Sleep(1000);
                        }
                    }
                    else {
                        desktopSession.Keyboard.SendKeys(variables[variables.Count - 1]);
                        Thread.Sleep(1000);
                    }
                }
            }
        }
    }
}
