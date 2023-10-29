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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmtipex
    {
        public static void KNmTipexCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> comboBoxes = desktopSession.FindElementsByClassName("TDBComboBox");

            if (flag == 0)
            {
               for(int i = 0; i < comboBoxes.Count; i++) {
                    desktopSession.Mouse.MouseMove(comboBoxes[i].Coordinates);
                    Thread.Sleep(1000);
                    comboBoxes[i].Click();
                    comboBoxes[i].Clear();
                    Thread.Sleep(1000);
                    comboBoxes[i].SendKeys(variables[i]);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                }
            }
            else {
                desktopSession.Mouse.MouseMove(comboBoxes[0].Coordinates);
                Thread.Sleep(1000);
                comboBoxes[0].Click();
                comboBoxes[0].Clear();
                Thread.Sleep(1000);
                comboBoxes[0].SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }

        }
    }
}
