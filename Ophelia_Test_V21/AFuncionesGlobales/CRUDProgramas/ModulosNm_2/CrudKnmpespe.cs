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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpespe
    {
        public static void KNmPespeCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> comboBoxes = desktopSession.FindElementsByClassName("TDBLookupComboBox");
            if (flag == 0)
            {
                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);


                //combobox Aportes
                desktopSession.Mouse.MouseMove(comboBoxes[1].Coordinates);
                desktopSession.Mouse.Click(null);
                comboBoxes[1].Clear();
                desktopSession.Keyboard.SendKeys(variables[1].Substring(0,1));
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);


            }
            else {
                //combobox Aportes
                desktopSession.Mouse.MouseMove(comboBoxes[1].Coordinates);
                desktopSession.Mouse.Click(null);
                comboBoxes[1].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1].Substring(0, 1));
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }
        }
    }
}
