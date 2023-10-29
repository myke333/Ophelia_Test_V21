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
    class CrudKnmtipsi
    {
        public static void KNmTipsiCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> comboBoxes = desktopSession.FindElementsByClassName("TDBComboBox");
            List<int> fieldIndex = new List<int>() { 5, 4};
            if (flag == 0)
            {
                for (int i = 0; i < fieldIndex.Count; i++)
                {
                    editFields[fieldIndex[i]].Click();
                    editFields[fieldIndex[i]].Clear();
                    Thread.Sleep(1000);
                    editFields[fieldIndex[i]].SendKeys(variables[i+1]);
                    Thread.Sleep(1000);
                }

                desktopSession.Mouse.MouseMove(comboBoxes[0].Coordinates, comboBoxes[0].Size.Width - 10, comboBoxes[0].Size.Height / 2);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(comboBoxes[0].Coordinates, comboBoxes[0].Size.Width/2, comboBoxes[0].Size.Height / 2);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                for (int i = 0; i < 5; i++) {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(1000);
                }
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);

                editFields[3].Click();
                Thread.Sleep(3000);

                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);




            }
            else {
                Thread.Sleep(1000);
                editFields[4].Click();
                editFields[4].Clear();
                Thread.Sleep(1000);
                editFields[4].SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }
        public static void KNmTipsiCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> grids = desktopSession.FindElementsByClassName("TDBGrid");
            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, 30, grids[0].Size.Height/6);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                for (int i = 0; i < variables.Count - 1; i++) {
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    if (i < variables.Count-2) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1000);
                }
            }
            else {
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }

    }
}
