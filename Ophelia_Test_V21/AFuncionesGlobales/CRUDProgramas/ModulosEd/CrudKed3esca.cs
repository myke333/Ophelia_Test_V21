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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd
{
    class CrudKed3esca
    {
        public static void KEd3escaCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

            if (flag == 0)
            {
                List<int> fieldIndex = new List<int>() { 3, 2 };
                for (int i = 0; i < fieldIndex.Count; i++) {
                    editFields[fieldIndex[i]].Click();
                    editFields[fieldIndex[i]].Clear();
                    Thread.Sleep(2000);
                    editFields[fieldIndex[i]].SendKeys(variables[i]);
                    Thread.Sleep(2000);
                }
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

        public static void KEd3escaCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
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
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, 20, grids[0].Size.Height/6);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                for (int i = 0; i < variables.Count - 1; i++) {
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    Thread.Sleep(1000);
                    if (i < variables.Count - 2) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
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
