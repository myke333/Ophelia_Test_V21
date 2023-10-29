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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGd
{
    class CrudKgdtiacc
    {
        public static void KGdTiaccCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
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
                List<int> fieldIndex = new List<int>() { 9, 2 };
                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);

                for (int i = 0; i < fieldIndex.Count; i++) {
                    desktopSession.Mouse.MouseMove(editFields[fieldIndex[i]].Coordinates);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    editFields[fieldIndex[i]].Clear();
                    Thread.Sleep(1000);
                    editFields[fieldIndex[i]].SendKeys(variables[3 + i]);
                    Thread.Sleep(1000);
                }
            }
            else {
                desktopSession.Mouse.MouseMove(editFields[2].Coordinates);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                editFields[2].Clear();
                Thread.Sleep(1000);
                editFields[2].SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }
        }

        public static void KGdTiaccCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
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
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width / 18, grids[0].Size.Height / 5);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);

                for (int i = 0; i < variables.Count-1; i++) {
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    Thread.Sleep(1000);
                    if(i < variables.Count-2 ) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
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
