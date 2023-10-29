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
    class CrudKnmranre
    {
        public static void KNmRanreCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            List<int> fieldIndex = new List<int>() { 2, 3 };
            if (flag == 0)
            {
                for (int i = 0; i < fieldIndex.Count; i++) {
                    desktopSession.Mouse.MouseMove(editFields[fieldIndex[i]].Coordinates);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    editFields[fieldIndex[i]].Clear();
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    Thread.Sleep(1000);
                }
            }
            else {
                desktopSession.Mouse.MouseMove(editFields[fieldIndex[0]].Coordinates);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                editFields[fieldIndex[0]].Clear();
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }

        public static void KNmRanreCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> grids = desktopSession.FindElementsByClassName("TDBGrid");

            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width / 25, grids[0].Size.Height / 6);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);

                for (int i = 0; i < variables.Count - 1; i++)
                {
                    Thread.Sleep(1000);
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
