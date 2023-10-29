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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmmdane
    {
        public static void KNmMdaneCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> buttons = desktopSession.FindElementsByClassName("TGroupButton");

            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(editFields[26].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[26].Clear();
                desktopSession.Keyboard.SendKeys(variables[4]);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(editFields[10].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[10].Clear();
                desktopSession.Keyboard.SendKeys(variables[3]);
                Thread.Sleep(1000);

                List<int> lupasOff = new List<int>() { 0 };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, variables, lupasOff);
                Thread.Sleep(1000);
                
                foreach (var elem in buttons) {
                    if (elem.Text == variables[variables.Count - 2]) {
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }
            }
            else {
                desktopSession.Mouse.MouseMove(editFields[10].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[10].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }
        }
    }
}
