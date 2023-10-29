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
    class CrudKnmdestu
    {
        public static void KNmDestuCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag)
        {
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
                int indexVariables = 0;
                foreach (var elem in editFields){
                    if (elem.Enabled)
                    {
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        elem.Clear();
                        desktopSession.Keyboard.SendKeys(variables[indexVariables]);
                        indexVariables++;
                        Thread.Sleep(1000);
                    }
                }
            }
            else {
                desktopSession.Mouse.MouseMove(editFields[4].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[4].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }
        }
    }
}
