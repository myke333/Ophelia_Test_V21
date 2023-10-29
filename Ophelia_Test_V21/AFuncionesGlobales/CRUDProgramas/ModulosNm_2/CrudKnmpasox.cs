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
    class CrudKnmpasox
    {
        public static void KNmPasoxCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            List<int> fieldPos = new List<int>() { 4, 5, 6, 7, 8 };
            List<string> fieldValues = variables.GetRange(1, 5);
            if (flag == 0)
            {
                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(2000);

                for (int i = 0; i < fieldPos.Count; i++)
                {
                    desktopSession.Mouse.MouseMove(editFields[fieldPos[i]].Coordinates);
                    Thread.Sleep(1000);
                    editFields[fieldPos[i]].Clear();
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(fieldValues[i]);
                    Thread.Sleep(1000);
                }

            }
            else {
                desktopSession.Mouse.MouseMove(editFields[8].Coordinates);
                Thread.Sleep(1000);
                editFields[8].Clear();
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }
    }
}
