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
    class CrudKnmpdpre
    {
        public static void KNmPdpreCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            List<int> fieldIndexes = new List<int>() { 7, 8, 9 };
            if (flag == 0)
            {
                PruebaCRUD.LupaDinamica(desktopSession, variables);

                for (int i = 0; i < editFields.Count; i++)
                {
                    if (i > 6 && i < 10)
                    {
                        desktopSession.Mouse.MouseMove(editFields[i].Coordinates);
                        Thread.Sleep(1000);
                        desktopSession.Mouse.Click(null);
                        editFields[i].Clear();
                        desktopSession.Keyboard.SendKeys(variables[i - 6]);
                        Thread.Sleep(1000);
                    }
                }

            }
            else {
                desktopSession.Mouse.MouseMove(editFields[7].Coordinates);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                editFields[7].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }
    }
}
