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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdparge
    {
        public static void KFdPargeCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> checkBoxes = desktopSession.FindElementsByClassName("TDBCheckBox");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButtons = desktopSession.FindElementsByClassName("TGroupButton");
            if (flag == 0)
            {

                desktopSession.Mouse.MouseMove(editFields[2].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[2].Clear();
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[2]);
                Thread.Sleep(1000);

                foreach (var elem in checkBoxes) {
                    if (elem.Text == "Actualizar Ausentismos") {
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }

                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);

                foreach (var elem in radioButtons) {
                    if (elem.Text == variables[variables.Count-2])
                    {
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }


            }
            else
            {

                foreach (var elem in radioButtons)
                {
                    if (elem.Text == variables[variables.Count - 1])
                    {
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }
            }
        }
    }
}
