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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc
{
    class CrudKacorgan
    {

        public static void KAcOrganCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> checkBoxes = desktopSession.FindElementsByClassName("TDBCheckBox");
            //System.Diagnostics.//Debugger.Launch();
            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(editFields[6].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[6].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 2]);
                Thread.Sleep(2000);
                foreach (var checkBox in checkBoxes)
                {
                    if (checkBox.Text == "Agrupar Nivel Árbol")
                    {
                        desktopSession.Mouse.MouseMove(checkBox.Coordinates);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                    }
                }
                newFunctions_4.LupaDinamica(desktopSession, variables, new Point(0, 0));
               
            }
            else { 
                List<string> editValue = new List<string>() { variables[variables.Count-1] };
                List<int> lupasOff = new List<int>() { 0, 2 };
                newFunctions_4.LupaDinamicaDiscriminatoria2(desktopSession, editValue, new Point(0,0), lupasOff);
                Thread.Sleep(1000);
            }


        }

    }
}
