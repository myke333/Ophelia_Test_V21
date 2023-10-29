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
    class CrudKnmpcoti
    {
        public static void KNmPcotiCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButtons = desktopSession.FindElementsByClassName("TGroupButton");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> comboBoxes = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

            if (flag == 0)
            {
                foreach (var radio in radioButtons)
                {
                    if (radio.Text == variables[2])
                    {
                        desktopSession.Mouse.MouseMove(radio.Coordinates);
                        desktopSession.Mouse.Click(null);
                    }
                }

                desktopSession.Mouse.MouseMove(comboBoxes[0].Coordinates);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                comboBoxes[0].Clear();
                desktopSession.Keyboard.SendKeys(variables[3]);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(editFields[4].Coordinates);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                editFields[4].Clear();
                desktopSession.Keyboard.SendKeys(variables[4]);
                Thread.Sleep(1000);


                desktopSession.Mouse.MouseMove(editFields[5].Coordinates);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                editFields[5].Clear();
                desktopSession.Keyboard.SendKeys(variables[5]);
                Thread.Sleep(1000);

                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);

            }
            else {
                List<int> lupasOff = new List<int>() { 0 };
                List<string> editVal = new List<string>() { variables[variables.Count - 1] };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, editVal, lupasOff);
                Thread.Sleep(1000);
                
            }

        }


        public static void KNmPcotiCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
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
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width/8, grids[0].Size.Height/6);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                
                foreach (var variable in variables) {
                    desktopSession.Keyboard.SendKeys(variable);
                    Thread.Sleep(1000);
                    if (variables[variables.Count - 2] != variable)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    else {
                        break;
                    }
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
