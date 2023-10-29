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
namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmparam
    {
        public static void KNmParamCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variablesPoliticas, List<string> variablesLeyCien, int flag) {
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
                procesoPoliticos(desktopSession, variablesPoliticas);
                Thread.Sleep(1000);
                newFunctions_4.changeWindow(desktopSession, " Ley 100 ");
                Thread.Sleep(1000);
                procesoLeyCien(desktopSession, variablesLeyCien);
                Thread.Sleep(2000);
                newFunctions_4.changeWindow(desktopSession, " Políticas Salariales ");

            }
            else {
                desktopSession.Mouse.MouseMove(editFields[9].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[9].Clear();
                desktopSession.Keyboard.SendKeys(variablesPoliticas[variablesPoliticas.Count-1]);
            }

        }

        public static void procesoPoliticos(WindowsDriver<WindowsElement> desktopSession, List<string> variables) {
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            int indexVariables = 0;
            foreach (var elem in editFields)
            {
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

        public static void procesoLeyCien(WindowsDriver<WindowsElement> desktopSession, List<string> variables) {
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            int index = 0;
            int indiceFields = 0; 
            List<int> indiceCampos = new List<int> { 18, 19, 21, 22 };
            foreach (var elem in editFields) {
                if (elem.Enabled) {
                    desktopSession.Mouse.MouseMove(elem.Coordinates);
                    Thread.Sleep(2000);
                    desktopSession.Mouse.Click(null);
                    elem.Clear();
                    if (indiceCampos.Contains(index))
                    {
                        desktopSession.Keyboard.SendKeys(variables[indiceFields]);
                        indiceFields++;
                        
                    }
                    else {
                        desktopSession.Keyboard.SendKeys("1");
                    }
                    
                }
                index++;
            }
        }


    }
}
