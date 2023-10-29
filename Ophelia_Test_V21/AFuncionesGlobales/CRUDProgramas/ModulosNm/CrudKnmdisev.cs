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
    class CrudKnmdisev
    {
        public static void KNmDisevCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag)
        {
        
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            desktopSession.FindElementByName("Tercero").Click();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (flag == 0)
            {
                PruebaCRUD.LupaDinamica(desktopSession, variables);
                desktopSession.Mouse.MouseMove(editFields[2].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[2].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 3]);
                Thread.Sleep(1000);

                /*desktopSession.Mouse.MouseMove(editFields[8].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[9].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 2]);
                Thread.Sleep(1000);*/
                ElementList[8].SendKeys(variables[6]);
            }
            else {
                List<string> editvalue = new List<string>(){ variables[variables.Count-1] };
                List<int> lupasOff = new List<int>() { 0, 1, 2, 3 };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, editvalue, lupasOff);
            }

            

        }
    }
}
