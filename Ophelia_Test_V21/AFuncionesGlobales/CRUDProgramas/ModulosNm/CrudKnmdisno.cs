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
    class CrudKnmdisno
    {
        public static void KNmDisnoCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag)
        {
            List<string> varLupas = variables.GetRange(0, 3);
            List<int> lupasOff = new List<int>() { 3, 4, 5 };
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
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, varLupas, lupasOff);
                desktopSession.Mouse.MouseMove(editFields[5].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[5].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 2]);
                Thread.Sleep(1000);
            }
            else
            {
                List<string> editvalue = new List<string>() { variables[variables.Count - 1] };
                lupasOff = new List<int>() { 0, 1, 3, 4, 5};
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, editvalue, lupasOff);
            }
        }
    }
}
