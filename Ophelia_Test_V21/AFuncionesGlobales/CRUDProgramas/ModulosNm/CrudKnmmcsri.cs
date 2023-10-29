using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    //TDBEdit
    class CrudKnmmcsri
    {
        public static void KNmMcsriCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag)
        {
            int indice = flag == 0 ? variables.Count - 2 : variables.Count - 1;
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            if (flag == 0) PruebaCRUD.LupaDinamica(desktopSession, variables);
            desktopSession.Mouse.MouseMove(editFields[28].Coordinates);
            desktopSession.Mouse.Click(null);
            editFields[28].Clear();
            desktopSession.Keyboard.SendKeys(variables[indice]);

        }
    }
}
