using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class crudKNmAutod : FuncionesVitales
    {

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            Debug.WriteLine($"la cantidad de elementos es: {ElementList.Count}");
            int cont = 0;
            List<string> variables = new List<string>() { crudVars[0], crudVars[1] };
            if (tipo == 1)
            {
                foreach (var elem in ElementList)
                {
                    if (elem.Enabled == true)
                    {
                        elem.SendKeys(variables[0]);
                        cont += 1;
                    }
                    Thread.Sleep(500);
                }

            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(crudVars[1]);
            }
        }

    }
}
