using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_6
{
    class CrudKnmbenes : FuncionesVitales
    {
        public static void CRUDKNmBenes(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            if (tipo == 1)
            {
                List<int> dis = new List<int>() { 2, 3 };
                List<string> data = new List<string>() { crudVars[0], crudVars[1] };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, data, dis);                
            }
            else
            {
                List<int> dis = new List<int>() { 0, 2, 3 };
                List<string> data = new List<string>() { crudVars[2] };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, data, dis);
            }
        }
    }
}
