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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmhorne : FuncionesVitales
    {
        public static void CRUDKNmHorne(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBLookupComboBox");
            if (tipo == 1)
            {
                List<string> data = new List<string>() { crudVars[0] };
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                ElementList2[0].Click();
                desktopSession.Keyboard.SendKeys(crudVars[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                for (int i = 2; i < ElementList.Count; i++)
                {
                    desktopSession.Mouse.MouseMove(ElementList[i].Coordinates, 110, 15);
                    desktopSession.Mouse.Click(null);
                    WindowsDriver<WindowsElement> rootSession = null;
                    rootSession = PruebaCRUD.RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmHorario");
                    var ElementList3 = rootSession.FindElementsByName("Aceptar   ");
                    ElementList3[0].Click();
                    if (i==3)
                    {
                        i = 5;
                    }
                }
            }
            else
            {
                ElementList2[0].Click();
                desktopSession.Keyboard.SendKeys(crudVars[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }
    }
}
