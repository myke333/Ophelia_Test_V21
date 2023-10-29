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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdencer : FuncionesVitales
    {
        public static void CRUDKFdEncer(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            var TGroupButton = desktopSession.FindElementsByClassName("TGroupButton"); 

            ElementList[2].SendKeys(crudVars[0]);
            TGroupButton[1].Click();
        }
    }
}