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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmplaca
    {
        public static void CRUDKNmPlaca(WindowsDriver<WindowsElement> desktopSession, int tipo, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TGroupButton");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBEdit"); 


            if (tipo == 1)
            {
                ElementList[14].Click();
            }
            else
            {
                ElementList[15].Click();
            }

        }
    }
}
