using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgninfsc : PruebaCRUD
    {
        public static void CRUDKgninfsc(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");

            if (tipo == 1)
            {
                ElementList[10].SendKeys(data[0]);
                ElementList[8].SendKeys(data[1]);
                ElementList[7].SendKeys(data[2]);
                ElementList[9].SendKeys(data[3]);
            }
            else
            {
                ElementList[9].Clear();
                ElementList[9].SendKeys(data[4]);
            }

        }
    }
}
