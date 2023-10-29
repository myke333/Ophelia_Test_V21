using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpasso : PruebaCRUD
    {
        public static void CRUDKnmpasso(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TScrollBox");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                //LupaDinamica2(rootSession, data);
                desktopSession.Mouse.MouseMove(ElementList2[0].Coordinates, 436, 52);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(ElementList2[0].Coordinates, 1065, 61);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                ElementList[6].SendKeys(data[1]);
                ElementList[5].SendKeys(data[2]);
                ElementList[7].SendKeys(data[3]);
                ElementList[8].SendKeys(data[4]);
            }
            else
            {
                ElementList[8].Clear();
                ElementList[8].SendKeys(data[5]);
            }
        }
    }
}
