using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmparei : PruebaCRUD
    {
        public static void CRUDKnmparei(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBCheckBox");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();

            if (tipo == 1)
            {
                Thread.Sleep(1000);
                ElementList[19].Click();
                Thread.Sleep(100);
                ElementList[18].Click();
                Thread.Sleep(100);
                ElementList[17].Click();
                Thread.Sleep(100);
                ElementList[16].Click();
                Thread.Sleep(100);
                ElementList[15].Click();
                Thread.Sleep(100);
                ElementList[14].Click();
                Thread.Sleep(100);

                ElementList[8].Click();
                Thread.Sleep(100);
                ElementList[7].Click();
                Thread.Sleep(100);
                ElementList[7].Click();
                Thread.Sleep(100);

                ElementList[6].Click();
                Thread.Sleep(100);
                ElementList[5].Click();
                Thread.Sleep(100);
                ElementList[4].Click();
                Thread.Sleep(100);
                ElementList[3].Click();
                Thread.Sleep(100);
                ElementList[2].Click();
                Thread.Sleep(100);
                ElementList[1].Click();
                Thread.Sleep(500);
            }

            else
            {
                ElementList[8].Click();
            }
        }
    }
}
