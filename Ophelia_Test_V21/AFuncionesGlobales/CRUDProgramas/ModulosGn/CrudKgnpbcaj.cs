﻿using System;
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
    class CrudKgnpbcaj : PruebaCRUD
    {
        public static void CRUDKgnpbcaj(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");
            if (tipo == 1)
            {

                ElementList[3].SendKeys(data[0]);
                ElementList[2].SendKeys(data[1]);
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(data[2]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(data[3]);
            }

        }
    }
}
