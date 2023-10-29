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
    class CrudKnmparco : FuncionesVitales
    {
        public static void CRUDKNmParco(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var TPageControl = desktopSession.FindElementsByClassName("TPageControl");

            WindowsDriver<WindowsElement> rootSession = null;
            //Debug.WriteLine("con: " + ElementList.Count);
            //for (int i = 0; i <= ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    Thread.Sleep(500);
            //}

            //List<string> data1 = new List<string>() { crudVars[15], crudVars[16], crudVars[17]};
            //List<string> data2 = new List<string>() { crudVars[18], crudVars[19], crudVars[20], crudVars[21], crudVars[22], crudVars[23] };
            //List<string> data3 = new List<string>() { crudVars[24], crudVars[25], crudVars[26], crudVars[27], crudVars[28], crudVars[29], crudVars[30], crudVars[31] };
            //List<string> data4 = new List<string>() { crudVars[32], crudVars[33], crudVars[34], crudVars[35], crudVars[36], crudVars[37] };
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (tipo == 1)
            {
                ElementList[15].SendKeys(crudVariables[0]);
                Thread.Sleep(500);

                ElementList[13].SendKeys(crudVariables[1]);
                Thread.Sleep(500);
                ElementList[11].SendKeys(crudVariables[2]);
                Thread.Sleep(500);
                ElementList[9].SendKeys(crudVariables[3]);
                Thread.Sleep(500);
                ElementList[7].SendKeys(crudVariables[4]);
                Thread.Sleep(500);
                ElementList[3].SendKeys(crudVariables[5]);
                Thread.Sleep(500);
                //desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 30, 10);
                //Thread.Sleep(500);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                //rootSession = PruebaCRUD.RootSession();
                //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                //ElementList = rootSession.FindElementsByClassName("TDBEdit");
                //TPageControl = rootSession.FindElementsByClassName("TPageControl");

                //desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 30, 10);
                //Thread.Sleep(500);
                //desktopSession.Mouse.Click(null);

                //ElementList[21].SendKeys(crudVars[6]);
                //Thread.Sleep(500);
                //ElementList[13].SendKeys(crudVars[7]);
                //Thread.Sleep(500);
                //ElementList[19].SendKeys(crudVars[8]);
                //Thread.Sleep(500);
                //ElementList[9].SendKeys(crudVars[9]);
                //Thread.Sleep(500);
                //ElementList[17].SendKeys(crudVars[10]);
                //Thread.Sleep(500);
                //ElementList[11].SendKeys(crudVars[11]);
                //Thread.Sleep(500);
                //ElementList[15].SendKeys(crudVars[12]);
                //Thread.Sleep(500);
                //ElementList[5].SendKeys(crudVars[13]);
                //Thread.Sleep(500);
                //ElementList[3].SendKeys(crudVars[14]);
                //Thread.Sleep(500);
                //rootSession.Mouse.MouseMove(TPageControl[0].Coordinates, 75, 10);
                //rootSession.Mouse.Click(null);
                //rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                //PruebaCRUD.LupaDinamica(desktopSession, data1);
                //Thread.Sleep(500);
                //rootSession = PruebaCRUD.RootSession();
                //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                //TPageControl = rootSession.FindElementsByClassName("TPageControl");
                //rootSession.Mouse.MouseMove(TPageControl[1].Coordinates, 30, 10);
                //rootSession.Mouse.Click(null);
                //rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                //PruebaCRUD.LupaDinamica(desktopSession, data2);
                //rootSession.Mouse.MouseMove(TPageControl[0].Coordinates, 300, 10);
                //rootSession.Mouse.Click(null);
                //PruebaCRUD.LupaDinamica(desktopSession, data3);
                //rootSession = PruebaCRUD.RootSession();
                //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                //TPageControl = rootSession.FindElementsByClassName("TPageControl");
                //rootSession.Mouse.MouseMove(TPageControl[0].Coordinates, 357, 10);
                //rootSession.Mouse.Click(null);
                //PruebaCRUD.LupaDinamica(desktopSession, data4);

            }
            else
            {

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 484, 308);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);



                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();



                RootSession1.Keyboard.SendKeys(crudVariables[6]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
             
                Thread.Sleep(1000);
            }
        }
    }
}