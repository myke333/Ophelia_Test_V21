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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmparhe : FuncionesVitales
    {
        public static void CRUDKNmParhe(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file, string motor)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBCheckBox");
            
            WindowsDriver<WindowsElement> rootSession = null;
            
            if (tipo == 1)
            {                
                if (motor == "SQL")
                {
                    List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3], crudVars[4], crudVars[5], crudVars[6], crudVars[7], crudVars[8] };
                    PruebaCRUD.LupaDinamica(desktopSession, data);
                    ElementList2[0].Click();
                    Thread.Sleep(1000);
                    ElementList2[0].Click();
                    Thread.Sleep(1000);
                    ElementList[2].SendKeys(crudVars[9]);
                    ElementList[3].SendKeys(crudVars[5]);
                    ElementList[8].SendKeys(crudVars[13]);
                    ElementList[9].SendKeys(crudVars[7]);
                    ElementList[11].SendKeys(crudVars[11]);
                    ElementList[7].SendKeys(crudVars[12]);
                    ElementList[6].SendKeys(crudVars[12]);
                    ElementList[10].SendKeys(crudVars[10]);
                    Thread.Sleep(1000);

                    //desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, /*27, 8*/30, 10);
                    //Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);
                    //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);



                    //rootSession = PruebaCRUD.RootSession();

                    //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");

                    //ElementList = rootSession.FindElementsByClassName("TDBEdit");

                    //ElementList2 = rootSession.FindElementsByClassName("TDBCheckBox");

                    //ElementList2[1].Click();
                    //Thread.Sleep(1000);
                    //ElementList[23].SendKeys(crudVars[14]);
                    //Thread.Sleep(1000);
                    //ElementList[25].SendKeys(crudVars[15]);
                    //Thread.Sleep(1000);
                    //ElementList[29].SendKeys(crudVars[16]);
                    //Thread.Sleep(1000);
                    //ElementList[27].SendKeys(crudVars[17]);
                    //Thread.Sleep(1000);
                }
                if (motor == "ORA")
                {
                    List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3], crudVars[4] };
                    List<int> Dislupa = new List<int>() { 2,4,6,8 };
                    newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, data, Dislupa);
                    ElementList2[0].Click();
                    ElementList2[0].Click();
                    ElementList[2].SendKeys(crudVars[5]);
                    ElementList[3].SendKeys(crudVars[5]);
                    ElementList[8].SendKeys(crudVars[9]);
                    ElementList[7].SendKeys(crudVars[8]);
                    ElementList[10].SendKeys(crudVars[6]);
                    ElementList[6].SendKeys(crudVars[9]);
                    ElementList[9].SendKeys(crudVars[7]);


                    Thread.Sleep(500);
                    //desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 30, 10);
                    //desktopSession.Mouse.Click(null);
                    //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                    //rootSession = PruebaCRUD.RootSession();
                    //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                    //ElementList = rootSession.FindElementsByClassName("TDBEdit");
                    //ElementList2 = rootSession.FindElementsByClassName("TDBCheckBox");
                    //ElementList2[1].Click();
                    //List<string> data2 = new List<string>() { crudVars[10], crudVars[11], crudVars[12], crudVars[13], crudVars[14], crudVars[15] };
                    //List<int> Dislupa2 = new List<int>() { 0,1,2,4,10 };
                    //newFunctions_4.LupaDinamicaDiscriminatoria(rootSession, data2, Dislupa2);
                }   
            }
            else
            {
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                //ElementList = rootSession.FindElementsByClassName("TDBEdit");
                ElementList[3].Clear();
                if (motor == "SQL")
                {
                    ElementList[3].SendKeys(crudVars[18]);
                }
                if (motor == "ORA")
                {
                   ElementList[3].SendKeys(crudVars[8]);
                }
            }
        }
    }
}