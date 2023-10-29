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

    class CrudKnmpreac : FuncionesVitales
    {
        
        public static void CRUDKNmPreac(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            
            //Debugger.Launch();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var TDBCmbCodigoBox = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            var TTabSheet = desktopSession.FindElementsByClassName("TTabSheet");

            //Debug.WriteLine("con: " + TDBCmbCodigoBox.Count);
            //for (int i = 0; i <= TDBCmbCodigoBox.Count; i++)
            //{
            //    TDBCmbCodigoBox[i].Click();
            //    Thread.Sleep(500);
            //}

            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3], crudVars[4], crudVars[5], crudVars[6], crudVars[7] };


            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(TTabSheet[0].Coordinates, 1046, 22);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                ElementList[16].SendKeys(crudVars[11]);
                
                //Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 76, 260);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(crudVars[11]);


                Thread.Sleep(1000);
                int cont = 8;
                for (int i = 0; i < TDBCmbCodigoBox.Count; i++)
                {
                    TDBCmbCodigoBox[i].SendKeys(crudVars[cont]);
                    cont += 1;
                }
                


            }
            else
            {                
                ElementList[14].Clear();
                ElementList[14].SendKeys(crudVars[12]);
            }
        }
    }
}