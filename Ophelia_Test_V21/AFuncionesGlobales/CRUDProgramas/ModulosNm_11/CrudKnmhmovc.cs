
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11
{
    class CrudKnmhmovc : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex

                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            /*for (int i = 0; i < btnTDVI.Count; i++)
            {
                btnTDVI[i].SendKeys(i.ToString());
            }*/

            if (bandera == 0)
            {
                //btnTDVI[4].SendKeys(crudPrincipal[0]);
                //btnTDVI[3].SendKeys(crudPrincipal[1]);
                //btnTDVI[2].SendKeys(crudPrincipal[2]);

                // Ingresar fecha
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 118, 120);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

                // primera lupa crud principal
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 682, 44);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                

                // Segunda lupa crud principal
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 663, 227);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                // Tercera lupa crud principal
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 519, 274);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession3 = null;
                RootSession3 = PruebaCRUD.RootSession();
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                

                // Ingresar Observacion
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 165, 310);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1500);


                
            }
            else
            {
                ////btnTDVI[4].Clear();
                //btnTDVI[3].Clear();
                ////btnTDVI[2].Clear();
                ////btnTDVI[4].SendKeys("394");
                //btnTDVI[3].SendKeys(crudPrincipal[3]);
                ////btnTDVI[2].SendKeys("30/10/2021");
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 165, 310);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);
            }

        }

        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1)
        {
            //var btnTDVIgrid = desktopSession.FindElementsByClassName("TDBGrid");
            var btnTDVIgrid = desktopSession.FindElementsByClassName("TDBGrid");
            if (bandera == 0)
            {
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 4, 592);
                //desktopSession.Mouse.Click(null);
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 661, 43);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDet1[0]);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[2]);
                Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys("1");
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 656, 224);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(3000);
                //rootSession = RootSession();
                //rootSession = ReloadSession(rootSession, "TFrmGnArbol");
                //Thread.Sleep(4000);
                //var Tfrmgnarbol = rootSession.FindElementsByClassName("TVirtualStringTree");
                //Thread.Sleep(4000);
                //rootSession.Mouse.MouseMove(Tfrmgnarbol[0].Coordinates, 109, 8);
                //Thread.Sleep(2000);
                //rootSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //rootSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //rootSession.FindElementByName("Aceptar").Click();
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 629, 30);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(crudDet1[2]);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {

                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 667, 46);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudDet1[3]);

            }

        }

        static public void ClickButton(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.FindElementByName("Cesiones al Contrato").Click();
            }
            else if (bandera == 1)
            {
                desktopSession.FindElementByName("Registro Suspensiones").Click();
            }
            else
            {
                desktopSession.FindElementByName("Adiciones al Contrato").Click();
            }

        }

        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet2)
        {
            var btnTDVIgrid = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 498, 229);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                var btnTDBGrid = desktopSession.FindElementsByClassName("TDBGrid");
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 652, 44);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudDet2[2]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 652, 44);
                desktopSession.Mouse.Click(null);
                //Cesionario
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 652, 44);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession3 = null;
                //RootSession3 = PruebaCRUD.RootSession();
                //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 652, 44);


                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.DoubleClick(null);
                //Thread.Sleep(1500);


                //desktopSession.Mouse.Click(null);
                //WindowsDriver<WindowsElement> rootSession = null;
                //rootSession = RootSession();
                //rootSession = ReloadSession(rootSession, "TCaptura");
                //var TKactusDBgrid = rootSession.FindElementsByClassName("TKactusDBgrid");
                //Thread.Sleep(3000);
                //rootSession.Mouse.MouseMove(TKactusDBgrid[0].Coordinates, 4, 589);
                //Thread.Sleep(3000);
                //rootSession.Mouse.Click(null);
                //rootSession.FindElementByName("Aceptar").Click();
                //Thread.Sleep(3000);
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 339, 31);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(crudDet2[1]);
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 435, 31);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(crudDet2[2]);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 652, 44);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[1]);
                Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }

        }

        static public void AgregarCrud3(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet3)
        {
            var btnTDVIgrid = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 919, 39);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudDet3[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudDet3[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudDet3[2]);
                Thread.Sleep(1000); desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudDet3[3]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudDet3[4]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudDet3[6]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession3 = null;
                RootSession3 = PruebaCRUD.RootSession();
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                WindowsDriver<WindowsElement> rootSession = null;
                rootSession.Mouse.Click(null);
                rootSession.Mouse.Click(null);
                rootSession.FindElementByName("Aceptar").Click();
                Thread.Sleep(1500);
                var btnTButton = desktopSession.FindElementsByClassName("TButton");
                desktopSession.Mouse.MouseMove(btnTButton[0].Coordinates, 35, 9);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.MouseMove(btnTButton[0].Coordinates, 44, 9);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.MouseMove(btnTButton[0].Coordinates, 35, 9);
                desktopSession.Mouse.Click(null);
                //Thread.Sleep(3000);
                //WindowsDriver<WindowsElement> rootSession = null;
                //rootSession = RootSession();
                //rootSession = ReloadSession(rootSession, "TCaptura");
                //var TKactusDBgrid = rootSession.FindElementsByClassName("TKactusDBgrid");
                //Thread.Sleep(3000);
                //rootSession.Mouse.MouseMove(TKactusDBgrid[0].Coordinates, 73, 174);
                //Thread.Sleep(4000);
                //rootSession.Mouse.Click(null);
                //rootSession.FindElementByName("Aceptar").Click();
                //Thread.Sleep(1500);
                //desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 338, 33);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //rootSession = RootSession();
                //rootSession = ReloadSession(rootSession, "TFrmGnArbol");
                //Thread.Sleep(4000);
                //var Tfrmgnarbol = rootSession.FindElementsByClassName("TVirtualStringTree");
                //Thread.Sleep(3000);
                //rootSession.Mouse.MouseMove(Tfrmgnarbol[0].Coordinates, 83, 47);
                //Thread.Sleep(2000);
                //rootSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //rootSession.FindElementByName("Aceptar").Click();
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 695, 23);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudDet3[5]);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession.Mouse.Click(null);
                rootSession.FindElementByName("Aceptar").Click();
                Thread.Sleep(1500);
                rootSession.Mouse.Click(null);
                rootSession.FindElementByName("OK").Click();
                Thread.Sleep(1500);
            }

        }

    }


}
