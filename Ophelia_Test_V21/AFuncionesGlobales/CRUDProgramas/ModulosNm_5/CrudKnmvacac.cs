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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmvacac : FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {

                /*for (int i = 0; i < ElementList.Count; i++)
                {
                    ElementList[i].SendKeys(i.ToString());
                    Thread.Sleep(500);
                }
                Debugger.Launch();*/
                ElementList[7].SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(3000);
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                ElementList[19].SendKeys("15");
                Thread.Sleep(1000);
                ElementList[20].SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[12].SendKeys(data[3]);
                ElementList[13].SendKeys(data[4]);
                ElementList[11].SendKeys(data[5]);
                ElementList[16].SendKeys(data[6]);
                ElementList[14].SendKeys("15");
                //Debugger.Launch();
            }
            else
            {
                ElementList[20].Clear();
                ElementList[20].SendKeys(data[7]);
            }
        }

        public static void Ventana(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

        public static void Preliminar(WindowsDriver<WindowsElement> desktopSession, int tipo, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> Session = null;
            Session = RootSession();
            Session = ReloadSession(Session, "TFrmActividad");
            Thread.Sleep(500);
            if (tipo == 1)
            {
                var ElementList = Session.FindElementByName("Activos");
                Thread.Sleep(2000);
                Session.Mouse.MouseMove(ElementList.Coordinates, 8, 12);
                Session.Mouse.Click(null);
            }
            else if(tipo == 2)
            {
                var ElementList = Session.FindElementByName("Inactivos");
                Thread.Sleep(2000);
                Session.Mouse.MouseMove(ElementList.Coordinates, 11, 13);
                Session.Mouse.Click(null);
            }
            else if (tipo == 3)
            {
                var ElementList = Session.FindElementByName("Pensionados");
                Thread.Sleep(2000);
                Session.Mouse.MouseMove(ElementList.Coordinates, 6, 23);
                Session.Mouse.Click(null);
            }
            else if (tipo == 4)
            {
                var ElementList = Session.FindElementByName("Retirados");
                Thread.Sleep(2000);
                Session.Mouse.MouseMove(ElementList.Coordinates, 9, 13);
                Session.Mouse.Click(null);
            }
            else
            {
                var ElementList = Session.FindElementByName("Todos");
                Thread.Sleep(2000);
                Session.Mouse.MouseMove(ElementList.Coordinates, 8, 14);
                Session.Mouse.Click(null);
            }
            Thread.Sleep(1000);
        }

        public static void Aceptar(WindowsDriver<WindowsElement> Session, string file)
        {
            List<string> errorMessages = new List<string>();
            Session = RootSession();
            Session = ReloadSession(Session, "TFrmActividad");
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Opción Seleccionada", true, file);
            //ACEPTAR
            Session.FindElementByName("Aceptar").Click();

        }


        //-------------------------------------
        //KRlEmple

        public static void CRUDEmple(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TEdit");
            /*int cont = 0;
            foreach (var item in ElementList)
            {
                item.SendKeys(Convert.ToString(cont));
                cont = cont + 1;
            }*/

           /* int cont1 = 0;
            foreach (var item1 in ElementList1)
            {
                item1.SendKeys(Convert.ToString(cont1));
                cont1 = cont1 + 1;
            }*/


            if (tipo == 1)
            {
                LupaDinamica(desktopSession, data);
                desktopSession.FindElementByName("Cédula Ciudadania").Click();
                ElementList[8].SendKeys(data[1]);
                ElementList1[3].SendKeys(data[2]);
                ElementList1[1].SendKeys(data[3]);
                ElementList[5].SendKeys(data[4]);
                // 1 - Datos Básicos
                ClickButtonExternoEmple(desktopSession, 1);
                desktopSession.FindElementByName("Masculino").Click();
                ElementList[22].SendKeys(data[5]);
                ElementList[11].SendKeys(data[6]);
                ElementList[10].SendKeys(data[7]);
                //  2 - Datos Personales 
                ClickButtonExternoEmple(desktopSession, 2);
                var ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList2[18].SendKeys(data[8]);
                // 1 - Int Lugar Residencia 
                ClickButtonInternoEmple(desktopSession, 1);
                var ElementList3 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList3[26].SendKeys(data[9]);
                ElementList3[29].SendKeys(data[10]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[11]);
                VentanaDireccion(desktopSession, data[12]);
                // 2 - Int Otros Datos 
                ClickButtonInternoEmple(desktopSession, 2);
                var ElementList4 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList4[22].SendKeys(data[13]);
                //  3 - Correo Electrónico
                ClickButtonInternoEmple(desktopSession, 3);
                var ElementList5 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList5[20].SendKeys(data[14]); ;
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[15]);
            }
            else
            {
                ElementList[11].Clear();
                ElementList[11].SendKeys(data[2]);
            }

        }

        public static void VentanaDireccion(WindowsDriver<WindowsElement> desktopSession, string data)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> Session = null;
            Session = RootSession();
            Session = ReloadSession(Session, "TFrmDireccion");
            Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
            Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Session.Keyboard.SendKeys(data);
            var ElementList = Session.FindElementByName("Aceptar");
            ElementList.Click(); 
        }

        public static void ClickButtonExternoEmple(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            List<string> errorMessages = new List<string>();
            // 1 - Datos Básicos 2 - Datos Personales 
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Datos Básicos ").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Datos Personales ").Click();
            }

        }

        public static void ClickButtonInternoEmple(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            List<string> errorMessages = new List<string>();
            // 1 - Lugar Residencia 2 - Otros Datos 3 - Correo Electrónico
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Lugar Residencia").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Otros Datos").Click();
            }
            else if (tipo == 3)
            {
                desktopSession.FindElementByName("Correo Electrónico").Click();
            }

        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                if (contLupa == 1)
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TCaptura");
                    Thread.Sleep(1000);
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    campo[1].Clear();
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    ////Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);
                        }
                    }
                    indexVal++;
                }
            }
        }



    }
}
