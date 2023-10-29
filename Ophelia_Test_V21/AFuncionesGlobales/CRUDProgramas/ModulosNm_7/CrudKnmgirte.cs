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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_7
{
    class CrudKnmgirte : FuncionesVitales
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


        public static void LupaDinamica2(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            int contadorLupa = 0;
            foreach (Point coord in coordenadas)
            {  
                Console.WriteLine(contadorLupa);
                if (contadorLupa == 1)
                {
                    Thread.Sleep(2000);
                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                    desktopSession.Mouse.Click(null);
                    if (contLupa == 1)
                    {
                        Thread.Sleep(2000);
                        rootSession = RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                        //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                        var campo = rootSession.FindElementsByClassName("TEdit");
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
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

                contadorLupa++;
            }

        }

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> dataLupa, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");

            switch (tipo)
            {
                case ("0"):
                    Elementlist[10].SendKeys(data[0]);
                    Elementlist[9].SendKeys(data[1]);
                    Elementlist[9].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[2].SendKeys(data[2]);
                    //Elementlist[2].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(2000);
                    LupaDinamica2(desktopSession, dataLupa);
                    Elementlist[5].SendKeys(data[3]);
                    Elementlist[3].SendKeys(data[4]);
                    break;
                case ("1"):
                    Elementlist[2].Clear();
                    Elementlist[2].SendKeys(data[5]);
                    //Elementlist[2].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
            }
        }
    }
}