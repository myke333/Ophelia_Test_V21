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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_8
{
    class CrudKnmarpls : FuncionesVitales
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



        static public void InData(WindowsDriver<WindowsElement> desktopSession, string file, List<string> Principal)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            //in Data Nro banco 
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 44, 84);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            desktopSession.Keyboard.SendKeys(Principal[0]);
            Thread.Sleep(1000);            



            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(1000);
            // in Data Valor Cheque
            desktopSession.Keyboard.SendKeys(Principal[1]);
            Thread.Sleep(1000);


            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(1000);
            desktopSession.Keyboard.SendKeys(Principal[2]);
            Thread.Sleep(1000);


            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(1000);
            desktopSession.Keyboard.SendKeys(Principal[3]);
            Thread.Sleep(1000);


            newFunctions_4.ScreenshotNuevo("Datos ingresado", true, file);

        }



        static public void ClickBtnAceptar(WindowsDriver<WindowsElement> desktopSession)
        {
            desktopSession.FindElementByName("Aceptar").Click();
            Thread.Sleep(1500);
        }



        static public void SaveArchivo(WindowsDriver<WindowsElement> desktopSession)
        {
            // Se ingresa ruta y nombre del archivo
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();

            string Txt = @"C:\Reportes\ReporteKNmArpls_" + "Txt" + "_" + Hora();
            RootSession.Keyboard.SendKeys(Txt);
            Thread.Sleep(1500);


            // Se da enter para guardar archivo
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);

        }



        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        }


    }
}
