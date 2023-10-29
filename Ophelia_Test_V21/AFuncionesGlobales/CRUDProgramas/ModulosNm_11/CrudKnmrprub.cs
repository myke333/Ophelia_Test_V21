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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
    class CrudKnmrprub : FuncionesVitales
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

        static public void AgregarNomina(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 330, 49);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys("01/10/2018");
                Thread.Sleep(1500);
            }
            else
            {
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 199, 284);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys("Prueba de calidad automatizada");
                Thread.Sleep(1500);
            }

        }

        static public void qbe2(WindowsDriver<WindowsElement> desktopSession)
        {


            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            //Coordenadas boton qbe2
            //string filterqb = "13";

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 910, 334);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);



            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");



            var btnTDVI2 = RootSession.FindElementsByClassName("TTabSheet");

            Thread.Sleep(1500);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }

        static public void preliminar1(WindowsDriver<WindowsElement> desktopSession)
        {


            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            //Coordenadas boton qbe2
            //string filterqb = "13";

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 660, 331);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);



            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmImpRep");



            var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");


            Thread.Sleep(1500);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }

        static public void preliminar2(WindowsDriver<WindowsElement> desktopSession)
        {


            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            //Coordenadas boton qbe2
            //string filterqb = "13";

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 660, 331);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);



            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmImpRep");



            var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");

            Thread.Sleep(1500);
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 145, 111);
            RootSession.Mouse.Click(null);
            Thread.Sleep(1500);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }


        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }




    }
}
