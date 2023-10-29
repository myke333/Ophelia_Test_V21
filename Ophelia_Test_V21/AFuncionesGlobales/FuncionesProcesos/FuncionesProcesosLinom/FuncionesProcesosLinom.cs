using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Threading;


namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosLinom
{
    public class FuncionesProcesosLinom
    {
        static public void CorrerNomina(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI2 = desktopSession.FindElementByName("Aceptar");
            var btnTDVI = desktopSession.FindElementsByClassName("TEdit");
            var btnTDVI1 = desktopSession.FindElementByName("Simulación");
         

            if (bandera == 0)
            {
                //Año en curso
                DateTime date = DateTime.Now;
                string ano = date.ToString("yyyy");

                string fecha = crudPrincipal[2] +"/"+ crudPrincipal[1] +"/"+ ano;
                //desktopSession.Mouse.MouseMove(btnTDVI1.Coordinates, 3, 3);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[1].Coordinates, 64, 14);
                Thread.Sleep(1500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                btnTDVI[1].Clear();
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(fecha);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(fecha);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);


            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2.Coordinates, 53, 14);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

            }
        }

        static public void VerificarPrenomina(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementByName("GridPanel1");
            var btnTDVI1 = desktopSession.FindElementByClassName("TKNavegador");



            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI.Coordinates, 437, 22);
                Thread.Sleep(1500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI1.Coordinates, 69, 12);
                Thread.Sleep(1500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
            }
        }
    }
}
