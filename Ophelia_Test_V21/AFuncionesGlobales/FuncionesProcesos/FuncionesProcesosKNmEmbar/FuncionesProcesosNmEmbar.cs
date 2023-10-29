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

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmEmbar
{
    class FuncionesProcesosNmEmbar
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

        static public WindowsDriver<WindowsElement> ReloadSession1(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        static public void AgregarRegistro(WindowsDriver<WindowsElement> desktopSession, string  bandera, List<string> crudPrincipal)
        {
            
            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI3 = desktopSession.FindElementByName("Porcentaje");
            var btnTDVI4 = desktopSession.FindElementByName("Valor Fijo");

            if (bandera == "P")
            {
                //IDENTIFICACION
                desktopSession.Mouse.MouseMove(btnTDVI[8].Coordinates, 50, 8);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);

                //FECHA INICIO 
                desktopSession.Mouse.MouseMove(btnTDVI[3].Coordinates, 29, 4);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

                //CONCEPTO
                desktopSession.Mouse.MouseMove(btnTDVI[10].Coordinates, 14, 7);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);

                //PORCENTAJE
                desktopSession.Mouse.MouseMove(btnTDVI3.Coordinates, 7, 10);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                //VALOR PORCENTAJE
                desktopSession.Mouse.MouseMove(btnTDVI[11].Coordinates, 9, 13);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);

                var btnTDVI7 = desktopSession.FindElementByName("Demandante");

                //PESTAÑA DEMANDANTE
                desktopSession.Mouse.MouseMove(btnTDVI7.Coordinates, 39, 5);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                var btnTDVI8 = desktopSession.FindElementByName("Cédula de Ciudadanía");
                //CEDULA CIUDADANIA
                desktopSession.Mouse.MouseMove(btnTDVI8.Coordinates, 8, 12);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                var btnTDVI10 = desktopSession.FindElementsByClassName("TDBEdit");

                //CEDULA
                desktopSession.Mouse.MouseMove(btnTDVI10[5].Coordinates, 37, 10);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);

                //NOMBRE
                desktopSession.Mouse.MouseMove(btnTDVI10[4].Coordinates, 107, 9);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1000);

                //APELLIDOS
                desktopSession.Mouse.MouseMove(btnTDVI10[3].Coordinates, 70, 9);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(1000);
                var btnTDVI9 = desktopSession.FindElementByName("Juzgado");

                //PESTAÑA JUZGADO
                desktopSession.Mouse.MouseMove(btnTDVI9.Coordinates, 22, 10);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                var btnTDVI11 = desktopSession.FindElementsByClassName("TDBEdit");

                //CODIGO
                desktopSession.Mouse.MouseMove(btnTDVI11[9].Coordinates, 29, 8);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(1000);

                //ENTIDAD
                desktopSession.Mouse.MouseMove(btnTDVI11[5].Coordinates, 47, 12);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(1000);

                ////TIPO CUENTA
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys("Corriente");
                Thread.Sleep(1000);

                ////MUNICIPIO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[10]);
                Thread.Sleep(1000);

                //OFICIO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[11]);
                Thread.Sleep(1000);

                //FECHA OFICIO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[12]);
                Thread.Sleep(1000);

                //BENEFICIARIO
                desktopSession.Mouse.MouseMove(btnTDVI11[22].Coordinates, 47, 7);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[13]);
                Thread.Sleep(1000);

                ////CUENTAS
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[14]);
                Thread.Sleep(1000);

                //NUMERO EXPEDIENTE
                desktopSession.Mouse.MouseMove(btnTDVI11[7].Coordinates, 99, 9);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[15]);
                Thread.Sleep(1000);
                //NUMERO PROCESOS
                desktopSession.Mouse.MouseMove(btnTDVI11[3].Coordinates, 30, 11);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[16]);
                Thread.Sleep(1000);
                //TIPO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                Thread.Sleep(1000);

            }



            if (bandera == "VF")
            {
                //IDENTIFICACION
                desktopSession.Mouse.MouseMove(btnTDVI[8].Coordinates, 50, 8);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);

                //FECHA INICIO 
                desktopSession.Mouse.MouseMove(btnTDVI[3].Coordinates, 29, 4);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

                //CONCEPTO
                desktopSession.Mouse.MouseMove(btnTDVI[10].Coordinates, 14, 7);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);

                //VALOR FIJO
                desktopSession.Mouse.MouseMove(btnTDVI4.Coordinates, 16, 1);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                //VALOR
                desktopSession.Mouse.MouseMove(btnTDVI[14].Coordinates, 21, 11);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);

                var btnTDVI7 = desktopSession.FindElementByName("Demandante");

                //PESTAÑA DEMANDANTE
                desktopSession.Mouse.MouseMove(btnTDVI7.Coordinates, 39, 5);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                var btnTDVI8 = desktopSession.FindElementByName("Cédula de Ciudadanía");
                //CEDULA CIUDADANIA
                desktopSession.Mouse.MouseMove(btnTDVI8.Coordinates, 8, 12);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                var btnTDVI10 = desktopSession.FindElementsByClassName("TDBEdit");

                //CEDULA
                desktopSession.Mouse.MouseMove(btnTDVI10[5].Coordinates, 37, 10);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);

                //NOMBRE
                desktopSession.Mouse.MouseMove(btnTDVI10[4].Coordinates, 107, 9);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1000);

                //APELLIDOS
                desktopSession.Mouse.MouseMove(btnTDVI10[3].Coordinates, 70, 9);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(1000);
                var btnTDVI9 = desktopSession.FindElementByName("Juzgado");

                //PESTAÑA JUZGADO
                desktopSession.Mouse.MouseMove(btnTDVI9.Coordinates, 22, 10);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                var btnTDVI11 = desktopSession.FindElementsByClassName("TDBEdit");

                //CODIGO
                desktopSession.Mouse.MouseMove(btnTDVI11[9].Coordinates, 29, 8);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(1000);

                //ENTIDAD
                desktopSession.Mouse.MouseMove(btnTDVI11[5].Coordinates, 47, 12);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(1000);

                ////TIPO CUENTA
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys("Corriente");
                Thread.Sleep(1000);

                ////MUNICIPIO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[10]);
                Thread.Sleep(1000);

                //OFICIO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[11]);
                Thread.Sleep(1000);

                //FECHA OFICIO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[12]);
                Thread.Sleep(1000);

                //BENEFICIARIO
                desktopSession.Mouse.MouseMove(btnTDVI11[22].Coordinates, 47, 7);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[13]);
                Thread.Sleep(1000);

                ////CUENTAS
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[14]);
                Thread.Sleep(1000);

                //NUMERO EXPEDIENTE
                desktopSession.Mouse.MouseMove(btnTDVI11[7].Coordinates, 99, 9);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[15]);
                Thread.Sleep(1000);
                //NUMERO PROCESOS
                desktopSession.Mouse.MouseMove(btnTDVI11[3].Coordinates, 30, 11);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[16]);
                Thread.Sleep(1000);
                //TIPO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                Thread.Sleep(1000);


            }

        }

        //static public void CorrerNomina(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        //{

        //    var btnTDVI2 = desktopSession.FindElementByName("Aceptar");
        //    var btnTDVI = desktopSession.FindElementsByClassName("TEdit");
        //    var btnTDVI1 = desktopSession.FindElementByName("Simulación");
        //    //Debugger.Launch();
        //    //for (int i = 0; i < btnTDVI.Count; i++)
        //    //{
        //    //    btnTDVI[i].SendKeys(i.ToString());
        //    //    Thread.Sleep(10);
        //    //}


        //    if (bandera == 0)
        //    {
        //        //Primero obtenemos el día actual
        //        DateTime date = DateTime.Now;

        //        //Asi obtenemos el primer dia del mes actual
        //        DateTime oPrimerDiaDelMes = new DateTime(date.Year, date.Month, 1);

        //        //Y de la siguiente forma obtenemos el ultimo dia del mes
        //        //agregamos 1 mes al objeto anterior y restamos 1 día.
        //        DateTime UltimoDiaDelMes = oPrimerDiaDelMes.AddMonths(1).AddDays(-1);

        //        string Findate = UltimoDiaDelMes.ToString("dd/MM/yyyy");

        //        //desktopSession.Mouse.MouseMove(btnTDVI1.Coordinates, 3, 3);
        //        //desktopSession.Mouse.Click(null);
        //        //Thread.Sleep(1500);
        //        desktopSession.Mouse.MouseMove(btnTDVI[1].Coordinates, 64, 14);
        //        Thread.Sleep(1500);
        //        desktopSession.Mouse.Click(null);
        //        Thread.Sleep(1500);
        //        btnTDVI[1].Clear();
        //        Thread.Sleep(1500);
        //        desktopSession.Keyboard.SendKeys(Findate);
        //        Thread.Sleep(1000);
        //        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
        //        Thread.Sleep(1000);
        //        desktopSession.Keyboard.SendKeys(Findate);
        //        Thread.Sleep(1000);
        //        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
        //        Thread.Sleep(1000);


        //    }
        //    else
        //    {
        //        desktopSession.Mouse.MouseMove(btnTDVI2.Coordinates, 53, 14);
        //        desktopSession.Mouse.Click(null);
        //        Thread.Sleep(2000);

        //    }
        //}

        //static public void VerificarPrenomina(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        //{

        //    var btnTDVI = desktopSession.FindElementByName("GridPanel1");
        //    var btnTDVI1 = desktopSession.FindElementByClassName("TKNavegador");



        //    if (bandera == 0)
        //    {

        //        desktopSession.Mouse.MouseMove(btnTDVI.Coordinates, 437, 22);
        //        Thread.Sleep(1500);
        //        desktopSession.Mouse.Click(null);
        //        Thread.Sleep(1500);

        //    }
        //    else
        //    {
        //        desktopSession.Mouse.MouseMove(btnTDVI1.Coordinates, 69, 12);
        //        Thread.Sleep(1500);
        //        desktopSession.Mouse.Click(null);
        //        Thread.Sleep(1500);
        //    }
        //}




        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }
    }
}
