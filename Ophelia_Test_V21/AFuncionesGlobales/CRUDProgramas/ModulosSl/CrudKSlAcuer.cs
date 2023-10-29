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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl
{
	class CrudKSlAcuer : FuncionesVitales
	{
        //Conectar
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

        static public WindowsDriver<WindowsElement> ReloadSession2(WindowsDriver<WindowsElement> session, string parametro)
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

        static public void CrudAcuer(WindowsDriver<WindowsElement> desktopSession, string tipo, List<string>crudPrincipal, string file)
		{
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");

			switch (tipo)
			{
                case ("0"):
                    Elementlist[4].SendKeys(crudPrincipal[0]);
                    Elementlist[4].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[5].SendKeys(crudPrincipal[1]);
                    Elementlist[5].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[3].SendKeys(crudPrincipal[2]);
                    Elementlist[3].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[2].SendKeys(crudPrincipal[3]);
                    break;
                case ("1"):
                    Elementlist[5].Clear();
                    Elementlist[5].SendKeys(crudPrincipal[4]);
                    Elementlist[3].Clear();
                    Elementlist[3].SendKeys(crudPrincipal[5]);
                    Elementlist[2].Clear();
                    Elementlist[2].SendKeys(crudPrincipal[6]);
                    break;
			}


/*

            int cont = 0;

            foreach (var item in Elementlist)
            {
                if (cont == 3)
                {
                    item.SendKeys("2");
                }
                else
                {
                    item.SendKeys(Convert.ToString(cont));

                }
                cont = cont + 1;

                Thread.Sleep(5000);
            }
*/

        }

         static public void BotonesReportes(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnPrueba = desktopSession.FindElementsByClassName("TPanel");


            if (bandera == 0)
            {
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }
            /*
                    if (bandera == 1)
                    {

                        desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 100, 124);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);

                        //RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 318, -54);
                        //RootSession.Mouse.Click(null);
                        //Thread.Sleep(2000);

                    }
                    if (bandera == 2) //ya usada
                    {
                        desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 300, 124);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);

                    }
                    if (bandera == 3)// add crud
                    {
                        desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1428, 472);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);

                    }
                    if (bandera == 4) //aplicar
                    {
                        desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1485, 472);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);

                    }
                    if (bandera == 5) //edit
                    {
                        desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1455, 472);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);

                    }
                    if (bandera == 6) //cancel
                    {
                        desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1512, 472);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);

                    }
             */
            if (bandera == 1) //preliminar consulta 
            {
                desktopSession.Mouse.MouseMove(btnPrueba[0].Coordinates, 115, -50);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession3 = null;
                RootSession3 = PruebaCRUD.RootSession();
                RootSession3 = ReloadSession(RootSession3, "TFrmReports");

                //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(2000);
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }
            if (bandera == 2) //preliminar consulta calendario
            {
                desktopSession.Mouse.MouseMove(btnPrueba[0].Coordinates, 122, 50);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession3 = null;
                RootSession3 = PruebaCRUD.RootSession();
                RootSession3 = ReloadSession(RootSession3, "TFrmReports");
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(2000);
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }
        }

        public static List<string> GenerarReportes(string TipoReporte, string file, string ruta, string nomPrograma = null, string codProgram = null)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession;
            string ReportePdf = "";
            string ReporteW = "";
            string ReporteE = "";
            string Ruta1 = "";
            string Extension1 = "";
            string Ruta2 = "";
            string Extension2 = "";
            string Ruta3 = "";
            string Extension3 = "";
            bool Pdf = false;
            bool Word = false;
            bool Excel = false;

            // Revisar que se abrio el preview
            bool IsDisplayedReview = false;

            try
            {
                for (int i = 0; i < 20; i++)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                        var validar = rootSession.FindElementsByClassName("TfrxPreviewForm");
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                    }
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

                int count = 0;
                while (count < 300)
                {
                    try
                    {
                        List<Point> carga = CoordinatesFinder(rootSession, 224, 32, 64);
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                        count++;
                    }
                }

                List<Point> coordenadas = CoordinatesFinder(rootSession, 247, 223, 99);
                int CoordX = coordenadas[0].X;
                int CoordY = coordenadas[0].Y;

                var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));
                bool IsDisplayedReport = false;

                // Clic en el icono de editar reporte
                rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                rootSession.Mouse.Click(null);
                Thread.Sleep(8000);
                newFunctions_4.ScreenshotNuevo("La opción 'Editar Reporte' " + TipoReporte + " está activada", true, file);

                errorMessages.Add("El Preview del reporte " + TipoReporte + " tiene habilitado la opcion 'Editar Reporte'");

                // Cerrar la ventana de editar Reporte
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxDesignerForm");

                var CerrarEditar = rootSession.FindElementByName("Cerrar");
                CerrarEditar.Click();
                Thread.Sleep(2000);
                newFunctions_4.ScreenshotNuevo("Ventana 'Editar Reporte' " + TipoReporte + " Cerrada", true, file);
            }
            catch
            {

            }

            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

            var ElementsPreview = rootSession.FindElementsByClassName("TToolBar");

            if (ElementsPreview.Count > 0)
            {
                IsDisplayedReview = true;
            }
            else
            {
                errorMessages.Add("No se puede visualizar la ventana Preview del Reporte " + TipoReporte);
            }

            string fecha = Hora();
            if (IsDisplayedReview == true)
            {
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                int cerrar = 0;
                try
                {
                    Thread.Sleep(2000);
                    List<Point> coordenadas1 = CoordinatesFinder(rootSession, 224, 32, 64);
                    int CoordX = coordenadas1[0].X;
                    int CoordY = coordenadas1[0].Y;

                    var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));

                    // Clic en el icono de exportar PDF
                    rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TfrxPDFExportDialog");

                    var ExportPDF = rootSession.FindElementByName("Aceptar");

                    if (ExportPDF != null)
                    {
                        //Clic en Aceptar exportar PDF
                        ExportPDF.Click();
                        Thread.Sleep(2000);
                        rootSession = RootSession();

                        try
                        {
                            //ingresar la ruta
                            rootSession.Keyboard.SendKeys(ruta);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                            //Abrir PDF
                            Process.Start(ruta + ".pdf");
                            Thread.Sleep(6000);
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                            //var TraerPantalla = rootSession.FindElementsByName("AVL_AVView");
                            newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " exportado a PDF 1", true, file);
                            LimpiarProcesos();
                        }
                        catch
                        {
                            var Ok = rootSession.FindElementByName("OK");
                            Ok.Click();
                            Thread.Sleep(1000);
                            errorMessages.Add("No existe el botón exportar reporte a PDF");
                        }
                    }
                }
                catch
                {
                    errorMessages.Add("En el Preview  del Reporte " + TipoReporte + " NO Existe el Botón Exportar Reporte a PDF");
                    cerrar = 1;
                }

                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                Thread.Sleep(1000);
                //Debugger.Launch();
                //if (codProgram != "KNmNoema")
                //{
                //    //Clic en Disquet
                //    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 57, 17);
                //    rootSession.Mouse.Click(null);
                //    Thread.Sleep(2000);

                //    //Reportes
                //    rootSession = RootSession();
                //    rootSession = ReloadSession(rootSession, "#32768");
                //    ////Debugger.Launch();
                //    var MenuItem = rootSession.FindElementsByXPath("/Menu[@Name=\"Contexto\"][@ClassName=\"#32768\"]/*");
                //    if (MenuItem.Count == 0)
                //    {
                //        //MenuItem = rootSession.FindElementsByXPath("/Menu[@ClassName='#32768'][@Name='Context']/*");
                //        MenuItem = rootSession.FindElementsByXPath("//Menu[contains(@ClassName,#32768)][contains(@Name,Context)]");
                //        //Desktop 1
                //    }
                //    for (int i = 0; i < MenuItem.Count; i++)
                //    {
                //        if (MenuItem[i].GetAttribute("Name").Contains("PDF"))
                //        {
                //            Pdf = true;
                //            ReportePdf = MenuItem[i].GetAttribute("Name");
                //            Ruta1 = @"C:\Reportes\Reporte_PDF_" + TipoReporte + "_" + fecha;
                //            Extension1 = ".pdf";
                //        }

                //    }

                //    if (Pdf == true)
                //    {
                //        try
                //        {
                //            ExportarReporte(ReportePdf, Ruta1, Extension1, TipoReporte, file);
                //        }
                //        catch
                //        {
                //            errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a PDF Falló");
                //        }
                //    }                                       

                //}


                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                Thread.Sleep(2500);

                if (cerrar == 0)
                {
                    //Cerrar Preview
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 600, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    try
                    {
                        ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                        //Cerrar Preview
                        rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        try
                        {
                            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Cerrar").Coordinates);
                            rootSession.Mouse.Click(null);
                            errorMessages.Add("El Botón 'Cerrar' no Existe en el Preview del Reporte " + TipoReporte);
                        }
                        catch
                        {

                        }
                    }
                    catch
                    {

                    }
                }
                else
                {
                    //Cerrar Preview
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 520, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    try
                    {
                        ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                        //Cerrar Preview
                        rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        try
                        {
                            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Cerrar").Coordinates);
                            rootSession.Mouse.Click(null);
                            errorMessages.Add("El Botón 'Cerrar' no Existe en el Preview del Reporte " + TipoReporte);
                        }
                        catch
                        {

                        }
                    }
                    catch
                    {

                    }
                }
            }
            return errorMessages;
        }

        private static List<Point> CoordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = FuncionesVitales.Hora();
            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            int Romper = 0;

            for (int x = 0; x < bmpSource.Width; x++)
            {
                for (int y = 0; y < bmpSource.Height; y++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    ////Debugger.Launch();
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 20 || y > lasty + 20)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            Romper++;
                            break;
                        }
                    }
                }
                if (Romper != 0)
                {
                    break;
                }
            }
            ////Debugger.Launch();
            return coordenadas;
        }


        public static List<string> generarExcel(WindowsDriver<WindowsElement> Session, string file, string codProgram, string ruta)
        {
            List<string> errorMessages = new List<string>();
            try
            {
                Session = RootSession();
                Session = ReloadSession2(Session, "TFrmExcelex");
                var element = Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelEx')]");
                if (element[0] != null)
                {
                    bool IsDisplayedExcel = false;
                    string Fecha = Hora();
                    Thread.Sleep(500);
                    if (element[0] != null)
                    {
                        IsDisplayedExcel = true;
                    }
                    else
                    {
                        errorMessages.Add($"No se puede visualizar la ventana Exportar a excel");
                    }

                    if (IsDisplayedExcel)
                    {
                        Session = RootSession();
                        Session = ReloadSession2(Session, "TFrmExcelEx");
                        var ReportElements = Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelEx')]");
                        Session.Mouse.MouseMove(ReportElements[0].Coordinates, 222, 207);
                        Session.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Exportar datos Excel", true, file);
                        var buttonAccept = Session.FindElementByName("Aceptar");
                        buttonAccept.Click();

                        Thread.Sleep(5000);
                        try
                        {
                            Session = RootSession();
                            Session = ReloadSession2(Session, "XLMAIN");
                            if (Session != null)
                            {
                                int count = 0;
                                while (count < 240)
                                {
                                    try
                                    {
                                        //Session = RootSession();
                                        //Session = ReloadSession1(Session, "XLMAIN");
                                        Thread.Sleep(500);
                                        var saveExcel1 = Session.FindElementsByName("Maximizar");
                                        if (saveExcel1.Count > 0)
                                        {
                                            break;
                                        }
                                    }
                                    catch
                                    {
                                        Thread.Sleep(1000);
                                        count++;
                                    }
                                }
                                var saveExcel = Session.FindElementsByName("Maximizar");
                                var cantidad = saveExcel.Count;
                                if (cantidad == 2)
                                {
                                    saveExcel[1].Click();
                                }
                                else
                                {
                                    saveExcel[0].Click();
                                }
                                //Cambio Jose
                                Session.FindElementByName("Pestaña Archivo").Click();
                                //Fin Cambio Jose

                                //Session.FindElementByName("Guardar").Click();
                                Session.FindElementByName("Guardar").Click();
                                Session.FindElementByName("Examinar").Click();
                                Thread.Sleep(1000);
                                Session.Keyboard.SendKeys(ruta);
                                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                                LimpiarProcesos();

                            }
                            else
                            {
                                //Debugger.Launch();
                                Thread.Sleep(1000);
                                Session = RootSession();
                                Session = ReloadSession2(Session, "Shell_TrayWnd");
                                Thread.Sleep(1000);
                                var venExcel = Session.FindElementByName("Excel - 1 ventana de ejecución");
                                if (venExcel != null)
                                {
                                    venExcel.Click();
                                    Thread.Sleep(4000);
                                    int count = 0;
                                    while (count < 240)
                                    {
                                        try
                                        {
                                            Session = RootSession();
                                            Session = ReloadSession2(Session, "XLMAIN");
                                            Thread.Sleep(500);
                                            var saveExcel1 = Session.FindElementsByName("Maximizar");
                                            if (saveExcel1.Count > 0)
                                            {
                                                break;
                                            }
                                        }
                                        catch
                                        {
                                            Thread.Sleep(1000);
                                            count++;
                                        }
                                    }
                                    var saveExcel = Session.FindElementsByName("Maximizar");
                                    var cantidad = saveExcel.Count;
                                    if (cantidad == 2)
                                    {
                                        saveExcel[1].Click();
                                    }
                                    else
                                    {
                                        saveExcel[0].Click();
                                    }
                                    //Cambio Jose
                                    Session.FindElementByName("Pestaña Archivo").Click();
                                    //Fin Cambio Jose

                                    //Session.FindElementByName("Guardar").Click();
                                    Session.FindElementByName("Guardar").Click();
                                    Session.FindElementByName("Examinar").Click();
                                    Thread.Sleep(1000);
                                    Session.Keyboard.SendKeys(ruta);
                                    Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                                    LimpiarProcesos();
                                }
                            }

                        }
                        catch
                        {
                            Thread.Sleep(1000);
                            /*string navigator = AbrirPrograma.SelectNavigator();
                            if (navigator == "Edge")
                            {
                                newFunctions_4.ScreenshotNuevo("Opción de Excel no trae registros", true, file);
                                PruebaCRUD.VentanaAzul(Session);
                                Thread.Sleep(2000);
                            }*/
                        }

                        if (Session == null)
                        {
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Opción de Excel no trae registros", true, file);
                            Thread.Sleep(1000);
                            Session = RootSession();
                            PruebaCRUD.VentanaAzul(Session);
                            Thread.Sleep(500);
                        }

                    }
                }
            }
            catch
            {
                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errorMessages.Add($"Error al intentar generar el excel");
            }

            return errorMessages;
        }


        public static List<Point> coordinatesFinder(WindowsDriver<WindowsElement> session)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = "ProgramaExcel";
            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            int romper = 0;
            for (int y = 0; y < bmpSource.Height; y++)
            {
                for (int x = 0; x < bmpSource.Width; x++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    if (clrPixel.R == 151 && clrPixel.G == 179 && clrPixel.B == 143)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 10 || y > lasty + 10)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            romper = 1;
                            break;
                        }
                    }
                }
                if (romper == 1)
                {
                    break;
                }
            }
            return coordenadas;
        }


        public static List<string> ExpExcel(WindowsDriver<WindowsElement> desktopSession, string banExcel, string file, string codProgram, string ruta)
        {
            List<string> errorMessages = new List<string>();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(10),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            try
            {
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;
                    List<Point> coordenadas = coordinatesFinder(desktopSession);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                    string navigator = AbrirPrograma.SelectNavigator();
                    //EDEGE
                    var ExcelOption1 = desktopSession.FindElementsByClassName("TKNavegador");

                    //Pasar datos
                    var pasDat = PruebaCRUD.NavClass(desktopSession);
                    //var pasDat = desktopSession.FindElementsByClassName("TFrmExcelEx");




                    if (navigator == "IE")
                    {
                        //IE
                        string ClickButtonExcel = "//Pane[contains(@ClassName, 'TKNavegador')]/..";
                        ExcelOption1 = desktopSession.FindElementsByXPath(ClickButtonExcel);
                        Thread.Sleep(1000);

                        //WindowsDriver<WindowsElement> RootSession3 = null;
                        //RootSession3 = PruebaCRUD.RootSession();
                        //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        //Thread.Sleep(2000);
                        //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        //Thread.Sleep(2000);

                        //rootSession = RootSession();
                        //rootSession = ReloadSession1(rootSession, "TForm");
                        //Thread.Sleep(1000);

                        //rootSession.FindElementByName("Exportación estándar").Click();
                        //Thread.Sleep(1000);
                        //newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                        //Thread.Sleep(1000);
                        //rootSession.FindElementByName("Aceptar").Click();
                        //Thread.Sleep(4000);

                    }



                    if (ExcelOption1.Count > 0)
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Opciones Exportar Excel", true, file);


                        bool optionExcel = false;
                        int cont = 0;
                        string clase = "";
                        if (banExcel == "0")
                        {
                            WindowsDriver<WindowsElement> RootSession3 = null;
                            RootSession3 = PruebaCRUD.RootSession();
                            RootSession3 = ReloadSession(RootSession3, "TFrmExcelEx");

                            //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            //Thread.Sleep(2000);


                            desktopSession.Mouse.MouseMove(pasDat[0].Coordinates, 810, 406);
                            desktopSession.Mouse.DoubleClick(null);
                            //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            newFunctions_4.ScreenshotNuevo("Exportacion Excel Maestro", true, file);

                           
                            Thread.Sleep(2000);

                            Thread.Sleep(1000);
                            try
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                                rootSession = RootSession();
                            }
                            catch
                            {
                                PruebaCRUD.VentanaAzul(desktopSession);
                                rootSession = RootSession();
                            }
                        }

                        else if (banExcel == "1")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession2(rootSession, "TForm");
                        }
                        else if (banExcel == "2")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession2(rootSession, "TFrmMenuExcel");
                        }
                        else if (banExcel == "3")
                        {

                            rootSession = RootSession();
                            rootSession = ReloadSession2(rootSession, "TFrmMenExpo");
                        }
                        else if (banExcel == "4")
                        {

                            rootSession = RootSession();
                            rootSession = ReloadSession2(rootSession, "TFrmTipExpo");
                        }

                        if (banExcel == "1" || banExcel == "2" || banExcel == "3" || banExcel == "4")
                        {
                            ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                            var contButton = rootSession.FindElementsByClassName("TGroupButton");
                            int cont2 = contButton.Count() - 1;
                            rootSession.FindElementsByClassName("TGroupButton");
                            rootSession.FindElementsByClassName("TGroupButton")[cont2].Click();
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            try
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                            }
                            catch
                            {
                                Thread.Sleep(2000);
                                PruebaCRUD.VentanaAzul(desktopSession);
                                Thread.Sleep(2000);
                            }
                            int i;
                            //Debugger.Launch();
                            for (i = cont2 - 1; i > -1; i--)
                            {
                                ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                                desktopSession.Mouse.Click(null);
                                if (banExcel == "1")
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                else
                                {
                                    rootSession.FindElementsByClassName("TGroupButton")[cont2 - 1].Click();
                                }
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                try
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                }
                                catch
                                {
                                    Thread.Sleep(7000);
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    Thread.Sleep(2000);
                                }
                                cont2 = cont2 - 1;
                            }
                        }

                        if (banExcel == "5")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession2(rootSession, "TForm");
                            Thread.Sleep(1000);
                            rootSession.FindElementByName("Exportación estándar").Click();
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                            Thread.Sleep(1000);
                            rootSession.FindElementByName("Aceptar").Click();
                            Thread.Sleep(4000);
                            rootSession = RootSession();
                            rootSession = ReloadSession2(rootSession, "XLMAIN");
                            Thread.Sleep(3000);
                            //Debugger.Launch();
                            if (rootSession == null)
                            {
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("No hay registros para exportar o el generar excel", true, file);
                                Thread.Sleep(2000);
                                PruebaCRUD.VentanaAzul(desktopSession);
                                Thread.Sleep(2000);
                            }
                            else
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                            }
                            for (int i = 1; i < 16; i++)
                            {
                                //Clic Excel
                                ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                                desktopSession.Mouse.Click(null);
                                Thread.Sleep(1000);
                                rootSession = RootSession();
                                rootSession = ReloadSession2(rootSession, "TForm");
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Exportación estándar").Click();
                                Thread.Sleep(1000);
                                for (int j = 0; j < i; j++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Aceptar").Click();
                                Thread.Sleep(4000);
                                rootSession = RootSession();
                                rootSession = ReloadSession2(rootSession, "XLMAIN");
                                if (rootSession == null)
                                {
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("No hay registros para exportar o el generar excel", true, file);
                                    Thread.Sleep(2000);
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    Thread.Sleep(2000);
                                }
                                else
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                }

                            }
                        }

                    }
                    return rootSession != null;
                });
            }

            catch
            {
                Debug.WriteLine("No se encuentra la opcion de excel para este programa o no hay registros para exportar");
            }
            //string pathImg = "C:\\imagenesTest\\Programa.Png";
            //if (File.Exists(pathImg))
            //{
            //    File.Delete(pathImg);
            //}
            return errorMessages;
        }

    }
}
