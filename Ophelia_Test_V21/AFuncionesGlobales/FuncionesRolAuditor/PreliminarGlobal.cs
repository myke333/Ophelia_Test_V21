using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosRl;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosCo;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn;

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesRolAuditor
{
    class PreliminarGlobal : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string className)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(className);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                ////Debugger.Launch();
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        public static List<Point> coordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
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
            int romper = 0;
            for (int y = 0; y < bmpSource.Height; y++)
            {
                for (int x = 0; x < bmpSource.Width; x++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 10 || y > lasty + 10)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            romper++;
                            break;
                        }
                    }
                }
                if (romper != 0)
                {
                    break;
                }
            }
            return coordenadas;
        }

        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string file, string pdf1, string nomProgram , string codProgram, string TipoQbe, string QbeFilter,string motor)
        {
            //Debugger.Launch();
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> crudVariables = null;
            try
            {
                //////CASOS PRELIMINAR DIFERENTE
                switch (codProgram)
                {                    
                    case "KNmDisrf":
                        BanderaPreli = "0";
                        for (int i = 0; i <= 1; i++)
                        {
                            errors = CrudKnmdisrf.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        BanderaPreli = "3";
                        break;

                    case "KNmRdncc":
                        errors = CrudKnmrdncc.PreliminarKNmrdncc(desktopSession, BanderaPreli, file, pdf1, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        BanderaPreli = "3";
                        break;

                    case "KNmAusen":
                        BanderaPreli = "0";
                        for (int i = 0; i < 2; i++)
                        {
                            errors = CrudKnmausen.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i + 1);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                            Thread.Sleep(1000);
                        }
                        BanderaPreli = "3";
                        break;

                    case "KNmRetce":
                        BanderaPreli = "0";
                        crudVariables = null;
                        for (int i = 0; i <= 2; i++)
                        {
                            errors = CrudKnmretce.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i, crudVariables);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        break;

                    case "Krlrhvex":
                        BanderaPreli = "0";
                        crudVariables = new List<string>() { "1/12/2020", "3/12/2020" };
                        for (int i = 0; i <= 1; i++)
                        {
                            errors = CrudKrlrhvex.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i, crudVariables);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        BanderaPreli = "3";
                        break;

                    case "KSoINfac":
                        BanderaPreli = "0";
                        crudVariables = new List<string>() { };
                        for (int i = 0; i <= 2; i++)
                        {
                            errors = CrudKsoinfac.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i, crudVariables);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        BanderaPreli = "3";
                        break;

                    case "KSoMaleg":
                        BanderaPreli = "0";
                        for (int i = 0; i <= 4; i++)
                        {
                            errors = CrudKsomaleg.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        BanderaPreli = "3";
                        break;

                    case "KSoAnvln":
                        BanderaPreli = "0";
                        for (int i = 0; i <= 2; i++)
                        {
                            errors = CrudKsoanvln.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        BanderaPreli = "3";
                        break;

                    default:
                        break;
                }

                //////FUNCION PRELIMINAR
                errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                ////Debugger.Launch();
                //////CASOS PRELIMINAR DIFERENTE
                switch (codProgram)
                {
                    case "KAcPlant":
                        errors = CrudKacplant.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                        break; //jhon beltran

                    case "KNmCuent":
                        errors = CrudKnmcuent.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                        break;

                    case "KNmPazys":
                        errors = CrudKnmpazys.PreliminarKnmpazys(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                        break;

                    case "KNmPoyca":
                        BanderaPreli = "0";                        
                        errors = CrudKnmpoyca.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                        break;

                    case "KNmPr23m":
                        BanderaPreli = "0";
                        errors = CrudKnmpr23m.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmPropa":
                        BanderaPreli = "0";
                        for (int i = 0; i <= 1; i++)
                        {
                            errors = CrudKnmpropa.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        break;

                    case "KNmProrr":
                        BanderaPreli = "0";
                        for (int i = 0; i <= 1; i++)
                        {
                            errors = CrudKnmprorr.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i, crudVariables);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        break;

                    case "KNmRcret":
                        string excel = @"C:\Reportes\ReporteExcel_" + "Preliminar" + "_" + Hora();
                        for (int i = 1; i <= 3; i++)
                        {
                            errors = CrudKnmrcret.PreliminarKNmrcret(i, pdf1, excel, file);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                            Thread.Sleep(2000);
                            if (i != 3)
                            {
                                errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                Thread.Sleep(2000);
                            }
                        }
                        break;

                    case "KNmPreno":
                        errors = CrudKnmpreno.ReportePreliminarIfnotOptions(desktopSession, pdf1, file, "Desprendible Standar", false);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmHcuen":
                        errors = CrudKnmhcuen.PreliminarKNmhcuen(desktopSession, BanderaPreli, file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmTprot":
                        errors = CrudKnmtprot.PreliminarKnmtprot(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmLiyca":
                        errors = CrudKnmliyca.PreliminarKNmliyca(desktopSession, BanderaPreli, file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmEmbar":
                        errors = CrudKnmembar.PreliKNmEmbar(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmFopev":
                        rootSession = PruebaCRUD.RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                        var allFrame = rootSession.FindElementsByClassName("IEFrame");
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 70, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(500);
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmRemba":
                        errors = CrudKnmremba.PreliminarKnmremba(desktopSession, file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmMovSe":
                        for (int i = 0; i <= 2; i++)
                        {
                            errors = CrudKnmmovse.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, i);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        }
                        break;

                    case "KNmRvape":
                        errors = CrudKnmrvape.Preliminar(desktopSession, file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmBases":
                        //CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmAutod":
                        errors = CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmRcert":
                        errors = CrudKnmrcert.PreliminarKNmRcert(desktopSession, file);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmBasli":
                        errors = CrudKnmbasli.PreliminarKNmBasli(desktopSession, file, pdf1, nomProgram);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KNmVacac":
                        for (int i = 0; i < 5; i++)
                        {
                            errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                            Thread.Sleep(2000);
                            if (i >= 1)
                            {
                                Thread.Sleep(1000);
                                CrudKnmvacac.Preliminar(desktopSession, i, file);
                            }
                            CrudKnmvacac.Preliminar(desktopSession, i + 1, file);
                            Thread.Sleep(1000);
                            CrudKnmvacac.Aceptar(desktopSession, file);
                            Thread.Sleep(2000);
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        }
                        break;

                    case "KNmRLidd":
                        Thread.Sleep(500);
                        rootSession = PruebaCRUD.RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                        allFrame = rootSession.FindElementsByClassName("IEFrame");
                        Thread.Sleep(500);
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2)+ 120, (allFrame[0].Size.Height / 2) + 90).Click().Perform();
                        Thread.Sleep(500);
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2)+ 110, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
                        Thread.Sleep(1500);
                        //BLOC DE NOTAS
                        string ruta = @"C:\Reportes\ExportarPreliminar" + "_" + codProgram + "_" + Hora();
                        CrudKnmrlidd.BlocNotas(desktopSession, file, ruta);
                        break;

                    case "KSoMaobs":
                        errors = CrudKsomaobs.PreliKSoMaobs(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KSoMasim":
                        errors = CrudKsomasim.PreliKSoMasim(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KSoViepi":
                        string FechaPreli = "18/08/2020";
                        errors = CrudKsoviepi.PreliKSoViepi(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli, FechaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KSoIndic":
                        errors = CrudKsoindic.PreliKSoIndic(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KSoTipin":
                        errors = CrudKsotipin.CrudKSotipinPreliminar(desktopSession, BanderaPreli, pdf1, file);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KSoSeyde":
                        errors = CrudKsoseyde.ReportePreliminarIfnotOptions(desktopSession, pdf1, file, "Reporte de Señalización", false);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        string pdf2 = @"C:\Reportes\ReportePDF1_" + "Preliminar2" + "_" + Hora();
                        errors = CrudKsoseyde.ReportePreliminarIfnotOptions(desktopSession, pdf2, file, "Reporte de Extintores", false);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    case "KSoPanEv":
                        errors = CrudKsopanev.PreliKSoPanEv(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli, motor);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    //Preliminares Diego

                    case "KAcRpeca":
                        errors = CrudKacrpeca.PreliKAcRpeca(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KAcMreno":
                        errors = CrudKacmreno.PreliKAcMreno(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiEntpr":
                        errors = CrudKbientpr.PreliminarKbientpr(desktopSession, file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiEvalu":
                        errors = KBiEvaluPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiRendo":
                        errors = CrudKbirendo.PreliminarKbirendo(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiEmpdo":
                        errors = KBiEmpdoPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiProdi":
                        errors = KBiProdiPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiExame":
                        errors = KBiExamePreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBISolqu":
                        errors = KBISolquPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiGruim":
                        errors = CrudKbigruim.PreliKBiGruim(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiRegdi":
                        errors = KBiRegdiPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter, motor);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBiEmple":
                        errors = KBiEmplePreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBpPresu":
                        errors = KBpPresuPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBpDequi":
                        errors = CrudKbpdequi.PreliminarKbpdequi(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBpRacev":
                        errors = CrudKbpracev.PreliKBpRacev(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBpVenre":
                        errors = KBpVenrePreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KBpBeven":
                        errors = KBpBevenPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KCoConen":
                        errors = KCoConenPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter, motor);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KCoEncue":
                        errors = CrudKcoencue.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KCoParam":
                        errors = KCoParamPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KEd3Ccom":
                        errors = KEd3CcomPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KEd3Licc":
                        errors = KEd3LiccmPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KEd3LiCa":
                        errors = KEd3LiCaPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KFdLicev":
                        errors = CrudKfdlicev.PreliminarKfdlicev(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KFdPresu":
                        errors = CrudKfdpresu.PreliminarKfdpresu(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KFdNefor":
                        errors = KFdNeforPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter, motor);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KFdEntid":
                        errors = CrudKfdentid.Preliminar(file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KFdPlcur":
                        errors = KFdPlcurPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KGeCpeor":
                        errors = KGeCpeorPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KGnUempr":
                        errors = KGnUemprPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KGnRperm":
                        errors = KGnRpermPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter, motor);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmCompe":
                        errors = CrudKnmcompe.PreliminarKnmcompe(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmCtinc":
                        errors = CrudKnmctinc.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmCtpre":
                        errors = KNmCtprePreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;
                    case "KNmCtrch":
                        errors = CrudKnmctrch.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmCtrre":
                        errors = CrudKnmctrre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmMpfac":
                        errors = CrudKnmmpfac.PreliKNmMpfac(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmMlsec":
                        errors = KNmMlsecPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    case "KNmRrepa":
                        errors = KNmRrepaPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;
                    case "KNmNoema":
                        errors = KNmNoemaPreliminar(desktopSession, BanderaPreli, file, pdf1, codProgram, TipoQbe, QbeFilter, nomProgram);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        break;

                    default:
                        errorMessages.Add($"No se encuentra el Codigo del programa en preliminares Diferentes");
                        break;
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }

        public static List<string> KBiEvaluPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }
                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmImpRep");
                ventana.Click();
                if (i == 1)
                {
                    var bot = rootSession.FindElementsByClassName("TGroupButton");
                    bot[0].Click();
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KBiEmpdoPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmRepOpc");
                ventana.Click();
                if (i == 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KBiProdiPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmImpRep");
                ventana.Click();
                if (i == 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KBiExamePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> preliminarButtons = new List<string>() { "General ", "Grafica", "Próximo Examen", "Asistencia" };
            List<string> Fechas = new List<string>() { "01/01/2010 ", "01/01/2020" };

            for (int i = 0; i < 4; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmBiProxExa");
                ventana.Click();

                ventana.FindElementByName(preliminarButtons[i]).Click();

                if (i > 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(Fechas[0]);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(Fechas[1]);
                }

                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                ventana.FindElementByName("Aceptar").Click();
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KBISolquPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;

            for (int i = 0; i < 3; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmTipRepo");
                ventana.Click();

                if (i >= 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    if (i == 2)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    }
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KBiRegdiPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter, string motor)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 7; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmTipRepo");
                ventana.Click();

                if (i == 6)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }

                if (i == 0)
                {
                    ventana.FindElementByName("Pendientes por Descargos").Click();
                }
                else
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    if (i == 5)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        if (motor == "SQL")
                        {
                            rootSession.Keyboard.SendKeys("1");
                        }
                        else if(motor == "ORA")
                        {
                            rootSession.Keyboard.SendKeys("19876");
                        }

                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                }

                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                Thread.Sleep(3000);
                try
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                    if (rootSession != null)
                    {
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        //break;
                    }
                }
                catch
                {
                    Thread.Sleep(1000);
                }

                if (rootSession == null)
                {
                    rootSession = RootSession();
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("No Existen Datos para Generar el Reporte", true, file);
                    Thread.Sleep(500);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }

            }
            return errorMessages;
        }

        public static List<string> KBiEmplePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmReport");
                ventana.Click();
                if (i == 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                if (i == 0)
                {
                    Thread.Sleep(40000);
                }
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KBpPresuPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }
                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmImpRep");
                ventana.Click();
                if (i == 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KNmNoemaPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter, string nomProgram)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 4; i++)
            {
                if (i > 0)
                {
                    errors = FuncionesGlobales.ReportePreliminar(desktopSession, "1", file, pdf1);
                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                }
                Thread.Sleep(1000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmTipoRepore");
                Thread.Sleep(1000);
                if (i == 0)
                {
                    rootSession.FindElementByName("Reporte preliminar").Click();
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    newFunctions_4.ScreenshotNuevo("Opción Preliminar 1", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(3000);
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram,codPrograma);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                } else if(i == 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    for (int j = 0; j < 4; j++)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    newFunctions_4.ScreenshotNuevo("Opción Preliminar 2", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(3000);
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram, codPrograma);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                } else if (i == 2)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    for (int j = 0; j < 4; j++)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    newFunctions_4.ScreenshotNuevo("Opción Preliminar 3", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(3000);
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram, codPrograma);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                } else if (i == 3)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    for (int j = 0; j < 4; j++)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    newFunctions_4.ScreenshotNuevo("Opción Preliminar 4", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(3000);
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram, codPrograma);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                } else if (i == 3)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    newFunctions_4.ScreenshotNuevo("Opción Preliminar 5", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(3000);
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram, codPrograma);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                }

            }
            return errorMessages;

        }
            public static List<string> KBpVenrePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            List<string> ClickButton = new List<string>() { "Reporte por Área y Cargo de Tipos de Retiro", "Reporte Comparativo por Área y Cargo de Tipos de Retiro" };
            List<string> data = new List<string>() { "01/01/2000", "01/01/2020" };

            WindowsDriver<WindowsElement> rootSession = null;

            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmSeleccion");
            var bot = rootSession.FindElementsByClassName("TGroupButton");
            bot[2].Click();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
            Thread.Sleep(500);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

            for (int i = 0; i < 2; i++)
            {
                pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                Thread.Sleep(500);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmSeleccion");
                rootSession.FindElementByName(ClickButton[i]).Click();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(data[0]);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(data[1]);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Space);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KBpBevenPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> errors = new List<string>();
            List<string> listErrors = new List<string> { };
            List<string> ClickButton = new List<string>() { "Listado", "Agenda", "Asistentes", "Materiales", "Citas", "Presupuesto", "Participantes Sorteo", "Conformación Parejas", "Participación" };
            for (int i = 0; i < 9; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmRepOpc");
                rootSession.FindElementByName(ClickButton[i]).Click();
                if (i == 2)
                {
                    rootSession.FindElementByName("Aceptada").Click();
                }
                if (i == 8)
                {
                    for (int k = 0; k < 5; k++)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    rootSession.Keyboard.SendKeys("01/01/2000");
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys("01/01/2020");
                }
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.FindElementByName("Aceptar").Click();

                for (int j = 0; j < 5; j++)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                        if (rootSession != null)
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                            break;
                        }
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                    }
                }
                if (rootSession == null)
                {
                    PruebaCRUD.VentanaAzul(desktopSession);
                }
            }


            return errorMessages;
        }

        public static List<string> KCoConenPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter, string motor)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> errors = new List<string>();
            List<string> data = new List<string> { "1" };

            for (int i = 0; i < 3; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmOpRep");
                ventana.Click();
                if (i >= 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(500);
                    LupaDinamica(rootSession, data);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KCoParamPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 5; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmMenu");
                ventana.Click();
                if (i >= 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                if (i != 4)
                {
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                }
                if (i == 4)
                {
                    string ruta = @"C:\Reportes\ExportarExcel" + "_" + codPrograma + "_" + Hora();
                    newFunctions_2.generarExcel(rootSession, file, codPrograma, ruta);
                }
            }
            return errorMessages;
        }

        public static List<string> KEd3CcomPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmMenuRpt");
                ventana.Click();
                if (i == 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KEd3LiccmPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 5; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmImpRep");
                ventana.Click();
                if (i == 0)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }
                else if (i >= 1)
                {

                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                }


                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                if (i == 3)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Space);
                    for (int j = 0; j < 5; j++)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                }
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                if (i == 3)
                {
                    Thread.Sleep(5000);
                }
                /*errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }*/
                for (int j = 0; j < 5; j++)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                        if (rootSession != null)
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                            break;
                        }
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                    }
                }
                if (rootSession == null)
                {
                    PruebaCRUD.VentanaAzul(desktopSession);
                    rootSession = RootSession();
                }
            }
            return errorMessages;
        }

        public static List<string> KEd3LiCaPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 4; i++)
            {

                if (i == 0)
                {
                    var rootSession1 = RootSession();
                    rootSession1.FindElementByClassName("TFrmEd3lica");
                    rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }

                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                    var rootSession1 = RootSession();
                    rootSession1.FindElementByClassName("TFrmEd3lica");
                    rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }
                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmMenu");
                ventana.Click();
                if (i >= 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KFdPlcurPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 8; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }
                CrudKfdplcur.Preliminar(desktopSession, i + 1, file);

                if (i != 5)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        try
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                            if (rootSession != null)
                            {
                                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                                break;
                            }
                        }
                        catch
                        {
                            Thread.Sleep(1000);
                        }
                    }
                    if (rootSession == null)
                    {
                        PruebaCRUD.VentanaAzul(desktopSession);
                        rootSession = RootSession();
                    }
                }
                else
                {
                    for (int k = 0; k < 5; k++)
                    {
                        try
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "TFrmExcelex");
                            if (rootSession != null)
                            {
                                string ruta = @"C:\Reportes\ExportarExcel" + "_" + codPrograma + "_" + Hora();
                                newFunctions_2.generarExcel(rootSession, file, codPrograma, ruta);
                                break;
                            }
                        }
                        catch
                        {
                            Thread.Sleep(1000);
                        }
                    }
                    if (rootSession == null)
                    {
                        PruebaCRUD.VentanaAzul(desktopSession);
                        rootSession = RootSession();
                    }
                }


            }
            return errorMessages;
        }

        public static List<string> KGeCpeorPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmTipRepo");
                ventana.Click();

                if (i >= 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                }

                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KGnUemprPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmOpc");

                var ElementList = rootSession.FindElementsByClassName("TGroupButton");
                if (ElementList[i].Text == "Todos")
                {
                    ElementList[i].Click();
                }
                else if (ElementList[i].Text == "Uno Solo")
                {
                    ElementList[i].Click();
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    campo[0].SendKeys("UCalida1");
                }
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KFdNeforPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter, string motor)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> errors = new List<string>();
            var data1 = "5";
            var data2 = "72213600";

            for (int i = 0; i < 3; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }

                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmSeleccion");
                ventana.Click();
                if (i == 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                }
                else if (i == 2)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                if (i == 1)
                {
                    var rootSession1 = RootSession();
                    rootSession1 = ReloadSession(rootSession1, "TFrmFdNxJef");
                    Thread.Sleep(500);
                    if (motor == "SQL")
                    {
                        FuncionesGlobales.QBEQry(rootSession1, TipoQbe, data1, file);
                    }
                    else if (motor == "ORA")
                    {
                        FuncionesGlobales.QBEQry(rootSession1, TipoQbe, data2, file);
                    }
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("Opción QBE Preliminar", true, file);
                    Thread.Sleep(500);
                    rootSession1.FindElementByName("Aceptar").Click();
                }
                Thread.Sleep(1000);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KNmCtprePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            for (int i = 0; i < 4; i++)
            {
                if (i >= 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmImpRep");

                var ElementList = rootSession.FindElementsByClassName("TGroupButton");
                ElementList[i].Click();
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            return errorMessages;
        }



        public static List<string> KGnRpermPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter, string motor)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            string ruta = @"C:\Reportes\ExportarExcel" + "_" + codPrograma + "_" + Hora();
            errors = Crudkgnrperm.CRUD(desktopSession, bandera, file, pdf1, ruta, motor);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            return errorMessages;
        }

        public static List<string> KNmMlsecPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            PruebaCRUD.VentanaAzul(desktopSession);
            Thread.Sleep(500);
            for (int i = 0; i < 2; i++)
            {
                if (i == 1)
                {
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("Error Preliminar Controlado", true, file);
                    Thread.Sleep(500);
                    PruebaCRUD.VentanaAzul(desktopSession);
                    Thread.Sleep(500);
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmRepoSS");

                if (i == 0)
                {
                    rootSession.FindElementByName("Normal").Click();
                    rootSession.FindElementByName("Salud").Click();
                }
                else
                {
                    rootSession.FindElementByName("Corrección").Click();
                    rootSession.FindElementByName("Salud").Click();
                }
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(500);
                rootSession.FindElementByName(" Aceptar").Click();
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }

        public static List<string> KNmRrepaPreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string codPrograma, string TipoQbe, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);

            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

            return errorMessages;
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
