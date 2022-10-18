using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Cargas.Data.Repositories;
using System;
using System.Threading;
using System.Net;
using Microsoft.AspNetCore.Mvc.Filters;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text;
using System.Collections.Generic;
using Cargas.Data.Models.NEO;
using Newtonsoft.Json;

namespace SrvAppCargaGoogleSheet.Controllers
{
    public static class Globals
    {
        public static Int32 a = 0;
        public static Int32 b = 0; //contador para operacional
        public static Int32 c = 0; //contador para planilha neo
    }

    //a fim de evitar 2 requisições, foi implementado o trecho de código abaixo para retornar um 503 caso venha outra requisição se uma já estiver em loop
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method)]
    public class ExclusiveActionAttribute : ActionFilterAttribute
    {
        private static int _isExecuting = 0;
        private static int _isDuplicated = 0;
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (Interlocked.CompareExchange(ref _isExecuting, 1, 0) == 0)
            {
                base.OnActionExecuting(filterContext);
                return;
            }

            Interlocked.Exchange(ref _isDuplicated, 1);
            filterContext.Result = new StatusCodeResult((int)HttpStatusCode.ServiceUnavailable);
        }

        public override void OnResultExecuted(ResultExecutedContext filterContext)
        {
            base.OnResultExecuted(filterContext);
            if (_isDuplicated == 1)
            {
                Interlocked.Exchange(ref _isDuplicated, 0);
                return;
            }
            Interlocked.Exchange(ref _isExecuting, 0);
        }
    }
    //fim do trecho

    [ExclusiveAction]
    [Route("api/[controller]/[action]")]
    public class CargaGoogleSheetController : Controller
    {
        #region CONSTRUTOR
        private readonly IHostingEnvironment _hostingEnvironment; //setar o environment para salvar os arquivos independente de onde for hosteado

        private readonly CargaGoogleSheetRepository _CargaGoogleSheetRepo;


        public CargaGoogleSheetController(IHostingEnvironment hostingEnvironment, IConfiguration Configuration)
        {
            _hostingEnvironment = hostingEnvironment;
            _CargaGoogleSheetRepo = new CargaGoogleSheetRepository(Configuration, hostingEnvironment);
        }
        #endregion

        public void Start()
        {
            while (1 > 0)
            {
                try
                {
                    Logistica();
                    new ManualResetEvent(false).WaitOne(10000);

                    Pessoas();
                    new ManualResetEvent(false).WaitOne(10000);

                    ManutencaoVTR();
                    new ManualResetEvent(false).WaitOne(10000);

                    Operacional();
                    new ManualResetEvent(false).WaitOne(10000);
                }
                catch (Exception e)
                {
                    _CargaGoogleSheetRepo.GetErro(Convert.ToString(e));
                }
            }

        }

        public void Logistica()
        {

            try
            {

                _CargaGoogleSheetRepo.TestConn(); 
                _CargaGoogleSheetRepo.loginGoogle();
                _CargaGoogleSheetRepo.atualizarLogistica();

            }
            catch (Exception e)
            {
                _CargaGoogleSheetRepo.GetErro(Convert.ToString(e));
                new ManualResetEvent(false).WaitOne(60000);
                Logistica();

            }
           
        }

        public void Pessoas()
        {

            try
            {
                _CargaGoogleSheetRepo.TestConn();
                _CargaGoogleSheetRepo.loginGoogle();
                _CargaGoogleSheetRepo.atualizarPessoa();


            }
            catch (Exception e)
            {
                _CargaGoogleSheetRepo.GetErro(Convert.ToString(e));
                new ManualResetEvent(false).WaitOne(60000);
                Pessoas();

            }
        }

        public void ManutencaoVTR()
        {

            try
            {
                _CargaGoogleSheetRepo.TestConn();
                _CargaGoogleSheetRepo.loginGoogle();
                _CargaGoogleSheetRepo.atualizarManutencao();

            }
            catch (Exception e)
            {
                _CargaGoogleSheetRepo.GetErro(Convert.ToString(e));
                new ManualResetEvent(false).WaitOne(60000);
                ManutencaoVTR();
            }


        }

        public void Operacional()
        {

            //neste métodos tem temporizadores para o método de atualizar o operacional ser executado mais vezes do que os outros
            try
            {
                while (Globals.a < 70)
                {
                    while (Globals.b < 600)
                    {
                        Globals.c++;
                        Globals.b++;
                        new ManualResetEvent(false).WaitOne(1000);
                    }
                    Globals.b = 0;
                    _CargaGoogleSheetRepo.TestConn();
                    _CargaGoogleSheetRepo.loginGoogle();
                    _CargaGoogleSheetRepo.atualizarOperacional();
                    Globals.a++;
                    if (Globals.c == 3600)
                    {
                        _CargaGoogleSheetRepo.TestConn();
                        _CargaGoogleSheetRepo.loginGoogle();
                        ManutencaoVTR();
                        Globals.c = 0;
                    }
                }
                Globals.a = 0;

            }
            catch (Exception e)
            {
                _CargaGoogleSheetRepo.GetErro(Convert.ToString(e));
                new ManualResetEvent(false).WaitOne(60000);
                Operacional();
            }

        }
    }
}
