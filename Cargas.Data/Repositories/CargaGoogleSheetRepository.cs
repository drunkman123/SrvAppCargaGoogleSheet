using Cargas.Data.Models;
using Cargas.Data.Models.Logistica;
using Cargas.Data.Models.NEO;
using Cargas.Data.Models.Operacional;
using Google.Apis.Auth.OAuth2;

using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Security.Cryptography.X509Certificates;

using System.Threading;
using System.Linq;
using System.Text;
using System.Diagnostics;
using Cargas.Data.Models.Pessoas;

namespace Cargas.Data.Repositories
{

    public class CargaGoogleSheetRepository

    {
        #region CONSTRUTOR;
        private readonly IHostingEnvironment _hostingEnvironment;
        private readonly string _scDBPad = "";
        ///ocultadas outras strings///


        readonly IConfiguration Configuration;

        public CargaGoogleSheetRepository(IConfiguration Configuration, IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
            this.Configuration = Configuration;
            _scDBPad = Configuration.GetValue<string>("ConnPadrao");
            ///ocultadas outras conexoes///



        }
        #endregion


        //Funçoes NEO
        private List<DadosGeral> GetDadosGeral(string tipo) //função q precisa alterar e pegar do banco em vez da internet
        {
            var retorno = new List<DadosGeral>();
            using (var client = new HttpClient())
            {
                var request = new HttpRequestMessage(HttpMethod.Get, "https://xxxxxxxx.com.br/api/xxxxxxxx/" + tipo);

                var response = client.SendAsync(request).Result;

                if (response.IsSuccessStatusCode)
                {
                    var content = response.Content.ReadAsStringAsync().Result;
                    var ObjOrcResp = JsonConvert.DeserializeObject<DadosGeral>(content);
                    int total_pag = ObjOrcResp.total_paginas;
                    int pag_atu = ObjOrcResp.pagina_atual;
                    while (total_pag > pag_atu || total_pag == 0)
                    {
                        pag_atu++;

                        var request2 = new HttpRequestMessage(HttpMethod.Get, "https://xxxxxxxx.com.br/api/xxxxxxxx/" + tipo + "/" + pag_atu);
                        var response2 = client.SendAsync(request2).Result;
                        if (response.IsSuccessStatusCode)
                        {
                            var content2 = response2.Content.ReadAsStringAsync().Result;

                            var ObjOrcResp2 = JsonConvert.DeserializeObject<DadosGeral>(content2);
                            var lenght = ObjOrcResp2.lista_os.Count();
                            var inic = 0;
                            while (inic < lenght)
                            {
                                ObjOrcResp.lista_os.Add(ObjOrcResp2.lista_os[inic]);
                                inic++;
                            }
                        }


                        if (total_pag == 0)
                        {
                            total_pag = 1;
                        }
                    }
                    retorno.Add(ObjOrcResp);
                }
                else
                {
                    return null;
                }

                return retorno;
            }
        }


        //Funções Pessoas
        private List<ContrefetivoModel> GetEfetivo()

        {
            SqlDataReader reader = null;
            var retorno = new List<ContrefetivoModel>();

            var query = @"SELECT **************
	                            ";

            using (SqlConnection con = new SqlConnection(_scDBPad))
            {
                SqlCommand com = new SqlCommand(query, con);

                try
                {
                    con.Open();
                    reader = com.ExecuteReader();
                    if (reader != null && reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            var ret = new ContrefetivoModel();

                            ret.exemplo = (reader["exemplo"] != DBNull.Value) ? reader["exemplo"].ToString().Trim() : null;


                            retorno.Add(ret);
                        }
                    }
                }
                catch (Exception e)
                {
                    GetErro(Convert.ToString(e));
                    con.Close();
                    new ManualResetEvent(false).WaitOne(60000);
                    //GetEfetivo();
                }
                finally
                {
                    con.Close();
                }
            }

            return retorno;

        }



        //Funções Logística
        private List<Materiais> GetMateriais()
        {
            SqlDataReader reader = null;
            var retorno = new List<Materiais>();

            var query = @"SELECT 
                          *********
                        ";

            using (SqlConnection con = new SqlConnection(_scDBSigpatSql))
            {
                SqlCommand com = new SqlCommand(query, con);
                try
                {
                    con.Open();
                    reader = com.ExecuteReader();
                    if (reader != null && reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            var ret = new Materiais();

                            ret.exemplo = (reader["exemplo"] != DBNull.Value) ? reader["exemplo"].ToString().Trim() : null;
                            retorno.Add(ret);
                        }
                    }
                }
                catch (Exception e)
                {
                    GetErro(Convert.ToString(e));
                    con.Close();
                    new ManualResetEvent(false).WaitOne(60000);
                    //GetMateriais();
                }
                finally
                {
                    con.Close();
                }
            }
            return retorno;


        }


        //Funções Operacional
        private List<USAtiva> GetUSAtiva()
        {

            SqlDataReader reader = null;
            var retorno = new List<USAtiva>();

            var query = @"SELECT  
	                            ************";

            using (SqlConnection con = new SqlConnection(_scDBProdOp))
            {
                SqlCommand com = new SqlCommand(query, con);
                try
                {
                    con.Open();
                    reader = com.ExecuteReader();
                    if (reader != null && reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            var ret = new USAtiva();

                            ret.exemplo = (reader["exemplo"] != DBNull.Value) ? reader["exemplo"].ToString().Trim() : null;


                            retorno.Add(ret);
                        }
                    }
                }
                catch (Exception e)
                {
                    GetErro(Convert.ToString(e));
                    con.Close();
                    new ManualResetEvent(false).WaitOne(10000);
                    //GetUSAtiva();
                }
                finally
                {
                    con.Close();
                }
            }

            return retorno;

        }



        //Funções Teste Conexão
        private bool CheckForInternetConnection() //Provavelmente algo bem antigo, deve necessitar de atualização
        {
            try
            {

                Ping myPing = new Ping();
                String host = "google.com";
                byte[] buffer = new byte[32];
                int timeout = 1000;
                PingOptions pingOptions = new PingOptions();
                PingReply reply = myPing.Send(host, timeout, buffer, pingOptions);
                if (reply.Status == IPStatus.Success)
                {
                    return true;
                }
                else if (reply.Status == IPStatus.TimedOut)
                {
                    return false;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void TestConn()
        {

            var test_con = CheckForInternetConnection();
            while (test_con == false)
            {
                test_con = CheckForInternetConnection();
                new ManualResetEvent(false).WaitOne(1000);
            }
            GC.Collect();
        }

        //API GOOGLE
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "XXXXXXXXXXXXXXX";
        static readonly string SpreadsheetIdPessoas = "XXXXXXXXXXXXXXX";
        static readonly string SpreadsheetIdLogistica = "XXXXXXXXXXXXXXX";
        static readonly string SpreadsheetIdOperacional = "XXXXXXXXXXXXXXX";
        static readonly string SpreadsheetIdManutencaoVTR = "XXXXXXXXXXXXXXX";

        //Abas Pessoas
        static readonly string EfeSheet = "XXXXXXXXXXXXXXX";
        static readonly string LaureaSheet = "XXXXXXXXXXXXXXX";
        static readonly string AlmaSgtSheet = "XXXXXXXXXXXXXXX";
        static readonly string AlmaOfSheet = "XXXXXXXXXXXXXXX";
        static readonly string AfastDsDrSheet = "XXXXXXXXXXXXXXX";
        static readonly string PAFGeralSheet = "XXXXXXXXXXXXXXX";
        static readonly string RestAtivSheet = "XXXXXXXXXXXXXXX";
        static readonly string SISQSESheet = "XXXXXXXXXXXXXXX";

        //Abas Logística
        static readonly string MatSheet = "XXXXXXXXXXXXXXX";
        static readonly string DetPorCodSheet = "XXXXXXXXXXXXXXX";
        static readonly string ComDefVTRSheet = "XXXXXXXXXXXXXXX";
        static readonly string DefVTRDetalSheet = "XXXXXXXXXXXXXXX";
        static readonly string ErivSheet = "XXXXXXXXXXXXXXX";
        static readonly string ResgateSheet = "XXXXXXXXXXXXXXX";
        static readonly string RomOpeSheet = "XXXXXXXXXXXXXXX";

        //Abas Operacional
        static readonly string USAtivaSheet = "XXXXXXXXXXXXXXX";
        static readonly string CompVtrSheet = "XXXXXXXXXXXXXXX";
        static readonly string MapForSheet = "XXXXXXXXXXXXXXX";

        //Abas Manutenção VTR
        static readonly string OSPendenteSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSPendenteObsSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSPendentePecasSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSAprovadaSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSAprovadaObsSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSAprovadaPecasSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSReavaliarSheet = " XXXXXXXXXXXXXXX";
        static readonly string OSReavaliarObsSheet = " XXXXXXXXXXXXXXX";
        static readonly string OSReavaliarPecasSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSAguardandoOrcSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSAguardandoOrcObsSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSAguardandoOrcPecasSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSVeicEntrSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSVeicEntrObsSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSVeicEntrPecasSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSFinalizadaSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSFinalizadaObsSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSFinalizadaPecasSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSPendenteAbClienteSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSPendenteAbClienteObsSheet = "XXXXXXXXXXXXXXX";
        static readonly string OSPendenteAbClientePecasSheet = "XXXXXXXXXXXXXXX";
        static SheetsService service;



        private void LimparPlanilha(string spreadsheet, string Limpar) //FUNÇÃO QUE LIMPA A PLANILHA DO EXCEL
        {
            var RequestBody = new ClearValuesRequest();
            var DeleteRequest = service.Spreadsheets.Values.Clear(RequestBody, spreadsheet, Limpar);
            var DeleteResponse = DeleteRequest.Execute();
            TestConn();
            new ManualResetEvent(false).WaitOne(2000);
        }

        private void AdicionarCabecalho(List<object> cabecalho, string spreadsheet, string Dados)//FUNÇÃO PARA ADICIONAR O CABEÇALHO
        {
            var ValueRange = new ValueRange();
            var ValueRangetmp = new ValueRange();
            ValueRange.Values = new List<IList<Object>> { cabecalho };
            var CabecalhoRequest = service.Spreadsheets.Values.Append(ValueRange, spreadsheet, Dados);
            CabecalhoRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            var Response = CabecalhoRequest.Execute();
            TestConn();
            new ManualResetEvent(false).WaitOne(2000);
        }

        private void UltimaAtualizacao(string spreadsheet, string Atualiza) //FUNÇÃO PARA LANÇAR A ULTIMA ATUALIZAÇÃO DA PLANILHA NUMA CÉLULA PRÉ-DEFINIDA
        {
            var DataAtu = new List<object>() { };
            dynamic ValueRange2 = new ValueRange();
            var DataAtualizacao = DateTime.Now.ToString();
            var StringAtualizacao = "Data Última Atualização";
            ValueRange2.Values = new List<IList<Object>> { };
            DataAtu.Add(StringAtualizacao);
            DataAtu.Add(DataAtualizacao);
            ValueRange2.Values.Add(DataAtu);
            var AppendRequest2 = service.Spreadsheets.Values.Append(ValueRange2, spreadsheet, Atualiza);
            AppendRequest2.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            var AppendResponse2 = AppendRequest2.Execute();
            TestConn();
            new ManualResetEvent(false).WaitOne(2000);
        }

        private void InsereObs(List<DadosGeral> obj, string spreadsheet, string dados)
        {
            var ValueRange = new ValueRange();
            var ObjectList = new List<object>() { };
            ValueRange.Values = new List<IList<Object>> { }; //ESTA LINHA FAZ COM QUE VARIAS LINHAS FOSSEM INSERIDAS COM APENAS 1 REQUEST, NÃO DANDO EXAUSTAO NA API
            var listatotal = new List<object> { };
            var cont1 = 0; //++ se aparecer algum id pela primeira vez
            var cont2 = 0; //escapa do primeiro if
            foreach (var item in obj[0].lista_os)
            {
                ObjectList.Add(item.id);
                if (item.id_status == 4 || item.id_status == 2 || item.id_status == 3 || item.id_status == 10)  //algumas os possuem a mesma OS nesses status, com isso evita de pular observações iguais de OS's diferentes
                {
                    ObjectList.Add(item.dc_observacao);
                    ValueRange.Values.Add(ObjectList);
                }
                else
                {
                    if (item.itens[0].dc_categoria.ToUpper().Trim() != "GUINCHO")
                    {
                        listatotal.Add(item.dc_observacao);
                        ObjectList.Add(item.dc_observacao);
                    }
                    else
                    {
                        listatotal.Add("Guincho");
                        ObjectList.Add("Guincho");
                    }
                    var listadistinta = listatotal.Distinct().ToList();
                    var contagem = listadistinta.Count();
                    if (cont2 == 0)
                    {
                        ValueRange.Values.Add(ObjectList);
                        cont1 = contagem;
                        cont2++;
                    }
                    else
                    {
                        if (contagem != cont1)
                        {
                            ValueRange.Values.Add(ObjectList);
                            cont1++;
                        }
                    }

                }
                ObjectList = new List<object>() { };
            }
            var InfosRequest = service.Spreadsheets.Values.Append(ValueRange, spreadsheet, dados);
            InfosRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            var InfosResponse = InfosRequest.Execute();
            TestConn();
            new ManualResetEvent(false).WaitOne(5000);
        }

        public void loginGoogle()
        {
            try
            {
                TestConn();
                String serviceAccountEmail = "xxxxxxxxxxxxxxxxxxxxxxxx";

                var certificate = new X509Certificate2(@"xxxxxxxx\xxxxxxxxxxx.p12", "xxxxxxxxxxxx", X509KeyStorageFlags.Exportable);


                ServiceAccountCredential credential = new ServiceAccountCredential(
                   new ServiceAccountCredential.Initializer(serviceAccountEmail)
                   {
                       Scopes = new[] { SheetsService.Scope.Spreadsheets }
                   }.FromCertificate(certificate));

                // Create the service.
                service = new SheetsService(new BaseClientService.Initializer()
                {

                    HttpClientInitializer = credential,
                    ApplicationName = "xxxxxxxxxxxx",
                });
            }
            catch (Exception e)
            {
                GetErro(Convert.ToString(e));
                new ManualResetEvent(false).WaitOne(10000);
                loginGoogle();
            }
            GC.Collect();
        }

        public void atualizarPessoa()
        {
            #region Efetivo
            var EfeObj = GetEfetivo();//pegar os dados do banco
            if (EfeObj.Count != 0)
            {
                try
                {
                    //Ranges  Efetivo Atual
                    var Dados = $"{EfeSheet}!A:O";
                    var Atualiza = $"{EfeSheet}!Q1:Q2";
                    var Limpar = $"{EfeSheet}!A:S";

                    //Limpar Planilha  Efetivo Atual
                    LimparPlanilha(SpreadsheetIdPessoas, Limpar);

                    //Adicionar Cabeçalho  Efetivo Atual
                    List<object> cabecalho = new List<object>();
                    string[] lista = { "RE", "Dígito", "Nome", "Nome de Guerra", "Quadro", "Posto/Graduação", "Data Nascimento", "Sexo", "Grande CMD", "OPM Efetiva", "OPM Atual", "Código OPM Atual", "Situação", "Admissão", "Posse" };
                    foreach (string str in lista)
                    {
                        cabecalho.Add(str);
                    }
                    AdicionarCabecalho(cabecalho, SpreadsheetIdPessoas, Dados);
                    cabecalho.Clear();

                    //Adicionar Informações Efetivo Atual
                    dynamic EfeValueRange = new ValueRange();
                    var EfeObjectList = new List<object>() { };
                    EfeValueRange.Values = new List<IList<Object>> { };//ESTA LINHA FAZ COM QUE VARIAS LINHAS FOSSEM INSERIDAS COM APENAS 1 REQUEST, NÃO DANDO EXAUSTAO NA API
                    foreach (var item in EfeObj)
                    {

                        EfeObjectList.Add(item.EXEMPLO);
                        ///continua os itens da model///
                        EfeValueRange.Values.Add(EfeObjectList);
                        EfeObjectList = new List<object>() { };

                    }
                    var EfeInfosRequest = service.Spreadsheets.Values.Append(EfeValueRange, SpreadsheetIdPessoas, Dados);
                    EfeInfosRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    var EfeInfosResponse = EfeInfosRequest.Execute();
                    TestConn();
                    new ManualResetEvent(false).WaitOne(5000);

                    //Última Atualização  Efetivo Atual
                    UltimaAtualizacao(SpreadsheetIdPessoas, Atualiza);
                    EfeObj.Clear();

                }
                catch (Exception e)
                {
                    GetErro(Convert.ToString(e));
                    new ManualResetEvent(false).WaitOne(10000);

                }
            }
            #endregion


        }

        public void atualizarLogistica()
        {

            #region Materiais
            var MatObj = GetMateriais();
            if (MatObj.Count != 0)
            {
                try
                {
                    //Ranges Materiais
                    var Dados = $"{MatSheet}!A:AG";
                    var Atualiza = $"{MatSheet}!AI1:AI2";
                    var Limpar = $"{MatSheet}!A:AJ";

                    //Limpar Planilha Materiais
                    LimparPlanilha(SpreadsheetIdLogistica, Limpar);



                    //Adicionar Cabeçalho Materiais
                    List<object> cabecalho = new List<object>();
                    string[] lista = { "Número Pat", "Descrição", "Especificação", "Marca", "Modelo", "Conta Contábil", "Nº Série", "Ambiente", "Domínio", "Cód Material", "Valor", "Data Inclusão", "Situação", "Estado", "OPM Detentora", "Doc1", "NumDoc1", "Doc2", "NumDoc2", "Doc3", "NumDoc3", "Doc4", "NumDoc4", "Data Movimentação", "Data Alteração", "Quantidade", "RE Detentor Executivo", "FAMDoc1", "FAMNumDoc1", "FAMDoc2", "FAMNumDoc2", "FAMDoc3", "FAMNumDoc3" };
                    foreach (string str in lista)
                    {
                        cabecalho.Add(str);
                    }
                    AdicionarCabecalho(cabecalho, SpreadsheetIdLogistica, Dados);
                    cabecalho.Clear();

                    //Adicionar Informações Materiais
                    dynamic MatValueRange = new ValueRange();
                    var MatObjectList = new List<object>() { };
                    MatValueRange.Values = new List<IList<Object>> { };
                    foreach (var item in MatObj)
                    {

                        MatObjectList.Add(item.EXEMPLO);
                        ///continua os itens da model///
                        MatValueRange.Values.Add(MatObjectList);
                        MatObjectList = new List<object>() { };




                    }
                    var MatInfosRequest = service.Spreadsheets.Values.Append(MatValueRange, SpreadsheetIdLogistica, Dados);
                    MatInfosRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    var MatInfosResponse = MatInfosRequest.Execute();
                    TestConn();
                    new ManualResetEvent(false).WaitOne(5000);

                    //Última Atualização  Materiais
                    UltimaAtualizacao(SpreadsheetIdLogistica, Atualiza);
                    MatObj.Clear();
                }
                catch (Exception e)
                {
                    GetErro(Convert.ToString(e));
                    new ManualResetEvent(false).WaitOne(10000);
                }


            }
            #endregion

          
        }

        public void atualizarOperacional()
        {
            #region Get US Ativa
            var USAtivaObj = GetUSAtiva();
            if (USAtivaObj.Count != 0)
            {
                try
                {

                    //Ranges USAtiva
                    var Dados = $"{USAtivaSheet}!A:K";
                    var Atualiza = $"{USAtivaSheet}!M1:M2";
                    var Limpar = $"{USAtivaSheet}!A:O";

                    //Limpar Planilha USAtiv
                    LimparPlanilha(SpreadsheetIdOperacional, Limpar);

                    //Adicionar Cabeçalho USAtiv
                    List<object> cabecalho = new List<object>();
                    string[] lista = { "CAD", "Patrulha", "Status", "Início Turno", "Fim Turno", "Cabine", "Desc Cabine", "PrtMsvCod", "PTROPMORI", "PrtOpmAtu", "PREFIXO" };
                    foreach (string str in lista)
                    {
                        cabecalho.Add(str);
                    }
                    AdicionarCabecalho(cabecalho, SpreadsheetIdOperacional, Dados);
                    cabecalho.Clear();

                    //Adicionar Informações USAtiv
                    dynamic USAtivaValueRange = new ValueRange();

                    var USAtivaObjectList = new List<object>() { };
                    USAtivaValueRange.Values = new List<IList<Object>> { };
                    foreach (var item in USAtivaObj)
                    {

                        USAtivaObjectList.Add(item.exemplo);
                        ///continua os itens da model///
                        USAtivaValueRange.Values.Add(USAtivaObjectList);
                        USAtivaObjectList = new List<object>() { };

                    }
                    var USAtivaInfosRequest = service.Spreadsheets.Values.Append(USAtivaValueRange, SpreadsheetIdOperacional, Dados);
                    USAtivaInfosRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    var USAtivaInfosResponse = USAtivaInfosRequest.Execute();
                    TestConn();
                    new ManualResetEvent(false).WaitOne(5000);

                    //Última Atualização USAtiv
                    UltimaAtualizacao(SpreadsheetIdOperacional, Atualiza);

                    USAtivaObj.Clear();
                }
                catch (Exception e)
                {
                    GetErro(Convert.ToString(e));
                    new ManualResetEvent(false).WaitOne(10000);
                }


            }
            #endregion


        }

        public void atualizarManutencao()

        {
            #region
            #region Get OS Pendente
            var OSPendenteObj = GetDadosGeral("1");
            if (OSPendenteObj.Count() != 0)
            {
                try
                {
                    //Ranges OSPendente
                    var Dados = $"{OSPendenteSheet}!A:X";
                    var Atualiza = $"{OSPendenteSheet}!Z1:Z2";
                    var Limpar = $"{OSPendenteSheet}!A:AA";

                    //Limpar Planilha OSPendente
                    LimparPlanilha(SpreadsheetIdManutencaoVTR, Limpar);


                    //Adicionar Cabeçalho OSPendente
                    List<object> cabecalho = new List<object>();
                    string[] lista = { "id_os", "id_status", "dc_status", "dt_abertura", "dt_orcamento", "km_veiculo_os", "id_veic", "prefixo", "placa", "km", "ano_fabricacao", "id_unidade", "dc_unidade", "id_sub_unidade", "dc_sub_unidade", "cod_estabelecimento", "cnpj", "dc_estabelecimento", "endereco", "cidade", "estado", "cep", "telefone", "cod_os_unico" };
                    foreach (string str in lista)
                    {
                        cabecalho.Add(str);
                    }
                    AdicionarCabecalho(cabecalho, SpreadsheetIdManutencaoVTR, Dados);
                    cabecalho.Clear();

                    //Adicionar Informações OSPendente
                    dynamic OSPendenteValueRange = new ValueRange();
                    var OSPendenteObjectList = new List<object>() { };
                    OSPendenteValueRange.Values = new List<IList<Object>> { };
                    foreach (var item in OSPendenteObj[0].lista_os)
                    {
                        OSPendenteObjectList.Add(item.exemplo);
                        OSPendenteValueRange.Values.Add(OSPendenteObjectList);
                        OSPendenteObjectList = new List<object>() { };

                    }
                    var OSPendenteInfosRequest = service.Spreadsheets.Values.Append(OSPendenteValueRange, SpreadsheetIdManutencaoVTR, Dados);
                    OSPendenteInfosRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    var OSPendenteInfosResponse = OSPendenteInfosRequest.Execute();
                    TestConn();
                    new ManualResetEvent(false).WaitOne(5000);

                    //Última Atualização OSPendente
                    UltimaAtualizacao(SpreadsheetIdManutencaoVTR, Atualiza);


                }
                catch (Exception e)
                {
                    GetErro(Convert.ToString(e));
                    new ManualResetEvent(false).WaitOne(10000);
                }
            }
            #endregion

            
            #endregion

        }

        public void GetErro(string exception)
        {
            string Folder = _hostingEnvironment.ContentRootPath + @"\Erro\erro" + (DateTime.Now).ToString().Replace("/", "_").Replace(":", ".") + ".txt";
            string erro = exception + "" + DateTime.Now;

            if (!File.Exists(Folder))
            {
                File.WriteAllText(Folder, erro);
            }


        }

    }
}
