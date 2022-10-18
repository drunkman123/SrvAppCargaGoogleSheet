# SrvAppCargaGoogleSheet
## Descrição do Projeto
### O que é?
Trata-se de um serviço para fazer GET em banco de dados SQL e em API's na internet e subir para o Google Sheet através do uso da API do Google.
### Tecnologias utilizadas
C# .NET core 2.1, SQL, API's da internet, API da Google
### Desafios e Futuro
Por ser meu primeiro projeto da vida com programação, tive muita dificuldade com a implementação da API do google (que inclusive precisa de atualização, pois está utilizando modelo antigo através de certificado).
```
String serviceAccountEmail = "xxxxxxxxxx@xxxxxx.com";

                var certificate = new X509Certificate2(@"xxxxx\xxxxxxx.p12", "xxxxxxx", X509KeyStorageFlags.Exportable);


                ServiceAccountCredential credential = new ServiceAccountCredential(
                   new ServiceAccountCredential.Initializer(serviceAccountEmail)
                   {
                       Scopes = new[] { SheetsService.Scope.Spreadsheets }
                   }.FromCertificate(certificate));

                // Create the service.
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    
                    HttpClientInitializer = credential,
                    ApplicationName = "xxxxxxxxxx",
                });
```
Outro ponto que também precisa de complemento, é a parte de solicitar dados de um api externa que traz muitos dados em páginas distintas, porém as páginas só incrementam dados. Em outro serviço eu fiz o tratamento próprio para dar carga em banco de dados, sendo que neste projeto, em vez de puxar da API, deve-se puxar do banco de dados para maior agilidade.
```
var retorno = new List<xxxxxxxxx>();
            using (var client = new HttpClient())
            {
                var request = new HttpRequestMessage(HttpMethod.Get, "https://xxxxxxxxxxxxxxxxxxxxxxx/" + tipo);

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

                            var request2 = new HttpRequestMessage(HttpMethod.Get, "https://xxxxxxxxxxxxxxxxxxxxxxx/" + tipo + "/" + pag_atu);
                            var response2 = client.SendAsync(request2).Result;
                            if (response.IsSuccessStatusCode)
                            {
                                var content2 = response2.Content.ReadAsStringAsync().Result;

                                var ObjOrcResp2 = JsonConvert.DeserializeObject<xxxxxxxxx>(content2);
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
```
O using de conexão HttpClient não causa socket exhaustion pois sao feitas poucas solicitações por dia, o ideal seria apenas uma conexão aberta e apenas sua reutilização.
```
using (var client = new HttpClient())
            {}
```
### Modo de Uso
Basta chamar o método Start() na controller que entrará em loop eterno, mesmo com erros(que estão sendo salvos), o programa só irá parar caso a máquina onde esteja hospedada, seja desligada ou o serviço parado
