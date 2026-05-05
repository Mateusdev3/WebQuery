using ClosedXML.Excel;
using System.Net;
using System.Net.Http.Json;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace Program
{
    public class JsonConfig
    {
        public Query[]? Queries { get; set; }
        public string? Tokenid { get; set; }
        public string UrlCheckToken { get; set; }
        public int? ForceDelay { get; set; }
    }

    public class Licences
    {
        public string? Ip { get; set; }
        public string? Name { get; set; }
    }

    public class Program
    {
        private static HttpClient httpClient = new HttpClient()
        {
            Timeout = TimeSpan.FromSeconds(10)
        };

        public static string user = string.Empty;
        public static string folderResult = "Resultados";

        public static async Task Main(string[] args)
        {
            string mainDrectory = Environment.CurrentDirectory;
           
            

            Console.Title = "WebQuery";
            string configPath = System.IO.Path.Combine(mainDrectory, "config.json");

            Directory.CreateDirectory(folderResult);
            if (!File.Exists(configPath))
            {
                var jsonConfig = new JsonConfig
                {
                    Tokenid = "",
                    UrlCheckToken = "",
                    ForceDelay = 0
                };
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                };
                string jsonString = JsonSerializer.Serialize(jsonConfig, options);
                string filename = "config.json";
                File.WriteAllText(filename, jsonString);
            }

            user = await GetUser();

            if (user != null || user != "Error")
            {
                await Initialize();
            }
            
        }
        private static async Task<bool> CheckToken(HttpClient httpClient)
        {
            try
            {
                string filename = "config.json";
                string jsonString = File.ReadAllText(filename);
                var jsonConfig = JsonSerializer.Deserialize<JsonConfig>(jsonString);
                if (jsonConfig != null)
                {
                    string? jsonToken = jsonConfig.Tokenid;

                    httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", jsonToken);
                    string url = "https://citgisnext.sitbus.com.br:9998/citgis-service-bhz/citgis/parametrosSistema/buscaCoordenadaPadrao";
                    using HttpResponseMessage response = await httpClient.GetAsync(url);

                    string body = await response.Content.ReadAsStringAsync();
                    if (body.Contains("pam_gmaps_coordenada_padrao"))
                    {
                        return true;
                    }
                    return false;
                }
                return false;
            }
            
            catch
            {
                return false;
            }
        }

        private static async Task<string> DeleteQuery(string queryname)
        {
            try
            {
                string filename = "config.json";
                var json = File.ReadAllText(filename);
                var config = JsonSerializer.Deserialize<JsonConfig>(json);
                if (config != null)
                {

                    if (queryname == null | queryname == "")
                    {
                        return "Nome inváido";
                    }

                    if (config.Queries == null)
                    {
                        return "Não há queries para excluir";
                    }

                    config.Queries = config.Queries.Where(q => q.Name != queryname).ToArray();
                    var options = new JsonSerializerOptions { WriteIndented = true };
                    string updatedJsonString = JsonSerializer.Serialize(config, options);
                    File.WriteAllText(filename, updatedJsonString);

                    return "Query deletada com sucesso";
                }

                return "Não há queries para excluir";
            }
            catch
            {
                return "Query não encontrada";
            }
        }

        private static async Task<string> GetIp()
        {
            try
            {
                foreach (var ip in await Dns.GetHostAddressesAsync(Dns.GetHostName()))
                {
                    if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        return ip.ToString();
                    }
                    else
                    {
                        return "Error";
                    }
                }
                return "Error";
            }
            catch
            {
                return "Error";
            }
        }

        private static async Task<string> GetUser()
        {
            try
            {
                string ips = await GetIp();
                string url = "https://raw.githubusercontent.com/Mateusdev3/WebQueryLicence/refs/heads/main/licences.json";
                using HttpResponseMessage response = await httpClient.GetAsync(url);

                string body = await response.Content.ReadAsStringAsync();
            
                Task.Delay(1000).Wait();
                var jsonDoc = JsonSerializer.Deserialize<User>(body);

                

                if(jsonDoc != null)
                {
                    
                    string? user = jsonDoc?.Licences?.FirstOrDefault(l => l.Ip == ips)?.Name;



                    if (!body.Contains(ips))
                    {
                        Console.Clear();
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine("Licença não encontrada para este dispositivo... ");
                        Task.Delay(5000).Wait();
                        Environment.Exit(0);
                    }


                    if(user != null)
                    {
                        return user;
                    }
                    else
                    {
                        return "null";
                    }
                    
                }

                return "null";

            }
            catch
            {
                Console.Clear();
                Console.WriteLine("Erro ao buscar licenca... ");
                return "Error";
            }
        }
        private static async Task Initialize()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("██╗    ██╗███████╗██████╗  ██████╗ ██╗   ██╗███████╗██████╗ ██╗   ██╗");
            Console.WriteLine("██║    ██║██╔════╝██╔══██╗██╔═══██╗██║   ██║██╔════╝██╔══██╗╚██╗ ██╔╝");
            Console.WriteLine("██║ █╗ ██║█████╗  ██████╔╝██║   ██║██║   ██║█████╗  ██████╔╝ ╚████╔╝ ");
            Console.WriteLine("██║███╗██║██╔══╝  ██╔══██╗██║▄▄ ██║██║   ██║██╔══╝  ██╔══██╗  ╚██╔╝  ");
            Console.WriteLine("╚███╔███╔╝███████╗██████╔╝╚██████╔╝╚██████╔╝███████╗██║  ██║   ██║   ");
            Console.WriteLine(" ╚══╝╚══╝ ╚══════╝╚═════╝  ╚══▀▀═╝  ╚═════╝ ╚══════╝╚═╝  ╚═╝   ╚═╝   ");
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("────────────────────────────────────────────────────────────────────────");
            Console.Write("Usuário: ");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(user);
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("────────────────────────────────────────────────────────────────────────");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("USE AS OPÇÕES ABAIXO PARA NAVEGAR");
            Console.WriteLine("                                                                         ");
            Console.WriteLine("[1] Realizar busca salva");
            Console.WriteLine("[2] Realizar nova busca");
            Console.WriteLine("[3] Configurações");
            Console.WriteLine("[4] Gerenciar buscas salvas");
            Console.WriteLine("[5] Atualizar Token");
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Escolha uma opção: ");

            var option = Console.ReadKey(true);
            string op = option.Key.ToString();
            switch (op)
            {
                case "D1" or "NumPad1":

                    await RealizeQuerySaved();
                    break;

                case "D2" or "NumPad2":
                    Console.WriteLine("   ");
                    Console.WriteLine("Insira a url da api: ");
                    string? url = Console.ReadLine();
                    if(url?.Length <= 1) { await Initialize(); }
                    Console.WriteLine("Insirao metodo da requisição Exemplo (POST)");
                    string? method = Console.ReadLine();
                    if (method?.Length < 1) { await Initialize(); }
                    Console.WriteLine("Insira o body da requsiição Exemplo ({\"linhaCodExternoSigla\":\"321\"}): ");
                    string? body = Console.ReadLine();
                    if(body?.Length < 1) { await Initialize(); }
                    Console.WriteLine("Insira o nome da planilha com os dados a serem buscados (Tecle enter se não houver): ");
                    string? sheet = Console.ReadLine();
                    Console.WriteLine("Insira o tipo de retorno que a api ira retornar Exemplo (kml)");
                    string? type = Console.ReadLine();
                    if (type?.Length < 1) { await Initialize(); }
               
                    await InitializeQuery(sheet, url, method, body, type);
                    break;

                case "D3" or "NumPad3":
                    await Settings();
                    Console.ReadKey(true);
                    break;
                case "D4" or "NumPad4":
                    await SetDefaultQuery();
                    break;

                case "D5" or "NumPad5":
                    Console.Clear();
                    await Token();
                    break;

                default:
                    await Initialize();
                    break;
            }
        }

        private static async Task RealizeQuerySaved()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine("██╗    ██╗███████╗██████╗  ██████╗ ██╗   ██╗███████╗██████╗ ██╗   ██╗");
            Console.WriteLine("██║    ██║██╔════╝██╔══██╗██╔═══██╗██║   ██║██╔════╝██╔══██╗╚██╗ ██╔╝");
            Console.WriteLine("██║ █╗ ██║█████╗  ██████╔╝██║   ██║██║   ██║█████╗  ██████╔╝ ╚████╔╝ ");
            Console.WriteLine("██║███╗██║██╔══╝  ██╔══██╗██║▄▄ ██║██║   ██║██╔══╝  ██╔══██╗  ╚██╔╝  ");
            Console.WriteLine("╚███╔███╔╝███████╗██████╔╝╚██████╔╝╚██████╔╝███████╗██║  ██║   ██║   ");
            Console.WriteLine(" ╚══╝╚══╝ ╚══════╝╚═════╝  ╚══▀▀═╝  ╚═════╝ ╚══════╝╚═╝  ╚═╝   ╚═╝   ");
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("────────────────────────────────────────────────────────────────────────");
            Console.Write("Usuário: ");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(user);
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("────────────────────────────────────────────────────────────────────────");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Selecione a query dessejadada: ");
            Console.WriteLine(" ");
            string filename = "config.json";
            string fileString = File.ReadAllText(filename);
            var jsonConfig = JsonSerializer.Deserialize<JsonConfig>(fileString);
            int i = 1;

            if(jsonConfig.Queries != null)
            {

                foreach (var config in jsonConfig.Queries)

                {
                    Console.WriteLine($"[{i}] {config.Name}");
                    i++;
                }
                Console.WriteLine("[0] Voltar");

                var key = Console.ReadLine();

                string? query = key.ToString();
                if (query == "0")
                {
                    await Initialize();
                }
                int queryForm = int.Parse(query) - 1;
                try
                {
                    string name = jsonConfig.Queries[queryForm].Name.ToString();
                    await InitSaveQuery(name);
                    await Task.Delay(10000);
                    await Initialize();
                }
                catch
                {
                    Console.WriteLine("Opção inválida");
                    await Task.Delay(1000);
                    await RealizeQuerySaved();
                }
            }
        }


        private static async Task InitializeQuery(string path, string url, string method, string body, string format)
        {
            if (path != "")
            {
                using var workbook = new XLWorkbook(path);
                var workseet = workbook.Worksheet(1);

                var planilha = new List<string>();
                var api = new List<string>();
                var responseApi = new List<Dictionary<string, object>>();
                foreach (var linha in workseet.RowsUsed())
                {
                    string data = linha.Cell(1).GetString();
                    string obj = body.Replace("*", data);

                    string response = await RealizeQuery(httpClient, url, method, obj);

                    bool isbase64 = response.Contains("[");
                  
                    if(isbase64)
                    {
                        var dados = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(response);
                        Console.WriteLine(response);
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("______________________________________________");
                        Console.ResetColor();

                        if (dados != null)
                            responseApi.AddRange(dados);

                        var colunas = responseApi.SelectMany(d => d.Keys).Distinct().ToList();

                        var wb = new XLWorkbook();
                        var ws = wb.Worksheets.Add("Resultado");

                        for (int i = 0; i < colunas.Count; i++)
                        {
                            ws.Cell(1, i + 1).Value = colunas[i];
                        }

                        for (int i = 0; i < responseApi.Count; i++)
                        {
                            for (int j = 0; j < colunas.Count; j++)
                            {
                                if (responseApi[i].ContainsKey(colunas[j]))
                                {
                                    if (responseApi[i].TryGetValue(colunas[j], out var cellObj) && cellObj != null)
                                    {
                                        ws.Cell(i + 2, j + 1).Value = XLCellValue.FromObject(cellObj);
                                    }
                                    else
                                    {
                                        ws.Cell(i + 2, j + 1).Value = XLCellValue.FromObject(string.Empty);
                                    }
                                }
                            }
                            ws.Columns().AdjustToContents();
                            string pathFull = Path.Combine(folderResult.ToString(), $"Resultado.{format}");
                            wb.SaveAs(pathFull);

                            Console.WriteLine($"Info: {i} Salva Com sucesso");
                        }

                        Console.WriteLine("Planilha gerada com sucesso.");
          
                    }
                    else
                    {
                        string dados = response.ToString();
                        dados = dados.Replace("\"", "");

                        dados = ClearResponse64(dados);

                        byte[] bytes = Convert.FromBase64String(dados);

                        string pathFull = Path.Combine(folderResult.ToString(), $"{data}.{format}");

                        File.WriteAllBytes(pathFull, bytes);

                        Console.WriteLine("OK:" + data);
                    }
                }
            }
            else
            {
                Console.WriteLine();
                var responseNobody = await RealizeQuery(httpClient, url, method, body);

                Console.WriteLine(responseNobody);

                await Task.Delay(1000);
                await Initialize();
            }
        }

        private static string ClearResponse64(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            string clear = value.Trim();

            if(clear.StartsWith("\"") && clear.EndsWith("\""))
            {
                try
                {
                    clear = JsonSerializer.Deserialize<string>(clear) ?? clear;

                }
                catch
                {

                }
            }

            clear = clear
             .Replace("\"", "")
            .Replace("\\r", "")
            .Replace("\\n", "")
            .Replace("\r", "")
            .Replace("\n", "")
            .Replace("\\t", "")
            .Replace("\t", "")
            .Replace("\\u003d", "=")
            .Replace("\\u002b", "+")
            .Replace("\\u002f", "/")
            .Trim();

            return clear;

        }

        private static async Task InitSaveQuery(string queryname)
        {
            string filename = "config.json";

            string config = File.ReadAllText(filename);
            var configJson = JsonSerializer.Deserialize<JsonConfig>(config);

            var query = configJson.Queries.FirstOrDefault(q => q.Name == queryname);

            if (query != null)
            {
                await InitializeQuery(query.Sheet, query.Url, query.Method, query.Body, query.ReturnType);
            }
        }
        private static async Task<string> RealizeQuery(HttpClient httpClient, string url, string method, string values)
        {
            try
            {
                string filename = "config.json";
                string jsonString = File.ReadAllText(filename);
                var configjson = JsonSerializer.Deserialize<JsonConfig>(jsonString);
                string token = configjson.Tokenid;
                httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                if (method == "GET")
                {
                    using HttpResponseMessage responseMessage = await httpClient.GetAsync($"{url}/{values}");
                    string body = await responseMessage.Content.ReadAsStringAsync();
                    return body;
                }
                var content = new StringContent(values, System.Text.Encoding.UTF8, "application/json");
                using HttpResponseMessage response = await httpClient.PostAsync(url, content);
                string responseBody = await response.Content.ReadAsStringAsync();
                return responseBody;
            }
            catch
            {
                return "Error";
            }
        }

        private static async Task SetDefaultQuery()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(" ██████╗ ██╗   ██╗███████╗██████╗ ██╗   ██╗");
            Console.WriteLine("██╔═══██╗██║   ██║██╔════╝██╔══██╗╚██╗ ██╔╝");
            Console.WriteLine("██║   ██║██║   ██║█████╗  ██████╔╝ ╚████╔╝ ");
            Console.WriteLine("██║▄▄ ██║██║   ██║██╔══╝  ██╔══██╗  ╚██╔╝  ");
            Console.WriteLine("╚██████╔╝╚██████╔╝███████╗██║  ██║   ██║   ");
            Console.WriteLine(" ╚══▀▀═╝  ╚═════╝ ╚══════╝╚═╝  ╚═╝   ╚═╝   ");

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("──────────────────────────────────────────────────────────");
            Console.WriteLine("INSIRA A OPÇÃO DESEJADA:");
            Console.WriteLine("                        ");
            Console.WriteLine("[1] Salvar nova busca");
            Console.WriteLine("[2] Exibir buscas salvas");
            Console.WriteLine("[3] Excluir buscas salvas");
            Console.WriteLine("[0] Voltar");
            var option = Console.ReadKey(true);
            string op = option.Key.ToString();

            switch (op)
            {
                case "D1" or "NumPad1":
                    Console.WriteLine("──────────────────────────────────────────────────");
                    Console.WriteLine(" ");
                    Console.Write("Defina um nome para a query: ");
                    string? name = Console.ReadLine();
                    if(name.Length < 1) { await SetDefaultQuery(); }
                    Console.WriteLine("Insira a url da api: ");
                    string? url = Console.ReadLine();
                    if(url.Length < 1) { await SetDefaultQuery(); }
                    Console.WriteLine("Insirao metodo da requisição Exemplo (POST):");
                    string? method = Console.ReadLine();
                    if (method.Length < 1) { await SetDefaultQuery(); }
                    Console.WriteLine("Insira o body da requsiição Exemplo ({\"linhaCodExternoSigla\":\"321\"}): ");
                    string? body = Console.ReadLine();
                    if (body.Length < 1) { await SetDefaultQuery(); }
                    Console.WriteLine("Insira o nome da planilha com os dados a serem buscados (Tecle enter se não houver): ");
                    string? sheet = Console.ReadLine();

                    Console.WriteLine("Insira o tipo de retorno da api Exemplo (kml)");
                    string? type = Console.ReadLine();
                    await WriteQuery(name, url, method, body, sheet, type);
                    Console.WriteLine("Query salva");
                    Task.Delay(2000).Wait();
                    await SetDefaultQuery();
                    break;

                case "D2" or "NumPad2":
                    await ShowAllQueries();
                    break;

                case "D3" or "NumPad3":
                    Console.Write("Insira o nome da query a ser excluida: ");
                    string? queryname = Console.ReadLine();
                    string result = await DeleteQuery(queryname);
                    Console.WriteLine(result);
                    Task.Delay(1000).Wait();
                    await SetDefaultQuery();

                    break;

                case "D0" or "NumPad0":
                    await Initialize();
                    break;

                default:
                    await Initialize();
                    break;
            }
        }

        private static async Task Settings()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("███████╗███████╗████████╗████████╗██╗███╗   ██╗ ██████╗ ███████╗");
            Console.WriteLine("██╔════╝██╔════╝╚══██╔══╝╚══██╔══╝██║████╗  ██║██╔════╝ ██╔════╝");
            Console.WriteLine("███████╗█████╗     ██║      ██║   ██║██╔██╗ ██║██║  ███╗███████╗");
            Console.WriteLine("╚════██║██╔══╝     ██║      ██║   ██║██║╚██╗██║██║   ██║╚════██║");
            Console.WriteLine("███████║███████╗   ██║      ██║   ██║██║ ╚████║╚██████╔╝███████║");
            Console.WriteLine("╚══════╝╚══════╝   ╚═╝      ╚═╝   ╚═╝╚═╝  ╚═══╝ ╚═════╝ ╚══════╝");

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("────────────────────────────────────────────────────────────────");
            Console.WriteLine("INSIRA A OPÇÃO DESEJADA: ");
            Console.WriteLine("                         ");
            Console.WriteLine("[1] Editar nome da pasta de resultados");
            Console.WriteLine("[2] Inserir delay a cada requisição");
            Console.WriteLine("[3] Inserir url para validação de token");
            Console.WriteLine("[0] Voltar");
            var option = Console.ReadKey(true);
            string op = option.Key.ToString();

            switch (op)
            {
                case "D1" or "NumPad1":
                    Console.WriteLine("Insira o nome do da pasta desejada: ");
                    string? folder = Console.ReadLine();
                    folderResult = folder;
                    Directory.CreateDirectory(folder);
                    Console.WriteLine("Pasta criada com sucesso!");
                    await Task.Delay(2000);
                    await Settings();
                    break;

                case "D2" or "NumPad2":
                    Console.WriteLine("────────────────────────────────────────────────────────────────");
                    Console.WriteLine(" ");
                    Console.Write("Insira o delay desejado: ");
                    string numberOfdelay = Console.ReadLine();
                    if(numberOfdelay != null || numberOfdelay != " ")
                    {
                        string filePath = "config.json";
                        string configString = File.ReadAllText(filePath);
                        var jsonDes = JsonSerializer.Deserialize<JsonConfig>(configString);

                        int delay = int.Parse(numberOfdelay);

                        if(delay is int)
                        {
                            jsonDes.ForceDelay = delay;
                        }
                        var options = new JsonSerializerOptions { WriteIndented = true };

                        var jsonUpdate = JsonSerializer.Serialize<JsonConfig>(jsonDes, options);
                        File.WriteAllText(filePath, jsonUpdate);

                        Console.WriteLine("Delay Salvo com sucesso.");
                        await Task.Delay(2000);
                        await Settings();
                        
                    } else
                    {
                        await Settings();
                    }
                    break;

                case "D3" or "NumPad3": 
                    Console.WriteLine("────────────────────────────────────────────────────────────────");
                    Console.WriteLine(" ");
                    Console.Write("Digite a url desejada: ");
                    string url = Console.ReadLine();
                    if(url.Length <= 1) { await Settings(); }
                    await WriteTokenUrl(url);
                    Console.WriteLine("Url salva!");
                    await Task.Delay(2000);
                    await Settings();
                    break;

                case "D0" or "NumPad0":
                    await Initialize();
                    break;

                default:
                    await Initialize();
                    break;
            }
        }

        private static async Task WriteTokenUrl(string url)
        {

            string fileName = "config.json";
            string json = File.ReadAllText(fileName);

            var jsonDes = JsonSerializer.Deserialize<JsonConfig>(json);

            jsonDes.UrlCheckToken = url;

            var options = new JsonSerializerOptions { WriteIndented = true };

            var updateJson = JsonSerializer.Serialize<JsonConfig>(jsonDes, options);

            File.WriteAllText(fileName, updateJson);
        }
        private static async Task ShowAllQueries()
        {
            string filename = "config.json";
            string jsonString = File.ReadAllText(filename);
            var configjson = JsonSerializer.Deserialize<JsonConfig>(jsonString);

            if (configjson.Queries == null || configjson.Queries.Length == 0)
            {
                Console.WriteLine("Nenhuma query salva...");
            }
            else
            {
                Console.Clear();
                foreach (var query in configjson.Queries)
                {
                    Console.WriteLine("────────────────────────────────────────────────────────────────────────────────");
                    Console.WriteLine($"Nome da query: {query.Name}");
                    Console.WriteLine($"Url da api: {query.Url}");
                    Console.WriteLine($"Metodo da requisição: {query.Method}");
                    Console.WriteLine($"Body da requisição: {query.Body}");
                    Console.WriteLine($"Nome da planilha fonte: {query.Sheet}");
                    Console.WriteLine($"Tipo de retorno: {query.ReturnType}");
                }
            }
            Console.WriteLine("────────────────────────────────────────────────────────────────────────────────────────");
            Console.WriteLine("                                            ");
            Console.WriteLine("Pressione qualquer tecla para voltar...");
            var key = Console.ReadKey(true);
            if (key.Key != null)
            {
                await SetDefaultQuery();
            }
        }

        private static async Task Token()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("████████╗ ██████╗ ██╗  ██╗███████╗███╗   ██╗");
            Console.WriteLine("╚══██╔══╝██╔═══██╗██║ ██╔╝██╔════╝████╗  ██║");
            Console.WriteLine("   ██║   ██║   ██║█████╔╝ █████╗  ██╔██╗ ██║");
            Console.WriteLine("   ██║   ██║   ██║██╔═██╗ ██╔══╝  ██║╚██╗██║");
            Console.WriteLine("   ██║   ╚██████╔╝██║  ██╗███████╗██║ ╚████║");
            Console.WriteLine("   ╚═╝    ╚═════╝ ╚═╝  ╚═╝╚══════╝╚═╝  ╚═══╝");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("────────────────────────────────────────────");
            Console.WriteLine("INSIRA A OPÇÃO DESEJADA: ");
            Console.WriteLine("                         ");
            Console.WriteLine("[1] Inserir novo token");
            Console.WriteLine("[2] Testar token salvo");
            Console.WriteLine("[0] Sair");
            var option = Console.ReadKey(true);
            string op = option.Key.ToString();
            switch (op)
            {
                case "D1" or "NumPad1":
                    Console.WriteLine("────────────────────────────────────────────");
                    Console.WriteLine(" ");
                    Console.Write("Insira o token desejado: ");
                    string? token = Console.ReadLine();
                    if (token?.Length != 36) { Console.WriteLine("Token inválido... "); Task.Delay(2000).Wait(); await Token(); }
                    ;
                    WriteToken(token);
                    Console.WriteLine("Token atualizado com sucesso!");
                    Task.Delay(1000).Wait();
                    await Token();

                    break;

                case "D2" or "NumPad2":
                    Console.WriteLine("────────────────────────────────────────────");
                    Console.WriteLine(" ");
                    if (await CheckToken(httpClient)) { Console.WriteLine("Token válido!"); }
                    else { Console.WriteLine("Token Inválido!"); }
                    Task.Delay(1000).Wait();
                    await Token();
                    break;

                case "D0" or "NumPad0":
                    await Initialize();
                    break;
                default:
                    await Initialize();
                    break;
            }
        }

        private static async Task WriteQuery(string name, string url, string method, string? body, string? sheet, string type)
        {
            var query = new Query
            {
                Name = name,
                Url = url,
                Method = method,
                Body = body,
                Sheet = sheet,
                ReturnType = type,
            };

            var options = new JsonSerializerOptions { WriteIndented = true };

            string filename = "config.json";
            string jsonString = File.ReadAllText(filename);
            var config = JsonSerializer.Deserialize<JsonConfig>(jsonString);
            config.Queries = config.Queries == null ? new Query[] { query } : config.Queries.Append(query).ToArray();
            string updatedJsonString = JsonSerializer.Serialize(config, options);
            File.WriteAllText(filename, updatedJsonString);
        }
        private static void WriteToken(string token)
        {
            string filename = "config.json";
            string jsonString = File.ReadAllText(filename);
            var jsonConfig = JsonSerializer.Deserialize<JsonConfig>(jsonString);
            jsonConfig.Tokenid = token;
            var options = new JsonSerializerOptions() { WriteIndented = true };
            string updatedJsonString = JsonSerializer.Serialize(jsonConfig, options);
            File.WriteAllText(filename, updatedJsonString);
        }
    }

    public class Query
    {
        public string? Body { get; set; }
        public string? Method { get; set; }
        public string? Name { get; set; }
        public string? Sheet { get; set; }
        public string? Url { get; set; }
        public string? ReturnType { get; set; }
    }

    public class User
    {
        public Licences[]? Licences { get; set; }
    }
}