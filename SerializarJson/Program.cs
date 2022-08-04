using System;
using System.Collections.Generic;
using System.IO;
using SerializarJson.Model;
using Newtonsoft;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Linq;

namespace SerializarJson
{
    class Program
    {
        static void Main(string[] args)
        {

            // SerializarCabecalhoPlanilha("Base de Clientes");
            //DeserializarObjetos("cabecalhoPlanilhaBaixada.json");

            //List<AtributosPlanilhaBaseClientes> cabecalhos = new List<AtributosPlanilhaBaseClientes>();

            //cabecalhos = DeserilizarListaDeObjetos("cabecalhoPlanilhaBaixada.json");

            TesteSerializarPlanilha();
        }

        #region Serialização

        private static void SerializarUmObjeto(string nomeArquivo)
        {
            Usuario usuario = new Usuario()
            {
                Nome = "Fabiana",
                Sobrenome = "Allana de Paula",
                Email = "ffabianaallanadepaula@runup.com.br",
                Endereco = new Endereco()
                {
                    Logradouro = "Rua 3",
                    Cidade = "Goiânia",
                    Estado = "GO",
                    Bairro = "Água Branca",
                    Numero = "160",
                    Cep = "74723-200"
                }
            };

            using (StreamWriter stream = new StreamWriter(Path.Combine(@"C:\Users\davin\Documents\Serializar", nomeArquivo)))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(stream, usuario);
            }
        }

        private static void SerializarListaDeObjetos(string nomeArquivo)
        {
            RepositorioDeUsuario repositorio = new RepositorioDeUsuario();

            using (StreamWriter stream = new StreamWriter(Path.Combine(@"C:\Users\davin\Documents\Serializar", nomeArquivo)))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(stream, repositorio);
            }
        }

        public static void SerializarCabecalhoPlanilha(string nomeGrafico)
        {

            string path = @$"C:\Users\Starline\Downloads";

            DirectoryInfo diretorio = new DirectoryInfo(path);
            var caminhoPlanilha = diretorio.GetFiles().Where(x => x.Name.StartsWith("Instalados_export_"))
            .Select(x => x.FullName).Last();
            string nomeArquivo = "cabecalhoPlanilhaBaixada.json";

            switch (nomeGrafico)
            {
                case "Base de Clientes":
                    var package = new ExcelPackage(new FileInfo(caminhoPlanilha));
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                    int totalColunas = sheet.Dimension.End.Column;

                    AtributosPlanilhaBaseClientes cabecalho = new AtributosPlanilhaBaseClientes();

                    for (int i = 1; i < totalColunas;)
                    {
                        cabecalho.dt_ativacao = sheet.Cells[1, i].Value.ToString();                        
                        cabecalho.cd_solicitacao = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nm_ocorrencia = sheet.Cells[1, ++i].Value.ToString();                        
                        cabecalho.dt_orcamento_pedido = sheet.Cells[1, ++i].Value.ToString();                        
                        cabecalho.cd_pedido = sheet.Cells[1, ++i].Value.ToString();                        
                        cabecalho.data_agendamento = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nm_canal_venda = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nm_parceiro = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nm_parceiro_estabelecimento = sheet.Cells[1, ++i].Value.ToString();                      
                        cabecalho.nm_vendedor = sheet.Cells[1, ++i].Value.ToString();                        
                        cabecalho.nm_loja = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nm_pessoa = sheet.Cells[1, ++i].Value.ToString();                        
                        cabecalho.nr_cnpj_cpf = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nm_produto = sheet.Cells[1, ++i].Value.ToString();                        
                        cabecalho.ds_placa = sheet.Cells[1, ++i].Value.ToString();                      
                        cabecalho.ds_chassi = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nm_tipo_plataforma = sheet.Cells[1, ++i].Value.ToString();                      
                        cabecalho.nm_marca = sheet.Cells[1, ++i].Value.ToString();                      
                        cabecalho.nm_modelo = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nu_ano_modelo = sheet.Cells[1, ++i].Value.ToString();                      
                        cabecalho.vl_total = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.dt_instalacao = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.nm_seguradora_ics = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.status_servico = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.dt_reativacao = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.dt_desativacao = sheet.Cells[1, ++i].Value.ToString();                        
                        cabecalho.status_pagamento_mensalidade = sheet.Cells[1, ++i].Value.ToString();                       
                        cabecalho.ds_apolice = sheet.Cells[1, ++i].Value.ToString();                     
                        cabecalho.dt_inicio_vigencia = sheet.Cells[1, ++i].Value.ToString();                  
                        cabecalho.dt_final_vigencia = sheet.Cells[1, ++i].Value.ToString();                     
                        cabecalho.vl_comissao_corretor = sheet.Cells[1, ++i].Value.ToString();                      
                        cabecalho.cd_canal_venda = sheet.Cells[1, ++i].Value.ToString();                     
                        cabecalho.fl_ativo = sheet.Cells[1, ++i].Value.ToString();                  
                        cabecalho.dt_desinstalacao = sheet.Cells[1, ++i].Value.ToString();                       

                    }
                    SerializarCabecalho(cabecalho, nomeArquivo);
                    break;

                case "Instalações - Agendamentos":
                    break;
                case "Instalações - Pagamentos":
                    break;
                case "Cotações Mensais":
                    break;

            }

        }

        private static void TesteSerializarPlanilha()
        {

            string path = @$"C:\Users\Starline\Downloads";

            DirectoryInfo diretorio = new DirectoryInfo(path);
            var fullPath = diretorio.GetFiles().Where(x => x.Name.StartsWith("Instalados_export_"))
            .Select(x => x.FullName).First();

            var package = new ExcelPackage(new FileInfo(fullPath));

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet sheet = package.Workbook.Worksheets["data"];
            var totalLinhas = sheet.Dimension.End.Row;
            var totalColunas = sheet.Dimension.End.Column;

            AtributosPlanilhaBaseClientes planilhaLinha;

            List<AtributosPlanilhaBaseClientes> planilhaCompleta = new List<AtributosPlanilhaBaseClientes>();

            string nomeArquivo = "planilhaBaixada.json";

            // primeira linha é o cabecalho


            for (int l = 2; l <= totalLinhas; l++)
            {
                planilhaLinha = new AtributosPlanilhaBaseClientes();

                for (int c = 0; c < totalColunas;)
                {
                    planilhaLinha.dt_ativacao = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.cd_solicitacao = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_ocorrencia = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.dt_orcamento_pedido = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.cd_pedido = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.data_agendamento = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_canal_venda = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_parceiro = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_parceiro_estabelecimento = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_vendedor = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_loja = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_pessoa = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nr_cnpj_cpf = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_produto = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.ds_placa = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.ds_chassi = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_tipo_plataforma = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_marca = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_modelo = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nu_ano_modelo = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.vl_total = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.dt_instalacao = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.nm_seguradora_ics = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.status_servico = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.dt_reativacao = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.dt_desativacao = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.status_pagamento_mensalidade = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.ds_apolice = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.dt_inicio_vigencia = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.dt_final_vigencia = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.vl_comissao_corretor = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.cd_canal_venda = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.fl_ativo = sheet.Cells[l, ++c].Value.ToString();
                    planilhaLinha.dt_desinstalacao = sheet.Cells[l, ++c].Value.ToString();

                }
                planilhaCompleta.Add(planilhaLinha);

            }

            Serializar(planilhaCompleta, nomeArquivo);

        }   

        public static void Serializar(List<AtributosPlanilhaBaseClientes> objeto, string nomeArquivo)
        {
            using (StreamWriter stream = new StreamWriter(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo)))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(stream, objeto);
            }
        }

        public static void SerializarCabecalho(AtributosPlanilhaBaseClientes objeto, string nomeArquivo)
        {
            using (StreamWriter stream = new StreamWriter(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo)))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(stream, objeto);
            }
        }

        #endregion

        #region Desserialização


        public static AtributosPlanilhaBaseClientes DeserializarObjetos(string nomeArquivo)
        {
            AtributosPlanilhaBaseClientes objeto = null;

            using (StreamReader stream = new StreamReader(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo)))
            {
                string jsonString = stream.ReadToEnd();
                objeto = JsonConvert.DeserializeObject<AtributosPlanilhaBaseClientes>(jsonString);
            }
            return objeto;
        }
        public static List<AtributosPlanilhaBaseClientes> DeserilizarListaDeObjetos(string nomeArquivo)
        {
            List<AtributosPlanilhaBaseClientes> objetos = null;
            using (StreamReader stream = new StreamReader(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo)))
            {
                string jsonString = stream.ReadToEnd();
                objetos = JsonConvert.DeserializeObject<List<AtributosPlanilhaBaseClientes>>(jsonString);
            }
            return objetos;
        }

        #endregion
    }
}
