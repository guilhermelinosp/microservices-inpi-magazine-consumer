using MongoDB.Bson;
using MongoDB.Driver;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System.IO.Compression;
using System.Xml;
using Formatting = Newtonsoft.Json.Formatting;

namespace INPI
{
	public static class Program
	{
		public static async Task Main()
		{
			try
			{
				await Console.Out.WriteLineAsync($"{DateTime.Now} - Starting INPI");

				if (!Directory.Exists(Config.DataStorage))
				{
					Directory.CreateDirectory(Config.DataStorage);
				}

				await DonwloadFiles();
				await ExtractFiles();
				await ProcessFiles();
			}
			catch (Exception ex)
			{
				Directory.Delete(Config.DataStorage, true);

				await Console.Error.WriteLineAsync($"{DateTime.Now} - Error Main(): {ex.Message}");
			}
			finally
			{
				Directory.Delete(Config.DataStorage, true);

				await Console.Out.WriteLineAsync($"{DateTime.Now} - Finished INPI");
			}
		}

		private static async Task DonwloadFiles()
		{
			try
			{
				var options = new EdgeOptions();
				options.AddUserProfilePreference("download.default_directory", Config.DataStorage);
				options.AddUserProfilePreference("download.prompt_for_download", false);
				options.AddUserProfilePreference("disable-popup-blocking", "true");
				options.AddArgument("--headless");

				using var driver = new EdgeDriver(options);

				driver.Navigate().GoToUrl($"http://revistas.inpi.gov.br/rpi/");

				var magazineNumber = driver.FindElement(By.XPath(@"/html/body/div[4]/div/table[1]/tbody/tr[2]/td[1]"))
					.Text;

				driver.Navigate()
					.GoToUrl(
						$"http://revistas.inpi.gov.br/txt/RM{magazineNumber}.zip"); // MARCAS
				await Task.Delay(500);
				driver.Navigate()
					.GoToUrl(
						$"http://revistas.inpi.gov.br/txt/P{magazineNumber}.zip"); // PATENTES
				driver.Navigate()
					.GoToUrl(
						$"http://revistas.inpi.gov.br/txt/PC{magazineNumber}.zip"); // PROGRAMA DE COMPUTADOR
				driver.Navigate()
					.GoToUrl(
						$"http://revistas.inpi.gov.br/txt/CT{magazineNumber}.zip"); // CONTRATOS DE TECNOLOGIA
				driver.Navigate()
					.GoToUrl($"http://revistas.inpi.gov.br/txt/DI{magazineNumber}.zip"); // DESENHOS INDUSTRIAIS
				await Task.Delay(10000);

				driver.Quit();
			}
			catch (Exception ex)
			{
				await Console.Error.WriteLineAsync($"{DateTime.Now} - Error DonwloadFiles(): {ex.Message}");
			}
		}

		private static async Task ExtractFiles()
		{
			try
			{
				foreach (var zipFilePath in Directory.EnumerateFiles(Config.DataStorage, "*.zip"))
				{
					var extractPath = Path.Combine(Config.DataStorage, Path.GetFileNameWithoutExtension(zipFilePath));

					using (var zipArchive = ZipFile.OpenRead(zipFilePath))
					{
						zipArchive.ExtractToDirectory(extractPath);
					}

					File.Delete(zipFilePath);

					foreach (var xmlFilePath in Directory.GetFiles(extractPath, "*.xml"))
					{
						var newFilePath = Path.Combine(Config.DataStorage,
							Path.GetFileNameWithoutExtension(zipFilePath) + ".xml");

						if (File.Exists(newFilePath))
						{
							File.Delete(newFilePath);
						}

						File.Move(xmlFilePath, newFilePath);

						await Task.Delay(500);
					}

					Directory.Delete(extractPath, true);
				}
			}
			catch (Exception ex)
			{
				await Console.Error.WriteLineAsync($"{DateTime.Now} - Error ExtractFiles(): {ex.Message}");
			}

			try
			{
				foreach (var xmlFilePath in Directory.GetFiles(Config.DataStorage, "*.xml"))
				{
					var xmlDocument = new XmlDocument();
					xmlDocument.Load(xmlFilePath);

					var jsonFile = JsonConvert.SerializeXmlNode(xmlDocument.SelectSingleNode("revista"),
						Formatting.Indented, true);

					var jsonFilePath = Path.ChangeExtension(xmlFilePath, ".json");

					if (File.Exists(jsonFilePath))
					{
						File.Delete(jsonFilePath);
					}

					await File.WriteAllTextAsync(jsonFilePath, jsonFile);

					await Task.Delay(500);

					File.Delete(xmlFilePath);
				}
			}
			catch (Exception ex)
			{
				await Console.Error.WriteLineAsync($"{DateTime.Now} - Error ConvertFiles(): {ex.Message}");
			}
		}

		private static async Task ProcessFiles()
		{
			try
			{
				foreach (var jsonFilePath in Directory.GetFiles(Config.DataStorage, "*.json"))
				{
					var content = await File.ReadAllTextAsync(jsonFilePath, CancellationToken.None);
					var fileName = Path.GetFileNameWithoutExtension(jsonFilePath);

					var jObject = JObject.Parse(content);

					jObject.SelectTokens("..@inid").ToList().ForEach(t => t.Parent?.Remove());
					jObject.SelectTokens("..@sequencia").ToList().ForEach(t => t.Parent?.Remove());
					jObject.SelectTokens("..@kindcode").ToList().ForEach(t => t.Parent?.Remove());

					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("comentario")).ToList()
						.ForEach(property => property.Replace(new JProperty("texto-complementar", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("concessao")).ToList()
						.ForEach(property => property.Replace(new JProperty("data-concessao", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("dataProtocolo")).ToList()
						.ForEach(property => property.Replace(new JProperty("data-protocolo", property.Value)));

					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("campoAplicacaoLista"))
						.ToList()
						.ForEach(property => property.Replace(new JProperty("campo-aplicacao-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("criadorLista")).ToList()
						.ForEach(property => property.Replace(new JProperty("criador-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("dataCriacao")).ToList()
						.ForEach(property => property.Replace(new JProperty("data-criacao", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("linguagemLista")).ToList()
						.ForEach(property => property.Replace(new JProperty("linguagem-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("tipoProgramaLista")).ToList()
						.ForEach(property => property.Replace(new JProperty("tipo-programa-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("titularLista")).ToList()
						.ForEach(property => property.Replace(new JProperty("titular-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("nome")).ToList()
						.ForEach(property => property.Replace(new JProperty("nome-completo", property.Value)));

					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("cedentes")).ToList()
						.ForEach(property => property.Replace(new JProperty("cedente-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("certificados")).ToList()
						.ForEach(property => property.Replace(new JProperty("certificado-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("cessionarias")).ToList()
						.ForEach(property => property.Replace(new JProperty("cessionaria-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("peticoes")).ToList()
						.ForEach(property => property.Replace(new JProperty("peticao-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("naturezaDocumento")).ToList()
						.ForEach(property => property.Replace(new JProperty("natureza-documento", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("textoObjeto")).ToList()
						.ForEach(property => property.Replace(new JProperty("texto-objeto", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("siglaCategoria")).ToList()
						.ForEach(property => property.Replace(new JProperty("sigla-categoria", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("descricaoMoeda")).ToList()
						.ForEach(property => property.Replace(new JProperty("descricao-moeda", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("valorContrato")).ToList()
						.ForEach(property => property.Replace(new JProperty("valor-contrato", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("formaPagamento")).ToList()
						.ForEach(property => property.Replace(new JProperty("forma-pagamento", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("prazoContrato")).ToList()
						.ForEach(property => property.Replace(new JProperty("prazo-contrato", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("prazoVigenciaPI")).ToList()
						.ForEach(property => property.Replace(new JProperty("prazo-vigencia", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("publicacao-nacional")).ToList()
						.ForEach(property => property.Replace(new JProperty("data-publicacao-nacional", property.Value)));

					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("titulares")).ToList()
						.ForEach(property => property.Replace(new JProperty("titular-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("despachos")).ToList()
						.ForEach(property => property.Replace(new JProperty("despacho-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("classes-vienna")).ToList()
						.ForEach(property => property.Replace(new JProperty("classe-vienna-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("lista-classe-nice")).ToList()
						.ForEach(property => property.Replace(new JProperty("classe-nice-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("sobrestadores")).ToList()
						.ForEach(property => property.Replace(new JProperty("sobrestador-lista", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("@numero")).ToList()
						.ForEach(property => property.Replace(new JProperty("numero", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("@data-deposito")).ToList()
						.ForEach(property => property.Replace(new JProperty("data-deposito", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("@data-vigencia")).ToList()
						.ForEach(property => property.Replace(new JProperty("data-vigencia", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("@apresentacao")).ToList()
						.ForEach(property => property.Replace(new JProperty("apresentacao", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("@natureza")).ToList()
						.ForEach(property => property.Replace(new JProperty("natureza", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("@pais")).ToList()
						.ForEach(property => property.Replace(new JProperty("pais", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("@uf")).ToList()
						.ForEach(property => property.Replace(new JProperty("uf", property.Value)));
					jObject.Descendants().OfType<JProperty>().Where(p => p.Name.Contains("@codigo")).ToList()
						.ForEach(property => property.Replace(new JProperty("codigo", property.Value)));

					var list = new List<BsonDocument>();

					if (fileName.Contains("RM"))
					{
						var magazine = $"{jObject["numero"]} - {jObject["@data"]}";

						foreach (var processo in jObject["processo"]!)
						{
							processo["revista"] = magazine;

							if (processo["titular-lista"] is JObject titularlista)
							{
								var titular = titularlista["titular"];
								if (titular is JArray)
									processo["titular-lista"] = titular;
								if (titular is JObject)
									processo["titular-lista"] = new JArray(titular);

								if (processo["titular-lista"] is JArray titularArray)
									foreach (var titularObj in titularArray)
									{
										if (titularObj["uf"] is JObject uf)
										{
											titularObj["endereco"] = $"{titularObj["pais"]}";
										}
										else
										{
											titularObj["endereco"] = $"{titularObj["pais"]}/{titularObj["uf"]}";
										}
										titularObj["pais"]?.Parent?.Remove();
										titularObj["uf"]?.Parent?.Remove();
									}
							}

							if (processo["classe-nice-lista"] is JObject classenicelista)
							{
								var classenice = classenicelista["classe-nice"];
								if (classenice is JArray)
									processo["classe-nice-lista"] = classenice;
								if (classenice is JObject)
									processo["classe-nice-lista"] = new JArray(classenice);
							}

							if (processo["classe-vienna-lista"] is JObject classeviennalista)
							{
								var classevienna = classeviennalista["classe-vienna"];
								if (classevienna is JArray)
									processo["classe-vienna-lista"] = classevienna;
								if (classevienna is JObject)
									processo["classe-vienna-lista"] = new JArray(classevienna);

								if (processo["classe-vienna-lista"] is JArray classeviennaArray)
									foreach (var classeviennaObj in classeviennaArray)
									{
										classeviennaObj["@edicao"]?.Parent?.Remove();
									}
							}

							if (processo["despacho-lista"] is JObject despacholista)
							{
								var despacho = despacholista["despacho"];
								if (despacho is JArray)
									processo["despacho-lista"] = despacho;
								if (despacho is JObject)
									processo["despacho-lista"] = new JArray(despacho);

								if (processo["despacho-lista"] is JArray despachoArray)
									foreach (var despachoObj in despachoArray)
									{
										var codigo = despachoObj["codigo"]?.ToString().Replace("IPAS", "");
										despachoObj["despacho"] = $"{codigo} - {despachoObj["nome-completo"]}";
										despachoObj["codigo"]?.Parent?.Remove();
										despachoObj["nome-completo"]?.Parent?.Remove();

										if (despachoObj["protocolo"] is JObject protocolo)
										{
											despachoObj["protocolo"]["numero-protocolo"] = protocolo["numero"];
											despachoObj["protocolo"]["data-protocolo"] = protocolo["@data"];
											despachoObj["protocolo"]["codigo-protocolo"] = protocolo["codigo"];

											protocolo["numero"]?.Parent?.Remove();
											protocolo["@data"]?.Parent?.Remove();
											protocolo["codigo"]?.Parent?.Remove();


											if (protocolo["requerente"] is JObject requetente)
											{
												if (requetente["uf"] is JObject uf)
												{
													protocolo["requerente"]["endereco"] = $"{requetente["pais"]}";
												}
												else
												{
													protocolo["requerente"]["endereco"] = $"{requetente["pais"]}/{requetente["uf"]}";
												}
												requetente["pais"]?.Parent?.Remove();
												requetente["uf"]?.Parent?.Remove();
											}
										}


									}
							}

							list.Add(BsonDocument.Parse(processo.ToString()));
						}

						await MongoRepository.Marcas.InsertManyAsync(list);

						list.Clear();
					}
					else
					{
						var magazine = $"{jObject["numero"]} - {jObject["@dataPublicacao"]}";

						if (fileName.Contains("P2"))
						{
							foreach (var despacho in jObject["despacho"]!)
							{
								var processo = despacho["processo-patente"]!;

								processo["revista"] = magazine;
								processo["despacho"] = $"{despacho["codigo"]} - {despacho["titulo"]}";

								if (despacho["texto-complementar"] is JObject comentario)
									processo["texto-complementar"] = comentario["#text"];
								if (processo["numero"] is JObject numero)
									processo["numero"] = numero["#text"];
								if (processo["data-deposito"] is JObject dataDeposito)
									processo["data-deposito"] = dataDeposito["#text"];
								if (processo["titulo"] is JObject titulo)
									processo["titulo"] = titulo["#text"];
								if (processo["data-concessao"] is JObject concessao)
									processo["data-concessao"] = concessao["#text"];
								if (processo["data-fase-nacional"] is JObject datafasenacional)
									processo["data-fase-nacional"] = datafasenacional["#text"];
								if (processo["data-publicacao-nacional"] is JObject datapublicacaonacional)
									processo["data-publicacao-nacional"] = datapublicacaonacional["data-rpi"];

								if (processo["titular-lista"] is JObject titularLista)
								{
									var titular = titularLista["titular"];
									if (titular is JArray)
										processo["titular-lista"] = titular;
									if (titular is JObject)
										processo["titular-lista"] = new JArray(titular);

									if (processo["titular-lista"] is JArray titularArray)
										foreach (var titularObj in titularArray)
											if (titularObj["endereco"]!["pais"] is JObject pais && pais.TryGetValue("sigla", out var sigla))
												if (titularObj["endereco"]!["uf"] != null)
												{
													titularObj["endereco"] = $"{sigla}/{titularObj["endereco"]!["uf"]}";
												}
												else
												{
													titularObj["endereco"] = $"{sigla}";
												}
								}

								if (processo["inventor-lista"] is JObject inventorLista)
								{
									var inventor = inventorLista["inventor"];
									if (inventor is JArray)
										processo["inventor-lista"] = inventor;
									if (inventor is JObject)
										processo["inventor-lista"] = new JArray(inventor);
								}

								if (processo["prioridade-unionista-lista"] is JObject prioridadeunionistalistaobj)
								{
									var prioridadeUnionista = prioridadeunionistalistaobj["prioridade-unionista"];
									if (prioridadeUnionista is JArray)
										processo["prioridade-unionista-lista"] = prioridadeUnionista;
									if (prioridadeUnionista is JObject)
										processo["prioridade-unionista-lista"] = new JArray(prioridadeUnionista);

									if (processo["prioridade-unionista-lista"] is JArray prioridadeLista)
										foreach (var prioridadeObj in prioridadeLista)
										{
											prioridadeObj["pais-prioridade"] =
												$"{prioridadeObj["sigla-pais"]?["#text"]}";
											prioridadeObj["numero-prioridade"] =
												$"{prioridadeObj["numero-prioridade"]?["#text"]}";
											prioridadeObj["data-prioridade"] =
												$"{prioridadeObj["data-prioridade"]?["#text"]}";
											prioridadeObj["sigla-pais"]?.Parent?.Remove();
										}
								}

								if (processo["classificacao-internacional-lista"] is JObject classificacaointernacionallistaobj)
								{
									var classificacaointernacional = classificacaointernacionallistaobj["classificacao-internacional"];
									if (classificacaointernacional is JArray)
										processo["classificacao-internacional-lista"] = classificacaointernacional;
									if (classificacaointernacional is JObject)
										processo["classificacao-internacional-lista"] = new JArray(classificacaointernacional);

									if (processo["classificacao-internacional-lista"] is JArray classificacaointernacionallistaArray)
										foreach (var classificacaointernacionalObj in classificacaointernacionallistaArray)
										{
											classificacaointernacionalObj["codigo"] = classificacaointernacionalObj["#text"];
											classificacaointernacionalObj["@ano"]?.Parent?.Remove();
											classificacaointernacionalObj["#text"]?.Parent?.Remove();
										}
								}

								if (processo["classificacao-nacional-lista"] is JObject classificacaonacionallistaobj)
								{
									var classificacaoNacional = classificacaonacionallistaobj["classificacao-nacional"];
									if (classificacaoNacional is JArray)
										processo["classificacao-nacional-lista"] = classificacaoNacional;
									if (classificacaoNacional is JObject)
										processo["classificacao-nacional-lista"] = new JArray(classificacaoNacional);

									if (processo["classificacao-nacional-lista"] is JArray campoAplicacaoListaArray)
										foreach (var campoAplicacaoListaObj in campoAplicacaoListaArray)
										{
											campoAplicacaoListaObj["codigo"] = campoAplicacaoListaObj["#text"];
											campoAplicacaoListaObj["#text"]?.Parent?.Remove();
										}
								}

								list.Add(BsonDocument.Parse(processo.ToString()));
							}
						}

						if (fileName.Contains("PC"))
						{
							foreach (var despacho in jObject["despacho"]!)
							{
								var processo = despacho["processo-programa"]!;

								processo["revista"] = magazine;
								processo["despacho"] = $"{despacho["codigo"]} - {despacho["titulo"]}";

								if (despacho["texto-complementar"] is JObject comentario)
									processo["texto-complementar"] = comentario["#text"];
								if (processo["numero"] is JObject numero)
									processo["numero"] = numero["#text"];
								if (processo["titulo"] is JObject titulo)
									processo["titulo"] = titulo["#text"];
								if (processo["data-criacao"] is JObject dataCriacao)
									processo["data-criacao"] = dataCriacao["#text"];

								if (processo["campo-aplicacao-lista"] is JObject campoaplicacaolistaobj)
								{
									var campoAplicacao = campoaplicacaolistaobj["campoAplicacao"]!;

									if (campoAplicacao is JArray)
										processo["campo-aplicacao-lista"] = campoAplicacao;

									if (campoAplicacao is JObject)
										processo["campo-aplicacao-lista"] = new JArray(campoAplicacao);

									if (processo["campo-aplicacao-lista"] is JArray campoAplicacaoListaArray)
										foreach (var campoAplicacaoListaObj in campoAplicacaoListaArray)
											if (campoAplicacaoListaObj["codigo"] is JObject campoAplicacaoListaCodigo)
												campoAplicacaoListaObj["codigo"] = campoAplicacaoListaCodigo["#text"];
								}

								if (processo["criador-lista"] is JObject criadorlistaobj)
								{
									var criador = criadorlistaobj["criador"];

									if (criador is JArray)
										processo["criador-lista"] = criador;

									if (criador is JObject)
										processo["criador-lista"] = new JArray(criador);
								}

								if (processo["linguagem-lista"] is JObject linguagemlistaobj)
								{
									var linguagem = linguagemlistaobj["linguagem"];

									if (linguagem is JArray)
										processo["linguagem-lista"] = linguagem;

									if (linguagem is JObject)
										processo["linguagem-lista"] = new JArray(linguagem);

									if (processo["linguagem-lista"] is JArray linguagemListaArray)
										foreach (var linguagemListaObj in linguagemListaArray)
										{
											linguagemListaObj["linguagem"] = linguagemListaObj["#text"];
											linguagemListaObj["#text"]?.Parent?.Remove();
										}
								}

								if (processo["tipo-programa-lista"] is JObject tipoProgramalistaobj)
								{
									var tipoPrograma = tipoProgramalistaobj["tipoPrograma"];

									if (tipoPrograma is JArray)
										processo["tipo-programa-lista"] = tipoPrograma;
									if (tipoPrograma is JObject)
										processo["tipo-programa-lista"] = new JArray(tipoPrograma);

									if (processo["tipo-programa-lista"] is JArray tipoProgramaListaArray)
										foreach (var tipoProgramaListaObj in tipoProgramaListaArray)
											if (tipoProgramaListaObj["codigo"] is JObject tipoProgramaListaCodigo)
												tipoProgramaListaObj["codigo"] = tipoProgramaListaCodigo["#text"];
								}

								if (processo["titular-lista"] is JObject titularlistaobj)
								{
									var titular = titularlistaobj["titular"];

									if (titular is JArray)
										processo["titular-lista"] = titular;
									if (titular is JObject)
										processo["titular-lista"] = new JArray(titular);
								}

								list.Add(BsonDocument.Parse(processo.ToString()));
							}
						}

						if (fileName.Contains("CT"))
						{
							foreach (var despacho in jObject["despacho"]!)
							{
								var processo = despacho["processo-contrato"]!;

								processo["revista"] = magazine;
								processo["despacho"] = $"{despacho["codigo"]} - {despacho["titulo"]}";

								if (despacho["texto-complementar"] is JObject comentario)
									processo["texto-complementar"] = comentario["#text"];
								if (processo["numero"] is JObject numero)
									processo["numero"] = numero["#text"];
								if (processo["data-protocolo"] is JObject dataProtocolo)
									processo["data-protocolo"] = dataProtocolo["#text"];

								if (processo["cedente-lista"] is JObject cedentelistaobj)
								{
									var cedente = cedentelistaobj["cedente"];
									if (cedente is JArray)
										processo["cedente-lista"] = cedente;
									if (cedente is JObject)
										processo["cedente-lista"] = new JArray(cedente);

									if (processo["cedente-lista"] is JArray cedentesArray)
										foreach (var cedentesObj in cedentesArray)
										{
											if (cedentesObj["nome-completo"] is JObject cedentesNomeCompleto)
												cedentesObj["nome-completo"] = cedentesNomeCompleto["#text"];
											if (cedentesObj["endereco"] is JObject cedentesEndereco)
												cedentesObj["endereco"] =
													cedentesEndereco["pais"]!["nome-completo"]!["#text"];
										}
								}

								if (processo["cessionaria-lista"] is JObject cessionarialistaobj)
								{
									var cessionaria = cessionarialistaobj["cessionaria"];
									if (cessionaria is JArray)
										processo["cessionaria-lista"] = cessionaria;
									if (cessionaria is JObject)
										processo["cessionaria-lista"] = new JArray(cessionaria);

									if (processo["cessionaria-lista"] is JArray cessionariasArray)
										foreach (var cessionariasObj in cessionariasArray)
										{
											if (cessionariasObj["nome-completo"] is JObject cessionariasNomeCompleto)
												cessionariasObj["nome-completo"] = cessionariasNomeCompleto["#text"];
											if (cessionariasObj["endereco"] is JObject cessionariasEndereco)
												cessionariasObj["endereco"] =
													cessionariasEndereco["pais"]!["nome-completo"]!["#text"];
											if (cessionariasObj["setor"] is JObject cessionariasSetor)
												cessionariasObj["setor"] = cessionariasSetor["#text"];
										}
								}

								if (processo["certificado-lista"] is JObject certificadolistaobj)
								{
									var certificado = certificadolistaobj["certificado"];
									if (certificado is JArray)
										processo["certificado-lista"] = certificado;
									if (certificado is JObject)
										processo["certificado-lista"] = new JArray(certificado);
								}

								if (processo["peticao-lista"] is JObject peticaolistaobj)
								{
									var peticao = peticaolistaobj["peticao"];
									if (peticao is JArray)
										processo["peticao-lista"] = peticao;
									if (peticao is JObject)
										processo["peticao-lista"] = new JArray(peticao);

									if (processo["peticao-lista"] is JArray peticoesArray)
										foreach (var peticoesObj in peticoesArray)
										{
											if (peticoesObj["numero"] is JObject peticoesNumero)
												peticoesObj["numero"] = peticoesNumero["#text"];
											if (peticoesObj["data-protocolo"] is JObject peticoesDataProtocoloro)
												peticoesObj["data-protocolo"] = peticoesDataProtocoloro["#text"];
											if (peticoesObj["requerente"] is JObject requerente)
												peticoesObj["requerente"] = requerente["nome-completo"]!["#text"];
										}
								}

								if (processo["certificado-lista"] is JArray certificadosArray)
									foreach (var certificadosObj in certificadosArray)
									{
										if (certificadosObj["numero"] is JObject certificadosNumero)
											certificadosObj["numero"] = certificadosNumero["#text"];
										if (certificadosObj["natureza-documento"] is JObject
											certificadosNaturezaDocumento)
											certificadosObj["natureza-documento"] =
												certificadosNaturezaDocumento["#text"];
										if (certificadosObj["texto-objeto"] is JObject certificadosTextoObjeto)
											certificadosObj["texto-objeto"] = certificadosTextoObjeto["#text"];
										if (certificadosObj["sigla-categoria"] is JObject certificadosSiglaCategoria)
											certificadosObj["sigla-categoria"] = certificadosSiglaCategoria["#text"];
										if (certificadosObj["descricao-moeda"] is JObject certificadosDescricaoMoeda)
											certificadosObj["descricao-moeda"] = certificadosDescricaoMoeda["#text"];
										if (certificadosObj["valor-contrato"] is JObject certificadosValorContrato)
											certificadosObj["valor-contrato"] = certificadosValorContrato["#text"];
										if (certificadosObj["forma-pagamento"] is JObject certificadosFormaPagamento)
											certificadosObj["forma-pagamento"] = certificadosFormaPagamento["#text"];
										if (certificadosObj["prazo-contrato"] is JObject certificadosPrazoContrato)
											certificadosObj["prazo-contrato"] = certificadosPrazoContrato["#text"];
										if (certificadosObj["prazo-vigencia"] is JObject certificadosPrazoVigenciaPi)
											certificadosObj["prazo-vigencia"] = certificadosPrazoVigenciaPi["#text"];
										if (certificadosObj["observacao"] is JObject certificadosObservacao)
											certificadosObj["observacao"] = certificadosObservacao["#text"];
									}

								list.Add(BsonDocument.Parse(processo.ToString()));
							}
						}

						if (fileName.Contains("DI"))
						{
							foreach (var despacho in jObject["despacho"]!)
							{
								var processo = despacho["processo-patente"]!;

								processo["revista"] = magazine;
								processo["despacho"] = $"{despacho["codigo"]} - {despacho["titulo"]}";

								if (despacho["texto-complementar"] is JObject comentario)
									processo["texto-complementar"] = comentario["#text"];
								if (processo["numero"] is JObject numero)
									processo["numero"] = numero["#text"];
								if (processo["data-deposito"] is JObject dataDeposito)
									processo["data-deposito"] = dataDeposito["#text"];
								if (processo["titulo"] is JObject titulo)
									processo["titulo"] = titulo["#text"];
								if (processo["data-concessao"] is JObject concessao)
									processo["data-concessao"] = concessao["#text"];
								if (processo["data-registro-prorrogacao"] is JObject dataregistroprorrogacao)
									processo["data-registro-prorrogacao"] = dataregistroprorrogacao["#text"];
								if (processo["data-publicacao-nacional"] is JObject datapublicacaonacional)
									processo["data-publicacao-nacional"] = datapublicacaonacional["data-rpi"];

								if (processo["titular-lista"] is JObject titularlistaobj)
								{
									var titular = titularlistaobj["titular"];
									if (titular is JArray)
										processo["titular-lista"] = titular;
									if (titular is JObject)
										processo["titular-lista"] = new JArray(titular);

									if (processo["titular-lista"] is JArray titularArray)
										foreach (var titularObj in titularArray)
											if (titularObj["endereco"]!["pais"] is JObject pais && pais.TryGetValue("sigla", out var sigla))
												if (titularObj["endereco"]!["uf"] != null)
												{
													titularObj["endereco"] = $"{sigla}/{titularObj["endereco"]!["uf"]}";
												}
												else
												{
													titularObj["endereco"] = $"{sigla}";
												}
								}

								if (processo["procurador-lista"] is JObject procuradorlistaobj)
								{
									var procurador = procuradorlistaobj["procurador"];
									if (procurador is JArray)
										processo["procurador-lista"] = procurador;
									if (procurador is JObject)
										processo["procurador-lista"] = new JArray(procurador);
								}

								if (processo["inventor-lista"] is JObject inventorlistaobj)
								{
									var inventor = inventorlistaobj["inventor"];
									if (inventor is JArray)
										processo["inventor-lista"] = inventor;
									if (inventor is JObject)
										processo["inventor-lista"] = new JArray(inventor);
								}

								if (processo["classificacao-nacional-lista"] is JObject classificacaonacionallistaobj)
								{
									var classificacaoNacional = classificacaonacionallistaobj["classificacao-nacional"];
									if (classificacaoNacional is JArray)
										processo["classificacao-nacional-lista"] = classificacaoNacional;
									if (classificacaoNacional is JObject)
										processo["classificacao-nacional-lista"] = new JArray(classificacaoNacional);

									if (processo["classificacao-nacional-lista"] is JArray campoAplicacaoListaArray)
										foreach (var campoAplicacaoListaObj in campoAplicacaoListaArray)
										{
											campoAplicacaoListaObj["codigo"] = campoAplicacaoListaObj["#text"];
											campoAplicacaoListaObj["#text"]?.Parent?.Remove();
										}
								}

								if (processo["prioridade-unionista-lista"] is JObject prioridadeunionistalistaobj)
								{
									var prioridadeUnionista = prioridadeunionistalistaobj["prioridade-unionista"];
									if (prioridadeUnionista is JArray)
										processo["prioridade-unionista-lista"] = prioridadeUnionista;
									if (prioridadeUnionista is JObject)
										processo["prioridade-unionista-lista"] = new JArray(prioridadeUnionista);

									if (processo["prioridade-unionista-lista"] is JArray prioridadeLista)
										foreach (var prioridadeObj in prioridadeLista)
										{
											prioridadeObj["pais-prioridade"] =
												$"{prioridadeObj["sigla-pais"]?["#text"]}";
											prioridadeObj["numero-prioridade"] =
												$"{prioridadeObj["numero-prioridade"]?["#text"]}";
											prioridadeObj["data-prioridade"] =
												$"{prioridadeObj["data-prioridade"]?["#text"]}";
											prioridadeObj["sigla-pais"]?.Parent?.Remove();
										}
								}

								list.Add(BsonDocument.Parse(processo.ToString())!);
							}
						}

						if (fileName.Contains("P2")) await MongoRepository.Patentes.InsertManyAsync(list);
						if (fileName.Contains("PC")) await MongoRepository.Programas.InsertManyAsync(list);
						if (fileName.Contains("CT")) await MongoRepository.Contratos.InsertManyAsync(list);
						if (fileName.Contains("DI")) await MongoRepository.Desenhos.InsertManyAsync(list);
					}
				}
			}
			catch (Exception ex)
			{
				await Console.Error.WriteLineAsync($"{DateTime.Now} - Error ProcessFiles(): {ex.Message}");
			}
		}
	}

	public static class Config
	{
		public static string DataStorage => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
	}

	public static class MongoRepository
	{
		private static readonly IMongoDatabase Database;

		static MongoRepository()
		{
			// mongodb+srv://sa:a7rCv73Wzjqlipvg@opencluster.gmh7wek.mongodb.net/?retryWrites=true&w=majority
			// mongodb://127.0.0.1:27018
			IMongoClient client =
				new MongoClient(
					"mongodb+srv://userMAGAZINE:7WYx9D4nm70gLHU3@opencluster.gmh7wek.mongodb.net/?retryWrites=true&w=majority");
			Database = client.GetDatabase("DB_MAGAZINE");
		}

		public static IMongoCollection<BsonDocument> Marcas => Database.GetCollection<BsonDocument>("marcas");
		public static IMongoCollection<BsonDocument> Patentes => Database.GetCollection<BsonDocument>("patentes");
		public static IMongoCollection<BsonDocument> Desenhos => Database.GetCollection<BsonDocument>("desenhos");
		public static IMongoCollection<BsonDocument> Programas => Database.GetCollection<BsonDocument>("programas");
		public static IMongoCollection<BsonDocument> Contratos => Database.GetCollection<BsonDocument>("contratos");

		public static async Task InsertManyAsync(string collectionName, IEnumerable<BsonDocument> documents)
		{
			var collection = Database.GetCollection<BsonDocument>(collectionName);
			await collection.InsertManyAsync(documents);
		}

		public static string GetLastMagazine(string collectionName)
		{
			var collection = Database.GetCollection<BsonDocument>(collectionName);

			return collection.AsQueryable().OrderByDescending(x => x["revista"]).Select(x => x["revista"].AsString)
				.FirstOrDefault() ?? string.Empty;
		}
	}
}