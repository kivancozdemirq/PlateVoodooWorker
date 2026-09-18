using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using PlateVoodooWorker.Context;
using PlateVoodooWorker.Model; 
using RestSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using PlateVoodooWorker.SqlStore;
using System.Text.Json;
using System.Text.Json.Serialization;

#region PowerOCR DTO Models
public class PowerOcrRequest
{
    [JsonPropertyName("SplitImage")]
    public int SplitImage { get; set; } = 0;

    [JsonPropertyName("ImageData")]
    public string ImageData { get; set; } = string.Empty;
}

public class PowerOcrResponse
{
    [JsonPropertyName("RetCode")]
    public int RetCode { get; set; }

    [JsonPropertyName("DetailedInfo")]
    public string DetailedInfo { get; set; } = string.Empty;

    [JsonPropertyName("UsageLeft")]
    public long UsageLeft { get; set; }

    [JsonPropertyName("LstPlate Result")]
    public List<PowerOcrPlateResult>? LstPlateResult { get; set; }
}

public class PowerOcrPlateResult
{
    [JsonPropertyName("RetCode")]
    public int RetCode { get; set; }

    [JsonPropertyName("Value")]
    public string Value { get; set; } = string.Empty;

    [JsonPropertyName("ValueProbability")]
    public double ValueProbability { get; set; }

    [JsonPropertyName("WarpedBoxProbability")]
    public double WarpedBoxProbability { get; set; }

    [JsonPropertyName("WarpedBox")]
    public List<int>? WarpedBox { get; set; }
}
#endregion

public class Worker
{
    private string GetBoxString(List<int>? box)
    {
        if (box == null || box.Count == 0) return string.Empty;
        return $"Box: [{string.Join(",", box)}]";
    }

    public void ExecuteProcess(int operationId, int caseState)
    {
        switch (caseState)
        {
            case 1:
                Console.WriteLine($"{operationId} için case 1 (Excel Okuma & DB Kayıt) çalışmaya başladı.");
                string filePath = @"C:\Users\kivanc.ozdemir\Downloads\deneme.xlsx";

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Excel dosyası bulunamadı: {filePath}");
                    break;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int row = 2;
                    int column = 1;

                    while (worksheet.Cells[row, column].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, column].Value.ToString()))
                    {
                        Vehicle vehicle = new Vehicle();
                        string id = worksheet.Cells[row, column].Value.ToString()!.Trim();

                        vehicle.TransactionId = id;
                        vehicle.State = 1;
                        vehicle.OperationId = operationId;

                        string companyId = id.StartsWith("0050") ? "212" : id.StartsWith("0060") ? "216" : "hatalı id : " + id;

                        try
                        {
                            Aselsan? aselsanData = null;

                            #region Aselsan Data Fetching
                            if (companyId == "216")
                            {
                                using (AselsanAsyaContext aselsanAsyaContext = new AselsanAsyaContext())
                                {
                                    aselsanData = aselsanAsyaContext.SAP_TRANSACTIONS_VIEW.FromSqlRaw(SqlStore.AselsanSql(id, companyId)).FirstOrDefault();
                                }
                            }
                            else if (companyId == "212")
                            {
                                using (AselsanAvrupaContext aselsanAvrupaContext = new AselsanAvrupaContext())
                                {
                                    aselsanData = aselsanAvrupaContext.SAP_TRANSACTIONS_VIEW.FromSqlRaw(SqlStore.AselsanSql(id, companyId)).FirstOrDefault();
                                }
                            }
                            else
                            {
                                Console.WriteLine(id + " 'da hata oluştu - Geçersiz Firma Kodu");
                                vehicle.State = 3;
                                vehicle.Description = "Hatalı ID : " + id;
                            }

                            if (aselsanData != null)
                            {
                                vehicle.Aselsan_FRONT = string.IsNullOrEmpty(aselsanData.LPR_FRONT) ? null : aselsanData.LPR_FRONT;
                                vehicle.Aselsan_FRONTCONF = aselsanData.LPR_FRONTCONF >= 0 ? Math.Round((decimal)aselsanData.LPR_FRONTCONF, 2) : null;
                                vehicle.Aselsan_REAR = string.IsNullOrEmpty(aselsanData.LPR_REAR) ? null : aselsanData.LPR_REAR;
                                vehicle.Aselsan_REARCONF = aselsanData.LPR_REARCONF >= 0 ? Math.Round((decimal)aselsanData.LPR_REARCONF, 2) : null;
                                vehicle.Aselsan_FINALCONF = aselsanData.LPR_FINALCONF >= 0 ? Math.Round((decimal)aselsanData.LPR_FINALCONF, 2) : null;
                                vehicle.Aselsan_FINAL = string.IsNullOrEmpty(aselsanData.LPR_FINAL) ? null : aselsanData.LPR_FINAL;
                                vehicle.Aselsan_APPROVEDCLASS = string.IsNullOrEmpty(aselsanData.APPROVEDCLASS) ? null : aselsanData.APPROVEDCLASS;
                                vehicle.Aselsan_APPROVEDPLATENUMBER = string.IsNullOrEmpty(aselsanData.APPROVEDPLATENUMBER) ? null : aselsanData.APPROVEDPLATENUMBER;
                                vehicle.Aselsan_TABLETYPE = string.IsNullOrEmpty(aselsanData.TABLETYPE) ? null : aselsanData.TABLETYPE;
                            }
                            #endregion
                        }
                        catch (Exception e)
                        {
                            vehicle.Description = e.Message + " -- " + e.InnerException;
                            vehicle.State = 3;
                        }

                        using (PlateVoodooContext plateContext = new PlateVoodooContext())
                        {
                            bool isExist = plateContext.Vehicles.Any(x => x.TransactionId == id);
                            if (!isExist)
                            {
                                plateContext.Vehicles.Add(vehicle);
                                plateContext.SaveChanges();
                            }
                        }

                        row++;
                    }
                }
                Console.WriteLine($"Case 1 Tamamlandı - {DateTime.Now}");
                break;

            case 2:
                Console.WriteLine(operationId + " için case 2 çalışmaya başladı");
                string baseUrl = "http://photoservices.kuzeymarmaraotoyolu.com:30080/";
                string endpoint = "/GetPhotoV2";
                string powerOcrUrl = "http://10.75.39.90:8901/recognize";

                using (PlateVoodooContext plateRecognizerContext = new PlateVoodooContext())
                {
                    foreach (var item in plateRecognizerContext.Vehicles.Where(x => x.State == 1 && x.OperationId == operationId).ToList())
                    {
                        try
                        {
                            string transactionId = item.TransactionId;
                            string companyId = item.TransactionId.Substring(0, 4) == "0050" ? "212" : item.TransactionId.Substring(0, 4) == "0060" ? "216" : "";
                            string tableType = item.Aselsan_TABLETYPE ?? "";


                            var client = new RestClient(baseUrl);
                            var request = new RestRequest(endpoint, Method.Get)
                                .AddQueryParameter("TransactionId", transactionId)
                                .AddQueryParameter("CompanyId", companyId)
                                .AddQueryParameter("TableType", tableType);

                            var response = client.Execute(request);

                            Console.WriteLine($"--- İSTEK LOGU [{transactionId}] ---");
                            Console.WriteLine($"Atılan URL: {client.BuildUri(request)}");
                            Console.WriteLine($"HTTP Status: {response.StatusCode} ({(int)response.StatusCode})");
                            Console.WriteLine($"Error Message: {response.ErrorMessage}");
                            Console.WriteLine($"Response Content: '{response.Content}'");
                            Console.WriteLine("----------------------------------");

                            var jsonContent = response.Content;
                            string imageValue = "";

                            if (!string.IsNullOrEmpty(jsonContent))
                            {
                                try
                                {
                                    using var document = JsonDocument.Parse(jsonContent);

                                    if (document.RootElement.ValueKind == JsonValueKind.Object)
                                    {
                                        foreach (var prop in document.RootElement.EnumerateObject())
                                        {
                                            if (prop.Name.Equals("image", StringComparison.OrdinalIgnoreCase) ||
                                                prop.Name.Equals("imageData", StringComparison.OrdinalIgnoreCase))
                                            {
                                                imageValue = prop.Value.GetString() ?? "";
                                                break;
                                            }
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(imageValue) && imageValue.Contains(","))
                                    {
                                        imageValue = imageValue.Split(',')[1];
                                    }
                                }
                                catch
                                {
                                    imageValue = jsonContent.Trim('"');
                                }
                            }

                            if (!string.IsNullOrEmpty(imageValue))
                            {
                                var clientPowerOcr = new RestClient(powerOcrUrl);
                                var requestPowerOcr = new RestRequest("", Method.Post);
                                requestPowerOcr.AddHeader("Content-Type", "application/json");

                                var payload = new PowerOcrRequest
                                {
                                    SplitImage = 0,
                                    ImageData = imageValue
                                };

                                requestPowerOcr.AddJsonBody(payload);
                                var responsePowerOcr = clientPowerOcr.Execute(requestPowerOcr);

                                if (responsePowerOcr.IsSuccessful && !string.IsNullOrEmpty(responsePowerOcr.Content))
                                {
                                    var powerOcrResult = JsonSerializer.Deserialize<PowerOcrResponse>(responsePowerOcr.Content);

                                    if (powerOcrResult != null && powerOcrResult.RetCode == 0 && powerOcrResult.LstPlateResult != null && powerOcrResult.LstPlateResult.Count > 0)
                                    {
                                        var result1 = powerOcrResult.LstPlateResult[0];

                                        item.FrontPlate = result1.Value?.ToUpper();
                                        item.FrontConf = Math.Round(result1.ValueProbability * 100, 1);
                                        item.FrontCoordination = GetBoxString(result1.WarpedBox);

                                        if (powerOcrResult.LstPlateResult.Count > 1)
                                        {
                                            var result2 = powerOcrResult.LstPlateResult[1];
                                            item.RearPlate = result2.Value?.ToUpper();
                                            item.RearConf = Math.Round(result2.ValueProbability * 100, 1);
                                            item.RearCoordination = GetBoxString(result2.WarpedBox);
                                        }

                                        item.State = 2;
                                    }
                                    else
                                    {
                                        item.Description = $"PowerOCR plaka okuyamadı - {transactionId}";
                                        item.State = 3;
                                    }
                                }
                                else
                                {
                                    item.Description = "PowerOCR Servisine Ulaşılamadı";
                                    item.State = 3;
                                }
                            }
                            else
                            {
                                item.Description = $"Görsel verisi boş - Servis Yanıtı: {jsonContent}";
                                item.State = 3;
                            }
                        }
                        catch (Exception e)
                        {
                            item.Description = e.Message + " -- " + e.InnerException;
                            item.State = 3;
                        }

                        item.UpdateDate = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss", new CultureInfo("tr-TR"));
                        plateRecognizerContext.SaveChanges();
                    }
                }
                break;

            default:
                break;
        }
    }
}