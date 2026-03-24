using System.Text.Json;
using DeDll;

namespace PhoneDecryptHelper;

internal sealed class DecryptRequest
{
    public string PrimaryKey { get; set; } = "";
    public string EncryptedPhone { get; set; } = "";
    public string SortCode { get; set; } = "";
    public string ExamDate { get; set; } = "";
    public string Province { get; set; } = "";
}

internal sealed class DecryptResponse
{
    public string PrimaryKey { get; set; } = "";
    public string EncryptedPhone { get; set; } = "";
    public string DecryptedPhone { get; set; } = "";
    public bool WasEncrypted { get; set; }
    public bool Success { get; set; }
    public string Error { get; set; } = "";
}

internal sealed class RequestEnvelope
{
    public List<DecryptRequest> Requests { get; set; } = new();
}

internal sealed class ResponseEnvelope
{
    public string BackendName { get; set; } = "csharp-x86-helper";
    public List<DecryptResponse> Responses { get; set; } = new();
}

internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.WriteLine("Usage: PhoneDecryptHelper <input.json> <output.json>");
            return 2;
        }

        try
        {
            var inputPath = args[0];
            var outputPath = args[1];
            var payload = JsonSerializer.Deserialize<RequestEnvelope>(File.ReadAllText(inputPath))
                         ?? new RequestEnvelope();

            var decryptor = new DeAESClass();
            var result = new ResponseEnvelope();

            foreach (var item in payload.Requests)
            {
                var response = new DecryptResponse
                {
                    PrimaryKey = item.PrimaryKey ?? "",
                    EncryptedPhone = item.EncryptedPhone ?? ""
                };

                try
                {
                    var province = NormalizeProvince(item.Province, item.SortCode);
                    var encryptedPhone = (item.EncryptedPhone ?? "").Trim();
                    if (decryptor.CheckEncrypted(encryptedPhone))
                    {
                        response.DecryptedPhone = (decryptor.Decrypt(
                            encryptedPhone,
                            (item.SortCode ?? "").Trim(),
                            (item.ExamDate ?? "").Trim(),
                            province
                        ) ?? "").Trim();
                        response.WasEncrypted = true;
                    }
                    else
                    {
                        response.DecryptedPhone = encryptedPhone;
                        response.WasEncrypted = false;
                    }

                    response.Success = !string.IsNullOrWhiteSpace(response.DecryptedPhone);
                    response.Error = response.Success ? "" : "解密结果为空。";
                }
                catch (Exception ex)
                {
                    response.Success = false;
                    response.Error = ex.Message;
                }

                result.Responses.Add(response);
            }

            var options = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(outputPath, JsonSerializer.Serialize(result, options));
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }

    private static string NormalizeProvince(string? rawProvince, string? sortCode)
    {
        var province = (rawProvince ?? "").Trim();
        var sort = (sortCode ?? "").Trim();
        if (province.StartsWith("141", StringComparison.OrdinalIgnoreCase) && sort == "96")
        {
            return "141";
        }

        return province;
    }
}
