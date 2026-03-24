using System.Text.Json;
using DeDll;
using Microsoft.Data.SqlClient;

namespace PhoneDecryptHelper;

#region Data Models

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
    public string Mode { get; set; } = "decrypt";

    // decrypt mode
    public List<DecryptRequest> Requests { get; set; } = new();

    // full mode
    public string Server { get; set; } = "";
    public int Port { get; set; } = 1433;
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public string SignupDatabase { get; set; } = "";
    public string PhoneDatabase { get; set; } = "";
    public string CandidateTable { get; set; } = "";
    public string FilterMode { get; set; } = "all";
    public List<string> IdCards { get; set; } = new();
    public bool UpdateResults { get; set; } = true;
}

internal sealed class RecordResult
{
    public string PrimaryKey { get; set; } = "";
    public string IdCard { get; set; } = "";
    public string Province { get; set; } = "";
    public string EncryptedPhone { get; set; } = "";
    public string DecryptedPhone { get; set; } = "";
    public string Status { get; set; } = "";
    public string Note { get; set; } = "";
}

internal sealed class ResponseEnvelope
{
    public string BackendName { get; set; } = "csharp-x86-helper";

    // decrypt mode
    public List<DecryptResponse> Responses { get; set; } = new();

    // full mode
    public List<RecordResult> Records { get; set; } = new();
    public int TotalRows { get; set; }
    public int MatchedInfoRows { get; set; }
    public int DecryptedRows { get; set; }
    public int UpdatedRows { get; set; }
    public int SkippedRows { get; set; }
    public int FailedRows { get; set; }
    public List<string> Logs { get; set; } = new();
    public string Error { get; set; } = "";
}

#endregion

internal static class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    private static int Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.WriteLine("Usage: PhoneDecryptHelper <input.json> <output.json>");
            return 2;
        }

        try
        {
            var envelope = JsonSerializer.Deserialize<RequestEnvelope>(
                File.ReadAllText(args[0]), JsonOpts) ?? new RequestEnvelope();

            var result = envelope.Mode == "full"
                ? RunFullWorkflow(envelope)
                : RunDecryptOnly(envelope);

            File.WriteAllText(args[1], JsonSerializer.Serialize(result, JsonOpts));
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }

    // ── decrypt-only mode (backward compatible) ──────────────────────────

    private static ResponseEnvelope RunDecryptOnly(RequestEnvelope envelope)
    {
        var decryptor = new DeAESClass();
        var result = new ResponseEnvelope();

        foreach (var item in envelope.Requests)
        {
            var response = new DecryptResponse
            {
                PrimaryKey = item.PrimaryKey ?? "",
                EncryptedPhone = item.EncryptedPhone ?? ""
            };

            try
            {
                var province = NormalizeProvince(item.Province, item.SortCode);
                var phone = (item.EncryptedPhone ?? "").Trim();
                if (decryptor.CheckEncrypted(phone))
                {
                    response.DecryptedPhone = (decryptor.Decrypt(
                        phone, (item.SortCode ?? "").Trim(),
                        (item.ExamDate ?? "").Trim(), province) ?? "").Trim();
                    response.WasEncrypted = true;
                }
                else
                {
                    response.DecryptedPhone = phone;
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

        return result;
    }

    // ── full workflow mode (connect + query + decrypt + update) ───────────

    private static ResponseEnvelope RunFullWorkflow(RequestEnvelope req)
    {
        var result = new ResponseEnvelope { BackendName = "csharp-x86-helper" };
        var logs = result.Logs;

        try
        {
            var connStr = new SqlConnectionStringBuilder
            {
                DataSource = $"{req.Server},{req.Port}",
                UserID = req.Username,
                Password = req.Password,
                TrustServerCertificate = true,
                Encrypt = false
            }.ConnectionString;

            var signupDb = Esc(req.SignupDatabase);
            var phoneDb = Esc(string.IsNullOrWhiteSpace(req.PhoneDatabase)
                ? req.SignupDatabase : req.PhoneDatabase);
            var table = Esc(req.CandidateTable);

            logs.Add($"连接 SQL Server：{req.Server},{req.Port}");
            using var conn = new SqlConnection(connStr);
            conn.Open();
            logs.Add("数据库连接成功。");

            // 1. query
            var rows = QueryRows(conn, signupDb, phoneDb, table, req);
            result.TotalRows = rows.Count;
            logs.Add($"查询到 {rows.Count} 条记录。");

            // 2. decrypt
            var decryptor = new DeAESClass();
            var updates = new List<(string phone, string pk)>();

            foreach (var r in rows)
            {
                var encrypted = (r.EncryptedPhone ?? "").Trim();
                if (string.IsNullOrEmpty(encrypted))
                {
                    result.SkippedRows++;
                    result.Records.Add(new RecordResult
                    {
                        PrimaryKey = r.PrimaryKey, IdCard = r.IdCard,
                        Province = r.Province, Status = "跳过",
                        Note = "未找到电话密文（web_info.info1 为空）。"
                    });
                    continue;
                }

                result.MatchedInfoRows++;
                var pk = (r.PrimaryKey ?? "").Trim();

                if (pk.Length < 8)
                {
                    result.FailedRows++;
                    result.Records.Add(new RecordResult
                    {
                        PrimaryKey = pk, IdCard = r.IdCard, Province = r.Province,
                        EncryptedPhone = encrypted, Status = "失败",
                        Note = $"主键编号长度不足（{pk.Length}），无法拆分考试代码和年月。"
                    });
                    continue;
                }

                var sortCode = pk[..2];
                var examDate = pk[2..8];
                var province = NormalizeProvince(r.Province, sortCode);

                if (string.IsNullOrWhiteSpace(province))
                {
                    result.FailedRows++;
                    result.Records.Add(new RecordResult
                    {
                        PrimaryKey = pk, IdCard = r.IdCard, Province = r.Province,
                        EncryptedPhone = encrypted, Status = "失败", Note = "考区代码为空。"
                    });
                    continue;
                }

                try
                {
                    string decrypted;
                    bool wasEnc;
                    if (decryptor.CheckEncrypted(encrypted))
                    {
                        decrypted = (decryptor.Decrypt(encrypted, sortCode, examDate, province) ?? "").Trim();
                        wasEnc = true;
                    }
                    else
                    {
                        decrypted = encrypted;
                        wasEnc = false;
                    }

                    if (string.IsNullOrWhiteSpace(decrypted))
                    {
                        result.FailedRows++;
                        result.Records.Add(new RecordResult
                        {
                            PrimaryKey = pk, IdCard = r.IdCard, Province = r.Province,
                            EncryptedPhone = encrypted, Status = "失败", Note = "解密结果为空。"
                        });
                        continue;
                    }

                    result.DecryptedRows++;
                    updates.Add((decrypted, pk));
                    result.Records.Add(new RecordResult
                    {
                        PrimaryKey = pk, IdCard = r.IdCard, Province = r.Province,
                        EncryptedPhone = encrypted, DecryptedPhone = decrypted,
                        Status = "成功",
                        Note = wasEnc ? "已调用 DLL 解密。" : "原值未加密，直接写回。"
                    });
                }
                catch (Exception ex)
                {
                    result.FailedRows++;
                    result.Records.Add(new RecordResult
                    {
                        PrimaryKey = pk, IdCard = r.IdCard, Province = r.Province,
                        EncryptedPhone = encrypted, Status = "失败", Note = ex.Message
                    });
                }
            }

            // 3. update
            if (updates.Count > 0 && req.UpdateResults)
            {
                logs.Add($"正在回写 {updates.Count} 条解密电话...");
                var sql = $"UPDATE [{signupDb}].[dbo].[{table}] SET [备用3]=@p WHERE CAST([主键编号] AS VARCHAR(50))=@k";
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Add("@p", System.Data.SqlDbType.NVarChar, 50);
                cmd.Parameters.Add("@k", System.Data.SqlDbType.NVarChar, 50);
                foreach (var (phone, pk) in updates)
                {
                    cmd.Parameters["@p"].Value = phone;
                    cmd.Parameters["@k"].Value = pk;
                    cmd.ExecuteNonQuery();
                }
                result.UpdatedRows = updates.Count;
                logs.Add($"已更新 {updates.Count} 条电话到备用3。");
            }
            else
            {
                logs.Add("没有可回写的电话记录。");
            }
        }
        catch (Exception ex)
        {
            result.Error = ex.ToString();
            logs.Add($"错误：{ex.Message}");
        }

        return result;
    }

    private sealed class CandidateRow
    {
        public string PrimaryKey { get; set; } = "";
        public string IdCard { get; set; } = "";
        public string Province { get; set; } = "";
        public string EncryptedPhone { get; set; } = "";
    }

    private static List<CandidateRow> QueryRows(
        SqlConnection conn, string signupDb, string phoneDb,
        string table, RequestEnvelope req)
    {
        var baseSql = $@"
            SELECT
                LTRIM(RTRIM(CAST(ks.[主键编号] AS VARCHAR(50)))),
                LTRIM(RTRIM(ISNULL(CAST(ks.[身份证号] AS VARCHAR(50)), ''))),
                LTRIM(RTRIM(ISNULL(CAST(ks.[考区] AS VARCHAR(50)), ''))),
                LTRIM(RTRIM(ISNULL(CAST(info.[info1] AS VARCHAR(255)), '')))
            FROM [{signupDb}].[dbo].[{table}] ks
            LEFT JOIN [{phoneDb}].[dbo].[web_info] info
                ON CAST(info.[zjbh] AS VARCHAR(50)) = CAST(ks.[主键编号] AS VARCHAR(50))
               AND CAST(info.[examsort] AS VARCHAR(10)) = LEFT(CAST(ks.[主键编号] AS VARCHAR(50)), 2)
            WHERE ks.[主键编号] IS NOT NULL";

        var rows = new List<CandidateRow>();

        if (req.FilterMode != "partial" || req.IdCards.Count == 0)
        {
            using var cmd = new SqlCommand(baseSql + " ORDER BY ks.[主键编号]", conn);
            ReadRows(cmd, rows);
        }
        else
        {
            for (int i = 0; i < req.IdCards.Count; i += 500)
            {
                var chunk = req.IdCards.Skip(i).Take(500).ToList();
                var ph = string.Join(",", chunk.Select((_, j) => $"@id{j}"));
                var sql = baseSql + $" AND CAST(ks.[身份证号] AS VARCHAR(50)) IN ({ph}) ORDER BY ks.[主键编号]";
                using var cmd = new SqlCommand(sql, conn);
                for (int j = 0; j < chunk.Count; j++)
                    cmd.Parameters.AddWithValue($"@id{j}", chunk[j]);
                ReadRows(cmd, rows);
            }
        }

        return rows;
    }

    private static void ReadRows(SqlCommand cmd, List<CandidateRow> rows)
    {
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            rows.Add(new CandidateRow
            {
                PrimaryKey = reader.IsDBNull(0) ? "" : reader.GetString(0),
                IdCard = reader.IsDBNull(1) ? "" : reader.GetString(1),
                Province = reader.IsDBNull(2) ? "" : reader.GetString(2),
                EncryptedPhone = reader.IsDBNull(3) ? "" : reader.GetString(3)
            });
        }
    }

    private static string Esc(string name) => name.Trim().Replace("]", "]]");

    private static string NormalizeProvince(string? rawProvince, string? sortCode)
    {
        var province = (rawProvince ?? "").Trim();
        if (province.StartsWith("141", StringComparison.OrdinalIgnoreCase)
            && (sortCode ?? "").Trim() == "96")
            return "141";
        return province;
    }
}
