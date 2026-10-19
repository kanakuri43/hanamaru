// tech-analysis / prime-hanamaru スキル(ta.py / screen.py)の C# 移植
//   hanamaru ta <コード>...   指標を出力し analysis/<日付>/<symbol>_<名前>.md に保存
//   hanamaru screen [上限]    プライム全銘柄を◎判定し analysis/<日付>/prime_hanamaru.md に保存
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using ExcelDataReader;

CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
Console.OutputEncoding = new UTF8Encoding(false);

if (args.Length == 0) args = ["screen"]; // 省略時は screen
if (args[0] is not ("ta" or "screen"))
{
    Console.WriteLine("usage: hanamaru ta <code>...  |  hanamaru screen [limit]");
    return 1;
}
return args[0] == "ta" ? await Ta.Run(args[1..]) : await Screen.Run(args.Length > 1 ? int.Parse(args[1]) : null);

record Bar(DateOnly D, double H, double L, double C, double V);

static class Yahoo
{
    static readonly HttpClient Http = new(new SocketsHttpHandler { PooledConnectionLifetime = TimeSpan.FromMinutes(5), MaxConnectionsPerServer = 128 }) { Timeout = TimeSpan.FromSeconds(60) };
    static Yahoo() => Http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");

    public static async Task<(string Sym, string Name, List<Bar> Rows)> Fetch(string sym)
    {
        if (!sym.Contains('.') && char.IsDigit(sym[0])) sym += ".T"; // 日本株コードは .T 補完
        using var doc = JsonDocument.Parse(await Http.GetStringAsync($"https://query1.finance.yahoo.com/v8/finance/chart/{sym}?range=5y&interval=1d"));
        var r = doc.RootElement.GetProperty("chart").GetProperty("result")[0];
        var meta = r.GetProperty("meta");
        long tz = meta.TryGetProperty("gmtoffset", out var g) ? g.GetInt64() : 0;
        var q = r.GetProperty("indicators").GetProperty("quote")[0];
        var ts = r.GetProperty("timestamp"); var h = q.GetProperty("high"); var l = q.GetProperty("low");
        var c = q.GetProperty("close"); var v = q.GetProperty("volume");
        var rows = new List<Bar>();
        for (int i = 0; i < ts.GetArrayLength(); i++)
        {
            if (c[i].ValueKind == JsonValueKind.Null || h[i].ValueKind == JsonValueKind.Null || l[i].ValueKind == JsonValueKind.Null) continue;
            var d = DateOnly.FromDateTime(DateTimeOffset.FromUnixTimeSeconds(ts[i].GetInt64() + tz).UtcDateTime);
            rows.Add(new(d, h[i].GetDouble(), l[i].GetDouble(), c[i].GetDouble(), v[i].ValueKind == JsonValueKind.Null ? 0 : v[i].GetDouble()));
        }
        return (sym, meta.TryGetProperty("longName", out var n) ? n.GetString() ?? "" : "", rows);
    }
}

// 時間軸ごとの高値・安値・終値
record Series(double[] H, double[] L, double[] C)
{
    public int N => C.Length;
    public static Series Of(IEnumerable<(double h, double l, double c)> x) { var a = x.ToArray(); return new(a.Select(t => t.h).ToArray(), a.Select(t => t.l).ToArray(), a.Select(t => t.c).ToArray()); }
    public static Series Resample<K>(List<Bar> b, Func<DateOnly, K> key) where K : notnull
    {
        var o = new Dictionary<K, (double h, double l, double c)>(); var order = new List<K>();
        foreach (var x in b)
        {
            var k = key(x.D);
            if (!o.TryGetValue(k, out var t)) { order.Add(k); o[k] = (x.H, x.L, x.C); }
            else o[k] = (Math.Max(t.h, x.H), Math.Min(t.l, x.L), x.C);
        }
        return Of(order.Select(k => o[k]));
    }
}

static class Ind
{
    public static double Sma(double[] a, int n) => a[^n..].Average();
    public static double[] Ema(double[] a, int n)
    {
        double k = 2.0 / (n + 1), e = a[0]; var o = new double[a.Length];
        for (int i = 0; i < a.Length; i++) { e = a[i] * k + e * (1 - k); o[i] = e; }
        return o;
    }
    public static double Rsi(double[] a, int n = 14)
    {
        double g = 0, l = 0;
        for (int i = 1; i <= n; i++) { var d = a[i] - a[i - 1]; g += Math.Max(d, 0); l += Math.Max(-d, 0); }
        g /= n; l /= n;
        for (int i = n + 1; i < a.Length; i++) { var d = a[i] - a[i - 1]; g = (g * (n - 1) + Math.Max(d, 0)) / n; l = (l * (n - 1) + Math.Max(-d, 0)) / n; }
        return l != 0 ? 100 - 100 / (1 + g / l) : 100;
    }
    public static double[] MacdHist(double[] c)
    {
        var m = Ema(c, 12).Zip(Ema(c, 26), (x, y) => x - y).ToArray();
        var s = Ema(m, 9);
        return m.Zip(s, (x, y) => x - y).ToArray();
    }
    public static double Atr(Series s)
    {
        double t = 0;
        for (int i = s.N - 14; i < s.N; i++)
            t += Math.Max(s.H[i] - s.L[i], Math.Max(Math.Abs(s.H[i] - s.C[i - 1]), Math.Abs(s.L[i] - s.C[i - 1])));
        return t / 14;
    }
    public static double Stoch(Series s)
    {
        var hh = s.H[^14..].Max(); var ll = s.L[^14..].Min();
        return hh > ll ? (s.C[^1] - ll) / (hh - ll) * 100 : 50;
    }
}

static class Ta
{

    public static (Series day, Series week, Series month) Frames(List<Bar> rows) => (
        Series.Of(rows.Select(r => (r.H, r.L, r.C))),
        Series.Resample(rows, d => (ISOWeek.GetYear(d.ToDateTime(default)), ISOWeek.GetWeekOfYear(d.ToDateTime(default)))),
        Series.Resample(rows, d => (d.Year, d.Month)));

    static void Report(StringBuilder o, string name, Series s, int[] mas)
    {
        var c = s.C; int n = s.N;
        o.AppendLine($"## {name} (n={n}, 終値 {c[^1]:#,0.##})\n");
        o.AppendLine("| 指標 | 値 | 備考 |\n|---|--:|---|");
        foreach (var k in mas.Where(k => n >= k))
        {
            var m = Ind.Sma(c, k);
            var prev = n >= k + 5 ? c[(n - k - 5)..(n - 5)].Average() : (double?)null;
            o.AppendLine($"| SMA{k} | {m:#,0.##} | 終値比 {Math.Round((c[^1] / m - 1) * 100, 1) + 0.0:+0.0;-0.0}%" + (prev is { } p ? $" / 5本前 {p:#,0.##} ({(m > p ? "上向き" : "下向き")})" : "") + " |");
        }
        if (n > 30)
        {
            var h = Ind.MacdHist(c); var r = Ind.Rsi(c);
            o.AppendLine($"| RSI14 | {r:F1} | {(r >= 70 ? "買われすぎ" : r <= 30 ? "売られすぎ" : "")} |");
            o.AppendLine($"| MACDヒストグラム | {h[^1]:F2} | 前 {h[^2]:F2} ({(h[^1] > h[^2] ? "拡大" : "縮小")}) |");
        }
        var mu = Ind.Sma(c, 20); var sd = Math.Sqrt(c[^20..].Sum(x => (x - mu) * (x - mu)) / 20);
        o.AppendLine($"| ボリンジャー20 | +2σ {mu + 2 * sd:#,0.##} | -2σ {mu - 2 * sd:#,0.##} |");
        var st = Ind.Stoch(s);
        o.AppendLine($"| Stoch%K14 | {st:F0} | {(st >= 80 ? "過熱" : st <= 20 ? "売られすぎ" : "")} |");
        foreach (var k in new[] { 20, 52, 60 }.Where(k => n >= k))
            o.AppendLine($"| {k}本 高値 | {s.H[^k..].Max():#,0.##} | 安値 {s.L[^k..].Min():#,0.##} |");
        o.AppendLine();
    }

    public static async Task<int> Run(string[] codes)
    {
        int rc = 0;
        foreach (var code in codes)
        {
            try
            {
                var (sym, name, rows) = await Yahoo.Fetch(code);
                var (d, w, m) = Frames(rows);
                var last = rows[^1]; var o = new StringBuilder();
                var v = rows.Select(r => r.V).ToArray();
                o.AppendLine($"# {sym} {name} テクニカル分析 ({last.D:yyyy-MM-dd}終値 {last.C:#,0.##})\n");
                o.AppendLine($"- ATR14: {Ind.Atr(d):#,0.##} / 出来高 5日/20日比: {v[^5..].Average() / Math.Max(v[^20..].Average(), 1):F2} / 5年高値: {rows.Max(r => r.H):#,0.##} / 日足数: {rows.Count}\n");
                o.AppendLine("## 直近の終値\n\n| 日付 | " + string.Join(" | ", rows[^10..].Select(r => r.D.ToString("MM/dd"))) + " |\n|---|" + string.Concat(Enumerable.Repeat("--:|", 10)));
                o.AppendLine("| 終値 | " + string.Join(" | ", rows[^10..].Select(r => r.C.ToString("#,0.##"))) + " |\n");
                Report(o, "日足", d, [5, 25, 75, 200]);
                Report(o, "週足", w, [13, 26, 52]);
                Report(o, "月足", m, [12, 24, 60]);
                o.AppendLine("---\nファンダ・決算日程は未考慮、投資判断は自己責任。");

                var dir = Path.Combine("analysis", last.D.ToString("yyyyMMdd")); Directory.CreateDirectory(dir);
                var safe = string.Concat(name.Where(ch => !Path.GetInvalidFileNameChars().Contains(ch) && ch != ' '));
                var path = Path.Combine(dir, $"{sym}_{safe}.md");
                File.WriteAllText(path, o.ToString(), new UTF8Encoding(false));
                Console.WriteLine(o);
                Console.WriteLine("保存先: " + path);
            }
            catch (Exception e) { Console.Error.WriteLine($"ERR {code}: {e.Message}"); rc = 1; }
        }
        return rc;
    }
}

// 週足が◎の銘柄だけ、Claude API で判定の妥当性を一言チェックする(ANTHROPIC_API_KEY 未設定なら何もしない)
static class Ai
{
    public const string Model ="claude-opus-5-5";
    const string System = """
        あなたは投資初心者に株を説明する案内役。機械判定で週足◎(上昇トレンド・MACDゴールデンクロス直後は判定済みで確認不要)となった銘柄を、渡された数値だけから、今エントリーする魅力で採点する。ファンダは考慮しない。
        採点は下の配点表で項目ごとに整数で行い、基準に沿って厳しめにつけ、銘柄ごとの違いをはっきり出す(◎を通った銘柄の平均が50点前後になるイメージ。満点級は滅多に出さない)。
        - pullback(押し目の質, 0〜25): 短期線との距離が0〜0.3ATRで20〜25、0.3〜1ATRで12〜19、1〜2ATRで5〜11。直近週足の終値が一直線に上がっていれば減点、押して反発した形なら加点。
        - upside(上値余地, 0〜25): RR3以上、または高値更新中で52週高値より上なら20〜25、RR2〜3で13〜19、RR1.5〜2で5〜12。52週高値までの距離が近いほど減点。
        - volume(出来高の傾向, 0〜25): 直近の出来高が増加傾向(5日÷20日が1.1超)かつ上昇日の出来高が下落日より多いなら18〜25、横ばいなら9〜17、細っている(5日÷20日が0.8未満など)なら0〜8。
        - heat(過熱感の無さ, 0〜15): RSIとStochが低いほど高得点。RSI60超またはStoch70超は0〜6、50前後は7〜11、45以下は12〜15。
        - trend(トレンドの強さ, 0〜10): 移動平均の傾き(5週間の変化)が大きく、現在値が平均線の上にあり、他の時間軸も◎なら高得点。傾きが弱い・横ばいなら低得点。
        comment は、初心者向けに2文・合計100字程度で、結論(買ってよさそう/様子見/見送り寄り)→理由・注意点の順。RSI・Stoch・ATR・RRなどの専門用語は『買われすぎの度合い』『平均より上がりすぎていないか』『損失に対する利益の見込み』『取引量』のようにやさしい言葉に言い換える。装飾記号は使わない。
        出力は次のJSONのみ(前後に文章やコードブロックを付けない): {"pullback":整数,"upside":整数,"volume":整数,"heat":整数,"trend":整数,"comment":"..."}
        """;
    static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(60) };

    public static async Task<Dictionary<string, (string Score, string Text)>> Check(List<Screen.Result> hits)
    {
        var targets = hits.Where(x => x.Hits.ContainsKey("週足")).ToList();
        var res = new Dictionary<string, (string Score, string Text)>();
        var key = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
        if (targets.Count == 0) return res;
        if (string.IsNullOrEmpty(key)) { Console.Error.WriteLine("ANTHROPIC_API_KEY 未設定のためAIチェックをスキップ"); return res; }
        Console.Error.WriteLine($"AIチェック中 ({targets.Count}銘柄)...");
        await Parallel.ForEachAsync(targets, new ParallelOptions { MaxDegreeOfParallelism = 4 }, async (x, _) =>
        {
            (string Score, string Text) r;
            try
            {
                try { r = Parse(await Ask(key, x)); }
                catch (Exception e) when (e is JsonException or FormatException) { r = Parse(await Ask(key, x)); } // 稀にJSONが崩れるので1回再試行
            }
            catch (Exception e) { r = ("-", "AIチェック不可: " + e.Message); }
            lock (res) res[x.Sym] = (r.Score, r.Text.Replace("|", "/").Replace("*", "").ReplaceLineEndings(" ").Trim());
        });
        return res;
    }

    // 返答から最初の { 〜 最後の } を取り出して score / comment を読む(コードブロックで囲まれても動く)
    static (string Score, string Text) Parse(string s)
    {
        int a = s.IndexOf('{'), b = s.LastIndexOf('}');
        if (a < 0 || b < a) throw new FormatException("JSONなし");
        using var d = JsonDocument.Parse(s[a..(b + 1)]);
        // 項目ごとの点数を上限で切って合計する(AIの足し算ミスや配点超過を防ぐ)
        (string Key, string Label, int Max)[] items = [("pullback", "押", 25), ("upside", "余", 25), ("volume", "出", 25), ("heat", "熱", 15), ("trend", "勢", 10)];
        var pts = items.Select(i => d.RootElement.TryGetProperty(i.Key, out var p) && p.TryGetInt32(out var n) ? Math.Clamp(n, 0, i.Max) : throw new FormatException(i.Key + "なし")).ToArray();
        var detail = string.Join(" ", items.Zip(pts, (i, p) => $"{i.Label}{p}"));
        return ($"<b>{pts.Sum()}</b><br><small>{detail}</small>", d.RootElement.TryGetProperty("comment", out var c) ? c.GetString() ?? "" : "");
    }

    static async Task<string> Ask(string key, Screen.Result x)
    {
        var w = x.Hits["週足"];
        var prompt = $"{x.Sym} {x.Name} 終値{x.Close:#,0.##}\n週足: RSI{w.Rsi:F0} Stoch{w.Stoch:F0} 短期線乖離{w.AtrDist:+0.0;-0.0}ATR 損切り{w.Stop:#,0.##} "
            + (w.Target is { } t ? $"目標{t:#,0.##} RR{w.Rr:F1}" : "高値更新中(上値抵抗なし)")
            + $"\n{x.Ctx}\n他に◎の時間軸: {string.Join(",", x.Hits.Keys.Where(k => k != "週足").DefaultIfEmpty("なし"))}";
        var body = JsonSerializer.Serialize(new { model = Model, max_tokens = 1500, system = System, messages = new[] { new { role = "user", content = prompt } } });
        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages") { Content = new StringContent(body, Encoding.UTF8, "application/json") };
        req.Headers.Add("x-api-key", key); req.Headers.Add("anthropic-version", "2023-06-01");
        using var rsp = await Http.SendAsync(req);
        var json = await rsp.Content.ReadAsStringAsync();
        if (!rsp.IsSuccessStatusCode) throw new Exception($"{(int)rsp.StatusCode}");
        using var doc = JsonDocument.Parse(json);
        // 先頭が thinking ブロックのことがあるので text ブロックを探す
        return string.Concat(doc.RootElement.GetProperty("content").EnumerateArray().Where(b => b.GetProperty("type").GetString() == "text").Select(b => b.GetProperty("text").GetString()));
    }
}

static class Screen
{
    public record Hit(double Rsi, double Stoch, double AtrDist, double Stop, double? Target, double? Rr);
    public record Result(string Sym, string Name, DateOnly Last, double Close, Dictionary<string, Hit> Hits, Dictionary<string, string> Drops, string Ctx);

    static async Task<List<(string code, string name)>> PrimeList()
    {
        System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");
        using var ms = new MemoryStream(await http.GetByteArrayAsync("https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xlsx"));
        using var rd = ExcelReaderFactory.CreateReader(ms);
        var list = new List<(string, string)>(); rd.Read(); // ヘッダ行
        while (rd.Read())
            if (rd.GetValue(3)?.ToString()?.StartsWith("プライム") == true) list.Add((rd.GetValue(1)!.ToString()!, rd.GetValue(2)?.ToString() ?? ""));
        return list;
    }

    // 基本4条件を満たさなければ null、精査で落ちたら理由(string)、全通過なら Hit
    static object? Judge(Series s, int sh, int md, double[]? v = null)
    {
        var c = s.C; int n = c.Length;
        if (n < Math.Max(md + 5, 61)) return null;
        double sm = Ind.Sma(c, sh), sm5 = c[(n - sh - 5)..(n - 5)].Average(), mm = Ind.Sma(c, md), mm5 = c[(n - md - 5)..(n - 5)].Average();
        if (!(sm > mm && sm > sm5 && mm > mm5 && c[^1] > mm)) return null;
        var hist = Ind.MacdHist(c);
        if (!(hist[^1] > 0 && hist[(n - 4)..(n - 1)].Min() <= 0)) return null;
        double r = Ind.Rsi(c), st = Ind.Stoch(s);
        if (r >= 70 || st >= 80) return null;
        double atr = Ind.Atr(s), d = (c[^1] - sm) / atr;
        if (d > 2) return null;
        double stop = sm - atr, res = s.H[(n - 61)..(n - 1)].Max(); // 損切り=短期線-1ATR、上値抵抗=直近60本高値(当日除く)
        double up5 = (c[^1] - c[^6]) / atr;
        if (up5 > 3) return $"急騰(5本で+{up5:F1}ATR)";
        double rr = (res - c[^1]) / (c[^1] - stop);
        if (c[^1] < res && rr < 1.5) return $"上値抵抗{res:G6}が近い(RR{rr:F1})";
        if (v != null)
        {
            double vr = v[^5..].Average() / Math.Max(v[^20..].Average(), 1);
            if (vr < 0.7) return $"出来高細り(5/20日比{vr:F2})";
        }
        bool up = c[^1] >= res;
        return new Hit(r, st, d, stop, up ? null : res, up ? null : rr);
    }

    // AIチェック用の判断材料(出来高の推移・上昇日/下落日の出来高・週足の推移・52週高値までの距離・移動平均の傾き)
    static string Context(List<Bar> rows, Series w)
    {
        var v = rows.Select(r => r.V).ToArray();
        double Avg(int n) => Math.Max(v[^n..].Average(), 1);
        var weeks = rows.GroupBy(r => (ISOWeek.GetYear(r.D.ToDateTime(default)), ISOWeek.GetWeekOfYear(r.D.ToDateTime(default))))
            .TakeLast(6).Select(g => $"{g.Average(r => r.V) / 10000:#,0.#}");
        var last20 = Enumerable.Range(rows.Count - 20, 20).ToList();
        double Mean(IEnumerable<double> x) => x.Any() ? x.Average() : 0;
        var upV = Mean(last20.Where(i => rows[i].C > rows[i - 1].C).Select(i => rows[i].V));
        var dnV = Mean(last20.Where(i => rows[i].C < rows[i - 1].C).Select(i => rows[i].V));
        var c = w.C; int n = c.Length;
        double hi = w.H[^52..].Max(), lo = w.L[^52..].Min();
        string Slope(int k) { var now = Ind.Sma(c, k); var prev = c[(n - k - 5)..(n - 5)].Average(); return $"{(now / prev - 1) * 100:+0.0;-0.0}%"; }
        return $"出来高(日平均,万株) 直近6週: {string.Join("→", weeks)} / 5日÷20日={Avg(5) / Avg(20):F2} 20日÷60日={Avg(20) / Avg(60):F2}\n"
            + $"直近20日の平均出来高 上昇日÷下落日={(dnV > 0 ? (upV / dnV).ToString("F2") : "下落日なし")}\n"
            + $"週足終値 直近12週: {string.Join("→", c[^12..].Select(x => x.ToString("#,0.##")))}\n"
            + $"52週高値{hi:#,0.##}(現在値から{(c[^1] / hi - 1) * 100:+0.0;-0.0}%) 52週安値{lo:#,0.##}(現在値は安値から{(c[^1] / lo - 1) * 100:+0.0;-0.0}%)\n"
            + $"週足SMA13の5週間の変化{Slope(13)} SMA26の5週間の変化{Slope(26)} 現在値のSMA13比{(c[^1] / Ind.Sma(c, 13) - 1) * 100:+0.0;-0.0}% SMA26比{(c[^1] / Ind.Sma(c, 26) - 1) * 100:+0.0;-0.0}%";
    }
    static async Task<Result?> Scan((string code, string name) it, List<string> errs)
    {
        (string Sym, string Name, List<Bar> Rows) f;
        try { f = await Yahoo.Fetch(it.code); }
        catch
        {
            try { f = await Yahoo.Fetch(it.code); } // Yahoo の一時エラー対策で1回だけ再試行
            catch (Exception e) { lock (errs) errs.Add($"{it.code} {e.Message}"); return null; }
        }
        if (f.Rows.Count < 61) return null;
        var (d, w, m) = Ta.Frames(f.Rows);
        var js = new (string tf, object? j)[] { ("日足", Judge(d, 25, 75, f.Rows.Select(r => r.V).ToArray())), ("週足", Judge(w, 13, 26)), ("月足", Judge(m, 12, 24)) }; // 出来高は日足のみ
        var last = f.Rows[^1];
        return new(f.Sym, it.name, last.D, last.C,
            js.Where(x => x.j is Hit).ToDictionary(x => x.tf, x => (Hit)x.j!),
            js.Where(x => x.j is string).ToDictionary(x => x.tf, x => (string)x.j!), js[1].j is Hit ? Context(f.Rows, w) : ""); // AI用の材料は週足◎のみ(履歴が短い銘柄での範囲外も避ける)
    }

    // 数値から組み立てる定型の一言コメント(AIは使わない)
    static string Comment(Hit j)
    {
        var pos = j.AtrDist < 0.5 ? "短期線付近の押し目圏" : j.AtrDist < 1.2 ? "短期線から適度な位置" : "やや伸びており追いかけ注意";
        var room = j.Target is null ? "高値更新中" : j.Rr >= 3 ? $"上値余地大(RR{j.Rr:F1})" : j.Rr >= 2 ? $"上値余地あり(RR{j.Rr:F1})" : $"上値余地は限定的(RR{j.Rr:F1})";
        var heat = j.Rsi >= 65 || j.Stoch >= 75 ? "、過熱気味" : "";
        return $"{pos}、{room}{heat}";
    }

    // 表付きの HTML を直接生成する(リンクは新しいタブで開く)
    static string Html(DateOnly last, int scanned, int errors, List<Result> hits, Dictionary<string, (string Score, string Text)> ai)
    {
        const string css = "body{font-family:'Yu Gothic UI',sans-serif;max-width:1500px;margin:2em auto;padding:0 1em;line-height:1.6;font-size:14px;color:#222}"
            + "table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px 10px;vertical-align:top}th{background:#eef1f5;white-space:nowrap;position:sticky;top:0}"
            + "tr:nth-child(even){background:#fafafa}tbody tr:hover,tr:hover{background:#fff8dc}"
            + ".r{text-align:right}.c{text-align:center;color:#c0392b;font-weight:bold}.nw{white-space:nowrap}.score{white-space:nowrap;text-align:center}.score b{font-size:1.3em}small,.note{color:#555}";
        string E(string s) => WebUtility.HtmlEncode(s);
        string Join(Result x, Func<Hit, string> f) => string.Join(" / ", x.Hits.Select(kv => (x.Hits.Count > 1 ? kv.Key + ":" : "") + f(kv.Value)));
        var h = new StringBuilder();
        h.Append($"<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>プライム◎スクリーニング {last:yyyy-MM-dd}</title><style>{css}</style></head><body>");
        h.Append($"<h1>プライム◎スクリーニング ({last:yyyy-MM-dd})</h1><p>対象 {scanned} 銘柄 / ◎ <b>{hits.Count}</b> 銘柄 / 取得エラー {errors} 件</p><h2>◎銘柄</h2>");
        h.Append("<table><tr><th>symbol</th><th>銘柄名</th><th>終値</th><th>月足</th><th>週足</th><th>日足</th><th>損切り</th><th>目標(RR)</th><th>コメント</th><th>AIスコア</th><th>AIチェック(週足◎のみ)</th></tr>");
        foreach (var x in hits)
        {
            h.Append($"<tr><td>{x.Sym}</td><td class=\"nw\"><a target=\"_blank\" rel=\"noopener\" href=\"https://finance.yahoo.co.jp/quote/{x.Sym}/chart\">{E(x.Name)}</a></td><td class=\"r\">{x.Close:#,0.##}</td>");
            foreach (var t in new[] { "月足", "週足", "日足" }) h.Append($"<td class=\"c\">{(x.Hits.ContainsKey(t) ? "◎" : "")}</td>");
            h.Append($"<td class=\"r\">{Join(x, j => j.Stop.ToString("#,0"))}</td><td class=\"r nw\">{Join(x, j => j.Target is { } t ? $"{t:#,0.##} ({j.Rr:F1})" : "高値更新中")}</td><td>{E(Join(x, Comment))}</td>");
            h.Append(ai.TryGetValue(x.Sym, out var a) ? $"<td class=\"score\">{a.Score}</td><td>{E(a.Text)}</td>" : "<td></td><td></td>");
            h.Append("</tr>");
        }
        if (hits.Count == 0) h.Append("<tr><td colspan=\"11\">◎なし</td></tr>");
        h.Append($"</table><p class=\"note\">AIスコア内訳(100点満点): 押=押し目の質(25) 余=上値余地(25) 出=出来高の傾向(25) 熱=過熱感の無さ(15) 勢=トレンドの強さ(10)<br>AIモデル: {Ai.Model}</p>");
        h.Append("<hr><p class=\"note\">損切り=短期線-1ATR、目標=直近60本高値。ファンダ・決算日程は未考慮、投資判断は自己責任。</p></body></html>");
        return h.ToString();
    }

    static void Bar(int done, int total)
    {
        const int w = 40; int n = done * w / total;
        Console.Error.Write($"\r[{new string('#', n)}{new string('-', w - n)}] {done}/{total} ({done * 100 / total}%)");
    }

    // 既定ブラウザを調べて --new-window で開く(既存ウインドウのタブにしない)。判別できなければ通常オープン
    static void OpenInNewWindow(string path)
    {
        try
        {
            var prog = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice", "ProgId", null) as string;
            var cmd = Microsoft.Win32.Registry.GetValue($@"HKEY_CLASSES_ROOT\{prog}\shell\open\command", "", null) as string;
            var exe = System.Text.RegularExpressions.Regex.Match(cmd ?? "", "\"([^\"]+)\"").Groups[1].Value;
            var name = Path.GetFileNameWithoutExtension(exe).ToLowerInvariant();
            var flag = name is "chrome" or "msedge" or "brave" or "vivaldi" or "opera" ? "--new-window" : name == "firefox" ? "-new-window" : null;
            if (flag != null && File.Exists(exe)) { System.Diagnostics.Process.Start(exe, $"{flag} \"{new Uri(path).AbsoluteUri}\""); return; }
        }
        catch { } // shortcut: 上記以外のブラウザ(Arc等)は通常オープン(既存ウインドウのタブになる)
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
    }

    public static async Task<int> Run(int? limit)
    {
        var list = await PrimeList(); if (limit != null) list = list.Take(limit.Value).ToList();
        var errs = new List<string>(); var res = new Result?[list.Count]; int done = 0;
        Bar(0, list.Count);
        await Parallel.ForEachAsync(Enumerable.Range(0, list.Count), new ParallelOptions { MaxDegreeOfParallelism = int.Parse(Environment.GetEnvironmentVariable("STOCKS_PAR") ?? "24") }, async (i, _) =>
        {
            res[i] = await Scan(list[i], errs);
            lock (errs) Bar(++done, list.Count);
        });
        Console.Error.WriteLine();
        var ok = res.Where(x => x != null).Select(x => x!).ToList();
        var hits = ok.Where(x => x.Hits.Count > 0)
            .OrderByDescending(x => x.Hits.Count).ThenByDescending(x => x.Hits.Values.Max(j => j.Rr ?? 9)).ToList();
        if (ok.Count == 0) { Console.Error.WriteLine($"全件取得失敗(例: {errs.FirstOrDefault()})。レート制限の可能性があるので時間をおいて再実行してください"); return 1; }
        var last = ok.Max(x => x.Last);

        var dir = Path.Combine(AppContext.BaseDirectory, "analysis"); Directory.CreateDirectory(dir); // exe と同じフォルダ配下
        var path = Path.Combine(dir, $"{last:yyyy-MM-dd}.html");
        File.WriteAllText(path, Html(last, list.Count, errs.Count, hits, await Ai.Check(hits)), new UTF8Encoding(false));
        Console.WriteLine($"完了: ◎{hits.Count}銘柄 / エラー{errs.Count}件");
        foreach (var e in errs.Take(5)) Console.Error.WriteLine("ERR " + e);
        Console.WriteLine("保存先: " + path);
        OpenInNewWindow(path);
        return 0;
    }
}
