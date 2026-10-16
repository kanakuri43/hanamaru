// tech-analysis / prime-hanamaru スキル(ta.py / screen.py)の C# 移植
//   stocks ta <コード>...   指標を出力し analysis/<日付>/<symbol>_<名前>.md に保存
//   stocks screen [上限]    プライム全銘柄を◎判定し analysis/<日付>/prime_hanamaru.md に保存
using System.Globalization;
using System.Text;
using System.Text.Json;
using ExcelDataReader;

CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
Console.OutputEncoding = new UTF8Encoding(false);

if (args.Length == 0) args = ["screen"]; // 省略時は screen
if (args[0] is not ("ta" or "screen"))
{
    Console.WriteLine("usage: stocks ta <code>...  |  stocks screen [limit]");
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

static class Screen
{
    record Hit(double Rsi, double Stoch, double AtrDist, double Stop, double? Target, double? Rr);
    record Result(string Sym, string Name, DateOnly Last, double Close, Dictionary<string, Hit> Hits, Dictionary<string, string> Drops);

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
            js.Where(x => x.j is string).ToDictionary(x => x.tf, x => (string)x.j!));
    }

    // 数値から組み立てる定型の一言コメント(AIは使わない)
    static string Comment(Hit j)
    {
        var pos = j.AtrDist < 0.5 ? "短期線付近の押し目圏" : j.AtrDist < 1.2 ? "短期線から適度な位置" : "やや伸びており追いかけ注意";
        var room = j.Target is null ? "高値更新中" : j.Rr >= 3 ? $"上値余地大(RR{j.Rr:F1})" : j.Rr >= 2 ? $"上値余地あり(RR{j.Rr:F1})" : $"上値余地は限定的(RR{j.Rr:F1})";
        var heat = j.Rsi >= 65 || j.Stoch >= 75 ? "、過熱気味" : "";
        return $"{pos}、{room}{heat}";
    }

    static string Markdown(DateOnly last, int scanned, int errors, List<Result> hits, List<Result> drops)
    {
        string[] tfs = ["月足", "週足", "日足"];
        string Join(Result x, Func<Hit, string> f) => hits.Count == 0 ? "" : string.Join(" / ", x.Hits.Select(kv => (x.Hits.Count > 1 ? kv.Key + ":" : "") + f(kv.Value)));
        var m = new StringBuilder();
        m.AppendLine($"# プライム◎スクリーニング ({last:yyyy-MM-dd})\n");
        m.AppendLine($"- 対象 {scanned} 銘柄 / ◎ **{hits.Count}** 銘柄 / 精査で除外 {drops.Count} 銘柄 / 取得エラー {errors} 件\n");
        m.AppendLine("## ◎銘柄\n");
        m.AppendLine("| symbol | 銘柄名 | 終値 | 月足 | 週足 | 日足 | RSI | 短期線乖離 | 損切り | 目標(RR) | コメント |");
        m.AppendLine("|---|---|--:|:-:|:-:|:-:|--:|--:|--:|--:|---|");
        foreach (var x in hits)
            m.AppendLine($"| {x.Sym} | {x.Name} | {x.Close:#,0.##} | " + string.Join(" | ", tfs.Select(t => x.Hits.ContainsKey(t) ? "◎" : ""))
                + $" | {Join(x, j => j.Rsi.ToString("F0"))} | {Join(x, j => $"{Math.Round(j.AtrDist, 1) + 0.0:+0.0;-0.0}ATR")} | {Join(x, j => j.Stop.ToString("#,0"))}"
                + $" | {Join(x, j => j.Target is { } t ? $"{t:#,0.##} ({j.Rr:F1})" : "高値更新中")} | {Join(x, Comment)} |");
        if (hits.Count == 0) m.AppendLine("| ◎なし | | | | | | | | | | |");
        m.AppendLine("\n## 精査で除外(基本条件は満たしたが…)\n");
        m.AppendLine("| symbol | 銘柄名 | 時間軸 | 理由 |\n|---|---|:-:|---|");
        foreach (var x in drops) foreach (var kv in x.Drops) m.AppendLine($"| {x.Sym} | {x.Name} | {kv.Key} | {kv.Value} |");
        m.AppendLine("\n---\n損切り=短期線-1ATR、目標=直近60本高値。ファンダ・決算日程は未考慮、投資判断は自己責任。");
        return m.ToString();
    }

    static void Bar(int done, int total)
    {
        const int w = 40; int n = done * w / total;
        Console.Error.Write($"\r[{new string('#', n)}{new string('-', w - n)}] {done}/{total} ({done * 100 / total}%)");
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
        var drops = ok.Where(x => x.Drops.Count > 0 && x.Hits.Count == 0).ToList();
        if (ok.Count == 0) { Console.Error.WriteLine($"全件取得失敗(例: {errs.FirstOrDefault()})。レート制限の可能性があるので時間をおいて再実行してください"); return 1; }
        var last = ok.Max(x => x.Last);

        var dir = Path.Combine("analysis", last.ToString("yyyyMMdd")); Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "prime_hanamaru.md");
        var md = Markdown(last, list.Count, errs.Count, hits, drops);
        File.WriteAllText(path, md, new UTF8Encoding(false));
        Console.WriteLine($"完了: ◎{hits.Count}銘柄 / 除外{drops.Count}銘柄 / エラー{errs.Count}件");
        foreach (var e in errs.Take(5)) Console.Error.WriteLine("ERR " + e);
        Console.WriteLine("保存先: " + path);
        return 0;
    }
}
