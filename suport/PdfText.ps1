#requires -Version 5.1
<#
  PdfText.ps1 - TREURE EL TEXT d'un PDF generat (sense cap biblioteca).

  Per al repas de contactes d'"Actualitzar base" (Contactes*.ps1): les
  instancies de la seu electronica i els formularis de l'e-TRAM son PDF
  GENERATS, amb text. Fa servir el mateix lector d'objectes que PdfUnio.ps1
  (PdfDoc: xref, fluxos d'objectes, FlateDecode amb predictor) -per que no
  PDFsharp ho explica la capcalera d'alla- i hi afegeix nomes el que cal per
  llegir: els operadors de text (Tj, TJ, ' i "), la posicio (Tm, Td, TD, T*,
  cm, q/Q) per tornar a fer les linies, i la ToUnicode (bfchar/bfrange) de
  cada font; sense ToUnicode, la codificacio (WinAnsi + Differences).

  QUE NO FA, i es a posta:
    - ESCANEJATS (nomes imatges): no hi ha text. Torna buit i qui crida ho
      posa a la llista "no llegibles" (l'OCR del Windows queda per mes
      endavant: l'usuari, "ara no").
    - XIFRATS: llanca. Qui crida prova el Word (_PdfTextWord).
    - Una font Identity-H SENSE ToUnicode dona glifs, no lletres: torna buit
      (i el Word, que si que ho sap fer).

  LES LINIES. Cada tros de text es desa amb la seva posicio a la pagina i, en
  acabar la pagina, s'ordena de dalt a baix i d'esquerra a dreta: els
  formularis posen l'etiqueta i el valor a la mateixa alcada pero en ordres
  qualssevol. Dos trossos de la mateixa linia van amb un espai si n'hi ha un
  forat (l'amplada surt de /Widths o /W de la font).

  El C# va DINS del namespace amb els seus 'using' a dins: es compila
  ENGANXAT al de PdfUnio (_PdfUnioCarrega), perque necessita PdfDoc i dos
  Add-Type separats no es poden referenciar en memoria al PowerShell 5.1.
  NOMES DEFINEIX: carregar-lo no compila res.
#>

$Script:PdfTextCs = @'
namespace InformesCornella
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Text;

    public sealed class PdfText
    {
        // ---- Una font: com es passa d'un codi a lletres, i quant ocupa ----
        sealed class Font
        {
            public int Bytes = 1;                 // 1 o 2 bytes per codi
            public Dictionary<int, string> Map;   // ToUnicode
            public string[] Base;                 // codificacio simple (256)
            public Dictionary<int, double> W = new Dictionary<int, double>();
            public double Dw = 500;
            public bool Glifs;                    // Identity sense ToUnicode: no se'n treu text

            public string Text(byte[] s, List<int> codis)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i + Bytes <= s.Length; i += Bytes)
                {
                    int c = Bytes == 2 ? (s[i] << 8) | s[i + 1] : s[i];
                    codis.Add(c);
                    string u;
                    if (Map != null && Map.TryGetValue(c, out u)) { sb.Append(u); continue; }
                    if (Glifs) continue;
                    if (Bytes == 1 && Base != null) { string t = Base[c]; if (t != null) sb.Append(t); continue; }
                    if (c >= 32 && c < 127) sb.Append((char)c);
                }
                return sb.ToString();
            }
            public double Amplada(int c) { double w; return W.TryGetValue(c, out w) ? w : Dw; }
        }

        sealed class Tros { public double X, Y, Fi, Mida; public string T; }

        readonly PdfDoc doc;
        readonly Dictionary<PdfObj, Font> fonts = new Dictionary<PdfObj, Font>();
        List<Tros> trossos;
        int prof;

        PdfText(PdfDoc d) { doc = d; }

        // El text de TOT el PDF: les linies amb \n i les pagines amb \f.
        public static string Extreu(string path)
        {
            PdfDoc d = new PdfDoc(path);
            if (d.Trailer.Get("Encrypt") != null) throw new Exception("PDF xifrat");
            PdfText t = new PdfText(d);
            StringBuilder sb = new StringBuilder();
            foreach (PdfDoc.Pagina pg in d.Pagines(new HashSet<int>()))
            {
                t.trossos = new List<Tros>();
                PdfDict res = d.Resol(pg.Heretat.Get("Resources")) as PdfDict;
                byte[] cont = t.Contingut(pg.Dict.Get("Contents"));
                try { t.prof = 0; t.Executa(cont, res, new double[] { 1, 0, 0, 1, 0, 0 }); }
                catch { }   // un contingut malmes: val el que s'hagi llegit
                if (sb.Length > 0) sb.Append('\f');
                sb.Append(t.Linies());
            }
            return sb.ToString();
        }

        byte[] Contingut(PdfObj c)
        {
            c = doc.Resol(c);
            if (c is PdfStream) return doc.Descodifica((PdfStream)c);
            PdfArr a = c as PdfArr;
            if (a == null) return new byte[0];
            List<byte> r = new List<byte>();
            foreach (PdfObj o in a.Items)
            {
                PdfStream s = doc.Resol(o) as PdfStream;
                if (s == null) continue;
                try { r.AddRange(doc.Descodifica(s)); r.Add(10); } catch { }
            }
            return r.ToArray();
        }

        // ---- Les linies de la pagina -------------------------------------
        string Linies()
        {
            List<Tros> l = new List<Tros>();
            foreach (Tros t in trossos) if (t.T.Trim().Length > 0) l.Add(t);
            // Ordre estable: per Y de dalt a baix; la mateixa linia, per X.
            List<List<Tros>> files = new List<List<Tros>>();
            List<Tros> perY = new List<Tros>(l);
            perY.Sort(delegate (Tros a, Tros b) { int c = b.Y.CompareTo(a.Y); return c != 0 ? c : a.X.CompareTo(b.X); });
            foreach (Tros t in perY)
            {
                List<Tros> f = files.Count > 0 ? files[files.Count - 1] : null;
                double tol = Math.Max(1.5, 0.45 * Math.Max(t.Mida, f != null ? f[0].Mida : 0));
                if (f != null && Math.Abs(f[0].Y - t.Y) <= tol) f.Add(t);
                else files.Add(new List<Tros>(new Tros[] { t }));
            }
            StringBuilder sb = new StringBuilder();
            foreach (List<Tros> f in files)
            {
                f.Sort(delegate (Tros a, Tros b) { return a.X.CompareTo(b.X); });
                StringBuilder ln = new StringBuilder();
                double fi = double.NaN;
                foreach (Tros t in f)
                {
                    if (ln.Length > 0)
                    {
                        bool forat = double.IsNaN(fi) || t.X - fi > 0.18 * Math.Max(1, t.Mida);
                        if (forat && ln[ln.Length - 1] != ' ' && t.T[0] != ' ') ln.Append(' ');
                    }
                    ln.Append(t.T);
                    fi = double.IsNaN(fi) ? t.Fi : Math.Max(fi, t.Fi);
                }
                sb.Append(ln.ToString().Trim()).Append('\n');
            }
            return sb.ToString();
        }

        // ---- L'interpret del contingut ------------------------------------
        static double[] Mul(double[] m, double[] n)
        {
            return new double[] {
                m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3],
                m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
                m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5] };
        }
        static double Num(PdfObj o) { PdfNum n = o as PdfNum; return n == null ? 0 : n.Val; }

        void Executa(byte[] c, PdfDict res, double[] ctm0)
        {
            if (++prof > 8) { prof--; return; }
            Stack<double[]> pila = new Stack<double[]>();
            double[] ctm = ctm0;
            double[] tm = { 1, 0, 0, 1, 0, 0 }, tlm = { 1, 0, 0, 1, 0, 0 };
            double tfs = 0, tc = 0, tw = 0, tz = 1, tl = 0, rise = 0;
            Font font = null;
            List<PdfObj> ops = new List<PdfObj>();
            PdfParser p = new PdfParser(c, 0);
            while (true)
            {
                p.SaltaBlancs();
                if (p.Final) break;
                PdfObj o;
                try { o = p.Objecte(); } catch { break; }
                PdfKw kw = o as PdfKw;
                if (kw == null || kw.Raw == "null" || kw.Raw == "true" || kw.Raw == "false") { ops.Add(o); continue; }
                string op = kw.Raw;
                switch (op)
                {
                    case "q": pila.Push(ctm); break;
                    case "Q": if (pila.Count > 0) ctm = pila.Pop(); break;
                    case "cm":
                        if (ops.Count >= 6) ctm = Mul(new double[] { Num(ops[0]), Num(ops[1]), Num(ops[2]), Num(ops[3]), Num(ops[4]), Num(ops[5]) }, ctm);
                        break;
                    case "BT": tm = new double[] { 1, 0, 0, 1, 0, 0 }; tlm = tm; break;
                    case "Tf":
                        if (ops.Count >= 2) { font = FontDe(res, ops[0] as PdfName); tfs = Num(ops[1]); }
                        break;
                    case "Tc": if (ops.Count >= 1) tc = Num(ops[0]); break;
                    case "Tw": if (ops.Count >= 1) tw = Num(ops[0]); break;
                    case "Tz": if (ops.Count >= 1) tz = Num(ops[0]) / 100.0; break;
                    case "TL": if (ops.Count >= 1) tl = Num(ops[0]); break;
                    case "Ts": if (ops.Count >= 1) rise = Num(ops[0]); break;
                    case "Td":
                    case "TD":
                        if (ops.Count >= 2)
                        {
                            if (op == "TD") tl = -Num(ops[1]);
                            tlm = Mul(new double[] { 1, 0, 0, 1, Num(ops[0]), Num(ops[1]) }, tlm); tm = tlm;
                        }
                        break;
                    case "Tm":
                        if (ops.Count >= 6) { tlm = new double[] { Num(ops[0]), Num(ops[1]), Num(ops[2]), Num(ops[3]), Num(ops[4]), Num(ops[5]) }; tm = tlm; }
                        break;
                    case "T*": tlm = Mul(new double[] { 1, 0, 0, 1, 0, -tl }, tlm); tm = tlm; break;
                    case "'":
                    case "\"":
                    case "Tj":
                    case "TJ":
                        if (op == "'" || op == "\"")
                        {
                            if (op == "\"" && ops.Count >= 3) { tw = Num(ops[0]); tc = Num(ops[1]); }
                            tlm = Mul(new double[] { 1, 0, 0, 1, 0, -tl }, tlm); tm = tlm;
                        }
                        if (ops.Count == 0 || font == null) break;
                        PdfObj arg = ops[ops.Count - 1];
                        List<PdfObj> peces = new List<PdfObj>();
                        if (arg is PdfArr) peces.AddRange(((PdfArr)arg).Items); else peces.Add(arg);
                        foreach (PdfObj pc in peces)
                        {
                            if (pc is PdfNum)
                            {
                                double adj = -((PdfNum)pc).Val / 1000.0 * tfs * tz;
                                // Un ajust gran dins d'un TJ es un espai entre paraules.
                                if (((PdfNum)pc).Val < -180 && trossos.Count > 0) { Tros ult = trossos[trossos.Count - 1]; if (!ult.T.EndsWith(" ")) ult.T += " "; }
                                tm = Mul(new double[] { 1, 0, 0, 1, adj, 0 }, tm);
                                continue;
                            }
                            PdfStr st = pc as PdfStr;
                            if (st == null) continue;
                            List<int> codis = new List<int>();
                            string txt = font.Text(Bytes(st.Raw), codis);
                            double[] m = Mul(new double[] { tfs * tz, 0, 0, tfs, 0, rise }, Mul(tm, ctm));
                            double avanc = 0;
                            foreach (int cd in codis)
                            {
                                double w = font.Amplada(cd) / 1000.0 * tfs + tc;
                                if (font.Bytes == 1 && cd == 32) w += tw;
                                avanc += w * tz;
                            }
                            double[] fiM = Mul(new double[] { 1, 0, 0, 1, avanc, 0 }, Mul(tm, ctm));
                            double mida = Math.Sqrt(Math.Abs(m[0] * m[3] - m[1] * m[2]));
                            if (txt.Length > 0)
                            {
                                Tros t = new Tros(); t.X = m[4]; t.Y = m[5]; t.Fi = fiM[4]; t.Mida = mida; t.T = txt;
                                trossos.Add(t);
                            }
                            tm = Mul(new double[] { 1, 0, 0, 1, avanc, 0 }, tm);
                        }
                        break;
                    case "Do":
                        if (ops.Count >= 1) Forma(res, ops[0] as PdfName, ctm);
                        break;
                    case "BI":
                        // Imatge en linia: les dades binaries van fins a "EI".
                        int id = PdfDoc.IndexOf(c, "ID", p.P, c.Length);
                        if (id < 0) { p.P = c.Length; break; }
                        int ei = id + 3;
                        while (true)
                        {
                            ei = PdfDoc.IndexOf(c, "EI", ei, c.Length);
                            if (ei < 0) { ei = c.Length; break; }
                            bool abans = ei > 0 && PdfParser.EsBlanc(c[ei - 1]);
                            bool despres = ei + 2 >= c.Length || PdfParser.EsBlanc(c[ei + 2]);
                            if (abans && despres) break;
                            ei += 2;
                        }
                        p.P = Math.Min(c.Length, ei + 2);
                        break;
                }
                ops.Clear();
            }
            prof--;
        }

        void Forma(PdfDict res, PdfName nom, double[] ctm)
        {
            if (res == null || nom == null) return;
            PdfDict xo = doc.Resol(res.Get("XObject")) as PdfDict;
            if (xo == null) return;
            PdfStream s = doc.Resol(xo.Get(nom.Raw)) as PdfStream;
            if (s == null) return;
            PdfName st = doc.Resol(s.Dict.Get("Subtype")) as PdfName;
            if (st == null || st.Raw != "Form") return;
            PdfArr ma = doc.Resol(s.Dict.Get("Matrix")) as PdfArr;
            double[] m = { 1, 0, 0, 1, 0, 0 };
            if (ma != null && ma.Items.Count >= 6) for (int i = 0; i < 6; i++) m[i] = Num(doc.Resol(ma.Items[i]));
            PdfDict r2 = doc.Resol(s.Dict.Get("Resources")) as PdfDict;
            byte[] d;
            try { d = doc.Descodifica(s); } catch { return; }
            Executa(d, r2 ?? res, Mul(m, ctm));
        }

        // ---- Les cadenes ---------------------------------------------------
        // PdfStr.Raw porta els delimitadors: "(...)" o "<...>".
        public static byte[] Bytes(byte[] raw)
        {
            List<byte> r = new List<byte>();
            if (raw.Length == 0) return r.ToArray();
            if (raw[0] == (byte)'<')
            {
                int hi = -1;
                for (int i = 1; i < raw.Length && raw[i] != (byte)'>'; i++)
                {
                    int v = Hex(raw[i]);
                    if (v < 0) continue;
                    if (hi < 0) hi = v; else { r.Add((byte)(hi * 16 + v)); hi = -1; }
                }
                if (hi >= 0) r.Add((byte)(hi * 16));
                return r.ToArray();
            }
            int fi = raw.Length - 1;
            if (fi > 0 && raw[fi] != (byte)')') fi = raw.Length;
            for (int i = 1; i < fi; i++)
            {
                byte b = raw[i];
                if (b != (byte)'\\') { r.Add(b); continue; }
                if (++i >= fi) break;
                byte e = raw[i];
                switch ((char)e)
                {
                    case 'n': r.Add(10); break;
                    case 'r': r.Add(13); break;
                    case 't': r.Add(9); break;
                    case 'b': r.Add(8); break;
                    case 'f': r.Add(12); break;
                    case '\r': if (i + 1 < fi && raw[i + 1] == 10) i++; break;
                    case '\n': break;
                    default:
                        if (e >= (byte)'0' && e <= (byte)'7')
                        {
                            int v = e - '0'; int n = 1;
                            while (n < 3 && i + 1 < fi && raw[i + 1] >= (byte)'0' && raw[i + 1] <= (byte)'7') { v = v * 8 + (raw[++i] - '0'); n++; }
                            r.Add((byte)(v & 0xFF));
                        }
                        else r.Add(e);
                        break;
                }
            }
            return r.ToArray();
        }
        static int Hex(byte c)
        {
            if (c >= (byte)'0' && c <= (byte)'9') return c - '0';
            if (c >= (byte)'a' && c <= (byte)'f') return c - 'a' + 10;
            if (c >= (byte)'A' && c <= (byte)'F') return c - 'A' + 10;
            return -1;
        }

        // ---- Les fonts -----------------------------------------------------
        Font FontDe(PdfDict res, PdfName nom)
        {
            if (res == null || nom == null) return null;
            PdfDict fd = doc.Resol(res.Get("Font")) as PdfDict;
            if (fd == null) return null;
            PdfObj clau = fd.Get(nom.Raw);
            if (clau == null) return null;
            PdfObj fo = doc.Resol(clau);
            if (fo == null) return null;
            Font f;
            if (fonts.TryGetValue(fo, out f)) return f;
            f = new Font();
            fonts[fo] = f;
            PdfDict d = fo as PdfDict;
            if (d == null) return f;
            PdfName sub = doc.Resol(d.Get("Subtype")) as PdfName;
            bool type0 = sub != null && sub.Raw == "Type0";
            if (type0)
            {
                f.Bytes = 2;
                PdfArr desc = doc.Resol(d.Get("DescendantFonts")) as PdfArr;
                PdfDict cid = desc != null && desc.Items.Count > 0 ? doc.Resol(desc.Items[0]) as PdfDict : null;
                if (cid != null)
                {
                    PdfNum dw = doc.Resol(cid.Get("DW")) as PdfNum; f.Dw = dw == null ? 1000 : dw.Val;
                    LlegeixW(f, doc.Resol(cid.Get("W")) as PdfArr);
                }
            }
            else
            {
                f.Base = Codificacio(d);
                PdfNum fc = doc.Resol(d.Get("FirstChar")) as PdfNum;
                PdfArr ws = doc.Resol(d.Get("Widths")) as PdfArr;
                if (fc != null && ws != null)
                    for (int i = 0; i < ws.Items.Count; i++) f.W[(int)fc.Val + i] = Num(doc.Resol(ws.Items[i]));
            }
            PdfStream tu = doc.Resol(d.Get("ToUnicode")) as PdfStream;
            if (tu != null)
            {
                try { int bytes; f.Map = CMap(doc.Descodifica(tu), out bytes); if (bytes > 0) f.Bytes = bytes; }
                catch { f.Map = null; }
            }
            if (type0 && f.Map == null) f.Glifs = true;
            return f;
        }

        void LlegeixW(Font f, PdfArr w)
        {
            if (w == null) return;
            int i = 0;
            while (i < w.Items.Count)
            {
                PdfObj a = doc.Resol(w.Items[i]);
                if (!(a is PdfNum) || i + 1 >= w.Items.Count) break;
                int c0 = (int)((PdfNum)a).Val;
                PdfObj b = doc.Resol(w.Items[i + 1]);
                if (b is PdfArr)
                {
                    PdfArr l = (PdfArr)b;
                    for (int k = 0; k < l.Items.Count; k++) f.W[c0 + k] = Num(doc.Resol(l.Items[k]));
                    i += 2;
                }
                else
                {
                    if (i + 2 >= w.Items.Count) break;
                    int c1 = (int)Num(b); double v = Num(doc.Resol(w.Items[i + 2]));
                    for (int k = c0; k <= c1 && k - c0 < 65536; k++) f.W[k] = v;
                    i += 3;
                }
            }
        }

        // La ToUnicode: bfchar i bfrange. Torna tambe els bytes per codi que
        // diu el codespacerange (0 si no en diu res).
        public static Dictionary<int, string> CMap(byte[] data, out int bytes)
        {
            Dictionary<int, string> m = new Dictionary<int, string>();
            bytes = 0;
            PdfParser p = new PdfParser(data, 0);
            List<PdfObj> ops = new List<PdfObj>();
            while (true)
            {
                p.SaltaBlancs();
                if (p.Final) break;
                PdfObj o;
                try { o = p.Objecte(); } catch { break; }
                PdfKw kw = o as PdfKw;
                if (kw == null) { ops.Add(o); continue; }
                string k = kw.Raw;
                if (k == "begincodespacerange" || k == "beginbfchar" || k == "beginbfrange") { ops.Clear(); continue; }
                if (k == "endcodespacerange")
                {
                    foreach (PdfObj x in ops) { PdfStr s = x as PdfStr; if (s != null) bytes = Math.Max(bytes, Bytes(s.Raw).Length); }
                }
                else if (k == "endbfchar")
                {
                    for (int i = 0; i + 1 < ops.Count; i += 2)
                    {
                        PdfStr a = ops[i] as PdfStr; PdfStr b = ops[i + 1] as PdfStr;
                        if (a == null || b == null) continue;
                        m[Codi(Bytes(a.Raw))] = Utf16(Bytes(b.Raw));
                    }
                }
                else if (k == "endbfrange")
                {
                    for (int i = 0; i + 2 < ops.Count; i += 3)
                    {
                        PdfStr a = ops[i] as PdfStr; PdfStr b = ops[i + 1] as PdfStr;
                        if (a == null || b == null) continue;
                        int c0 = Codi(Bytes(a.Raw)), c1 = Codi(Bytes(b.Raw));
                        if (c1 < c0 || c1 - c0 > 65535) continue;
                        if (ops[i + 2] is PdfArr)
                        {
                            List<PdfObj> l = ((PdfArr)ops[i + 2]).Items;
                            for (int c = c0; c <= c1 && c - c0 < l.Count; c++) { PdfStr d = l[c - c0] as PdfStr; if (d != null) m[c] = Utf16(Bytes(d.Raw)); }
                        }
                        else
                        {
                            PdfStr d = ops[i + 2] as PdfStr;
                            if (d == null) continue;
                            byte[] bd = Bytes(d.Raw);
                            for (int c = c0; c <= c1; c++)
                            {
                                byte[] x = (byte[])bd.Clone();
                                int inc = c - c0, j = x.Length - 1;
                                while (inc > 0 && j >= 0) { int v = x[j] + (inc & 0xFF); x[j] = (byte)v; inc = (inc >> 8) + (v >> 8); j--; }
                                m[c] = Utf16(x);
                            }
                        }
                    }
                }
                ops.Clear();
            }
            return m;
        }
        static int Codi(byte[] b) { int c = 0; foreach (byte x in b) c = (c << 8) | x; return c; }
        static string Utf16(byte[] b)
        {
            if (b.Length == 1) return ((char)b[0]).ToString();
            try { return Encoding.BigEndianUnicode.GetString(b); } catch { return ""; }
        }

        // ---- Fonts simples sense ToUnicode: WinAnsi + /Differences --------
        static readonly int[] Ansi80 = {
            0x20AC, 0, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0, 0x017D, 0,
            0, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0, 0x017E, 0x0178 };
        static string[] WinAnsi()
        {
            string[] t = new string[256];
            for (int i = 32; i < 256; i++)
            {
                int u = i >= 0x80 && i < 0xA0 ? Ansi80[i - 0x80] : i;
                if (u != 0) t[i] = ((char)u).ToString();
            }
            return t;
        }
        static Dictionary<string, string> glifs;
        static string Glif(string n)
        {
            if (glifs == null)
            {
                glifs = new Dictionary<string, string>();
                string[] noms = {
                    "space"," ","exclam","!","quotedbl","\"","numbersign","#","dollar","$","percent","%","ampersand","&",
                    "quotesingle","'","quoteright","’","quoteleft","‘","parenleft","(","parenright",")","asterisk","*",
                    "plus","+","comma",",","hyphen","-","minus","-","period",".","slash","/","colon",":","semicolon",";",
                    "less","<","equal","=","greater",">","question","?","at","@","bracketleft","[","backslash","\\",
                    "bracketright","]","underscore","_","braceleft","{","bar","|","braceright","}","endash","–",
                    "emdash","—","bullet","•","periodcentered","·","middot","·","degree","°",
                    "ordfeminine","ª","ordmasculine","º","quotedblleft","“","quotedblright","”",
                    "guillemotleft","«","guillemotright","»","euro","€","nbspace"," ","uni00A0"," ",
                    "zero","0","one","1","two","2","three","3","four","4","five","5","six","6","seven","7","eight","8","nine","9" };
                for (int i = 0; i + 1 < noms.Length; i += 2) glifs[noms[i]] = noms[i + 1];
                string acc = "aAeEiIoOuU";
                foreach (char v in acc)
                {
                    glifs[v + "acute"] = Composa(v, '́'); glifs[v + "grave"] = Composa(v, '̀');
                    glifs[v + "dieresis"] = Composa(v, '̈'); glifs[v + "circumflex"] = Composa(v, '̂');
                }
                glifs["ccedilla"] = "ç"; glifs["Ccedilla"] = "Ç"; glifs["ntilde"] = "ñ"; glifs["Ntilde"] = "Ñ";
            }
            string r;
            if (glifs.TryGetValue(n, out r)) return r;
            if (n.Length == 1) return n;
            if (n.StartsWith("uni") && n.Length == 7)
            {
                int u;
                if (int.TryParse(n.Substring(3), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out u)) return ((char)u).ToString();
            }
            return null;
        }
        static string Composa(char v, char marca) { return (v.ToString() + marca).Normalize(NormalizationForm.FormC); }

        string[] Codificacio(PdfDict d)
        {
            string[] t = WinAnsi();
            PdfObj e = doc.Resol(d.Get("Encoding"));
            PdfDict ed = e as PdfDict;
            if (ed == null) return t;
            PdfArr dif = doc.Resol(ed.Get("Differences")) as PdfArr;
            if (dif == null) return t;
            int c = 0;
            foreach (PdfObj o in dif.Items)
            {
                PdfObj x = doc.Resol(o);
                if (x is PdfNum) { c = (int)((PdfNum)x).Val; continue; }
                PdfName n = x as PdfName;
                if (n == null || c < 0 || c > 255) { c++; continue; }
                string g = Glif(n.Raw);
                if (g != null) t[c] = g;
                c++;
            }
            return t;
        }
    }
}
'@

# El text d'un PDF amb el lector propi. Torna '' si no n'hi ha (escanejat, o
# glifs sense ToUnicode); LLANCA si el PDF es xifrat o no es pot llegir.
function Get-PdfText([string]$pdf) {
    _PdfUnioCarrega
    return [string][InformesCornella.PdfText]::Extreu($pdf)
}

# El RECURS AL WORD, per als pocs que no surten amb el lector propi (xifrats,
# fonts sense ToUnicode). El Word converteix el PDF a document
# (ConfirmConversions = $false: sense preguntar) i se'n llegeix el text. Lent:
# nomes per a aquests. $wordApp el porta qui crida (New-WordApp -Opcional, amb
# la Quit() al seu finally); sense, torna ''.
function _PdfTextWord($wordApp, [string]$pdf) {
    if ($null -eq $wordApp) { return '' }
    $doc = $null
    try {
        $doc = $wordApp.Documents.Open($pdf, $false, $true, $false)
        return [string]$doc.Content.Text
    } catch { return '' }
    finally { if ($null -ne $doc) { try { $doc.Close($false) } catch { } } }
}

# Hi ha prou text per fer-ne res? Un escanejat pot portar quatre lletres d'una
# capa d'OCR trencada o el numero de pagina. PURA.
function _PdfTeText([string]$t) {
    return (([regex]::Matches([string]$t, '[\p{L}\d]')).Count -ge 40)
}
