#requires -Version 5.1
<#
.SYNOPSIS
  Ajuntar PDF: l'informe de llicencia + els informes dels organismes que posen
  les condicions (OGAU, Agencia de Residus...), adjunts darrere.

.DESCRIPTION
  PER QUE CODI PROPI I NO UNA LLIBRERIA. Es va provar PDFsharp 1.50 i 1.51 (una
  sola DLL de .NET Framework): amb un PDF SIGNAT en revisio incremental sobre un
  flux d'objectes comprimit -el format que fan servir els signadors- llegeix la
  versio ANTIGA de la pagina i la signatura DESAPAREIX sense cap error. PDFsharp
  6 si que ho llegeix be, pero per a .NET Framework arrossega mitja dotzena de
  DLL (Logging.Abstractions, Cryptography.Pkcs, System.Memory...) amb conflictes
  de versions coneguts dins del PowerShell 5.1. El que cal aqui es poc i concret
  -llegir, copiar pagines, escriure-, i escrit aqui es pot provar a fons.

  LES SIGNATURES DELS ADJUNTS S'APLANEN. Un camp de signatura copiat a un PDF
  nou ja no quadra amb els bytes que va signar: l'Adobe diria "signatura no
  valida". Per aixo l'APARENCA de cada signatura visible passa al CONTINGUT de
  la pagina (es veu exactament igual) i el camp desapareix; les invisibles
  simplement es treuen. Els originals, signats i valids, queden a la carpeta de
  la llicencia (local\base-dades-llicencies\GIA <id>).

  El PDF que surt es net (una sola revisio, taula xref classica): es pot signar
  despres amb l'Adobe o amb l'AutoFirma com qualsevol altre.

  El C# es compila EN VIU el primer cop que es fa servir (Add-Type), en C# 5 que
  es el que compila el PowerShell 5.1. Aquest fitxer NOMES DEFINEIX: carregar-lo
  no compila res.
#>

$Script:PdfUnioCs = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace InformesCornella
{
    // ---- El model d'objectes d'un PDF -------------------------------------
    // Els valors ATOMICS es guarden amb el text ORIGINAL (Raw): un nom amb
    // #xx, un numero 1.500 o una cadena amb escapades surten exactament com
    // han entrat. Aixi no hi ha cap conversio que pugui canviar res.
    public abstract class PdfObj { }
    public sealed class PdfName : PdfObj { public string Raw; public PdfName(string r) { Raw = r; } }
    public sealed class PdfNum : PdfObj
    {
        public string Raw;
        public PdfNum(string r) { Raw = r; }
        public double Val
        {
            get
            {
                double d;
                double.TryParse(Raw, NumberStyles.Float, CultureInfo.InvariantCulture, out d);
                return d;
            }
        }
    }
    public sealed class PdfStr : PdfObj { public byte[] Raw; public PdfStr(byte[] r) { Raw = r; } }
    public sealed class PdfKw : PdfObj { public string Raw; public PdfKw(string r) { Raw = r; } }
    public sealed class PdfRef : PdfObj
    {
        public int Num; public int Gen;
        // Sortida = ja es un numero del PDF NOU (no s'ha de tornar a traduir).
        public bool Sortida;
        public PdfRef(int n, int g, bool s) { Num = n; Gen = g; Sortida = s; }
    }
    public sealed class PdfArr : PdfObj { public List<PdfObj> Items = new List<PdfObj>(); }
    public sealed class PdfDict : PdfObj
    {
        public List<string> Keys = new List<string>();
        public Dictionary<string, PdfObj> Map = new Dictionary<string, PdfObj>();
        public PdfObj Get(string k) { PdfObj o; return Map.TryGetValue(k, out o) ? o : null; }
        public void Set(string k, PdfObj v) { if (!Map.ContainsKey(k)) Keys.Add(k); Map[k] = v; }
        public void Remove(string k) { if (Map.Remove(k)) Keys.Remove(k); }
        public PdfDict Clone()
        {
            PdfDict d = new PdfDict();
            foreach (string k in Keys) d.Set(k, Map[k]);
            return d;
        }
    }
    public sealed class PdfStream : PdfObj
    {
        public PdfDict Dict; public byte[] Data;
        public PdfStream(PdfDict d, byte[] data) { Dict = d; Data = data; }
    }

    // ---- Lectura -----------------------------------------------------------
    public sealed class PdfParser
    {
        readonly byte[] b; public int P; readonly int fi;
        public PdfParser(byte[] data, int pos) { b = data; P = pos; fi = data.Length; }
        public PdfParser(byte[] data, int pos, int end) { b = data; P = pos; fi = end; }

        public static bool EsBlanc(byte c) { return c == 0 || c == 9 || c == 10 || c == 12 || c == 13 || c == 32; }
        public static bool EsDelim(byte c)
        {
            return c == (byte)'(' || c == (byte)')' || c == (byte)'<' || c == (byte)'>' || c == (byte)'[' ||
                   c == (byte)']' || c == (byte)'{' || c == (byte)'}' || c == (byte)'/' || c == (byte)'%';
        }
        public bool Final { get { return P >= fi; } }

        public void SaltaBlancs()
        {
            while (P < fi)
            {
                byte c = b[P];
                if (EsBlanc(c)) { P++; continue; }
                if (c == (byte)'%') { while (P < fi && b[P] != 10 && b[P] != 13) P++; continue; }
                break;
            }
        }

        public bool Comenca(string s)
        {
            if (P + s.Length > fi) return false;
            for (int i = 0; i < s.Length; i++) if (b[P + i] != (byte)s[i]) return false;
            return true;
        }

        public string Paraula()
        {
            SaltaBlancs();
            int ini = P;
            while (P < fi && !EsBlanc(b[P]) && !EsDelim(b[P])) P++;
            return Encoding.ASCII.GetString(b, ini, P - ini);
        }

        static bool EsInt(string s)
        {
            if (s.Length == 0) return false;
            for (int i = 0; i < s.Length; i++) if (s[i] < '0' || s[i] > '9') return false;
            return true;
        }

        public PdfObj Objecte()
        {
            SaltaBlancs();
            if (P >= fi) throw new Exception("final inesperat");
            byte c = b[P];
            if (c == (byte)'/')
            {
                P++;
                int ini = P;
                while (P < fi && !EsBlanc(b[P]) && !EsDelim(b[P])) P++;
                return new PdfName(Encoding.ASCII.GetString(b, ini, P - ini));
            }
            if (c == (byte)'<' && P + 1 < fi && b[P + 1] == (byte)'<')
            {
                P += 2;
                PdfDict d = new PdfDict();
                while (true)
                {
                    SaltaBlancs();
                    if (P >= fi) throw new Exception("diccionari sense tancar");
                    if (b[P] == (byte)'>' && P + 1 < fi && b[P + 1] == (byte)'>') { P += 2; break; }
                    PdfObj k = Objecte();
                    PdfName kn = k as PdfName;
                    if (kn == null) continue;   // brossa: s'ignora
                    SaltaBlancs();
                    if (P < fi && b[P] == (byte)'>' && P + 1 < fi && b[P + 1] == (byte)'>') { d.Set(kn.Raw, new PdfKw("null")); continue; }
                    d.Set(kn.Raw, Objecte());
                }
                return d;
            }
            if (c == (byte)'<')
            {
                int ini = P;
                while (P < fi && b[P] != (byte)'>') P++;
                P++;
                return new PdfStr(Tros(ini, P));
            }
            if (c == (byte)'(')
            {
                int ini = P; int prof = 0;
                while (P < fi)
                {
                    byte x = b[P];
                    if (x == (byte)'\\') { P += 2; continue; }
                    if (x == (byte)'(') prof++;
                    else if (x == (byte)')') { prof--; if (prof == 0) { P++; break; } }
                    P++;
                }
                return new PdfStr(Tros(ini, P));
            }
            if (c == (byte)'[')
            {
                P++;
                PdfArr a = new PdfArr();
                while (true)
                {
                    SaltaBlancs();
                    if (P >= fi) throw new Exception("taula sense tancar");
                    if (b[P] == (byte)']') { P++; break; }
                    a.Items.Add(Objecte());
                }
                return a;
            }
            if ((c >= (byte)'0' && c <= (byte)'9') || c == (byte)'+' || c == (byte)'-' || c == (byte)'.')
            {
                int ini = P;
                P++;
                while (P < fi && ((b[P] >= (byte)'0' && b[P] <= (byte)'9') || b[P] == (byte)'.' || b[P] == (byte)'-' || b[P] == (byte)'+')) P++;
                string n1 = Encoding.ASCII.GetString(b, ini, P - ini);
                if (EsInt(n1))
                {
                    // "n g R": una referencia. Si no ho es, es torna enrere.
                    int desa = P;
                    string n2 = Paraula();
                    if (EsInt(n2))
                    {
                        string r = Paraula();
                        if (r == "R") return new PdfRef(int.Parse(n1, CultureInfo.InvariantCulture), int.Parse(n2, CultureInfo.InvariantCulture), false);
                    }
                    P = desa;
                }
                return new PdfNum(n1);
            }
            if (c == (byte)')' || c == (byte)'>' || c == (byte)']' || c == (byte)'{' || c == (byte)'}') { P++; return new PdfKw("null"); }
            string w = Paraula();
            if (w.Length == 0) { P++; return new PdfKw("null"); }
            return new PdfKw(w);
        }

        byte[] Tros(int ini, int fin)
        {
            if (fin > fi) fin = fi;
            byte[] r = new byte[Math.Max(0, fin - ini)];
            Array.Copy(b, ini, r, 0, r.Length);
            return r;
        }
    }

    public sealed class PdfDoc
    {
        struct Entrada { public int Tipus; public long Off; public int Stm; public int Idx; }
        sealed class ObjStm { public byte[] Data; public int[] Nums; public int[] Offs; public int First; }

        public readonly string Nom;
        public readonly byte[] B;
        public PdfDict Trailer = new PdfDict();
        readonly Dictionary<int, Entrada> xref = new Dictionary<int, Entrada>();
        readonly Dictionary<int, PdfObj> cache = new Dictionary<int, PdfObj>();
        readonly Dictionary<int, ObjStm> stms = new Dictionary<int, ObjStm>();
        int hdr;
        bool refet;

        public PdfDoc(string path)
        {
            Nom = Path.GetFileName(path);
            B = File.ReadAllBytes(path);
            hdr = IndexOf(B, "%PDF-", 0, Math.Min(B.Length, 1024));
            if (hdr < 0) throw new Exception("no es un PDF: " + Nom);
            try { LlegeixXrefs(); }
            catch { Refes(); }
            if (Trailer.Get("Root") == null) Refes();
            if (Trailer.Get("Encrypt") != null)
                throw new Exception("el PDF esta xifrat (protegit) i no es pot ajuntar: " + Nom);
        }

        public static int IndexOf(byte[] b, string s, int ini, int fi)
        {
            for (int i = ini; i <= fi - s.Length; i++)
            {
                int j = 0;
                while (j < s.Length && b[i + j] == (byte)s[j]) j++;
                if (j == s.Length) return i;
            }
            return -1;
        }
        static int LastIndexOf(byte[] b, string s)
        {
            for (int i = b.Length - s.Length; i >= 0; i--)
            {
                int j = 0;
                while (j < s.Length && b[i + j] == (byte)s[j]) j++;
                if (j == s.Length) return i;
            }
            return -1;
        }

        void LlegeixXrefs()
        {
            int sx = LastIndexOf(B, "startxref");
            if (sx < 0) throw new Exception("sense startxref");
            PdfParser pp = new PdfParser(B, sx + 9);
            long off = long.Parse(pp.Paraula(), CultureInfo.InvariantCulture);
            HashSet<long> vistos = new HashSet<long>();
            bool primer = true;
            while (off > 0 && !vistos.Contains(off))
            {
                vistos.Add(off);
                PdfDict tr = LlegeixXrefA(off);
                if (primer) { Trailer = tr; primer = false; }
                else
                {
                    foreach (string k in new string[] { "Root", "Info", "Encrypt", "ID" })
                        if (Trailer.Get(k) == null && tr.Get(k) != null) Trailer.Set(k, tr.Get(k));
                }
                PdfObj xs = tr.Get("XRefStm");
                if (xs is PdfNum)
                {
                    long o2 = (long)((PdfNum)xs).Val;
                    if (!vistos.Contains(o2)) { vistos.Add(o2); LlegeixXrefA(o2); }
                }
                PdfObj pv = tr.Get("Prev");
                off = (pv is PdfNum) ? (long)((PdfNum)pv).Val : 0;
            }
        }

        int Posicio(long off, string esperat)
        {
            // Hi ha PDF amb brossa abans de "%PDF-": els desplacaments poden
            // ser des del principi del fitxer o des de la capcalera.
            foreach (long o in new long[] { off, off + hdr })
            {
                if (o < 0 || o >= B.Length) continue;
                PdfParser p = new PdfParser(B, (int)o);
                p.SaltaBlancs();
                if (esperat == "xref" && p.Comenca("xref")) return p.P;
                if (esperat == "obj")
                {
                    string a = p.Paraula(); string g = p.Paraula(); string k = p.Paraula();
                    if (k == "obj" && a.Length > 0 && g.Length > 0) return (int)o;
                }
            }
            return -1;
        }

        PdfDict LlegeixXrefA(long off)
        {
            int p0 = Posicio(off, "xref");
            if (p0 >= 0)
            {
                PdfParser p = new PdfParser(B, p0 + 4);
                while (true)
                {
                    p.SaltaBlancs();
                    if (p.Comenca("trailer")) { p.P += 7; break; }
                    int ini = int.Parse(p.Paraula(), CultureInfo.InvariantCulture);
                    int n = int.Parse(p.Paraula(), CultureInfo.InvariantCulture);
                    for (int i = 0; i < n; i++)
                    {
                        long o = long.Parse(p.Paraula(), CultureInfo.InvariantCulture);
                        p.Paraula();
                        string t = p.Paraula();
                        int num = ini + i;
                        if (xref.ContainsKey(num)) continue;
                        Entrada e = new Entrada();
                        e.Tipus = (t == "n") ? 1 : 0; e.Off = o;
                        xref[num] = e;
                    }
                }
                PdfDict tr = p.Objecte() as PdfDict;
                if (tr == null) throw new Exception("trailer");
                return tr;
            }
            int po = Posicio(off, "obj");
            if (po < 0) throw new Exception("xref no trobada");
            PdfStream s = LlegeixIndirecte(po) as PdfStream;
            if (s == null) throw new Exception("xref no es un flux");
            byte[] d = Descodifica(s);
            int[] w = new int[3];
            PdfArr wa = s.Dict.Get("W") as PdfArr;
            for (int i = 0; i < 3; i++) w[i] = (int)((PdfNum)wa.Items[i]).Val;
            List<int> idx = new List<int>();
            PdfArr ia = s.Dict.Get("Index") as PdfArr;
            if (ia != null) foreach (PdfObj o in ia.Items) idx.Add((int)((PdfNum)o).Val);
            else { idx.Add(0); idx.Add((int)((PdfNum)s.Dict.Get("Size")).Val); }
            int pos = 0; int ample = w[0] + w[1] + w[2];
            for (int k = 0; k + 1 < idx.Count; k += 2)
            {
                for (int i = 0; i < idx[k + 1]; i++)
                {
                    if (pos + ample > d.Length) break;
                    long f0 = w[0] == 0 ? 1 : Camp(d, pos, w[0]);
                    long f1 = Camp(d, pos + w[0], w[1]);
                    long f2 = Camp(d, pos + w[0] + w[1], w[2]);
                    pos += ample;
                    int num = idx[k] + i;
                    if (xref.ContainsKey(num)) continue;
                    Entrada e = new Entrada();
                    e.Tipus = (int)f0;
                    if (f0 == 1) e.Off = f1;
                    else if (f0 == 2) { e.Stm = (int)f1; e.Idx = (int)f2; }
                    xref[num] = e;
                }
            }
            return s.Dict;
        }

        static long Camp(byte[] d, int pos, int n)
        {
            long v = 0;
            for (int i = 0; i < n; i++) v = (v << 8) | d[pos + i];
            return v;
        }

        // Reconstrueix la taula recorrent el fitxer ("n g obj"). Per a PDF amb
        // la taula malmesa: el darrer que surt mana, com en una revisio.
        void Refes()
        {
            if (refet) return;
            refet = true;
            xref.Clear(); cache.Clear(); stms.Clear();
            PdfDict trailerTrobat = null;
            for (int i = 0; i < B.Length - 4; i++)
            {
                if (B[i] == (byte)'t' && IndexOf(B, "trailer", i, Math.Min(B.Length, i + 7)) == i)
                {
                    try
                    {
                        PdfParser tp = new PdfParser(B, i + 7);
                        PdfDict td = tp.Objecte() as PdfDict;
                        if (td != null && td.Get("Root") != null) trailerTrobat = td;
                    }
                    catch { }
                    continue;
                }
                if (B[i] < (byte)'0' || B[i] > (byte)'9') continue;
                if (i > 0 && !PdfParser.EsBlanc(B[i - 1])) continue;
                PdfParser p = new PdfParser(B, i);
                string a = p.Paraula(); string g = p.Paraula(); string k = p.Paraula();
                int num; int gen;
                if (k == "obj" && int.TryParse(a, NumberStyles.None, CultureInfo.InvariantCulture, out num) &&
                    int.TryParse(g, NumberStyles.None, CultureInfo.InvariantCulture, out gen))
                {
                    Entrada e = new Entrada(); e.Tipus = 1; e.Off = i;
                    xref[num] = e;
                    i = p.P - 1;
                }
            }
            // Els objectes de dins dels fluxos d'objectes, si no hi son directes.
            List<int> nums = new List<int>(xref.Keys);
            foreach (int n in nums)
            {
                try
                {
                    PdfStream s = Get(n) as PdfStream;
                    if (s == null) continue;
                    PdfName t = s.Dict.Get("Type") as PdfName;
                    if (t != null && t.Raw == "XRef" && trailerTrobat == null) trailerTrobat = s.Dict;
                    if (t == null || t.Raw != "ObjStm") continue;
                    ObjStm os = CarregaStm(n);
                    for (int i = 0; i < os.Nums.Length; i++)
                    {
                        if (xref.ContainsKey(os.Nums[i])) continue;
                        Entrada e = new Entrada(); e.Tipus = 2; e.Stm = n; e.Idx = i;
                        xref[os.Nums[i]] = e;
                    }
                }
                catch { }
            }
            cache.Clear();
            Trailer = trailerTrobat ?? new PdfDict();
            if (Trailer.Get("Root") == null)
            {
                foreach (int n in new List<int>(xref.Keys))
                {
                    PdfDict d = null;
                    try { d = Get(n) as PdfDict; } catch { }
                    PdfName t = d == null ? null : d.Get("Type") as PdfName;
                    if (t != null && t.Raw == "Catalog") { Trailer.Set("Root", new PdfRef(n, 0, false)); break; }
                }
            }
            if (Trailer.Get("Root") == null) throw new Exception("no es troba el cataleg del PDF: " + Nom);
        }

        public PdfObj Resol(PdfObj o)
        {
            int guarda = 0;
            while (o is PdfRef && !((PdfRef)o).Sortida && guarda++ < 32) o = Get(((PdfRef)o).Num);
            return o;
        }

        public PdfObj Get(int num)
        {
            PdfObj o;
            if (cache.TryGetValue(num, out o)) return o;
            cache[num] = new PdfKw("null");   // contra cicles (Length que apunta a si mateix...)
            o = new PdfKw("null");
            Entrada e;
            try
            {
                if (xref.TryGetValue(num, out e))
                {
                    if (e.Tipus == 1)
                    {
                        int po = Posicio(e.Off, "obj");
                        if (po >= 0) o = LlegeixIndirecte(po);
                        else if (!refet) { Refes(); return Get(num); }
                    }
                    else if (e.Tipus == 2)
                    {
                        ObjStm os = CarregaStm(e.Stm);
                        if (e.Idx < os.Offs.Length)
                            o = new PdfParser(os.Data, os.First + os.Offs[e.Idx]).Objecte();
                    }
                }
            }
            catch
            {
                if (!refet) { Refes(); return Get(num); }
            }
            cache[num] = o;
            return o;
        }

        ObjStm CarregaStm(int num)
        {
            ObjStm os;
            if (stms.TryGetValue(num, out os)) return os;
            PdfStream s = Get(num) as PdfStream;
            if (s == null) throw new Exception("flux d'objectes " + num);
            os = new ObjStm();
            os.Data = Descodifica(s);
            int n = (int)((PdfNum)Resol(s.Dict.Get("N"))).Val;
            os.First = (int)((PdfNum)Resol(s.Dict.Get("First"))).Val;
            os.Nums = new int[n]; os.Offs = new int[n];
            PdfParser p = new PdfParser(os.Data, 0);
            for (int i = 0; i < n; i++)
            {
                os.Nums[i] = int.Parse(p.Paraula(), CultureInfo.InvariantCulture);
                os.Offs[i] = int.Parse(p.Paraula(), CultureInfo.InvariantCulture);
            }
            stms[num] = os;
            return os;
        }

        PdfObj LlegeixIndirecte(int pos)
        {
            PdfParser p = new PdfParser(B, pos);
            p.Paraula(); p.Paraula(); p.Paraula();   // n g obj
            PdfObj o = p.Objecte();
            PdfDict d = o as PdfDict;
            if (d == null) return o;
            p.SaltaBlancs();
            if (!p.Comenca("stream")) return o;
            p.P += 6;
            if (p.P < B.Length && B[p.P] == 13) p.P++;
            if (p.P < B.Length && B[p.P] == 10) p.P++;
            int ini = p.P;
            int len = -1;
            PdfObj lo = d.Get("Length");
            if (lo is PdfNum) len = (int)((PdfNum)lo).Val;
            else if (lo is PdfRef) { PdfNum ln = Get(((PdfRef)lo).Num) as PdfNum; if (ln != null) len = (int)ln.Val; }
            bool bo = false;
            if (len >= 0 && ini + len <= B.Length)
            {
                PdfParser q = new PdfParser(B, ini + len);
                q.SaltaBlancs();
                bo = q.Comenca("endstream");
            }
            if (!bo)
            {
                int fi = IndexOf(B, "endstream", ini, B.Length);
                if (fi < 0) fi = B.Length;
                int f = fi;
                if (f > ini && B[f - 1] == 10) f--;
                if (f > ini && B[f - 1] == 13) f--;
                len = f - ini;
            }
            byte[] data = new byte[len];
            Array.Copy(B, ini, data, 0, len);
            return new PdfStream(d, data);
        }

        public byte[] Descodifica(PdfStream s)
        {
            PdfObj f = Resol(s.Dict.Get("Filter"));
            PdfObj dp = Resol(s.Dict.Get("DecodeParms"));
            List<string> filtres = new List<string>();
            List<PdfDict> parms = new List<PdfDict>();
            if (f is PdfName) { filtres.Add(((PdfName)f).Raw); parms.Add(dp as PdfDict); }
            else if (f is PdfArr)
            {
                PdfArr fa = (PdfArr)f; PdfArr pa = dp as PdfArr;
                for (int i = 0; i < fa.Items.Count; i++)
                {
                    PdfName nm = Resol(fa.Items[i]) as PdfName;
                    filtres.Add(nm == null ? "" : nm.Raw);
                    parms.Add(pa != null && i < pa.Items.Count ? Resol(pa.Items[i]) as PdfDict : null);
                }
            }
            byte[] d = s.Data;
            for (int i = 0; i < filtres.Count; i++)
            {
                if (filtres[i] == "FlateDecode" || filtres[i] == "Fl") d = Inflate(d);
                else throw new Exception("filtre no suportat: " + filtres[i]);
                d = Predictor(d, parms[i]);
            }
            return d;
        }

        static byte[] Inflate(byte[] d)
        {
            int ini = (d.Length > 2 && (d[0] & 0x0F) == 8) ? 2 : 0;
            MemoryStream sortida = new MemoryStream();
            try
            {
                using (DeflateStream z = new DeflateStream(new MemoryStream(d, ini, d.Length - ini), CompressionMode.Decompress))
                {
                    byte[] buf = new byte[8192]; int n;
                    while ((n = z.Read(buf, 0, buf.Length)) > 0) sortida.Write(buf, 0, n);
                }
            }
            catch { if (sortida.Length == 0) throw; }   // cua malmesa: val el que s'ha llegit
            return sortida.ToArray();
        }

        PdfObj Num(PdfDict d, string k) { return d == null ? null : Resol(d.Get(k)); }

        byte[] Predictor(byte[] d, PdfDict p)
        {
            PdfNum pr = Num(p, "Predictor") as PdfNum;
            if (pr == null || pr.Val < 10) return d;
            PdfNum co = Num(p, "Columns") as PdfNum; PdfNum cl = Num(p, "Colors") as PdfNum; PdfNum bc = Num(p, "BitsPerComponent") as PdfNum;
            int cols = co == null ? 1 : (int)co.Val; int colors = cl == null ? 1 : (int)cl.Val; int bpc = bc == null ? 8 : (int)bc.Val;
            int bpp = Math.Max(1, colors * bpc / 8);
            int fila = (cols * colors * bpc + 7) / 8;
            MemoryStream o = new MemoryStream();
            byte[] ant = new byte[fila];
            int pos = 0;
            while (pos + 1 + fila <= d.Length)
            {
                int t = d[pos]; byte[] act = new byte[fila];
                Array.Copy(d, pos + 1, act, 0, fila);
                for (int i = 0; i < fila; i++)
                {
                    int a = i >= bpp ? act[i - bpp] : 0; int up = ant[i]; int c = i >= bpp ? ant[i - bpp] : 0;
                    int v = act[i];
                    switch (t)
                    {
                        case 1: v += a; break;
                        case 2: v += up; break;
                        case 3: v += (a + up) / 2; break;
                        case 4:
                            int pa = Math.Abs(up - c), pb = Math.Abs(a - c), pc = Math.Abs(a + up - 2 * c);
                            v += (pa <= pb && pa <= pc) ? a : (pb <= pc ? up : c); break;
                    }
                    act[i] = (byte)v;
                }
                o.Write(act, 0, fila);
                ant = act; pos += 1 + fila;
            }
            return o.ToArray();
        }

        // Les pagines en ordre, amb els atributs HERETATS ja resolts.
        public sealed class Pagina { public int Num; public PdfDict Dict; public PdfDict Heretat = new PdfDict(); }
        public List<Pagina> Pagines(HashSet<int> nodesArbre)
        {
            List<Pagina> r = new List<Pagina>();
            PdfDict cat = Resol(Trailer.Get("Root")) as PdfDict;
            if (cat == null) throw new Exception("sense cataleg: " + Nom);
            Recorre(cat.Get("Pages"), new PdfDict(), r, nodesArbre, 0);
            return r;
        }
        void Recorre(PdfObj node, PdfDict her, List<Pagina> r, HashSet<int> nodes, int prof)
        {
            if (prof > 64) return;
            int num = node is PdfRef ? ((PdfRef)node).Num : -1;
            if (num >= 0 && nodes.Contains(num)) return;   // cicle
            PdfDict d = Resol(node) as PdfDict;
            if (d == null) return;
            PdfName t = d.Get("Type") as PdfName;
            PdfObj kids = Resol(d.Get("Kids"));
            bool esNode = (t != null && t.Raw == "Pages") || (t == null && kids is PdfArr);
            PdfDict h = her.Clone();
            foreach (string k in new string[] { "Resources", "MediaBox", "CropBox", "Rotate" })
                if (d.Get(k) != null) h.Set(k, d.Get(k));
            if (esNode)
            {
                if (num >= 0) nodes.Add(num);
                PdfArr ka = kids as PdfArr;
                if (ka != null) foreach (PdfObj k in ka.Items) Recorre(k, h, r, nodes, prof + 1);
                return;
            }
            Pagina pg = new Pagina(); pg.Num = num; pg.Dict = d; pg.Heretat = h;
            if (num >= 0) nodes.Add(num);
            r.Add(pg);
        }
    }

    // ---- La unio -----------------------------------------------------------
    public sealed class PdfUnio
    {
        readonly Dictionary<int, PdfObj> sortida = new Dictionary<int, PdfObj>();
        int seguent = 1;
        sealed class Pendent { public PdfDoc Doc; public int Src; public int Nou; }
        readonly Queue<Pendent> cua = new Queue<Pendent>();

        public int Pagines;
        public int SignaturesAplanades;
        public int SignaturesInvisibles;

        int Reserva() { return seguent++; }
        static PdfRef RefSortida(int n) { return new PdfRef(n, 0, true); }

        // Uneix $entrades (el primer es l'informe) a $sortida. Retorna un resum.
        public static string Unir(string[] entrades, string sortida)
        {
            PdfUnio u = new PdfUnio();
            u.Fes(entrades, sortida);
            return string.Format(CultureInfo.InvariantCulture, "pagines={0};aplanades={1};invisibles={2}",
                u.Pagines, u.SignaturesAplanades, u.SignaturesInvisibles);
        }

        void Fes(string[] entrades, string fitxerSortida)
        {
            int numPages = Reserva();
            int numCat = Reserva();
            PdfArr kids = new PdfArr();
            PdfRef info = null;
            bool primer = true;
            foreach (string f in entrades)
            {
                PdfDoc doc = new PdfDoc(f);
                Dictionary<int, int> mapa = new Dictionary<int, int>();
                HashSet<int> nodes = new HashSet<int>();
                List<PdfDoc.Pagina> pgs = doc.Pagines(nodes);
                if (pgs.Count == 0) throw new Exception("el PDF no te pagines: " + doc.Nom);
                // Els nodes de l'arbre de pagines de l'original apunten al NOU.
                foreach (int n in nodes) mapa[n] = numPages;
                foreach (PdfDoc.Pagina pg in pgs) if (pg.Num >= 0) mapa[pg.Num] = Reserva();
                foreach (PdfDoc.Pagina pg in pgs)
                {
                    int nou = pg.Num >= 0 ? mapa[pg.Num] : Reserva();
                    sortida[nou] = Pagina(doc, mapa, pg, numPages);
                    kids.Items.Add(RefSortida(nou));
                    Pagines++;
                }
                if (primer && doc.Trailer.Get("Info") != null)
                {
                    PdfObj i = Tradueix(doc, mapa, doc.Trailer.Get("Info"));
                    info = i as PdfRef;
                }
                Buida(doc, mapa);
                primer = false;
            }
            PdfDict pages = new PdfDict();
            pages.Set("Type", new PdfName("Pages"));
            pages.Set("Kids", kids);
            pages.Set("Count", new PdfNum(kids.Items.Count.ToString(CultureInfo.InvariantCulture)));
            sortida[numPages] = pages;
            PdfDict cat = new PdfDict();
            cat.Set("Type", new PdfName("Catalog"));
            cat.Set("Pages", RefSortida(numPages));
            sortida[numCat] = cat;
            Escriu(fitxerSortida, numCat, info);
        }

        void Buida(PdfDoc doc, Dictionary<int, int> mapa)
        {
            while (cua.Count > 0)
            {
                Pendent p = cua.Dequeue();
                PdfObj o = doc.Get(p.Src);
                PdfStream s = o as PdfStream;
                if (s != null) sortida[p.Nou] = new PdfStream((PdfDict)Tradueix(doc, mapa, s.Dict), s.Data);
                else sortida[p.Nou] = Tradueix(doc, mapa, o);
            }
        }

        // Copia un valor de l'original al PDF nou, renumerant les referencies.
        PdfObj Tradueix(PdfDoc doc, Dictionary<int, int> mapa, PdfObj o)
        {
            PdfRef r = o as PdfRef;
            if (r != null)
            {
                if (r.Sortida) return r;
                int n;
                if (!mapa.TryGetValue(r.Num, out n))
                {
                    n = Reserva();
                    mapa[r.Num] = n;
                    Pendent p = new Pendent(); p.Doc = doc; p.Src = r.Num; p.Nou = n;
                    cua.Enqueue(p);
                }
                return RefSortida(n);
            }
            PdfDict d = o as PdfDict;
            if (d != null)
            {
                PdfDict nd = new PdfDict();
                foreach (string k in d.Keys) nd.Set(k, Tradueix(doc, mapa, d.Map[k]));
                return nd;
            }
            PdfArr a = o as PdfArr;
            if (a != null)
            {
                PdfArr na = new PdfArr();
                foreach (PdfObj x in a.Items) na.Items.Add(Tradueix(doc, mapa, x));
                return na;
            }
            PdfStream s = o as PdfStream;
            if (s != null)
            {
                // Un flux ha de ser un objecte indirecte.
                int n = Reserva();
                sortida[n] = new PdfStream((PdfDict)Tradueix(doc, mapa, s.Dict), s.Data);
                return RefSortida(n);
            }
            return o;
        }

        static bool EsSignatura(PdfDoc doc, PdfDict a)
        {
            PdfName st = doc.Resol(a.Get("Subtype")) as PdfName;
            if (st == null || st.Raw != "Widget") return false;
            PdfDict x = a;
            for (int i = 0; i < 32 && x != null; i++)
            {
                PdfName ft = doc.Resol(x.Get("FT")) as PdfName;
                if (ft != null) return ft.Raw == "Sig";
                x = doc.Resol(x.Get("Parent")) as PdfDict;
            }
            return false;
        }

        static double[] Nums(PdfDoc doc, PdfObj o, int n)
        {
            PdfArr a = doc.Resol(o) as PdfArr;
            if (a == null || a.Items.Count < n) return null;
            double[] r = new double[n];
            for (int i = 0; i < n; i++)
            {
                PdfNum x = doc.Resol(a.Items[i]) as PdfNum;
                if (x == null) return null;
                r[i] = x.Val;
            }
            return r;
        }

        static string F(double v)
        {
            if (Math.Abs(v) < 0.000001) v = 0;
            return v.ToString("0.######", CultureInfo.InvariantCulture);
        }

        // Una pagina nova. LES SIGNATURES S'APLANEN: l'aparenca de cada camp de
        // signatura passa al CONTINGUT de la pagina (es segueix veient igual) i
        // el camp desapareix. Deixar-lo voldria dir una signatura que ja no
        // quadra amb el document nou: "signatura no valida".
        PdfDict Pagina(PdfDoc doc, Dictionary<int, int> mapa, PdfDoc.Pagina pg, int numPages)
        {
            PdfDict np = new PdfDict();
            foreach (string k in pg.Dict.Keys)
            {
                if (k == "Parent" || k == "Annots" || k == "B" || k == "Resources" || k == "Contents") continue;
                np.Set(k, pg.Dict.Map[k]);
            }
            foreach (string k in new string[] { "MediaBox", "CropBox", "Rotate" })
                if (np.Get(k) == null && pg.Heretat.Get(k) != null) np.Set(k, pg.Heretat.Get(k));
            if (np.Get("MediaBox") == null)
            {
                PdfArr mb = new PdfArr();
                foreach (string v in new string[] { "0", "0", "595.276", "841.89" }) mb.Items.Add(new PdfNum(v));
                np.Set("MediaBox", mb);
            }

            PdfArr annotsNous = new PdfArr();
            StringBuilder ops = new StringBuilder();
            PdfDict xobjNous = new PdfDict();
            PdfDict recursos = doc.Resol(pg.Heretat.Get("Resources")) as PdfDict;
            PdfDict xobjVells = recursos == null ? null : doc.Resol(recursos.Get("XObject")) as PdfDict;
            int k2 = 0;
            PdfArr annots = doc.Resol(pg.Dict.Get("Annots")) as PdfArr;
            if (annots != null)
            {
                foreach (PdfObj ao in annots.Items)
                {
                    PdfDict a = doc.Resol(ao) as PdfDict;
                    if (a == null) continue;
                    if (!EsSignatura(doc, a)) { annotsNous.Items.Add(ao); continue; }
                    // Invisible (amagada, sense mida o sense aparenca): fora i prou.
                    PdfNum fl = doc.Resol(a.Get("F")) as PdfNum;
                    int flags = fl == null ? 0 : (int)fl.Val;
                    double[] rect = Nums(doc, a.Get("Rect"), 4);
                    PdfDict ap = doc.Resol(a.Get("AP")) as PdfDict;
                    PdfObj n = ap == null ? null : ap.Get("N");
                    PdfObj nr = doc.Resol(n);
                    if (nr is PdfDict && !(nr is PdfStream))
                    {
                        PdfName ast = doc.Resol(a.Get("AS")) as PdfName;
                        n = ast == null ? null : ((PdfDict)nr).Get(ast.Raw);
                        nr = doc.Resol(n);
                    }
                    PdfStream form = nr as PdfStream;
                    if ((flags & 2) != 0 || (flags & 32) != 0 || rect == null || form == null) { SignaturesInvisibles++; continue; }
                    double rx1 = Math.Min(rect[0], rect[2]), rx2 = Math.Max(rect[0], rect[2]);
                    double ry1 = Math.Min(rect[1], rect[3]), ry2 = Math.Max(rect[1], rect[3]);
                    double[] bb = Nums(doc, form.Dict.Get("BBox"), 4);
                    double[] m = Nums(doc, form.Dict.Get("Matrix"), 6) ?? new double[] { 1, 0, 0, 1, 0, 0 };
                    if (bb == null || rx2 - rx1 <= 0 || ry2 - ry1 <= 0) { SignaturesInvisibles++; continue; }
                    // El BBox transformat per la Matrix (algorisme de l'apartat
                    // 12.5.5 de l'ISO 32000): el rectangle que ocupa l'aparenca.
                    double bx1 = double.MaxValue, by1 = double.MaxValue, bx2 = double.MinValue, by2 = double.MinValue;
                    foreach (double[] c in new double[][] { new double[] { bb[0], bb[1] }, new double[] { bb[2], bb[1] }, new double[] { bb[0], bb[3] }, new double[] { bb[2], bb[3] } })
                    {
                        double x = m[0] * c[0] + m[2] * c[1] + m[4];
                        double y = m[1] * c[0] + m[3] * c[1] + m[5];
                        bx1 = Math.Min(bx1, x); bx2 = Math.Max(bx2, x); by1 = Math.Min(by1, y); by2 = Math.Max(by2, y);
                    }
                    if (bx2 - bx1 <= 0 || by2 - by1 <= 0) { SignaturesInvisibles++; continue; }
                    double sx = (rx2 - rx1) / (bx2 - bx1), sy = (ry2 - ry1) / (by2 - by1);
                    double tx = rx1 - bx1 * sx, ty = ry1 - by1 * sy;
                    string nom;
                    do { nom = "SigAplanada" + (++k2).ToString(CultureInfo.InvariantCulture); }
                    while (xobjVells != null && xobjVells.Get(nom) != null);
                    PdfObj valor = n;
                    PdfName sub = doc.Resol(form.Dict.Get("Subtype")) as PdfName;
                    if (sub == null || sub.Raw != "Form" || !(n is PdfRef))
                    {
                        PdfDict fd = form.Dict.Clone();
                        fd.Set("Type", new PdfName("XObject"));
                        fd.Set("Subtype", new PdfName("Form"));
                        valor = new PdfStream(fd, form.Data);
                    }
                    xobjNous.Set(nom, valor);
                    ops.Append("q ").Append(F(sx)).Append(" 0 0 ").Append(F(sy)).Append(' ').Append(F(tx)).Append(' ').Append(F(ty))
                       .Append(" cm /").Append(nom).Append(" Do Q\n");
                    SignaturesAplanades++;
                }
            }
            if (annotsNous.Items.Count > 0) np.Set("Annots", annotsNous);

            // Recursos: una COPIA (poden ser compartits amb altres pagines) amb
            // les aparences afegides.
            if (xobjNous.Keys.Count > 0)
            {
                PdfDict nr = recursos == null ? new PdfDict() : recursos.Clone();
                PdfDict nx = xobjVells == null ? new PdfDict() : xobjVells.Clone();
                foreach (string k in xobjNous.Keys) nx.Set(k, xobjNous.Map[k]);
                nr.Set("XObject", nx);
                np.Set("Resources", nr);
            }
            else if (pg.Heretat.Get("Resources") != null) np.Set("Resources", pg.Heretat.Get("Resources"));
            else np.Set("Resources", new PdfDict());

            // Contingut: el de sempre entre q/Q (que no deixi res canviat) i les
            // aparences al damunt.
            PdfObj cont = pg.Dict.Get("Contents");
            if (ops.Length > 0)
            {
                PdfArr ca = new PdfArr();
                ca.Items.Add(new PdfStream(new PdfDict(), Encoding.ASCII.GetBytes("q\n")));
                PdfObj cr = doc.Resol(cont);
                if (cr is PdfArr) foreach (PdfObj x in ((PdfArr)cr).Items) ca.Items.Add(x);
                else if (cont != null) ca.Items.Add(cont);
                ca.Items.Add(new PdfStream(new PdfDict(), Encoding.ASCII.GetBytes("\nQ\n" + ops.ToString())));
                np.Set("Contents", ca);
            }
            else if (cont != null) np.Set("Contents", cont);

            PdfDict res = (PdfDict)Tradueix(doc, mapa, np);
            res.Set("Parent", RefSortida(numPages));
            return res;
        }

        // ---- Escriptura: un PDF net, sense revisions ni fluxos d'objectes --
        void Escriu(string fitxer, int numCat, PdfRef info)
        {
            MemoryStream ms = new MemoryStream();
            Posa(ms, "%PDF-1.7\n%");
            ms.Write(new byte[] { 0xE2, 0xE3, 0xCF, 0xD3, 10 }, 0, 5);
            long[] offs = new long[seguent];
            for (int n = 1; n < seguent; n++)
            {
                offs[n] = ms.Position;
                PdfObj o;
                if (!sortida.TryGetValue(n, out o)) o = new PdfKw("null");
                Posa(ms, n.ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
                Serialitza(ms, o);
                Posa(ms, "\nendobj\n");
            }
            long xr = ms.Position;
            StringBuilder sb = new StringBuilder();
            sb.Append("xref\n0 ").Append(seguent.ToString(CultureInfo.InvariantCulture)).Append("\n0000000000 65535 f\r\n");
            for (int n = 1; n < seguent; n++) sb.Append(offs[n].ToString("0000000000", CultureInfo.InvariantCulture)).Append(" 00000 n\r\n");
            sb.Append("trailer\n<< /Size ").Append(seguent.ToString(CultureInfo.InvariantCulture))
              .Append(" /Root ").Append(numCat.ToString(CultureInfo.InvariantCulture)).Append(" 0 R");
            if (info != null) sb.Append(" /Info ").Append(info.Num.ToString(CultureInfo.InvariantCulture)).Append(" 0 R");
            sb.Append(" >>\nstartxref\n").Append(xr.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
            Posa(ms, sb.ToString());
            File.WriteAllBytes(fitxer, ms.ToArray());
        }

        static void Posa(Stream s, string t) { byte[] b = Encoding.ASCII.GetBytes(t); s.Write(b, 0, b.Length); }

        static void Serialitza(Stream s, PdfObj o)
        {
            if (o is PdfName) { Posa(s, "/" + ((PdfName)o).Raw); return; }
            if (o is PdfNum) { Posa(s, ((PdfNum)o).Raw); return; }
            if (o is PdfKw) { Posa(s, ((PdfKw)o).Raw); return; }
            if (o is PdfStr) { byte[] r = ((PdfStr)o).Raw; s.Write(r, 0, r.Length); return; }
            if (o is PdfRef) { Posa(s, ((PdfRef)o).Num.ToString(CultureInfo.InvariantCulture) + " 0 R"); return; }
            if (o is PdfArr)
            {
                Posa(s, "[");
                bool pr = true;
                foreach (PdfObj x in ((PdfArr)o).Items) { if (!pr) Posa(s, " "); Serialitza(s, x); pr = false; }
                Posa(s, "]");
                return;
            }
            if (o is PdfDict)
            {
                PdfDict d = (PdfDict)o;
                Posa(s, "<<");
                foreach (string k in d.Keys) { Posa(s, "/" + k + " "); Serialitza(s, d.Map[k]); Posa(s, "\n"); }
                Posa(s, ">>");
                return;
            }
            if (o is PdfStream)
            {
                PdfStream st = (PdfStream)o;
                PdfDict d = st.Dict.Clone();
                d.Remove("Length");
                d.Set("Length", new PdfNum(st.Data.Length.ToString(CultureInfo.InvariantCulture)));
                Serialitza(s, d);
                Posa(s, "\nstream\r\n");
                s.Write(st.Data, 0, st.Data.Length);
                Posa(s, "\r\nendstream");
                return;
            }
            Posa(s, "null");
        }

        // ---- Per comprovar (proves i resum) ---------------------------------
        public static string Descriu(string fitxer)
        {
            PdfDoc doc = new PdfDoc(fitxer);
            HashSet<int> nodes = new HashSet<int>();
            List<PdfDoc.Pagina> pgs = doc.Pagines(nodes);
            int sig = 0;
            foreach (PdfDoc.Pagina pg in pgs)
            {
                PdfArr annots = doc.Resol(pg.Dict.Get("Annots")) as PdfArr;
                if (annots == null) continue;
                foreach (PdfObj ao in annots.Items)
                {
                    PdfDict a = doc.Resol(ao) as PdfDict;
                    if (a != null && EsSignatura(doc, a)) sig++;
                }
            }
            PdfDict cat = doc.Resol(doc.Trailer.Get("Root")) as PdfDict;
            bool acro = cat != null && cat.Get("AcroForm") != null;
            return string.Format(CultureInfo.InvariantCulture, "pagines={0};signatures={1};acroform={2}", pgs.Count, sig, acro ? 1 : 0);
        }
    }
}
'@

# Compila el C# si encara no ho esta. Un sol cop per sessio.
function _PdfUnioCarrega {
    if ('InformesCornella.PdfUnio' -as [type]) { return }
    Add-Type -TypeDefinition $Script:PdfUnioCs -ErrorAction Stop
}

# Ajunta $informe + $adjunts (en aquest ordre) a $sortida. Retorna
# @{ Pagines; Aplanades; Invisibles }. Llanca si algun PDF no es pot llegir
# (xifrat, malmes sense remei...): qui crida decideix que fa.
function Join-PdfAmbAdjunts([string]$informe, $adjunts, [string]$sortida) {
    _PdfUnioCarrega
    $entrades = [string[]]@(@($informe) + @($adjunts))
    return (_PdfUnioResum ([InformesCornella.PdfUnio]::Unir($entrades, $sortida)))
}

# Pagines, camps de signatura i si te formulari (AcroForm). Per al resum i les
# proves: @{ Pagines; Signatures; AcroForm }.
function Get-PdfDescripcio([string]$pdf) {
    _PdfUnioCarrega
    return (_PdfUnioResum ([InformesCornella.PdfUnio]::Descriu($pdf)))
}

# "clau=valor;clau=valor" -> hashtable amb enters. PURA.
function _PdfUnioResum([string]$text) {
    $noms = @{ pagines = 'Pagines'; aplanades = 'Aplanades'; invisibles = 'Invisibles'; signatures = 'Signatures'; acroform = 'AcroForm' }
    $r = @{}
    foreach ($t in ([string]$text -split ';')) {
        $kv = $t -split '=', 2
        if ($kv.Count -ne 2) { continue }
        $k = if ($noms.ContainsKey($kv[0])) { $noms[$kv[0]] } else { $kv[0] }
        $r[$k] = [int]$kv[1]
    }
    return $r
}
