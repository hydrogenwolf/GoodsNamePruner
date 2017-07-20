using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace GoodsNamePruner.Controllers
{
    public class Entry
    {
        public object Key;
        public object Value;

        public Entry()
        {
        }

        public Entry(object key, object value)
        {
            Key = key;
            Value = value;
        }
    }

    public class HomeController : Controller
    {
        private Dictionary<string, string> BasicRules = new Dictionary<string, string>();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public static void Serialize(TextWriter writer, IDictionary dictionary)
        {
            List<Entry> entries = new List<Entry>(dictionary.Count);
            foreach (object key in dictionary.Keys)
            {
                entries.Add(new Entry(key, dictionary[key]));
            }
            XmlSerializer serializer = new XmlSerializer(typeof(List<Entry>));
            serializer.Serialize(writer, entries);
        }

        public static void Deserialize(TextReader reader, IDictionary dictionary)
        {
            dictionary.Clear();
            XmlSerializer serializer = new XmlSerializer(typeof(List<Entry>));
            List<Entry> list = (List<Entry>)serializer.Deserialize(reader);
            foreach (Entry entry in list)
            {
                dictionary[entry.Key] = entry.Value;
            }
        }

        protected void ReadBasicRules(string filename)
        {
            Excel.Application excelApplication = new Excel.Application();
            //excelApplication.Visible = true;
            var workbooks = excelApplication.Workbooks;
            var workbook = workbooks.Open(filename);

            int count = 0;
            for (int row = 1; ; row++)
            {
                string a = workbook.ActiveSheet.Range("A" + row).Value;
                if (String.IsNullOrWhiteSpace(a)) break;
                a = Crop(a);

                string b = workbook.ActiveSheet.Range("B" + row).Value;
                b = Crop(b);

                if (!BasicRules.ContainsKey(a))
                {
                    BasicRules.Add(a, b);
                    count++;
                }
            }

            workbook.Close();
            workbooks.Close();
            excelApplication.Quit();
        }

        public string Checklist()
        {
            ReadBasicRules(Server.MapPath(Path.Combine("~/App_Data/Checklist", "Checklist.xlsx")));

            using (StreamWriter writer = System.IO.File.CreateText(Server.MapPath(Path.Combine("~/App_Data/Checklist", "Checklist.txt"))))
            {
                Serialize(writer, BasicRules);
            }

            return "Checklist file created...";
        }

        [HttpPost]
        public ActionResult Convert(HttpPostedFileBase file)
        {
            try
            {
                using (TextReader reader = System.IO.File.OpenText(Server.MapPath(Path.Combine("~/App_Data/Checklist", "Checklist.xml"))))
                {
                    Deserialize(reader, BasicRules);
                }

                if (file.ContentLength > 0)
                {
                    var fileName = Path.GetFileName(file.FileName);

                    string uploadFile = Server.MapPath(Path.Combine("~/App_Data/Upload", fileName));
                    string resultFile = Server.MapPath(Path.Combine("~/App_Data/Result", fileName));

                    if (System.IO.File.Exists(uploadFile)) System.IO.File.Delete(uploadFile);
                    if (System.IO.File.Exists(resultFile)) System.IO.File.Delete(resultFile);

                    file.SaveAs(uploadFile);

                    Excel.Application excelApplication = new Excel.Application();
                    excelApplication.DisplayAlerts = false;
                    //excelApplication.Visible = true;
                    var workbooks = excelApplication.Workbooks;
                    var workbook = workbooks.Open(uploadFile);

                    char goodsNameColumn = 'H';
                    for (char c = 'A'; c <= 'Z'; c++)
                    {
                        string column = c + "1";

                        if (workbook.ActiveSheet.Range(column).Value == "품목"
                            || workbook.ActiveSheet.Range(column).Value == "품목명"
                            || workbook.ActiveSheet.Range(column).Value == "상품"
                            || workbook.ActiveSheet.Range(column).Value == "상품명"
                            )
                        {
                            goodsNameColumn = c;
                            break;
                        }
                    }

                    for (int row = 1; ; row++)
                    {
                        string column = goodsNameColumn + row.ToString();

                        string s = workbook.ActiveSheet.Range(column).Value;
                        if (String.IsNullOrWhiteSpace(s)) break;
                        //s = Crop(s);

                        s = s.Trim();
                        string count = String.Empty;
                        string pattern = @"^(.+)(\[\d+\])$";
                        Regex rgx = new Regex(pattern);
                        Match match = Regex.Match(s, pattern);
                        if (match.Success)
                        {
                            s = match.Groups[1].Value;
                            count = match.Groups[2].Value;
                        }
                        s = s.Trim();

                        string title = "";
                        if (BasicRules.TryGetValue(s, out title))
                        {
                            workbook.ActiveSheet.Range(column).Value = title + " " + count;
                        }
                        else
                        {
                            workbook.ActiveSheet.Range(column).Value = Beautify(RemoveDuplicates(RemoveOptionNumbers(RemoveWords(s)))) + " " + count;
                        }
                    }

                    workbook.SaveAs(resultFile);

                    workbook.Close();
                    workbooks.Close();
                    excelApplication.Quit();

                    return new FilePathResult(resultFile, "application/vnd.ms-excel");
                }

                return RedirectToAction("GoodsNameSimplifier");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);

                return RedirectToAction("GoodsNameSimplifier");
            }
        }

        protected string Crop(string s)
        {
            s = s.Trim();

            // 기능 협의 후 삭제
            string pattern = @"★\s*수량\s*\d+$";
            Regex rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\[\d+\]$";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            return s;
        }

        protected string RemoveWords(string s)
        {
            s = s.Trim();

            string pattern = @"\d+원";
            Regex rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            s = s.Replace("2피스 반팔+반바지 세트", "하복").Replace("2종세트 (반팔티셔츠+반바지)", "하복");
            s = s.Replace("|네이비", "네이비").Replace("|레드", "레드");
            s = s.Replace("23A01-2314", "");
            s = s.Replace("자전거모자/자전거의류", "");
            s = s.Replace("무료배송", "").Replace("(단일상품)", "").Replace("사입", "").Replace("free", "").Replace("FREE", "")
                    .Replace("특가 모음", "").Replace("1개", "").Replace("1+1", "").Replace("균일가", "").Replace("사은품증정", "")
                    .Replace("남여", "").Replace("|여성", "").Replace("티셔츠 구매시 한개 추가증정", "");
            s = s.Replace("블랙 / 화이트 택1", "").Replace("블랙/화이트 택1", "").Replace("화이트,블랙", "");
            s = s.Replace("초극세사", "").Replace("비치타월", "").Replace("null", "").Replace("빕숏", "");
            s = s.Replace("기능성", "").Replace("언더레이어", "").Replace("미세먼지 차단", "")
                .Replace("(패드없음)", "").Replace("패드없음", "").Replace("(패드있음)", "").Replace("패드있음", "")
                .Replace("(속바지포함)", "").Replace("(속바지)", "").Replace("테크핏", "");
            s = s.Replace("자전거세트", "").Replace("자전거의류", "").Replace("자전거상의", "").Replace("자전거바지", "").Replace("자전거티셔츠", "")
                    .Replace("자전거팬츠", "").Replace("자전거장갑", "").Replace("자전거마스크", "").Replace("자전거악세사리", "")
                    .Replace("자전거반바지", "").Replace("자전거가방", "");
            s = s.Replace("트레이닝복", "").Replace("등산복", "");
            s = s.Replace("바람막이자켓", "").Replace("윈드자켓", "").Replace("경량바람막이", "");
            s = s.Replace("스판멀티7부바지", "").Replace("스판멀티바지", "").Replace("테크핏반바지", "").Replace("3부 반바지", "")
                    .Replace("7부패드바지", "").Replace("5부패드바지", "").Replace("9부바지", "").Replace("7부바지", "")
                    .Replace("등산바지", "").Replace("스판바지", "").Replace("골프바지", "").Replace("스포츠바지", "")
                    .Replace("트레이닝 바지", "").Replace("트레이닝바지", "")
                    .Replace("운동복바지", "").Replace("운동복", "");
            s = s.Replace("카라넥 반팔티셔츠", "").Replace("카라넥반팔티", "").Replace("카라티셔츠", "")
                    .Replace("라운드 반팔티셔츠", "").Replace("라운드반팔티", "").Replace("라운드티셔츠", "")
                    .Replace("여성 반팔티셔츠", "").Replace("긴팔티셔츠", "").Replace("베이직티셔츠", "").Replace("티셔츠", "");

            s = s.Replace("스포츠 스커트", "");
            s = s.Replace("스포츠 반장갑", "").Replace("백팩", "");

            s = s.Trim();

            return s;
        }

        protected string RemoveOptionNumbers(string s)
        {
            string pattern = @"\d{4}-\d{2,4}";
            Regex rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"옵션\/사이즈:";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"사이즈:\d{1,2}(\)|\.)";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"사이즈(:|\/)";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"색상:(\d+\.)?";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\s[A-Z]\d{2}_\s";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, " ");
            s = s.Trim();

            pattern = @"0{4}\d";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\d+종\s*중?\s*택1 (옵션선택:)?0\d\)";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\d+종\s*중?\s*택1\/";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"옵션선택\d:0\d\)";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            return s;
        }

        protected string RemoveDuplicates(string s)
        {
            var words = new HashSet<string>();
            s = Regex.Replace(s, "\\w+", m => words.Add(m.Value.ToUpperInvariant()) ? m.Value : String.Empty);
            s = s.Trim();

            return s;
        }

        protected string Beautify(string s)
        {
            s = s.Trim();

            string pattern = @"(\d{4})\|(\d{2})";
            Regex rgx = new Regex(pattern);
            s = rgx.Replace(s, "$1-$2");
            s = s.Trim();

            pattern = @"\^$";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\-$";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\[\]";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\,$";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\s+\,";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, ",");
            s = s.Trim();

            pattern = @"\/\,";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, "");
            s = s.Trim();

            pattern = @"\/\s*\/";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, "/");
            s = s.Trim();

            pattern = @"\s\/+\s";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, " ");
            s = s.Trim();

            pattern = @"\/{2,}";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, " ");
            s = s.Trim();

            pattern = @"\/$";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\,$";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, String.Empty);
            s = s.Trim();

            pattern = @"\s{2,}";
            rgx = new Regex(pattern);
            s = rgx.Replace(s, " ");
            s = s.Trim();

            return s;
        }
    }
}