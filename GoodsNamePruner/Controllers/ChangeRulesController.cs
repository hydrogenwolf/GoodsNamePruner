using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using GoodsNamePruner.Models;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.SqlClient;

namespace GoodsNamePruner.Controllers
{
    [Authorize]
    public class ChangeRulesController : Controller
    {
        private ChangeRuleDBContext db = new ChangeRuleDBContext();

        // GET: ChangeRules
        public ActionResult Index()
        {
            return View(db.ChangeRules.ToList());
        }

        // GET: ChangeRules/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ChangeRule changeRule = db.ChangeRules.Find(id);
            if (changeRule == null)
            {
                return HttpNotFound();
            }
            return View(changeRule);
        }

        // GET: ChangeRules/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ChangeRules/Create
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 https://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,OwnerID,Before,After")] ChangeRule changeRule)
        {
            changeRule.Before = changeRule.Before.Trim();
            changeRule.After = changeRule.After.Trim();

            var rules = from r in db.ChangeRules
                            where r.Before == changeRule.Before
                            select r;

            if (rules.Count() > 0)
            {
                int ID = 0;
                foreach (var rule in rules)
                {
                    ID = rule.ID;
                }

                ChangeRule theRule = db.ChangeRules.Find(ID);
                theRule.AdjustmentDate = DateTime.Now;
                theRule.After = changeRule.After;
                db.Entry(theRule).State = EntityState.Modified;
            }
            else
            {
                changeRule.OwnerID = "5f39e5a7-f4c7-49af-9b55-796eb8c33d33";
                changeRule.DefinitionDate = DateTime.Now;

                db.ChangeRules.Add(changeRule);
            }

            db.SaveChanges();
            return RedirectToAction("Index");

            /*
            if (ModelState.IsValid)
            {
                db.ChangeRules.Add(changeRule);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(changeRule);
            */
        }

        // GET: ChangeRules/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ChangeRule changeRule = db.ChangeRules.Find(id);
            if (changeRule == null)
            {
                return HttpNotFound();
            }
            return View(changeRule);
        }

        // POST: ChangeRules/Edit/5
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 https://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,OwnerID,DefinitionDate,AdjustmentDate,Before,After")] ChangeRule changeRule)
        {
            if (ModelState.IsValid)
            {
                changeRule.Before = changeRule.Before.Trim();
                changeRule.After = changeRule.After.Trim();
                changeRule.AdjustmentDate = DateTime.Now;

                db.Entry(changeRule).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(changeRule);
        }

        // GET: ChangeRules/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ChangeRule changeRule = db.ChangeRules.Find(id);
            if (changeRule == null)
            {
                return HttpNotFound();
            }
            return View(changeRule);
        }

        // POST: ChangeRules/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            ChangeRule changeRule = db.ChangeRules.Find(id);
            db.ChangeRules.Remove(changeRule);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult Prune()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Convert(HttpPostedFileBase file)
        {
            try
            {
                if (file.ContentLength > 0)
                {
                    DataTable changeRules = new DataTable("ChangeRules");

                    using (SqlConnection connection = new SqlConnection("Data Source=192.168.1.120;Initial Catalog=Noition;Persist Security Info=True;User ID=noition;Password=prune2017$"))
                    {
                        connection.Open();
                        SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM ChangeRules", connection);
                        DataSet ds = new DataSet();

                        adapter.Fill(ds, "ChangeRules");
                        changeRules = ds.Tables["ChangeRules"];
                    }

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

                        bool found = false;
                        //foundRows = changeRules.Select(String.Format("Before = '{0}'", s));
                        // DataTable.Select 메소드에 제약이 많아 대신 foreach 사용
                        foreach (DataRow rule in changeRules.Rows)
                        {
                            if(rule["Before"].ToString() == s)
                            {
                                workbook.ActiveSheet.Range(column).Value = rule["After"] + " " + count;
                                workbook.ActiveSheet.Range(column).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                                //workbook.ActiveSheet.Range(column).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Brown);
                                //workbook.ActiveSheet.Range(column).Font.Bold = true;
                                //workbook.ActiveSheet.Range(column).Font.Italic = true;

                                found = true;
                                break;
                            }
                        }

                        /*
                        if(!found)
                        {
                            workbook.ActiveSheet.Range(column).Value = Beautify(RemoveDuplicates(RemoveOptionNumbers(RemoveWords(s)))) + " " + count;
                        }
                        */
                    }

                    workbook.SaveAs(resultFile);

                    workbook.Close();
                    workbooks.Close();
                    excelApplication.Quit();

                    return new FilePathResult(resultFile, "application/vnd.ms-excel");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return RedirectToAction("Index");
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
