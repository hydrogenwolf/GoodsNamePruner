using System;
using System.Data.Entity;
using System.ComponentModel.DataAnnotations;

namespace GoodsNamePruner.Models
{
    public class ChangeRule
    {
        public int ID { get; set; }
        public string OwnerID { get; set; }
        public DateTime DefinitionDate { get; set; }
        public Nullable<DateTime> AdjustmentDate { get; set; }

        [Display(Name = "상품명")]
        public string Before { get; set; }

        [Display(Name = "간략상품명")]
        public string After { get; set; }
    }

    public class ChangeRuleDBContext : DbContext
    {
        public DbSet<ChangeRule> ChangeRules { get; set; }
    }
}