
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;

namespace UAStat
{
    [Table("UserAccount", Schema = "public")]
   public class UserAccount
    {
        [Key]
        [Column("ID")]
        public int Id { get; set; }

        public string Login { get; set; }
        public string INN { get; set; }
        public string OGRN { get; set; }
        public string Company { get; set; }
        public string MarketMembersTypes { get; set; }
    }
}
