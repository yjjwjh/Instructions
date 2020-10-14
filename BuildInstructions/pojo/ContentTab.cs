using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuildInstructions.pojo
{
    [Table("contentTab")]
    public class ContentTab
    {
        [Key]
        public int Id { get; set; }

        public int section_id { get; set; }

        public string content { get; set; }
    }
}
