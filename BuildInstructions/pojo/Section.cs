using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuildInstructions.pojo
{
    [Table("section")]
    public class Section
    {
        [Key]
        public int Id { get; set; }
        public string name { get; set; }
        public string titleLevel { get; set; }
        public string fontName { get; set; }
        public float fontSZ { get; set; }
        public int afterLine { get; set; }
        public int  beforeLine { get; set; }
        public string spacingType { get; set; }
        public float spacingSize { get; set; }
        public int outlineLevel { get; set; }



    }
}
