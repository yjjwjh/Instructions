using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuildInstructions.pojo
{
    /// <summary>
    /// 中间级
    /// </summary>
    public class CentreNode
    {
        public string Name { get; set; }
        public string DataContent { get; set; }
        public List<LastNode> LastNodeList{ get; set; }


    }
}
