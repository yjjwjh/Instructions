using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuildInstructions.pojo
{
    /// <summary>
    /// 最首端
    /// </summary>
    public class LeaderNode
    {
        public string Name { get; set; }
        public string DataContent { get; set; }
        public List<CentreNode> CentreNodeList { get; set; }



    }
}
