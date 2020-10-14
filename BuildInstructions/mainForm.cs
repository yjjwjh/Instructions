using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BuildInstructions.createDocx;
using BuildInstructions.dao;
using BuildInstructions.pojo;

namespace BuildInstructions
{
    public partial class MainForm : Form
    {
        public static Form form;
        private delegate void MyDelegate(Section section);
        public MainForm()
        {
            InitializeComponent();
            form = this;
        }
        /// <summary>
        /// 程序加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mainForm_Load(object sender, EventArgs e)
        {
            treeView1.ExpandAll();
        }
        /// <summary>
        /// 树形结构节点复选框勾选后
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            TreeNode node = e.Node;
            //查看当前勾选节点的等级,0:根，1：中间级，2：最后一级   目前总共3级
            int level = node.Level;
            TreeNode parentNode = node.Parent;
            if (node.Checked == true)//勾选
            {
                switch (level)
                {
                    case 0:
                        treeView2.Nodes.Add(node.Text, node.Text);//添加子节点
                        break;
                    case 1:
                        TreeNode newParentNode = treeView2.Nodes[parentNode.Text];
                        newParentNode.Nodes.Add(node.Text, node.Text);//添加子节点
                        newParentNode.ExpandAll();
                        break;
                    case 2:
                        //获取最顶级
                        TreeNode supNode = parentNode.Parent;
                        TreeNode newSupNode = treeView2.Nodes[supNode.Text];
                        //获取中间级
                        newParentNode = newSupNode.Nodes[parentNode.Text];
                        //添加子节点
                        newParentNode.Nodes.Add(node.Text, node.Text);
                        newParentNode.ExpandAll();
                        break;
                }
            }
            else //取消勾选
            {
                switch (level)
                {
                    case 0:
                        treeView2.Nodes.RemoveByKey(node.Text);//删除子节点
                        break;
                    case 1:
                        TreeNode newParentNode = treeView2.Nodes[parentNode.Text];
                        newParentNode.Nodes.RemoveByKey(node.Text);//删除子节点
                        break;
                    case 2:
                        //获取最顶级
                        TreeNode supNode = parentNode.Parent;
                        TreeNode newSupNode = treeView2.Nodes[supNode.Text];
                        //获取中间级
                        newParentNode = newSupNode.Nodes[parentNode.Text];
                        //删除子节点
                        newParentNode.Nodes.RemoveByKey(node.Text);
                        break;
                }

            }
        }
        /// <summary>
        /// 树形结构节点复选框勾选前
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_BeforeCheck(object sender, TreeViewCancelEventArgs e)
        {
            TreeNode node = e.Node;
            //查看当前勾选节点的等级,0:根，1：中间级，2：最后一级   目前总共3级
            int level = node.Level;
            if (node.Checked == false)//勾选
            {
                if (level != 0)
                {
                    //获取父级
                    TreeNode parentNode = node.Parent;
                    if (parentNode.Checked == false)
                    {
                        parentNode.Checked = true;
                    }
                }
            }
            else //取消勾选
            {
                if (level != 2)
                {
                    //遍历下级的勾选状态
                    foreach (TreeNode n in node.Nodes)
                    {
                        if (n.Checked == true)
                        {
                            n.Checked = false;
                        }

                    }
                }
            }
        }
        /// <summary>
        /// 向上移动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            TreeNode node = treeView2.SelectedNode;
            if (node == null)
            {
                MessageBox.Show("未选择任何节点！");
                return;
            }
            //获取当前节点在同级目录的位置（index）
            int index = node.Index;
            if (index == 0)
            {
                //获取当前节点的父级
                TreeNode parentNode = node.Parent;
                if (parentNode == null)
                {
                    MessageBox.Show("已在最上级！");
                    return;
                }
                //判断父级节点的等级
                int level = parentNode.Level;
                if (level == 0)
                {
                    node.Remove();
                    treeView2.Nodes.Insert(parentNode.Index, node);
                    //treeView2.Nodes[parentNode.Index].BackColor = Color.Blue;
                    treeView2.SelectedNode = treeView2.Nodes[parentNode.Index - 1];

                }
                else
                {
                    TreeNode supNode = parentNode.Parent;
                    node.Remove();
                    supNode.Nodes.Insert(parentNode.Index, node);
                    //supNode.Nodes[parentNode.Index].BackColor = Color.Blue;
                    treeView2.SelectedNode = supNode.Nodes[parentNode.Index - 1];
                }
            }
            else
            {
                //获取上一个树节点
                TreeNode upNode = node.PrevNode;
                node.Remove();
                upNode.Nodes.Insert(upNode.Nodes.Count, node);
                //upNode.Nodes[upNode.Nodes.Count].BackColor = Color.Blue;
                treeView2.SelectedNode = upNode.Nodes[upNode.Nodes.Count - 1];
            }
            //treeView2.SelectedNode.ImageIndex = 0;
            treeView2.ExpandAll();
        }
        /// <summary>
        /// 向下移动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            TreeNode node = treeView2.SelectedNode;
            if (node == null)
            {
                MessageBox.Show("未选择任何节点！");
                return;
            }
            //获取当前节点在同级目录的位置（index）
            int index = node.Index;
            int level = node.Level;
            //获取下一树节点
            TreeNode nextNode = node.NextNode;
            if (nextNode == null)//未有下一个节点
            {
                if (level == 0 && index == treeView2.Nodes.Count - 1)
                {
                    MessageBox.Show("已在最下级");
                    return;
                }
                else
                {
                    //获取父级节点
                    TreeNode parentNode = node.Parent;
                    //判断父级节点的等级
                    if (parentNode.Level == 0)
                    {
                        node.Remove();
                        treeView2.Nodes.Insert(parentNode.Index + 1, node);
                        //treeView2.Nodes[parentNode.Index + 1].BackColor = Color.Blue;
                        treeView2.SelectedNode = treeView2.Nodes[parentNode.Index + 1];
                    }
                    else
                    {
                        //获取父级的父级
                        TreeNode supNode = parentNode.Parent;
                        node.Remove();
                        supNode.Nodes.Insert(parentNode.Index + 1, node);
                        //supNode.Nodes[parentNode.Index + 1].BackColor = Color.Blue;
                        treeView2.SelectedNode = supNode.Nodes[parentNode.Index + 1];
                    }
                }
            }
            else
            {
                node.Remove();
                nextNode.Nodes.Insert(0, node);
                //nextNode.Nodes[0].BackColor = Color.Blue;
                treeView2.SelectedNode = nextNode.Nodes[0];
            }
            //treeView2.SelectedNode.ImageIndex = 0;
            treeView2.ExpandAll();//展开所有节点
        }
        /// <summary>
        /// 向组合目录添加按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            //获取目录框中选定的节点
            TreeNode node = treeView1.SelectedNode;
            if (node == null)
            {
                MessageBox.Show("目录树形结构中未选择节点！");
                return;
            }
            TreeNode parentNode = node.Parent;
            if (parentNode == null)//根级
            {
                //判断组合目录树形中是否有
                TreeNode newNode = treeView2.Nodes[node.Text];
                if (newNode == null)//没有
                {
                    treeView2.Nodes.Add(node.Text, node.Text);
                }
                else
                {
                    MessageBox.Show("该节点已存在！");

                }
            }
            else
            {
                //获取父级的父级
                TreeNode supNode = parentNode.Parent;
                //判断父级的父级
                if (supNode == null)//parentNode就是根级，node就是第二级
                {
                    //判断parentNode是否在组合中存在
                    TreeNode newParentNode = treeView2.Nodes[parentNode.Text];
                    if (newParentNode == null)
                    {
                        newParentNode = treeView2.Nodes.Add(parentNode.Text, parentNode.Text);
                        newParentNode.Nodes.Add(node.Text, node.Text);
                    }
                    else
                    {
                        //查看组合中newParentNode是否存在node
                        TreeNode newNode = newParentNode.Nodes[node.Text];
                        if (newNode == null)
                        {
                            newParentNode.Nodes.Add(node.Text, node.Text);
                        }
                        else
                        {
                            MessageBox.Show("该节点已存在！");
                        }
                    }
                }
                else//supNode就是根级，parentNode就是第二级，node就是第三级
                {
                    //判断supNode是否在组合中存在
                    TreeNode newSupNode = treeView2.Nodes[supNode.Text];
                    if (newSupNode == null)
                    {
                        newSupNode = treeView2.Nodes.Add(supNode.Text, supNode.Text);
                        TreeNode newParentNode = newSupNode.Nodes.Add(parentNode.Text, parentNode.Text);
                        newParentNode.Nodes.Add(node.Text, node.Text);
                    }
                    else
                    {
                        //在newSupNode中判断newParentNode是否存在
                        //判断parentNode是否在组合中存在
                        TreeNode newParentNode = newSupNode.Nodes[parentNode.Text];
                        if (newParentNode == null)
                        {
                            newParentNode = newSupNode.Nodes.Add(parentNode.Text, parentNode.Text);
                            newParentNode.Nodes.Add(node.Text, node.Text);
                        }
                        else
                        {
                            //查看组合中newParentNode是否存在node
                            TreeNode newNode = newParentNode.Nodes[node.Text];
                            if (newNode == null)
                            {
                                newParentNode.Nodes.Add(node.Text, node.Text);
                            }
                            else
                            {
                                MessageBox.Show("该节点已存在！");
                            }
                        }

                    }
                }
            }
            treeView2.ExpandAll();

        }
        /// <summary>
        /// 删除组合目录节点
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            //获取目录框中选定的节点
            TreeNode node = treeView2.SelectedNode;
            if (node == null)
            {
                MessageBox.Show("目录树形结构中未选择节点！");
                return;
            }
            node.Remove();
        }
        /// <summary>
        /// 双击节点
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode treeNode= e.Node;
            //从数据库中查询节点信息
            using (var myEntity = new MyEntity())
            {
                var se = (from s in myEntity.Section
                              where s.name==treeNode.Text
                          select s).ToList();

                //MessageBox.Show(se.Count().ToString());
                Feature feature = new Feature();
                MyDelegate myDelegate = new MyDelegate(feature.showType);
                myDelegate(se.First());
                feature.Show();
            }
           

           


        }
        /// <summary>
        /// 生成word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //获取组合目录的节点个数
            int count = treeView2.Nodes.Count;
            if (count == 0)
            {
                MessageBox.Show("组合目录中未有章节！");
                return;
            }
            List<LeaderNode> leaderNodes = new List<LeaderNode>();
            foreach (TreeNode node in treeView2.Nodes)
            {
                LeaderNode leaderNode = new LeaderNode();
                leaderNode.Name = node.Text;
                if (node.Nodes.Count != 0)
                {
                    List<CentreNode> centreNodes = new List<CentreNode>();
                    foreach (TreeNode node1 in node.Nodes)
                    {
                        CentreNode centreNode = new CentreNode();
                        centreNode.Name = node1.Text;
                        if (node1.Nodes.Count != 0)
                        {
                            List<LastNode> lastNodes = new List<LastNode>();
                            foreach (TreeNode node2 in node1.Nodes)
                            {
                                LastNode lastNode = new LastNode();
                                lastNode.Name = node2.Text;
                                lastNodes.Add(lastNode);
                            }
                            centreNode.LastNodeList = lastNodes;
                        }
                        centreNodes.Add(centreNode);
                    }
                    leaderNode.CentreNodeList = centreNodes;
                }
                leaderNodes.Add(leaderNode);
            }
            CreateWord createWord= CreateWord.GetInstance(leaderNodes);

            //CreateWord cw = new CreateWord();
            createWord.Create();
        }


        //段落缩进   返回值为对应的缩进距离
        //(fontname：文字类型名称   fontsize：文字大小    fontcount：缩进数目 fontstyle：文字类型（斜体、粗体...）)
        public static int Indentation(String fontname, int fontsize, int fontnum, FontStyle fontstyle)
        {
            Graphics gp = form.CreateGraphics();
            gp.PageUnit = GraphicsUnit.Point;
            SizeF size = gp.MeasureString("字", new Font(fontname, fontsize * 0.75F, fontstyle));
            return (int)size.Width * fontnum * 10;
        }
    }
}

