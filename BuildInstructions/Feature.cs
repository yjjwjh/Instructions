using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BuildInstructions.pojo;

namespace BuildInstructions
{
    public partial class Feature : Form
    {
        

        private Section section;
        public Feature()
        {
            InitializeComponent();
        }
        public void showType(Section section) 
        {
            if (section == null) 
            {
                MessageBox.Show("未查询到此节点！");
                return;
            }
            //在listview中赋值
            this.Text = section.name + "节点特征信息";
            listView1.Items.Clear();//清除所有项和列
            ListViewItem lv1 = new ListViewItem();
            lv1.Text = "标题";
            //listView1.Columns[0].TextAlign = HorizontalAlignment.Center;
            lv1.SubItems.Add(section.titleLevel);
            listView1.Items.Add(lv1);
            lv1 = new ListViewItem();
            lv1.Text = "字体";
            lv1.SubItems.Add(section.fontName);
            listView1.Items.Add(lv1);
            lv1 = new ListViewItem();
            lv1.Text = "字体大小";
            lv1.SubItems.Add(section.fontSZ.ToString());
            listView1.Items.Add(lv1);
            lv1 = new ListViewItem();
            lv1.Text = "大纲级别";
            lv1.SubItems.Add(section.outlineLevel.ToString());
            listView1.Items.Add(lv1);
            lv1 = new ListViewItem();
            lv1.Text = "行距";
            lv1.SubItems.Add("固定   值为" + section.spacingSize.ToString() + "磅");
            listView1.Items.Add(lv1);
            lv1 = new ListViewItem();
            lv1.Text = "间距";
            lv1.SubItems.Add("段前:" + section.afterLine.ToString() + "磅" +","+ "段后:" + section.afterLine.ToString() + "磅");
            listView1.Items.Add(lv1);
        }
    }
}
