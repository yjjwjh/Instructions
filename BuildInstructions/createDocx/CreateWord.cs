using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.Drawing;
using BuildInstructions.pojo;
using BuildInstructions.dao;
using System.Text.RegularExpressions;

namespace BuildInstructions.createDocx
{
    public class CreateWord
    {

        private static CreateWord uniqueInstance;

        private List<LeaderNode> leaderNodes;

        private CreateWord() { }

        /// <summary>
        /// 单例
        /// </summary>
        /// <param name="nodes"></param>
        /// <returns></returns>
        public static CreateWord GetInstance(List<LeaderNode> nodes)
        {
            if (uniqueInstance == null)
            {
                uniqueInstance = new CreateWord();
            }
            uniqueInstance.leaderNodes = nodes;
            return uniqueInstance;
        }

        /// <summary>
        /// 创建文档
        /// </summary>
        public void Create()
        {
            //readWord();
            XWPFDocument doc = new XWPFDocument();
            //设置业内边距
            doc.Document.body.sectPr = new CT_SectPr();
            CT_SectPr m_sectpr = doc.Document.body.sectPr;
            m_sectpr.pgMar.top = "2265";//上40mm
            m_sectpr.pgMar.bottom = "1985";//下350mm
            m_sectpr.pgMar.left = 1417;//左25mm
            m_sectpr.pgMar.right = 1133;//右20mm


            CreateHomePage(doc);//首页
            CreateProjectData(doc);//项目信息
            //CreatePageHeaderFooter(doc,m_sectpr);//页眉页脚
            CreateChaptersAndSections(doc);

            //生成Word文件,保存对话框
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //设置文件类型
            saveFileDialog.Filter = "Microsoft Word文件(*.docx)|*.docx";
            //保存对话框是否记忆上次打开的目录
            saveFileDialog.RestoreDirectory = true;
            //设置默认的文件名称
            saveFileDialog.FileName = "施工说明书（组合）";
            if (saveFileDialog.ShowDialog() != DialogResult.OK) return;
            //获取文件路径
            string flieName = saveFileDialog.FileName.ToString();
            FileStream out1 = new FileStream(flieName, FileMode.Create);
            doc.Write(out1);
            out1.Close();



        }
        /// <summary>
        /// 创建首页
        /// </summary>
        /// <param name="doc"></param>
        private void CreateHomePage(XWPFDocument doc)
        {
            CT_SectPr cT_Sect = new CT_SectPr();
            cT_Sect.pgMar.top = "1700";//上30mm
            cT_Sect.pgMar.bottom = "1700";//下30mm
                                          //cT_Sect.pgMar.left = 1417;//左25mm
                                          //cT_Sect.pgMar.right = 1133;//右20mm


            XWPFParagraph p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.BOTH;
            p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            XWPFRun r1 = p1.CreateRun();
            r1.SetText("图号：S123456S-D0101-01                       电力行业（送电、变电）专业甲级");
            r1.FontFamily = "宋体";
            r1.FontSize = 12;

            p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.BOTH;
            p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            r1 = p1.CreateRun();
            r1.SetText("版本：A	                                    	 勘测设计证号：A144000587");
            r1.FontFamily = "宋体";
            r1.FontSize = 12;


            for (int i = 0; i < 4; i++)
            {
                p1 = doc.CreateParagraph();
                p1.Alignment = ParagraphAlignment.CENTER;
                p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            }

            p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.CENTER;
            p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            r1 = p1.CreateRun();
            r1.SetText("110kV");
            r1.FontFamily = "Times New Roman";
            r1.FontSize = 22;
            r1 = p1.CreateRun();
            r1.SetText("濂泉送电线路工程");
            r1.FontFamily = "宋体";
            r1.FontSize = 22;

            p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.CENTER;
            p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);

            p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.CENTER;
            p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            r1 = p1.CreateRun();
            r1.SetText("施工图设计说明书");
            r1.FontFamily = "宋体";
            r1.FontSize = 18;

            for (int i = 0; i < 14; i++)
            {
                p1 = doc.CreateParagraph();
                p1.Alignment = ParagraphAlignment.CENTER;
                p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            }
            p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.CENTER;
            p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            r1 = p1.CreateRun();
            r1.SetText("广 州 电 力 设 计 院 有 限 公 司");
            r1.FontFamily = "宋体";
            r1.FontSize = 14;

            p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.CENTER;
            p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            r1 = p1.CreateRun();
            r1.SetText(DateTime.Today.Year + "年" + DateTime.Today.Month + "月  广州");
            r1.FontFamily = "宋体";
            r1.FontSize = 14;

            p1.CreateRun().AddBreak(BreakType.COLUMN);//插入空白页



        }
        /// <summary>
        /// 创建工程信息
        /// </summary>
        /// <param name="doc"></param>
        private void CreateProjectData(XWPFDocument doc)
        {
            XWPFParagraph p1 = doc.CreateParagraph();
            p1.CreateRun().AddBreak();//新建页
            for (int i = 0; i < 2; i++)
            {
                p1 = doc.CreateParagraph();
                p1.Alignment = ParagraphAlignment.CENTER;
                p1.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            }
            XWPFParagraph p2 = doc.CreateParagraph();
            p2.Alignment = ParagraphAlignment.CENTER;
            p2.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            XWPFRun r1 = p2.CreateRun();
            r1.SetText("110KV");
            r1.FontFamily = "Times New Roman";
            r1.FontSize = 22;
            XWPFRun r2 = p2.CreateRun();
            r2.SetText("濂泉送电线路工程");
            r2.FontFamily = "宋体";
            r2.FontSize = 22;
            r2.AddCarriageReturn();
            r2.AddCarriageReturn();
            r2.AddCarriageReturn();



            p2 = doc.CreateParagraph();
            p2.Alignment = ParagraphAlignment.CENTER;
            p2.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            r2 = p2.CreateRun();
            r2.SetText("施工图设计说明书");
            r2.FontFamily = "宋体";
            r2.FontSize = 18;
            r2.AddCarriageReturn();
            r2.AddCarriageReturn();
            r2.AddCarriageReturn();

            p2 = doc.CreateParagraph();
            p2.Alignment = ParagraphAlignment.BOTH;
            p2.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            p2.IndentationLeft = 1276;
            p2.SpacingAfter = 360;
            r2 = p2.CreateRun();
            r2.SetText("批      准： 叶其革");
            r2.FontFamily = "宋体";
            r2.FontSize = 14;

            p2 = doc.CreateParagraph();
            p2.Alignment = ParagraphAlignment.BOTH;
            p2.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            p2.IndentationLeft = 1276;
            p2.SpacingAfter = 360;
            r2 = p2.CreateRun();
            r2.SetText("审      核： 陈沛民     李锋");
            r2.FontFamily = "宋体";
            r2.FontSize = 14;

            p2 = doc.CreateParagraph();
            p2.Alignment = ParagraphAlignment.BOTH;
            p2.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            p2.IndentationLeft = 1276;
            p2.SpacingAfter = 360;
            r2 = p2.CreateRun();
            r2.SetText("校      核： 刘莉华     唐兴佳");
            r2.FontFamily = "宋体";
            r2.FontSize = 14;

            p2 = doc.CreateParagraph();
            p2.Alignment = ParagraphAlignment.BOTH;
            p2.setSpacingBetween(1.5, LineSpacingRule.AUTO);
            p2.IndentationLeft = 1276;
            p2.SpacingAfter = 360;
            r2 = p2.CreateRun();
            r2.SetText("设      计： 杨健锐     刘俊勇");
            r2.FontFamily = "宋体";
            r2.FontSize = 14;
            p2.CreateRun().AddBreak();
        }
        /// <summary>
        /// 创建页眉页脚
        /// </summary>
        /// <param name="doc"></param>
        private void CreatePageHeaderFooter(XWPFDocument doc, CT_SectPr m_Sectpr)
        {
            XWPFParagraph p1 = doc.CreateParagraph();
            p1.CreateRun().AddBreak();//新建页

            //创建页眉

            CT_Hdr m_hdr = new CT_Hdr();
            m_hdr.Items = new System.Collections.ArrayList();


            CT_P m_p = m_hdr.AddNewP();
            CT_PPr cT_PPr = m_p.AddNewPPr();
            cT_PPr.AddNewJc().val = ST_Jc.both;//两端对齐
            cT_PPr.AddNewSpacing().beforeLines = "370";


            CT_R cT_R = m_p.AddNewR();
            cT_R.AddNewT().Value = "110kV濂泉（沙河）送电线路工程                 施工图设计说明书                     S123456S-D0101-01";//页眉内容
            CT_RPr cT_RPr = cT_R.AddNewRPr();
            cT_RPr.AddNewSz().val = (ulong)18;
            cT_RPr.AddNewSzCs().val = (ulong)18;
            cT_RPr.AddNewRFonts().ascii = "宋体";
            //cT_RPr.AddNewU().val=ST_Underline.single;//下划线

            //创建页眉关系（headern.xml）
            XWPFRelation Hrelation = XWPFRelation.HEADER;
            XWPFHeader m_h = (XWPFHeader)doc.CreateRelationship(Hrelation, XWPFFactory.GetInstance(), 3);
            //doc.CreateFootnotes();
            //设置页眉
            m_h.SetHeaderFooter(m_hdr);
            CT_HdrFtrRef m_HdrFtr = m_Sectpr.AddNewHeaderReference();
            m_HdrFtr.type = ST_HdrFtr.@default;
            //m_h.GetRelationById(m_HdrFtr.id);
            m_HdrFtr.id = m_h.GetPackageRelationship().Id;



            //创建页脚
            CT_Ftr m_ftr = new CT_Ftr();
            m_ftr.Items = new System.Collections.ArrayList();
            CT_SdtBlock m_Sdt = new CT_SdtBlock();
            CT_SdtPr m_SdtPr = m_Sdt.AddNewSdtPr();
            CT_SdtDocPart m_SdDocPartObj = m_SdtPr.AddNewDocPartObj();
            m_SdDocPartObj.AddNewDocPartGallery().val = "PageNumbers (Bottom of Page)";
            m_SdDocPartObj.docPartUnique = new CT_OnOff();
            CT_SdtContentBlock m_SdtContent = m_Sdt.AddNewSdtContent();
            CT_P m_SdtContentP = m_SdtContent.AddNewP();
            CT_PPr m_SdtContentPpr = m_SdtContentP.AddNewPPr();
            m_SdtContentPpr.AddNewJc().val = ST_Jc.center;
            m_SdtContentP.Items = new System.Collections.ArrayList();
            CT_SimpleField m_fldSimple = new CT_SimpleField();
            m_fldSimple.instr = " PAGE   \\*MERGEFORMAT";

            //页码字体大小
            CT_R m_r = new CT_R();
            CT_RPr m_Rpr = m_r.AddNewRPr();
            m_Rpr.AddNewRFonts().ascii = "宋体";
            m_Rpr.AddNewSz().val = (ulong)18;
            m_Rpr.AddNewSzCs().val = (ulong)18;
            m_r.AddNewT().Value = "1";//页数
            m_fldSimple.Items.Add(m_r);

            m_SdtContentP.Items.Add(m_fldSimple);
            m_ftr.Items.Add(m_Sdt);

            //m_ftr.AddNewP().AddNewR().AddNewT().Value = "fff";//页脚内容
            //m_ftr.AddNewP().AddNewPPr().AddNewJc().val = ST_Jc.center;
            //创建页脚关系（footern.xml）
            XWPFRelation Frelation = XWPFRelation.FOOTER;
            XWPFFooter m_f = (XWPFFooter)doc.CreateRelationship(Frelation, XWPFFactory.GetInstance(), doc.FooterList.Count + 1);
            //设置页脚
            m_f.SetHeaderFooter(m_ftr);
            m_HdrFtr = m_Sectpr.AddNewFooterReference();
            m_HdrFtr.type = ST_HdrFtr.@default;
            m_HdrFtr.id = m_f.GetPackageRelationship().Id;

        }

        /// <summary>
        /// 创建章节
        /// </summary>
        /// <param name="doc"></param>
        private void CreateChaptersAndSections(XWPFDocument doc)
        {
            using (var myEntity = new MyEntity())
            {
                //读取模板
                FileStream stream = File.OpenRead(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "模板\\施工说明书(模板).docx");
                XWPFDocument readDoc = new XWPFDocument(stream);
                //创建段落
                XWPFParagraph paragraph = doc.CreateParagraph();
                paragraph.CreateRun().AddBreak();//新建页
                for (int i = 0; i < leaderNodes.Count; i++)
                {
                    LeaderNode leaderNode = leaderNodes[i];
                    paragraph = doc.CreateParagraph();//创建段落
                    //行距
                    paragraph.setSpacingBetween(30, LineSpacingRule.EXACT);//固定值，30磅
                    //间距
                    paragraph.SpacingAfterLines = 8;//上
                    paragraph.SpacingBeforeLines = 8;//下
                    //对齐方式
                    paragraph.Alignment = ParagraphAlignment.BOTH;//两端对齐
                    //大纲级别
                    CT_DecimalNumber cT_DecimalNumber = new CT_DecimalNumber();
                    cT_DecimalNumber.val = "0级";//1级目录
                    paragraph.GetCTP().AddNewPPr().outlineLvl = cT_DecimalNumber;
                    //标题信息
                    XWPFRun run = paragraph.CreateRun();
                    run.SetText((i + 1).ToString());//1级
                    run.FontSize = 14;//字体大小
                    run.FontFamily = "Times New Roman";//字体
                    run = paragraph.CreateRun();
                    run.SetText( " " + leaderNode.Name);//1级
                    run.FontSize = 14;//字体大小
                    run.FontFamily = "宋体";//字体
                    //根据节点信息在数据库中查询
                    var se = (from s in myEntity.Section
                              where s.name == leaderNode.Name
                              select s).ToList();
                    Section section = se.First();//章节对象
                    //判断章节中是否存在内容
                    bool flag = IsContainTable(leaderNode.Name, doc, readDoc);
                    if (flag)
                    {
                        //根据节点对象id信息在数据库内容表中查询
                        var contents = (from c in myEntity.ContentTab
                                        where c.section_id == section.Id
                                        select c).ToList();
                        if (contents.Count > 0)
                        {
                            foreach (ContentTab contentTab in contents)
                            {
                                paragraph = doc.CreateParagraph();//创建段落
                                paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, new FontStyle());//首行缩进
                                XWPFRun r = paragraph.CreateRun();
                                r.SetText(contentTab.content);
                                r.FontFamily = "宋体";
                                r.FontSize = 14;
                            }
                        }
                    }
                    if (leaderNode.CentreNodeList != null) 
                    {
                        //遍历二级目录
                        for (int j = 0; j < leaderNode.CentreNodeList.Count; j++)
                        {
                            CentreNode centreNode = leaderNode.CentreNodeList[j];
                            paragraph = doc.CreateParagraph();//创建段落
                            //行距
                            paragraph.setSpacingBetween(30, LineSpacingRule.EXACT);//固定值，30磅
                            //对齐方式
                            paragraph.Alignment = ParagraphAlignment.BOTH;//两端对齐
                            //大纲级别
                            cT_DecimalNumber = new CT_DecimalNumber();
                            cT_DecimalNumber.val = "1级";//2级目录
                            paragraph.GetCTP().AddNewPPr().outlineLvl = cT_DecimalNumber;
                            //标题信息
                            run = paragraph.CreateRun();
                            run.SetText((i + 1) + "." + (j + 1));//2级
                            run.FontSize = 14;//字体大小
                            run.FontFamily = "Times New Roman";//字体
                            run = paragraph.CreateRun();
                            run.SetText(" " + centreNode.Name);//2级
                            run.FontSize = 14;//字体大小
                            run.FontFamily = "宋体";//字体
                            //根据节点信息在数据库中查询
                            se = (from s in myEntity.Section
                                  where s.name == centreNode.Name
                                  select s).ToList();
                            section = se.First();//章节对象
                            //判断章节中是否存在内容
                            flag = IsContainTable(centreNode.Name, doc, readDoc);
                            if (flag)
                            {
                                //根据节点对象id信息在数据库内容表中查询
                                var contents = (from c in myEntity.ContentTab
                                                where c.section_id == section.Id
                                                select c).ToList();
                                if (contents.Count > 0)
                                {
                                    foreach (ContentTab contentTab in contents)
                                    {
                                        paragraph = doc.CreateParagraph();//创建段落
                                        paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, new FontStyle());//首行缩进
                                        XWPFRun r = paragraph.CreateRun();
                                        r.SetText(contentTab.content);
                                        r.FontFamily = "宋体";
                                        r.FontSize = 14;
                                    }
                                }
                            }
                            if (centreNode.LastNodeList != null)
                            {
                                //遍历三级目录
                                for (int k = 0; k < centreNode.LastNodeList.Count; k++)
                                {
                                    LastNode lastNode = centreNode.LastNodeList[k];
                                    paragraph = doc.CreateParagraph();//创建段落
                                    //行距
                                    paragraph.setSpacingBetween(30, LineSpacingRule.EXACT);//固定值，30磅
                                    //对齐方式
                                    paragraph.Alignment = ParagraphAlignment.BOTH;//两端对齐
                                    //大纲级别
                                    cT_DecimalNumber = new CT_DecimalNumber();
                                    cT_DecimalNumber.val = "2级";//3级目录
                                    paragraph.GetCTP().AddNewPPr().outlineLvl = cT_DecimalNumber;
                                    //标题信息
                                    run = paragraph.CreateRun();
                                    run.SetText((i + 1) + "." + (j + 1) + "." + (k + 1));//3级
                                    run.FontSize = 14;//字体大小
                                    run.FontFamily = "Times New Roman";//字体
                                    run = paragraph.CreateRun();
                                    run.SetText(" " + lastNode.Name);//3级
                                    run.FontSize = 14;//字体大小
                                    run.FontFamily = "宋体";//字体
                                    //根据节点信息在数据库中查询
                                    se = (from s in myEntity.Section
                                          where s.name == lastNode.Name
                                          select s).ToList();
                                    section = se.First();//章节对象
                                                         //判断章节中是否存在内容
                                    flag = IsContainTable(lastNode.Name, doc, readDoc);
                                    if (flag)
                                    {
                                        //根据节点对象id信息在数据库内容表中查询
                                        var contents = (from c in myEntity.ContentTab
                                                        where c.section_id == section.Id
                                                        select c).ToList();
                                        if (contents.Count > 0)
                                        {
                                            foreach (ContentTab contentTab in contents)
                                            {
                                                paragraph = doc.CreateParagraph();//创建段落
                                                paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, new FontStyle());//首行缩进
                                                XWPFRun r = paragraph.CreateRun();
                                                r.SetText(contentTab.content);
                                                r.FontFamily = "宋体";
                                                r.FontSize = 14;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                stream.Close();
            }
        }
        /// <summary>
        /// 插入内容表格
        /// </summary>
        /// <param name="nodeName"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        private bool IsContainTable(string nodeName, XWPFDocument doc, XWPFDocument readDoc)
        {
            bool flag = false;
            XWPFParagraph paragraph;
            XWPFRun run;
            int index = doc.Tables.Count;
            switch (nodeName) 
            {
                case "工程技术特性表":
                    CopyTable(readDoc,0, index, doc);//复制表
                    break;
                case "对初步设计评审意见的执行情况（暂缺）":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("对初步设计评审意见的执行情况详见表1.5。");
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.Alignment = ParagraphAlignment.CENTER;
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("表1.5  对初步设计评审意见的执行情况");
                    CopyTable(readDoc, 1, index, doc);//复制表
                    break;
                case "强制性条文执行情况":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据国家和电力行业现行的强制性条文及《基建工程强制性条文实施办法》，编制《工程建设标准强制性条文执行计划》，指导施工图设计严格执行强制性条文及南方电网公司电网反事故措施，将相关强制性条文落实到每一册图；并在相关施工图完成后对照强条执行计划完成《工程设计强制性条文执行检查表》，从编制执行计划到落实执行检查表，加强设计过程中对执行《工程建设标准强制性条文》管理。");
                    
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("本工程送电线路部分共执行强制性条文4条，其中电缆线路部分执行4条。");
                    
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    //paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("1）《电力工程电缆设计标准》GB50217-2018电缆线路工程质量强制性条文执行方案");
                    CopyTable(readDoc, 2, index, doc);//复制表

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    //paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("2）《电气装置安装工程电缆线路施工及验收标准》GB50168-2018强制性条文执行方案");
                    index = doc.Tables.Count;
                    CopyTable(readDoc, 3, index, doc);//复制表
                    break;
                case "南网反措执行情况":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("《关于主网基建项目落实防范重大电气火灾及故障专项反事故措施相关工作的通知》（广供电基部[2019]28 号）、《南网反事故措施（2020年版）》设计执行情况如下：");
                    CopyTable(readDoc, 4, index, doc);//复制表
                    break;
                case "气象条件":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据广东省气象局提供的资料，参照架空线路的气象条件，本期电缆线路工程按下表条件设计。");
                    CopyTable(readDoc, 5, index, doc);//复制表
                    break;
                case "电缆型式和导体截面积":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据初设批复，本期110kV电缆截面采用1200mm2。设计推荐采用干式交联聚乙烯绝缘电力电缆。本工程电缆暂未订货，按JB/T 10181.11~10181.32-2014标准计算1200mm2导体截面电缆导体载流量。载流量计算结果如表3.4-1、3.4-2所示（最终电缆导体载流量应以订货后电缆生产厂家提供为准）。");
                    
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.Alignment = ParagraphAlignment.CENTER;
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("表3.4-1    电缆导体载流量表一");
                    CopyTable(readDoc, 6, index, doc);//复制表

                    paragraph = doc.CreateParagraph();
                    index = doc.Tables.Count;
                    CopyTable(readDoc, 7, index, doc);//复制表

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据上表计算结果可知：按隧道内环境空气温度为35℃考虑（当隧道内环境空气温度超过35℃，自动开启隧道通风设备；同时，需要运行调度部门采取监测措施，适当控制温度及载流量），110kV濂泉至永福双回电缆线路采用导体截面为1200 mm2电缆在隧道内敷设、双回直埋（33℃、深1. 2 m）、双回穿管（33℃、深1.5m）时，满足规划输送容量要求；但采用三回电缆沟、三回穿管、双回穿管（26℃、深5.0m）敷设型式时均无法满足规划输送容量要求。");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("由于实际运行中，电缆线路主要在不等负荷的情况下运行（电缆结构与负荷均不同），即并行的多回路电缆线路同时出现N-1的运行方式的概率极小，因此设计建议考虑采用不等负荷运行方式下的载流量计算方式来校核3回路电缆同路径敷设时导体载流量，具体为：");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("假设与濂泉至永福双回110kV电缆线路同路径敷设的第3回路电缆也为3T接线第一段，按带3×63MVA主变考虑输送容量，则其中任1回处于N-1运行方式，并且导体温度达到90℃，其余2回处于正常运行方式（变电站规模为3×63MVA，线变组3T接线第一段，正常运行时输送662A负荷考虑），则处于N-1运行方式的电缆导体载流量计算结果如表3.2.4-2所示：");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.Alignment = ParagraphAlignment.CENTER;
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("表3.4-2    电缆导体载流量表二");
                    paragraph = doc.CreateParagraph();
                    index = doc.Tables.Count;
                    CopyTable(readDoc, 8, index, doc);//复制表

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据上表计算可知结果：本工程双回110kV电缆线路均采用导体截面为1200 mm2电缆可以满足系统输送容量要求。");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("另外，为满足埋地电缆的防蚁要求，本工程在隧道外敷设的电缆选用聚乙烯（PE-ST7）与绿色环保型防蚁材料双层结构混合护套，并需采用防蚁措施，如埋设防蚁药包；为满足隧道内敷设的电缆的防火要求，在隧道内敷设的电缆选用聚氯乙烯（PVC-ST2）材料绝缘外护套，阻燃等级按阻燃A级考虑。具体电缆型号如下：");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("隧道外采用交联聚乙烯绝缘皱纹铝套或焊接皱纹铝套聚乙烯护套纵向阻水电力电缆，型号YJLW03-Z  64/110  1×1200  GB/T 11017.2-2014；");


                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("隧道内采用交联聚乙烯绝缘皱纹铝套或焊接皱纹铝套聚氯乙烯护套纵向阻水电力电缆，型号YJLW02-Z  64/110  1×1200  GB/T 11017.2-2014。");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("由于涉及迁改的一段部分在综合管廊内，部分在综合管廊外，根据生纪[2018]1号，110kV空机线在综合管廊内敷设小于100m，为满足防蚁的要求，电缆选用聚乙烯（PE-ST7）与绿色环保型防蚁材料双层结构混合护套，具体为交联聚乙烯绝缘皱纹铝套或焊接皱纹铝套聚乙烯护套纵向阻水电力电缆，型号为YJLW03-Z  64/110  1×800  GB/T 11017.2-2014。110kV金机线在综合管廊内敷设大于100m，为满足防火要求，电缆选用聚氯乙烯（PVC-ST2）材料绝缘外护套，阻燃等级按阻燃A级考虑。具体为交联聚乙烯绝缘皱纹铝套或焊接皱纹铝套聚氯乙烯护套纵向阻水电力电缆，型号为YJLW02-Z 64/110 1×800 GB/T 11017.2-2014。");

                    break;
                case "正常情况下，电缆金属护套的感应电压最大计算结果如下：":
                    CopyTable(readDoc, 9, index, doc);//复制表

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据《电力工程电缆设计标准》，电缆金属护套感应电压一般要求控制在50V以内，当电缆金属护层采取隔离措施后不得超过300V。本工程在变电站外的电缆全部敷设在电缆专用隧道内或埋置在地下，同时电缆金属护层外尚有非金属外护套，除运行检修人员外其他人不易触及电缆金属护层。另外，交叉互联接地保护箱设计亦放置于专用工作井内，除运行检修人员外其他人不易触及电缆金属护层。从上表可见，电缆金属护套感应电压满足电缆线路设计规程规定。");

                    break;
                case "敷设方式统计":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("本工程电缆线路主要采用直埋、穿管、电缆沟、顶管以及电力隧道等敷设型式。各种敷设方式统计如下：");

                    CopyTable(readDoc, 10, index, doc);//复制表
                    break;
                case "3C绿色电网说明及评价":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据南方电网公司《3C 绿色电网建设评价标准（输电线路绿色部分）》的要求，对本工程新建电缆线路所达到的3C绿色指标情况进行评价，具体列表如下：");
                    
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.Alignment = ParagraphAlignment.CENTER;
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("表7-1 电缆线路3C绿色指标情况评价表");
                    CopyTable(readDoc, 11, index, doc);//复制表
                    paragraph = doc.CreateParagraph();
                    index = doc.Tables.Count;
                    CopyTable(readDoc, 12, index, doc);//复制表

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    //paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("电缆线路3C绿色电网建设评价等级划分如表7-2。");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.Alignment = ParagraphAlignment.CENTER;
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("表7-2濂泉绿色电缆线路的项数要求及情况");
                    index = doc.Tables.Count;
                    CopyTable(readDoc, 13, index, doc);//复制表

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("本工程符合三级标准。");
                    break;
                case "濂泉（沙河）站出线间隔":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据110kV濂泉（沙河）站 “电气总平面布置图”，本工程本期电缆线路采用电缆型式向东出线，出线间隔布置如下图所示。");

                    //插入图片函数
                    CopyPicture("濂泉（沙河）站出线间隔", doc,11.26f, 4.61f);
                    paragraph = doc.CreateParagraph();
                    paragraph.Alignment = ParagraphAlignment.CENTER;
                    run = paragraph.CreateRun();
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("图2.1.1 濂泉（沙河）站110kV出线间隔示意图");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("本工程本期110kV双回电缆线路使用 “永福甲”、“永福乙”间隔。");

                    break;
                case "永福站出线间隔":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("根据220kV永福站 “电气总平面布置图”， 本工程本期电缆线路采用电缆型式向东南出线，出线间隔布置如下图所示。");

                    //插入图片函数
                    CopyPicture("永福站出线间隔", doc, 16.81f, 2.81f);

                    paragraph = doc.CreateParagraph();
                    paragraph.Alignment = ParagraphAlignment.CENTER;
                    run = paragraph.CreateRun();
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("图2.1.2 永福站110kV出线间隔示意图");

                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("本工程本期110kV双回电缆线路使用“濂泉甲”、“濂泉乙”间隔。");

                    break;
                case "接地电流监测系统":
                    paragraph = doc.CreateParagraph();
                    run = paragraph.CreateRun();
                    paragraph.IndentationFirstLine = MainForm.Indentation("宋体", 14, 4, FontStyle.Regular);//首行缩进
                    run.FontFamily = "宋体";
                    run.FontSize = 14;
                    run.SetText("本工程新建电缆金属护套均采用交叉互联两端直接接地的接地方式，根据要求，在每个接头位置安装一个电流互感器监测电缆接地电流，通过变送器将监测到的信号传输至通信子站。具体如下图所示：");

                    //插入图片函数
                    CopyPicture("接地电流监测系统", doc, 14.64f, 7.57f);
                    break;
                default:
                    flag = true;
                    break;
            }
            return flag;
        
        
        }


        
        /// <summary>
        /// 合并行、垂直合并列单元格
        /// </summary>
        /// <param name="table"></param>
        /// <param name="fromRow"></param>
        /// <param name="toRow"></param>
        /// <param name="colIndex"></param>
        public void MYMergeRows(XWPFTable table, int fromRow, int toRow, int colIndex)
        {
            for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++)
            {
                XWPFTableCell rowcell = table.GetRow(rowIndex).GetCell(colIndex);
                rowcell.SetVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                CT_Tc cttc = rowcell.GetCTTc();
                CT_TcPr ctTcPr = cttc.tcPr;
                if (ctTcPr == null)
                {
                    ctTcPr = cttc.AddNewTcPr();
                }

                if (rowIndex == fromRow)
                {
                    // The first merged cell is set with RESTART merge value
                    ctTcPr.AddNewVMerge().val = ST_Merge.restart;
                }
                else
                {
                    // Cells which join (merge) the first one, are set with CONTINUE
                    ctTcPr.AddNewVMerge().val = ST_Merge.@continue;//继续合并行
                }
                ctTcPr.AddNewVAlign().val = ST_VerticalJc.center;//垂直
            }
        }


        private void readWord() 
        {

            using (FileStream stream=File.OpenRead("C:\\Users\\NBJZ\\Desktop\\施工说明书(模板).docx"))
            //using (FileStream stream=File.OpenRead("C:\\Users\\NBJZ\\Desktop\\施工说明书（组合）.docx"))
            {
                XWPFDocument doc = new XWPFDocument(stream);

                foreach (XWPFTable table in doc.Tables)
                {
                    //循环表格行
                    foreach (XWPFTableRow row in table.Rows)
                    {
                        foreach (XWPFTableCell cell in row.GetTableCells())
                        {
                            //sb.Append(cell.GetText());
                        }
                    }
                }


            }








                FileStream fs = new FileStream("C:\\Users\\NBJZ\\Desktop\\01 施工说明书(2).doc", FileMode.Open, FileAccess.Read);
            XWPFDocument myDocx = new XWPFDocument(fs);//打开07（.docx）以上的版本的文档
                                                       //读取表格
            foreach (XWPFTable table in myDocx.Tables)
            {
                //循环表格行
                foreach (XWPFTableRow row in table.Rows)
                {
                    foreach (XWPFTableCell cell in row.GetTableCells())
                    {
                        //sb.Append(cell.GetText());
                    }
                }
            }


        }

        /// <summary>
        /// 为XWPFDocument文档复制指定索引的表
        /// </summary>
        /// <param name="readDoc">模板文件</param>
        /// <param name="tableIndex">需要复制模板的table的索引</param>
        /// <param name="targetIndex">复制到目标位置的table索引(如果目标位置原来有表格，会被覆盖)</param>
        /// <param name="myDoc">新创建的文件</param>
        public static void CopyTable(XWPFDocument readDoc, int tableIndex, int targetIndex, XWPFDocument myDoc)
        {
            var sourceTable = readDoc.Tables[tableIndex];
            CT_Tbl sourceCTTbl = readDoc.Document.body.GetTblArray(8);

            var targetTable = myDoc.CreateTable();
            myDoc.SetTable(targetIndex, targetTable);
            var targetCTTbl = myDoc.Document.body.GetTblArray()[myDoc.Document.body.GetTblArray().Length-1 ];
            targetCTTbl.tblPr = sourceCTTbl.tblPr;
            targetCTTbl.tblPr.jc.val = ST_Jc.left;//表格在页面水平位置
            //targetCTTbl.tblGrid = sourceCTTbl.tblGrid;

            for (int i = 0; i < sourceTable.Rows.Count; i++)
            {
                var tbRow = targetTable.CreateRow();
                var targetRow = tbRow.GetCTRow();
                tbRow.RemoveCell(0);
                XWPFTableRow row = sourceTable.Rows[i];
                targetRow.trPr = row.GetCTRow().trPr;
                for (int c = 0; c < row.GetTableCells().Count; c++)
                {
                    var tbCell = tbRow.CreateCell();
                    tbCell.RemoveParagraph(0);
                    var targetCell = tbCell.GetCTTc();

                    XWPFTableCell cell = row.GetTableCells()[c];
                    targetCell.tcPr = cell.GetCTTc().tcPr;
                    for (int p = 0; p < cell.Paragraphs.Count; p++)
                    {
                        var tbPhs = tbCell.AddParagraph();
                        CT_P targetPhs = tbPhs.GetCTP();
                        XWPFParagraph para = cell.Paragraphs[p];
                        var paraCTP = para.GetCTP();
                        targetPhs.pPr = paraCTP.pPr;
                        targetPhs.rsidR = paraCTP.rsidR;
                        targetPhs.rsidRPr = paraCTP.rsidRPr;
                        targetPhs.rsidRDefault = paraCTP.rsidRDefault;
                        targetPhs.rsidP = paraCTP.rsidP;

                        for (int r = 0; r < para.Runs.Count; r++)
                        {
                            var tbRun = tbPhs.CreateRun();
                            CT_R targetRun = tbRun.GetCTR();

                            XWPFRun run = para.Runs[r];
                            var runCTR = run.GetCTR();
                            targetRun.rPr = runCTR.rPr;
                            targetRun.rsidRPr = runCTR.rsidRPr;
                            targetRun.rsidR = runCTR.rsidR;
                            CT_Text text = targetRun.AddNewT();
                            text.Value = run.Text;
                        }
                    }
                }
            }
            targetTable.RemoveRow(0);
        }

        /// <summary>
        /// 在word中插入文档
        /// </summary>
        /// <param name="picName">图片名称</param>
        /// <param name="myDoc">文档对象</param>
        /// <param name="w">图片宽度</param>
        /// <param name="h">图片高度</param>
        public void CopyPicture(string  picName, XWPFDocument myDoc,float w,float h)
        {
            //XWPFPictureData pic = myDoc.AllPictures[0];
            int width = (int)(w *38.665* 9325);
            int height = (int)(h * 38.67 * 9325);
            XWPFParagraph paragraph = myDoc.CreateParagraph();
            paragraph.Alignment = ParagraphAlignment.CENTER;
            XWPFRun run = paragraph.CreateRun();
            string path = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "图库\\"+picName+".jpg";
            using (FileStream stream=new FileStream(path,FileMode.Open,FileAccess.Read))
            {
                //图片的文件流，图片类型、图片名称  ，设置的宽度及高度
                run.AddPicture(stream, (int)PictureType.PNG, picName, width,height);
            }
        }

    }
}
