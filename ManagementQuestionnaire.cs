using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using K12.Questionnaire.UDT;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using DevComponents.DotNetBar;

namespace K12.Questionnaire
{
    public partial class ManagementQuestionnaire : BaseForm
    {

        List<QuestionnaireForm> formList = new List<QuestionnaireForm>();

        QuestionnaireForm currentQ = new QuestionnaireForm();

        public ManagementQuestionnaire()
        {
            InitializeComponent();

 


            ButtonItem btnItem_preload = new ButtonItem();

            btnItem_preload.Text = "資料讀取中...";

            btnItem_preload.ButtonStyle = eButtonStyle.ImageAndText;

            itmPnlQuestionnaire.Items.Add(btnItem_preload);



            
           
            var bkw = new BackgroundWorker();
            
            bkw.DoWork += delegate
            {

                //formList = new AccessHelper().Select<QuestionnaireForm>("ref_teacher_id = null AND NOT(end_time is null)");
                
                formList = new AccessHelper().Select<QuestionnaireForm>();

                formList.Sort(delegate(QuestionnaireForm f1, QuestionnaireForm f2)
                {
                    
                    //return f1.EndTime.Value.CompareTo(f2.EndTime.Value);

                    //不是每張Table 每個時候都有 StartTime、 EndTime  先用Name 排序
                    return f1.Name.CompareTo(f2.Name);

                });
            };

            bkw.RunWorkerCompleted += delegate
            {

                RefreshitmPnlQuestionnaire();


            };
            bkw.RunWorkerAsync();


        }

        void btnItem_Click(object sender, EventArgs e)
        {


            #region 儲存currentQ Memo 至物件 QuestionnaireForm
            //  儲存Memo
            XmlDocument doc = new XmlDocument();

            foreach (QuestionnaireForm q in formList)
            {
                if (q == currentQ && !q.Deleted)
                {
                    if (q.ContentString != null && q.ContentString != "")
                    {
                        doc.LoadXml(q.ContentString);

                        // 問卷說明備註                  
                        doc.SelectSingleNode("Content").SelectSingleNode("Memo").InnerText = MemoTxt.Text;





                        


                        q.ContentString = doc.OuterXml;
                    }
                }
            }  
            #endregion







            dataGridViewX1.Rows.Clear();

            //若左方 問卷只選一個選項
            if (itmPnlQuestionnaire.SelectedItems.Count == 1)
            {
                //取得目前所選擇的Button
                ButtonItem Buttonitem = itmPnlQuestionnaire.SelectedItems[0] as ButtonItem;
               
                //取得問卷Record
                QuestionnaireForm Q = (QuestionnaireForm)Buttonitem.Tag;

                currentQ = Q;

                XmlDocument content = new XmlDocument();


                if (Q.ContentString != null && Q.ContentString != "") 
                {

                    content.LoadXml(Q.ContentString);

                    if (content.SelectSingleNode("Content") != null ) 
                    {
                        if (content.SelectSingleNode("Content").SelectNodes("Section") != null) 
                        {
                            foreach (XmlElement section in content.SelectSingleNode("Content").SelectNodes("Section"))
                            {

                                foreach (XmlElement question in section.SelectNodes("Question"))
                                {
                                    DataGridViewRow dr = new DataGridViewRow();

                                    dr.CreateCells(dataGridViewX1);

                                    //群組
                                    dr.Cells[0].Value = section.SelectSingleNode("Title").InnerText;

                                    //問題
                                    dr.Cells[1].Value = question.SelectSingleNode("Title").InnerText;

                                    //題型
                                    dr.Cells[2].Value = question.SelectSingleNode("Type").InnerText == "option" ? "選項" : "填答";

                                    //必填
                                    dr.Cells[3].Value = question.SelectSingleNode("Require").InnerText == "true" ? "是" : "否";

                                    //最少勾選
                                    dr.Cells[4].Value = question.SelectSingleNode("Max") != null ? question.SelectSingleNode("Max").InnerText : "";

                                    //最多勾選
                                    dr.Cells[5].Value = question.SelectSingleNode("Min") != null ? question.SelectSingleNode("Min").InnerText : "";



                                    List<string> option_list = new List<string>();

                                    foreach (XmlElement option in question.SelectNodes("Option"))
                                    {

                                        option_list.Add(option.InnerText);

                                    }

                                    DataGridViewComboBoxCell c = new DataGridViewComboBoxCell();

                                    c.DataSource = option_list;

                                    //選項
                                    dr.Cells[6] = c;

                                    dataGridViewX1.Rows.Add(dr);
                                }
                            }                                                                        
                        }

                        if (content.SelectSingleNode("Content").SelectSingleNode("Memo") != null) 
                        {
                            // 問卷說明備註
                            MemoTxt.Text = content.SelectSingleNode("Content").SelectSingleNode("Memo").InnerText;                        
                        }                        
                    }                                        
                }
                // 問卷名稱
                NameTxt.Text = Q.Name;

                
                
            }
        }


        // Save
        private void buttonX1_Click(object sender, EventArgs e)
        {

            #region 儲存currentQ Memo 至物件 QuestionnaireForm
            //  儲存Memo
            XmlDocument doc = new XmlDocument();

            foreach (QuestionnaireForm q in formList)
            {
                if (q == currentQ && !q.Deleted)
                {
                    if (q.ContentString != null && q.ContentString != "")
                    {
                        doc.LoadXml(q.ContentString);

                        // 問卷說明備註                  
                        doc.SelectSingleNode("Content").SelectSingleNode("Memo").InnerText = MemoTxt.Text;

                        q.ContentString = doc.OuterXml;
                    }
                }
            }
            #endregion



            // 檢查 是否 有重覆的問卷名稱，基本上不希望有

            List<string> questionnaire_name_list = new List<string>();
            foreach (QuestionnaireForm q in formList)
            {
                if (!questionnaire_name_list.Contains(q.Name))
                {
                    questionnaire_name_list.Add(q.Name);
                }
                else 
                {
                    

                    MsgBox.Show("問卷:" + q.Name + "名稱重覆，請修正。", "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }
            }



            formList.SaveAll();

            MsgBox.Show("儲存成功");
            
            //this.Close();
        }


        // Close
        private void buttonX2_Click(object sender, EventArgs e)
        {

            //MsgBox.Show("有尚未儲存的問卷更動，確定要離開?", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            this.Close();

        }

        //新增問卷
        private void buttonX3_Click(object sender, EventArgs e)
        {

            QuestionnaireForm Q = new QuestionnaireForm();

            Q.Name = "新增問卷";

            Q.ParentReply = true;

            Q.StudentReply = true;

            Q.TeacherReply = true;

            
            //拚content xml

            XmlDocument doc = new XmlDocument();

            XmlElement content = doc.CreateElement("Content");

            XmlElement memo = doc.CreateElement("Memo");

            content.AppendChild(memo);

            doc.AppendChild(content);




            Q.ContentString = doc.OuterXml;

            formList.Add(Q);
            
            RefreshitmPnlQuestionnaire();

        }

        // 刪除問卷
        private void buttonX4_Click(object sender, EventArgs e)
        {                        
            //若左方 問卷只選一個選項
            if (itmPnlQuestionnaire.SelectedItems.Count == 1)
            {
                //取得目前所選擇的Button
                ButtonItem Buttonitem = itmPnlQuestionnaire.SelectedItems[0] as ButtonItem;

                //取得問卷Record
                QuestionnaireForm Q = (QuestionnaireForm)Buttonitem.Tag;


                //formList.Remove(Q);


                foreach (QuestionnaireForm q in formList)
                {
                    if (q == Q)
                    {
                        q.Deleted = true;
                    }
                }

                RefreshitmPnlQuestionnaire();

                // 將右側UI 清空
                NameTxt.Text = "";
                MemoTxt.Text = "";
                dataGridViewX1.Rows.Clear();

            }
            else 
            {
                MsgBox.Show("請選擇一個欲刪除課程問卷。");
                        
            }            
        }


        // 更新左邊問卷清單
        private void RefreshitmPnlQuestionnaire() 
        {

            #region 問卷

            itmPnlQuestionnaire.Items.Clear();

            foreach (QuestionnaireForm Q in formList)
            {
                //假如問卷不是有要被刪除，就列出來
                if (!Q.Deleted) 
                {
                    ButtonItem btnItem = new ButtonItem();
                    btnItem.Text = Q.Name;
                    btnItem.Tag = Q;
                    btnItem.OptionGroup = "itmPnlTimeName";
                    btnItem.ButtonStyle = eButtonStyle.ImageAndText;

                    btnItem.Click += new EventHandler(btnItem_Click);

                    itmPnlQuestionnaire.Items.Add(btnItem);                                                
                }                
            }

            itmPnlQuestionnaire.ResumeLayout();
            itmPnlQuestionnaire.Refresh();
            #endregion
        
        
        }


        // 問卷 Name 輸入 >>為即時更新
        private void NameTxt_TextChanged(object sender, EventArgs e)
        {
                                           
            foreach (QuestionnaireForm q in formList)             
            {
                if (q == currentQ && !q.Deleted )                
                {
                 
                    q.Name = NameTxt.Text;                
                }              
            }             
            RefreshitmPnlQuestionnaire();            
        }

        
       
    }
}
