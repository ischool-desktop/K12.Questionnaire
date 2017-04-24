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

        // 目前選擇的問卷
        QuestionnaireForm currentQ = new QuestionnaireForm();

        // 紀錄 問卷是否有改變
        bool questionnaire_modified = false;
        // 紀錄 問卷_name是否有改變
        bool questionnaire_name_modified = false;
        // 紀錄 問卷_memo是否有改變
        bool questionnaire_memo_modified = false;
        // 紀錄 問卷_start_time是否有改變
        bool questionnaire_start_time_modified = false;
        // 紀錄 問卷_end_time是否有改變
        bool questionnaire_end_time_modified = false;

        public ManagementQuestionnaire()
        {
            InitializeComponent();

            //初始化左方 問卷List
            ButtonItem btnItem_preload = new ButtonItem();
            btnItem_preload.Text = "資料讀取中...";
            btnItem_preload.ButtonStyle = eButtonStyle.ImageAndText;
            itmPnlQuestionnaire.Items.Add(btnItem_preload);

            //題型--選項
            Column3.Items.Add("選擇");
            Column3.Items.Add("填答");

            //必選--選項
            Column4.Items.Add("是");
            Column4.Items.Add("否");

            var bkw = new BackgroundWorker();

            bkw.DoWork += delegate
            {
                //formList = new AccessHelper().Select<QuestionnaireForm>("ref_teacher_id = null AND NOT(end_time is null)");

                formList = new AccessHelper().Select<QuestionnaireForm>();

                formList.Sort(delegate(QuestionnaireForm f1, QuestionnaireForm f2)
                {
                    //不是每張Table 每個時候都有 StartTime、 EndTime  先用Name 排序
                    return f1.Name.CompareTo(f2.Name);
                });
            };

            bkw.RunWorkerCompleted += delegate
            {
                // 更新左方 問卷List
                RefreshitmPnlQuestionnaire();

                RenewRightUI();
            };
            bkw.RunWorkerAsync();

        }

        //點選左邊問卷List 項目
        void btnItem_Click(object sender, EventArgs e)
        {
            // 更新右方UI
            RenewRightUI();
        }

        // Save
        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 資料合理性檢查
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

            DateTime start_time;
            DateTime end_time;

            // 開始時間 檢查
            if (!DateTime.TryParse(StartTimetxt.Text, out start_time))
            {
                MsgBox.Show("開始時間格式輸入錯誤，請修正。", "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
            }

            // 結束時間 檢查
            if (!DateTime.TryParse(EndTimetxt.Text, out end_time))
            {
                MsgBox.Show("結束時間格式輸入錯誤，請修正。", "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
            }

            // 開始時間與結束時間關係 檢查
            if (start_time > end_time)
            {
                MsgBox.Show("開始時間必須早於結束時間，請修正。", "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
            }

            // 傳遞目前右邊UI畫面的輸入值 至 QuestionnaireForm 物件                                           
            SectionDeliver();

            formList.SaveAll();

            MsgBox.Show("儲存成功");

            questionnaire_modified = false;

            Modified_Indicator.Text = "已儲存";

            //this.Close();
        } 
            #endregion

        //新增問卷
        private void buttonX3_Click(object sender, EventArgs e)
        {

            QuestionnaireForm Q = new QuestionnaireForm();

            Q.Name = "新增問卷";

            Q.ParentReply = true;
            Q.StudentReply = true;
            Q.TeacherReply = true;

            #region 拚基礎的 content xml
            //拚content xml

            XmlDocument doc = new XmlDocument();

            XmlElement content = doc.CreateElement("Content");

            XmlElement memo = doc.CreateElement("Memo");

            memo.InnerText = "";

            content.AppendChild(memo);

            //XmlElement section = doc.CreateElement("Section");

            //XmlElement section_title = doc.CreateElement("Title");

            //XmlElement question = doc.CreateElement("Question");

            //XmlElement question_title = doc.CreateElement("Title");

            //XmlElement type = doc.CreateElement("Type");

            //XmlElement require = doc.CreateElement("Require");

            //XmlElement max = doc.CreateElement("Max");

            //XmlElement min = doc.CreateElement("Min");

            //XmlElement option = doc.CreateElement("Option");

            //question.AppendChild(question_title);
            //question.AppendChild(type);
            //question.AppendChild(require);
            //question.AppendChild(max);
            //question.AppendChild(min);
            //question.AppendChild(option);

            //section.AppendChild(section_title);
            //section.AppendChild(question);


            //content.AppendChild(section);

            doc.AppendChild(content);
            #endregion


            Q.ContentString = doc.OuterXml;

            // 整理目前有的問卷名稱， 避免 新增出現同名問卷
            List<string> q_name_ori_list = new List<string>();

            foreach (QuestionnaireForm q in formList)
            {
                q_name_ori_list.Add(q.Name);
            }

            InsertQuestionnaireInputNewName IQINN = new InsertQuestionnaireInputNewName(q_name_ori_list);

            IQINN.ShowDialog();

            if (IQINN.DialogResult == DialogResult.OK)
            {

                Q.Name = IQINN.q_name;

                formList.Add(Q);

                formList.SaveAll();

                RefreshitmPnlQuestionnaire(Q);

                RenewRightUI();
            }
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

                if (MsgBox.Show("點選確認將會刪除本問卷，且不能回復，確認刪除?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    foreach (QuestionnaireForm q in formList)
                    {
                        if (q == Q)
                        {
                            q.Deleted = true;
                        }
                    }

                    formList.SaveAll();

                    //刪除後 ，取得新的formList，可以避免ㄧ些新舊資料交雜的錯誤
                    formList = new AccessHelper().Select<QuestionnaireForm>();

                    formList.Sort(delegate(QuestionnaireForm f1, QuestionnaireForm f2)
                    {                        
                        //不是每張Table 每個時候都有 StartTime、 EndTime  先用Name 排序
                        return f1.Name.CompareTo(f2.Name);
                    });

                    RefreshitmPnlQuestionnaire();



                    // 將右側UI 清空
                    //NameTxt.Text = "";
                    //MemoTxt.Text = "";
                    //dataGridViewX1.Rows.Clear();

                    RenewRightUI();
                }
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

                    // 預設 先選List 第一個項目 為"選擇狀態"
                    if(Q == formList[0])
                    {
                        btnItem.Checked = true;                                       
                    }

                    itmPnlQuestionnaire.Items.Add(btnItem);
                }
            }

         
            itmPnlQuestionnaire.ResumeLayout();
            itmPnlQuestionnaire.Refresh();
            #endregion
        }

        // 更新左邊問卷清單，並指定選項為選擇狀態  (此為上面 方法 RefreshitmPnlQuestionnaire() 的多載，主要供 新增問卷、複製問卷 按鈕使用)
        private void RefreshitmPnlQuestionnaire(QuestionnaireForm select_q)
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

                    if (Q == select_q || Q.Name == select_q.Name)
                    {
                        btnItem.Checked = true;
                    }

                    btnItem.Click += new EventHandler(btnItem_Click);

                    itmPnlQuestionnaire.Items.Add(btnItem);
                }
            }

            itmPnlQuestionnaire.ResumeLayout();
            itmPnlQuestionnaire.Refresh();
            #endregion
        }

        //  傳遞Section (其實後來也不只傳遞Section 資料了，整張QuestionnaireForm表 都在此傳遞)
        private void SectionDeliver()
        {
            XmlDocument doc = new XmlDocument();

            Dictionary<string, XmlElement> section_title_dict = new Dictionary<string, XmlElement>();

            #region 整理 secion 的資料
            foreach (QuestionnaireForm q in formList)
            {
                if (q == currentQ && !q.Deleted)
                {
                    foreach (DataGridViewRow dr in dataGridViewX1.Rows)
                    {
                        //非最後一編輯列 才納入
                        if (!dr.IsNewRow)
                        {
                            if (!section_title_dict.ContainsKey("" + dr.Cells[0].Value))
                            {
                                XmlElement section = doc.CreateElement("Section");

                                XmlElement section_title = doc.CreateElement("Title");

                                section_title.InnerText = "" + dr.Cells[0].Value;

                                XmlElement question = doc.CreateElement("Question");

                                XmlElement question_title = doc.CreateElement("Title");

                                question_title.InnerText = "" + dr.Cells[1].Value;

                                XmlElement type = doc.CreateElement("Type");

                                type.InnerText = ("" + dr.Cells[2].Value == "選擇" ? "option" : "text");

                                XmlElement require = doc.CreateElement("Require");

                                require.InnerText = ("" + dr.Cells[3].Value == "是") ? "true" : "false";

                                XmlElement min = doc.CreateElement("Min");

                                min.InnerText = "" + dr.Cells[4].Value;

                                if ("" + dr.Cells[4].Value == "" && "" + dr.Cells[3].Value == "是")
                                {
                                    min.InnerText = "1";                                                                
                                }
                                if ("" + dr.Cells[4].Value == "" && "" + dr.Cells[3].Value != "是")
                                {
                                    min.InnerText = "0";
                                }

                                XmlElement max = doc.CreateElement("Max");

                                max.InnerText = "" + dr.Cells[5].Value;
                                
                                question.AppendChild(question_title);
                                question.AppendChild(type);
                                question.AppendChild(require);
                                question.AppendChild(max);
                                question.AppendChild(min);

                                //dr.Tag != null 意旨為 此行 為選擇題型資料 有option 要儲存
                                if (dr.Tag != null)
                                {
                                    foreach (string opt in (List<string>)dr.Tag)
                                    {
                                        XmlElement option = doc.CreateElement("Option");

                                        option.InnerText = opt;

                                        question.AppendChild(option);

                                    }
                                }

                                section.AppendChild(section_title);
                                section.AppendChild(question);

                                section_title_dict.Add("" + section_title.InnerText, section);

                            }
                            else
                            {
                                XmlElement question = doc.CreateElement("Question");

                                XmlElement question_title = doc.CreateElement("Title");

                                question_title.InnerText = "" + dr.Cells[1].Value;

                                XmlElement type = doc.CreateElement("Type");

                                type.InnerText = ("" + dr.Cells[2].Value == "選擇" ? "option" : "text");

                                XmlElement require = doc.CreateElement("Require");

                                require.InnerText = ("" + dr.Cells[3].Value == "是") ? "true" : "false";

                                XmlElement min = doc.CreateElement("Min");

                                min.InnerText = "" + dr.Cells[4].Value;

                                if ("" + dr.Cells[4].Value == "" && "" + dr.Cells[3].Value == "是")
                                {
                                    min.InnerText = "1";
                                }
                                if ("" + dr.Cells[4].Value == "" && "" + dr.Cells[3].Value != "是")
                                {
                                    min.InnerText = "0";
                                }
                                
                                XmlElement max = doc.CreateElement("Max");

                                max.InnerText = "" + dr.Cells[5].Value;
                                
                                question.AppendChild(question_title);
                                question.AppendChild(type);
                                question.AppendChild(require);
                                question.AppendChild(max);
                                question.AppendChild(min);

                                if (dr.Tag != null) 
                                {
                                    foreach (string opt in (List<string>)dr.Tag)
                                    {
                                        XmlElement option = doc.CreateElement("Option");

                                        option.InnerText = opt;

                                        question.AppendChild(option);
                                    }                                
                                }                                
                                section_title_dict["" + dr.Cells[0].Value].AppendChild(question);
                            }
                        }
                    }
                }
            }
            #endregion

            #region 傳遞currentQ xml資訊 至物件 QuestionnaireForm

            foreach (QuestionnaireForm q in formList)
            {
                if (q == currentQ && !q.Deleted)
                {
                    if (q.ContentString != null && q.ContentString != "")
                    {
                        doc.LoadXml(q.ContentString);

                        // 問卷說明備註                  
                        doc.SelectSingleNode("Content").SelectSingleNode("Memo").InnerText = MemoTxt.Text;

                        //把舊的 xml Section  都先刪除，以利下方避免重覆加入
                        foreach (XmlElement section in doc.SelectSingleNode("Content").SelectNodes("Section"))
                        {
                            doc.SelectSingleNode("Content").RemoveChild(section);
                        }

                        // 加入問卷 群組section 部分(裡面包含了Question)
                        foreach (XmlElement section in section_title_dict.Values)
                        {
                            doc.SelectSingleNode("Content").AppendChild(section);
                        }

                        q.ContentString = doc.OuterXml;
                    }
          
                    q.StartTime = DateTime.Parse(StartTimetxt.Text);
                    q.EndTime = DateTime.Parse(EndTimetxt.Text);

                    q.Name = NameTxt.Text;
                }
            }
            #endregion
        }

        private void dataGridViewX1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //不是表頭， 處理 當題型為"選擇" 時 才可以填答 最多、最少勾選數、選項，避免使用者存到髒資料
            if (e.RowIndex != -1 && e.ColumnIndex == 2)
            {
                foreach (DataGridViewRow dr in dataGridViewX1.Rows)
                {
                    if (dr.Index == e.RowIndex)
                    {
                        if (("" + dr.Cells[e.ColumnIndex].Value) == "選擇")
                        {
                            dr.Cells[6].Value = "點選以編輯";

                            dr.Cells[4].ReadOnly = false;
                            dr.Cells[5].ReadOnly = false;
                            dr.Cells[6].ReadOnly = false;

                            //將背景改回白
                            //dr.Cells[4].Style.BackColor = Color.White;
                            //dr.Cells[5].Style.BackColor = Color.White;
                            //dr.Cells[6].Style.BackColor = Color.White;                                                                              
                        }
                        else  //填答
                        {
                            //清空原本的值
                            dr.Cells[4].Value = "";
                            dr.Cells[5].Value = "";
                            dr.Cells[6].Value = "";

                            dr.Cells[4].ReadOnly = true;
                            dr.Cells[5].ReadOnly = true;
                            dr.Cells[6].ReadOnly = true;

                            //將背景改灰
                            //dr.Cells[4].Style.BackColor = Color.Gray;
                            //dr.Cells[5].Style.BackColor = Color.Gray;
                            //dr.Cells[6].Style.BackColor = Color.Gray;                                                                        
                        }
                    }
                }
            }
            dataGridViewX1_xml_compare_to_currentQ();
        }

        private void dataGridViewX1_xml_compare_to_currentQ()
        {
            #region 使用現在的Datagridview 產生 xml 和 currentQ  做比較

            if (currentQ.ContentString != null)
            {
                XmlDocument doc = new XmlDocument();

                Dictionary<string, XmlElement> section_title_dict = new Dictionary<string, XmlElement>();

                foreach (DataGridViewRow dr in dataGridViewX1.Rows)
                {
                    //非最後一編輯列 才納入
                    if (!dr.IsNewRow)
                    {
                        if (!section_title_dict.ContainsKey("" + dr.Cells[0].Value))
                        {

                            XmlElement section = doc.CreateElement("Section");

                            XmlElement section_title = doc.CreateElement("Title");

                            section_title.InnerText = "" + dr.Cells[0].Value;

                            XmlElement question = doc.CreateElement("Question");

                            XmlElement question_title = doc.CreateElement("Title");

                            question_title.InnerText = "" + dr.Cells[1].Value;

                            XmlElement type = doc.CreateElement("Type");

                            type.InnerText = ("" + dr.Cells[2].Value == "選擇" ? "option" : "text");

                            XmlElement require = doc.CreateElement("Require");

                            require.InnerText = ("" + dr.Cells[3].Value == "是") ? "true" : "false";


                            XmlElement min = doc.CreateElement("Min");

                            min.InnerText = "" + dr.Cells[4].Value;


                            if ("" + dr.Cells[4].Value == "" && "" + dr.Cells[3].Value == "是")
                            {
                                min.InnerText = "1";
                            }
                            if ("" + dr.Cells[4].Value == "" && "" + dr.Cells[3].Value != "是")
                            {
                                min.InnerText = "0";
                            }

                            XmlElement max = doc.CreateElement("Max");

                            max.InnerText = "" + dr.Cells[5].Value;
                            
                            question.AppendChild(question_title);
                            question.AppendChild(type);
                            question.AppendChild(require);
                            question.AppendChild(max);
                            question.AppendChild(min);

                            //dr.Tag != null 意旨為 此行 為選擇題型資料 有option 要儲存
                            if (dr.Tag != null)
                            {
                                foreach (string opt in (List<string>)dr.Tag)
                                {
                                    XmlElement option = doc.CreateElement("Option");

                                    option.InnerText = opt;

                                    question.AppendChild(option);
                                }
                            }

                            section.AppendChild(section_title);
                            section.AppendChild(question);

                            section_title_dict.Add("" + section_title.InnerText, section);
                        }
                        else
                        {
                            XmlElement question = doc.CreateElement("Question");

                            XmlElement question_title = doc.CreateElement("Title");

                            question_title.InnerText = "" + dr.Cells[1].Value;

                            XmlElement type = doc.CreateElement("Type");

                            type.InnerText = ("" + dr.Cells[2].Value == "選擇" ? "option" : "text");

                            XmlElement require = doc.CreateElement("Require");

                            require.InnerText = ("" + dr.Cells[3].Value == "是") ? "true" : "false";
                            
                            XmlElement min = doc.CreateElement("Min");

                            min.InnerText = "" + dr.Cells[4].Value;

                            if ("" + dr.Cells[4].Value == "" && "" + dr.Cells[3].Value == "是")
                            {
                                min.InnerText = "1";
                            }
                            if ("" + dr.Cells[4].Value == "" && "" + dr.Cells[3].Value != "是")
                            {
                                min.InnerText = "0";
                            }

                            XmlElement max = doc.CreateElement("Max");

                            max.InnerText = "" + dr.Cells[5].Value;
                            
                            question.AppendChild(question_title);
                            question.AppendChild(type);
                            question.AppendChild(require);
                            question.AppendChild(max);
                            question.AppendChild(min);

                            if (dr.Tag != null) 
                            {
                                foreach (string opt in (List<string>)dr.Tag)
                                {
                                    XmlElement option = doc.CreateElement("Option");

                                    option.InnerText = opt;

                                    question.AppendChild(option);

                                }                                                        
                            }
                            
                            section_title_dict["" + dr.Cells[0].Value].AppendChild(question);
                        }
                    }
                }

                XmlElement content = doc.CreateElement("Content");

                XmlElement memo = doc.CreateElement("Memo");

                memo.InnerText = MemoTxt.Text;

                content.AppendChild(memo);

                doc.AppendChild(content);

                // 加入問卷 群組section 部分(裡面包含了Question)
                foreach (XmlElement section in section_title_dict.Values)
                {
                    doc.SelectSingleNode("Content").AppendChild(section);
                }

                XmlDocument doc_current = new XmlDocument();

                doc_current.LoadXml(currentQ.ContentString);

                if (doc_current.SelectSingleNode("Content").OuterXml != doc.OuterXml)
                {
                    questionnaire_modified = true;
                    Modified_Indicator.Text = "已修改，尚未儲存。";
                    Modified_Indicator.Visible = true;

                }
                else
                {
                    questionnaire_modified = false;
                    Modified_Indicator.Text = "";
                    Modified_Indicator.Visible = false;
                }
            }
            #endregion
        }
   
        //複製問卷
        private void buttonX5_Click(object sender, EventArgs e)
        {

            //若左方 問卷只選一個選項
            if (itmPnlQuestionnaire.SelectedItems.Count == 1)
            {
                //取得目前所選擇的Button
                ButtonItem btnItem = itmPnlQuestionnaire.SelectedItems[0] as ButtonItem;

                // 複製出來的 問卷q
                QuestionnaireForm q_copy = new QuestionnaireForm();

                //取得問卷Record
                QuestionnaireForm q = (QuestionnaireForm)btnItem.Tag;

                //  輸入新問卷Name Form
                CopyQuestionnaireInputNewName CQINN = new CopyQuestionnaireInputNewName(q.Name);

                CQINN.ShowDialog();

                if (CQINN.DialogResult == DialogResult.OK)
                {
                    q_copy.Name = CQINN.q_name;

                    q_copy.ContentString = q.ContentString;

                    q_copy.StudentReply = q.StudentReply;
                    q_copy.ParentReply = q.ParentReply;
                    q_copy.TeacherReply = q.TeacherReply;

                    q_copy.StartTime = q.StartTime;
                    q_copy.EndTime = q.EndTime;

                    //加入 formList 儲存
                    formList.Add(q_copy);
                    formList.SaveAll();

                    formList = new AccessHelper().Select<QuestionnaireForm>();

                    formList.Sort(delegate(QuestionnaireForm f1, QuestionnaireForm f2)
                    {                      
                        //不是每張Table 每個時候都有 StartTime、 EndTime  先用Name 排序
                        return f1.Name.CompareTo(f2.Name);
                    });

                    q_copy = new AccessHelper().Select<QuestionnaireForm>("name =" + "'" + CQINN.q_name + "'")[0];

                    // 重心整理，並指定q_copy 為選擇狀態
                    RefreshitmPnlQuestionnaire(q_copy);

                    RenewRightUI();
                }
            }
        }

        // 當使用者 正在關閉表格時 
        private void ManagementQuestionnaire_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (questionnaire_modified)
            {
                if (MsgBox.Show("有尚未儲存的問卷更動，確定要離開?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    e.Cancel = false;
                }
                else
                {
                    //阻止離開程序
                    e.Cancel = true;
                }
            }
        }

        // 當 TextBox 文字改變時 ，
        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            // 與currentQ 的值比較，檢查是否有更動
            // 若NameTxt 有更動
            if (sender == NameTxt)
            {
                if (currentQ.Name != NameTxt.Text)
                {
                    questionnaire_modified = true;
                    Modified_Indicator.Text = "已修改，尚未儲存。";
                    Modified_Indicator.Visible = true;

                    questionnaire_name_modified = true;
                }
                else
                {
                    questionnaire_name_modified = false;

                    if (!questionnaire_name_modified && !questionnaire_memo_modified && !questionnaire_start_time_modified && !questionnaire_end_time_modified)
                    {
                        questionnaire_modified = false;
                        Modified_Indicator.Text = "";
                        Modified_Indicator.Visible = false;
                    }
                }
            }
            // 若MemoTxt 有更動
            if (sender == MemoTxt)
            {
                XmlDocument doc = new XmlDocument();

                doc.LoadXml(currentQ.ContentString);

                if (doc.SelectSingleNode("Content").SelectSingleNode("Memo") != null)
                {
                    if (doc.SelectSingleNode("Content").SelectSingleNode("Memo").InnerText != MemoTxt.Text)
                    {
                        questionnaire_modified = true;
                        Modified_Indicator.Text = "已修改，尚未儲存。";
                        Modified_Indicator.Visible = true;

                        questionnaire_memo_modified = true;
                    }
                    else
                    {
                        questionnaire_memo_modified = false;

                        if (!questionnaire_name_modified && !questionnaire_memo_modified && !questionnaire_start_time_modified && !questionnaire_end_time_modified)
                        {
                            questionnaire_modified = false;
                            Modified_Indicator.Text = "";
                            Modified_Indicator.Visible = false;

                        }
                    }
                }
            }
            // 若StartTimetxt 有更動
            if (sender == StartTimetxt)
            {
                if (currentQ.StartTime != null) 
                {
                    if (((DateTime)currentQ.StartTime).ToString("yyyy/MM/dd HH:mm:ss") != StartTimetxt.Text)
                    {
                        questionnaire_modified = true;
                        Modified_Indicator.Text = "已修改，尚未儲存。";
                        Modified_Indicator.Visible = true;

                        questionnaire_start_time_modified = true;
                    }
                    else
                    {
                        questionnaire_start_time_modified = false;

                        if (!questionnaire_name_modified && !questionnaire_memo_modified && !questionnaire_start_time_modified && !questionnaire_end_time_modified)
                        {
                            questionnaire_modified = false;
                            Modified_Indicator.Text = "";
                            Modified_Indicator.Visible = false;
                        }
                    }                                                
                }                
            }
            // 若EndTimetxt 有更動
            if (sender == EndTimetxt)
            {
                if (currentQ.EndTime != null) 
                {
                    if (((DateTime)currentQ.EndTime).ToString("yyyy/MM/dd HH:mm:ss") != EndTimetxt.Text)
                    {
                        questionnaire_modified = true;
                        Modified_Indicator.Text = "已修改，尚未儲存。";
                        Modified_Indicator.Visible = true;

                        questionnaire_end_time_modified = true;
                    }
                    else
                    {
                        questionnaire_end_time_modified = false;

                        if (!questionnaire_name_modified && !questionnaire_memo_modified && !questionnaire_start_time_modified && !questionnaire_end_time_modified)
                        {
                            questionnaire_modified = false;
                            Modified_Indicator.Text = "";
                            Modified_Indicator.Visible = false;
                        }
                    }                                
                }                
            }
        }

        // 如果有整行Row 被刪除，也要檢查 資料是否有更動
        private void dataGridViewX1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            dataGridViewX1_xml_compare_to_currentQ();
        }
        
        private void dataGridViewX1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 處理 點選左邊表頭 可以選 整ROW ，點Cell 可以直接進入編輯模式 的功能
            if (e.ColumnIndex != -1)
            {
                dataGridViewX1.BeginEdit(true);

                dataGridViewX1.EditMode = DataGridViewEditMode.EditOnEnter;
            }
            else
            {
                //dataGridViewX1.BeginEdit(false);
                dataGridViewX1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            }

            //處理 "點選以編輯" 選項， 開新的 Form 讓使用者調整 選項option
            if (e.RowIndex != -1 && (e.ColumnIndex == 6) && ("" + dataGridViewX1.Rows[e.RowIndex].Cells[2].Value == "選擇"))
            {

                QuestionnaireOptionSetting QOS = new QuestionnaireOptionSetting((List<string>)dataGridViewX1.Rows[e.RowIndex].Tag);

                QOS.ShowDialog();

                if (QOS.DialogResult == DialogResult.OK)
                {

                    dataGridViewX1.Rows[e.RowIndex].Tag = QOS.opt_list;

                    //QOS.Close();
                }
                dataGridViewX1_xml_compare_to_currentQ();
            }
        }


        private void RenewRightUI()
        {
            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (questionnaire_modified)
            {
                if (MsgBox.Show("有尚未儲存的問卷更動，確定要離開?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {

                }
                else
                {
                    return;
                }
            }

            Modified_Indicator.Text = "";
            Modified_Indicator.Visible = false;

            //SectionDeliver();

            //NameTxt.Text = "";
            //MemoTxt.Text = "";

            StartTimetxt.Text = "";
            EndTimetxt.Text = "";

            dataGridViewX1.Rows.Clear();

            //若左方 問卷只選一個選項
            if (itmPnlQuestionnaire.SelectedItems.Count == 1)
            {
                //取得目前所選擇的Button
                ButtonItem Buttonitem = itmPnlQuestionnaire.SelectedItems[0] as ButtonItem;

                //取得問卷Record
                QuestionnaireForm Q = (QuestionnaireForm)Buttonitem.Tag;

                //用來存放目前有的群組標題
                List<string> title_list = new List<string>();

                currentQ = Q;

                XmlDocument content = new XmlDocument();


                if (Q.ContentString != null && Q.ContentString != "")
                {
                    content.LoadXml(Q.ContentString);

                    if (content.SelectSingleNode("Content") != null)
                    {
                        if (content.SelectSingleNode("Content").SelectNodes("Section") != null)
                        {
                            foreach (XmlElement section in content.SelectSingleNode("Content").SelectNodes("Section"))
                            {
                                if (!title_list.Contains(section.SelectSingleNode("Title").InnerText))
                                {
                                    title_list.Add(section.SelectSingleNode("Title").InnerText);
                                }

                                foreach (XmlElement question in section.SelectNodes("Question"))
                                {
                                    DataGridViewRow dr = new DataGridViewRow();

                                    dr.CreateCells(dataGridViewX1);

                                    //群組
                                    dr.Cells[0].Value = section.SelectSingleNode("Title").InnerText;

                                    //問題
                                    dr.Cells[1].Value = question.SelectSingleNode("Title").InnerText;

                                    //題型
                                    dr.Cells[2].Value = question.SelectSingleNode("Type").InnerText == "option" ? "選擇" : "填答";

                                    //必填
                                    dr.Cells[3].Value = question.SelectSingleNode("Require").InnerText == "true" ? "是" : "否";


                                    if (question.SelectSingleNode("Type").InnerText == "option")
                                    {
                                        //最少勾選
                                        dr.Cells[4].Value = question.SelectSingleNode("Min") != null ? question.SelectSingleNode("Min").InnerText : "";

                                        //最多勾選
                                        dr.Cells[5].Value = question.SelectSingleNode("Max") != null ? question.SelectSingleNode("Max").InnerText : "";

                                        dr.Cells[4].ReadOnly = false;
                                        dr.Cells[5].ReadOnly = false;
                                        dr.Cells[6].ReadOnly = false;

                                        ////將背景改回白
                                        //dr.Cells[4].Style.BackColor = Color.White;
                                        //dr.Cells[5].Style.BackColor = Color.White;
                                        //dr.Cells[6].Style.BackColor = Color.White;

                                        dr.Cells[6].Value = "點選以編輯";
                                    }
                                    else
                                    {
                                        //清空原本的值
                                        dr.Cells[4].Value = "";
                                        dr.Cells[5].Value = "";
                                        dr.Cells[6].Value = "";

                                        dr.Cells[4].ReadOnly = true;
                                        dr.Cells[5].ReadOnly = true;
                                        dr.Cells[6].ReadOnly = true;

                                        ////將背景改灰
                                        //dr.Cells[4].Style.BackColor = Color.Gray;
                                        //dr.Cells[5].Style.BackColor = Color.Gray;
                                        //dr.Cells[6].Style.BackColor = Color.Gray;
                                    }

                                    List<string> option_list = new List<string>();

                                    foreach (XmlElement option in question.SelectNodes("Option"))
                                    {
                                        option_list.Add(option.InnerText);
                                    }

                                    // 利用Tag 來傳遞 每一行問題的opt
                                    dr.Tag = option_list;
                                    
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

                //開始時間
                if (Q.StartTime != null)
                {
                    StartTimetxt.Text = ((DateTime)Q.StartTime).ToString("yyyy/MM/dd HH:mm:ss");
                }
                //結束時間
                if (Q.EndTime != null)
                {
                    EndTimetxt.Text = ((DateTime)Q.EndTime).ToString("yyyy/MM/dd HH:mm:ss");
                }

                //提供目前使用者已經有的 群組標題，當作選項，以利使用者快速新增問題
                Column1.Items.Clear();
                foreach (string title in title_list)
                {
                    //群組--選項
                    Column1.Items.Add(title);
                }
                foreach (DataGridViewRow dr in dataGridViewX1.Rows)
                {
                    if (("" + dr.Cells[2].Value) == "選擇")
                    {

                        dr.Cells[6].Value = "點選以編輯";

                        dr.Cells[4].ReadOnly = false;
                        dr.Cells[5].ReadOnly = false;
                        dr.Cells[6].ReadOnly = false;

                        //將背景改回白                        
                        //dr.Cells[4].Style.BackColor = Color.White;                        
                        //dr.Cells[5].Style.BackColor = Color.White;                        
                        //dr.Cells[6].Style.BackColor = Color.White;                        
                    }

                    else  //填答
                    {
                        //清空原本的值

                        dr.Cells[4].Value = "";
                        dr.Cells[5].Value = "";
                        dr.Cells[6].Value = "";

                        dr.Cells[4].ReadOnly = true;
                        dr.Cells[5].ReadOnly = true;
                        dr.Cells[6].ReadOnly = true;

                        //將背景改灰                        
                        //dr.Cells[4].Style.BackColor = Color.Gray;                        
                        //dr.Cells[5].Style.BackColor = Color.Gray;                        
                        //dr.Cells[6].Style.BackColor = Color.Gray;

                    }
                }
                //dataGridViewX1.BeginEdit(false);
                dataGridViewX1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            }
        }
    }
}
