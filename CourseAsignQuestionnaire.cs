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
    public partial class CourseAsignQuestionnaire : BaseForm
    {

        bool asign_record_modified = false;

        QuestionnaireForm currentQ;
        
        List<QuestionnaireForm> formList = new List<QuestionnaireForm>();

        List<QuestionnaireAsignGroup> form_asign_group_List = new List<QuestionnaireAsignGroup>();

        //原本的List 作為與form_asign_group_List比較 確認是否已儲存使用
        List<QuestionnaireAsignGroup> form_asign_group_List_ori = new List<QuestionnaireAsignGroup>();

        //用來記錄現在已有幾項項目被修改， 用此int紀錄方法會比較高效。
        int asign_record_modified_counter = 0;

        // 儲存 由學年度、學期 抓下來的 table 資料
        DataTable dt_course = new DataTable();

        // 將上面 dt_course 經由 年級、科目 、 是否列出已勾選問卷 等三個條件篩選過後的 table，最後顯示在UI上的 依靠此table
        DataTable dt_course_after_filter = new DataTable();

        // 紀錄 listview 是否已填好
        bool listview_already = false;

        List<string> subject_list = new List<string>(); 

        public CourseAsignQuestionnaire()
        {
            InitializeComponent();

            form_asign_group_List_ori = new AccessHelper().Select<QuestionnaireAsignGroup>();

            // 抓取 學校預設 學年度、學期
            string default_school_year =  K12.Data.School.DefaultSchoolYear;
            string default_school_semester = K12.Data.School.DefaultSemester;
                        
            schoolyear_cbox.Items.Add(""+( int.Parse(default_school_year)+1));
            schoolyear_cbox.Items.Add(default_school_year);
            schoolyear_cbox.Items.Add("" + (int.Parse(default_school_year) - 1));

            schoolyear_cbox.Text = default_school_year;

            semester_cbox.Items.Add("1");
            semester_cbox.Items.Add("2");

            semester_cbox.Text = default_school_semester;

            grade_year_cbox.Items.Add("1");
            grade_year_cbox.Items.Add("2");
            grade_year_cbox.Items.Add("3");
            grade_year_cbox.Items.Add("");

            load_questionnaire_form();            
        }

        //save
        private void buttonX1_Click(object sender, EventArgs e)
        {
            save();
        }

        //close
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        // 若學年度 有改變
        private void schoolyear_cbox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (asign_record_modified)
            {
                if (MsgBox.Show("有尚未儲存的問卷更動，是否要儲存?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    save();
                }                
            }
            asign_record_modified = false;
            Modified_Indicator.Visible = false;
            Modified_Indicator.Text = "";


            if (schoolyear_cbox.Text != "" && semester_cbox.Text != "") 
            {
                renew_course_data(schoolyear_cbox.Text, semester_cbox.Text); 
                            
            }
            
        }

        // 若學期 有改變
        private void semester_cbox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (asign_record_modified)
            {
                if (MsgBox.Show("有尚未儲存的問卷更動，是否要儲存?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    save();
                }
            }
            asign_record_modified = false;
            Modified_Indicator.Visible = false;
            Modified_Indicator.Text = "";

            if (schoolyear_cbox.Text != "" && semester_cbox.Text != "")
            {
                renew_course_data(schoolyear_cbox.Text, semester_cbox.Text);
            }
        }
        
        // 載入問卷，並更新左方問卷清單，結束後 載入課程
        private void load_questionnaire_form()
        {
            BackgroundWorker bkw = new BackgroundWorker();

            bkw.DoWork += delegate
            {
                formList = new AccessHelper().Select<QuestionnaireForm>();

                formList.Sort(delegate(QuestionnaireForm f1, QuestionnaireForm f2)
                {
                    //不是每張Table 每個時候都有 StartTime、 EndTime  先用Name 排序
                    return f1.Name.CompareTo(f2.Name);
                });


            };

            bkw.RunWorkerCompleted += delegate
            {
                itmPnlQuestionnaire.Items.Clear();

                

                foreach (QuestionnaireForm Q in formList)
                {
                    XmlDocument doc = new XmlDocument();

                    //假如問卷不是有要被刪除，就列出來
                    if (!Q.Deleted)
                    {
                        ButtonItem btnItem = new ButtonItem();
                        btnItem.Text = Q.Name;
                        btnItem.Tag = Q;
                        btnItem.OptionGroup = "itmPnlTimeName";
                        btnItem.ButtonStyle = eButtonStyle.ImageAndText;
                        
                        btnItem.Click += new EventHandler(btnItem_Click);

                        doc.LoadXml(Q.ContentString);

                        if (doc.SelectSingleNode("Content").SelectNodes("Section").Count == 0) 
                        {
                            btnItem.Enabled = false;                                                
                        }

                        // 預設 先選List 第一個項目 為"選擇狀態"
                        if (currentQ==null && Q == formList[0])
                        {
                            btnItem.Checked = true;
                            currentQ = Q;
                        }

                        itmPnlQuestionnaire.Items.Add(btnItem);
                    }
                }

                itmPnlQuestionnaire.Refresh();

                //結束後 載入課程
                renew_course_data(schoolyear_cbox.Text, semester_cbox.Text);
            };

            bkw.RunWorkerAsync();
               
        }

        //載入課程
        private void renew_course_data(string school_year,string semester) 
        {
            BackgroundWorker bkw = new BackgroundWorker();
            
            bkw.DoWork += delegate
            {                                                
                QueryHelper q = new QueryHelper();

                //  取得 課程之 課程名稱、group_id、年級、科目名稱
                dt_course = q.Select("SELECT course.course_name,course.subject,course.group_id,class.grade_year FROM course LEFT JOIN class ON course.ref_class_id = class.id where school_year =" + "'" + school_year + "'" + "and semester =" + "'" + semester + "'");
            };

            bkw.RunWorkerCompleted += delegate
            {
                // 將抓下來的 dt_course  Column 同步給 dt_course_after_filter  Columns，後續才能做資料處理
                foreach (DataColumn dc in dt_course.Columns)
                {
                    if (!dt_course_after_filter.Columns.Contains(dc.ColumnName))
                    {
                        dt_course_after_filter.Columns.Add(dc.ColumnName, dc.DataType);                    
                    }                    
                }

                //使選擇年級 回到預設(空白)
                grade_year_cbox.SelectedIndex = -1;

                listview_already = false;

                
            };

            bkw.RunWorkerAsync();                        
        }

        private void renew_course_data_with_now_setting(string school_year, string semester)
        {
            BackgroundWorker bkw = new BackgroundWorker();

            bkw.DoWork += delegate
            {
                QueryHelper q = new QueryHelper();

                //  取得 課程之 課程名稱、group_id、年級、科目名稱
                dt_course = q.Select("SELECT course.course_name,course.subject,course.group_id,class.grade_year FROM course LEFT JOIN class ON course.ref_class_id = class.id where school_year =" + "'" + school_year + "'" + "and semester =" + "'" + semester + "'");
            };

            bkw.RunWorkerCompleted += delegate
            {
                // 將抓下來的 dt_course  Column 同步給 dt_course_after_filter  Columns，後續才能做資料處理
                foreach (DataColumn dc in dt_course.Columns)
                {
                    if (!dt_course_after_filter.Columns.Contains(dc.ColumnName))
                    {
                        dt_course_after_filter.Columns.Add(dc.ColumnName, dc.DataType);
                    }
                }

                //使選擇年級 回到預設(空白)
                //grade_year_cbox.SelectedIndex = -1;

                listview_already = false;

                show_all_selected_course();
            };

            bkw.RunWorkerAsync();
        }

        //點選左邊問卷List 項目
        void btnItem_Click(object sender, EventArgs e)
        {
            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (asign_record_modified)
            {
                if (MsgBox.Show("有尚未儲存的問卷更動，是否要儲存?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    save();
                }
            }
            asign_record_modified = false;
            Modified_Indicator.Visible = false;
            Modified_Indicator.Text = "";


            if (itmPnlQuestionnaire.SelectedItems.Count == 1)
            {
                //取得目前所選擇的Button
                ButtonItem Buttonitem = itmPnlQuestionnaire.SelectedItems[0] as ButtonItem;

                //取得問卷Record
                QuestionnaireForm Q = (QuestionnaireForm)Buttonitem.Tag;

                //設定新的currentQ 為現在選的問卷項目
                currentQ = Q;

                // 更新課程項目(使用原有UI 學年度、學期 、年級、科目)
                renew_course_data_with_now_setting(schoolyear_cbox.Text, semester_cbox.Text);

            }
        }

        // 年級改變時
        private void grade_year_cbox_SelectedIndexChanged(object sender, EventArgs e)
        {

            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (asign_record_modified)
            {
                if (MsgBox.Show("有尚未儲存的問卷更動，是否要儲存?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    save();
                }
            }
            asign_record_modified = false;
            Modified_Indicator.Visible = false;
            Modified_Indicator.Text = "";

            grade_year_filter_process();

        }


        //科目改變時
        private void subject_cbox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (asign_record_modified)
            {
                if (MsgBox.Show("有尚未儲存的問卷更動，是否要儲存?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    save();
                }
            }
            asign_record_modified = false;
            Modified_Indicator.Visible = false;
            Modified_Indicator.Text = "";

            subject_filter_process();
               
        }

        private void grade_year_filter_process()
        {
            // 將原設定清空
            dt_course_after_filter.Clear();
            course_listViewEx.Items.Clear();
            subject_list.Clear();
            subject_cbox.Items.Clear();

            //整理新的 subject_list                    
            foreach (DataRow row in dt_course.Rows)
            {
                if ("" + row["grade_year"] == grade_year_cbox.Text) 
                {                                        
                    if (!subject_list.Contains("" + row["subject"])) 
                    {
                        subject_list.Add("" + row["subject"]);                                        
                    }
                }                
            }
            foreach (string subject_name in subject_list)
            {
                subject_cbox.Items.Add(subject_name);
                                    
            }                                                        
        }

        private void subject_filter_process()
        {

            dt_course_after_filter.Clear();

            List<string> already_in_form_group_list = new List<string>();

            course_listViewEx.Items.Clear();

            asign_record_modified = false;
            listview_already = false;
            asign_record_modified_counter = 0;

            course_listViewEx.SuspendLayout();

            form_asign_group_List = new AccessHelper().Select<QuestionnaireAsignGroup>();

            if (currentQ != null)
            {
                foreach (QuestionnaireAsignGroup item in form_asign_group_List)
                {
                    if ("" + item.RefFormId == currentQ.UID)
                    {
                        already_in_form_group_list.Add("" + item.RefGroupId);
                    }
                }
            }

            // 如果 年級、科目 條件相同 才加入 dt_course_after_filter
            foreach (DataRow row in dt_course.Rows)
            {
                if ("" + row["grade_year"] == grade_year_cbox.Text && "" + row["subject"] == subject_cbox.Text)
                {
                    // 自另一DT 引入row 至新DT
                    dt_course_after_filter.ImportRow(row);                    
                }
            }
            foreach (DataRow row in dt_course_after_filter.Rows)
            {
                ListViewItem lv_item = new ListViewItem();

                lv_item.Text = "" + row["course_name"];

                lv_item.Tag = "" + row["group_id"];

                // 如果 原本是以加入的課程group_id，將其勾起
                if (already_in_form_group_list.Contains("" + row["group_id"]))
                {
                    lv_item.Checked = true;
                }

                course_listViewEx.Items.Add(lv_item);
            }

            course_listViewEx.ResumeLayout();

            listview_already = true;

        }

        // 當 "列出所有已設定本問卷課程 " Chk 被 點選時
        private void list_all_checked_chkbox_Click(object sender, EventArgs e)
        {
            show_all_selected_course();
            
        }


        private void show_all_selected_course() 
        {
            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (asign_record_modified)
            {
                if (MsgBox.Show("有尚未儲存的問卷更動，是否要儲存?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    save();
                }
            }
            asign_record_modified = false;
            Modified_Indicator.Visible = false;
            Modified_Indicator.Text = "";

            if (list_all_checked_chkbox.Checked)
            {
                // 將上面項目停止
                schoolyear_cbox.Enabled = false;
                semester_cbox.Enabled = false;
                grade_year_cbox.Enabled = false;
                subject_cbox.Enabled = false;

                asign_record_modified = false;
                listview_already = false;
                asign_record_modified_counter = 0;

                dt_course_after_filter.Clear();

                List<string> already_in_form_group_list = new List<string>();

                course_listViewEx.Items.Clear();

                course_listViewEx.SuspendLayout();

                form_asign_group_List = new AccessHelper().Select<QuestionnaireAsignGroup>();

                if (currentQ != null)
                {
                    foreach (QuestionnaireAsignGroup item in form_asign_group_List)
                    {
                        if ("" + item.RefFormId == currentQ.UID)
                        {
                            already_in_form_group_list.Add("" + item.RefGroupId);
                        }
                    }
                }

                //僅列出已加入 才加入 dt_course_after_filter
                foreach (DataRow row in dt_course.Rows)
                {
                    if (already_in_form_group_list.Contains("" + row["group_id"]))
                    {
                        // 自另一DT 引入row 至新DT
                        dt_course_after_filter.ImportRow(row);
                    }
                }
                foreach (DataRow row in dt_course_after_filter.Rows)
                {
                    ListViewItem lv_item = new ListViewItem();

                    lv_item.Text = "" + row["course_name"];

                    lv_item.Tag = "" + row["group_id"];

                    // 如果 原本是以加入的課程group_id，將其勾起
                    if (already_in_form_group_list.Contains("" + row["group_id"]))
                    {
                        lv_item.Checked = true;
                    }

                    course_listViewEx.Items.Add(lv_item);
                }

                course_listViewEx.ResumeLayout();

                listview_already = true;
            }
            else
            {
                //將上面項目開啟
                schoolyear_cbox.Enabled = true;
                semester_cbox.Enabled = true;
                grade_year_cbox.Enabled = true;
                subject_cbox.Enabled = true;

                if (schoolyear_cbox.Text != "" && semester_cbox.Text != "")
                {
                    subject_filter_process();
                }
            }
        
        
        
        }


        //判斷是否有修改
        private void course_listViewEx_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            // 假如表已填好， 避免在讀取匯入時 觸發
            if (listview_already) 
            {
                // 假如點下去 是勾選
                if (e.Item.Checked)
                {
                    bool in_ori_list = false;

                    
                    foreach (QuestionnaireAsignGroup ori_item in form_asign_group_List_ori)
                    {
                        if (("" + ori_item.RefFormId == currentQ.UID) && ("" + ori_item.RefGroupId == "" + e.Item.Tag))
                        {
                            asign_record_modified_counter--;

                            in_ori_list = true;
                        }
                    }

                    if (!in_ori_list) 
                    {
                        asign_record_modified_counter++;
                    }

                }
                // 假如點下去 是取消勾選
                else
                {
                    bool not_in_ori_list = true;

                    foreach (QuestionnaireAsignGroup ori_item in form_asign_group_List_ori)
                    {
                        if (("" + ori_item.RefFormId == currentQ.UID) && ("" + ori_item.RefGroupId == "" + e.Item.Tag))
                        {


                            asign_record_modified_counter++;

                            not_in_ori_list = false;
                        }
                    }
                    if (not_in_ori_list)
                    {
                        asign_record_modified_counter--;
                    }
                }

                // > 0 就是有修改
                if (asign_record_modified_counter > 0)
                {
                    asign_record_modified = true;
                }
                if (asign_record_modified_counter == 0)
                {
                    asign_record_modified = false;
                }

                if (asign_record_modified)
                {
                    Modified_Indicator.Visible = true;
                    Modified_Indicator.Text = "已修改，尚未儲存。";
                }
                else
                {
                    Modified_Indicator.Visible = false;
                    Modified_Indicator.Text = "";
                }                        
            }            
        }

        private void CourseAsignQuestionnaire_FormClosing(object sender, FormClosingEventArgs e)
        {
            
            // 確認使用者 是否有遺忘儲存 改變後的資料
            if (asign_record_modified)
            {                
                if (MsgBox.Show("有尚未儲存的問卷更動，是否要儲存?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    e.Cancel = true;                    
                    save();
                    e.Cancel = false;
                }
                else
                {
                    //阻止離開程序
                    e.Cancel = true;
                }
            }
        }

        private void save() 
        {
            if (currentQ != null)
            {
                foreach (ListViewItem item in course_listViewEx.Items)
                {
                    if (item.Checked)
                    {
                        // 確認 原本是否就已經勾選
                        bool alreadyAsign = false;

                        foreach (QuestionnaireAsignGroup asignitem in form_asign_group_List)
                        {
                            if (("" + asignitem.RefFormId == currentQ.UID) && ("" + item.Tag == "" + asignitem.RefGroupId))
                            {
                                alreadyAsign = true;
                            }
                        }
                        //還沒有加入問卷，就加入
                        if (!alreadyAsign)
                        {
                            QuestionnaireAsignGroup new_asign_item = new QuestionnaireAsignGroup();

                            new_asign_item.RefFormId = int.Parse(currentQ.UID);
                            new_asign_item.RefGroupId = int.Parse("" + item.Tag);

                            form_asign_group_List.Add(new_asign_item);
                        }
                    }
                    else
                    {
                        foreach (QuestionnaireAsignGroup asignitem in form_asign_group_List)
                        {
                            if (("" + asignitem.RefFormId == currentQ.UID) && ("" + item.Tag == "" + asignitem.RefGroupId))
                            {
                                asignitem.Deleted = true;
                            }
                        }
                    }
                }
                form_asign_group_List.SaveAll();
            }


            Modified_Indicator.Visible = true;
            Modified_Indicator.Text = "已儲存";
            asign_record_modified = false;


            form_asign_group_List_ori = new AccessHelper().Select<QuestionnaireAsignGroup>();

            MsgBox.Show("儲存成功");
        
        
        
        }

    }
}
