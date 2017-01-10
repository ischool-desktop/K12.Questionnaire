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

namespace K12.Questionnaire
{
    public partial class ExportReply : BaseForm
    {
        public ExportReply()
        {
            InitializeComponent();

            comboBoxEx1.Items.Add("資料讀取中...");
            comboBoxEx1.SelectedIndex = 0;


            List<QuestionnaireForm> formList = new List<QuestionnaireForm>();
            var bkw = new BackgroundWorker();
            bkw.DoWork += delegate
            {
                formList = new AccessHelper().Select<QuestionnaireForm>("ref_teacher_id = null AND NOT(end_time is null)");
                formList.Sort(delegate (QuestionnaireForm f1, QuestionnaireForm f2)
                {
                    return f1.EndTime.Value.CompareTo(f2.EndTime.Value);
                });
            };
            bkw.RunWorkerCompleted += delegate
            {
                comboBoxEx1.Items.Clear();
                comboBoxEx1.DisplayMember = "Name";
                comboBoxEx1.Items.AddRange(formList.ToArray());
                if (comboBoxEx1.Items.Count > 0)
                    comboBoxEx1.SelectedIndex = 0;
            };
            bkw.RunWorkerAsync();
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxEx1.SelectedItem is QuestionnaireForm)
            {
                var form = comboBoxEx1.SelectedItem as QuestionnaireForm;
                listViewEx1.Items.Clear();
                btnExport.Enabled = false;
                DataTable dt = null;
                var bkw = new BackgroundWorker();
                bkw.DoWork += delegate
                {
                    dt = new QueryHelper().Select(@"
SELECT groupx.group_name, groupx.uid
FROM
    $ischool.1campus.questionnaire.group
	LEFT OUTER JOIN (  
		SELECT 'group' as kind, groupx.uid as uid, group_name as group_name
		FROM $group as groupx 
		union all  

        SELECT 'class' as kind, class.group_id as uid, class_name as group_name
		FROM class
		union all  

        SELECT 'course' as kind, course.group_id as uid, course_name as group_name
		FROM course
	)  as groupx on groupx.uid = $ischool.1campus.questionnaire.group.ref_group_id
WHERE
    $ischool.1campus.questionnaire.group.ref_form_id = " + form.UID);
                };
                bkw.RunWorkerCompleted += delegate
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        listViewEx1.Items.Add("" + row["group_name"]).Tag = "" + row["uid"];
                    }
                };
                bkw.RunWorkerAsync();
            }

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (comboBoxEx1.SelectedItem != null && comboBoxEx1.SelectedItem is QuestionnaireForm)
            {
                var form = comboBoxEx1.SelectedItem as QuestionnaireForm;
                List<string> groupIDList = new List<string>();
                groupIDList.Add("-1");
                foreach (ListViewItem item in listViewEx1.CheckedItems)
                {
                    groupIDList.Add("" + item.Tag);
                }

                Workbook wb = new Workbook();
                var bkw = new BackgroundWorker();
                bkw.DoWork += delegate
                {
                    string sql = @"
SELECT 
    form.name as form_name, student_questionnaire.group_name
    , class.class_name, student.seat_no, student.student_number, student.name as student_name
    , form.content, reply.reply 
FROM 
	( 
		SELECT ref_form_id, ref_group_id, group_name, ref_student_id  
		FROM 
			(  
				SELECT 'group' as kind, groupx.uid as uid, group_name as group_name, $sg_attend.ref_student_id as ref_student_id, group_name as icon  
				FROM $group as groupx LEFT OUTER JOIN $sg_attend on $sg_attend.ref_group_id = groupx.uid  
				union all  

                SELECT 'class' as kind, class.group_id as uid, class_name as group_name, student.id as ref_student_id, '班' as icon   
				FROM class LEFT OUTER JOIN student on class.id = student.ref_class_id  
				union all  

                SELECT 'course' as kind, course.group_id as uid, course_name as group_name, sc_attend.ref_student_id as ref_student_id, subject as icon  
				FROM course LEFT OUTER JOIN sc_attend on course.id = sc_attend.ref_course_id  
			)  as groupx 
			LEFT OUTER JOIN $ischool.1campus.questionnaire.group on groupx.uid = $ischool.1campus.questionnaire.group.ref_group_id  
            LEFT OUTER JOIN student on student.id = ref_student_id
        WHERE 
            student.status in (1, 2) 
        
        UNION 
        
        SELECT ref_form_id, ref_group_id, group_name, ref_student_id
        FROM 
            $ischool.1campus.questionnaire.reply
            LEFT OUTER JOIN (
                SELECT 'group' as kind, groupx.uid as uid, group_name as group_name
                FROM $group as groupx

                UNION ALL

                SELECT 'class' as kind, class.group_id as uid, class_name as group_name
                FROM class

                UNION ALL

                SELECT 'course' as kind, course.group_id as uid, course_name as group_name
                FROM course                
            ) as gp on gp.uid = ref_group_id
	) as student_questionnaire 
	LEFT OUTER JOIN $ischool.1campus.questionnaire.reply as reply on reply.ref_group_id = student_questionnaire.ref_group_id AND reply.ref_form_id = student_questionnaire.ref_form_id AND reply.ref_student_id = student_questionnaire.ref_student_id
	LEFT OUTER JOIN $ischool.1campus.questionnaire.form as form on form.uid = student_questionnaire.ref_form_id 
    LEFT OUTER JOIN student on student.id = student_questionnaire.ref_student_id
    LEFT OUTER JOIN class on class.id = student.ref_class_id
WHERE form.uid = " + form.UID + " AND student_questionnaire.ref_group_id in ( " + string.Join(",", groupIDList) + ")";
                    var queryHelper = new QueryHelper();
                    var dt = queryHelper.Select(sql);
                    XmlDocument doc = new XmlDocument();

                    Worksheet ws = wb.Worksheets[0];

                    var rowIndex = 0;
                    var colIndex = 0;
                    var dicPathCellIndex = new Dictionary<string, int>();
                    foreach (DataRow row in dt.Rows)
                    {
                        if (rowIndex == 0)
                        {
                            #region 產出表頭
                            colIndex = 0;
                            ws.Cells[rowIndex, colIndex++].PutValue("問卷名稱");
                            ws.Cells[rowIndex, colIndex++].PutValue("調查群組");
                            ws.Cells[rowIndex, colIndex++].PutValue("班級");
                            ws.Cells[rowIndex, colIndex++].PutValue("座號");
                            ws.Cells[rowIndex, colIndex++].PutValue("學號");
                            ws.Cells[rowIndex, colIndex++].PutValue("姓名");

                            doc.LoadXml("" + row["content"]);
                            #region 填寫值
                            foreach (XmlElement sectionElement in doc.SelectNodes("Content/Section"))
                            {
                                var sectionTitle = sectionElement.SelectSingleNode("Title").InnerText;
                                foreach (XmlElement questionElement in sectionElement.SelectNodes("Question"))
                                {
                                    var questionTitle = questionElement.SelectSingleNode("Title").InnerText;
                                    var questionType = questionElement.SelectSingleNode("Type").InnerText;

                                    var path = sectionTitle + "_" + questionTitle;
                                    dicPathCellIndex.Add(path, colIndex);
                                    ws.Cells[rowIndex, colIndex++].PutValue(questionTitle);
                                }
                            }
                            #endregion
                            #region 勾選選項
                            foreach (XmlElement sectionElement in doc.SelectNodes("Content/Section"))
                            {
                                var sectionTitle = sectionElement.SelectSingleNode("Title").InnerText;
                                foreach (XmlElement questionElement in sectionElement.SelectNodes("Question"))
                                {
                                    var questionTitle = questionElement.SelectSingleNode("Title").InnerText;
                                    var questionType = questionElement.SelectSingleNode("Type").InnerText;
                                    if (questionType == "option")
                                    {
                                        foreach (XmlElement optionElement in questionElement.SelectNodes("Option"))
                                        {
                                            var option = optionElement.InnerText;
                                            var path = sectionTitle + "_" + questionTitle + "_" + option;
                                            dicPathCellIndex.Add(path, colIndex);
                                            ws.Cells[rowIndex, colIndex++].PutValue("勾選[" + questionTitle + "][" + option.Replace("%TEXT%", "____") + "]");
                                        }
                                    }
                                }
                            }
                            #endregion
                            #region 選項補充填寫
                            foreach (XmlElement sectionElement in doc.SelectNodes("Content/Section"))
                            {
                                var sectionTitle = sectionElement.SelectSingleNode("Title").InnerText;
                                foreach (XmlElement questionElement in sectionElement.SelectNodes("Question"))
                                {
                                    var questionTitle = questionElement.SelectSingleNode("Title").InnerText;
                                    var questionType = questionElement.SelectSingleNode("Type").InnerText;
                                    if (questionType == "option")
                                    {
                                        foreach (XmlElement optionElement in questionElement.SelectNodes("Option"))
                                        {
                                            var option = optionElement.InnerText;
                                            if (option.Contains("%TEXT%"))
                                            {
                                                var split = option.Split(new string[] { "%TEXT%" }, StringSplitOptions.None);
                                                for (int i = 1; i < split.Length; i++)
                                                {
                                                    var path = sectionTitle + "_" + questionTitle + "_" + option + "_split" + i;
                                                    dicPathCellIndex.Add(path, colIndex);
                                                    ws.Cells[rowIndex, colIndex++].PutValue("填值[" + questionTitle + "][" + option.Replace("%TEXT%", "____") + "][值" + i + "]");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion
                            rowIndex++;
                            #endregion
                        }

                        colIndex = 0;
                        ws.Cells[rowIndex, colIndex++].PutValue("" + row["form_name"]);
                        ws.Cells[rowIndex, colIndex++].PutValue("" + row["group_name"]);
                        ws.Cells[rowIndex, colIndex++].PutValue("" + row["class_name"]);
                        ws.Cells[rowIndex, colIndex++].PutValue("" + row["seat_no"]);
                        ws.Cells[rowIndex, colIndex++].PutValue("" + row["student_number"]);
                        ws.Cells[rowIndex, colIndex++].PutValue("" + row["student_name"]);
                        if (("" + row["reply"]) != "")
                        {
                            doc.LoadXml("" + row["reply"]);
                            foreach (XmlElement sectionElement in doc.SelectNodes("Reply/Section"))
                            {
                                var sectionTitle = sectionElement.SelectSingleNode("Title").InnerText;
                                foreach (XmlElement questionElement in sectionElement.SelectNodes("Question"))
                                {
                                    var questionTitle = questionElement.SelectSingleNode("Title").InnerText;
                                    var answerElement = questionElement.SelectSingleNode("Answer");
                                    if (answerElement != null)
                                    {//填入值
                                        var path = sectionTitle + "_" + questionTitle;
                                        if (dicPathCellIndex.ContainsKey(path))
                                            ws.Cells[rowIndex, dicPathCellIndex[path]].PutValue(answerElement.InnerText);
                                    }
                                    if (questionElement.SelectNodes("Chose").Count > 0)
                                    {
                                        List<string> choseValues = new List<string>();
                                        foreach (XmlElement choseElement in questionElement.SelectNodes("Chose"))
                                        {
                                            var option = choseElement.SelectSingleNode("Option").InnerText;
                                            var splitList = new List<string>();
                                            {//填入勾選
                                                var path = sectionTitle + "_" + questionTitle + "_" + option;
                                                if (dicPathCellIndex.ContainsKey(path))
                                                    ws.Cells[rowIndex, dicPathCellIndex[path]].PutValue("是");
                                            }
                                            var spliteIndex = 0;
                                            foreach (XmlElement splitElement in choseElement.SelectNodes("Split"))
                                            {
                                                spliteIndex++;
                                                splitList.Add(splitElement.InnerText);
                                                if (spliteIndex % 2 == 0)
                                                {//填入自填值
                                                    var path = sectionTitle + "_" + questionTitle + "_" + option + "_split" + (spliteIndex / 2);
                                                    if (dicPathCellIndex.ContainsKey(path))
                                                        ws.Cells[rowIndex, dicPathCellIndex[path]].PutValue(splitElement.InnerText);
                                                }
                                            }
                                            choseValues.Add(string.Join("", splitList));
                                        }
                                        {//填入選取值
                                            var path = sectionTitle + "_" + questionTitle;
                                            if (dicPathCellIndex.ContainsKey(path))
                                                ws.Cells[rowIndex, dicPathCellIndex[path]].PutValue(string.Join("、", choseValues));
                                        }
                                    }
                                }
                            }
                        }
                        rowIndex++;
                    }
                    //ws.AutoFitColumns(0, 0, 100, colIndex);
                };
                bkw.RunWorkerCompleted += delegate
                {

                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.FileName = "匯出課程回饋填寫明細";
                    saveFileDialog1.Filter = "Excel (*.xls)|*.xls";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        wb.Save(saveFileDialog1.FileName);
                        System.Diagnostics.Process.Start(saveFileDialog1.FileName);
                    }
                };
                bkw.RunWorkerAsync();
            }
        }

        private void listViewEx1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            btnExport.Enabled = listViewEx1.CheckedItems.Count > 0;
        }
    }
}
