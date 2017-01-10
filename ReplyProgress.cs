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
    public partial class ReplyProgress : BaseForm
    {
        public ReplyProgress()
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
                dataGridViewX1.Rows.Clear();
                DataTable dt = null;
                var bkw = new BackgroundWorker();
                bkw.DoWork += delegate
                {
                    dt = new QueryHelper().Select(@"
SELECT groupx.group_name, groupx.uid, count(groupx.ref_student_id) as student_count, count(reply.uid) as reply_count
FROM
    $ischool.1campus.questionnaire.group
	LEFT OUTER JOIN (  
		SELECT 'group' as kind, groupx.uid as uid, group_name as group_name, $sg_attend.ref_student_id as ref_student_id
		FROM $group as groupx LEFT OUTER JOIN $sg_attend on $sg_attend.ref_group_id = groupx.uid  
		union all  

        SELECT 'class' as kind, class.group_id as uid, class_name as group_name, student.id as ref_student_id
		FROM class LEFT OUTER JOIN student on class.id = student.ref_class_id  
		union all  

        SELECT 'course' as kind, course.group_id as uid, course_name as group_name, sc_attend.ref_student_id as ref_student_id
		FROM course LEFT OUTER JOIN sc_attend on course.id = sc_attend.ref_course_id  
	)  as groupx on groupx.uid = $ischool.1campus.questionnaire.group.ref_group_id
    LEFT OUTER JOIN $ischool.1campus.questionnaire.reply as reply on reply.ref_group_id = $ischool.1campus.questionnaire.group.ref_group_id AND reply.ref_form_id = $ischool.1campus.questionnaire.group.ref_form_id AND reply.ref_student_id = groupx.ref_student_id
    LEFT OUTER JOIN student on student.id = groupx.ref_student_id
WHERE
    student.status in (1, 2)
    AND $ischool.1campus.questionnaire.group.ref_form_id = " + form.UID + @"
GROUP BY groupx.group_name, groupx.uid
");
                };
                bkw.RunWorkerCompleted += delegate
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        dataGridViewX1.Rows[dataGridViewX1.Rows.Add("" + row["group_name"], "" + row["reply_count"], "" + row["student_count"])].DefaultCellStyle.ForeColor = (("" + row["student_count"]) == ("" + row["reply_count"]) ? Color.Gray : Color.Red);
                    }
                };
                bkw.RunWorkerAsync();
            }

        }
    }
}
