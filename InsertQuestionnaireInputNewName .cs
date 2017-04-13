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
    public partial class InsertQuestionnaireInputNewName : BaseForm
    {

        // 所有現有問卷名子的清單
        List<string> q_name_ori_list = new List<string>();


        public InsertQuestionnaireInputNewName(List<string> _q_name_ori_list)
        {
            InitializeComponent();

            q_name_ori_list = _q_name_ori_list;                    
        }

        //儲存
        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (q_name_ori_list.Contains(textBoxX1.Text))
            {
                MsgBox.Show("問卷:" + textBoxX1.Text + "名稱重覆，請修正。", "新增問卷失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // 名稱 重覆 傳遞 DialogResult.Cancel 回 母form ，避免儲存。
                this.DialogResult = DialogResult.Cancel;

                return;            
            }

            this.Close();
        }

        //儲存
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        // 用來傳遞值回 母 Form
        public string q_name
        {
            set
            {
                textBoxX1.Text = value;
            }
            get
            {
                return textBoxX1.Text;
            }
        }
    }
}
