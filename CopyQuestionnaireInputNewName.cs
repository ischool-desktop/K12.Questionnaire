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
    public partial class CopyQuestionnaireInputNewName : BaseForm
    {
        // 原問卷名稱
        string q_name_ori;

        public CopyQuestionnaireInputNewName(string q_name)
        {
            InitializeComponent();

            q_name_ori = q_name;

            textBoxX1.Text = q_name;
        }

        //儲存
        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (textBoxX1.Text == q_name_ori)
            {
                MsgBox.Show("問卷:" + q_name_ori + "名稱重覆，請修正。", "儲存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);

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
