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
    public partial class QuestionnaireOptionSetting : BaseForm
    {

        List<string> opt_list_temp = new List<string>();

        public QuestionnaireOptionSetting(List<string> opt_list)
        {
            InitializeComponent();

            if (opt_list != null) 
            {
                // 載入
                foreach (string opt in opt_list)
                {
                    DataGridViewRow dr = new DataGridViewRow();

                    dr.CreateCells(dataGridViewX1);

                    dr.Cells[0].Value = opt;

                    dataGridViewX1.Rows.Add(dr);

                }                                    
            }           
        }


        // 儲存
        private void buttonX1_Click(object sender, EventArgs e)
        {           
            foreach (DataGridViewRow dr in dataGridViewX1.Rows)         
            {
                if (!dr.IsNewRow) 
                {
                    opt_list_temp.Add("" + dr.Cells[0].Value);                           
                }                
            }
            this.Close();
        }

        //取消
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 用來傳遞值回 母 Form
        public List<string> opt_list
        {
            set
            {
                opt_list_temp = value;
            }
            get
            {
                return opt_list_temp;
            }
        }

        // 動態設定 EditMode
        private void dataGridViewX1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != -1) 
            {
                dataGridViewX1.BeginEdit(true);
                dataGridViewX1.EditMode = DataGridViewEditMode.EditOnEnter;                                    
            }
            else
            {
                dataGridViewX1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            }
        }
    }
}
