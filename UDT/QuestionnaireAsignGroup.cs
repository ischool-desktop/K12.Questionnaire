using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Questionnaire.UDT
{
    [TableName("ischool.1campus.questionnaire.group")]
    class QuestionnaireAsignGroup : ActiveRecord
    {
        [Field(Field = "ref_group_id", Indexed =false)]
        public int RefGroupId { get; set; }

        [Field(Field = "ref_form_id", Indexed = false)]
        public int RefFormId { get; set; }
        
    }
}
