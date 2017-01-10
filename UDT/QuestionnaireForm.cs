using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Questionnaire.UDT
{
    [TableName("ischool.1campus.questionnaire.form")]
    class QuestionnaireForm:ActiveRecord
    {
        [Field(Field = "ref_teacher_id", Indexed =false)]
        public int? RefTeacherID { get; set; }
        [Field(Field = "name", Indexed = false)]
        public string Name { get; set; }
        [Field(Field = "student_reply", Indexed = false)]
        public bool StudentReply { get; set; }
        [Field(Field = "parent_reply", Indexed = false)]
        public bool ParentReply { get; set; }
        [Field(Field = "teacher_reply", Indexed = false)]
        public bool TeacherReply { get; set; }
        [Field(Field = "start_time", Indexed = false)]
        public DateTime? StartTime { get; set; }
        [Field(Field = "end_time", Indexed = false)]
        public DateTime? EndTime { get; set; }
        [Field(Field = "content", Indexed = false)]
        public string ContentString { get; set; }
    }
}
