using FISCA;
using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace K12.Questionnaire
{
    public static class Program
    {
        [MainMethod()]
        public static void Main()
        {
            {
                Catalog catalog = RoleAclSource.Instance["課程回饋"]["功能"];
                catalog.Add(new RibbonFeature("FEE2BDD7-29F1-48D4-A02E-7CA10737324D", "管理問卷"));

                var btn = K12.Presentation.NLDPanels.Course.RibbonBarItems["課程回饋"]["管理問卷"];
                btn.Enable = UserAcl.Current["FEE2BDD7-29F1-48D4-A02E-7CA10737324D"].Executable;
                btn.Click += delegate
                {
                    new ManagementQuestionnaire().ShowDialog();


                };
            }
            {
                Catalog catalog = RoleAclSource.Instance["課程回饋"]["功能"];
                catalog.Add(new RibbonFeature("62421DBC-34E1-4AEC-A832-6EDAC49A40A5", "設定調查群組"));

                var btn = K12.Presentation.NLDPanels.Course.RibbonBarItems["課程回饋"]["設定調查群組"];
                btn.Enable = UserAcl.Current["62421DBC-34E1-4AEC-A832-6EDAC49A40A5"].Executable;
                btn.Click += delegate
                {
                    FISCA.Presentation.Controls.MsgBox.Show("功能尚未完成。");
                };
            }
            {
                Catalog catalog = RoleAclSource.Instance["課程回饋"]["功能"];
                catalog.Add(new RibbonFeature("0F97E4CF-C699-48D2-86D5-6D5E195C412B", "檢查填寫進度"));

                var btn = K12.Presentation.NLDPanels.Course.RibbonBarItems["課程回饋"]["檢查填寫進度"];
                btn.Enable = UserAcl.Current["0F97E4CF-C699-48D2-86D5-6D5E195C412B"].Executable;
                btn.Click += delegate
                {
                    new ReplyProgress().ShowDialog();
                };
            }
            {
                Catalog catalog = RoleAclSource.Instance["課程回饋"]["功能"];
                catalog.Add(new RibbonFeature("7E12122B-A98C-4E71-B411-C186535C09A8", "匯出填寫明細"));

                var btn = K12.Presentation.NLDPanels.Course.RibbonBarItems["課程回饋"]["匯出填寫明細"];
                btn.Enable = UserAcl.Current["7E12122B-A98C-4E71-B411-C186535C09A8"].Executable;
                btn.Click += delegate
                {
                    new ExportReply().ShowDialog();
                };
            }
        }
    }
}
