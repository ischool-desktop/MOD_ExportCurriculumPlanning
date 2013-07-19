using System;
using System.Collections.Generic;
using System.Text;
using SmartSchool;
using FISCA.Presentation;
using FISCA;
using FISCA.Permission;

namespace ExportCurriculumPlanning
{
    public static class Program
    {
        [MainMethod()]
        public static void Main()
        {
            RibbonBarItem item = MotherForm.RibbonBarItems["教務作業", "資料統計"];
            item["匯出"].Image = Properties.Resources.Export_Image;
            item["匯出"].Size = RibbonBarButton.MenuButtonSize.Large;
            item["匯出"]["匯出課程規劃表"].Enable = 匯出課程規劃表權限;
            item["匯出"]["匯出課程規劃表"].Click += delegate
            {
                ClickButton f = new ClickButton();
                f.ShowDialog();
            };

            Catalog detail1;
            detail1 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            detail1.Add(new RibbonFeature(匯出課程規劃表, "匯出課程規劃表"));
        }

        public static string 匯出課程規劃表 { get { return "20130719.new.ExportCurriculumPlanning"; } }
        public static bool 匯出課程規劃表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[匯出課程規劃表].Executable;
            }
        }
    }
}
