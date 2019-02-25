using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using NPOI.SS.UserModel;

namespace testNpoi
{
    public partial class import51xuanxiao : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        

        protected void Button1_Click(object sender, EventArgs e) //导入专业
        {
            IWorkbook wb = WorkbookFactory.Create(FileUpload1.PostedFile.InputStream);

            ISheet sheet = wb.GetSheetAt(0);

            var rows = sheet.GetEnumerator();

            Dictionary<string, int> dict = new Dictionary<string, int>();
            List<Major> majorList = new List<Major>();

            while (rows.MoveNext())
            {
                IRow row = rows.Current as IRow;
                List<ICell> cells = row.Cells;
                if (row != null && cells != null)
                {
                    if (row.RowNum == 0)
                    {
                        for (int i = 0; i < cells.Count; i++)
                        {
                            var txt = cells[i].ToString().Trim();
                            switch (txt)
                            {
                                case "名称": dict.Add("名称", i); break;
                                case "专业": dict.Add("专业", i); break;
                                case "所属学科": dict.Add("所属学科", i); break;
                                case "所属门类": dict.Add("所属门类", i); break;
                                case "专业代码": dict.Add("专业代码", i); break;
                                case "专业介绍": dict.Add("专业介绍", i); break;
                            }
                        }
                    }  //第一行

                    else //数据行
                    {
                        
                        ICell cellName = row.GetCell(dict["名称"]);
                        ICell cellZhuanYe = row.GetCell(dict["专业"]);
                        ICell cellXueKe = row.GetCell(dict["所属学科"]);
                        ICell cellMenLei = row.GetCell(dict["所属门类"]);
                        ICell cellDaiMa = row.GetCell(dict["专业代码"]);
                        ICell cellJieShao = row.GetCell(dict["专业介绍"]);

                        if (cellName != null)
                        {
                            string name= cellName.ToString().Trim();
                            if (name != "")
                            {
                                Major major = new Major();
                                major.Name = name;
                                if (cellZhuanYe != null)
                                {
                                    major.MajorType = (cellZhuanYe.ToString().Trim() == "本科" ? 1 : 2);
                                }

                                if (cellXueKe != null)
                                {
                                    major.SubjectName = cellXueKe.ToString().Trim();
                                }

                                if (cellMenLei != null)
                                {
                                    major.CategoryName = cellMenLei.ToString().Trim();
                                }

                                if (cellDaiMa != null)
                                {
                                    major.MajorCode = cellDaiMa.ToString().Trim();
                                }

                                if (cellJieShao != null)
                                {
                                    major.Introduction = cellJieShao.ToString().Trim();
                                }
                                majorList.Add(major);
                            }
                        }
                    }

                }
            }


            foreach (var item in majorList)
            {
                AddMajor(item);
            }

            

            
        }

        protected void Button2_Click(object sender, EventArgs e) //导入职业
        {
            IWorkbook wb = WorkbookFactory.Create(FileUpload1.PostedFile.InputStream);

            ISheet sheet = wb.GetSheetAt(0);

            var rows = sheet.GetEnumerator();

            Dictionary<string, int> dict = new Dictionary<string, int>();
            List<Job> jobList = new List<Job>();

            while (rows.MoveNext())
            {
                IRow row = rows.Current as IRow;
                List<ICell> cells = row.Cells;
                if (row != null && cells != null)
                {
                    if (row.RowNum == 0)
                    {
                        for (int i = 0; i < cells.Count; i++)
                        {
                            var txt = cells[i].ToString().Trim();
                            switch (txt)
                            {
                                case "职业": dict.Add("职业", i); break;
                                case "所属行业一级分类": dict.Add("所属行业一级分类", i); break;
                                case "所属行业二级分类": dict.Add("所属行业二级分类", i); break;
                                case "介绍": dict.Add("介绍", i); break;
                            }
                        }
                    }  //第一行

                    else //数据行
                    {

                        ICell cellName = row.GetCell(dict["职业"]);
                        ICell cellPname = row.GetCell(dict["所属行业一级分类"]);
                        ICell cellCname = row.GetCell(dict["所属行业二级分类"]);
                        ICell cellJieShao = row.GetCell(dict["介绍"]);

                        if (cellName != null)
                        {
                            string name=cellName.ToString().Trim();
                            if (name != "")
                            {
                                Job job = new Job();
                                job.Name = name;

                                if (cellPname != null)
                                {
                                    job.SubjectName = cellPname.ToString().Trim();
                                }

                                if (cellCname != null)
                                {
                                    job.CategoryName = cellCname.ToString().Trim();
                                }

                                if (cellJieShao != null)
                                {
                                    job.Introduction = cellJieShao.ToString().Trim();
                                }
                                jobList.Add(job);
                            }
                        }
                    }

                }
            }


            foreach (var item in jobList)
            {
                AddJob(item);
            }

        }




        /// <summary>
        /// 添加专业
        /// </summary>
        /// <param name="majorCategory"></param>
        /// <returns></returns>
        public int AddMajor(Major major)
        {
            string sql = @"
DECLARE @pid INT,@cid INT
IF(EXISTS(SELECT * FROM dbo.MajorCategory WHERE Name=@SubjectName AND ParentId=0))
BEGIN
	SET @pid=(SELECT TOP 1 CategoryId FROM dbo.MajorCategory WHERE Name=@SubjectName AND ParentId=0)
END
ELSE
BEGIN
    IF(ISNULL(@pid,0)=0)
    BEGIN
	    SET @pid=0
    END
    ELSE
    BEGIN
	    	INSERT INTO [MajorCategory]( 
            [Name], [ParentId], [SortOrder] 
        ) VALUES  ( 
            @SubjectName, 0, 0
        )
	    SET @pid=@@IDENTITY
    END

END

IF(EXISTS(SELECT * FROM dbo.MajorCategory WHERE Name=@CategoryName AND ParentId!=0))
BEGIN
	SET @cid=(SELECT TOP 1 CategoryId FROM dbo.MajorCategory WHERE Name=@CategoryName AND ParentId!=0)
END
ELSE
BEGIN
    IF(ISNULL(@pid,0)=0)
    BEGIN
	    SET @cid=0
    END
    ELSE
    BEGIN
	    INSERT INTO [MajorCategory]( 
            [Name], [ParentId], [SortOrder] 
        ) VALUES  ( 
            @CategoryName, @pid, 0
        )
	    SET @cid=@@IDENTITY
    END
	
END

IF(NOT EXISTS(SELECT * FROM dbo.Major WHERE Name=@Name))
BEGIN
	INSERT INTO dbo.Major
	        ( Name ,
	          MajorType ,
	          SubjectId ,
	          CategoryId ,
	          MajorCode ,
	          Introduction ,
	          HotMajor
	        )
	VALUES  ( @Name , -- Name - nvarchar(100)
	          @MajorType , -- MajorType - tinyint
	          @pid , -- SubjectId - int
	          @cid , -- CategoryId - int
	          @MajorCode , -- MajorCode - nvarchar(20)
	          @Introduction , -- Introduction - ntext
	          0  -- HotMajor - bit
	        )
END
";
            object[,] par = 
            {
                {"@SubjectName",major.SubjectName},
                {"@CategoryName",major.CategoryName},
                {"@Name",major.Name},
                {"@MajorType",major.MajorType},
                {"@MajorCode",major.MajorCode},
                {"@Introduction",major.Introduction}
                

            };
            object value= SQLHelper2.GetSingle(sql,par);

            return Convert.ToInt32(value);
        }


        /// <summary>
        /// 添加职业
        /// </summary>
        /// <param name="job"></param>
        public void AddJob(Job job)
        {
            string sql = @"

DECLARE @pid INT,@cid INT
IF(EXISTS(SELECT * FROM dbo.JobCategory WHERE Name=@SubjectName AND ParentId=0))
BEGIN
	SET @pid=(SELECT TOP 1 CategoryId FROM dbo.JobCategory WHERE Name=@SubjectName AND ParentId=0)
END
ELSE
BEGIN
    IF(ISNULL(@pid,0)=0)
    BEGIN
	    SET @pid=0
    END
    ELSE
    BEGIN
	    INSERT INTO [JobCategory]( 
            [Name], [ParentId], [SortOrder] 
        ) VALUES  ( 
            @SubjectName, 0, 0
        )
	    SET @pid=@@IDENTITY
    END
	
END

IF(EXISTS(SELECT * FROM dbo.JobCategory WHERE Name=@CategoryName AND ParentId!=0))
BEGIN
	SET @cid=(SELECT TOP 1 CategoryId FROM dbo.JobCategory WHERE Name=@CategoryName AND ParentId!=0)
END
ELSE
BEGIN
    IF(ISNULL(@pid,0)=0)
    BEGIN
	    SET @cid=0
    END
    ELSE
    BEGIN
	    INSERT INTO [JobCategory]( 
            [Name], [ParentId], [SortOrder] 
        ) VALUES  ( 
            @CategoryName, @pid, 0
        )
	    SET @cid=@@IDENTITY
    END
END

IF(NOT EXISTS(SELECT * FROM dbo.Job WHERE Name=@Name))
BEGIN
	INSERT INTO dbo.Job
	        ( Name ,
	          TopCategoryId ,
	          SecondCategoryId ,
	          Introduction ,
	          HotJob
	        )
	VALUES  ( @Name , -- Name - nvarchar(100)
	          @pid, -- TopCategoryId - int
	          @cid, -- SecondCategoryId - int
	          @Introduction , -- Introduction - ntext
	          0  -- HotJob - bit
	        )
END
";

            object[,] par = 
            {
                {"@SubjectName",job.SubjectName},
                {"@CategoryName",job.CategoryName},
                {"@Name",job.Name},
                {"@Introduction",job.Introduction}
                

            };

            SQLHelper2.ExecuteSql(sql, par);
        }

    }
}