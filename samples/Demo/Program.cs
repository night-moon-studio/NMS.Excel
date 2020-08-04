using QE.CMS.Entities;
using System;
using System.Collections.Generic;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {

            NatashaComponentRegister.RegistDomain<NatashaAssemblyDomain>();
            NatashaComponentRegister.RegistCompiler<NatashaCSharpCompiler>();
            NatashaComponentRegister.RegistSyntax<NatashaCSharpSyntax>();

            NSucceedLog.Enabled = true;
            //List<Student> students = new List<Student>();
            //for (int i = 0; i < 5; i++)
            //{
            //    students.Add(new Student()
            //    {
            //        Age = i,
            //        Description = "This is 描述！",
            //        Name = "test" + i,
            //        Sex = i % 2 == 0 ? "男" : "女",
            //        Flag = i % 2 == 0 ? (short?)1 : null,
            //        UpdateTime = i % 2 == 0 ? (short?)1000 : null,
            //    });
            //}


            var dict = new Dictionary<string, string> {
    {"Id","分类Id"},
    { "Pid","父节点ID"},
    {"CategoryName","分类名称"},
    {"CreateTime", "创建时间"},
    {"Status", "状态"},
    { "UpdateTime","修改时间"},
    { "ViewCount","浏览量"}
            };

            ExcelOperator.SetWritterMapping<QuestionCategory>(dict, "描述");
            //ExcelOperator.WriteToFile("1.xlsx", students);

            ExcelOperator.SetReaderMapping<QuestionCategory>(dict);
            var list = ExcelOperator.FileToEntities<QuestionCategory>("2020-08-04 09-08 QuestionCategory.xlsx");
            //var a = ExcelOperator<QuestionCategory>.Get("2020-08-04 09-08 QuestionCategory.xlsx", 0);
            //Invoke(a.Item1, a.Item2);
           Console.ReadKey();



            //using (ExcelFile file = new ExcelFile(AppDomain.CurrentDomain.BaseDirectory + "1.xlsx"))
            //{
            //    ExcelStyle style = ExcelStyle.Create(file);
            //    file.LoadTemplate("1.txt")
            //    .UseTemplate("Test2")
            //    .FillHeader(style.Header())
            //    .FillCollection(students)
            //    .Save();
            //}

            //using (ExcelFile file = new ExcelFile(AppDomain.CurrentDomain.BaseDirectory + "2.xlsx"))
            //{
            //    file.Select(1);

            //    //阶梯
            //    file.CurrentCell(0.ToString());
            //    for (int i = 1; i < 10; i+=1)
            //    {
            //        file
            //            .NextRow(false)
            //            .NextCell(i.ToString());
            //    }

            //    file.Save();
            //}

            //using (ExcelOperator file = new ExcelOperator(AppDomain.CurrentDomain.BaseDirectory + "2.xlsx"))
            //{
            //    var sheet = file[1];

            //    //阶梯
            //    Console.WriteLine(sheet[0][0].StringValue);
            //    for (int i = 1; i < 10; i += 1)
            //    {
            //        Console.WriteLine(sheet[i][i].StringValue);
            //    }
            //}
           
        }

        public static System.Collections.Generic.IEnumerable<QE.CMS.Entities.QuestionCategory> Invoke(NPOI.SS.UserModel.ISheet arg1, System.Int32[] arg2)
        {
            var list = new List<QE.CMS.Entities.QuestionCategory>(arg1.LastRowNum);
            var tempNullableValue = String.Empty;
            for (int i = 1; i <= arg1.LastRowNum; i += 1)
            {
                var row = arg1.GetRow(i);
                var instance = new QE.CMS.Entities.QuestionCategory();
                instance.CategoryName = row.GetCell(arg2[2]).StringCellValue;
                instance.ViewCount = Convert.ToInt64(row.GetCell(arg2[6]).NumericCellValue);
                instance.Id = Convert.ToInt64(row.GetCell(arg2[0]).NumericCellValue);
                var a = row.GetCell(arg2[5]);
                tempNullableValue = row.GetCell(arg2[5]).StringCellValue;
                if (string.IsNullOrEmpty(tempNullableValue)) { instance.UpdateTime = null; }
                else { instance.UpdateTime = Convert.ToInt64(tempNullableValue); }
                instance.Status = Convert.ToInt16(row.GetCell(arg2[4]).NumericCellValue);
                instance.CreateTime = Convert.ToInt64(row.GetCell(arg2[3]).NumericCellValue);
                instance.Pid = Convert.ToInt64(row.GetCell(arg2[1]).NumericCellValue);
                list.Add(instance);
            }
            return list;
        }

    }
}
