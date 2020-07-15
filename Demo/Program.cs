using System;
using System.Collections.Generic;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            AssemblyDomain.Init();
            List<Student> students = new List<Student>();
            for (int i = 0; i < 5; i++)
            {
                students.Add(new Student()
                {
                    Age = i,
                    Description = "This is 描述！",
                    Name = "test" + i,
                    Sex = i % 2 == 0 ? "男" : "女"
                });
            }


            ExcelOperator.ConfigWritter<Student>(new Dictionary<string, string> {

                { "名字","Name"},
                { "年龄","Age"},
                { "性别","Sex"},
                { "描述","Description"},

            }, "描述");

            ExcelOperator.WriteToFile("1.xlsx", students);


            ExcelOperator.ConfigReader<Student>(("名字", "Name"), ("年龄", "Age"), ("性别", "Sex"));
            var list = ExcelOperator.FileToEntities<Student>("1.xlsx");
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
    }
}
