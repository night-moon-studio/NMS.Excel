using System;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            //List<Student> students = new List<Student>();
            //for (int i = 0; i < 5; i++)
            //{
            //    students.Add(new Student()
            //    {
            //        Age = i,
            //        Description = "This is 描述！",
            //        Name = "test" + i,
            //        Sex = i % 2 == 0 ? "男" : "女"
            //    });
            //}

            
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

            using (ExcelOperator file = new ExcelOperator(AppDomain.CurrentDomain.BaseDirectory + "2.xlsx"))
            {
                var sheet = file[1];

                //阶梯
                Console.WriteLine(sheet[0][0].StringValue);
                for (int i = 1; i < 10; i += 1)
                {
                    Console.WriteLine(sheet[i][i].StringValue);
                }
            }
            Console.ReadKey();
        }
    }
}
