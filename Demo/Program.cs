using AzulX.NPOI;
using System;
using System.Collections.Generic;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
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

            
            using (ExcelFile file = new ExcelFile(AppDomain.CurrentDomain.BaseDirectory + "1.xlsx"))
            {
                ExcelStyle style = ExcelStyle.Create(file);
                file.LoadTemplate("1.txt")
                .UseTemplate("Test2")
                .FillHeader(style.Header())
                .FillCollection(students)
                .Save();
            }

            using (ExcelFile file = new ExcelFile(AppDomain.CurrentDomain.BaseDirectory + "2.xlsx"))
            {
                file.Select(1);

                //阶梯
                file.CurrentCell(0.ToString());
                for (int i = 1; i < 10; i+=1)
                {
                    file
                        .NextRow(false)
                        .NextCell(i.ToString());
                }

                file.Save();
            }

            using (ExcelFile file = new ExcelFile(AppDomain.CurrentDomain.BaseDirectory + "2.xlsx"))
            {
                file.Select(1);

                //阶梯
                Console.WriteLine(file.StringValue);
                for (int i = 1; i < 10; i += 1)
                {
                    file
                        .NextRow(false);
                    Console.WriteLine(file.NextStringValue);
                }
            }
            Console.ReadKey();
        }
    }
}
