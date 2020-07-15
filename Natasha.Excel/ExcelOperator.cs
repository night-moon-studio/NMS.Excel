using Natasha.CSharp;
using Natasha.Excel;
using NPOI.SS.Formula.Functions;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;

namespace System
{
    public class ExcelOperator
    {

        public static void ConfigWritter<TEntity>(Dictionary<string, string> mappers, params string[] ignores)
        {
            ExcelOperator<TEntity>.CreateWriteDelegate(mappers, ignores);
        }
        public static void ConfigReader<TEntity>(params (string Key, string Value)[] mappers)
        {
            ExcelOperator<TEntity>.CreateReadDelegate(mappers);
        }

        public static void WriteToFile<TEntity>(string filePath, IEnumerable<TEntity> entities, int sheetPage = 0)
        {
            ExcelOperator<TEntity>.WriteToFile(filePath, entities, sheetPage);
        }
        public static IEnumerable<TEntity> FileToEntities<TEntity>(string filePath, int sheetPage = 0)
        {
            return ExcelOperator<TEntity>.FileToEntities(filePath, sheetPage);
        }
    }


    public class ExcelOperator<TEntity>
    {

        private static ImmutableDictionary<string, string> _mappers;
        private static Dictionary<string, int> _fields;
        private static Action<RowOperator, IEnumerable<TEntity>> Writter;
        private static Func<RowOperator, int, int[], IEnumerable<TEntity>> Reader;

        public static Action<RowOperator, IEnumerable<TEntity>> CreateWriteDelegate(Dictionary<string, string> mappers, params string[] ignores)
        {

            _mappers = ImmutableDictionary.CreateRange(mappers);
            //给字段排序
            int index = 0;
            foreach (var item in _mappers)
            {
                _fields[item.Key] = index;
                index += 1;
            }

            HashSet<string> ignorSets = new HashSet<string>(ignores);
            StringBuilder excelBody = new StringBuilder();
            StringBuilder excelHeader = new StringBuilder();
            excelHeader.Append("var col = row.Columns;");
            excelBody.Append(@"foreach(var item in arg2){");
            excelBody.Append($"arg1 = arg1.Next;");
            excelBody.Append($"col = row.Columns;");
            foreach (var item in mappers)
            {

                if (!ignorSets.Contains(item.Key))
                {
                    excelHeader.Append($"col.SetValue(\"{item.Key}\");");
                    excelHeader.Append("col = col.Next;");
                    excelBody.Append($"col.SetValue(item.{item.Value});");
                    excelBody.Append("col = col.Next;");
                }

            }
            excelBody.Append("}");
            excelHeader.Append(excelBody);
            return Writter = NDelegate
                .UseDomain(typeof(TEntity).GetDomain())
                .Action<RowOperator, IEnumerable<TEntity>>(excelHeader.ToString());
        }
        public static Func<RowOperator, int, IEnumerable<TEntity>> CreateReadDelegate(params (string Key, string Value)[] mappers)
        {

            StringBuilder excelBody = new StringBuilder();
            excelBody.Append($"var list = new List<{typeof(TEntity).GetDevelopName()}>(arg.Count);");
            excelBody.Append(@"for(int i=1;i<arg2;i+=1){");
            excelBody.Append("arg1 = arg1[i];");
            excelBody.Append("var columns = row[i].Columns;");
            excelBody.Append($"var instance = new {typeof(TEntity).GetDevelopName()}();");
            for (int i = 0; i < mappers.Length; i++)
            {
                
                var prop = typeof(TEntity).GetProperty(mappers[i].Value);
               
                if (prop.PropertyType == typeof(string))
                {
                    excelBody.Append($"instance.{mappers[i].Value} = columns[{i}].StringValue;");
                }
                else if (prop.PropertyType == typeof(DateTime))
                {
                    excelBody.Append($"instance.{mappers[i].Value} = columns[{i}].DateValue;");
                }
                else if (prop.PropertyType != typeof(bool))
                {
                    excelBody.Append($"instance.{mappers[i].Value} = Convert.To{prop.PropertyType.Name}(columns[{i}].NumValue);");
                }
                else
                {
                    excelBody.Append($"instance.{mappers[i].Value} = columns[{i}].BoolValue;");
                }
                 
            }
            excelBody.Append("list.Add(instance);");
            excelBody.Append("}");
            excelBody.Append("return list;");
            return Reader = NDelegate
                .UseDomain(typeof(TEntity).GetDomain())
                .Func<RowOperator, int, int[], IEnumerable<TEntity>>(excelBody.ToString());
        }


        public static void WriteToFile(string filePath, IEnumerable<TEntity> entities,int sheetPage)
        {

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            using (var builder = new ExcelBuilder(filePath))
            {
                Writter(builder[sheetPage], entities);
                builder.Save();
            }
            

        }


        public static IEnumerable<TEntity> FileToEntities(string filePath,int sheetPage)
        {

            using (var builder = new ExcelBuilder(filePath))
            {

                var indexs = new int[_mappers.Count];
                var row = builder[sheetPage];
                var columns = row.Columns;
                for (int i = 0; i < row.Count; i+=1)
                {

                    if (_mappers.TryGetValue(columns.StringValue, out var field))
                    {
                        if (_fields.TryGetValue(field, out var value))
                        {
                            indexs[value] = i;
                        }
                    }

                }
                return Reader(row.Next, builder.Count, indexs);

            }

        }

    }

}
