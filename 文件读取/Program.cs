using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Collections.Generic;

namespace 文件读取
{
    class Program
    {
        static void Main(string[] args)
        {
            //var sheetName = "基础设施";
            //var folder = string.Format(@"F:\大四下\助管\同路人交科论文集\{0}", sheetName);
            //var workBook = new HSSFWorkbook();
            //var table = workBook.CreateSheet(sheetName);
            //DirectoryInfo directoryInfos = new DirectoryInfo(folder);
            //int i = 0;
            //foreach (var directoryInfo_item in directoryInfos.GetFiles())
            //{
            //    var row = table.CreateRow(i);
            //    var cell = row.CreateCell(1);
            //    cell.SetCellValue(directoryInfo_item.Name.Remove(directoryInfo_item.Name.LastIndexOf('.')));//读取文件名，并将其写入excel
            //    i++;
            //}
            //using (var fs = File.OpenWrite(string.Format(@"F:\大四下\助管\同路人交科论文集\{0}.xls",sheetName)))
            //{
            //    workBook.Write(fs);   //向打开的这个xls文件中写入mySheet表并保存。
            //    Console.WriteLine("生成成功");
            //}

            var sheetName = "基础设施";
            DataTable dt=ExcelToDataTable(@"F:\大四下\助管\同路人交科论文集\论文分类（没问题版）.xlsx",sheetName,true);

            List<string> FileNewNames = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                //新的文件夹名存进列表
                FileNewNames.Add(row["年份"].ToString() + "_" + row["作品名称"].ToString() + "_" + row["指导教师"].ToString() + "_" + row["第一完成人"].ToString());
            }
            Console.WriteLine("------------------------------\n excel读取结束");


            var NewFolder =string.Format(@"F:\大四下\助管\同路人交科论文集\{0}\",sheetName+"demo") ;
            if (new DirectoryInfo(NewFolder).Exists)//如果新的路径存在
            {
                DeleteFolder(NewFolder);//清空新建的文件夹
            }
            else
            {
                new DirectoryInfo(NewFolder).Create();
            }

            Console.WriteLine("---------------------------\n" + "文件清空结束");
            var OldBasicFolder =string.Format(@"F:\大四下\助管\同路人交科论文集\{0}\",sheetName) ;
            DirectoryInfo directoryInfos = new DirectoryInfo(OldBasicFolder);
            int i = 0;
            foreach(var item in directoryInfos.GetFiles())
            {
                Console.WriteLine( FileNewNames[i]+ "正在复制文件");
                try
                {
                    item.CopyTo(NewFolder + FileNewNames[i] + item.Extension);
                    Console.WriteLine(FileNewNames[i] + "已经复制完成");
                }
                catch(Exception e)
                {
                    Console.WriteLine(e.ToString());
                }

                i++;
            }
            Console.WriteLine("---------------------------\n"+"文件复制结束");
            Console.ReadKey();
        }
        //读入excel为datatable
        public static DataTable ExcelToDataTable(string filePath, string sheetName,bool isColumnName)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        //sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet  
                        sheet = workbook.GetSheet(sheetName);
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数  
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行  
                                int cellCount = firstRow.LastCellNum;//列数  

                                //构建datatable的列  
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取  
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }

                                //填充行  
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null) continue;

                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)  
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }
        //删除判断文件夹是否存在，若不存在，新建，否则删除当前文件夹下所有文件
        public static void DeleteFolder(string path)
        {
            if (!new DirectoryInfo(path).Exists)//如果新的路径不存在
            {
                new DirectoryInfo(path).Create();
            }
            else
            {
                foreach (string d in Directory.GetFileSystemEntries(path))
                {
                    if (File.Exists(d))
                    {
                        FileInfo fi = new FileInfo(d);
                        if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                            fi.Attributes = FileAttributes.Normal;
                        File.Delete(d);//直接删除其中的文件  
                    }
                    else
                    {
                        DirectoryInfo d1 = new DirectoryInfo(d);
                        if (d1.GetFiles().Length != 0)
                        {
                            DeleteFolder(d1.FullName);////递归删除子文件夹
                        }
                        Directory.Delete(d);
                    }
                }
            }
        }
    }
    
}

