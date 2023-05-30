using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace extension0529.Extension
{
   

        public static class ExcelExtensions
        {

            /// <summary>
            /// 獲得DisplayName所設定的名稱
            /// </summary>
            /// <param name="memberInfo"></param>
            /// <returns>名稱值</returns>
            /// 根據 MemberInfo 物件是否具有 DisplayNameAttribute 屬性來獲取其顯示名稱。如果找到了該屬性，則返回屬性的值；如果沒有找到該屬性，則返回空字串。
            private static string GetDisplayName(this MemberInfo memberInfo)
            {
                var titleName = string.Empty;
                var attribute = memberInfo.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault();
                if (attribute != null)
                {
                    titleName = (attribute as DisplayNameAttribute).DisplayName;
                }
                else
                {
                    //titleName = memberInfo.Name;
                }
                return titleName;
            }
            /// <summary>
            /// 獲得屬性名displayName的特性名
            /// </summary>
            /// <param name="type">類型</param>
            /// <returns>特性名的值</returns>
            /// 根據指定類型的屬性，遍歷每個屬性並調用其 GetDisplayName 方法，取得屬性的顯示名稱，並將這些顯示名稱存放在一個清單中返回。
            private static List<string> GetPropertyDisplayNames(this Type type)
            {
                var titleList = new List<string>();
                var propertyInfos = type.GetProperties();
                foreach (var propertyInfo in propertyInfos)
                {
                    var titleName = propertyInfo.GetDisplayName();
                    titleList.Add(titleName);
                }
                return titleList;
            }

            private static void SetBorderStyle(int startRow, int endRow, int startCol, int endCol, IWorkbook workBook, ISheet workSheet)
            { //設定工作表中儲存格的邊框樣式
              //for (int r = startRow; r <= endRow; r++)
              //{
              //    IRow row = workSheet.GetRow(r);
              //    for (int c = startCol; c <= endCol; c++)
              //    {
              //        ICellStyle style = workBook.CreateCellStyle();
              //        style.BorderBottom = BorderStyle.THIN;
              //        style.BorderLeft = BorderStyle.THIN;
              //        style.BorderRight = BorderStyle.THIN;
              //        style.BorderTop = BorderStyle.THIN;
              //        style.Alignment = HorizontalAlignment.CENTER;
              //        ICell cell = row.GetCell(c);
              //        cell.CellStyle = style;
              //        workSheet.AutoSizeColumn(c);
              //    }
              //}
                for (int r = startRow; r <= endRow; r++)
                {
                    IRow row = workSheet.GetRow(r);

                    for (int c = startCol; c <= endCol; c++)
                    {
                        //設定 ICellStyle 的邊框樣式，將 BorderStyle 設定為 Thin，即細邊框
                        ICellStyle style = workBook.CreateCellStyle();
                        style.BorderBottom = BorderStyle.Thin;
                        style.BorderLeft = BorderStyle.Thin;
                        style.BorderRight = BorderStyle.Thin;
                        style.BorderTop = BorderStyle.Thin;


                        if (r == 1)  //當前處理的是第一行
                        {
                            style.Alignment = HorizontalAlignment.Center;
                            //var color = Color.FromArgb(80, 124, 209);
                            //style.FillPattern = FillPatternType.SOLID_FOREGROUND;
                            //style.FillForegroundColor = HSSFColor.RoyalBlue.Index;

                            //文字水平對齊為居中，並設定粗體字型、字體顏色為白色、字體大小為 12
                            IFont font = workBook.CreateFont();
                            font.Boldweight = (short)400;
                            font.Color = HSSFColor.White.Index;
                            font.FontHeightInPoints = (short)12;
                            style.SetFont(font);

                        }
                        if (r > 1) //處理的是其他行
                        {
                            //var color = Color.FromArgb(239, 243, 251);
                            //文字自動換行、垂直對齊為居中
                            style.WrapText = true;
                            style.VerticalAlignment = VerticalAlignment.Center;
                            if (r % 2 == 0)
                            {
                                //style.FillPattern = FillPatternType.SOLID_FOREGROUND;
                                //style.FillForegroundColor = HSSFColor.LIGHT_CORNFLOWER_BLUE.index;
                            }
                        }

                        //將設定好的 ICellStyle 賦值給當前儲存格，使用 cell.CellStyle = style。
                        ICell cell = row.GetCell(c);
                        cell.CellStyle = style;
                        workSheet.AutoSizeColumn(c);
                        // 自動調整儲存格所在的列的寬度
                    }
                }
            }

        /// <summary>
        /// for Report1 use
        /// one class has lots list ,then you got sheets
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        /// 將資料匯出到 Excel 檔案中
        //單一的物件
        public static string ExportExcel<T>(this T dataList, string fileName)
            {   //(要匯出到 Excel 的資料列表,匯出的 Excel 檔案的檔名)

                //Create workbook
                var datatype = typeof(T);

                IWorkbook workbook; //= new HSSFWorkbook();
                var extension = Path.GetFileNameWithoutExtension(fileName);
                if (extension == "xls")
                { //根據 fileName 的擴展名判斷要使用的工作簿類型。如果擴展名為 "xls"
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    workbook = new XSSFWorkbook();
                }

                var test = datatype.GetPropertyDisplayNamesMap();
                var subClassMapName = datatype.GetClassNameList();
                Assembly a = Assembly.Load("extension0529");

            foreach (var item in subClassMapName)
                {
                    var dyType = a.GetType(item.Value.FullName);
                    var worksheet = workbook.CreateSheet(string.Format("{0}", item.Value.DisplayName));
                    var row = worksheet.CreateRow(0);
                    var titleListMap = dyType.GetPropertyDisplayNamesMap();


                    var cellIdx = 0;
                    foreach (var subItem in titleListMap)
                    {
                        var cell = row.CreateCell(cellIdx);


                        var title = subItem.Value;
                        //cell.CellStyle = fontStyle;

                        cell.SetCellValue(title.Name);

                        cellIdx++;
                    }


                    var ttt = datatype.GetProperty(item.Key);
                    var t_type = ttt.GetType();
                    var sub_dataList = (IList)ttt.GetValue(dataList, null);

                    //Insert data values
                    InsertDataValues(sub_dataList, workbook, worksheet, titleListMap);
                    //自動篩選
                    var endRange = IntToAlphabet.IndexToColumn(titleListMap.Count) + "1";
                    var headerRange = CellRangeAddress.ValueOf("A1:" + endRange);
                    worksheet.SetAutoFilter(headerRange);

                    //自動設寬
                    for (int i = 1; i < titleListMap.Count + 1; i++)
                    {
                        worksheet.AutoSizeColumn(i);
                    }
                }


                //Save file
                var savePath = Path.Combine(Path.GetTempPath(), fileName);
                FileStream file = new FileStream(savePath, FileMode.Create);
                workbook.Write(file);
                file.Close();

                return savePath;


                //將資料列表 dataList 匯出到 Excel 檔案中。
                //它會根據泛型類型 T 的屬性和子類，創建工作表並寫入標題和資料。
                //最後，將工作簿的內容保存到指定的檔案中
            }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <param name="fileName">檔名: myExcel.xls</param>
        /// <returns>filePath</returns>
        //傳遞一個包含多個元素的集合
        public static string ExportExcel<T>(this IEnumerable<T> dataList, string fileName, string sheetName)
            {//(資料,excel檔名,創建的工作表的名稱)
             //Create workbook
                var datatype = typeof(T); //得泛型型別 T 的類型，並賦值給 datatype

                var extension = Path.GetFileNameWithoutExtension(fileName);

                IWorkbook workbook; //= new HSSFWorkbook();

                if (extension == "xls")
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    workbook = new XSSFWorkbook();
                }

                //創建一個工作表，名稱為 sheetName
                //在工作表中創建一行，第一行用於標題
                var worksheet = workbook.CreateSheet(string.Format("{0}", sheetName));
                //Insert titles
                var row = worksheet.CreateRow(0);
                var titleListMap = datatype.GetPropertyDisplayNamesMap();
                var wrapRowCount = 0;

                var cellIdx = 0;
                //使用迴圈將標題寫入第一行的每個儲存格
                foreach (var item in titleListMap)
                {
                    var cell = row.CreateCell(cellIdx);


                    var title = item.Value;
                    //cell.CellStyle = fontStyle;

                    cell.SetCellValue(title.Name);

                    cellIdx++;
                }

                //Insert data values
                InsertDataValues(dataList.ToList(), workbook, worksheet, titleListMap);


                //自動篩選
                var endRange = IntToAlphabet.IndexToColumn(titleListMap.Count) + "1";
                var headerRange = CellRangeAddress.ValueOf("A1:" + endRange);
                worksheet.SetAutoFilter(headerRange);

                //自動設寬
                for (int i = 1; i < titleListMap.Count + 1; i++)
                {
                    worksheet.AutoSizeColumn(i);
                }

                //Save file
                var savePath = Path.Combine(Path.GetTempPath(), fileName);
                FileStream file = new FileStream(savePath, FileMode.Create);
                workbook.Write(file);
                file.Close();

                return savePath;
            }

            public static Dictionary<string, ClassNameValue> GetClassNameList(this Type type)
            {
                var result = new Dictionary<string, ClassNameValue>();
                var propertyInfos = type.GetProperties();
                foreach (var item in propertyInfos)
                {
                    var titleName = item.GetDisplayName();
                    Regex r1 = new Regex(@"(\[\[)([^,]+)");

                    // C
                    // Match the input and write results
                    Match match = r1.Match(item.PropertyType.FullName);
                    if (match.Success)
                    {
                        string match_value = match.Groups[0].Value.Replace(@"[[", string.Empty);
                        if (string.IsNullOrEmpty(titleName))
                        {
                            titleName = item.Name;
                        }
                        result.Add(item.Name, new ClassNameValue { DisplayName = titleName, FullName = match_value });
                    }

                }


                return result;

                //獲取類型的屬性名稱和對應的類名，並將其存儲在一個字典中返回
            }

            public class ClassNameValue
            {
                public string DisplayName { get; set; }
                public string FullName { get; set; }
            }

            /// <summary>
            /// 取得屬性的顯示名稱 (字典類)
            /// Note:因類名稱規定不能重覆，這邊就不再判斷會不會加入同樣的key
            /// 調整形態判斷規則:
            /// 如果沒有 指定特殊形態 則使用系統內建的
            /// </summary>
            /// <param name="type"></param>
            /// <returns></returns>
            public static Dictionary<string, DisplayItem> GetPropertyDisplayNamesMap(this Type type)
            {
                var titleListMap = new Dictionary<string, DisplayItem>();
                var propertyInfos = type.GetProperties();
                var isEnableFormat = false;
                foreach (var propertyInfo in propertyInfos)
                {
                    var titleName = propertyInfo.GetDisplayName();
                    //default
                    var cellType = CellType.String;

                    var excelType = propertyInfo.GetCustomAttributes(typeof(ExcelDataTypeAttribute), true);
                    if (excelType.Length == 1)
                    {
                        var checkType = (ExcelDataTypeAttribute)excelType.FirstOrDefault();
                        if (checkType.DataType == DataType.Currency)
                        {
                            cellType = CellType.Numeric;
                            isEnableFormat = true;
                        }
                        else if (checkType.DataType == DataType.DateTime)
                        {
                            cellType = CellType.Formula;
                        }
                    }
                    else
                    {
                        var proType = propertyInfo.PropertyType;
                        //沒有強制指定Column Type則使用預設所偵測到的反射型態
                        switch (proType.Name)
                        {

                            case "Int32":
                            case "Float":
                            case "Decimal":
                                cellType = CellType.Numeric;
                                break;
                        }
                    }

                    if (string.IsNullOrEmpty(titleName))
                        continue;

                    titleListMap.Add(propertyInfo.Name, new DisplayItem { Name = titleName, CellType = cellType, IsEnableFormat = isEnableFormat });
                }

                return titleListMap;
            }

            /// <summary>
            /// 塞資料用的 not pretty
            /// 供上面產出Excel Data 共用
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="dataList"></param>
            /// <param name="workbook"></param>
            /// <param name="worksheet"></param>
            /// <param name="titleListMap"></param>
            private static void InsertDataValues(IList dataList, IWorkbook workbook, ISheet worksheet, Dictionary<string, DisplayItem> titleListMap)
            { //將數據值插入到 Excel 工作表中
              //(要插入的數據列表, Excel 工作簿,要插入數據的工作表,屬性名稱,對象的字典映射)

                //Insert data values
                for (int i = 1; i < dataList.Count + 1; i++)
                { //從索引 1 開始（因為索引 0 是標題行），獲取每個數據對象。
                    var tmpRow = worksheet.CreateRow(i); //行
                    var valueList = dataList[i - 1].GetPropertyValues(titleListMap); //獲取數據值列表

                    for (int j = 0; j < valueList.Count; j++)
                    {
                        var rowCell = tmpRow.CreateCell(j);
                        var valueItem = valueList[j];
                        var tempValue = valueItem.Name;

                        switch (valueItem.CellType) //根據 CellType 設置單元格值
                        {
                            case CellType.Numeric: //根據 valueItem.CellType
                                if (string.IsNullOrEmpty(tempValue) || tempValue == "------" || tempValue == "--")
                                {
                                    rowCell.SetCellValue("");
                                }
                                else
                                {
                                    var intValue = 0.00;
                                    var flag = double.TryParse(tempValue.Replace(",", string.Empty), out intValue);
                                    if (flag)
                                        rowCell.SetCellValue(intValue);
                                    else
                                        rowCell.SetCellValue(tempValue);
                                }

                                if (valueItem.IsEnableFormat)
                                {
                                    var cellStyle = workbook.CreateCellStyle();
                                    var format = workbook.CreateDataFormat();
                                    cellStyle.DataFormat = format.GetFormat("#,##0.000");
                                    rowCell.CellStyle = cellStyle;
                                }

                                break;
                            case CellType.Formula:
                                if (!string.IsNullOrEmpty(tempValue))
                                {
                                    DateTime dateValue;
                                    var flag = DateTime.TryParseExact(tempValue, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue);
                                    if (flag)
                                        rowCell.SetCellValue(dateValue);
                                    else
                                        rowCell.SetCellValue(tempValue);
                                }
                                else
                                {
                                    rowCell.SetCellValue(tempValue);
                                }
                                var cellStyle2 = workbook.CreateCellStyle();
                                var format2 = workbook.CreateDataFormat();
                                cellStyle2.DataFormat = format2.GetFormat("yyyy/m/d");
                                rowCell.CellStyle = cellStyle2;
                                break;
                            default:
                                rowCell.SetCellValue(tempValue);
                                break;

                                //將數據列表插入到 Excel 工作表中，並根據屬性的類型和設定進行相應的數值格式化
                        }

                    }

                }
            }

            /// <summary>
            /// 將T類型的公共屬性全部轉換成字符串
            /// </summary>
            /// <typeparam name="T">T類型</typeparam>
            /// <param name="data">需要轉換的對象</param>
            /// <returns>公共類型的屬性字符串集合</returns>
            private static List<string> GetPropertyValues<T>(this T data)
            {
                var properValues = new List<string>();
                var properInfos = data.GetType().GetProperties();
                foreach (var properInfoItem in properInfos)
                {
                    //var value = properInfoItem.GetValue(data, null).ToString();
                    var rowValue = properInfoItem.GetValue(data, null) != null ? properInfoItem.GetValue(data, null).ToString() : "";
                    properValues.Add(rowValue);
                }
                return properValues;

                ///從給定對象中獲取屬性的值列表。
                ///它使用反射遍歷對象的屬性信息，並
                ///將每個屬性的值轉換為字符串後添加到列表中，最後返回該列表
            }

            /// <summary>
            /// 取得屬性值
            /// Note:原作者取值的方法，會有排序上的錯亂。
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="data"></param>
            /// <param name="type"></param>
            /// <returns></returns>
            public static List<RowItem> GetPropertyValues<T>(this T data, Dictionary<string, DisplayItem> columnMap)
            {
                var propertyValues = new List<RowItem>();

                var sourceData = data.GetType();
                foreach (var item in columnMap)
                {

                    var propertyValue = sourceData.GetProperty(item.Key).GetValue(data, null);
                    var rowValue = propertyValue != null ? propertyValue.ToString() : "";

                    propertyValues.Add(new RowItem { Name = rowValue, CellType = item.Value.CellType, IsEnableFormat = item.Value.IsEnableFormat });//排除null的情況
                }

                return propertyValues;

                //從給定對象中獲取屬性的值列表，同時根據給定的 columnMap 字典對屬性進行映射。
                //它使用反射遍歷屬性映射字典中的每個屬性，
                //並根據對象中對應屬性的值創建 RowItem 對象，然後將其添加到列表中，最後返回該列表
            }
        }

        /// <summary>
        /// 自定義欄位名稱
        /// </summary>
        public class ColNames
        {
            public string RootItem { get; set; } //根項目
            public List<string> ChildrenItem { get; set; } //子項目的字串清單
            public int Length { get; set; } //子項目清單的長度（元素數量）
            public bool HasChildrenItem()
            { //是否存在子項目
                if (this.ChildrenItem == null || this.Length == 0)
                { //為空或長度為零
                    Length = 0;
                    return false;
                }
                else
                {
                    Length = ChildrenItem.Count();
                    return true;
                }
            }
            public ColNames(string root, List<string> children)
            {//接受一個根項目的字串和子項目的字串清單作為參數
                RootItem = root;
                if (children == null)
                {
                    //ChildrenItem = null;
                    Length = 0;
                }
                else
                {
                    ChildrenItem = children;
                    Length = children.Count();
                }
            }
        }

        public class ColNamesMap
        {
            public int Length { get; set; } //映射中所有項目的總數
            public List<ColNames> colMap { get; set; }  //ColNames 的列表，即根項目和子項目的映射
            public ColNamesMap(List<ColNames> map)
            { //接受一個 ColNames 的列表作為參數
              //該建構函式將傳入的映射列表設置給 colMap 屬性
                var lenth = 0;
                colMap = map;
                foreach (var item0 in map)
                {
                    if (item0.ChildrenItem != null)
                    {
                        foreach (var item1 in item0.ChildrenItem)
                        {
                            lenth += 1;
                        }
                    }
                    else
                    {
                        lenth += 1;
                    }
                }
                this.Length = lenth;
            }
        }

        /// <summary>
        /// 自訂Column物件
        /// </summary>
        public class DisplayItem
        { //存儲顯示項目的相關資訊，包括名稱、儲存格類型和格式設定標誌
          //這些資訊通常用於將資料導出到 Excel 檔案時，指定每個儲存格的類型和格式
            public string Name { get; set; }

            public CellType CellType { get; set; }

            /// <summary>
            /// 彈性不夠 要在修
            /// </summary>
            public bool IsEnableFormat { get; set; }
        }

        /// <summary>
        /// 自訂列物件
        /// </summary>
        public class RowItem
        { //導出到 Excel 檔案中的一個儲存格的相關資訊
          //這些資訊通常用於將資料插入到 Excel 表格的特定儲存格中
            public string Name { get; set; }

            public CellType CellType { get; set; }

            public bool IsEnableFormat { get; set; }
        }

        public class ExcelDataTypeAttribute : Attribute
        {//自定義屬性
         //標記特定屬性的 Excel 資料類型，
         //以在導出到 Excel 時能夠正確地處理這些屬性的資料
            public ExcelDataTypeAttribute(DataType dataType)
            {
                this.DataType = dataType;
            }
            public DataType DataType { get; set; }

        }



        public class IntToAlphabet
        {
            const int ColumnBase = 26;
            const int DigitMax = 7; // ceil(log26(Int32.Max))
            const string Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            public static string IndexToColumn(int index)
            {
                var result = "A";
                try
                {

                    if (index <= ColumnBase)
                        return Digits[index - 1].ToString();

                    var sb = new StringBuilder().Append(' ', DigitMax);
                    var current = index;
                    var offset = DigitMax;
                    while (current > 0)
                    {
                        sb[--offset] = Digits[--current % ColumnBase];
                        current /= ColumnBase;
                    }
                    result = sb.ToString(offset, DigitMax - offset);
                }
                catch (Exception ex)
                {


                }
                return result;
            }
        }
    
}