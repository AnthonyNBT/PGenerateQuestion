using System.Data;
using Newtonsoft.Json;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.ComponentModel.DataAnnotations;

namespace GenerateQuestion
{
    public class QuestionGenerator
    {
        public DataResponse<TDataType> ReadData<TDataType>(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return new DataResponse<TDataType>("File path is null or empty.", StatusDataType.Error);

            try
            {
                if (!File.Exists(filePath))
                    return new DataResponse<TDataType>("File path is not exists.", StatusDataType.Error);

                string fileExt = Path.GetExtension(filePath);
                if (fileExt.Equals(".xlsx"))
                {
                    using XLWorkbook excelWorkbook = new(filePath);
                    IXLWorksheet ws = excelWorkbook.Worksheet(1);
                    IXLRows rows = ws.RowsUsed();

                    return E_ReadToDataType<TDataType>(rows);
                }
                else if (fileExt.Equals(".docx"))
                {
                    List<string> rows = new();
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                    {
                        Body? body = wordDoc.MainDocumentPart?.Document.Body;

                        if (body == null) return new DataResponse<TDataType>("Can not read data from file.", StatusDataType.Error);

                        var reg = new Regex(@"^[\s\p{L}\d\•\-\►]");

                        foreach (Paragraph co in
                            body.Descendants<Paragraph>().Where<Paragraph>(somethingElse =>
                            reg.IsMatch(somethingElse.InnerText)))
                        {
                            rows.Add(co.InnerText);
                        }
                    }
                    return W_ReadToDataType<TDataType>(rows); ;
                }
                else
                {
                    return new($"File extension '{fileExt}' is not supported. Supported extensions are '.xlsx', '.docx'.", StatusDataType.Error);
                }
            }
            catch (Exception ex)
            {
                return new DataResponse<TDataType>(ex.Message, StatusDataType.Error);
            }
        }

        public DataResponse<TDataType> ExportData<TDataType>(TDataType exportValues, string filePath = "", string fileName = "Question_Export")
        {
            try
            {
                return ExportToDataType<TDataType>(filePath, fileName, exportValues);
            }
            catch (Exception ex)
            {
                return new(ex.Message, StatusDataType.Error);
            }
        }

        public DataResponse<TDataType> RandomQuestionsFromBank<TDataType>(TDataType questionBank, List<DifficultStructure> difficultStructures)
        {
            if (difficultStructures.Select(item => item.DifficultLevel).Distinct().Count() != difficultStructures.Count())
                return new("DifficultLevel is not unique from input.", StatusDataType.Error);

            List<ModelQuestion> outputQuestions = new List<ModelQuestion>();

            List<ModelQuestion>? questionBankConvert;

            switch (typeof(TDataType))
            {
                case Type type when type == typeof(DataTable):
                    DataTable? tempDataDT = questionBank as DataTable;

                    if (tempDataDT == null) return new DataResponse<TDataType>("Data source can not read.", StatusDataType.Error);
                    if (tempDataDT.Rows.Count <= 0) return new DataResponse<TDataType>("Data source is empty.", StatusDataType.Error);

                    questionBankConvert = ConvertData<DataTable, List<ModelQuestion>>(tempDataDT);
                    break;
                case Type type when (type == typeof(List<ModelQuestion>)):
                    List<ModelQuestion>? tempDataLs = questionBank as List<ModelQuestion>;

                    if (tempDataLs == null) return new DataResponse<TDataType>("Data source can not read.", StatusDataType.Error);
                    if (tempDataLs.Count <= 0) return new DataResponse<TDataType>("Data source is empty.", StatusDataType.Error);
                    
                    questionBankConvert = tempDataLs;
                    break;
                default:
                    return new DataResponse<TDataType>("Type is not supported.", StatusDataType.Error);
            }
            
            if (questionBankConvert == null) return new DataResponse<TDataType>("Data convert process fail.", StatusDataType.Error);

            foreach (DifficultStructure item in difficultStructures)
            {
                DataResponse<List<ModelQuestion>> responseEachDiffLevel = RandomQuestionFromBank(questionBankConvert, item);

                if (responseEachDiffLevel.Status == StatusDataType.Error)
                {
                    return new DataResponse<TDataType>(responseEachDiffLevel.Message, responseEachDiffLevel.Status);
                }
                if (responseEachDiffLevel == null || responseEachDiffLevel.Data == null) return new DataResponse<TDataType>("Data gennerate is null in process", StatusDataType.Error);

                outputQuestions = outputQuestions.Concat(responseEachDiffLevel.Data).ToList();
            }

            TDataType? outputDataTypeQuestionsConvert = ConvertData<List<ModelQuestion>, TDataType>(outputQuestions);

            if (outputDataTypeQuestionsConvert == null) return new DataResponse<TDataType>("Data gennerate out process convert is null", StatusDataType.Error);

            return new DataResponse<TDataType>(outputDataTypeQuestionsConvert, "Gennerate success", StatusDataType.Success);
        }

        private DataResponse<TDataType> E_ReadToDataType<TDataType>(IXLRows rows)
        {
            TDataType? resultDataTableConvert;

            switch (typeof(TDataType))
            {
                case Type type when type == typeof(DataTable):
                    DataResponse<DataTable> resultDataTable = E_ReadToDataTable(rows);
                    if (resultDataTable == null || resultDataTable.Data == null) return new DataResponse<TDataType>("Data process is null.", StatusDataType.Error);
                    resultDataTableConvert = ConvertData<DataTable, TDataType>(resultDataTable.Data);
                    if (resultDataTableConvert == null) return new DataResponse<TDataType>("Data convert process fail.", StatusDataType.Error);
                    return new DataResponse<TDataType>(resultDataTableConvert,
                                                       resultDataTable.Message,
                                                       resultDataTable.Status);

                case Type type when (type == typeof(List<ModelQuestion>)):
                    DataResponse<List<ModelQuestion>> resultDataListModel = E_ReadToList(rows);
                    if (resultDataListModel == null || resultDataListModel.Data == null) return new DataResponse<TDataType>("Data process is null.", StatusDataType.Error);
                    resultDataTableConvert = ConvertData<List<ModelQuestion>, TDataType>(resultDataListModel.Data);
                    if (resultDataTableConvert == null) return new DataResponse<TDataType>("Data convert process fail.", StatusDataType.Error);
                    return new DataResponse<TDataType>(resultDataTableConvert,
                                                       resultDataListModel.Message,
                                                       resultDataListModel.Status);

                default:
                    return new DataResponse<TDataType>("Type is not supported.", StatusDataType.Error);
            }
        }

        private DataResponse<TDataType> W_ReadToDataType<TDataType>(List<string> rows)
        {
            TDataType? resultDataTableConvert;
            
            switch (typeof(TDataType))
            {
                case Type type when type == typeof(DataTable):
                    DataResponse<DataTable> resultDataTable = W_ReadToDataTable(rows);
                    if (resultDataTable == null || resultDataTable.Data == null) return new DataResponse<TDataType>("Data process is null.", StatusDataType.Error);
                    resultDataTableConvert = ConvertData<DataTable, TDataType>(resultDataTable.Data);
                    if (resultDataTableConvert == null) return new DataResponse<TDataType>("Data convert process fail.", StatusDataType.Error);
                    return new DataResponse<TDataType>(resultDataTableConvert,
                                                       resultDataTable.Message,
                                                       resultDataTable.Status);

                case Type type when (type == typeof(List<ModelQuestion>) || type == typeof(IEnumerable<ModelQuestion>) || type == typeof(List<object>) || type == typeof(IEnumerable<object>)):
                    DataResponse<List<ModelQuestion>> resultDataListModel = W_ReadToList(rows);
                    if (resultDataListModel == null || resultDataListModel.Data == null) return new DataResponse<TDataType>("Data process is null.", StatusDataType.Error);
                    resultDataTableConvert = ConvertData<List<ModelQuestion>, TDataType>(resultDataListModel.Data);
                    if (resultDataTableConvert == null) return new DataResponse<TDataType>("Data convert process fail.", StatusDataType.Error);
                    return new DataResponse<TDataType>(JsonConvert.DeserializeObject<TDataType>(JsonConvert.SerializeObject(resultDataListModel.Data)) ?? default,
                                                       resultDataListModel.Message,
                                                       resultDataListModel.Status);

                default:
                    return new DataResponse<TDataType>("Type is not supported.", StatusDataType.Error);
            }
        }

        private DataResponse<DataTable> E_ReadToDataTable(IXLRows rows)
        {
            DataTable excelData = new();
            excelData.Columns.Add("Question");
            excelData.Columns.Add("AnswerA");
            excelData.Columns.Add("AnswerB");
            excelData.Columns.Add("AnswerC");
            excelData.Columns.Add("AnswerD");
            excelData.Columns.Add("CorrectAnswer");
            excelData.Columns.Add("DifficultLevel");

            foreach (IXLRow row in rows)
            {
                DataRow rowData = excelData.NewRow();

                rowData.SetField<string>(0, row.Cell(1).Value.ToString());
                rowData.SetField<string>(1, row.Cell(2).Value.ToString());
                rowData.SetField<string>(2, row.Cell(3).Value.ToString());
                rowData.SetField<string>(3, row.Cell(4).Value.ToString());
                rowData.SetField<string>(4, row.Cell(5).Value.ToString());
                rowData.SetField<string>(5, row.Cell(6).Value.ToString());
                rowData.SetField<int>(6, int.Parse(row.Cell(7).Value.ToString()));

                excelData.Rows.Add(rowData);
            }

            return new DataResponse<DataTable>(excelData, "Read data success.", StatusDataType.Success);
        }

        private DataResponse<DataTable> W_ReadToDataTable(List<string> rows)
        {
            DataTable excelData = new();
            excelData.Columns.Add("Question");
            excelData.Columns.Add("AnswerA");
            excelData.Columns.Add("AnswerB");
            excelData.Columns.Add("AnswerC");
            excelData.Columns.Add("AnswerD");
            excelData.Columns.Add("CorrectAnswer");
            excelData.Columns.Add("DifficultLevel");

            if (rows.Count % 7 != 0)
            {
                return new DataResponse<DataTable>("Number of data lines is missing. Each question need 7 line { 'Question', 'AnswerA', 'AnswerB', 'AnswerC', 'AnswerD', 'CorrectAnswer', 'DifficultLevel' }", StatusDataType.Error);
            }

            for (int i = 0; i < rows.Count / 7; i++)
            {
                DataRow rowData = excelData.NewRow();

                rowData.SetField<string>(0, rows[7 * i].ToString());
                rowData.SetField<string>(1, rows[7 * i + 1].ToString());
                rowData.SetField<string>(2, rows[7 * i + 2].ToString());
                rowData.SetField<string>(3, rows[7 * i + 3].ToString());
                rowData.SetField<string>(4, rows[7 * i + 4].ToString());
                rowData.SetField<string>(5, rows[7 * i + 5].ToString());
                rowData.SetField<int>(6, int.Parse(rows[7 * i + 6].ToString()));

                excelData.Rows.Add(rowData);
            }

            return new DataResponse<DataTable>(excelData, "Read data success.", StatusDataType.Success);
        }

        private DataResponse<List<ModelQuestion>> E_ReadToList(IXLRows rows)
        {
            List<ModelQuestion> excelData = new();

            foreach (IXLRow row in rows)
            {
                ModelQuestion rowData = new()
                {
                    Question = row.Cell(1).Value.ToString(),
                    AnswerA = row.Cell(2).Value.ToString(),
                    AnswerB = row.Cell(3).Value.ToString(),
                    AnswerC = row.Cell(4).Value.ToString(),
                    AnswerD = row.Cell(5).Value.ToString(),
                    CorrectAnswer = row.Cell(6).Value.ToString(),
                    DifficultLevel = int.Parse(row.Cell(7).Value.ToString())
                };

                excelData.Add(rowData);
            }

            return new DataResponse<List<ModelQuestion>>(excelData, "Read data success.", StatusDataType.Success);
        }

        private DataResponse<List<ModelQuestion>> W_ReadToList(List<string> rows)
        {
            List<ModelQuestion> excelData = new();

            if (rows.Count % 7 != 0)
            {
                return new DataResponse<List<ModelQuestion>>("Number of data lines is missing. Each question need 7 line { 'Question', 'AnswerA', 'AnswerB', 'AnswerC', 'AnswerD', 'CorrectAnswer', 'DifficultLevel' }", StatusDataType.Error);
            }

            for (int i = 0; i < rows.Count / 7; i++)
            {
                ModelQuestion rowData = new()
                {
                    Question = rows[7 * i].ToString(),
                    AnswerA = rows[7 * i + 1].ToString(),
                    AnswerB = rows[7 * i + 2].ToString(),
                    AnswerC = rows[7 * i + 3].ToString(),
                    AnswerD = rows[7 * i + 4].ToString(),
                    CorrectAnswer = rows[7 * i + 5].ToString(),
                    DifficultLevel = int.Parse(rows[7 * i + 6].ToString()),
                };

                excelData.Add(rowData);
            } 

            return new DataResponse<List<ModelQuestion>>(excelData, "Read data success.", StatusDataType.Success);
        }

        private DataResponse<TDataType> ExportToDataType<TDataType>(string filePath, string fileName, TDataType exportValues)
        {
            switch (typeof(TDataType))
            {
                case Type type when type == typeof(DataTable):
                    DataResponse<DataTable> resultDataTable = ExportFromDataTable(filePath, fileName, exportValues as DataTable ?? new());
                    return new DataResponse<TDataType>(resultDataTable.Message,
                                                       resultDataTable.Status);

                case Type type when (type == typeof(List<ModelQuestion>) || type == typeof(IEnumerable<ModelQuestion>) || type == typeof(List<object>) || type == typeof(IEnumerable<object>)):
                    DataResponse<List<ModelQuestion>> resultDataListModel = ExportFromList(filePath, fileName, exportValues as List<ModelQuestion> ?? new());
                    return new DataResponse<TDataType>(resultDataListModel.Message,
                                                       resultDataListModel.Status);

                default:
                    return new DataResponse<TDataType>("Type is not supported.", StatusDataType.Error);
            }
        }

        private DataResponse<DataTable> ExportFromDataTable(string filePath, string fileName, DataTable exportValues)
        {
            ExportToFile(filePath, fileName, exportValues);
            return new DataResponse<DataTable>("Export success.", StatusDataType.Success);
        }

        private DataResponse<List<ModelQuestion>> ExportFromList(string filePath, string fileName, List<ModelQuestion> exportValues)
        {
            DataTable dataTableConvert = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(exportValues)) ?? new();
            ExportToFile(filePath, fileName, dataTableConvert);

            return new DataResponse<List<ModelQuestion>>("Export success.", StatusDataType.Success);
        }

        private void ExportToFile(string filePath, string fileName, DataTable data)
        {
            using XLWorkbook wb = new();
            wb.AddWorksheet(data, "Sheet1");
            using MemoryStream stream = new();
            wb.SaveAs(stream);
            File.WriteAllBytes($"{filePath}/{fileName}.xlsx", stream.ToArray());
        }

        private DataResponse<List<ModelQuestion>> RandomQuestionFromBank(List<ModelQuestion> questionBank, DifficultStructure difficultStructure)
        {
            List<ModelQuestion> questions = questionBank.Where(item => item.DifficultLevel == difficultStructure.DifficultLevel).ToList<ModelQuestion>();
            if (questions.Count < difficultStructure.NumberOfQuestion) return new DataResponse<List<ModelQuestion>>($"Question Bank not enough question with DifficultLevel equal {difficultStructure.DifficultLevel}", StatusDataType.Error);
            return new DataResponse<List<ModelQuestion>>(questions.OrderBy(x => Guid.NewGuid()).Take(difficultStructure.NumberOfQuestion).ToList<ModelQuestion>(), "Success", StatusDataType.Success);
        }

        private TDataTypeB? ConvertData<TDataTypeA, TDataTypeB>(TDataTypeA sourceData)
        {
            return JsonConvert.DeserializeObject<TDataTypeB>(JsonConvert.SerializeObject(sourceData)) ?? default;
        }
    }
}
