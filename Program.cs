using GenerateQuestion;
using System.Data;

QuestionGenerator generateQuestion = new();
DataResponse<DataTable> dataResponse = generateQuestion.ReadData<DataTable>(@"Z:\ReadFile_Test.xlsx");

if (dataResponse.Status == StatusDataType.Success)
{
    Console.WriteLine(dataResponse.Message);

    List<DifficultStructure> ds = new List<DifficultStructure>() {
        new DifficultStructure() { DifficultLevel = 1, NumberOfQuestion = 10 },
        new DifficultStructure() { DifficultLevel = 2, NumberOfQuestion = 20 },
        new DifficultStructure() { DifficultLevel = 3, NumberOfQuestion = 30 },
        new DifficultStructure() { DifficultLevel = 4, NumberOfQuestion = 40 },
        new DifficultStructure() { DifficultLevel = 5, NumberOfQuestion = 50 }
    };

    DataResponse<DataTable> dataResponseGenerate = generateQuestion.RandomQuestionsFromBank<DataTable>(dataResponse.Data, ds);
    if (dataResponseGenerate.Status == StatusDataType.Success)
    {
        Console.WriteLine(dataResponseGenerate.Message);

        DataResponse<DataTable> dataResponseExport = generateQuestion.ExportData<DataTable>(dataResponseGenerate.Data, @"Z:\Output", "QuestionExport");
        Console.WriteLine(dataResponseExport.Message);
    }
    else
    {
        Console.WriteLine(dataResponseGenerate.Message);
    }
}
else
{
    Console.WriteLine(dataResponse.Message);
}