namespace GenerateQuestion
{
    public enum StatusDataType
    {
        Success,
        Error
    }

    public class DataResponse<TDataType>
    {
        public TDataType? Data { get; set; } = default;
        public string Message { get; set; } = string.Empty;
        public StatusDataType Status { get; set; } = StatusDataType.Error;

        public DataResponse() { }

        public DataResponse(TDataType? data, string message, StatusDataType status)
        {
            Data = data;
            Message = message;
            Status = status;
        }

        public DataResponse(string message, StatusDataType status)
        {
            Message = message;
            Status = status;
        }
    }

    public class DifficultStructure
    {
        public int DifficultLevel { get; set; }
        public int NumberOfQuestion { get; set; }
    }
}
