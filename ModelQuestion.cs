namespace GenerateQuestion
{
    public class ModelQuestion
    {
        public string Question { get; set; } = string.Empty;
        public string AnswerA { get; set; } = string.Empty;
        public string AnswerB { get; set; } = string.Empty;
        public string AnswerC { get; set; } = string.Empty;
        public string AnswerD { get; set; } = string.Empty;
        public string CorrectAnswer { get; set; } = string.Empty;
        public int DifficultLevel { get; set; } = 1;
    }
}
