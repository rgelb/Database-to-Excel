namespace DatabaseToExcel
{
    class AppArgs
    {
        [Argument(ArgumentType.Required, HelpText = "Database server name.")]
        public string server;

        [Argument(ArgumentType.Required, HelpText = "Database name.")]
        public string database;

        [Argument(ArgumentType.Required, HelpText = "User login.")]
        public string user;

        [Argument(ArgumentType.Required, HelpText = "User password.")]
        public string password;

        [Argument(ArgumentType.Required, HelpText = "Query file to run")]
        public string queryFile;

        [Argument(ArgumentType.AtMostOnce, HelpText = "Sheet file to run")]
        public string sheetFile;

        [Argument(ArgumentType.Required, HelpText = "Output Excel File name")]
        public string outputFile;
        

        public override string ToString()
        {
            return string.Format("/s:{0} /d:{1} /u:{2} /p:{3} /queryFile:{4} /sheetFile:{5} /outputFile:{6}", server, database, user, password, queryFile, sheetFile, outputFile);
        }
    }
}
