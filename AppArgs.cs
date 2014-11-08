namespace DatabaseToExcel
{
    class AppArgs
    {
        [Argument(ArgumentType.Required, HelpText = "Database server name.")]
        public string server;

        [Argument(ArgumentType.Required, HelpText = "Database name.")]
        public string database;

        [Argument(ArgumentType.AtMostOnce, HelpText = "User login.")]
        public string user;

        [Argument(ArgumentType.AtMostOnce, HelpText = "User password.")]
        public string password;

        [Argument(ArgumentType.AtMostOnce, HelpText = "Login via Windows Authentication", DefaultValue = false)]
        public bool integratedSecurity;

        [Argument(ArgumentType.Required, HelpText = "Query file to run")]
        public string queryFile;

        [Argument(ArgumentType.AtMostOnce, HelpText = "Sheet file to run")]
        public string sheetFile;

        [Argument(ArgumentType.Required, HelpText = "Output Excel File name")]
        public string outputFile;

        [Argument(ArgumentType.AtMostOnce, HelpText = "Launch after creation", DefaultValue = false)]
        public bool launchAfterCreation;



        

        public override string ToString()
        {
            return string.Format("/s:{0} /d:{1} /u:{2} /p:{3} /i /queryFile:{4} /sheetFile:{5} /outputFile:{6}", server, database, user, password, queryFile, sheetFile, outputFile);
        }
    }
}
