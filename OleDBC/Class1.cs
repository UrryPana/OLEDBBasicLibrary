using System;
using System.Data.OleDb;

namespace OleDBC 
{
    /// <summary>
    /// It's a simple OLEDB library that can do basic Insert, Update, and Query.
    /// I Tried it on windowsForms from my projects that need oledb access like ms access. I will test it soon on others too. 
    /// Framework: .NET 9.0
    /// Written by: Urry John Paña
    /// </summary>
    public class OledB 
    {   
        public OleDbConnection conn { get; set; } // ConnectionString
        public OleDbCommand cmd { get; set; }  //Command
        
        /// <param name="connString">Refers to connection String. You can easily find it in the internet. </param>
        public OledB(string connString) 
        { 
            conn = new OleDbConnection(connString); 
            cmd = new OleDbCommand { Connection = conn }; 
        } 
        /// <summary>
        /// A method to query data. It only return boolean (true or false) if it found an  entry 
        /// </summary>
        /// <param name="commandText">An SQL query string</param>
        /// <param name="flag">
        /// <returns> it returns a boolean value through out keyword. Don't forget to initialize</returns>
        /// </param>
        public void QueryData(string commandText, out bool flag) 
        { 
            flag = false; 
            using (OleDbConnection conn = new OleDbConnection(this.conn.ConnectionString)) 
            { 
                using (OleDbCommand cmd = new OleDbCommand(commandText, conn)) 
                { 
                    try 
                    { 
                        conn.Open(); 
                        using (OleDbDataReader reader = cmd.ExecuteReader()) 
                        { 
                            while (reader.Read()) 
                            { 
                                flag = true; 
                            } 
                        } 
                    } 
                    catch (Exception ex) 
                    {
                        // Handle exception (e.g., log the error)
                       Console.WriteLine($"An error occurred: {ex.Message}"); 
                    } 
                } 
            } 
        }
        /// <summary>
        /// A method for inserting Data
        /// </summary>
        /// <param name="commandText"> Your SQL Query for inserting or updating data</param>
        /// <param name="paramName">the parameters that you need to add in the query</param>
        /// <param name="value">the objects your parameter are pointing</param>
        public void InsertUpdateData(string commandText, string[]? paramName, object[]? value)
        {
            using (OleDbConnection conn = new OleDbConnection(this.conn.ConnectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand(commandText, conn))
                {
                    conn.Open();
                    try
                    {
                        if (paramName != null && value != null) //checks if theres data present
                        {
                            if(paramName.Length == value.Length) //checks if the number of parameters and objects are the same
                            {
                                throw new ArgumentException("Parameter names and Parameter objects should be the same.");
                            }
                            for (int i = 0; i < paramName.Length; ++i)
                            {
                                cmd.Parameters.AddWithValue(paramName[i], value[i]);
                            }

                        }
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    { //Handles any exception
                        Console.WriteLine($"{ex.Message}");
                    }
                }
            }
        }
    }
} 